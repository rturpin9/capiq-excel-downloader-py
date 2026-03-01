"""Non-destructive Excel lifecycle management.

Launches an isolated Excel instance via subprocess with the /x flag
(forces a new process) and finds it via the Running Object Table by
workbook name -- so existing user Excel windows are never touched.
"""
from __future__ import annotations

import os
import subprocess
import time
from dataclasses import dataclass, field
from typing import Optional

import logging

import pythoncom
import win32com.client

from exceldriver.path import get_excel_path

log = logging.getLogger("capiq_mcp")


@dataclass
class ExcelSession:
    """Holds references to an isolated Excel instance we launched."""
    workbook_path: str
    workbook_name: str
    process: Optional[subprocess.Popen] = None
    excel: Optional[object] = None       # COM Excel.Application
    workbook: Optional[object] = None    # COM Workbook


def launch_excel_isolated(
    workbook_path: str,
    addin_init_timeout: float = 60.0,
    addin_poll_interval: float = 2.0,
    rot_poll_interval: float = 2.0,
    rot_poll_timeout: float = 60.0,
) -> ExcelSession:
    """Launch Excel in a new isolated process and connect via the ROT.

    Instead of blindly sleeping for add-in initialization, this polls
    the ROT for the workbook first, then checks whether UDFs are
    registered by evaluating a lightweight probe formula.

    Parameters
    ----------
    workbook_path : str
        Absolute path to the .xlsx file to open.
    addin_init_timeout : float
        Max seconds to wait for add-in UDFs to register (default 60).
    addin_poll_interval : float
        Seconds between UDF readiness checks (default 2).
    rot_poll_interval : float
        Seconds between ROT polling attempts (default 2).
    rot_poll_timeout : float
        Max seconds to poll the ROT before giving up (default 60).

    Returns
    -------
    ExcelSession
        Session object with .excel, .workbook, .process handles.
    """
    abs_path = os.path.abspath(workbook_path)
    wb_name = os.path.basename(abs_path)

    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"Workbook not found: {abs_path}")

    # Get Excel path from Windows registry (no hardcoded path)
    excel_exe = get_excel_path()

    # /x forces a new separate Excel process (does NOT suppress add-ins)
    proc = subprocess.Popen([excel_exe, "/x", abs_path])

    # Phase 1: Find our workbook in the Running Object Table
    app, wb_com = _poll_rot_for_workbook(
        wb_name,
        interval=rot_poll_interval,
        timeout=rot_poll_timeout,
    )

    # Suppress modal dialogs that would block COM calls from MCP/headless callers
    try:
        app.DisplayAlerts = False
    except Exception:
        pass

    # Phase 2: Poll until add-in UDFs are registered
    _wait_for_udfs(app, timeout=addin_init_timeout, interval=addin_poll_interval)

    return ExcelSession(
        workbook_path=abs_path,
        workbook_name=wb_name,
        process=proc,
        excel=app,
        workbook=wb_com,
    )


def close_session(
    session: ExcelSession,
    save: bool = False,
    delete_workbook: bool = True,
) -> None:
    """Close only our workbook/instance, leaving other Excel windows alone.

    Parameters
    ----------
    session : ExcelSession
        The session returned by launch_excel_isolated().
    save : bool
        Whether to save the workbook before closing (default False).
    delete_workbook : bool
        Whether to delete the temp XLSX file (default True).
    """
    # Close our workbook
    if session.workbook is not None:
        try:
            session.workbook.Close(SaveChanges=save)
        except Exception:
            pass

    # Quit our Excel instance
    if session.excel is not None:
        try:
            session.excel.Quit()
        except Exception:
            pass

    # Terminate our subprocess
    if session.process is not None:
        try:
            session.process.terminate()
        except Exception:
            pass

    # Remove the temp file
    if delete_workbook and session.workbook_path:
        try:
            if os.path.exists(session.workbook_path):
                os.remove(session.workbook_path)
        except Exception:
            pass


# ── ROT helpers ──────────────────────────────────────────────────────────


def _find_workbook_in_rot(workbook_name: str):
    """Search the Running Object Table for a workbook by filename.

    Adapted from exceldriver.tools._get_excel_running_workbook().

    Returns
    -------
    tuple[Application, Workbook] or None
        The COM Application and Workbook objects, or None if not found.
    """
    pythoncom.CoInitialize()
    rot = pythoncom.GetRunningObjectTable()
    rotenum = rot.EnumRunning()
    target_len = len(workbook_name)
    obj = None

    while True:
        monikers = rotenum.Next()
        if not monikers:
            break
        try:
            ctx = pythoncom.CreateBindCtx(0)
            display_name = monikers[0].GetDisplayName(ctx, None)
            if display_name[-target_len:] == workbook_name:
                obj = rot.GetObject(monikers[0])
        except Exception:
            continue

    if obj is None:
        return None

    wb = win32com.client.Dispatch(
        obj.QueryInterface(pythoncom.IID_IDispatch)
    )
    return wb.Application, wb


def _poll_rot_for_workbook(
    workbook_name: str,
    interval: float = 2.0,
    timeout: float = 60.0,
):
    """Poll the ROT until the workbook appears, with timeout.

    Returns
    -------
    tuple[Application, Workbook]

    Raises
    ------
    TimeoutError
        If the workbook doesn't appear in the ROT within *timeout* seconds.
    """
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        result = _find_workbook_in_rot(workbook_name)
        if result is not None:
            return result
        time.sleep(interval)

    raise TimeoutError(
        f"Workbook '{workbook_name}' did not appear in the Running Object "
        f"Table within {timeout:.0f}s. Excel may not have started correctly."
    )


def _wait_for_udfs(
    app,
    timeout: float = 60.0,
    interval: float = 2.0,
) -> None:
    """Poll until CIQ/SPG UDFs are registered in Excel.

    Uses Application.Evaluate() to test a lightweight probe formula.
    When UDFs aren't registered yet, Evaluate returns a COM error code
    (-2146826259 = #NAME?) or None. Once registered, it returns a
    string or numeric value (even an error like #INVALID COMPANY ID
    means the UDF is working).

    Tries SPG first (Pro add-in), then CIQ (compat mode). Either one
    succeeding means the add-in is ready.
    """
    # Probe formulas — use nonsense identifier so they evaluate fast
    # (we only care that the UDF is callable, not that it returns data)
    probes = [
        'SPG("__PROBE__","SP_COMPANY_NAME")',
        'CIQ("__PROBE__","IQ_COMPANY_NAME")',
    ]

    start = time.monotonic()
    while True:
        elapsed = time.monotonic() - start
        if elapsed > timeout:
            # Don't raise — proceed anyway; formulas may still work
            # once RefreshSheet is called
            log.info("UDF readiness: timeout after %.0fs, proceeding anyway", timeout)
            return

        for probe in probes:
            try:
                result = app.Evaluate(probe)
            except Exception:
                continue

            # COM error -2146826259 = #NAME? → UDF not registered yet
            if isinstance(result, int) and result == -2146826259:
                continue

            # None → not evaluated yet
            if result is None:
                continue

            # Any other result (string, number, or a different COM error)
            # means the UDF is callable — add-in is ready
            log.info("UDF ready after %.1fs", elapsed)
            return

        time.sleep(interval)
