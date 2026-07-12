"""Non-destructive Excel lifecycle management.

Launches an isolated Excel instance via subprocess with the /x flag
(forces a new process) and finds it via the Running Object Table by
workbook name -- so existing user Excel windows are never touched.
"""
from __future__ import annotations

import os
import subprocess
import time
from dataclasses import dataclass
from functools import wraps
from typing import Optional

import logging

import pythoncom
import win32api
import win32com.client

from exceldriver.path import get_excel_path
from exceldriver.tools import (
    _start_excel_with_addins_and_attach as _ed_start_excel,
    _restart_excel_with_addins_and_attach as _ed_restart_excel,
)

log = logging.getLogger("capiq_excel")


def com_initialized(func):
    """Run a callable inside a balanced COM apartment initialization."""
    @wraps(func)
    def wrapper(*args, **kwargs):
        pythoncom.CoInitialize()
        try:
            return func(*args, **kwargs)
        finally:
            pythoncom.CoUninitialize()
    return wrapper


# ── Visibility control ─────────────────────────────────────────────────────
# Excel is hidden by default so automated runs (comps / chart / lookup and the
# batch download pipeline) don't pop a window or steal focus.  Set the env var
# CAPIQ_EXCEL_VISIBLE=1 (or pass visible=True) to show Excel for debugging.

_TRUTHY = {"1", "true", "yes", "on"}


def _env_visible_default() -> bool:
    """Default Excel visibility from the environment — hidden unless opted in.

    Returns True only when CAPIQ_EXCEL_VISIBLE is set to a truthy value
    (1/true/yes/on); otherwise False (hidden).
    """
    raw = os.environ.get("CAPIQ_EXCEL_VISIBLE")
    if raw is None:
        return False
    return raw.strip().lower() in _TRUTHY


def _resolve_visible(visible: Optional[bool]) -> bool:
    """An explicit argument always wins; otherwise use the env default."""
    return _env_visible_default() if visible is None else bool(visible)


def apply_excel_visibility(excel, visible: Optional[bool] = None) -> bool:
    """Best-effort set Excel window visibility.  Never raises.

    Should be called on a fully-initialized Excel.Application (i.e. after the
    add-ins have loaded and UDFs registered) so it can't interfere with UDF
    registration or the add-in's window-handle acquisition.

    Returns the effective visibility that was requested.
    """
    eff = _resolve_visible(visible)
    try:
        excel.Visible = eff
    except Exception:
        log.debug("Could not set Excel.Visible=%s", eff, exc_info=True)
    return eff


def start_excel_with_addins_and_attach(*args, visible: Optional[bool] = None, **kwargs):
    """exceldriver `_start_excel_with_addins_and_attach` + visibility control.

    Launches Excel and waits for add-ins to attach (exceldriver handles the
    subprocess launch and UDF-safe startup), then applies the requested
    visibility.  Used by the legacy download / ids pipeline so it honors
    CAPIQ_EXCEL_VISIBLE the same way as `launch_excel_isolated`.
    """
    excel = _ed_start_excel(*args, **kwargs)
    apply_excel_visibility(excel, visible)
    return excel


def restart_excel_with_addins_and_attach(*args, visible: Optional[bool] = None, **kwargs):
    """exceldriver `_restart_excel_with_addins_and_attach` + visibility control.

    Used in the batch pipeline's restart/retry loops so Excel stays hidden
    across mid-job restarts.
    """
    excel = _ed_restart_excel(*args, **kwargs)
    apply_excel_visibility(excel, visible)
    return excel


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
    visible: Optional[bool] = None,
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
    visible : bool, optional
        Whether the Excel window should be shown.  Default None resolves to
        the CAPIQ_EXCEL_VISIBLE env var (hidden unless set truthy).  Excel is
        launched, add-ins initialize, UDFs register, and only *then* is the
        window hidden — so visibility never affects data retrieval.

    Returns
    -------
    ExcelSession
        Session object with .excel, .workbook, .process handles.
    """
    abs_path = os.path.abspath(workbook_path)
    wb_name = os.path.basename(abs_path)
    eff_visible = _resolve_visible(visible)

    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"Workbook not found: {abs_path}")

    # Get Excel path from Windows registry (no hardcoded path)
    excel_exe = get_excel_path()

    # Launch Excel via ShellExecute so it's created as a child of the
    # Windows Shell (explorer.exe), fully outside the MCP server's process
    # tree and any Windows Job Objects constraining it.  subprocess.Popen
    # creates a child process that inherits the parent's job object, which
    # can block Excel from launching until the parent is interrupted.
    # SW_SHOWNORMAL (1) when visible; SW_SHOWMINNOACTIVE (7) when hiding, so the
    # launch doesn't grab focus before we hide the window after add-ins init.
    show_cmd = 1 if eff_visible else 7
    log.info("Launching Excel via ShellExecute: %s /x %s (visible=%s)",
             excel_exe, abs_path, eff_visible)
    win32api.ShellExecute(
        0,                       # hwnd (no parent window)
        "open",                  # verb
        excel_exe,               # program
        f'/x "{abs_path}"',      # parameters
        None,                    # working directory
        show_cmd,                # SW_SHOWNORMAL (visible) / SW_SHOWMINNOACTIVE (hidden)
    )
    log.info("ShellExecute returned — Excel launch initiated")

    # Phase 1: Find our workbook in the Running Object Table
    log.info("Phase 1: Polling ROT for workbook '%s' ...", wb_name)
    app, wb_com = _poll_rot_for_workbook(
        wb_name,
        interval=rot_poll_interval,
        timeout=rot_poll_timeout,
        full_path=abs_path,
    )

    # Suppress modal dialogs that would block COM calls from MCP/headless callers.
    # Interactive=False prevents ALL user-facing dialogs (including add-in dialogs).
    # DisplayAlerts=False suppresses Excel's own confirmation prompts.
    # EnableEvents=False prevents event-driven interruptions during automation.
    try:
        app.Interactive = False
        app.DisplayAlerts = False
        app.EnableEvents = False
    except Exception:
        pass

    # Phase 2: Poll until add-in UDFs are registered
    log.info("Phase 1 complete. Phase 2: Waiting for UDFs ...")
    _wait_for_udfs(app, timeout=addin_init_timeout, interval=addin_poll_interval)
    log.info("Phase 2 complete. Excel session ready.")

    # Add-ins are now fully initialized and UDFs are registered — safe to hide
    # the window.  Doing it here (rather than at launch) guarantees visibility
    # never interferes with UDF registration or data retrieval.
    apply_excel_visibility(app, eff_visible)
    log.info("Excel visibility set to %s", eff_visible)

    return ExcelSession(
        workbook_path=abs_path,
        workbook_name=wb_name,
        process=None,  # ShellExecute doesn't return a process handle; cleanup uses COM app.Quit()
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
    # Restore Interactive mode before closing (some add-ins need it for cleanup)
    if session.excel is not None:
        try:
            session.excel.Interactive = True
            session.excel.EnableEvents = True
        except Exception:
            pass

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
            os.remove(session.workbook_path)
        except Exception:
            pass


# ── ROT helpers ──────────────────────────────────────────────────────────


def _find_workbook_in_rot(workbook_name: str, full_path: Optional[str] = None):
    """Search the Running Object Table for a workbook by filename.

    Adapted from exceldriver.tools._get_excel_running_workbook().

    When *full_path* is provided, an exact (case-insensitive) path match is
    preferred and short-circuits the scan — this binds precisely to the
    workbook we launched instead of risking a same-basename workbook open in a
    different folder/instance (which close_session would then Quit and delete).
    A basename-suffix match is kept as a fallback for resilience.

    Returns
    -------
    tuple[Application, Workbook] or None
        The COM Application and Workbook objects, or None if not found.
    """
    pythoncom.CoInitialize()
    rot = pythoncom.GetRunningObjectTable()
    rotenum = rot.EnumRunning()
    target_len = len(workbook_name)
    exact_target = os.path.normcase(os.path.abspath(full_path)) if full_path else None
    exact_obj = None    # matched by full path (preferred)
    suffix_obj = None   # matched by basename suffix (fallback)

    while True:
        monikers = rotenum.Next()
        if not monikers:
            break
        try:
            ctx = pythoncom.CreateBindCtx(0)
            display_name = monikers[0].GetDisplayName(ctx, None)
            if exact_target is not None and os.path.normcase(display_name) == exact_target:
                exact_obj = rot.GetObject(monikers[0])
                break  # definitively our workbook — stop scanning
            if display_name[-target_len:] == workbook_name:
                suffix_obj = rot.GetObject(monikers[0])
        except Exception:
            continue

    obj = exact_obj if exact_obj is not None else suffix_obj
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
    full_path: Optional[str] = None,
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
    attempt = 0
    while time.monotonic() < deadline:
        attempt += 1
        log.debug("ROT poll attempt %d (%.1fs elapsed)", attempt,
                  timeout - (deadline - time.monotonic()))
        result = _find_workbook_in_rot(workbook_name, full_path=full_path)
        if result is not None:
            log.info("Found workbook '%s' in ROT after %.1fs",
                     workbook_name, timeout - (deadline - time.monotonic()))
            return result
        # Pump COM message queue to prevent STA deadlocks
        pythoncom.PumpWaitingMessages()
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

        # Pump COM message queue to prevent STA deadlocks
        pythoncom.PumpWaitingMessages()
        time.sleep(interval)
