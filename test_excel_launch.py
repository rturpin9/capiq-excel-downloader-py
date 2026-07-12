"""Standalone test: does Excel launch and appear in the ROT?

Run this directly from a terminal (NOT through Claude Code / MCP):
    cd C:\Claude
    python test_excel_launch.py

This isolates whether the launch issue is MCP-specific.
"""
import os
import subprocess
import sys
import time
import tempfile

import pythoncom
import win32api
import win32com.client
from openpyxl import Workbook

from exceldriver.path import get_excel_path


def create_test_workbook():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "test"
    path = os.path.join(tempfile.gettempdir(), "capiq_launch_test.xlsx")
    wb.save(path)
    return path


def find_in_rot(workbook_name):
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
    wb = win32com.client.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
    return wb.Application, wb


def test_launch(method="shellexecute"):
    wb_path = create_test_workbook()
    wb_name = os.path.basename(wb_path)
    excel_exe = get_excel_path()

    print(f"Excel: {excel_exe}")
    print(f"Workbook: {wb_path}")
    print(f"Method: {method}")
    print()

    t0 = time.monotonic()

    if method == "shellexecute":
        print(f"[{0:.1f}s] Calling ShellExecute ...")
        win32api.ShellExecute(0, "open", excel_exe, f'/x "{wb_path}"', None, 1)
        print(f"[{time.monotonic()-t0:.1f}s] ShellExecute returned")
    elif method == "popen":
        print(f"[{0:.1f}s] Calling subprocess.Popen ...")
        proc = subprocess.Popen([excel_exe, "/x", wb_path])
        print(f"[{time.monotonic()-t0:.1f}s] Popen returned (PID {proc.pid})")
    elif method == "popen_devnull":
        print(f"[{0:.1f}s] Calling subprocess.Popen (DEVNULL) ...")
        proc = subprocess.Popen(
            [excel_exe, "/x", wb_path],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
        )
        print(f"[{time.monotonic()-t0:.1f}s] Popen returned (PID {proc.pid})")

    print(f"[{time.monotonic()-t0:.1f}s] Polling ROT for '{wb_name}' ...")

    for i in range(30):
        result = find_in_rot(wb_name)
        elapsed = time.monotonic() - t0
        if result is not None:
            print(f"[{elapsed:.1f}s] FOUND in ROT!")
            app, wb = result
            print(f"[{elapsed:.1f}s] Quitting Excel ...")
            try:
                app.DisplayAlerts = False
                wb.Close(SaveChanges=False)
                app.Quit()
            except Exception:
                pass
            os.remove(wb_path)
            print(f"[{time.monotonic()-t0:.1f}s] Done. SUCCESS.")
            return
        print(f"[{elapsed:.1f}s] Not found yet (attempt {i+1}/30)")
        pythoncom.PumpWaitingMessages()
        time.sleep(2)

    print(f"FAILED — workbook never appeared in ROT after 60s")
    os.remove(wb_path)


if __name__ == "__main__":
    method = sys.argv[1] if len(sys.argv) > 1 else "shellexecute"
    test_launch(method)
