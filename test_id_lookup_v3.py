"""
Follow-up test:
  1. SPG with various metric names to find the right one for ID lookup
  2. CIQRANGEA with the blank-column layout (formula in col A, result lands in col B)
"""
import os
import time
import openpyxl
import pythoncom

from capiq_excel.excel_lifecycle import launch_excel_isolated, close_session

pythoncom.CoInitialize()

XLSX_PATH = r"C:\Claude\test_id_v3.xlsx"

# Test identifiers
TEST_IDS = ["Microsoft", "03783310", "NVDA", "US0378331005"]

# SPG metric names to try for ID resolution
SPG_METRICS = [
    "IQ_COMPANY_ID",
    "IQ_COMPANY_ID_QUICK_MATCH",
    "COMPANY_ID",
    "SP_COMPANY_ID",
    "IQ_ENTITY_ID",
]


def create_workbook():
    wb = openpyxl.Workbook()

    # Sheet 1: SPG metric name exploration
    ws1 = wb.active
    ws1.title = "SPG Metrics"

    # Headers
    ws1.cell(row=1, column=1, value="Input")
    for j, metric in enumerate(SPG_METRICS):
        ws1.cell(row=1, column=j + 2, value=metric)

    for i, identifier in enumerate(TEST_IDS):
        row = i + 2
        ws1.cell(row=row, column=1, value=identifier)
        for j, metric in enumerate(SPG_METRICS):
            ws1.cell(row=row, column=j + 2,
                     value=f'=SPG("{identifier}","{metric}")')

    # Sheet 2: CIQRANGEA with blank column layout (legacy pattern)
    ws2 = wb.create_sheet("CIQRANGEA Layout")
    ws2.cell(row=1, column=1, value="Input")
    ws2.cell(row=1, column=2, value="CIQRANGEA Formula")
    ws2.cell(row=1, column=3, value="Result (spill)")
    ws2.cell(row=1, column=4, value="CIQ() Fallback")

    for i, identifier in enumerate(TEST_IDS):
        row = i + 2
        ws2.cell(row=row, column=1, value=identifier)
        # CIQRANGEA in col B — result should spill into col C
        ws2.cell(row=row, column=2,
                 value=f'=CIQRANGEA("{identifier}","IQ_COMPANY_ID_QUICK_MATCH",1,1)')
        # Col C left blank for spill
        # Col D: CIQ() as fallback
        ws2.cell(row=row, column=4,
                 value=f'=CIQ("{identifier}","IQ_COMPANY_ID")')

    wb.save(XLSX_PATH)
    print(f"Created workbook: {XLSX_PATH}")


def open_refresh_and_extract():
    print("\nLaunching Excel...")
    session = launch_excel_isolated(XLSX_PATH)
    excel = session.excel

    try:
        excel.Run("SNLXLAddin.xla!RefreshWorkbook")
        print("Triggered RefreshWorkbook")
    except Exception as e:
        print(f"Refresh warning: {e}")

    print("Waiting for evaluation...")
    for tick in range(30):
        time.sleep(5)
        elapsed = (tick + 1) * 5
        ws = session.workbook.Sheets(1)
        val = ws.Cells(2, 2).Value
        print(f"  [{elapsed}s] B2={val!r}")
        if val is not None and str(val) not in ("#REFRESH", "#PEND", ""):
            time.sleep(10)
            break
    else:
        print("  Timeout")

    # --- Sheet 1: SPG metrics ---
    ws1 = session.workbook.Sheets("SPG Metrics")
    print("\n" + "=" * 100)
    print("SPG() with various metric names:")
    print("-" * 100)
    header = f"{'Input':<20}"
    for metric in SPG_METRICS:
        header += f" {metric:<18}"
    print(header)
    print("-" * 100)

    for i in range(len(TEST_IDS)):
        row = i + 2
        line = f"{str(ws1.Cells(row, 1).Value):<20}"
        for j in range(len(SPG_METRICS)):
            val = ws1.Cells(row, j + 2).Value
            s = str(val)[:17] if val is not None else "(empty)"
            line += f" {s:<18}"
        print(line)
    print("=" * 100)

    # --- Sheet 2: CIQRANGEA layout ---
    ws2 = session.workbook.Sheets("CIQRANGEA Layout")
    print("\nCIQRANGEA with blank-column layout:")
    print("-" * 90)
    print(f"{'Input':<20} {'Formula Cell (B)':<22} {'Spill Cell (C)':<22} {'CIQ() Fallback (D)':<22}")
    print("-" * 90)

    for i in range(len(TEST_IDS)):
        row = i + 2
        input_id = str(ws2.Cells(row, 1).Value)
        formula_cell = str(ws2.Cells(row, 2).Value) if ws2.Cells(row, 2).Value else "(empty)"
        spill_cell = str(ws2.Cells(row, 3).Value) if ws2.Cells(row, 3).Value else "(empty)"
        ciq_cell = str(ws2.Cells(row, 4).Value) if ws2.Cells(row, 4).Value else "(empty)"
        print(f"{input_id:<20} {formula_cell[:21]:<22} {spill_cell[:21]:<22} {ciq_cell[:21]:<22}")

    print("=" * 90)

    close_session(session, save=False, delete_workbook=False)


if __name__ == "__main__":
    create_workbook()
    open_refresh_and_extract()
