"""Lightweight identifier resolution via CIQRANGEA + CIQ company name lookup.

Builds a small workbook with just ID resolution formulas, launches Excel,
reads results. Much faster than a full comp pull (~10s vs ~40s).
"""
from __future__ import annotations

import logging
import os
import time

from openpyxl import Workbook

log = logging.getLogger("capiq_mcp")


def build_id_workbook(path: str, identifiers: list[str]) -> None:
    """Build a workbook with CIQRANGEA ID lookup + CIQ company name per identifier.

    Layout:
      Col A: identifier (raw input)
      Col B: =CIQRANGEA(id, "IQ_COMPANY_ID_QUICK_MATCH", 1, 1)  (formula)
      Col C: (blank — CIQRANGEA spill result = resolved IQ ID)
      Col D: =CIQ(spill_ref, "IQ_COMPANY_NAME")  (name from resolved ID)
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Lookup"

    ws.cell(row=1, column=1, value="Input")
    ws.cell(row=1, column=2, value="CIQRANGEA_Formula")
    ws.cell(row=1, column=3, value="Resolved_IQ_ID")
    ws.cell(row=1, column=4, value="Company_Name")

    for idx, ident in enumerate(identifiers):
        row = idx + 2
        ws.cell(row=row, column=1, value=ident)
        ws.cell(row=row, column=2,
                value=f'=CIQRANGEA("{ident}","IQ_COMPANY_ID_QUICK_MATCH",1,1)')
        # Col C is blank — CIQRANGEA result spills here
        # Company name from resolved IQ ID (references spill cell)
        ws.cell(row=row, column=4,
                value=f'=CIQ($C${row},"IQ_COMPANY_NAME")')

    wb.save(path)
    log.info("Created ID lookup workbook: %s (%d identifiers)", path, len(identifiers))


def run_id_lookup(identifiers: list[str], max_wait: int = 60) -> dict:
    """Resolve identifiers to CIQ IQ IDs and company names.

    Parameters
    ----------
    identifiers : list[str]
        Tickers, company names, CUSIPs, ISINs, etc.
    max_wait : int
        Max seconds to wait for formula refresh.

    Returns
    -------
    dict
        {"results": [{"input": ..., "iq_id": ..., "company_name": ..., "status": ...}]}
    """
    import pythoncom
    pythoncom.CoInitialize()

    from capiq_excel.excel_lifecycle import launch_excel_isolated, close_session

    xlsx_path = os.path.abspath("_id_lookup_mcp_temp.xlsx")
    build_id_workbook(xlsx_path, identifiers)

    session = launch_excel_isolated(xlsx_path)
    try:
        excel = session.excel
        log.info("Connected to Excel %s for ID lookup", excel.Version)
        ws_com = session.workbook.Sheets(1)

        # Trigger refresh
        try:
            excel.Run("SNLXLAddin.xla!RefreshSheet")
            log.info("RefreshSheet executed for ID lookup")
        except Exception as e:
            log.warning("RefreshSheet warning: %s", e)

        # Poll for completion — just check CIQRANGEA spill cells
        start = time.monotonic()
        while True:
            elapsed = time.monotonic() - start
            if elapsed > max_wait:
                log.warning("ID lookup timeout after %ds", max_wait)
                break

            all_done = True
            for idx in range(len(identifiers)):
                row = idx + 2
                val = ws_com.Cells(row, 3).Value  # spill column
                if val is None:
                    all_done = False
                    break
                if isinstance(val, str) and val.upper() in ('#PEND', '#REFRESH'):
                    all_done = False
                    break

            if all_done:
                log.info("ID lookup resolved in %.0fs", elapsed)
                break

            time.sleep(2)

        # Read results
        results = []
        for idx, ident in enumerate(identifiers):
            row = idx + 2
            iq_id = ws_com.Cells(row, 3).Value  # spill cell
            company_name = ws_com.Cells(row, 4).Value

            # Classify result
            if iq_id is None or (isinstance(iq_id, str) and
                                  any(tok in iq_id.upper() for tok in
                                      ("#ERROR", "#INVALID", "#NAME", "KEYERROR"))):
                results.append({
                    "input": ident,
                    "iq_id": None,
                    "company_name": None,
                    "status": "FAILED",
                })
            else:
                name = None
                if company_name and isinstance(company_name, str):
                    upper = company_name.upper()
                    if not any(tok in upper for tok in ("#ERROR", "#INVALID", "#NAME")):
                        name = company_name
                results.append({
                    "input": ident,
                    "iq_id": str(iq_id),
                    "company_name": name,
                    "status": "OK",
                })
    finally:
        close_session(session, save=False, delete_workbook=True)

    log.info("ID lookup complete: %d/%d resolved",
             sum(1 for r in results if r["status"] == "OK"), len(results))
    return {"results": results}
