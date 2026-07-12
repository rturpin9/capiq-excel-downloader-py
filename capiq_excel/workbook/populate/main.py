import sys
import time
import traceback
from typing import Dict

import pythoncom
from win32com.client import constants
from pywintypes import com_error

from capiq_excel.excel_lifecycle import restart_excel_with_addins_and_attach as _restart_excel_with_addins_and_attach
from capiq_excel.workbook.wait import _wait_for_capiq_result
from capiq_excel.exceptions import (
    WorkbookClosedException,
    CapitalIQInactiveException,
    RefreshTimeoutError,
)
from capiq_excel.workbook.populate.extract import extract_capiq_df_from_sheet
from capiq_excel.workbook.populate.replace import write_df_to_ws_values


def _wait_for_data(
    excel,
    config=None,
    addin_mode=None,
    timeout_seconds=None,
    allow_cell_errors=False,
):
    """Wait for plugin to finish evaluating formulas.

    Uses the new refresh engine when *config* is provided (Pro-aware with
    expanded token detection).  Falls back to legacy cell-A2 polling otherwise.
    """
    if config is not None:
        from capiq_excel.config import AddinMode
        from capiq_excel.refresh.engine import refresh_and_wait
        resolved_mode = addin_mode
        if resolved_mode is None:
            resolved_mode = (
                AddinMode.LEGACY
                if config.addin_mode == AddinMode.LEGACY
                else AddinMode.PRO
            )
        refresh_and_wait(
            excel,
            config,
            addin_mode=resolved_mode,
            allow_cell_errors=allow_cell_errors,
        )
    else:
        _wait_for_capiq_result(
            excel,
            timeout_seconds=timeout_seconds or 240,
        )


def populate_capiq_for_file(filepath, excel, financial_data_items_dict: Dict[str, str],
                            market_data_items_dict: Dict[str, str], retries_remaining=3,
                            close_workbook=False, index=0, config=None,
                            addin_mode=None, refresh_timeout=None):
    """
    Private function has main functionality. This is a wrapper to add retries afer com errors
    """

    # Balance COM initialization even when a workbook attempt fails.
    pythoncom.CoInitialize()
    try:
        restart_interval = 500
        if config is not None:
            restart_interval = config.retry.restart_interval

        # Even healthy automation sessions are recycled periodically to bound
        # Excel/add-in memory growth.
        if restart_interval > 0 and index > 0 and index % restart_interval == 0:
            excel = _restart_excel_with_addins_and_attach()

        max_attempts = max(1, retries_remaining)
        for attempt in range(max_attempts):
            try:
                if close_workbook or attempt > 0:
                    excel.CutCopyMode = False
                    if excel.ActiveWorkbook:
                        excel.ActiveWorkbook.Close(SaveChanges=False)

                _populate_capiq_for_file(
                    filepath,
                    excel,
                    financial_data_items_dict,
                    market_data_items_dict,
                    config=config,
                    addin_mode=addin_mode,
                    refresh_timeout=refresh_timeout,
                )
                return excel, True
            except (
                com_error,
                WorkbookClosedException,
                CapitalIQInactiveException,
                RefreshTimeoutError,
            ) as e:
                print(
                    f'Error {e} populating {filepath} '
                    f'(attempt {attempt + 1}/{max_attempts}).'
                )
                traceback.print_tb(sys.exc_info()[2])
                if attempt + 1 >= max_attempts:
                    break
                retry_delay = 30
                if config is not None:
                    retry_delay = config.retry.retry_delay_seconds
                if retry_delay > 0:
                    time.sleep(retry_delay)
                excel = _restart_excel_with_addins_and_attach()

        print(fr'ERROR: Could not process {filepath}. Skipping and moving to "..\failed".')
        return excel, False
    finally:
        pythoncom.CoUninitialize()


def _populate_capiq_for_file(filepath, excel, financial_data_items_dict: Dict[str, str],
                            market_data_items_dict: Dict[str, str], config=None,
                            addin_mode=None, refresh_timeout=None):
    wb = excel.Workbooks.Open(filepath)
    _wait_for_data(
        excel,
        config=config,
        addin_mode=addin_mode,
        timeout_seconds=refresh_timeout,
    )
    _set_date_format(excel, wb, cell_range='A:A')  # column A is automatically included date
    _extract_unaligned_data_align_and_replace(wb, market_data_items_dict)
    excel.ActiveWorkbook.Close(SaveChanges=True)
    return True


def populate_capiq_ids_for_file(
    filepath,
    excel,
    config=None,
    addin_mode=None,
    refresh_timeout=None,
):
    wb = excel.Workbooks.Open(filepath)
    _wait_for_data(
        excel,
        config=config,
        addin_mode=addin_mode,
        timeout_seconds=refresh_timeout,
        allow_cell_errors=True,
    )
    _copy_paste_values(excel, wb, cell_range='A:Z')
    excel.ActiveWorkbook.Close(SaveChanges=True)
    return True


def _copy_paste_values(excel, wb, cell_range='A1:ZZ20000'):
    ws = wb.Sheets('Sheet')
    ws.Range(cell_range).Copy()
    range_begin = cell_range.split(':')[0]
    if not range_begin[-1].isdigit():
        # Got a full column, e.g. A instead of A1. Start from the first row
        range_begin += '1'
    ws.Range(range_begin).PasteSpecial(Paste=constants.xlPasteValues, Operation=constants.xlNone)
    excel.CutCopyMode = False


def _set_date_format(excel, wb, cell_range='B:B'):
    ws = wb.Sheets('Sheet')
    ws.Range(cell_range).NumberFormat = 'mm/dd/yyyy'


def _extract_unaligned_data_align_and_replace(wb, market_data_items_dict: Dict[str, str]):
    """
    Market data items become unaligned with the date axis. Extract the date from each individual formula,
    then combine everything and replace the contents of the worksheet
    """
    ws = wb.Sheets('Sheet')
    df = extract_capiq_df_from_sheet(ws, market_data_items_dict)
    write_df_to_ws_values(df, ws)
