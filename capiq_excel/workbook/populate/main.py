import sys
import time
import traceback
from typing import Dict, Optional

import pythoncom
from win32com.client import constants
from pywintypes import com_error

from capiq_excel.excel_lifecycle import restart_excel_with_addins_and_attach as _restart_excel_with_addins_and_attach
from capiq_excel.workbook.wait import _wait_for_capiq_result
from capiq_excel.exceptions import WorkbookClosedException, CapitalIQInactiveException
from capiq_excel.workbook.populate.extract import extract_capiq_df_from_sheet
from capiq_excel.workbook.populate.replace import write_df_to_ws_values


def _wait_for_data(excel, config=None):
    """Wait for plugin to finish evaluating formulas.

    Uses the new refresh engine when *config* is provided (Pro-aware with
    expanded token detection).  Falls back to legacy cell-A2 polling otherwise.
    """
    if config is not None:
        from capiq_excel.config import AddinMode
        from capiq_excel.refresh.engine import refresh_and_wait
        # For AUTO mode, try Pro refresh (it handles both Pro and CIQ-compat).
        # Only use LEGACY when explicitly configured.
        if config.addin_mode == AddinMode.LEGACY:
            addin_mode = AddinMode.LEGACY
        else:
            addin_mode = AddinMode.PRO
        refresh_and_wait(excel, config, addin_mode=addin_mode)
    else:
        _wait_for_capiq_result(excel)


def populate_capiq_for_file(filepath, excel, financial_data_items_dict: Dict[str, str],
                            market_data_items_dict: Dict[str, str], retries_remaining=3,
                            close_workbook=False, index=0, config=None):
    """
    Private function has main functionality. This is a wrapper to add retries afer com errors
    """

    # Necessary to be called in each new thread or process which uses COM (communicate with Microsoft products)
    pythoncom.CoInitialize()

    restart_interval = 500
    if config is not None:
        restart_interval = config.retry.restart_interval

    # Even if things are going normally, restart every N worksheets as there is a memory leak
    if index % restart_interval == 0 and retries_remaining == 3:
        excel = _restart_excel_with_addins_and_attach()

    # Stop retries
    if retries_remaining <= 0:
        print(fr'ERROR: Could not process {filepath}. Skipping and moving to "..\failed".')
        return excel, False

    try:
        # If we are retrying, need to close the workbook before trying to populate
        if close_workbook:
            excel.CutCopyMode = False
            time.sleep(1)
            if excel.ActiveWorkbook:
                excel.ActiveWorkbook.Close(SaveChanges=False)
            time.sleep(5)

        _populate_capiq_for_file(filepath, excel, financial_data_items_dict, market_data_items_dict,
                                 config=config)
        return excel, True
    except (com_error, WorkbookClosedException, CapitalIQInactiveException) as e:
        print(f'Error {e} populating {filepath}. Will wait 30 seconds, restart Excel, and try again.')
        traceback.print_tb(sys.exc_info()[2])
        retry_delay = 30
        if config is not None:
            retry_delay = config.retry.retry_delay_seconds
        time.sleep(retry_delay)
        excel = _restart_excel_with_addins_and_attach()
        return populate_capiq_for_file(
            filepath,
            excel,
            financial_data_items_dict,
            market_data_items_dict,
            retries_remaining=retries_remaining - 1,
            close_workbook=True,
            index=index,
            config=config,
        )


def _populate_capiq_for_file(filepath, excel, financial_data_items_dict: Dict[str, str],
                            market_data_items_dict: Dict[str, str], config=None):
    wb = excel.Workbooks.Open(filepath)
    _wait_for_data(excel, config=config)
    _set_date_format(excel, wb, cell_range='A:A')  # column A is automatically included date
    _extract_unaligned_data_align_and_replace(wb, market_data_items_dict)
    excel.ActiveWorkbook.Close(SaveChanges=True)
    return True


def populate_capiq_ids_for_file(filepath, excel, config=None):
    wb = excel.Workbooks.Open(filepath)
    _wait_for_data(excel, config=config)
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
