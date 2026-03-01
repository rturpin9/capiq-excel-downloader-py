from typing import Dict
import pandas as pd
import warnings
from concurrent.futures import ThreadPoolExecutor, TimeoutError

from capiq_excel.fileops import get_path_of_failed_folder_add_if_necessary, move_file_to_failed_folder, get_path_of_additional_failed_folder_add_if_necessary
from capiq_excel.workbook.populate.main import populate_capiq_for_file
from exceldriver.tools import _start_excel_with_addins_and_attach, _restart_excel_with_addins_and_attach
from processfiles.files import FileProcessTracker



def populate_all_files_in_folder(folder, financial_data_items_dict: Dict[str, str],
                                 market_data_items_dict: Dict[str, str], restart=True, timeout=240,
                                 run_failed=False, config=None):
    """
    Populate all XLSX files in a folder by opening them in Excel and letting
    the Capital IQ plugin evaluate the formulas.

    When *config* is provided, uses the new refresh engine and runtime detection.
    """
    _validate_populate_inputs(folder, restart, run_failed)

    excel = _start_excel_with_addins_and_attach()

    try:
        # When config is available, detect runtime and load the right add-in
        if config is not None:
            from capiq_excel.addin import load_capiq_addin
            try:
                profile, addin_name = load_capiq_addin(excel, config)
            except Exception as e:
                print(f'Warning: Could not detect/load add-in ({e}), falling back to default behavior')

        failed_folder = get_path_of_failed_folder_add_if_necessary(folder)

        if run_failed:
            # Set main folder as 'failed', then set failed folder as another failed folder inside the original
            folder = failed_folder
            failed_folder = get_path_of_additional_failed_folder_add_if_necessary(folder)

        file_tracker = FileProcessTracker(folder=folder, restart=restart, file_types=('xlsx',))

        effective_timeout = timeout
        if config is not None:
            effective_timeout = config.retry.timeout_seconds

        with ThreadPoolExecutor(max_workers=1) as e:
            for i, file in enumerate(file_tracker.file_generator()):

                excel, successful = _try_to_get_result_if_fail_restart_excel(
                    e,
                    i,
                    file,
                    excel,
                    financial_data_items_dict=financial_data_items_dict,
                    market_data_items_dict=market_data_items_dict,
                    config=config,
                    timeout=effective_timeout,
                )

                if not successful:
                    move_file_to_failed_folder(file, failed_folder)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass


def _validate_populate_inputs(folder, restart, run_failed):
    if restart and run_failed:
        raise ValueError("Cannot set both restart=True and run_failed=True")

    if run_failed:
        warnings.warn(f'run_failed flag passed. Folder {folder}, will not be run, instead the failed folder will')

### Functions below to assist with multiprocessing/timeout handling

def _try_to_get_result_if_fail_restart_excel(threadpool, i, file, excel, tries_remaining=3,
                                              timeout=300, **kwargs):

    if tries_remaining <= 0:
        print(fr'ERROR: Timed out processing {file}. Skipping and moving to "..\failed".')
        return excel, False

    # Submit result on threadpool asynchronously
    async_result = threadpool.submit(populate_capiq_for_file, file, excel, index=i + 1, **kwargs)

    # Try to get the result, waiting up to timeout seconds.
    try:
        excel, successful = async_result.result(timeout=timeout)
    except Exception as e:
        if not isinstance(e, TimeoutError):
            print(f'ERROR: Unexpected error processing {file}: {e}')
        try:
            excel.Quit()
        except Exception:
            pass
        excel = _restart_excel_with_addins_and_attach()
        return _try_to_get_result_if_fail_restart_excel(threadpool, i, file, excel,
                                                         tries_remaining=tries_remaining - 1,
                                                         timeout=timeout, **kwargs)

    return excel, successful
