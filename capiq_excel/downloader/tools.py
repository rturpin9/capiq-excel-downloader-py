from typing import Dict
import warnings

from capiq_excel.fileops import get_path_of_failed_folder_add_if_necessary, move_file_to_failed_folder, get_path_of_additional_failed_folder_add_if_necessary
from capiq_excel.workbook.populate.main import populate_capiq_for_file
from capiq_excel.excel_lifecycle import (
    start_excel_with_addins_and_attach as _start_excel_with_addins_and_attach,
    restart_excel_with_addins_and_attach as _restart_excel_with_addins_and_attach,
)
from processfiles.files import FileProcessTracker
from capiq_excel.config import AddinMode



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
        effective_addin_mode = None
        if config is not None:
            from capiq_excel.addin import load_capiq_addin
            try:
                profile, _ = load_capiq_addin(excel, config)
                if config.addin_mode == AddinMode.AUTO:
                    effective_addin_mode = (
                        AddinMode.LEGACY
                        if profile.addin_mode == "legacy"
                        else AddinMode.PRO
                    )
                else:
                    effective_addin_mode = config.addin_mode
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

        max_attempts = 3
        if config is not None:
            max_attempts = max(1, config.retry.max_retries)

        for i, file in enumerate(file_tracker.file_generator()):
            excel, successful = _try_to_get_result_if_fail_restart_excel(
                i,
                file,
                excel,
                tries_remaining=max_attempts,
                financial_data_items_dict=financial_data_items_dict,
                market_data_items_dict=market_data_items_dict,
                config=config,
                addin_mode=effective_addin_mode,
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

# Helpers for bounded retries and Excel restart handling.

def _try_to_get_result_if_fail_restart_excel(i, file, excel, tries_remaining=3,
                                              timeout=300, **kwargs):
    try:
        excel, successful = populate_capiq_for_file(
            file,
            excel,
            index=i + 1,
            retries_remaining=tries_remaining,
            refresh_timeout=timeout,
            **kwargs,
        )
    except Exception as e:
        print(f'ERROR: Unexpected error processing {file}: {e}')
        try:
            excel.Quit()
        except Exception:
            pass
        excel = _restart_excel_with_addins_and_attach()
        return excel, False

    if not successful:
        try:
            excel.Quit()
        except Exception:
            pass
        excel = _restart_excel_with_addins_and_attach()

    return excel, successful
