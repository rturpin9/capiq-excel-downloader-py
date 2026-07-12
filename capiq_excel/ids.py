import math
from typing import Sequence, List
import pandas as pd

from capiq_excel.excel_lifecycle import start_excel_with_addins_and_attach as _start_excel_with_addins_and_attach
from processfiles.files import FileProcessTracker
from capiq_excel.workbook.populate.main import populate_capiq_ids_for_file
from capiq_excel.workbook.create import create_all_xlsx_with_id_commands
from capiq_excel.fileops import clear_generated_xlsx_files


def download_capiq_ids(ids: Sequence[str], outpath: str = 'capiq ids.csv', folder: str = 'in_process_ids',
                       config=None) -> List[str]:
    """
    Downloads Capital IQ identifiers when passed other identifiers such as CUSIP,
    ISIN, ticker, name, etc.

    Stores in a CSV with matched names included and also returns capiq ids as a list

    :param ids: identifiers such as CUSIP, ISIN, ticker, name. Can be a mixture.
    :param folder: folder which will hold in process files
    :param outpath: filepath to output csv, including file extension
    :param config: Optional CapiqConfig for dialect and add-in mode.
    :return: capiq ids
    """
    if not ids:
        raise ValueError("ids must not be empty")

    # Resolve builder for formula generation
    builder = None
    if config is not None:
        from capiq_excel.formulas import get_builder
        from capiq_excel.config import FormulaDialect
        dialect = config.formula_dialect
        if dialect == FormulaDialect.AUTO:
            # AUTO wasn't resolved upstream; default to CIQ (safe fallback)
            dialect = config.resolve_dialect(ciq_compat_available=True)
        builder = get_builder(dialect)

    print('Creating XLSX files with commands to get ids')
    clear_generated_xlsx_files(folder)
    num_files = math.ceil(len(ids) / 100)
    create_all_xlsx_with_id_commands(ids, folder, num_files=num_files, builder=builder)

    print('Populating XLSX files for ids')
    populate_all_ids_in_folder(folder, config=config)

    print('Combining all ids into a single CSV file')
    combine_all_capiq_ids_xlsx(folder, outpath)

    return _get_ids_from_csv_path(outpath)


def populate_all_ids_in_folder(folder, restart=True, config=None):
    excel = _start_excel_with_addins_and_attach()

    try:
        effective_addin_mode = None
        # When config is available, detect runtime and load the right add-in
        if config is not None:
            from capiq_excel.addin import load_capiq_addin
            from capiq_excel.config import AddinMode
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

        file_tracker = FileProcessTracker(folder=folder, restart=restart, file_types=('xlsx',))

        for file in file_tracker.file_generator():
            populate_capiq_ids_for_file(
                file,
                excel,
                config=config,
                addin_mode=effective_addin_mode,
                refresh_timeout=(config.retry.timeout_seconds if config else 240),
            )
    finally:
        try:
            excel.Quit()
        except Exception:
            pass


def combine_all_capiq_ids_xlsx(infolder, outpath, restart=True):
    file_tracker = FileProcessTracker(folder=infolder, restart=restart, file_types=('xlsx',))

    dfs = [pd.read_excel(f) for f in file_tracker.file_generator()]
    df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

    _remove_useless_cols(df)
    df.to_csv(outpath, index=False)

    return df


def _remove_useless_cols(df):
    """
    NOTE: inplace
    """
    blank_cols = [col for col in df.columns if 'blank' in col.lower()]
    useless_cols = ['ID'] + blank_cols
    existing_cols = [col for col in useless_cols if col in df.columns]
    if existing_cols:
        df.drop(existing_cols, axis=1, inplace=True)


def _get_ids_from_csv_path(csv_path: str) -> List[str]:
    df = pd.read_csv(csv_path)
    if 'IQID' not in df.columns:
        raise ValueError(f"ID output is missing the IQID column: {csv_path}")

    ids = []
    for value in df['IQID']:
        if pd.isna(value):
            continue
        normalized = str(value).strip()
        if not normalized or normalized.upper().startswith(("#", "INVALID")):
            continue
        ids.append(normalized)

    if not ids:
        raise ValueError(f"No identifiers resolved successfully in {csv_path}")
    return ids
