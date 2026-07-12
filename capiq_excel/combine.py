import os
import re
import tempfile
import math
import pandas as pd

from capiq_excel.tools.ext_pandas import _get_outpath_and_df_of_headers, _append_df_to_csv
from processfiles.files import FileProcessTracker


def combine_all_capiq_xlsx(infolder, outpath, restart=True, num_parts=100):
    if num_parts <= 0:
        raise ValueError("num_parts must be greater than zero")
    target_outpath = outpath or 'combined.csv'
    file_tracker = FileProcessTracker(folder=infolder, restart=restart, file_types=('xlsx',))
    num_files = len(file_tracker.process_list)
    if num_files == 0:
        if not restart and os.path.exists(target_outpath):
            return target_outpath
        raise ValueError(f"No XLSX files found to combine in {infolder}")

    target_dir = os.path.dirname(os.path.abspath(target_outpath))
    os.makedirs(target_dir, exist_ok=True)
    temporary_output = None
    working_outpath = target_outpath
    if restart:
        fd, temporary_output = tempfile.mkstemp(
            prefix=".capiq-combine-",
            suffix=".csv",
            dir=target_dir,
        )
        os.close(fd)
        os.remove(temporary_output)
        working_outpath = temporary_output

    try:
        _, df_of_headers = _get_outpath_and_df_of_headers(working_outpath)
        all_columns = list(df_of_headers.columns)
        num_files_per_part = max(1, math.ceil(num_files / num_parts))
        file_num = 0

        with tempfile.TemporaryDirectory(prefix="capiq-combine-parts-") as temp_dir:
            for i, file in enumerate(file_tracker.file_generator()):
                if i % num_files_per_part == 0:
                    file_num += 1
                part_path = os.path.join(temp_dir, f'{file_num}.csv')
                df_of_headers, all_columns = _append_capiq_xlsx_to_csv(
                    file,
                    part_path,
                    df_of_headers,
                    all_columns,
                )

            part_tracker = FileProcessTracker(
                folder=temp_dir,
                restart=True,
                file_types=('csv',),
            )
            _, df_of_headers = _get_outpath_and_df_of_headers(working_outpath)
            all_columns = list(df_of_headers.columns)
            for file in part_tracker.file_generator():
                df_for_append = pd.read_csv(file)
                df_of_headers, all_columns = _append_df_to_csv(
                    df_for_append,
                    df_of_headers,
                    working_outpath,
                    all_columns,
                )

        if restart:
            os.replace(working_outpath, target_outpath)
        return target_outpath
    finally:
        if temporary_output and os.path.exists(temporary_output):
            os.remove(temporary_output)

def _append_capiq_xlsx_to_csv(file, outpath, df_of_headers, all_columns):
    df_for_append = pd.read_excel(file)  # load new data
    if _filepath_has_date(file):
        id_, date = _capiq_filepath_to_iq_id_and_date(file)
        df_for_append['CQID'] = id_
        df_for_append['Date'] = date
    else:
        df_for_append['CQID'] = _capiq_filepath_to_iq_id(file)
    df_of_headers, all_columns = _append_df_to_csv(df_for_append, df_of_headers, outpath, all_columns)

    return df_of_headers, all_columns

_PATTERN_ID_DATE = re.compile(r'(IQ\d+) ([\d-]+)([.]xlsx)')
_PATTERN_ID_ONLY = re.compile(r'(IQ\d+)([.]xlsx)')


def _filepath_has_date(filepath):
    filename = os.path.basename(filepath)  # strips folders, etc.
    return bool(_PATTERN_ID_DATE.match(filename))

def _capiq_filepath_to_iq_id(filepath):
    filename = os.path.basename(filepath)  # strips folders, etc.
    match = _PATTERN_ID_ONLY.match(filename)
    if not match:
        raise ValueError(f"Could not extract IQ ID from filename: {filename}")
    return match.group(1)

def _capiq_filepath_to_iq_id_and_date(filepath):
    filename = os.path.basename(filepath)  # strips folders, etc.
    match = _PATTERN_ID_DATE.match(filename)
    if not match:
        raise ValueError(f"Could not extract IQ ID and date from filename: {filename}")
    return match.group(1), match.group(2)
