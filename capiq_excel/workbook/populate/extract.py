import re
import pandas as pd
import numpy as np
from exceldriver.columns import excel_cols


CIQ_MARKET_ITEM_FORMULA_PATTERN = re.compile(r'=CIQ\("[\w]+", "[\w]+", "([\w\/]+)"\)')


def extract_capiq_df_from_sheet(ws, market_data_items):
    market_values = set(market_data_items.values())
    col_gen = excel_cols()

    all_series = []
    while True:
        col = next(col_gen)
        if col == 'A':
            continue  # A is date column, don't need to extract
        header = ws.Range(f'{col}1').Value
        if header is None:
            # Blank column label, data is finished
            break
        if header in market_values:
            all_series.append(get_series_from_ciq_market_item_col(ws, col))
        else:
            all_series.append(get_series_from_ciq_financial_col(ws, col))

    df = pd.concat(all_series, axis=1)
    df.index.name = 'Date'
    return df


# xlUp constant for End() method
_XL_UP = -4162


def _find_last_row(ws, col: str) -> int:
    """Find the last non-empty row in a column using Excel's End(xlUp)."""
    return ws.Cells(ws.Rows.Count, col).End(_XL_UP).Row


def get_series_from_ciq_market_item_col(ws, col: str):
    name = ws.Range(f'{col}1').Value
    last_row = _find_last_row(ws, col)
    if last_row < 2:
        series = pd.Series(name=name, dtype=float)
        series.index = pd.to_datetime(series.index)
        return series

    # Batch read values and formulas in 2 COM calls instead of 2*N
    values = ws.Range(f'{col}2:{col}{last_row}').Value
    formulas = ws.Range(f'{col}2:{col}{last_row}').Formula

    data = {}
    for val_tuple, formula_tuple in zip(values, formulas):
        val = val_tuple[0] if isinstance(val_tuple, tuple) else val_tuple
        formula = formula_tuple[0] if isinstance(formula_tuple, tuple) else formula_tuple
        if val is None:
            break
        date = get_date_from_ciq_market_item_formula(formula)
        data[date] = val

    series = pd.Series(data, name=name)
    series.index = pd.to_datetime(series.index)
    return series


def get_series_from_ciq_financial_col(ws, col: str):
    name = ws.Range(f'{col}1').Value
    last_row = _find_last_row(ws, col)
    if last_row < 2:
        series = pd.Series(name=name, dtype=float)
        series.index = pd.to_datetime(series.index)
        return series

    # Batch read data column and date column A in 2 COM calls instead of 2*N
    values = ws.Range(f'{col}2:{col}{last_row}').Value
    dates = ws.Range(f'A2:A{last_row}').Value

    data = {}
    for val_tuple, date_tuple in zip(values, dates):
        val = val_tuple[0] if isinstance(val_tuple, tuple) else val_tuple
        if val is None:
            break
        date = date_tuple[0] if isinstance(date_tuple, tuple) else date_tuple
        if isinstance(date, str):
            if date == 'NA':
                date = np.nan
        else:
            date = str(date.date())
        data[date] = val

    series = pd.Series(data, name=name)
    series.index = pd.to_datetime(series.index)
    return series


def get_date_from_ciq_market_item_formula(formula: str) -> str:
    match = CIQ_MARKET_ITEM_FORMULA_PATTERN.match(formula)
    if not match:
        return None
    return match.group(1)