from typing import Sequence, Dict, Iterable, Callable, Optional
import os
import itertools
import math
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows

from exceldriver.workbook.create import get_workbook_and_worksheet
from exceldriver.columns import excel_cols
from .commands import (
    financial_data_command, id_command, name_command, holdings_command, market_data_command,
    make_financial_command, make_market_command, make_holdings_command, make_id_command, make_name_command,
)

from capiq_excel.config import FormulaOptions
from capiq_excel.formulas.base import DialectBuilder


def create_all_xlsx_with_commands(folder: str, company_id_list: Sequence[str],
                                  financial_data_items_dict: Dict[str, str],
                                  market_data_items_dict: Dict[str, str],
                                  builder: Optional[DialectBuilder] = None,
                                  formula_options: Optional[FormulaOptions] = None,
                                  **financials_kwargs):
    for company_id in company_id_list:
        create_xlsx_with_commands(
            folder,
            company_id,
            financial_data_items_dict,
            market_data_items_dict,
            builder=builder,
            formula_options=formula_options,
            **financials_kwargs
        )


def create_xlsx_with_commands(folder: str, company_id: str, financial_data_items_dict: Dict[str, str],
                              market_data_items_dict: Dict[str, str],
                              builder: Optional[DialectBuilder] = None,
                              formula_options: Optional[FormulaOptions] = None,
                              **financials_kwargs):
    wb, ws = get_workbook_and_worksheet()
    _fill_with_commands(ws, company_id, financial_data_items_dict, market_data_items_dict,
                        builder=builder, formula_options=formula_options,
                        **financials_kwargs)

    os.makedirs(folder, exist_ok=True)

    filepath = os.path.join(folder, f'{company_id}.xlsx')
    wb.save(filepath)

    return os.path.abspath(filepath)

def create_all_xlsx_with_holdings_commands(folder, company_id_list, date_str_list, data_items_dict,
                                           builder: Optional[DialectBuilder] = None):
    for company_id, date_str in itertools.product(company_id_list, date_str_list):
        create_xlsx_with_holdings_commands(folder, company_id, date_str, data_items_dict, builder=builder)


def create_xlsx_with_holdings_commands(folder, company_id, date_str, data_items_dict,
                                       builder: Optional[DialectBuilder] = None):
    wb, ws = get_workbook_and_worksheet()
    _fill_with_holdings_commands(ws, company_id, date_str, data_items_dict, builder=builder)

    os.makedirs(folder, exist_ok=True)
    filepath = os.path.join(folder, f'{company_id} {_date_str_to_file_format(date_str)}.xlsx')
    wb.save(filepath)

    return os.path.abspath(filepath)

def create_all_xlsx_with_id_commands(ids: Sequence[str], folder, num_files=100,
                                     builder: Optional[DialectBuilder] = None):
    if not ids:
        raise ValueError("ids must not be empty")
    if num_files <= 0:
        raise ValueError("num_files must be greater than zero")
    os.makedirs(folder, exist_ok=True)

    df = pd.DataFrame()
    _fill_id_column(df, ids)
    _fill_capiq_id_column(df, builder=builder)
    _fill_capiq_name_column(df, builder=builder)

    rows_per_ws = math.ceil(len(df) / num_files)
    for file_index, start in enumerate(range(0, len(df), rows_per_ws), start=1):
        wb, ws = get_workbook_and_worksheet()
        chunk = df.iloc[start:start + rows_per_ws]
        for row in dataframe_to_rows(chunk, index=False, header=True):
            ws.append(row)
        filepath = os.path.join(folder, f'ids {file_index}.xlsx')
        wb.save(filepath)

##### Helper functions ####

def _fill_id_column(df: pd.DataFrame, ids: Sequence[str]):
    """
    NOTE: inplace
    """
    df['ID'] = ids

def _fill_capiq_id_column(df, builder: Optional[DialectBuilder] = None):
    """
    NOTE: inplace
    """
    id_cmd = make_id_command(builder) if builder is not None else id_command

    if builder is not None and not builder.identifier_lookup_spills_right():
        df['IQID'] = df['ID'].apply(id_cmd)
    else:
        # CIQRANGEA expands one column to the right — formula goes in Blank,
        # resolved value spills into IQID.
        df['Blank 1'] = df['ID'].apply(id_cmd)
        df['IQID'] = ''


def _fill_capiq_name_column(df, builder: Optional[DialectBuilder] = None):
    """
    NOTE: inplace
    """
    name_cmd = make_name_command(builder) if builder is not None else name_command

    if builder is not None and not builder.identifier_lookup_spills_right():
        df['IQ Name'] = df['ID'].apply(name_cmd)
    else:
        # CIQRANGEA expands one column to the right — formula goes in Blank,
        # resolved value spills into IQ Name.
        df['Blank 2'] = df['ID'].apply(name_cmd)
        df['IQ Name'] = ''

def _fill_with_commands(ws, company_id: str, financial_data_items_dict: Dict[str, str],
                        market_data_items_dict: Dict[str, str],
                        builder: Optional[DialectBuilder] = None,
                        formula_options: Optional[FormulaOptions] = None,
                        **financials_kwargs):
    """
    Note: inplace
    """
    # Resolve command functions: builder-aware when available, legacy otherwise
    if builder is not None:
        fin_cmd = make_financial_command(builder, formula_options)
        mkt_cmd = make_market_command(builder, formula_options)
    else:
        fin_cmd = financial_data_command
        mkt_cmd = market_data_command

    # Set default freq
    try:
        freq = financials_kwargs['freq']
    except KeyError:
        freq = 'Q'

    date_dict = _get_date_var_dict_from_freq(freq)

    column_generator = excel_cols()

    common_args = (
        column_generator,
        ws,
        company_id
    )

    _fill_ws_by_data_item_dict(date_dict, fin_cmd, *common_args, **financials_kwargs)
    _fill_ws_by_data_item_dict(financial_data_items_dict, fin_cmd, *common_args, **financials_kwargs)
    _fill_ws_by_data_item_dict(market_data_items_dict, mkt_cmd, *common_args, **financials_kwargs)



def _fill_ws_by_data_item_dict(data_items_dict: Dict[str, str], command_func: Callable,
                               column_generator: Iterable, ws,
                               company_id: str, **financials_kwargs):
    for item in data_items_dict:
        current_column = next(column_generator)
        ws[f'{current_column}1'] = command_func(
            company_id,
            item,
            **financials_kwargs,
            data_item_label=data_items_dict[item]
        )


def _fill_with_holdings_commands(ws, company_id, date_str, data_items_dict,
                                  builder: Optional[DialectBuilder] = None):
    hold_cmd = make_holdings_command(builder) if builder is not None else holdings_command
    column_generator = excel_cols()

    for item in data_items_dict:
        current_column = next(column_generator)
        ws[f'{current_column}1'] = hold_cmd(
            company_id, item, date_str,
            data_item_label=data_items_dict[item]
        )

def _get_date_var_dict_from_freq(freq) -> Dict[str, str]:
    date_dict = {
        'IQ_PERIODDATE_BS': 'Date'
    }
    if freq == 'Y':
        date_dict.update(dict(
            IQ_FISCAL_Y='Fiscal Year'
        ))
    elif freq == 'Q':
        date_dict.update(dict(
            IQ_ABS_PERIOD='Fiscal Quarter'
        ))
    else:
        raise ValueError('Must pass Y or Q for freq')

    return date_dict

def _date_str_to_file_format(date_str):
    return date_str.replace('/','-')
