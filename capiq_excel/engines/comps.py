"""Core comp table engine — extracted from pull_comps.py.

All business logic for building, refreshing, and reading CIQ comp tables.
No print() — uses logging throughout. Returns structured dicts for consumption.
"""
from __future__ import annotations

import logging
import os
import time
from dataclasses import dataclass
from datetime import datetime

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

import pythoncom

log = logging.getLogger("capiq_excel")

# ── Constants ─────────────────────────────────────────────────────────────

_US_EXCHANGES = frozenset({
    "NYSE", "NASDAQGS", "NASDAQGM", "NASDAQCM",
    "AMEX", "NYSEAMERICAN", "NYSEARCA", "BATS",
})

_ERROR_TOKENS = (
    "#ERROR", "#INVALID", "#PEND", "#REFRESH", "#NAME",
    "#OUTSIDE", "KEYERROR", "DEFUNCT", "INVALID", "NM", "(INVALID",
)

# ── Metric definitions ─────────────────────────────────────────────────────

METRICS_LEASE_ADJUSTED = [
    # (label,           mnemonic,                           call_type)
    ("Company Name",    "SP_COMPANY_NAME",                  "name"),
    ("Market Cap",      "SP_MARKETCAP",                     "market"),
    ("Total Debt",      "IQ_TOTAL_DEBT",                    "bs"),
    ("Oper. Leases",    "IQ_TOTAL_OPER_LEASES",             "bs"),
    ("Cash",            "IQ_CASH_ST_INVEST",                "bs"),
    ("Net Debt",        "IQ_NET_DEBT",                      "bs"),
    ("TEV",             "IQ_TEV",                           "market"),
    ("LTM Revenue",     "IQ_TOTAL_REV",                     "ltm"),
    ("NTM Revenue",     "SP_REV_EST",                       "ntm"),
    ("LTM EBITDA",      "IQ_EBITDA_EQ_INC",                 "ltm"),
    ("Lease Adj",       "IQ_LEASE_ADJUSTMENT_EBITDA",        "ltm"),
    ("NTM EBITDA",      "SP_EBITDA_EST",                    "ntm"),
    ("GAAP",            "IQ_GAAP_BS",                       "gaap"),
    ("P/BV",            "IQ_PBV_X",                         "ltm_raw"),
    ("Gross Margin",    "IQ_GROSS_MARGIN",                  "ltm_raw"),
    ("CapEx",           "IQ_CAPEX",                         "ltm"),
]

METRICS_EXCLUDING_LEASES = [
    ("Company Name",    "SP_COMPANY_NAME",                  "name"),
    ("Market Cap",      "SP_MARKETCAP",                     "market"),
    ("Total Debt",      "IQ_TOTAL_DEBT_EXCL_OPER_LEASES",  "bs"),
    ("Oper. Leases",    "IQ_TOTAL_OPER_LEASES",             "bs"),
    ("Cash",            "IQ_CASH_ST_INVEST",                "bs"),
    ("TEV",             "IQ_TEV_EXCL_OPER_LEASES",          "market"),
    ("LTM Revenue",     "IQ_TOTAL_REV",                     "ltm"),
    ("NTM Revenue",     "SP_REV_EST",                       "ntm"),
    ("LTM EBITDA",      "IQ_EBITDA_EQ_INC_EXCL_OPER_LEASE_ADJ", "ltm"),
    ("NTM EBITDA",      "SP_EBITDA_EST",                    "ntm"),
    ("P/BV",            "IQ_PBV_X",                         "ltm_raw"),
    ("Gross Margin",    "IQ_GROSS_MARGIN",                  "ltm_raw"),
    ("CapEx",           "IQ_CAPEX",                         "ltm"),
]


# ── Extra (optional) metrics ─────────────────────────────────────────────

@dataclass(frozen=True)
class ExtraMetricSpec:
    key: str            # CLI flag value, e.g. "roe"
    label: str          # Display label / DataFrame column name
    mnemonic: str       # CIQ mnemonic
    call_type: str      # "ltm", "ltm_raw", "bs", "market", "ntm"
    fmt: str            # "pct", "dollar", "mult"
    is_summary: bool    # Include in avg/median row
    transform: str | None = None  # "abs" for cash-outflow items


def _extra(key, label, mnemonic, call_type, fmt, is_summary, transform=None):
    return ExtraMetricSpec(key, label, mnemonic, call_type, fmt, is_summary, transform)


EXTRA_METRICS_CATALOG: dict[str, ExtraMetricSpec] = {s.key: s for s in [
    # ── Profitability ──
    _extra("roe",          "ROE %",              "IQ_RETURN_EQUITY",          "ltm_raw", "pct",    True),
    _extra("roa",          "ROA %",              "IQ_RETURN_ASSETS",          "ltm_raw", "pct",    True),
    _extra("ni-margin",    "NI Margin %",        "IQ_NI_MARGIN",             "ltm_raw", "pct",    True),
    _extra("sga-margin",   "SGA % Rev",          "IQ_SGA_MARGIN",            "ltm_raw", "pct",    True),
    _extra("fcf-margin",   "FCF Margin %",       "IQ_FCF_MARGIN",            "ltm_raw", "pct",    True),
    # ── Dollar amounts ──
    _extra("gross-profit", "Gross Profit",       "IQ_GP",                    "ltm",     "dollar", False),
    _extra("ebit",         "EBIT",               "IQ_EBIT",                  "ltm",     "dollar", False),
    _extra("net-income",   "Net Income",         "IQ_NI",                    "ltm",     "dollar", False),
    _extra("levered-fcf",  "Levered FCF",        "IQ_LEVERED_FCF",           "ltm",     "dollar", False),
    _extra("unlevered-fcf","Unlevered FCF",      "IQ_UNLEVERED_FCF",         "ltm",     "dollar", False),
    _extra("cash-from-ops","Cash from Ops",      "IQ_CASH_OPER",             "ltm",     "dollar", False),
    _extra("da",           "D&A",                "IQ_DA",                    "ltm",     "dollar", False),
    # ── Growth ──
    _extra("rev-growth",   "Rev Growth %",       "IQ_TOTAL_REV_1YR_ANN_GROWTH", "ltm_raw", "pct", True),
    _extra("ebitda-growth","EBITDA Growth %",     "IQ_EBITDA_1YR_ANN_GROWTH","ltm_raw", "pct",    True),
    _extra("ni-growth",    "NI Growth %",        "IQ_NI_1YR_ANN_GROWTH",    "ltm_raw", "pct",    True),
    _extra("eps-growth",   "EPS Growth %",       "IQ_DILUT_EPS_NORM_1YR_ANN_GROWTH", "ltm_raw", "pct", True),
    # ── Multiples ──
    _extra("pe-ltm",       "P/E LTM",            "IQ_PE_EXCL",              "ltm_raw", "mult",   True),
    _extra("pe-ntm",       "P/E NTM",            "SP_PE_NTM",               "ntm",     "mult",   True),
    _extra("ev-ebit",      "EV/EBIT",            "IQ_TEV_EBIT",             "ltm_raw", "mult",   True),
    _extra("div-yield",    "Div Yield %",        "IQ_DIV_YIELD",            "ltm_raw", "pct",    True),
    # ── Other ──
    _extra("capex-pct-rev","CapEx % Rev",        "IQ_CAPEX_PCT_REV",        "ltm_raw", "pct",    True),
    _extra("current-ratio","Current Ratio",      "IQ_CURRENT_RATIO",        "bs",      "mult",   True),
    _extra("debt-equity",  "Debt/Equity",        "IQ_TOTAL_DEBT_EQUITY",    "ltm_raw", "mult",   True),
    _extra("employees",    "Employees",          "IQ_EMPLOYEES",            "bs",      "dollar", False),
]}


def resolve_extras(keys: list[str]) -> tuple[list[ExtraMetricSpec], list[tuple]]:
    """Validate extra metric keys and return specs + metric tuples to append."""
    specs = []
    metric_tuples = []
    for key in keys:
        spec = EXTRA_METRICS_CATALOG.get(key)
        if not spec:
            available = ", ".join(sorted(EXTRA_METRICS_CATALOG.keys()))
            raise ValueError(f"Unknown extra metric '{key}'. Available: {available}")
        specs.append(spec)
        metric_tuples.append((spec.label, spec.mnemonic, spec.call_type))
    return specs, metric_tuples


# ── Formula builders ───────────────────────────────────────────────────────

def _build_spg_formula(id_expr: str, mnemonic: str, call_type: str,
                       date: str, curr_opt: str) -> str:
    """Build an SPG formula string for a given metric type.

    Parameters
    ----------
    id_expr : str
        Either a quoted ticker like ``'"DSGX"'`` or a cell reference like ``'$P$2'``.
    """
    m = f'"{mnemonic}"'
    d = f'"{date}"'
    c = f'"{curr_opt}"'

    if call_type == "name":
        return f"=SPG({id_expr}, {m})"
    elif call_type == "market":
        return f"=SPG({id_expr}, {m}, {d}, {c})"
    elif call_type == "bs":
        return f'=SPG({id_expr}, {m}, "FQ0", {d}, {c})'
    elif call_type == "ltm":
        return f'=SPG({id_expr}, {m}, "LTM", {d}, {c})'
    elif call_type == "ntm":
        return f'=SPG({id_expr}, {m}, "NTM", {d}, {c})'
    elif call_type == "ltm_raw":
        return f'=SPG({id_expr}, {m}, "LTM", {d})'
    elif call_type == "gaap":
        return f'=CIQ({id_expr}, {m})'
    else:
        raise ValueError(f"Unknown call_type: {call_type}")


def build_spg_formula(ticker: str, mnemonic: str, call_type: str,
                      date: str, curr_opt: str) -> str:
    """Build an SPG formula string with a quoted ticker identifier."""
    return _build_spg_formula(f'"{ticker}"', mnemonic, call_type, date, curr_opt)


# ── Workbook creation ──────────────────────────────────────────────────────

def build_workbook(path: str, tickers: list[str], metrics: list,
                   date: str, curr_opt: str) -> None:
    """Build workbook with primary SPG formulas + CIQRANGEA fallback.

    Layout per company (two rows):
      Row N (primary):  ticker in col A, SPG formulas using original ticker
      Row N+1 (fallback): SPG formulas using CIQRANGEA-resolved IQ ID via cell ref

    CIQRANGEA columns sit after the metrics:
      col M+2: =CIQRANGEA(ticker, "IQ_COMPANY_ID_QUICK_MATCH", 1, 1)
      col M+3: (blank — CIQRANGEA spill result = resolved IQ ID)
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Comp"

    num_metrics = len(metrics)

    # Headers
    ws.cell(row=1, column=1, value="Ticker")
    for col_idx, (label, _, _) in enumerate(metrics, start=2):
        ws.cell(row=1, column=col_idx, value=label)

    # CIQRANGEA columns after metrics
    ciqrangea_col = num_metrics + 2
    ciqrangea_spill_col = ciqrangea_col + 1
    ws.cell(row=1, column=ciqrangea_col, value="CIQRANGEA_Formula")
    ws.cell(row=1, column=ciqrangea_spill_col, value="Resolved_IQ_ID")

    spill_col_letter = get_column_letter(ciqrangea_spill_col)

    for comp_idx, ticker in enumerate(tickers):
        # Strip US exchange prefixes — SPG validates them against the actual
        # listing exchange, so NYSE:WMT fails because WMT is on NASDAQGS.
        # Plain tickers resolve fine for US stocks.  Non-US prefixes (TSX,
        # LSE, ASX, etc.) are kept because plain tickers don't resolve
        # reliably for international listings.
        prefix, _, sym = ticker.partition(":")
        if sym and prefix.upper() in _US_EXCHANGES:
            plain_ticker = sym
        else:
            plain_ticker = ticker
        primary_row = 2 + comp_idx * 2
        ws.cell(row=primary_row, column=1, value=plain_ticker)
        for col_idx, (_, mnemonic, call_type) in enumerate(metrics, start=2):
            formula = build_spg_formula(plain_ticker, mnemonic, call_type,
                                        date, curr_opt)
            ws.cell(row=primary_row, column=col_idx, value=formula)

        # CIQRANGEA lookup on primary row (also uses plain ticker)
        lookup_id = plain_ticker
        ws.cell(row=primary_row, column=ciqrangea_col,
                value=f'=CIQRANGEA("{lookup_id}","IQ_COMPANY_ID_QUICK_MATCH",1,1)')

        # Fallback row: SPG formulas referencing the CIQRANGEA spill cell
        fallback_row = primary_row + 1
        spill_ref = f"${spill_col_letter}${primary_row}"
        ws.cell(row=fallback_row, column=1, value=f"(fallback for {ticker})")
        for col_idx, (_, mnemonic, call_type) in enumerate(metrics, start=2):
            formula = _build_spg_formula(spill_ref, mnemonic, call_type,
                                         date, curr_opt)
            ws.cell(row=fallback_row, column=col_idx, value=formula)

    wb.save(path)
    total_formulas = len(tickers) * num_metrics * 2
    log.info("Created %s (%d companies x %d metrics, %d formulas + %d CIQRANGEA lookups)",
             path, len(tickers), num_metrics, total_formulas, len(tickers))


# ── Value helpers ──────────────────────────────────────────────────────────

def safe_float(val) -> float:
    """Convert a CIQ cell value to float, returning NaN for errors."""
    if val is None:
        return np.nan
    if isinstance(val, (int, float)):
        if val < -2000000000:
            return np.nan
        return float(val)
    if isinstance(val, str):
        upper = val.upper()
        if any(tok in upper for tok in _ERROR_TOKENS):
            return np.nan
        try:
            return float(val.replace(",", ""))
        except ValueError:
            return np.nan
    return np.nan


def is_error_value(val) -> bool:
    """Check if a cell value is a CIQ/SPG error."""
    if val is None:
        return True
    if isinstance(val, (int, float)) and val < -2000000000:
        return True
    if isinstance(val, str):
        upper = val.upper()
        return any(tok in upper for tok in _ERROR_TOKENS)
    return False


def safe_div(num: float, denom: float) -> float:
    """Safe division returning NaN for invalid or non-positive results."""
    if np.isnan(num) or np.isnan(denom) or denom == 0:
        return np.nan
    result = num / denom
    return result if result > 0 else np.nan


# ── Result reading ─────────────────────────────────────────────────────────

def read_results(ws_com, num_companies: int, metrics: list) -> tuple[list[dict], list[dict]]:
    """Read results with automatic fallback to CIQRANGEA-resolved IQ IDs.

    Returns (results, resolution_notes).
    """
    num_metrics = len(metrics)
    last_data_col = num_metrics + 1
    ciqrangea_spill_col = num_metrics + 3

    # Read the entire data region at once (batch COM read)
    total_rows = num_companies * 2
    last_row = 1 + total_rows
    data_range = ws_com.Range(
        ws_com.Cells(2, 1),
        ws_com.Cells(last_row, last_data_col)
    ).Value

    # Batch-read CIQRANGEA spill column for all primary rows (single COM call)
    if num_companies == 1:
        id_values = [ws_com.Cells(2, ciqrangea_spill_col).Value]
    else:
        all_id_vals = ws_com.Range(
            ws_com.Cells(2, ciqrangea_spill_col),
            ws_com.Cells(last_row, ciqrangea_spill_col),
        ).Value
        id_values = [all_id_vals[comp_idx * 2][0] for comp_idx in range(num_companies)]

    results = []
    resolution_notes = []

    for comp_idx in range(num_companies):
        primary_idx = comp_idx * 2
        fallback_idx = primary_idx + 1

        primary_vals = data_range[primary_idx]
        fallback_vals = data_range[fallback_idx]
        ticker = primary_vals[0]
        resolved_id = id_values[comp_idx]

        # Check if primary SPG resolved
        primary_name = primary_vals[1]  # Company Name is first metric
        primary_ok = not is_error_value(primary_name)

        if primary_ok:
            row_data = {"_ticker": ticker}
            for met_idx, (label, _, _) in enumerate(metrics):
                row_data[label] = primary_vals[met_idx + 1]
            results.append(row_data)
            resolution_notes.append({
                "ticker": ticker, "status": "OK", "used": "primary"
            })
        else:
            fallback_name = fallback_vals[1]
            fallback_ok = not is_error_value(fallback_name)

            if fallback_ok:
                row_data = {"_ticker": ticker}
                for met_idx, (label, _, _) in enumerate(metrics):
                    row_data[label] = fallback_vals[met_idx + 1]
                results.append(row_data)
                iq_id = str(resolved_id) if resolved_id and not is_error_value(resolved_id) else "?"
                resolution_notes.append({
                    "ticker": ticker, "status": "RESOLVED",
                    "used": "fallback", "resolved_iq_id": iq_id
                })
            else:
                row_data = {"_ticker": ticker}
                for met_idx, (label, _, _) in enumerate(metrics):
                    row_data[label] = primary_vals[met_idx + 1]
                results.append(row_data)
                resolution_notes.append({
                    "ticker": ticker, "status": "FAILED", "used": "primary"
                })

    return results, resolution_notes


# ── Table builders ─────────────────────────────────────────────────────────

def _append_extras(row: dict, d: dict, extra_specs: list[ExtraMetricSpec]) -> None:
    """Append extra metric values to a row dict (mutates row in place)."""
    for spec in extra_specs:
        raw = safe_float(d.get(spec.label))
        if spec.transform == "abs":
            raw = abs(raw) if not np.isnan(raw) else np.nan
        row[spec.label] = raw


def build_table_lease_adjusted(results: list[dict], lease_adjust_ntm: bool,
                               extra_specs: list[ExtraMetricSpec] | None = None) -> pd.DataFrame:
    """Build the lease-adjusted comp table DataFrame."""
    extra_specs = extra_specs or []
    rows = []
    for d in results:
        ticker = d.get("_ticker", "")
        name = d.get("Company Name", ticker)
        if not isinstance(name, str) or len(name) < 2:
            name = ticker

        mkt_cap = safe_float(d.get("Market Cap"))
        total_debt = safe_float(d.get("Total Debt"))
        oper_leases = safe_float(d.get("Oper. Leases"))
        cash = safe_float(d.get("Cash"))
        net_debt = safe_float(d.get("Net Debt"))
        tev = safe_float(d.get("TEV"))
        ltm_rev = safe_float(d.get("LTM Revenue"))
        ntm_rev = safe_float(d.get("NTM Revenue"))
        ltm_ebitda = safe_float(d.get("LTM EBITDA"))
        lease_adj = safe_float(d.get("Lease Adj"))
        ntm_ebitda_raw = safe_float(d.get("NTM EBITDA"))
        gaap = d.get("GAAP", "")
        pbv = safe_float(d.get("P/BV"))

        # Lease-adjusted NTM EBITDA
        ntm_ebitda = ntm_ebitda_raw
        adj_applied = False
        if lease_adjust_ntm and not np.isnan(ntm_ebitda_raw) and not np.isnan(lease_adj):
            if isinstance(gaap, str) and "IFRS" not in gaap.upper():
                ntm_ebitda = ntm_ebitda_raw + lease_adj
                adj_applied = True

        gross_margin_raw = safe_float(d.get("Gross Margin"))
        capex_raw = safe_float(d.get("CapEx"))

        # IQ_GROSS_MARGIN returns a percentage directly (e.g. 35.0 for 35%)
        gross_margin_pct = gross_margin_raw if not np.isnan(gross_margin_raw) else np.nan
        # IQ_CAPEX returns a negative number (cash outflow); take absolute value
        capex = abs(capex_raw) if not np.isnan(capex_raw) else np.nan

        ebitda_margin = safe_div(ltm_ebitda, ltm_rev)
        nd_ebitda = safe_div(net_debt, ltm_ebitda)

        row = {
            "Company": name,
            "Ticker": ticker,
            "GAAP": gaap if isinstance(gaap, str) else "N/A",
            "Mkt Cap": mkt_cap,
            "Total Debt": total_debt,
            "Leases": oper_leases,
            "Cash": cash,
            "Net Debt": net_debt,
            "TEV": tev,
            "LTM Rev": ltm_rev,
            "NTM Rev": ntm_rev,
            "LTM EBITDA": ltm_ebitda,
            "Lease Adj": lease_adj,
            "NTM EBITDA": ntm_ebitda,
            "Lease Adj Applied": adj_applied,
            "EBITDA Margin %": ebitda_margin * 100 if not np.isnan(ebitda_margin) else np.nan,
            "Gross Margin %": gross_margin_pct,
            "CapEx": capex,
            "EV/Rev": safe_div(tev, ltm_rev),
            "EV/LTM EBITDA": safe_div(tev, ltm_ebitda),
            "EV/NTM EBITDA": safe_div(tev, ntm_ebitda),
            "P/BV": pbv,
            "ND/EBITDA": nd_ebitda,
        }
        _append_extras(row, d, extra_specs)
        rows.append(row)
    return pd.DataFrame(rows)


def build_table_excluding_leases(results: list[dict],
                                 extra_specs: list[ExtraMetricSpec] | None = None) -> pd.DataFrame:
    """Build the excluding-leases comp table DataFrame."""
    extra_specs = extra_specs or []
    rows = []
    for d in results:
        ticker = d.get("_ticker", "")
        name = d.get("Company Name", ticker)
        if not isinstance(name, str) or len(name) < 2:
            name = ticker

        mkt_cap = safe_float(d.get("Market Cap"))
        total_debt = safe_float(d.get("Total Debt"))
        oper_leases = safe_float(d.get("Oper. Leases"))
        cash = safe_float(d.get("Cash"))
        net_debt = total_debt - cash if not np.isnan(total_debt) and not np.isnan(cash) else np.nan
        tev = safe_float(d.get("TEV"))
        ltm_rev = safe_float(d.get("LTM Revenue"))
        ntm_rev = safe_float(d.get("NTM Revenue"))
        ltm_ebitda = safe_float(d.get("LTM EBITDA"))
        ntm_ebitda = safe_float(d.get("NTM EBITDA"))
        pbv = safe_float(d.get("P/BV"))

        gross_margin_raw = safe_float(d.get("Gross Margin"))
        capex_raw = safe_float(d.get("CapEx"))

        # IQ_GROSS_MARGIN returns a percentage directly (e.g. 35.0 for 35%)
        gross_margin_pct = gross_margin_raw if not np.isnan(gross_margin_raw) else np.nan
        capex = abs(capex_raw) if not np.isnan(capex_raw) else np.nan

        ebitda_margin = safe_div(ltm_ebitda, ltm_rev)
        nd_ebitda = safe_div(net_debt, ltm_ebitda)

        row = {
            "Company": name,
            "Ticker": ticker,
            "Mkt Cap": mkt_cap,
            "Total Debt": total_debt,
            "Leases": oper_leases,
            "Cash": cash,
            "Net Debt": net_debt,
            "TEV": tev,
            "LTM Rev": ltm_rev,
            "NTM Rev": ntm_rev,
            "LTM EBITDA": ltm_ebitda,
            "NTM EBITDA": ntm_ebitda,
            "EBITDA Margin %": ebitda_margin * 100 if not np.isnan(ebitda_margin) else np.nan,
            "Gross Margin %": gross_margin_pct,
            "CapEx": capex,
            "EV/Rev": safe_div(tev, ltm_rev),
            "EV/LTM EBITDA": safe_div(tev, ltm_ebitda),
            "EV/NTM EBITDA": safe_div(tev, ntm_ebitda),
            "P/BV": pbv,
            "ND/EBITDA": nd_ebitda,
        }
        _append_extras(row, d, extra_specs)
        rows.append(row)
    return pd.DataFrame(rows)


# ── JSON serialization ─────────────────────────────────────────────────────

# Precision mapping for numeric columns
_ROUND_0 = frozenset({
    "Mkt Cap", "Total Debt", "Leases", "Cash", "Net Debt", "TEV",
    "LTM Rev", "NTM Rev", "LTM EBITDA", "Lease Adj", "NTM EBITDA", "CapEx",
})
_ROUND_1 = frozenset({"EBITDA Margin %", "Gross Margin %"})
_ROUND_2 = frozenset({"EV/Rev", "EV/LTM EBITDA", "EV/NTM EBITDA", "P/BV", "ND/EBITDA"})


def _make_round_fn(extra_specs: list[ExtraMetricSpec] | None = None):
    """Return a rounding function that handles base + extra columns."""
    # Build local sets that include extras
    round_0 = set(_ROUND_0)
    round_1 = set(_ROUND_1)
    round_2 = set(_ROUND_2)
    for spec in (extra_specs or []):
        if spec.fmt == "dollar":
            round_0.add(spec.label)
        elif spec.fmt == "pct":
            round_1.add(spec.label)
        elif spec.fmt == "mult":
            round_2.add(spec.label)

    def _round_value(col: str, val: float) -> float | int:
        if col in round_0:
            return round(val)
        if col in round_1:
            return round(val, 1)
        if col in round_2:
            return round(val, 2)
        return val

    return _round_value


def dataframe_to_json(df: pd.DataFrame, mode: str, currency: str,
                      date: str, lease_adjust_ntm: bool,
                      extra_specs: list[ExtraMetricSpec] | None = None) -> dict:
    """Convert a comp table DataFrame to a structured dict.

    Applies intelligent rounding and strips None/NaN values to minimize tokens.
    """
    round_value = _make_round_fn(extra_specs)

    output = {
        "mode": mode,
        "currency": currency,
        "date": date,
    }
    if mode == "lease-adjusted":
        output["lease_adjust_ntm"] = lease_adjust_ntm

    companies = []
    for _, row in df.iterrows():
        company = {}
        for col in df.columns:
            val = row[col]
            if isinstance(val, float) and np.isnan(val):
                continue  # strip None/NaN values
            elif isinstance(val, (np.floating, np.integer)):
                company[col] = round_value(col, float(val))
            elif isinstance(val, float):
                company[col] = round_value(col, val)
            else:
                company[col] = val
        companies.append(company)
    output["companies"] = companies
    return output


# ── Polling ────────────────────────────────────────────────────────────────

def poll_for_completion(ws_com, num_companies: int, num_metrics: int,
                        max_wait: float) -> None:
    """Poll formula cells until data is ready or timeout.

    Uses a single batch COM read per cycle instead of cell-by-cell access.
    """
    num_formulas = num_companies * num_metrics * 2  # primary + fallback
    log.info("Waiting for %d formulas (max %ds)...", num_formulas, max_wait)
    start = time.monotonic()

    total_rows = num_companies * 2
    first_row = 2
    last_row = 1 + total_rows
    first_col = 2
    last_col = num_metrics + 1

    # Primary rows are odd-indexed (rows 2, 4, 6, ...) — these have the real data.
    # Fallback rows are even-indexed (rows 3, 5, 7, ...) — CIQRANGEA lookups that
    # may stay #PEND indefinitely.  We exit as soon as primary rows are resolved,
    # rather than waiting for fallback rows that may never settle.
    primary_row_indices = list(range(0, total_rows, 2))      # 0, 2, 4, ...
    num_primary_cells = num_companies * num_metrics

    while True:
        elapsed = time.monotonic() - start
        if elapsed > max_wait:
            log.warning("Timeout after %ds — proceeding with available data", max_wait)
            break

        # Single batch COM read for all formula cells
        data = ws_com.Range(
            ws_com.Cells(first_row, first_col),
            ws_com.Cells(last_row, last_col),
        ).Value

        # Check primary rows only for exit condition
        primary_pending = False
        primary_resolved = 0
        for i in primary_row_indices:
            for val in data[i]:
                if isinstance(val, str) and val.upper() in ('#PEND', '#REFRESH'):
                    primary_pending = True
                elif val is not None and not (isinstance(val, (int, float)) and val < -2000000000):
                    primary_resolved += 1

        if not primary_pending and primary_resolved > num_primary_cells * 0.5:
            log.info("Data ready: %d/%d primary cells resolved in %.0fs",
                     primary_resolved, num_primary_cells, elapsed)
            break

        if int(elapsed) % 15 < 5:
            log.debug("%ds: %d/%d primary resolved, pending=%s",
                      elapsed, primary_resolved, num_primary_cells, primary_pending)

        # Pump COM message queue to prevent STA deadlocks
        pythoncom.PumpWaitingMessages()
        time.sleep(5)


# ── Main pipeline ──────────────────────────────────────────────────────────

def run_comps(
    tickers: list[str],
    currency: str = "CAD",
    mode: str = "lease-adjusted",
    lease_adjust_ntm: bool = True,
    date: str | None = None,
    max_wait: int = 180,
    extras: list[str] | None = None,
) -> dict:
    """Full pipeline: build workbook -> launch Excel -> refresh -> read -> return JSON.

    Parameters
    ----------
    tickers : list[str]
        Company identifiers (EXCHANGE:TICKER format preferred).
    currency : str
        Output currency (default CAD).
    mode : str
        "lease-adjusted" or "excluding-leases".
    lease_adjust_ntm : bool
        Whether to add TTM lease adj to NTM EBITDA for US GAAP reporters.
    date : str or None
        As-of date in M/D/YYYY format. Defaults to today.
    max_wait : int
        Max seconds to wait for formula refresh.
    extras : list[str] or None
        Optional extra metric keys (e.g. ["roe", "rev-growth"]).

    Returns
    -------
    dict
        Structured result with companies array and resolution_notes.
    """
    pythoncom.CoInitialize()

    from capiq_excel.excel_lifecycle import launch_excel_isolated, close_session

    # Resolve date
    if date is None:
        now = datetime.now()
        date = f"{now.month}/{now.day}/{now.year}"

    curr_opt = f"Options: Curr={currency.upper()}"

    # Select metrics
    metrics = METRICS_LEASE_ADJUSTED if mode == "lease-adjusted" else METRICS_EXCLUDING_LEASES

    # Resolve extras and append to metrics list
    extra_specs = []
    if extras:
        extra_specs, extra_tuples = resolve_extras(extras)
        metrics = list(metrics) + extra_tuples

    xlsx_path = os.path.abspath("_comps_mcp_temp.xlsx")

    log.info("Comp pull: %d tickers, mode=%s, currency=%s, date=%s, extras=%s",
             len(tickers), mode, currency, date,
             [s.key for s in extra_specs] if extra_specs else "none")

    # Build workbook
    build_workbook(xlsx_path, tickers, metrics, date, curr_opt)

    # Launch Excel and process
    session = launch_excel_isolated(xlsx_path)
    try:
        excel = session.excel
        log.info("Connected to Excel %s", excel.Version)
        ws_com = session.workbook.Sheets(1)

        # Trigger refresh
        log.info("Calling RefreshSheet...")
        try:
            pythoncom.PumpWaitingMessages()
            excel.Run("SNLXLAddin.xla!RefreshSheet")
            log.info("RefreshSheet executed")
        except Exception as e:
            log.warning("RefreshSheet warning: %s", e)
        pythoncom.PumpWaitingMessages()

        # Poll for completion
        poll_for_completion(ws_com, len(tickers), len(metrics), max_wait)

        # Read results
        log.info("Reading data...")
        results, resolution_notes = read_results(ws_com, len(tickers), metrics)
    finally:
        close_session(session, save=False, delete_workbook=True)

    # Build table
    if mode == "lease-adjusted":
        df = build_table_lease_adjusted(results, lease_adjust_ntm, extra_specs)
    else:
        df = build_table_excluding_leases(results, extra_specs)

    # Convert to JSON-serializable dict
    data = dataframe_to_json(df, mode, currency, date, lease_adjust_ntm, extra_specs)

    # Merge resolution into company dicts (only non-OK statuses)
    for note, company in zip(resolution_notes, data["companies"]):
        status = note.get("status", "OK")
        if status != "OK":
            company["_resolution"] = status
            if note.get("resolved_iq_id"):
                company["_resolved_iq_id"] = note["resolved_iq_id"]

    log.info("Comp pull complete: %d companies", len(data["companies"]))
    return data
