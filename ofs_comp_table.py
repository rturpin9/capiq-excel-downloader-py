"""
Oilfield Services Comparable Companies Table — Lease Adjusted.

Replicates the Lease Adjusted tab logic from "Comp Set - New.xlsx":
  - SPG() formulas for fundamentals + estimates
  - CIQ() for GAAP standard identification
  - NTM EBITDA adjusted: adds TTM lease adjustment for US GAAP reporters
    to normalize against IFRS reporters (where leases are below EBITDA line)
  - All values in CAD (matching template)
"""
import time
import os
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import Workbook

import pythoncom
from capiq_excel.excel_lifecycle import launch_excel_isolated, close_session

pythoncom.CoInitialize()

# -- Configuration ----------------------------------------------------------
DATE = "2/28/2026"
CURR_OPT = "Options: Curr=CAD"
LEASE_ADJUST_NTM = True  # Add TTM lease adj to NTM EBITDA for US GAAP reporters

# -- Companies --------------------------------------------------------------
US_COMPANIES = [
    "NYSE:HAL",       # Halliburton
    "NYSE:SLB",       # SLB (Schlumberger)
    "NASDAQGS:BKR",   # Baker Hughes
    "NYSE:NOV",       # NOV Inc
    "NYSE:LBRT",      # Liberty Energy
]

CA_COMPANIES = [
    "TSX:PD",         # Precision Drilling
    "TSX:TCW",        # Trican Well Service
    "TSX:CEU",        # CES Energy Solutions
    "TSX:CFW",        # Calfrac Well Services
    "TSX:STEP",       # STEP Energy Services
]

ALL_TICKERS = US_COMPANIES + CA_COMPANIES

# -- SPG/CIQ Metrics to pull -----------------------------------------------
# (column_label, mnemonic, spg_call_type)
# spg_call_type:
#   "name"    -> SPG(ticker, mnemonic)
#   "market"  -> SPG(ticker, mnemonic, date, curr)        [point-in-time]
#   "bs"      -> SPG(ticker, mnemonic, "FQ0", date, curr) [balance sheet]
#   "ltm"     -> SPG(ticker, mnemonic, "LTM", date, curr) [trailing]
#   "ntm"     -> SPG(ticker, mnemonic, "NTM", date, curr) [forward est]
#   "ltm_raw" -> SPG(ticker, mnemonic, "LTM", date)       [no currency]
#   "gaap"    -> CIQ(ticker, mnemonic)                     [legacy CIQ]

METRICS = [
    ("Company Name",    "SP_COMPANY_NAME",               "name"),
    ("Market Cap",      "SP_MARKETCAP",                  "market"),
    ("Total Debt",      "IQ_TOTAL_DEBT",                 "bs"),
    ("Oper. Leases",    "IQ_TOTAL_OPER_LEASES",          "bs"),
    ("Cash",            "IQ_CASH_ST_INVEST",             "bs"),
    ("Net Debt",        "IQ_NET_DEBT",                   "bs"),
    ("TEV",             "IQ_TEV",                        "market"),
    ("LTM Revenue",     "IQ_TOTAL_REV",                  "ltm"),
    ("NTM Revenue",     "SP_REV_EST",                    "ntm"),
    ("LTM EBITDA",      "IQ_EBITDA_EQ_INC",             "ltm"),
    ("Lease Adj",       "IQ_LEASE_ADJUSTMENT_EBITDA",    "ltm"),
    ("NTM EBITDA",      "SP_EBITDA_EST",                 "ntm"),
    ("GAAP",            "IQ_GAAP_BS",                    "gaap"),
    ("P/BV",            "IQ_PBV_X",                      "ltm_raw"),
    ("ND/EBITDA",       "IQ_NET_DEBT_EBITDA",            "ltm_raw"),
]


def build_spg_formula(ticker: str, mnemonic: str, call_type: str) -> str:
    """Build an SPG or CIQ formula string matching the comp set template."""
    t = f'"{ticker}"'
    m = f'"{mnemonic}"'
    d = f'"{DATE}"'
    c = f'"{CURR_OPT}"'

    if call_type == "name":
        return f"=SPG({t}, {m})"
    elif call_type == "market":
        return f"=SPG({t}, {m}, {d}, {c})"
    elif call_type == "bs":
        return f'=SPG({t}, {m}, "FQ0", {d}, {c})'
    elif call_type == "ltm":
        return f'=SPG({t}, {m}, "LTM", {d}, {c})'
    elif call_type == "ntm":
        return f'=SPG({t}, {m}, "NTM", {d}, {c})'
    elif call_type == "ltm_raw":
        return f'=SPG({t}, {m}, "LTM", {d})'
    elif call_type == "gaap":
        return f'=CIQ({t}, {m})'
    else:
        raise ValueError(f"Unknown call_type: {call_type}")


def build_workbook(path: str):
    """Create XLSX with SPG/CIQ formulas for all companies and metrics."""
    wb = Workbook()
    ws = wb.active
    ws.title = "OFS Comp"

    # Row 1: headers
    ws.cell(row=1, column=1, value="Ticker")
    for col_idx, (label, _, _) in enumerate(METRICS, start=2):
        ws.cell(row=1, column=col_idx, value=label)

    # Rows 2+: one per company
    for row_idx, ticker in enumerate(ALL_TICKERS, start=2):
        ws.cell(row=row_idx, column=1, value=ticker)
        for col_idx, (_, mnemonic, call_type) in enumerate(METRICS, start=2):
            formula = build_spg_formula(ticker, mnemonic, call_type)
            ws.cell(row=row_idx, column=col_idx, value=formula)

    wb.save(path)
    print(f"  Created: {path}")
    print(f"  {len(ALL_TICKERS)} companies x {len(METRICS)} metrics = {len(ALL_TICKERS) * len(METRICS)} formulas")


def safe_float(val):
    """Convert a cell value to float, returning NaN for errors."""
    if val is None:
        return np.nan
    if isinstance(val, (int, float)):
        if val < -2000000000:  # COM error code
            return np.nan
        return float(val)
    if isinstance(val, str):
        # Check for CIQ error tokens
        upper = val.upper()
        error_tokens = ("#ERROR", "#INVALID", "#PEND", "#REFRESH", "#NAME",
                        "#OUTSIDE", "KEYERROR", "DEFUNCT", "INVALID", "NM")
        if any(tok in upper for tok in error_tokens):
            return np.nan
        try:
            return float(val.replace(",", ""))
        except ValueError:
            return np.nan
    return np.nan


def read_results(ws_com, num_companies: int, num_metrics: int) -> list[dict]:
    """Read all formula results from the COM worksheet."""
    results = []

    # Batch read: grab the entire data range at once
    last_row = num_companies + 1
    last_col = num_metrics + 1
    data_range = ws_com.Range(
        ws_com.Cells(2, 1),
        ws_com.Cells(last_row, last_col)
    ).Value  # returns tuple of tuples

    for row_idx in range(num_companies):
        row_data = {}
        row_vals = data_range[row_idx]
        row_data["_ticker"] = row_vals[0]
        for met_idx, (label, _, _) in enumerate(METRICS):
            row_data[label] = row_vals[met_idx + 1]
        results.append(row_data)

    return results


def build_comp_table(results: list[dict]) -> pd.DataFrame:
    """Build the lease-adjusted comp table with derived multiples."""
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
        nd_ebitda = safe_float(d.get("ND/EBITDA"))

        # Lease-adjusted NTM EBITDA: add TTM lease adj for US GAAP reporters
        ntm_ebitda = ntm_ebitda_raw
        lease_adj_applied = False
        if LEASE_ADJUST_NTM and not np.isnan(ntm_ebitda_raw) and not np.isnan(lease_adj):
            if isinstance(gaap, str) and "IFRS" not in gaap.upper():
                ntm_ebitda = ntm_ebitda_raw + lease_adj
                lease_adj_applied = True

        # Derived multiples
        def safe_div(num, denom):
            if np.isnan(num) or np.isnan(denom) or denom == 0:
                return np.nan
            result = num / denom
            return result if result > 0 else np.nan

        ebitda_margin = safe_div(ltm_ebitda, ltm_rev) * 100 if not np.isnan(safe_div(ltm_ebitda, ltm_rev)) else np.nan
        ev_rev = safe_div(tev, ltm_rev)
        ev_ltm_ebitda = safe_div(tev, ltm_ebitda)
        ev_ntm_ebitda = safe_div(tev, ntm_ebitda)

        rows.append({
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
            "Lease Adj Applied": lease_adj_applied,
            "EBITDA Margin %": ebitda_margin,
            "EV/Rev": ev_rev,
            "EV/LTM EBITDA": ev_ltm_ebitda,
            "EV/NTM EBITDA": ev_ntm_ebitda,
            "P/BV": pbv,
            "ND/EBITDA": nd_ebitda,
        })

    return pd.DataFrame(rows)


def format_table(df: pd.DataFrame):
    """Print a formatted lease-adjusted comp table."""
    print("\n" + "=" * 160)
    print(f"  OILFIELD SERVICES — COMPARABLE COMPANIES ANALYSIS (LEASE ADJUSTED)")
    print(f"  As at {DATE} (Millions $CAD)")
    print(f"  NTM EBITDA lease-adjusted for US GAAP reporters (TTM lease adj added as proxy)")
    print("=" * 160)

    # Format specs
    money_fmt = "${:,.0f}"
    mult_fmt = "{:.1f}x"
    pct_fmt = "{:.1f}%"

    # Column definitions for display
    display_cols = [
        ("Company",         35, None,      "left"),
        ("GAAP",             7, None,      "left"),
        ("Mkt Cap",         10, money_fmt, "right"),
        ("Total Debt",      10, money_fmt, "right"),
        ("Leases",           8, money_fmt, "right"),
        ("Cash",            10, money_fmt, "right"),
        ("Net Debt",        10, money_fmt, "right"),
        ("TEV",             10, money_fmt, "right"),
        ("LTM Rev",         10, money_fmt, "right"),
        ("NTM Rev",         10, money_fmt, "right"),
        ("LTM EBITDA",      10, money_fmt, "right"),
        ("Lease Adj",        9, money_fmt, "right"),
        ("NTM EBITDA",      10, money_fmt, "right"),
        ("EBITDA Margin %",  8, pct_fmt,   "right"),
        ("EV/Rev",           7, mult_fmt,  "right"),
        ("EV/LTM EBITDA",    7, mult_fmt,  "right"),
        ("EV/NTM EBITDA",    7, mult_fmt,  "right"),
        ("P/BV",             6, mult_fmt,  "right"),
        ("ND/EBITDA",        6, mult_fmt,  "right"),
    ]

    def fmt_val(val, fmt_str, width, align):
        if isinstance(val, str):
            if align == "left":
                return f"{val:<{width}}"
            return f"{val:>{width}}"
        if pd.isna(val):
            return f"{'N/A':>{width}}" if align == "right" else f"{'N/A':<{width}}"
        if fmt_str:
            try:
                formatted = fmt_str.format(val)
                return f"{formatted:>{width}}" if align == "right" else f"{formatted:<{width}}"
            except (ValueError, TypeError):
                return f"{'N/A':>{width}}"
        return f"{str(val):>{width}}"

    # Print sections
    def print_section(section_name, section_df):
        print(f"\n  {section_name}")
        print("  " + "-" * 155)

        # Header
        header = "  "
        for col_name, width, _, align in display_cols:
            short = col_name.replace("EBITDA ", "").replace("Margin ", "Mgn ")
            if align == "left":
                header += f"{short:<{width}}  "
            else:
                header += f"{short:>{width}}  "
        print(header)
        print("  " + "-" * 155)

        # Data rows
        for _, row in section_df.iterrows():
            line = "  "
            for col_name, width, fmt_str, align in display_cols:
                val = row.get(col_name, np.nan)
                # Truncate company name
                if col_name == "Company" and isinstance(val, str):
                    val = val[:width - 2]
                line += fmt_val(val, fmt_str, width, align) + "  "
            # Mark lease-adjusted NTM EBITDA
            if row.get("Lease Adj Applied", False):
                line += " *adj"
            print(line)

        print("  " + "-" * 155)

        # Averages and medians
        numeric_cols = [(c, w, f, a) for c, w, f, a in display_cols
                        if f is not None and c not in ("Company", "GAAP")]

        for stat_name, stat_fn in [("Average", "mean"), ("Median", "median")]:
            line = f"  {stat_name:<35}{'':>7}  "
            for col_name, width, fmt_str, align in display_cols:
                if col_name in ("Company", "GAAP"):
                    continue
                vals = section_df[col_name].apply(lambda x: safe_float(x) if not isinstance(x, (int, float)) else x).dropna()
                if col_name in ("EBITDA Margin %", "EV/Rev", "EV/LTM EBITDA", "EV/NTM EBITDA", "P/BV", "ND/EBITDA"):
                    if len(vals) > 0:
                        s = getattr(vals, stat_fn)()
                        line += fmt_val(s, fmt_str, width, "right") + "  "
                    else:
                        line += fmt_val(np.nan, fmt_str, width, "right") + "  "
                else:
                    line += f"{'':>{width}}  "
            print(line)

    # Split into US and Canadian
    us_df = df.iloc[:len(US_COMPANIES)]
    ca_df = df.iloc[len(US_COMPANIES):]

    print_section("U.S. Oilfield Services", us_df)
    print_section("Canadian Oilfield Services", ca_df)

    # Overall summary
    print(f"\n  Overall Summary")
    print("  " + "-" * 155)
    mult_cols = ["EBITDA Margin %", "EV/Rev", "EV/LTM EBITDA", "EV/NTM EBITDA", "P/BV", "ND/EBITDA"]
    for stat_name, stat_fn in [("Overall Average", "mean"), ("Overall Median", "median")]:
        line = f"  {stat_name:<35}{'':>7}  "
        for col_name, width, fmt_str, align in display_cols:
            if col_name in ("Company", "GAAP"):
                continue
            if col_name in mult_cols:
                vals = df[col_name].dropna()
                if len(vals) > 0:
                    s = getattr(vals, stat_fn)()
                    line += fmt_val(s, fmt_str, width, "right") + "  "
                else:
                    line += fmt_val(np.nan, fmt_str, width, "right") + "  "
            else:
                line += f"{'':>{width}}  "
        print(line)

    print("=" * 160)
    print("  * adj = NTM EBITDA lease-adjusted (TTM lease expense added for US GAAP reporters)")
    print("  IFRS 16 reporters already have leases below EBITDA; no adjustment needed.")


# -- Main -------------------------------------------------------------------
def main():
    xlsx_path = os.path.abspath("ofs_comp_temp.xlsx")

    print("Building OFS comp table workbook...")
    build_workbook(xlsx_path)

    print("\nLaunching Excel (isolated instance)...")
    session = launch_excel_isolated(xlsx_path)
    excel = session.excel
    print(f"  Connected to Excel {excel.Version}")
    ws_com = session.workbook.Sheets(1)

    # Trigger refresh
    print("Triggering RefreshSheet...")
    try:
        excel.Run("SNLXLAddin.xla!RefreshSheet")
        print("  RefreshSheet executed")
    except Exception as e:
        print(f"  RefreshSheet warning: {e}")

    # Poll for completion
    MAX_WAIT = 180  # 3 min - more formulas than comp_table.py
    POLL_INTERVAL = 5
    print(f"Waiting for formulas to evaluate (max {MAX_WAIT}s)...")
    start = time.monotonic()
    num_formulas = len(ALL_TICKERS) * len(METRICS)

    while True:
        elapsed = time.monotonic() - start
        if elapsed > MAX_WAIT:
            print(f"  Timeout after {MAX_WAIT}s -- proceeding with available data")
            break

        # Batch-read a sample: first row of data
        pending = False
        resolved = 0
        for comp_idx in range(len(ALL_TICKERS)):
            row = comp_idx + 2
            for met_idx in range(len(METRICS)):
                col = met_idx + 2
                val = ws_com.Cells(row, col).Value
                if isinstance(val, str) and val.upper() in ('#PEND', '#REFRESH'):
                    pending = True
                elif val is not None and not (isinstance(val, (int, float)) and val < -2000000000):
                    resolved += 1

        if not pending and resolved > num_formulas * 0.5:
            print(f"  Data ready! {resolved}/{num_formulas} cells resolved after {elapsed:.0f}s")
            break

        if int(elapsed) % 15 < POLL_INTERVAL:
            print(f"  {elapsed:.0f}s: {resolved}/{num_formulas} resolved, pending={pending}")

        time.sleep(POLL_INTERVAL)

    # Read all data at once
    print("\nReading data...")
    results = read_results(ws_com, len(ALL_TICKERS), len(METRICS))

    # Debug: show raw values
    print("\n-- Raw Values --")
    for d in results:
        ticker = d.get("_ticker", "?")
        gaap = d.get("GAAP", "?")
        name = d.get("Company Name", "?")
        mkt = d.get("Market Cap", "?")
        tev = d.get("TEV", "?")
        ltm_eb = d.get("LTM EBITDA", "?")
        lease = d.get("Lease Adj", "?")
        ntm_eb = d.get("NTM EBITDA", "?")
        print(f"  {ticker}: {name}, GAAP={gaap}, MktCap={mkt}, TEV={tev}, "
              f"LTM_EBITDA={ltm_eb}, LeaseAdj={lease}, NTM_EBITDA={ntm_eb}")

    # Close Excel
    close_session(session, save=False, delete_workbook=True)

    # Build and display
    df = build_comp_table(results)
    format_table(df)

    # Save CSV
    csv_path = os.path.abspath("ofs_comp_table.csv")
    df.to_csv(csv_path, index=False)
    print(f"\nCSV saved to: {csv_path}")

    return df


if __name__ == "__main__":
    main()
