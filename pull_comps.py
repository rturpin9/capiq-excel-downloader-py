"""Capital IQ Comparable Companies Table — CLI wrapper.

Thin CLI interface that delegates to capiq_mcp.comps_engine for business logic.
Supports the same arguments as before. Standalone usage still works.

Usage:
  python pull_comps.py NYSE:HAL NYSE:SLB TSX:PD TSX:TCW
  python pull_comps.py NYSE:HAL NYSE:SLB --currency USD
  python pull_comps.py NYSE:HAL NYSE:SLB --mode excluding-leases
  python pull_comps.py NYSE:HAL NYSE:SLB --no-lease-adjust-ntm
  python pull_comps.py NYSE:HAL NYSE:SLB --groups "U.S. OFS: NYSE:HAL NYSE:SLB" "Canadian OFS: TSX:PD TSX:TCW"
"""
import argparse
import json
import logging
import os
import sys
from datetime import datetime

import numpy as np
import pandas as pd

import pythoncom

pythoncom.CoInitialize()

# Configure logging for CLI usage (to stderr so stdout stays clean)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)

from capiq_mcp.comps_engine import (
    METRICS_LEASE_ADJUSTED,
    METRICS_EXCLUDING_LEASES,
    run_comps,
    dataframe_to_json,
    safe_float,
    safe_div,
)


# ── Formatting (CLI-only presentation) ─────────────────────────────────────

def format_value(val, fmt_str, width):
    if pd.isna(val):
        return f"{'N/A':>{width}}"
    try:
        return f"{fmt_str.format(val):>{width}}"
    except (ValueError, TypeError):
        return f"{'N/A':>{width}}"


def print_group_stats(df, mult_cols, fmts, widths, col_order, label):
    fn = "mean" if "Average" in label or "Avg" in label else "median"
    line = f"  {label:<35}"
    for col in col_order:
        w = widths[col]
        if col in mult_cols and col in fmts:
            vals = df[col].dropna()
            if len(vals) > 0:
                s = getattr(vals, fn)()
                line += f"  {format_value(s, fmts[col], w)}"
            else:
                line += f"  {'N/A':>{w}}"
        else:
            line += f"  {'':>{w}}"
    print(line)


def format_and_print(df, mode: str, currency: str, date: str,
                     lease_adjust_ntm: bool, groups: dict | None):
    """Format and print the comp table to stdout."""
    mode_label = "Lease Adjusted" if mode == "lease-adjusted" else "Excluding Leases"

    money = "${:,.0f}"
    mult = "{:.1f}x"
    pct = "{:.1f}%"

    if mode == "lease-adjusted":
        col_order = ["GAAP", "Mkt Cap", "Total Debt", "Leases", "Cash",
                     "Net Debt", "TEV", "LTM Rev", "NTM Rev", "LTM EBITDA",
                     "Lease Adj", "NTM EBITDA", "EBITDA Margin %", "EV/Rev",
                     "EV/LTM EBITDA", "EV/NTM EBITDA", "P/BV", "ND/EBITDA"]
        fmts = {
            "Mkt Cap": money, "Total Debt": money, "Leases": money,
            "Cash": money, "Net Debt": money, "TEV": money,
            "LTM Rev": money, "NTM Rev": money, "LTM EBITDA": money,
            "Lease Adj": money, "NTM EBITDA": money,
            "EBITDA Margin %": pct, "EV/Rev": mult,
            "EV/LTM EBITDA": mult, "EV/NTM EBITDA": mult,
            "P/BV": mult, "ND/EBITDA": mult,
        }
    else:
        col_order = ["Mkt Cap", "Total Debt", "Leases", "Cash",
                     "Net Debt", "TEV", "LTM Rev", "NTM Rev", "LTM EBITDA",
                     "NTM EBITDA", "EBITDA Margin %", "EV/Rev",
                     "EV/LTM EBITDA", "EV/NTM EBITDA", "P/BV", "ND/EBITDA"]
        fmts = {
            "Mkt Cap": money, "Total Debt": money, "Leases": money,
            "Cash": money, "Net Debt": money, "TEV": money,
            "LTM Rev": money, "NTM Rev": money, "LTM EBITDA": money,
            "NTM EBITDA": money,
            "EBITDA Margin %": pct, "EV/Rev": mult,
            "EV/LTM EBITDA": mult, "EV/NTM EBITDA": mult,
            "P/BV": mult, "ND/EBITDA": mult,
        }

    widths = {}
    for col in col_order:
        widths[col] = max(len(col), 8)
        if col == "GAAP":
            widths[col] = 9

    mult_cols = {"EBITDA Margin %", "EV/Rev", "EV/LTM EBITDA",
                 "EV/NTM EBITDA", "P/BV", "ND/EBITDA"}

    sep_len = 35 + sum(widths[c] + 2 for c in col_order) + 10
    sep = "-" * sep_len

    print(f"\n{'=' * sep_len}")
    print(f"  COMPARABLE COMPANIES ANALYSIS ({mode_label.upper()})")
    print(f"  As at {date} (Millions ${currency.upper()})")
    if mode == "lease-adjusted" and lease_adjust_ntm:
        print(f"  NTM EBITDA lease-adjusted for US GAAP reporters (TTM lease adj added as proxy)")
    elif mode == "lease-adjusted":
        print(f"  NTM EBITDA: straight consensus (lease adjustment DISABLED)")
    else:
        print(f"  TEV/Debt exclude operating leases; EBITDA excludes lease adjustment")
    print(f"{'=' * sep_len}")

    def print_header():
        header = f"  {'Company':<35}"
        for col in col_order:
            w = widths[col]
            header += f"  {col:>{w}}"
        print(header)
        print(f"  {sep}")

    def print_rows(section_df):
        for _, row in section_df.iterrows():
            cname = str(row.get("Company", ""))[:33]
            line = f"  {cname:<35}"
            for col in col_order:
                w = widths[col]
                val = row.get(col, np.nan)
                if col == "GAAP":
                    s = str(val) if isinstance(val, str) else "N/A"
                    line += f"  {s:>{w}}"
                elif col in fmts:
                    line += f"  {format_value(val, fmts[col], w)}"
                else:
                    line += f"  {str(val):>{w}}"
            if mode == "lease-adjusted" and row.get("Lease Adj Applied", False):
                line += "  *adj"
            print(line)

    if groups:
        for group_name, group_tickers in groups.items():
            group_df = df[df["Ticker"].isin(group_tickers)]
            if group_df.empty:
                continue

            print(f"\n  {group_name}")
            print(f"  {sep}")
            print_header()
            print_rows(group_df)
            print(f"  {sep}")

            print_group_stats(group_df, mult_cols, fmts, widths, col_order,
                              f"Average {group_name}")
            print_group_stats(group_df, mult_cols, fmts, widths, col_order,
                              f"Median {group_name}")
    else:
        print_header()
        print_rows(df)
        print(f"  {sep}")

    # Overall
    print(f"\n  Overall Summary")
    print(f"  {sep}")
    for stat_label, stat_fn in [("Overall Average", "mean"), ("Overall Median", "median")]:
        line = f"  {stat_label:<35}"
        for col in col_order:
            w = widths[col]
            if col in mult_cols and col in fmts:
                vals = df[col].dropna()
                if len(vals) > 0:
                    s = getattr(vals, stat_fn)()
                    line += f"  {format_value(s, fmts[col], w)}"
                else:
                    line += f"  {'N/A':>{w}}"
            else:
                line += f"  {'':>{w}}"
        print(line)

    print(f"{'=' * sep_len}")

    if mode == "lease-adjusted" and lease_adjust_ntm:
        print("  * adj = NTM EBITDA lease-adjusted (TTM lease expense added for US GAAP reporters)")
        print("  IFRS 16 reporters already have leases below EBITDA; no adjustment needed.")


# ── CLI ────────────────────────────────────────────────────────────────────

def parse_groups(group_args: list[str] | None, all_tickers: list[str]) -> dict | None:
    """Parse --groups 'Label: T1 T2' 'Label2: T3 T4' into dict."""
    if not group_args:
        return None

    groups = {}
    for g in group_args:
        if ":" not in g:
            print(f"  Warning: ignoring malformed group '{g}' (expected 'Label: T1 T2 ...')",
                  file=sys.stderr)
            continue
        label, tickers_str = g.split(":", 1)
        tickers = tickers_str.strip().split()
        groups[label.strip()] = tickers
    return groups if groups else None


def main():
    parser = argparse.ArgumentParser(
        description="Pull comparable company data via S&P Capital IQ SPG formulas.")
    parser.add_argument("tickers", nargs="+",
                        help="Company tickers (e.g. NYSE:HAL TSX:PD DSGX)")
    parser.add_argument("--currency", default="CAD",
                        help="Currency for output (default: CAD)")
    parser.add_argument("--mode", choices=["lease-adjusted", "excluding-leases"],
                        default="lease-adjusted",
                        help="Comp table mode (default: lease-adjusted)")
    parser.add_argument("--no-lease-adjust-ntm", action="store_true",
                        help="Disable NTM EBITDA lease adjustment for US GAAP reporters")
    parser.add_argument("--groups", nargs="*",
                        help='Group tickers: "U.S.: NYSE:HAL NYSE:SLB" "Canada: TSX:PD TSX:TCW"')
    parser.add_argument("--date", default=None,
                        help="As-of date in M/D/YYYY format (default: today)")
    parser.add_argument("--json-output", default=None,
                        help="Write structured JSON to this path")
    parser.add_argument("--csv-output", default=None,
                        help="Write CSV to this path (default: comps_output.csv)")
    parser.add_argument("--max-wait", type=int, default=180,
                        help="Max seconds to wait for formula refresh (default: 180)")

    args = parser.parse_args()

    # Resolve date
    if args.date:
        date_str = args.date
    else:
        now = datetime.now()
        date_str = f"{now.month}/{now.day}/{now.year}"

    lease_adjust_ntm = not args.no_lease_adjust_ntm
    mode = args.mode
    tickers = args.tickers
    csv_path = args.csv_output or os.path.abspath("comps_output.csv")
    groups = parse_groups(args.groups, tickers)

    print(f"Comparable Companies Analysis", file=sys.stderr)
    print(f"  Mode: {mode}", file=sys.stderr)
    print(f"  Currency: {args.currency.upper()}", file=sys.stderr)
    print(f"  Date: {date_str}", file=sys.stderr)
    print(f"  Tickers: {' '.join(tickers)}", file=sys.stderr)
    if mode == "lease-adjusted":
        print(f"  NTM lease adjustment: {'ON' if lease_adjust_ntm else 'OFF'}", file=sys.stderr)

    # Run the engine (all business logic is in capiq_mcp.comps_engine)
    data = run_comps(
        tickers=tickers,
        currency=args.currency,
        mode=mode,
        lease_adjust_ntm=lease_adjust_ntm,
        date=date_str,
        max_wait=args.max_wait,
    )

    # Reconstruct DataFrame from JSON result for CLI formatting
    df = pd.DataFrame(data["companies"])

    # Report identifier resolution
    resolution_notes = data.get("resolution_notes", [])
    resolved_via_fallback = [n for n in resolution_notes if n["status"] == "RESOLVED"]
    failed = [n for n in resolution_notes if n["status"] == "FAILED"]

    if resolved_via_fallback or failed:
        print("\n" + "=" * 70)
        print("  IDENTIFIER RESOLUTION REPORT")
        print("=" * 70)
        for note in resolved_via_fallback:
            iq_id = note.get("resolved_iq_id", "?")
            print(f"  {note['ticker']}: resolved via CIQRANGEA -> {iq_id} (data OK)")
        for note in failed:
            print(f"  {note['ticker']}: FAILED — not recognized by SPG or CIQRANGEA")
            print(f"    -> Check spelling, or try EXCHANGE:TICKER format (e.g., NYSE:HAL, TSX:PD)")
        print("=" * 70 + "\n")

    # Print formatted output
    format_and_print(df, mode, args.currency, date_str, lease_adjust_ntm, groups)

    # Save CSV
    df.to_csv(csv_path, index=False)
    print(f"\nCSV saved to: {csv_path}", file=sys.stderr)

    # Save JSON if requested
    if args.json_output:
        with open(args.json_output, "w") as f:
            json.dump(data, f, indent=2)
        print(f"JSON saved to: {args.json_output}", file=sys.stderr)

    return df


if __name__ == "__main__":
    main()
