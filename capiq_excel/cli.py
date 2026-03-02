"""
CLI entry point for capiq-excel.

Usage:
    capiq status            Show detected runtime profile and config
    capiq detect-addins     Detect installed add-ins (requires Excel)
    capiq download          Run the data download pipeline
    capiq doctor            Run environment diagnostics
    capiq comps             Pull comparable companies analysis table
    capiq chart             Create time-series chart from CIQ data
    capiq lookup            Resolve company identifiers to CIQ IQ IDs
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from typing import Optional

import numpy as np
import pandas as pd

from capiq_excel.config import CapiqConfig, FormulaDialect, AddinMode


# ── Markdown table formatter ──────────────────────────────────────────────

def _fmt_dollar(val) -> str:
    """Format a dollar value as $X,XXX (no decimals)."""
    if pd.isna(val):
        return "N/A"
    try:
        return f"${val:,.0f}"
    except (ValueError, TypeError):
        return "N/A"


def _fmt_mult(val) -> str:
    """Format a multiple as X.Xx (one decimal)."""
    if pd.isna(val):
        return "N/A"
    try:
        return f"{val:.1f}x"
    except (ValueError, TypeError):
        return "N/A"


def _fmt_pct(val) -> str:
    """Format a percentage as XX.X%."""
    if pd.isna(val):
        return "N/A"
    try:
        return f"{val:.1f}%"
    except (ValueError, TypeError):
        return "N/A"


def _fmt_count(val) -> str:
    """Format a count value as X,XXX (no decimals, no dollar sign)."""
    if pd.isna(val):
        return "N/A"
    try:
        return f"{val:,.0f}"
    except (ValueError, TypeError):
        return "N/A"


def _fmt_str(val) -> str:
    """Format a string value."""
    if pd.isna(val) or val is None:
        return "N/A"
    return str(val)


# Column format specs: (column_name, formatter, is_summary_col)
_COLS_LEASE_ADJ = [
    ("GAAP", _fmt_str, False),
    ("Mkt Cap", _fmt_dollar, False),
    ("Total Debt", _fmt_dollar, False),
    ("Leases", _fmt_dollar, False),
    ("Cash", _fmt_dollar, False),
    ("Net Debt", _fmt_dollar, False),
    ("TEV", _fmt_dollar, False),
    ("LTM Rev", _fmt_dollar, False),
    ("NTM Rev", _fmt_dollar, False),
    ("LTM EBITDA", _fmt_dollar, False),
    ("Lease Adj", _fmt_dollar, False),
    ("NTM EBITDA", _fmt_dollar, False),
    ("EBITDA Margin %", _fmt_pct, True),
    ("Gross Margin %", _fmt_pct, True),
    ("CapEx", _fmt_dollar, False),
    ("EV/Rev", _fmt_mult, True),
    ("EV/LTM EBITDA", _fmt_mult, True),
    ("EV/NTM EBITDA", _fmt_mult, True),
    ("P/BV", _fmt_mult, True),
    ("ND/EBITDA", _fmt_mult, True),
]

_COLS_EXCL_LEASES = [
    ("Mkt Cap", _fmt_dollar, False),
    ("Total Debt", _fmt_dollar, False),
    ("Leases", _fmt_dollar, False),
    ("Cash", _fmt_dollar, False),
    ("Net Debt", _fmt_dollar, False),
    ("TEV", _fmt_dollar, False),
    ("LTM Rev", _fmt_dollar, False),
    ("NTM Rev", _fmt_dollar, False),
    ("LTM EBITDA", _fmt_dollar, False),
    ("NTM EBITDA", _fmt_dollar, False),
    ("EBITDA Margin %", _fmt_pct, True),
    ("Gross Margin %", _fmt_pct, True),
    ("CapEx", _fmt_dollar, False),
    ("EV/Rev", _fmt_mult, True),
    ("EV/LTM EBITDA", _fmt_mult, True),
    ("EV/NTM EBITDA", _fmt_mult, True),
    ("P/BV", _fmt_mult, True),
    ("ND/EBITDA", _fmt_mult, True),
]


def _stat_row(df: pd.DataFrame, cols: list, fn: str, label: str) -> str:
    """Build a markdown table row with mean/median of summary columns."""
    parts = [label]
    for col_name, fmt, is_summary in cols:
        if is_summary and col_name in df.columns:
            vals = df[col_name].dropna()
            if len(vals) > 0:
                parts.append(fmt(getattr(vals, fn)()))
            else:
                parts.append("N/A")
        else:
            parts.append("")
    return "| " + " | ".join(parts) + " |"


def parse_groups(group_args: list[str] | None) -> dict | None:
    """Parse --groups 'Label: T1 T2' 'Label2: T3 T4' into dict."""
    if not group_args:
        return None
    groups = {}
    for g in group_args:
        if ":" not in g:
            print(f"Warning: ignoring malformed group '{g}' (expected 'Label: T1 T2 ...')",
                  file=sys.stderr)
            continue
        label, tickers_str = g.split(":", 1)
        tickers = tickers_str.strip().split()
        groups[label.strip()] = tickers
    return groups if groups else None


def _format_comps_markdown(
    data: dict,
    mode: str,
    currency: str,
    date: str,
    lease_adjust_ntm: bool,
    groups: dict | None = None,
    extra_specs: list | None = None,
) -> str:
    """Format comp table data as a markdown string."""
    lines: list[str] = []
    df = pd.DataFrame(data["companies"])

    mode_label = "Lease Adjusted" if mode == "lease-adjusted" else "Excluding Leases"
    cols = list(_COLS_LEASE_ADJ if mode == "lease-adjusted" else _COLS_EXCL_LEASES)

    # Append extra columns dynamically
    for spec in (extra_specs or []):
        fmt_fn = {"pct": _fmt_pct, "dollar": _fmt_dollar, "mult": _fmt_mult, "count": _fmt_count}[spec.fmt]
        cols.append((spec.label, fmt_fn, spec.is_summary))

    # Header
    lines.append(f"**Comparable Companies Analysis ({mode_label})**")
    lines.append(f"As at {date} (Millions ${currency.upper()})")
    if mode == "lease-adjusted" and lease_adjust_ntm:
        lines.append("NTM EBITDA lease-adjusted for US GAAP reporters (TTM lease adj added as proxy)")
    elif mode == "lease-adjusted":
        lines.append("NTM EBITDA: straight consensus (lease adjustment disabled)")
    else:
        lines.append("TEV/Debt exclude operating leases; EBITDA excludes lease adjustment")
    lines.append("")

    # Build table header
    col_names = ["Company"] + [c[0] for c in cols]
    header = "| " + " | ".join(col_names) + " |"
    separator = "| " + " | ".join("---" for _ in col_names) + " |"

    def _company_rows(section_df: pd.DataFrame) -> list[str]:
        row_lines = []
        for _, row in section_df.iterrows():
            cname = str(row.get("Company", ""))[:40]
            if mode == "lease-adjusted" and row.get("Lease Adj Applied", False):
                cname += " *"
            parts = [cname]
            for col_name, fmt, _ in cols:
                parts.append(fmt(row.get(col_name, np.nan)))
            row_lines.append("| " + " | ".join(parts) + " |")
        return row_lines

    if groups:
        for group_name, group_tickers in groups.items():
            group_df = df[df["Ticker"].isin(group_tickers)]
            if group_df.empty:
                continue
            lines.append(f"### {group_name}")
            lines.append("")
            lines.append(header)
            lines.append(separator)
            lines.extend(_company_rows(group_df))
            lines.append(_stat_row(group_df, cols, "mean", f"**Avg {group_name}**"))
            lines.append(_stat_row(group_df, cols, "median", f"**Med {group_name}**"))
            lines.append("")
    else:
        lines.append(header)
        lines.append(separator)
        lines.extend(_company_rows(df))

    # Overall stats
    lines.append(_stat_row(df, cols, "mean", "**Average**"))
    lines.append(_stat_row(df, cols, "median", "**Median**"))

    # Footnotes
    if mode == "lease-adjusted" and lease_adjust_ntm:
        lines.append("")
        lines.append("\\* NTM EBITDA lease-adjusted (TTM lease expense added for US GAAP reporters)")

    return "\n".join(lines)


# ── CLI entry point ───────────────────────────────────────────────────────

def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        prog="capiq",
        description="Capital IQ Excel data downloader and analysis CLI",
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true",
        help="Enable verbose/debug logging",
    )
    sub = parser.add_subparsers(dest="command")

    # --- status ---
    sub.add_parser("status", help="Show current config and environment info")

    # --- detect-addins ---
    sub.add_parser("detect-addins", help="Detect installed Capital IQ add-ins (starts Excel)")

    # --- download ---
    dl = sub.add_parser("download", help="Run the legacy data download pipeline")
    dl.add_argument("--ids", nargs="+", help="Company identifiers (tickers, CUSIPs, etc.)")
    dl.add_argument("--ids-file", help="Path to file with one identifier per line")
    dl.add_argument("--financial-items", nargs="+", help="Financial data items (e.g. IQ_TOTAL_REVENUE)")
    dl.add_argument("--market-items", nargs="+", help="Market data items (e.g. IQ_CLOSEPRICE)")
    dl.add_argument("--outpath", default="capiq data.csv", help="Output CSV path")
    dl.add_argument("--dialect", choices=["auto", "ciq", "spg"], default="auto",
                    help="Formula dialect to use")
    dl.add_argument("--addin-mode", choices=["auto", "legacy", "pro"], default="auto",
                    help="Add-in mode")
    dl.add_argument("--freq", choices=["Q", "Y"], default="Q", help="Data frequency")
    dl.add_argument("--num-periods", type=int, default=80, help="Number of periods")
    dl.add_argument("--timeout", type=int, default=240, help="Timeout per file (seconds)")

    # --- doctor ---
    sub.add_parser("doctor", help="Run environment diagnostics")

    # --- comps ---
    cp = sub.add_parser("comps", help="Pull comparable companies analysis table via SPG formulas")
    cp.add_argument("tickers", nargs="+",
                    help="Company tickers (e.g. NYSE:HAL TSX:PD DSGX)")
    cp.add_argument("--currency", default="CAD",
                    help="Output currency code (default: CAD)")
    cp.add_argument("--mode", choices=["lease-adjusted", "excluding-leases"],
                    default="lease-adjusted",
                    help="Comp table mode (default: lease-adjusted)")
    cp.add_argument("--no-lease-adjust-ntm", action="store_true",
                    help="Disable NTM EBITDA lease adjustment for US GAAP reporters")
    cp.add_argument("--groups", nargs="*",
                    help='Group tickers by sector: "U.S.: NYSE:HAL NYSE:SLB" "Canada: TSX:PD TSX:TCW"')
    cp.add_argument("--date", default=None,
                    help="As-of date in M/D/YYYY format (default: today)")
    cp.add_argument("--max-wait", type=int, default=180,
                    help="Max seconds to wait for formula refresh (default: 180)")
    cp.add_argument("--json", action="store_true", dest="json_output",
                    help="Output raw JSON dict instead of markdown table")
    cp.add_argument("--csv", default=None, metavar="PATH",
                    help="Also write CSV to this path")
    cp.add_argument("--extra", nargs="*", default=None,
                    help='Optional extra metrics: "roe" "rev-growth" "div-yield" (run --list-extras to see all)')
    cp.add_argument("--list-extras", action="store_true",
                    help="Print available extra metrics and exit")

    # --- chart ---
    ch = sub.add_parser("chart", help="Create time-series chart from Capital IQ data")
    ch.add_argument("tickers", nargs="+",
                    help="Company tickers (e.g. DSGX NYSE:HAL TSX:PD)")
    ch.add_argument("--metrics", nargs="+", required=True,
                    help="CIQ mnemonics (e.g. IQ_CLOSEPRICE IQ_TEV_EBITDA)")
    ch.add_argument("--metric-type", required=True,
                    choices=["market", "multiple", "financial"],
                    help="Data category: market (prices), multiple (EV/EBITDA), financial (revenue)")
    ch.add_argument("--chart-type", default="line",
                    choices=["line", "bar", "line_marker", "dual_axis"],
                    help="Chart style (default: line)")
    ch.add_argument("--start-date", default=None,
                    help="Start date M/D/YYYY (default: 1Y ago for market/multiple)")
    ch.add_argument("--end-date", default=None,
                    help="End date M/D/YYYY (default: today)")
    ch.add_argument("--period-type", default="IQ_LTM",
                    help="Period type for multiples: IQ_LTM, IQ_NTM, etc. (default: IQ_LTM)")
    ch.add_argument("--frequency", default="Q", choices=["Q", "Y"],
                    help="Frequency for financial data: Q (quarterly) or Y (annual)")
    ch.add_argument("--num-periods", type=int, default=12,
                    help="Number of periods back for financial data (default: 12)")
    ch.add_argument("--indexed", action="store_true",
                    help="Index market data to base 100 for relative comparison")
    ch.add_argument("--currency", default="USD",
                    help="Output currency code (default: USD)")
    ch.add_argument("--title", default=None,
                    help="Chart title (auto-generated if omitted)")
    ch.add_argument("--output", default=None, metavar="PATH",
                    help="Output PNG path (default: chart_output.png)")
    ch.add_argument("--max-wait", type=int, default=180,
                    help="Max seconds to wait for formula refresh (default: 180)")

    # --- lookup ---
    lk = sub.add_parser("lookup", help="Resolve company identifiers to CIQ IQ IDs and names")
    lk.add_argument("identifiers", nargs="+",
                    help="Identifiers to resolve (tickers, company names, CUSIPs, ISINs)")
    lk.add_argument("--max-wait", type=int, default=60,
                    help="Max seconds to wait for resolution (default: 60)")

    args = parser.parse_args(argv)

    # Configure logging to stderr so stdout stays clean for data output
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s %(name)s %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
        stream=sys.stderr,
    )

    if args.command is None:
        parser.print_help()
        return 0

    if args.command == "status":
        return _cmd_status()
    elif args.command == "detect-addins":
        return _cmd_detect_addins()
    elif args.command == "download":
        return _cmd_download(args)
    elif args.command == "doctor":
        return _cmd_doctor()
    elif args.command == "comps":
        return _cmd_comps(args)
    elif args.command == "chart":
        return _cmd_chart(args)
    elif args.command == "lookup":
        return _cmd_lookup(args)
    else:
        parser.print_help()
        return 1


def _cmd_status() -> int:
    """Show current config from environment."""
    config = CapiqConfig.from_env()
    print("Capital IQ Excel Downloader — Current Configuration")
    print("=" * 52)
    print(f"  Formula dialect:  {config.formula_dialect.value}")
    print(f"  Add-in mode:      {config.addin_mode.value}")
    print(f"  Refresh scope:    {config.refresh_scope.value}")
    print(f"  Frequency:        {config.freq}")
    print(f"  Num periods:      {config.num_periods}")
    print(f"  Timeout:          {config.retry.timeout_seconds}s")
    print(f"  Max retries:      {config.retry.max_retries}")
    print(f"  Restart interval: {config.retry.restart_interval}")
    opts = config.formula_options.to_spg_options_string()
    print(f"  Formula options:  {opts or '(none)'}")
    print()
    print("Environment variables:")
    import os
    for var in ("CAPIQ_FORMULA_DIALECT", "CAPIQ_ADDIN_MODE", "CAPIQ_REFRESH_SCOPE",
                "CAPIQ_FREQ", "CAPIQ_NUM_PERIODS", "CAPIQ_TIMEOUT", "CAPIQ_MAX_RETRIES",
                "CAPIQ_RESTART_INTERVAL"):
        val = os.environ.get(var)
        print(f"  {var} = {val or '(not set)'}")
    return 0


def _cmd_detect_addins() -> int:
    """Detect installed add-ins by starting Excel and scanning COM."""
    print("Starting Excel and scanning for Capital IQ add-ins...")
    try:
        from exceldriver.tools import _start_excel_with_addins_and_attach
        from capiq_excel.runtime.addin_detection import detect_runtime
    except ImportError as e:
        print(f"Error: Required packages not available: {e}")
        print("This command requires pypiwin32 and exceldriver to be installed.")
        return 1

    try:
        excel = _start_excel_with_addins_and_attach()
        profile = detect_runtime(excel)
    except Exception as e:
        print(f"Error: Could not start Excel or detect add-ins: {e}")
        return 1

    print()
    print("Runtime Profile")
    print("=" * 40)
    print(f"  Excel version:      {profile.excel_version or 'unknown'}")
    print(f"  Excel bitness:      {profile.excel_bitness or 'unknown'}")
    print(f"  Add-in mode:        {profile.addin_mode}")
    print(f"  Legacy add-in:      {profile.legacy_addin_name or '(not found)'}")
    print(f"  Pro add-in:         {profile.pro_addin_name or '(not found)'}")
    print(f"  Pro installed:      {profile.pro_installed}")
    print(f"  Pro load behavior:  {profile.pro_load_behavior}")
    print(f"  CIQ compat:         {profile.ciq_compat_enabled}")
    print(f"  Available dialects: {', '.join(sorted(profile.available_dialects)) or '(none)'}")
    print(f"  Install path:       {profile.install_path or '(unknown)'}")
    print(f"  Log path:           {profile.log_path or '(unknown)'}")
    return 0


def _cmd_download(args) -> int:
    """Run the download pipeline."""
    # Collect company IDs
    ids = []
    if args.ids:
        ids.extend(args.ids)
    if args.ids_file:
        try:
            with open(args.ids_file) as f:
                ids.extend(line.strip() for line in f if line.strip())
        except FileNotFoundError:
            print(f"Error: IDs file not found: {args.ids_file}")
            return 1

    if not ids:
        print("Error: No company identifiers provided. Use --ids or --ids-file.")
        return 1

    financial_items = args.financial_items
    market_items = args.market_items
    if not financial_items and not market_items:
        print("Error: Must provide --financial-items and/or --market-items.")
        return 1

    config = CapiqConfig.from_env()
    # Override with CLI arguments
    config.formula_dialect = FormulaDialect(args.dialect)
    config.addin_mode = AddinMode(args.addin_mode)
    config.freq = args.freq
    config.num_periods = args.num_periods
    config.retry.timeout_seconds = args.timeout

    from capiq_excel.main import download_data

    try:
        download_data(
            company_ids=ids,
            financial_data_items=financial_items,
            market_data_items=market_items,
            data_outpath=args.outpath,
            config=config,
            freq=config.freq,
            num_periods=config.num_periods,
        )
        print(f"\nData saved to: {args.outpath}")
        return 0
    except Exception as e:
        print(f"\nError during download: {e}")
        logging.getLogger(__name__).debug("Download error details:", exc_info=True)
        return 1


def _cmd_doctor() -> int:
    """Run environment diagnostics."""
    print("Capital IQ Excel Downloader — Environment Check")
    print("=" * 50)
    issues = []

    # Check Python version
    print(f"\n[1] Python: {sys.version}")

    # Check required packages
    print("\n[2] Required packages:")
    for pkg in ("pandas", "openpyxl", "win32com", "pythoncom", "pywintypes",
                "exceldriver", "processfiles"):
        try:
            __import__(pkg)
            print(f"  {pkg}: OK")
        except ImportError:
            print(f"  {pkg}: MISSING")
            issues.append(f"Package '{pkg}' is not installed")

    # Check config
    print("\n[3] Configuration:")
    config = CapiqConfig.from_env()
    print(f"  Dialect: {config.formula_dialect.value}")
    print(f"  Mode:    {config.addin_mode.value}")

    # Check registry (Windows only)
    print("\n[4] Registry (Windows):")
    try:
        import winreg
        from capiq_excel.runtime.addin_detection import _REG_EXCEL_ADDINS, _REG_SNL_OFFICE

        for name, path in [("Excel Add-ins", _REG_EXCEL_ADDINS), ("SNL Office", _REG_SNL_OFFICE)]:
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path):
                    print(f"  {name}: Found")
            except FileNotFoundError:
                print(f"  {name}: Not found")
    except ImportError:
        print("  winreg: Not available (non-Windows platform)")
        issues.append("Running on non-Windows platform — COM automation will not work")

    # Summary
    print(f"\n{'=' * 50}")
    if issues:
        print(f"Found {len(issues)} issue(s):")
        for issue in issues:
            print(f"  - {issue}")
        return 1
    else:
        print("All checks passed!")
        return 0


# ── New subcommands ───────────────────────────────────────────────────────

def _cmd_comps(args) -> int:
    """Pull comparable companies analysis table."""
    from capiq_excel.engines.comps import run_comps, resolve_extras, EXTRA_METRICS_CATALOG

    # Handle --list-extras
    if args.list_extras:
        print("Available extra metrics for --extra flag:\n")
        print(f"{'Key':<16} {'Label':<20} {'Type':<8} {'Summary':<8} {'CIQ Mnemonic'}")
        print(f"{'-'*16} {'-'*20} {'-'*8} {'-'*8} {'-'*30}")
        for key in sorted(EXTRA_METRICS_CATALOG):
            s = EXTRA_METRICS_CATALOG[key]
            print(f"{s.key:<16} {s.label:<20} {s.fmt:<8} {'yes' if s.is_summary else 'no':<8} {s.mnemonic}")
        print(f"\nUsage: capiq comps TICKER1 TICKER2 --extra roe rev-growth div-yield")
        return 0

    lease_adjust_ntm = not args.no_lease_adjust_ntm
    groups = parse_groups(args.groups)

    # Resolve extras early to fail fast on bad keys
    extra_specs = []
    if args.extra:
        try:
            extra_specs, _ = resolve_extras(args.extra)
        except ValueError as e:
            print(json.dumps({"error": "validation", "message": str(e)}))
            return 1

    try:
        data = run_comps(
            tickers=args.tickers,
            currency=args.currency,
            mode=args.mode,
            lease_adjust_ntm=lease_adjust_ntm,
            date=args.date,
            max_wait=args.max_wait,
            extras=args.extra,
        )
    except Exception as e:
        print(json.dumps({"error": "execution", "message": str(e)}))
        return 1

    if args.json_output:
        print(json.dumps(data, indent=2))
    else:
        md = _format_comps_markdown(
            data, args.mode, args.currency, data["date"],
            lease_adjust_ntm, groups, extra_specs,
        )
        print(md)

    if args.csv:
        df = pd.DataFrame(data["companies"])
        df.to_csv(args.csv, index=False)
        print(f"CSV saved to: {args.csv}", file=sys.stderr)

    return 0


def _cmd_chart(args) -> int:
    """Create time-series chart from CIQ data."""
    from capiq_excel.engines.chart import run_chart

    try:
        result = run_chart(
            tickers=args.tickers,
            metrics=args.metrics,
            metric_type=args.metric_type,
            start_date=args.start_date,
            end_date=args.end_date,
            period_type=args.period_type,
            frequency=args.frequency,
            num_periods=args.num_periods,
            chart_type=args.chart_type,
            indexed=args.indexed,
            currency=args.currency,
            title=args.title,
            output_path=args.output,
            max_wait=args.max_wait,
        )
    except Exception as e:
        print(json.dumps({"error": "execution", "message": str(e)}))
        return 1

    print(json.dumps(result, indent=2))
    return 0


def _cmd_lookup(args) -> int:
    """Resolve company identifiers to CIQ IQ IDs."""
    from capiq_excel.engines.id_lookup import run_id_lookup

    try:
        result = run_id_lookup(
            identifiers=args.identifiers,
            max_wait=args.max_wait,
        )
    except Exception as e:
        print(json.dumps({"error": "execution", "message": str(e)}))
        return 1

    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
