"""
CLI entry point for capiq-excel.

Usage:
    capiq status            Show detected runtime profile and config
    capiq detect-addins     Detect installed add-ins (requires Excel)
    capiq download          Run the data download pipeline
    capiq doctor            Run diagnostics and check environment health
"""
from __future__ import annotations

import argparse
import logging
import sys
from typing import Optional

from capiq_excel.config import CapiqConfig, FormulaDialect, AddinMode


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        prog="capiq",
        description="Capital IQ Excel data downloader CLI",
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
    dl = sub.add_parser("download", help="Run the data download pipeline")
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

    args = parser.parse_args(argv)

    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s %(name)s %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
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


if __name__ == "__main__":
    sys.exit(main())
