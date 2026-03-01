"""
Indexed equity price chart for Constellation Software, Salesforce, and Nvidia.

Uses CIQRANGE with market data syntax:
  =CIQRANGE(ticker, "IQ_CLOSEPRICE", "start_date", "end_date", , , , , , "Label")

Prices are indexed to 100 at the start of 2026.
"""
import os
import time
import openpyxl
import pandas as pd
import numpy as np
import pythoncom
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from capiq_excel.excel_lifecycle import launch_excel_isolated, close_session

pythoncom.CoInitialize()

# ---- Configuration ----
COMPANIES = [
    ("TSX:CSU", "Constellation Software"),
    ("CRM", "Salesforce"),
    ("NVDA", "Nvidia"),
]

START_DATE = "1/2/2026"
END_DATE = "2/28/2026"

BASE_DIR = r"C:\Claude"
XLSX_PATH = os.path.join(BASE_DIR, "indexed_equity.xlsx")
CHART_PATH = os.path.join(BASE_DIR, "indexed_equity_chart.png")


def create_workbook():
    """Create XLSX with CIQRANGE close price formulas, one sheet per company."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for ticker, name in COMPANIES:
        ws = wb.create_sheet(title=name)
        ws.cell(row=1, column=1,
                value=f'=CIQRANGE("{ticker}","IQ_CLOSEPRICE","{START_DATE}","{END_DATE}",,,,,,"Close")')

    wb.save(XLSX_PATH)
    print(f"Created workbook: {XLSX_PATH}")


def open_refresh_and_extract():
    """Launch Excel, refresh, extract close price data."""
    print("Launching Excel (isolated instance)...")
    session = launch_excel_isolated(XLSX_PATH)
    excel = session.excel
    print(f"Connected to Excel {excel.Version}")

    try:
        excel.Run("SNLXLAddin.xla!RefreshWorkbook")
        print("Triggered RefreshWorkbook")
    except Exception as e:
        print(f"Refresh warning: {e}")

    # Poll for data
    print("Waiting for CIQ to evaluate...")
    for tick in range(30):
        time.sleep(5)
        elapsed = (tick + 1) * 5
        ws = session.workbook.Sheets(1)
        val = ws.Cells(2, 1).Value
        print(f"  [{elapsed}s] A2={val!r}")
        if isinstance(val, (int, float)) and val > 0:
            time.sleep(10)
            break
    else:
        print("  Timeout — extracting available data")

    # Extract from each sheet
    all_data = {}
    for sheet_idx in range(1, session.workbook.Sheets.Count + 1):
        ws = session.workbook.Sheets(sheet_idx)
        company = ws.Name
        print(f"  Extracting: {company}")

        # Find last row via End(xlUp)
        last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        if last_row < 2:
            print(f"    No data")
            continue

        # Batch read all values in one COM call
        values = ws.Range(f'A2:A{last_row}').Value
        prices = [v[0] if isinstance(v, tuple) else v for v in values]
        prices = [_safe_float(p) for p in prices]

        n = len(prices)
        # Generate trading-day dates
        dates = pd.bdate_range(start=START_DATE, periods=n)

        df = pd.DataFrame({'Date': dates, 'Close': prices})
        all_data[company] = df
        print(f"    {n} data points")

    close_session(session, save=False, delete_workbook=False)
    return all_data


def _safe_float(val):
    if val is None:
        return np.nan
    if isinstance(val, (int, float)):
        return float(val) if val > 0 else np.nan
    if isinstance(val, str):
        try:
            return float(val)
        except ValueError:
            return np.nan
    return np.nan


def index_to_100(all_data):
    """Index each series so the first valid price = 100."""
    indexed = {}
    for company, df in all_data.items():
        series = df[['Date', 'Close']].dropna(subset=['Close'])
        if len(series) == 0:
            continue
        base = series['Close'].iloc[0]
        series = series.copy()
        series['Indexed'] = (series['Close'] / base) * 100
        indexed[company] = series
    return indexed


def plot_chart(indexed_data):
    """Create indexed price chart."""
    fig, ax = plt.subplots(figsize=(12, 6))

    colors = {'Constellation Software': '#1f77b4', 'Salesforce': '#ff7f0e', 'Nvidia': '#2ca02c'}

    for company, df in sorted(indexed_data.items()):
        ax.plot(df['Date'], df['Indexed'],
                label=company, linewidth=2.2,
                color=colors.get(company, 'gray'))

    ax.axhline(y=100, color='gray', linestyle='--', linewidth=0.8, alpha=0.5)
    ax.set_ylabel('Indexed Price (Jan 2 = 100)')
    ax.set_title('YTD Equity Performance — Indexed to 100', fontsize=14, fontweight='bold')
    ax.legend(loc='best', fontsize=10, framealpha=0.9)
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %d'))
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=2))

    fig.autofmt_xdate()
    plt.tight_layout()
    plt.savefig(CHART_PATH, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"\nChart saved: {CHART_PATH}")


def print_summary(indexed_data):
    """Print performance summary table."""
    print("\n" + "=" * 65)
    print(f"{'Company':<25} {'Start':>8} {'Current':>8} {'Change':>8} {'Points':>7}")
    print("-" * 65)
    for company, df in sorted(indexed_data.items()):
        start = df['Close'].iloc[0]
        current = df['Close'].iloc[-1]
        pct = (current / start - 1) * 100
        print(f"{company:<25} ${start:>7.2f} ${current:>7.2f} {pct:>+7.1f}% {len(df):>7}")
    print("=" * 65)


if __name__ == "__main__":
    print("=" * 55)
    print("  Indexed Equity Price Chart (via CIQ)")
    print(f"  Companies: {', '.join(n for _, n in COMPANIES)}")
    print(f"  Period: {START_DATE} to {END_DATE}")
    print("=" * 55)

    create_workbook()
    all_data = open_refresh_and_extract()

    if all_data:
        indexed = index_to_100(all_data)
        plot_chart(indexed)
        print_summary(indexed)
    else:
        print("\nNo data extracted.")
