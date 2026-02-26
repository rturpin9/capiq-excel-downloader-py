"""
EV/EBITDA and EV/Revenue trend charts for 6 companies via CIQ CIQRANGEV.

Uses the correct multiples syntax:
  =CIQRANGEV(ticker, metric, IQ_LTM, start, end, , , , , label)

This gives weekly pre-calculated LTM multiples directly — no manual
rolling-sum calculation needed.
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
    ("DSGX", "Descartes"),
    ("IEX", "IDEX"),
    ("MANH", "Manhattan"),
    ("ROP", "Roper"),
    ("WTC", "WiseTech"),
    ("CMDXF", "Comp Modelling"),
]

LOOKBACK = "-1Y"           # relative: 1 year back
END_DATE = "2/26/2026"     # today

BASE_DIR = r"C:\Claude"
XLSX_PATH = os.path.join(BASE_DIR, "ev_mult_v2.xlsx")
CHART_PATH = os.path.join(BASE_DIR, "ev_multiples_chart.png")


def create_workbook():
    """Create XLSX with CIQRANGEV formulas for EV/EBITDA and EV/Revenue per company."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for ticker, name in COMPANIES:
        ws = wb.create_sheet(title=name)
        # Column A: EV/EBITDA (LTM)
        ws.cell(row=1, column=1,
                value=f'=CIQRANGEV("{ticker}","IQ_TEV_EBITDA",IQ_LTM,"{LOOKBACK}","{END_DATE}",,,,"TEV/EBITDA")')
        # Column B: EV/Revenue (LTM)
        ws.cell(row=1, column=2,
                value=f'=CIQRANGEV("{ticker}","IQ_TEV_TOTAL_REV",IQ_LTM,"{LOOKBACK}","{END_DATE}",,,,"TEV/Revenue")')

    wb.save(XLSX_PATH)
    print(f"Created workbook: {XLSX_PATH}")


def open_refresh_and_extract():
    """Launch Excel, refresh, extract multiples data."""
    print("Launching Excel (isolated instance)...")
    session = launch_excel_isolated(XLSX_PATH)
    excel = session.excel
    print(f"Connected to Excel {excel.Version}")

    try:
        excel.Run("SNLXLAddin.xla!RefreshWorkbook")
        print("Triggered RefreshWorkbook")
    except Exception as e:
        print(f"Refresh warning: {e}")

    # Poll for completion
    print("Waiting for CIQ to evaluate...")
    for tick in range(30):
        time.sleep(5)
        elapsed = (tick + 1) * 5
        ws = session.workbook.Sheets(1)
        val = ws.Cells(2, 1).Value
        print(f"  [{elapsed}s] A2={val!r}")
        if isinstance(val, (int, float)) and val > 0:
            time.sleep(15)  # let all sheets settle
            break
    else:
        print("  Timeout — extracting available data")

    # Extract from each sheet
    all_data = {}
    for sheet_idx in range(1, session.workbook.Sheets.Count + 1):
        ws = session.workbook.Sheets(sheet_idx)
        company = ws.Name
        print(f"  Extracting: {company}")

        rows_data = []
        for row in range(2, 200):
            a = ws.Cells(row, 1).Value  # EV/EBITDA
            b = ws.Cells(row, 2).Value  # EV/Revenue
            if a is None and b is None:
                if row > 3:
                    break
                continue
            rows_data.append((_safe_float(a), _safe_float(b)))

        if rows_data:
            ev_ebitda = [r[0] for r in rows_data]
            ev_rev = [r[1] for r in rows_data]
            n = len(rows_data)

            # CIQ returns weekly data; generate evenly-spaced dates
            end = pd.Timestamp(END_DATE)
            start = end - pd.DateOffset(years=1)
            dates = pd.date_range(start=start, end=end, periods=n)

            df = pd.DataFrame({
                'Date': dates,
                'EV/EBITDA': ev_ebitda,
                'EV/Revenue': ev_rev,
            })
            all_data[company] = df
            print(f"    {n} data points")
        else:
            print(f"    No data")

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


def plot_charts(all_data):
    """Create two stacked charts: EV/EBITDA and EV/Revenue."""
    fig, axes = plt.subplots(2, 1, figsize=(14, 10))

    palette = plt.cm.tab10(np.linspace(0, 0.6, len(COMPANIES)))
    color_map = {name: palette[i] for i, (_, name) in enumerate(COMPANIES)}

    for ax, metric in zip(axes, ['EV/EBITDA', 'EV/Revenue']):
        for company, df in sorted(all_data.items()):
            series = df[['Date', metric]].dropna()
            if len(series) == 0:
                continue
            ax.plot(series['Date'], series[metric],
                    label=company, linewidth=2,
                    color=color_map.get(company, 'gray'))

        ax.set_ylabel(f'{metric}x')
        ax.set_title(f'{metric} (LTM) — Weekly, Past Year', fontsize=13, fontweight='bold')
        ax.legend(loc='best', fontsize=9, framealpha=0.9)
        ax.grid(True, alpha=0.3, linestyle='--')
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))

    fig.autofmt_xdate()
    plt.tight_layout()
    plt.savefig(CHART_PATH, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"\nChart saved: {CHART_PATH}")


def print_summary(all_data):
    """Print summary table."""
    print("\n" + "=" * 80)
    print(f"{'Company':<18} {'Metric':<14} {'Current':>9} {'1Y Low':>9} {'1Y High':>9} {'Points':>7}")
    print("-" * 80)
    for company, df in sorted(all_data.items()):
        for metric in ['EV/EBITDA', 'EV/Revenue']:
            vals = df[metric].dropna()
            if len(vals) == 0:
                continue
            print(f"{company:<18} {metric:<14} {vals.iloc[-1]:>8.1f}x {vals.min():>8.1f}x "
                  f"{vals.max():>8.1f}x {len(vals):>7}")
    print("=" * 80)


if __name__ == "__main__":
    print("=" * 60)
    print("  EV Multiples Trend Chart (Direct from CIQ)")
    print(f"  Companies: {', '.join(n for _, n in COMPANIES)}")
    print(f"  Period: {LOOKBACK} to {END_DATE}")
    print("=" * 60)

    create_workbook()
    all_data = open_refresh_and_extract()

    if all_data:
        plot_charts(all_data)
        print_summary(all_data)
    else:
        print("\nNo data extracted.")
