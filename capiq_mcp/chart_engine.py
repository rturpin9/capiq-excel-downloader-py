"""Core chart engine — builds CIQ workbooks, extracts time-series data, renders charts.

Supports three metric types:
  - market: CIQRANGE with date range (stock prices, TEV, market cap)
  - multiple: CIQRANGEV with period type + date range (EV/EBITDA, P/E, etc.)
  - financial: CIQRANGE with period offset (revenue, EBITDA, etc.)

All COM/Excel work follows the same pattern as comps_engine.py:
  build workbook -> launch isolated Excel -> RefreshSheet -> poll -> extract -> close.

Chart rendering uses matplotlib with an IB-quality professional template.
"""
from __future__ import annotations

import logging
import os
import time
from datetime import datetime, timedelta

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from capiq_mcp.comps_engine import safe_float, is_error_value

log = logging.getLogger("capiq_mcp")

# ── IB-quality chart style constants ──────────────────────────────────────

IB_COLORS = [
    "#1B3A5C",  # navy
    "#2A7B8E",  # teal
    "#5C6B7A",  # slate
    "#8B2252",  # burgundy
    "#2D5A3D",  # forest
    "#C4922A",  # amber
    "#6B4C8A",  # plum
    "#A05A2C",  # copper
]

CHART_DEFAULTS = {
    "figsize_line": (14, 7),
    "figsize_bar": (14, 7),
    "figsize_dual": (14, 8),
    "dpi": 150,
    "title_size": 13,
    "axis_label_size": 10,
    "tick_size": 9,
    "legend_size": 9,
    "source_size": 8,
    "line_width": 2.0,
    "grid_alpha": 0.3,
    "font_family": "sans-serif",
}

VALID_METRIC_TYPES = ("market", "multiple", "financial")
VALID_CHART_TYPES = ("line", "bar", "line_marker", "dual_axis")

# ── Formula builders (by metric type) ─────────────────────────────────────


def _build_market_formula(
    id_ref: str, metric: str, start: str, end: str, label: str
) -> str:
    """CIQRANGE formula for market data (date range)."""
    return (
        f'=CIQRANGE({id_ref},"{metric}","{start}","{end}"'
        f',,,,,,"{label}")'
    )


def _build_multiple_formula(
    id_ref: str, metric: str, period_type: str, start: str, end: str, label: str
) -> str:
    """CIQRANGEV formula for pre-calculated multiples (period type + date range).

    CIQRANGEV params: id(1), metric(2), period_type(3), start(4), end(5),
                      currency(6), primary(7), metatag(8), label(9)
    """
    return (
        f'=CIQRANGEV({id_ref},"{metric}",{period_type}'
        f',"{start}","{end}",,,,"{label}")'
    )


def _build_financial_formula(
    id_ref: str, metric: str, freq_period: str, label: str
) -> str:
    """CIQRANGE formula for financial data (period offset).

    CIQRANGE params: id(1), metric(2), start_period(3), end_period(4),
                     periodicity(5), report_type(6), primary(7), metatag(8),
                     currency(9), label(10)
    """
    return (
        f'=CIQRANGE({id_ref},"{metric}",{freq_period}'
        f',,,,,,,"{label}")'
    )


# ── Workbook creation ─────────────────────────────────────────────────────


def _build_workbook(
    path: str,
    tickers: list[str],
    metrics: list[str],
    metric_type: str,
    start_date: str | None,
    end_date: str | None,
    period_type: str,
    frequency: str,
    num_periods: int,
) -> dict:
    """Build a two-sheet workbook (Resolve + Data) and return column metadata.

    Returns
    -------
    dict with keys:
        columns: list of (col_index, ticker, metric, label) tuples
        num_tickers: int
        num_metrics: int
    """
    wb = Workbook()

    # ── Sheet 1: Resolve ──
    ws_resolve = wb.active
    ws_resolve.title = "Resolve"
    ws_resolve.cell(row=1, column=1, value="Input")
    ws_resolve.cell(row=1, column=2, value="CIQRANGEA_Formula")
    ws_resolve.cell(row=1, column=3, value="Resolved_IQ_ID")
    ws_resolve.cell(row=1, column=4, value="Company_Name")

    for idx, ticker in enumerate(tickers):
        row = idx + 2
        # CIQRANGEA doesn't handle EXCHANGE:TICKER format — strip prefix
        lookup_id = ticker.split(":", 1)[1] if ":" in ticker else ticker
        ws_resolve.cell(row=row, column=1, value=ticker)
        ws_resolve.cell(
            row=row, column=2,
            value=f'=CIQRANGEA("{lookup_id}","IQ_COMPANY_ID_QUICK_MATCH",1,1)',
        )
        # Col C: blank — CIQRANGEA spill target
        ws_resolve.cell(
            row=row, column=4,
            value=f'=CIQ($C${row},"IQ_COMPANY_NAME")',
        )

    # ── Sheet 2: Data ──
    ws_data = wb.create_sheet(title="Data")

    columns = []
    col = 1

    # Data sheet uses tickers directly (not cross-sheet references to Resolve)
    # because CIQ UDFs don't reliably chain cross-sheet dependencies.
    # The Resolve sheet runs independently for name resolution only.

    if metric_type == "market":
        # One column per (ticker, metric) — data expands downward
        for metric in metrics:
            for t_idx, ticker in enumerate(tickers):
                id_ref = f'"{ticker}"'
                label = f"{ticker} {metric}"
                formula = _build_market_formula(
                    id_ref, metric, start_date, end_date, label
                )
                ws_data.cell(row=1, column=col, value=formula)
                columns.append((col, ticker, metric, label))
                col += 1

    elif metric_type == "multiple":
        # One column per (ticker, metric) — data expands downward
        for metric in metrics:
            for t_idx, ticker in enumerate(tickers):
                id_ref = f'"{ticker}"'
                label = f"{ticker} {metric}"
                formula = _build_multiple_formula(
                    id_ref, metric, period_type, start_date, end_date, label
                )
                ws_data.cell(row=1, column=col, value=formula)
                columns.append((col, ticker, metric, label))
                col += 1

    elif metric_type == "financial":
        # Grouped by company: date col + metric cols
        freq_prefix = "IQ_FQ" if frequency == "Q" else "IQ_FY"
        period_offset = f"{freq_prefix}-{num_periods}"

        for t_idx, ticker in enumerate(tickers):
            id_ref = f'"{ticker}"'

            # Date column
            date_label = f"{ticker} Date"
            date_formula = _build_financial_formula(
                id_ref, "IQ_PERIODDATE_BS", period_offset, date_label
            )
            ws_data.cell(row=1, column=col, value=date_formula)
            columns.append((col, ticker, "_DATE_", date_label))
            col += 1

            # Metric columns
            for metric in metrics:
                label = f"{ticker} {metric}"
                formula = _build_financial_formula(
                    id_ref, metric, period_offset, label
                )
                ws_data.cell(row=1, column=col, value=formula)
                columns.append((col, ticker, metric, label))
                col += 1

    wb.save(path)
    total_cols = col - 1
    log.info(
        "Created chart workbook: %s (%d tickers, %d metrics, %d data columns)",
        path, len(tickers), len(metrics), total_cols,
    )

    return {
        "columns": columns,
        "num_tickers": len(tickers),
        "num_metrics": len(metrics),
    }


# ── Resolution reading ────────────────────────────────────────────────────


def _read_resolution(ws_resolve, num_tickers: int) -> tuple[list, list]:
    """Read resolved IQ IDs and company names from the Resolve sheet.

    Returns (resolution_notes, ticker_names).
    """
    resolution_notes = []
    ticker_names = {}

    if num_tickers == 1:
        iq_id = ws_resolve.Cells(2, 3).Value
        company_name = ws_resolve.Cells(2, 4).Value
        ticker = ws_resolve.Cells(2, 1).Value

        if iq_id and not is_error_value(iq_id):
            name = company_name if isinstance(company_name, str) and not is_error_value(company_name) else None
            resolution_notes.append({"ticker": ticker, "status": "OK", "iq_id": str(iq_id)})
            if name:
                ticker_names[ticker] = name
        else:
            resolution_notes.append({"ticker": ticker, "status": "FAILED"})
    else:
        # Batch read columns C-D
        vals = ws_resolve.Range(
            ws_resolve.Cells(2, 1),
            ws_resolve.Cells(1 + num_tickers, 4),
        ).Value

        for idx in range(num_tickers):
            ticker = vals[idx][0]
            iq_id = vals[idx][2]
            company_name = vals[idx][3]

            if iq_id and not is_error_value(iq_id):
                name = company_name if isinstance(company_name, str) and not is_error_value(company_name) else None
                resolution_notes.append({"ticker": ticker, "status": "OK", "iq_id": str(iq_id)})
                if name:
                    ticker_names[ticker] = name
            else:
                # CIQRANGEA didn't resolve, but data formulas use ticker directly
                # so data may still be present
                resolution_notes.append({"ticker": ticker, "status": "UNRESOLVED",
                                          "note": "Name lookup failed; data uses ticker directly"})

    return resolution_notes, ticker_names


def _display_name(ticker: str, ticker_names: dict) -> str:
    """Get a display-friendly name for a ticker.

    Uses resolved company name if available, otherwise strips exchange prefix
    (e.g. 'TSX:CSU' -> 'CSU').
    """
    name = ticker_names.get(ticker)
    if name:
        return name
    # Strip EXCHANGE: prefix for cleaner labels
    if ":" in ticker:
        return ticker.split(":", 1)[1]
    return ticker


# ── Data extraction ───────────────────────────────────────────────────────


def _find_last_data_row(ws_com, col: int, max_scan: int = 500) -> int:
    """Find the last row with data in a given column using End(xlUp)."""
    try:
        last_row = ws_com.Cells(ws_com.Rows.Count, col).End(-4162).Row  # xlUp
        return last_row
    except Exception:
        # Fallback: scan downward
        for row in range(max_scan, 1, -1):
            val = ws_com.Cells(row, col).Value
            if val is not None:
                return row
        return 1


def _read_column_data(ws_com, col: int, last_row: int) -> list:
    """Read a column from row 2 to last_row as a flat list."""
    if last_row < 2:
        return []
    if last_row == 2:
        val = ws_com.Cells(2, col).Value
        return [val]
    vals = ws_com.Range(
        ws_com.Cells(2, col),
        ws_com.Cells(last_row, col),
    ).Value
    return [v[0] if isinstance(v, tuple) else v for v in vals]


def _extract_market_or_multiple_data(
    ws_com, columns: list, tickers: list[str], metrics: list[str],
) -> dict:
    """Extract data for market or multiple metric types.

    Returns {ticker: {"dates": [...], metric: [...]}} where dates are
    synthesized from the data point count + known date range.
    """
    data = {}

    for col_idx, ticker, metric, label in columns:
        last_row = _find_last_data_row(ws_com, col_idx)
        raw_values = _read_column_data(ws_com, col_idx, last_row)
        values = [safe_float(v) for v in raw_values]

        if ticker not in data:
            data[ticker] = {"_num_points": len(values)}
        data[ticker][metric] = values
        # Track max length for date synthesis
        if len(values) > data[ticker]["_num_points"]:
            data[ticker]["_num_points"] = len(values)

    return data


def _extract_financial_data(
    ws_com, columns: list, tickers: list[str], metrics: list[str],
) -> dict:
    """Extract financial data grouped by company (date col + metric cols).

    Returns {ticker: {"dates": [...], metric: [...]}}.
    """
    data = {}

    for col_idx, ticker, metric, label in columns:
        last_row = _find_last_data_row(ws_com, col_idx)
        raw_values = _read_column_data(ws_com, col_idx, last_row)

        if ticker not in data:
            data[ticker] = {}

        if metric == "_DATE_":
            # Parse dates — CIQ returns Excel serial dates or date strings
            dates = []
            for v in raw_values:
                if isinstance(v, (int, float)) and v > 30000:
                    # Excel serial date
                    try:
                        dt = datetime(1899, 12, 30) + timedelta(days=int(v))
                        dates.append(dt)
                    except Exception:
                        dates.append(None)
                elif isinstance(v, str):
                    try:
                        dates.append(pd.to_datetime(v))
                    except Exception:
                        dates.append(None)
                elif hasattr(v, "year"):
                    # Already a datetime
                    dates.append(v)
                else:
                    dates.append(None)
            data[ticker]["dates"] = dates
        else:
            data[ticker][metric] = [safe_float(v) for v in raw_values]

    return data


# ── Polling ───────────────────────────────────────────────────────────────


def _poll_chart_data(ws_com, columns: list, max_wait: float) -> None:
    """Poll data sheet until formulas resolve or timeout."""
    log.info("Waiting for %d data columns (max %ds)...", len(columns), max_wait)
    start = time.monotonic()

    # Check first column as sentinel
    sentinel_col = columns[0][0]

    while True:
        elapsed = time.monotonic() - start
        if elapsed > max_wait:
            log.warning("Chart data timeout after %ds — proceeding with available data", max_wait)
            break

        val = ws_com.Cells(2, sentinel_col).Value
        if val is not None:
            if isinstance(val, str) and val.upper() in ("#PEND", "#REFRESH"):
                pass  # still pending
            elif isinstance(val, (int, float)) or (isinstance(val, str) and val.upper() not in ("#PEND", "#REFRESH")):
                # Data arrived — give extra time for all columns to settle
                log.info("Sentinel data arrived after %.0fs, waiting 15s for all columns...", elapsed)
                time.sleep(15)
                break

        if int(elapsed) % 15 < 5:
            log.debug("%ds: sentinel=%r", elapsed, val)

        time.sleep(5)


# ── Chart rendering ───────────────────────────────────────────────────────


def _apply_chart_style(ax, title: str, ylabel: str) -> None:
    """Apply IB-quality styling to an axes object."""
    ax.set_title(title, fontsize=CHART_DEFAULTS["title_size"], fontweight="bold", pad=12)
    ax.set_ylabel(ylabel, fontsize=CHART_DEFAULTS["axis_label_size"])
    ax.tick_params(labelsize=CHART_DEFAULTS["tick_size"])
    ax.grid(True, alpha=CHART_DEFAULTS["grid_alpha"], linestyle="--", color="#cccccc")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.margins(x=0)  # Lines start at y-axis, no left/right padding
    ax.legend(loc="best", fontsize=CHART_DEFAULTS["legend_size"], framealpha=0.9)


def _add_source_label(fig) -> None:
    """Add 'Source: S&P Capital IQ' attribution."""
    fig.text(
        0.98, 0.01, "Source: S&P Capital IQ",
        ha="right", va="bottom",
        fontsize=CHART_DEFAULTS["source_size"],
        fontstyle="italic", color="gray",
    )


def _format_date_axis(ax, num_points: int) -> None:
    """Auto-format date axis based on data density."""
    if num_points > 200:
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    elif num_points > 50:
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    else:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %d"))


def _render_line_chart(
    data: dict,
    metrics: list[str],
    ticker_names: dict,
    start_date: str,
    end_date: str,
    indexed: bool,
    title: str,
    output_path: str,
) -> str:
    """Render a line chart for market or multiple data."""
    fig, ax = plt.subplots(figsize=CHART_DEFAULTS["figsize_line"])
    color_idx = 0

    for ticker, ticker_data in data.items():
        num_points = ticker_data.get("_num_points", 0)
        if num_points == 0:
            continue

        # Synthesize dates from known range
        end_dt = pd.Timestamp(end_date)
        start_dt = pd.Timestamp(start_date)
        dates = pd.date_range(start=start_dt, end=end_dt, periods=num_points)

        display_name = _display_name(ticker, ticker_names)

        for metric in metrics:
            values = ticker_data.get(metric, [])
            if not values:
                continue

            series = pd.Series(values, index=dates[:len(values)])
            series = series.dropna()
            if len(series) == 0:
                continue

            if indexed:
                first_valid = series.first_valid_index()
                if first_valid is not None:
                    base = series[first_valid]
                    if base != 0:
                        series = (series / base) * 100

            label = display_name if len(metrics) == 1 else f"{display_name} — {metric}"
            color = IB_COLORS[color_idx % len(IB_COLORS)]
            ax.plot(series.index, series.values, label=label,
                    linewidth=CHART_DEFAULTS["line_width"], color=color)
            color_idx += 1

    if indexed:
        ax.axhline(y=100, color="gray", linestyle="--", linewidth=0.8, alpha=0.5)

    ylabel = "Indexed (Base = 100)" if indexed else (metrics[0] if len(metrics) == 1 else "Value")
    _apply_chart_style(ax, title, ylabel)
    _format_date_axis(ax, max((d.get("_num_points", 0) for d in data.values()), default=0))
    _add_source_label(fig)

    fig.autofmt_xdate()
    plt.tight_layout()
    plt.savefig(output_path, dpi=CHART_DEFAULTS["dpi"], bbox_inches="tight")
    plt.close()
    return output_path


def _render_bar_chart(
    data: dict,
    metrics: list[str],
    ticker_names: dict,
    frequency: str,
    title: str,
    output_path: str,
) -> str:
    """Render a grouped bar chart for financial data."""
    fig, ax = plt.subplots(figsize=CHART_DEFAULTS["figsize_bar"])

    # Collect all unique period dates across tickers for x-axis alignment
    all_dates = set()
    for ticker, ticker_data in data.items():
        dates = ticker_data.get("dates", [])
        for d in dates:
            if d is not None:
                all_dates.add(d)

    if not all_dates:
        log.warning("No dates found for bar chart")
        plt.close()
        return output_path

    sorted_dates = sorted(all_dates)

    # Format period labels
    period_labels = []
    for d in sorted_dates:
        if hasattr(d, "strftime"):
            if frequency == "Q":
                quarter = (d.month - 1) // 3 + 1
                period_labels.append(f"Q{quarter}'{d.strftime('%y')}")
            else:
                period_labels.append(f"FY{d.strftime('%Y')}")
        else:
            period_labels.append(str(d))

    x = np.arange(len(sorted_dates))
    tickers = list(data.keys())
    num_tickers = len(tickers)
    # For each metric, group bars by ticker
    total_groups = num_tickers * len(metrics)
    bar_width = 0.8 / max(total_groups, 1)

    color_idx = 0
    for m_idx, metric in enumerate(metrics):
        for t_idx, ticker in enumerate(tickers):
            ticker_data = data[ticker]
            dates = ticker_data.get("dates", [])
            values = ticker_data.get(metric, [])

            # Align values to sorted_dates
            aligned = []
            for target_date in sorted_dates:
                found = np.nan
                for i, d in enumerate(dates):
                    if d is not None and hasattr(d, "date") and hasattr(target_date, "date"):
                        if d.date() == target_date.date() and i < len(values):
                            found = values[i]
                            break
                    elif d == target_date and i < len(values):
                        found = values[i]
                        break
                aligned.append(found)

            offset = (m_idx * num_tickers + t_idx - total_groups / 2 + 0.5) * bar_width
            display_name = _display_name(ticker, ticker_names)
            label = display_name if len(metrics) == 1 else f"{display_name} — {metric}"
            color = IB_COLORS[color_idx % len(IB_COLORS)]

            ax.bar(x + offset, aligned, bar_width, label=label, color=color, alpha=0.85)
            color_idx += 1

    ax.set_xticks(x)
    ax.set_xticklabels(period_labels, fontsize=CHART_DEFAULTS["tick_size"])
    ylabel = metrics[0] if len(metrics) == 1 else "Value"
    _apply_chart_style(ax, title, ylabel)

    # Format y-axis with commas for large numbers
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, p: f"{v:,.0f}"))
    _add_source_label(fig)

    plt.tight_layout()
    plt.savefig(output_path, dpi=CHART_DEFAULTS["dpi"], bbox_inches="tight")
    plt.close()
    return output_path


def _render_line_marker_chart(
    data: dict,
    metrics: list[str],
    ticker_names: dict,
    frequency: str,
    title: str,
    output_path: str,
) -> str:
    """Render a line chart with markers for financial data."""
    fig, ax = plt.subplots(figsize=CHART_DEFAULTS["figsize_line"])
    color_idx = 0

    for ticker, ticker_data in data.items():
        dates = ticker_data.get("dates", [])
        display_name = _display_name(ticker, ticker_names)

        for metric in metrics:
            values = ticker_data.get(metric, [])
            if not values or not dates:
                continue

            # Pair dates with values, dropping None/NaN
            paired = [(d, v) for d, v in zip(dates, values) if d is not None and not np.isnan(v)]
            if not paired:
                continue

            plot_dates, plot_values = zip(*paired)
            label = display_name if len(metrics) == 1 else f"{display_name} — {metric}"
            color = IB_COLORS[color_idx % len(IB_COLORS)]

            ax.plot(plot_dates, plot_values, label=label,
                    linewidth=CHART_DEFAULTS["line_width"], color=color,
                    marker="o", markersize=5)
            color_idx += 1

    ylabel = metrics[0] if len(metrics) == 1 else "Value"
    _apply_chart_style(ax, title, ylabel)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, p: f"{v:,.0f}"))
    _add_source_label(fig)

    fig.autofmt_xdate()
    plt.tight_layout()
    plt.savefig(output_path, dpi=CHART_DEFAULTS["dpi"], bbox_inches="tight")
    plt.close()
    return output_path


def _render_dual_axis_chart(
    data: dict,
    metrics: list[str],
    ticker_names: dict,
    metric_type: str,
    start_date: str | None,
    end_date: str | None,
    frequency: str,
    title: str,
    output_path: str,
) -> str:
    """Render a dual-axis chart with two metrics on separate y-axes."""
    if len(metrics) != 2:
        raise ValueError("dual_axis chart requires exactly 2 metrics")

    fig, ax1 = plt.subplots(figsize=CHART_DEFAULTS["figsize_dual"])
    ax2 = ax1.twinx()

    axes = [ax1, ax2]
    line_styles = ["-", "--"]
    color_idx = 0

    for m_idx, metric in enumerate(metrics):
        ax = axes[m_idx]
        ls = line_styles[m_idx]

        for ticker, ticker_data in data.items():
            display_name = _display_name(ticker, ticker_names)
            label = f"{display_name} — {metric}"
            color = IB_COLORS[color_idx % len(IB_COLORS)]

            if metric_type == "financial":
                dates = ticker_data.get("dates", [])
                values = ticker_data.get(metric, [])
                if not values or not dates:
                    continue
                paired = [(d, v) for d, v in zip(dates, values) if d is not None and not np.isnan(v)]
                if not paired:
                    continue
                plot_dates, plot_values = zip(*paired)
                ax.plot(plot_dates, plot_values, label=label,
                        linewidth=CHART_DEFAULTS["line_width"], color=color,
                        linestyle=ls, marker="o" if m_idx == 1 else None,
                        markersize=4)
            else:
                num_points = ticker_data.get("_num_points", 0)
                values = ticker_data.get(metric, [])
                if not values or num_points == 0:
                    continue

                end_dt = pd.Timestamp(end_date)
                start_dt = pd.Timestamp(start_date)
                dates = pd.date_range(start=start_dt, end=end_dt, periods=num_points)

                series = pd.Series(values, index=dates[:len(values)]).dropna()
                if len(series) == 0:
                    continue

                ax.plot(series.index, series.values, label=label,
                        linewidth=CHART_DEFAULTS["line_width"], color=color,
                        linestyle=ls)

            color_idx += 1

        ax.set_ylabel(metric, fontsize=CHART_DEFAULTS["axis_label_size"])
        ax.tick_params(labelsize=CHART_DEFAULTS["tick_size"])

    # Styling
    ax1.set_title(title, fontsize=CHART_DEFAULTS["title_size"], fontweight="bold", pad=12)
    ax1.grid(True, alpha=CHART_DEFAULTS["grid_alpha"], linestyle="--", color="#cccccc")

    # Combined legend
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2,
               loc="best", fontsize=CHART_DEFAULTS["legend_size"], framealpha=0.9)

    _add_source_label(fig)

    fig.autofmt_xdate()
    plt.tight_layout()
    plt.savefig(output_path, dpi=CHART_DEFAULTS["dpi"], bbox_inches="tight")
    plt.close()
    return output_path


# ── JSON serialization ────────────────────────────────────────────────────


def _serialize_data(data: dict) -> dict:
    """Convert extracted data to JSON-serializable format."""
    serialized = {}
    for ticker, ticker_data in data.items():
        s = {}
        for key, values in ticker_data.items():
            if key == "_num_points":
                continue
            if key == "dates":
                s["dates"] = [
                    d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d)
                    for d in values
                ]
            else:
                s[key] = [
                    None if (isinstance(v, float) and np.isnan(v)) else v
                    for v in values
                ]
        serialized[ticker] = s
    return serialized


# ── Main pipeline ─────────────────────────────────────────────────────────


def run_chart(
    tickers: list[str],
    metrics: list[str],
    metric_type: str,
    start_date: str | None = None,
    end_date: str | None = None,
    period_type: str = "IQ_LTM",
    frequency: str = "Q",
    num_periods: int = 12,
    chart_type: str = "line",
    indexed: bool = False,
    currency: str = "USD",
    title: str | None = None,
    output_path: str | None = None,
    max_wait: int = 180,
) -> dict:
    """Full pipeline: build workbook -> Excel -> refresh -> extract -> chart -> return.

    Parameters
    ----------
    tickers : list[str]
        Company identifiers (EXCHANGE:TICKER or plain ticker).
    metrics : list[str]
        CIQ mnemonics (e.g. ["IQ_CLOSEPRICE"], ["IQ_TEV_EBITDA", "IQ_TEV_TOTAL_REV"]).
    metric_type : str
        "market", "multiple", or "financial".
    start_date : str or None
        Start date M/D/YYYY. Defaults: 1Y ago for market/multiple, None for financial.
    end_date : str or None
        End date M/D/YYYY. Defaults to today.
    period_type : str
        For multiples: IQ_LTM, IQ_NTM, etc. Default IQ_LTM.
    frequency : str
        For financial: "Q" (quarterly) or "Y" (annual). Default "Q".
    num_periods : int
        For financial: how many periods back. Default 12.
    chart_type : str
        "line", "bar", "line_marker", or "dual_axis".
    indexed : bool
        Index market data to base 100.
    currency : str
        Output currency. Default "USD".
    title : str or None
        Chart title. Auto-generated if None.
    output_path : str or None
        PNG output path. Default: C:\\Claude\\chart_output.png.
    max_wait : int
        Max seconds to wait for formula refresh.

    Returns
    -------
    dict with keys: chart_path, data, resolution_notes, chart_config.
    """
    import pythoncom
    pythoncom.CoInitialize()

    from capiq_excel.excel_lifecycle import launch_excel_isolated, close_session

    # ── Validate inputs ──
    if metric_type not in VALID_METRIC_TYPES:
        raise ValueError(f"Invalid metric_type '{metric_type}'. Use: {VALID_METRIC_TYPES}")
    if chart_type not in VALID_CHART_TYPES:
        raise ValueError(f"Invalid chart_type '{chart_type}'. Use: {VALID_CHART_TYPES}")
    if chart_type == "dual_axis" and len(metrics) != 2:
        raise ValueError("dual_axis chart requires exactly 2 metrics")

    # ── Resolve dates ──
    now = datetime.now()
    if end_date is None:
        end_date = f"{now.month}/{now.day}/{now.year}"

    if start_date is None:
        if metric_type in ("market", "multiple"):
            one_year_ago = now - timedelta(days=365)
            start_date = f"{one_year_ago.month}/{one_year_ago.day}/{one_year_ago.year}"
        # financial: start_date stays None (uses period offset instead)

    if output_path is None:
        output_path = os.path.abspath("chart_output.png")

    # ── Auto-generate title ──
    if title is None:
        metric_labels = " / ".join(m.replace("IQ_", "").replace("_", " ").title() for m in metrics)
        ticker_str = ", ".join(tickers[:4])
        if len(tickers) > 4:
            ticker_str += f" +{len(tickers) - 4} more"

        if indexed:
            title = f"Indexed {metric_labels} — {ticker_str}"
        elif metric_type == "financial":
            period_label = "Quarterly" if frequency == "Q" else "Annual"
            title = f"{metric_labels} ({period_label}) — {ticker_str}"
        else:
            title = f"{metric_labels} — {ticker_str}"

    # ── Build workbook ──
    xlsx_path = os.path.abspath("_chart_mcp_temp.xlsx")

    col_meta = _build_workbook(
        xlsx_path, tickers, metrics, metric_type,
        start_date, end_date, period_type, frequency, num_periods,
    )
    columns = col_meta["columns"]

    log.info(
        "Chart pull: %d tickers, %d metrics, type=%s, chart=%s",
        len(tickers), len(metrics), metric_type, chart_type,
    )

    # ── Launch Excel and process ──
    session = launch_excel_isolated(xlsx_path)
    try:
        excel = session.excel
        log.info("Connected to Excel %s for chart data", excel.Version)

        # Trigger refresh on both sheets
        try:
            excel.Run("SNLXLAddin.xla!RefreshWorkbook")
            log.info("RefreshWorkbook executed")
        except Exception as e:
            log.warning("RefreshWorkbook warning: %s", e)

        # Poll Data sheet for completion
        ws_data = session.workbook.Sheets("Data")
        _poll_chart_data(ws_data, columns, max_wait)

        # Read resolution from Resolve sheet
        ws_resolve = session.workbook.Sheets("Resolve")
        resolution_notes, ticker_names = _read_resolution(ws_resolve, len(tickers))

        # Extract time-series data
        log.info("Extracting chart data...")
        if metric_type in ("market", "multiple"):
            data = _extract_market_or_multiple_data(ws_data, columns, tickers, metrics)
        else:
            data = _extract_financial_data(ws_data, columns, tickers, metrics)

    finally:
        close_session(session, save=False, delete_workbook=True)

    # ── Render chart ──
    log.info("Rendering %s chart...", chart_type)

    if chart_type == "dual_axis":
        chart_path = _render_dual_axis_chart(
            data, metrics, ticker_names, metric_type,
            start_date, end_date, frequency, title, output_path,
        )
    elif chart_type == "bar":
        chart_path = _render_bar_chart(
            data, metrics, ticker_names, frequency, title, output_path,
        )
    elif chart_type == "line_marker":
        chart_path = _render_line_marker_chart(
            data, metrics, ticker_names, frequency, title, output_path,
        )
    else:  # "line"
        chart_path = _render_line_chart(
            data, metrics, ticker_names,
            start_date, end_date, indexed, title, output_path,
        )

    log.info("Chart saved: %s", chart_path)

    return {
        "chart_path": chart_path.replace("\\", "/"),
        "data": _serialize_data(data),
        "resolution_notes": resolution_notes,
        "chart_config": {
            "chart_type": chart_type,
            "metric_type": metric_type,
            "indexed": indexed,
            "currency": currency,
            "start_date": start_date,
            "end_date": end_date,
            "period_type": period_type if metric_type == "multiple" else None,
            "frequency": frequency if metric_type == "financial" else None,
            "num_periods": num_periods if metric_type == "financial" else None,
        },
    }
