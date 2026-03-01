"""Capital IQ MCP server — stdio transport.

Provides three tools:
  - pull_comps: Full comparable companies analysis via SPG formulas
  - pull_chart_data: Time-series charts (market, multiple, financial data)
  - lookup_identifiers: Lightweight identifier resolution via CIQRANGEA

All COM/Excel work runs in a thread via asyncio.to_thread().
An asyncio.Lock() serializes Excel sessions (one at a time).

Logging goes to stderr only — stdout is reserved for MCP JSON-RPC protocol.
"""
from __future__ import annotations

import asyncio
import logging
import os
import sys

from mcp.server.fastmcp import FastMCP

# ── Logging to stderr only ─────────────────────────────────────────────────

_log_file = os.path.join(os.path.expanduser("~"), "capiq_mcp_debug.log")
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    handlers=[
        logging.FileHandler(_log_file, mode="w"),
        logging.StreamHandler(sys.stderr),
    ],
)
log = logging.getLogger("capiq_mcp")

# ── MCP server ─────────────────────────────────────────────────────────────

mcp = FastMCP(
    "capiq",
    instructions="S&P Capital IQ data retrieval via Excel COM automation",
)

# Serialize Excel sessions — only one at a time
_excel_lock = asyncio.Lock()


@mcp.tool()
async def pull_comps(
    tickers: list[str],
    currency: str = "CAD",
    mode: str = "lease-adjusted",
    lease_adjust_ntm: bool = True,
    date: str | None = None,
    max_wait: int = 180,
) -> dict:
    """Pull a comparable companies analysis table from S&P Capital IQ.

    Builds an Excel workbook with SPG formulas, launches Excel via COM,
    waits for the CIQ Pro add-in to refresh all formulas, reads results,
    and returns structured data.

    Args:
        tickers: Company identifiers (e.g. ["NYSE:HAL", "TSX:PD", "DSGX"]).
                 Supports EXCHANGE:TICKER, plain tickers, company names,
                 CUSIPs, and ISINs (auto-resolved via CIQRANGEA fallback).
        currency: Output currency code (default "CAD").
        mode: "lease-adjusted" (default) or "excluding-leases".
        lease_adjust_ntm: Add TTM lease adj to NTM EBITDA for US GAAP
                          reporters (default True, only applies in lease-adjusted mode).
        date: As-of date in M/D/YYYY format (default: today).
        max_wait: Max seconds to wait for formula refresh (default 180).

    Returns:
        Dict with keys: mode, currency, date, lease_adjust_ntm, companies
        (list of company dicts with all metrics and derived multiples),
        and resolution_notes (list of per-ticker resolution status).
    """
    if not tickers:
        return {"error": "validation", "message": "No tickers provided"}

    if mode not in ("lease-adjusted", "excluding-leases"):
        return {"error": "validation",
                "message": f"Invalid mode '{mode}'. Use 'lease-adjusted' or 'excluding-leases'"}

    async with _excel_lock:
        try:
            # Lazy import to avoid COM init at module level
            from capiq_mcp.comps_engine import run_comps
            result = await asyncio.to_thread(
                run_comps,
                tickers=tickers,
                currency=currency,
                mode=mode,
                lease_adjust_ntm=lease_adjust_ntm,
                date=date,
                max_wait=max_wait,
            )
            return result
        except Exception as e:
            log.exception("pull_comps failed")
            return {"error": "execution", "message": str(e)}


@mcp.tool()
async def lookup_identifiers(
    identifiers: list[str],
    max_wait: int = 60,
) -> dict:
    """Resolve company identifiers to CIQ IQ IDs and company names.

    Accepts tickers, company names, CUSIPs, ISINs, or any identifier
    that CIQ's CIQRANGEA can resolve. Much faster than a full comp pull
    (~10s vs ~40s).

    Args:
        identifiers: List of identifiers to resolve
                     (e.g. ["NYSE:HAL", "Precision Drilling", "US4062161017"]).
        max_wait: Max seconds to wait for resolution (default 60).

    Returns:
        Dict with "results" key: list of dicts, each containing
        input, iq_id, company_name, and status ("OK" or "FAILED").
    """
    if not identifiers:
        return {"error": "validation", "message": "No identifiers provided"}

    async with _excel_lock:
        try:
            from capiq_mcp.id_lookup import run_id_lookup
            result = await asyncio.to_thread(
                run_id_lookup,
                identifiers=identifiers,
                max_wait=max_wait,
            )
            return result
        except Exception as e:
            log.exception("lookup_identifiers failed")
            return {"error": "execution", "message": str(e)}


@mcp.tool()
async def pull_chart_data(
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
    """Fetch time-series data from S&P Capital IQ and render a professional chart.

    Builds an Excel workbook with CIQ formulas, launches Excel via COM,
    waits for the CIQ Pro add-in to refresh all formulas, extracts data,
    renders a chart PNG, and returns both the chart path and underlying data.

    Args:
        tickers: Company identifiers (e.g. ["DSGX", "NYSE:HAL", "TSX:PD"]).
                 Supports EXCHANGE:TICKER, plain tickers, company names,
                 CUSIPs, and ISINs (auto-resolved via CIQRANGEA).
        metrics: CIQ mnemonics (e.g. ["IQ_CLOSEPRICE"], ["IQ_TEV_EBITDA"]).
                 1-2 metrics supported (2 required for dual_axis).
        metric_type: "market" (prices, TEV, market cap),
                     "multiple" (EV/EBITDA, P/E, etc.),
                     or "financial" (revenue, EBITDA, etc.).
        start_date: Start date M/D/YYYY (default: 1Y ago for market/multiple).
        end_date: End date M/D/YYYY (default: today).
        period_type: For multiples: IQ_LTM, IQ_NTM, IQ_FY, etc. Default IQ_LTM.
        frequency: For financial: "Q" (quarterly) or "Y" (annual). Default "Q".
        num_periods: For financial: how many periods back. Default 12.
        chart_type: "line", "bar", "line_marker", or "dual_axis".
        indexed: Index market data to base 100 (default false).
        currency: Output currency code (default "USD").
        title: Chart title (auto-generated if None).
        output_path: PNG file path (default: chart_output.png in working dir).
        max_wait: Max seconds to wait for formula refresh (default 180).

    Returns:
        Dict with keys: chart_path (PNG location), data (time-series per ticker),
        resolution_notes (per-ticker resolution status), chart_config.
    """
    if not tickers:
        return {"error": "validation", "message": "No tickers provided"}
    if not metrics:
        return {"error": "validation", "message": "No metrics provided"}
    if metric_type not in ("market", "multiple", "financial"):
        return {"error": "validation",
                "message": f"Invalid metric_type '{metric_type}'. Use 'market', 'multiple', or 'financial'"}
    if chart_type not in ("line", "bar", "line_marker", "dual_axis"):
        return {"error": "validation",
                "message": f"Invalid chart_type '{chart_type}'. Use 'line', 'bar', 'line_marker', or 'dual_axis'"}
    if chart_type == "dual_axis" and len(metrics) != 2:
        return {"error": "validation",
                "message": "dual_axis chart requires exactly 2 metrics"}

    async with _excel_lock:
        try:
            from capiq_mcp.chart_engine import run_chart
            result = await asyncio.to_thread(
                run_chart,
                tickers=tickers,
                metrics=metrics,
                metric_type=metric_type,
                start_date=start_date,
                end_date=end_date,
                period_type=period_type,
                frequency=frequency,
                num_periods=num_periods,
                chart_type=chart_type,
                indexed=indexed,
                currency=currency,
                title=title,
                output_path=output_path,
                max_wait=max_wait,
            )
            return result
        except Exception as e:
            log.exception("pull_chart_data failed")
            return {"error": "execution", "message": str(e)}


if __name__ == "__main__":
    mcp.run(transport="stdio")
