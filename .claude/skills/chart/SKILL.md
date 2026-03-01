---
name: chart
description: Create professional time-series charts from S&P Capital IQ data. Use when the user wants stock price charts, valuation multiple trends, financial metric charts, or asks to "chart" or "plot" company data.
argument-hint: [tickers...] [metric] [--period 1Y] [--bar] [--indexed] [--dual-axis]
allowed-tools: Agent
user-invocable: true
---

# Capital IQ Chart Tool

Create professional IB-quality charts from Capital IQ data via the `capiq-analyst` subagent, which calls the `capiq` MCP server's `pull_chart_data` tool.

## Arguments from user: $ARGUMENTS

Parse the user's request for:

1. **Tickers** — extract all company identifiers. Use `EXCHANGE:TICKER` format (e.g., `NYSE:HAL`, `TSX:PD`, `NASDAQGS:BKR`). Plain US tickers work without prefix.
2. **Metric(s)** — map natural language to CIQ mnemonics (see reference below). Max 2 metrics (2 required for dual-axis).
3. **Metric type** — determine from the metric name:
   - `"market"` → stock prices, TEV, market cap (daily date-ranged data)
   - `"multiple"` → EV/EBITDA, EV/Revenue, P/E (pre-calculated multiples with period type)
   - `"financial"` → revenue, EBITDA, margins, capex (period-offset quarterly/annual data)
4. **Chart type** — default based on metric type:
   - `"line"` → default for market and multiple data
   - `"bar"` → default for financial data
   - `"line_marker"` → optional for financial (line with data point markers)
   - `"dual_axis"` → when user asks to overlay 2 metrics with different scales
5. **Date range / periods**:
   - Market/multiple: `start_date` and `end_date` in M/D/YYYY (default: 1 year lookback)
   - Financial: `frequency` ("Q" or "Y") and `num_periods` (default: 12 quarters)
6. **Indexed** — if user says "indexed" or "relative performance", set `indexed: true` (normalizes to base 100)
7. **Period type** — for multiples: `IQ_LTM` (default), `IQ_NTM`, `IQ_FY`, etc.
8. **Currency** — default is `USD` unless specified

## Metric Mapping Reference

### Market data (metric_type: "market")
| User says | CIQ mnemonic |
|---|---|
| stock price, share price, close price | `IQ_CLOSEPRICE` |
| enterprise value, TEV | `IQ_TEV` |
| market cap | `IQ_MARKETCAP` |

### Multiples (metric_type: "multiple")
| User says | CIQ mnemonic | Period type |
|---|---|---|
| EV/EBITDA | `IQ_TEV_EBITDA` | `IQ_LTM` |
| forward EV/EBITDA | `IQ_TEV_EBITDA_FWD` | `IQ_NTM` |
| EV/Revenue | `IQ_TEV_TOTAL_REV` | `IQ_LTM` |
| forward EV/Revenue | `IQ_TEV_TOTAL_REV_FWD` | `IQ_NTM` |
| P/E | `IQ_PE_EXCL` | `IQ_LTM` |
| forward P/E | `IQ_PE_EXCL_FWD` | `IQ_NTM` |
| P/BV, price to book | `IQ_PBV_X` | `IQ_LTM` |

### Financial data (metric_type: "financial")
| User says | CIQ mnemonic |
|---|---|
| revenue, sales | `IQ_TOTAL_REV` |
| EBITDA | `IQ_EBITDA` |
| gross margin | `IQ_GROSS_MARGIN` |
| EBITDA margin | `IQ_EBITDA_MARGIN` |
| net income | `IQ_NI` |
| capex, capital expenditure | `IQ_CAPEX` |
| free cash flow, FCF | `IQ_LEVERED_FCF` |
| EPS | `IQ_DILUTED_EPS_EXCL` |

## Ticker format guide

- **U.S. (NYSE):** `NYSE:HAL`, `NYSE:SLB`
- **U.S. (NASDAQ):** `NASDAQGS:BKR`, `NASDAQGS:LBRT`
- **Canada (TSX):** `TSX:PD`, `TSX:TCW`, `TSX:CEU`
- **Plain tickers** (US-listed): `DSGX`, `MANH`, `ROP`
- **OTC:** `CMDXF`, `CNSWF`
- **Class shares:** `TSX:CCL.B`, `TSX:BBD.B`
- **Australian:** Plain ticker (e.g., `WTC`)

## Execution

Spawn the `capiq:capiq-analyst` subagent via the Agent tool with:
- The parsed parameters (tickers, metrics, metric_type, chart_type, etc.)
- Instruction to call `pull_chart_data` and report the chart path + any resolution notes

```
subagent_type: "capiq-analyst"
model: "sonnet"
prompt: "Call pull_chart_data with: tickers=[...], metrics=[...], metric_type=..., chart_type=..., [other params]. Report the chart path and any resolution issues."
```

**IMPORTANT:** Use `model: "sonnet"` — the chart skill requires judgement for metric mapping and chart type selection that Haiku may not handle well.

The subagent will:
1. Call `mcp__capiq__pull_chart_data` with the parameters
2. Return the chart path and resolution notes

**Present the chart path to the user** so they can view the PNG. If the user asks about the data, summarize key values from the returned data dict.

## Defaults Summary

| Setting | Default | Override |
|---|---|---|
| Currency | **USD** | "in CAD" or "--currency CAD" |
| Chart type (market/multiple) | **line** | "--bar", "--dual-axis" |
| Chart type (financial) | **bar** | "--line", "--line-marker" |
| Indexed | **OFF** | "--indexed" or "indexed" |
| Period type (multiples) | **IQ_LTM** | "forward" or "NTM" → IQ_NTM |
| Frequency (financial) | **Q** (quarterly) | "--annual" or "--yearly" → Y |
| Num periods (financial) | **12** | "--periods N" |
| Lookback (market/multiple) | **1 year** | "--period 6M", "--from M/D/YYYY" |
