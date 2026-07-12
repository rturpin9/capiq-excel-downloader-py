---
name: chart
description: Create professional time-series charts from S&P Capital IQ data. Use when the user wants stock price charts, valuation multiple trends, financial metric charts, or asks to "chart" or "plot" company data.
argument-hint: [tickers...] [metric] [--period 1Y] [--bar] [--indexed] [--dual-axis]
allowed-tools: Bash
user-invocable: true
---

# Capital IQ Chart Tool

Create professional IB-quality charts from Capital IQ data via `capiq chart` CLI.

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
   - **Financial data has no continuous date axis**, so use `bar` or `line_marker` for `--metric-type financial`. A `line` chart requested with financial data is auto-converted to `line_marker` (a plain `line` would render empty).
5. **Date range / periods**:
   - Market/multiple: `--start-date` and `--end-date` in M/D/YYYY (default: 1 year lookback)
   - Financial: `--frequency` ("Q" or "Y") and `--num-periods` (default: 12 quarters)
6. **Indexed** — if user says "indexed" or "relative performance", add `--indexed` (normalizes to base 100)
7. **Period type** — for multiples: `--period-type IQ_LTM` (default), `IQ_NTM`, `IQ_FY`, etc.
8. **Currency** — default is `USD` unless specified

## Metric Mapping Reference

### Market data (--metric-type market)
| User says | CIQ mnemonic |
|---|---|
| stock price, share price, close price | `IQ_CLOSEPRICE` |
| enterprise value, TEV | `IQ_TEV` |
| market cap | `IQ_MARKETCAP` |

### Multiples (--metric-type multiple)
| User says | CIQ mnemonic | Period type |
|---|---|---|
| EV/EBITDA | `IQ_TEV_EBITDA` | `IQ_LTM` |
| forward EV/EBITDA | `IQ_TEV_EBITDA_FWD` | `IQ_NTM` |
| EV/Revenue | `IQ_TEV_TOTAL_REV` | `IQ_LTM` |
| forward EV/Revenue | `IQ_TEV_TOTAL_REV_FWD` | `IQ_NTM` |
| P/E | `IQ_PE_EXCL` | `IQ_LTM` |
| forward P/E | `IQ_PE_EXCL_FWD` | `IQ_NTM` |
| P/BV, price to book | `IQ_PBV_X` | `IQ_LTM` |

### Financial data (--metric-type financial)
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

Run `capiq chart --help` if you need to discover available flags.

Build the command:
```bash
capiq chart TICKER1 TICKER2 ... --metrics IQ_CLOSEPRICE --metric-type market [--chart-type line] [--start-date M/D/YYYY] [--end-date M/D/YYYY] [--indexed] [--currency USD] [--output path.png]
```

The CLI outputs JSON to stdout with `chart_path`, `data` (summary stats), and `chart_config`. Present the chart path to the user so they can view the PNG.

## Defaults Summary

| Setting | Default | Override |
|---|---|---|
| Currency | **USD** | "--currency CAD" |
| Chart type (market/multiple) | **line** | "--chart-type bar", "--chart-type dual_axis" |
| Chart type (financial) | **bar** | "--chart-type line_marker" (a requested "line" auto-converts to line_marker) |
| Indexed | **OFF** | "--indexed" |
| Period type (multiples) | **IQ_LTM** | "--period-type IQ_NTM" |
| Frequency (financial) | **Q** (quarterly) | "--frequency Y" |
| Num periods (financial) | **12** | "--num-periods N" |
| Lookback (market/multiple) | **1 year** | "--start-date M/D/YYYY" |
