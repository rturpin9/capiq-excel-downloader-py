---
name: comps
description: Pull a comparable companies analysis table from S&P Capital IQ Pro. Use when the user provides company tickers and wants a comp table, valuation multiples, or asks to "pull comps".
argument-hint: [tickers...] [--currency USD] [--excluding-leases] [--no-lease-adjust]
allowed-tools: Agent
user-invocable: true
---

# Capital IQ Comparable Companies Analysis

Pull comp tables from Capital IQ via the `capiq-analyst` subagent, which calls the `capiq` MCP server to retrieve data and formats it as a clean markdown table.

## Arguments from user: $ARGUMENTS

Parse the user's request for:
1. **Tickers** — extract all company identifiers. Format: `EXCHANGE:TICKER` (e.g., `NYSE:HAL`, `TSX:PD`, `NASDAQGS:BKR`). If the user gives plain tickers without exchange prefix, use your knowledge to add the correct exchange prefix. Common exchanges: NYSE, NASDAQGS (NASDAQ Global Select), NasdaqGM (NASDAQ Global Market), TSX, TSXV.
2. **Currency** — default is **CAD** unless the user specifies otherwise.
3. **Mode** — default is **lease-adjusted** unless user says "excluding leases" or "ex-leases".
4. **NTM lease adjustment** — default is **ON**. User can disable with "no lease adjust" or "straight consensus NTM".
5. **Groups** — if the user specifies categories/sectors, pass them to the agent.
6. **Date** — defaults to today. User can specify with "as of [date]".

## Ticker format guide

- **U.S. (NYSE):** `NYSE:HAL`, `NYSE:SLB`, `NYSE:NOV`
- **U.S. (NASDAQ):** `NASDAQGS:BKR`, `NASDAQGS:LBRT`, `NasdaqGM:WLDN`
- **Canada (TSX):** `TSX:PD`, `TSX:TCW`, `TSX:CEU`, `TSX:STN`
- **Plain tickers** (US-listed, no exchange needed): `DSGX`, `MANH`, `ROP`
- **OTC:** `CMDXF`, `CNSWF`
- **Class shares:** `TSX:CCL.B`, `TSX:BBD.B`, `TSX:RPI.UN`
- **Australian:** Plain ticker works (e.g., `WTC`)

### Known ticker gotchas

| Company | Wrong | Correct |
|---|---|---|
| Baker Hughes | NYSE:BKR | NASDAQGS:BKR |
| Canadian tickers | PD, TCW | TSX:PD, TSX:TCW |
| NASDAQ Global Select | NASDAQ:LBRT | NASDAQGS:LBRT |
| Class shares | TSX:CCL | TSX:CCL.B |

## Execution

Spawn the `capiq:capiq-analyst` subagent via the Agent tool with:
- The list of tickers (with exchange prefixes resolved)
- Currency, mode, lease adjustment preference, groups, and date
- Instruction to call the `pull_comps` MCP tool and format results as a markdown table

Example Agent tool call:

```
subagent_type: "capiq-analyst"
prompt: "Pull a comp table for the following tickers: NYSE:HAL, NYSE:SLB, TSX:PD, TSX:TCW. Currency: CAD, mode: lease-adjusted, NTM lease adjustment: ON. Format as a markdown table with summary statistics."
```

The subagent will:
1. Call `mcp__capiq__pull_comps` with the tickers and parameters
2. Format the result as a professional markdown table
3. Return the table to you

**Present the agent's formatted table directly to the user.** Include any resolution notes or warnings the agent reports.

## Defaults summary

| Setting | Default | Override |
|---|---|---|
| Currency | **CAD** | "in USD" or "--currency USD" |
| Mode | **Lease Adjusted** | "excluding leases" or "--excluding-leases" |
| NTM lease adj | **ON** | "no lease adjust" or "--no-lease-adjust" |
| Date | **Today** | "as of M/D/YYYY" |
