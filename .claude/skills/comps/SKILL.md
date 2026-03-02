---
name: comps
description: Pull a comparable companies analysis table from S&P Capital IQ Pro. Use when the user provides company tickers and wants a comp table, valuation multiples, or asks to "pull comps".
argument-hint: [tickers...] [--currency USD] [--excluding-leases] [--no-lease-adjust]
allowed-tools: Bash
user-invocable: true
---

# Capital IQ Comparable Companies Analysis

Pull comp tables from Capital IQ via `capiq comps` CLI and format results as a clean markdown table.

## Arguments from user: $ARGUMENTS

Parse the user's request for:
1. **Tickers** — extract all company identifiers. Format: `EXCHANGE:TICKER` (e.g., `NYSE:HAL`, `TSX:PD`, `NASDAQGS:BKR`). If the user gives plain tickers without exchange prefix, use your knowledge to add the correct exchange prefix. Common exchanges: NYSE, NASDAQGS (NASDAQ Global Select), NasdaqGM (NASDAQ Global Market), TSX, TSXV.
2. **Currency** — default is **CAD** unless the user specifies otherwise.
3. **Mode** — default is **lease-adjusted** unless user says "excluding leases" or "ex-leases".
4. **NTM lease adjustment** — default is **ON**. User can disable with "no lease adjust" or "straight consensus NTM".
5. **Groups** — if the user specifies categories/sectors, pass them via `--groups`.
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

Run `capiq comps --help` if you need to discover available flags.

Build the command:
```bash
capiq comps TICKER1 TICKER2 ... [--currency USD] [--mode excluding-leases] [--no-lease-adjust-ntm] [--groups "Label: T1 T2" "Label: T3 T4"] [--date M/D/YYYY]
```

The default output is a markdown table printed to stdout. The CLI logs progress to stderr.

Present the markdown table output directly to the user. If there are resolution warnings (RESOLVED or FAILED tickers), report them.

## Defaults summary

| Setting | Default | Override |
|---|---|---|
| Currency | **CAD** | "in USD" or "--currency USD" |
| Mode | **Lease Adjusted** | "excluding leases" or "--mode excluding-leases" |
| NTM lease adj | **ON** | "no lease adjust" or "--no-lease-adjust-ntm" |
| Date | **Today** | "as of M/D/YYYY" or "--date M/D/YYYY" |
