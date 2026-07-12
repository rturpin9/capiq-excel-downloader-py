---
name: comps
description: Pull a comparable companies analysis table from S&P Capital IQ Pro. Use when the user provides company tickers and wants a comp table, valuation multiples, or asks to "pull comps". Can also output a formatted Excel (.xlsx) deliverable styled like the firm's comp-set template (Lease Adjusted + Excluding Leases tabs) via --xlsx — use when the user asks for an Excel file, a "formatted"/"polished" comp table, or something to drop into a deck or model.
argument-hint: [tickers...] [--currency USD] [--mode excluding-leases] [--no-lease-adjust-ntm] [--groups "Label: T1 T2"] [--xlsx out.xlsx] [--xlsx-sheets both] [--columns ...]
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
7. **Deliverable format** — if the user wants an Excel file, a "formatted" / "polished" comp table, something "to drop into a deck/model", or mentions the firm's comp set template, also write a formatted workbook with `--xlsx PATH`. If they want specific columns, pass `--columns` (see below).

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
capiq comps TICKER1 TICKER2 ... [--currency USD] [--mode excluding-leases] [--no-lease-adjust-ntm] [--groups "Label: T1 T2" "Label: T3 T4"] [--date M/D/YYYY] [--xlsx PATH] [--columns ...]
```

The default output is a markdown table printed to stdout. The CLI logs progress to stderr.

Present the markdown table output directly to the user. If there are resolution warnings (RESOLVED or FAILED tickers), report them.

**Reading the numbers:** all dollar figures are in **millions** of the chosen currency. A **negative Net Debt or ND/EBITDA is correct, not an error** — it means a **net cash** position (cash exceeds debt). Present it as-is (e.g. `-0.3x`); don't flag it or drop it.

## Formatted Excel deliverable (`--xlsx`)

When the user wants a deliverable (Excel file / "formatted" / "for the deck" / matching the firm comp set), add `--xlsx PATH`. This writes a presentation-ready workbook styled like the firm's "Comp Set" template: navy column headers, teal sector group headers, gold per-group Average/Median rows, navy Overall Average/Median, accounting / percent / multiple number formats, and **live formulas** for the derived columns (margins, EV multiples, ND/EBITDA) and the Average/Median rows — so the file recalculates if a number is edited. `--groups` produces the sector sections. Raw pulled metrics are written as values (sourced from Capital IQ).

**Both tabs by default.** Like the firm template, the workbook contains **two tabs — `Lease Adjusted` and `Excluding Leases`** — and the tool runs a separate pull for each (the `--mode` tab is the one active on open). To limit it, use `--xlsx-sheets {both|lease-adjusted|excluding-leases}` (default `both`). Pulling both tabs means two Excel round-trips, so it takes about twice as long; pass `--xlsx-sheets lease-adjusted` (or the mode you want) for a single-tab file. If the output file is open in Excel the write fails with a clear "file may be open" message — close it and re-run.

The XLSX is written **in addition** to the markdown (still present that). Example:
```bash
capiq comps NYSE:HAL TSX:PD NASDAQGS:BKR --currency USD \
  --groups "U.S.: NYSE:HAL NASDAQGS:BKR" "Canada: TSX:PD" \
  --xlsx "comps_OFS.xlsx"
```

### Column flexibility (`--columns`)

By default the XLSX shows the firm template's standard columns. To include/reorder a specific set, pass `--columns` with space-separated column keys (the `company` column is always included first). Run `capiq comps --list-columns [--mode ...]` to see all keys (e.g. `mkt_cap`, `tev`, `ltm_rev`, `ev_ltm_ebitda`, `pbv`, `nd_ebitda`, plus optional `capex`, `gross_margin`, `gaap`). Any `--extra` metric key is also usable as a column. Derived columns whose input columns are dropped fall back to static values automatically. Example:
```bash
capiq comps DSGX ROP MANH --extra roe rev-growth \
  --columns company mkt_cap tev ev_ltm_ebitda ev_ntm_ebitda roe rev-growth \
  --xlsx "comps_multiples.xlsx"
```

## Defaults summary

| Setting | Default | Override |
|---|---|---|
| Currency | **CAD** | "in USD" or "--currency USD" |
| Mode | **Lease Adjusted** | "excluding leases" or "--mode excluding-leases" |
| NTM lease adj | **ON** | "no lease adjust" or "--no-lease-adjust-ntm" |
| Date | **Today** | "as of M/D/YYYY" or "--date M/D/YYYY" |
| Output | **Markdown only** | add `--xlsx PATH` for a formatted Excel deliverable |
| XLSX tabs | **Both** (Lease Adjusted + Excluding Leases) | `--xlsx-sheets lease-adjusted` / `excluding-leases` for one tab |
| XLSX columns | **Firm template set** | `--columns ...` (see `--list-columns`); `--extra` keys also usable |
