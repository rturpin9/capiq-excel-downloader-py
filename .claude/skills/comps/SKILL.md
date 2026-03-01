---
name: comps
description: Pull a comparable companies analysis table from S&P Capital IQ Pro via SPG formulas. Use when the user provides company tickers and wants a comp table, valuation multiples, or asks to "pull comps".
argument-hint: [tickers...] [--currency USD] [--excluding-leases] [--no-lease-adjust]
allowed-tools: Bash, Read, Write, Edit
user-invocable: true
---

# Capital IQ Comparable Companies Analysis

Pull comp tables via SPG formulas using the `pull_comps.py` CLI tool.

## Quick reference

```bash
# Basic (defaults: CAD, lease-adjusted, NTM lease adj ON)
python C:/Claude/pull_comps.py NYSE:HAL NYSE:SLB TSX:PD TSX:TCW

# With groups and currency override
python C:/Claude/pull_comps.py NYSE:HAL NYSE:SLB TSX:PD TSX:TCW \
  --currency USD \
  --groups "U.S. OFS: NYSE:HAL NYSE:SLB" "Canadian OFS: TSX:PD TSX:TCW"

# Excluding leases mode
python C:/Claude/pull_comps.py NYSE:HAL NYSE:SLB --mode excluding-leases

# Disable NTM EBITDA lease adjustment
python C:/Claude/pull_comps.py NYSE:HAL NYSE:SLB --no-lease-adjust-ntm

# JSON output for further processing
python C:/Claude/pull_comps.py NYSE:HAL NYSE:SLB --json-output comps.json
```

## Arguments from user: $ARGUMENTS

Parse the user's request for:
1. **Tickers** — extract all company identifiers. Format: `EXCHANGE:TICKER` (e.g., `NYSE:HAL`, `TSX:PD`, `NASDAQGS:BKR`). If the user gives plain tickers without exchange prefix, use your knowledge to add the correct exchange prefix. Common exchanges: NYSE, NASDAQGS (NASDAQ Global Select), NasdaqGM (NASDAQ Global Market), TSX, TSXV.
2. **Currency** — default is **CAD** unless the user specifies otherwise. Always note in output which currency was used.
3. **Mode** — default is **lease-adjusted** unless user says "excluding leases" or "ex-leases".
4. **NTM lease adjustment** — default is **ON** (adds TTM lease adj to NTM EBITDA for US GAAP reporters). User can disable with "no lease adjust" or "straight consensus NTM".
5. **Groups** — if the user specifies categories/sectors, pass them as `--groups` arguments.
6. **Date** — defaults to today. User can specify with "as of [date]".

## Two modes explained

### Lease Adjusted (default: `--mode lease-adjusted`)

Matches the "Lease Adjusted" tab in the firm's comp set template:

| Column | Mnemonic | Period | Notes |
|---|---|---|---|
| Company Name | SP_COMPANY_NAME | — | |
| Market Cap | SP_MARKETCAP | as-of date | Includes all equity |
| Total Debt | IQ_TOTAL_DEBT | FQ0 | **Includes** operating leases |
| Oper. Leases | IQ_TOTAL_OPER_LEASES | FQ0 | For reference |
| Cash | IQ_CASH_ST_INVEST | FQ0 | |
| Net Debt | IQ_NET_DEBT | FQ0 | Pulled from CIQ (debt incl leases - cash) |
| TEV | IQ_TEV | as-of date | **Includes** operating leases |
| LTM Revenue | IQ_TOTAL_REV | LTM | |
| NTM Revenue | SP_REV_EST | NTM | Consensus estimate |
| LTM EBITDA | IQ_EBITDA_EQ_INC | LTM | |
| Lease Adj | IQ_LEASE_ADJUSTMENT_EBITDA | LTM | TTM lease adjustment amount |
| NTM EBITDA | SP_EBITDA_EST | NTM | + lease adj if US GAAP (when toggle ON) |
| GAAP | IQ_GAAP_BS | — | Via CIQ(); returns "IFRS" or "US GAAP" |
| P/BV | IQ_PBV_X | LTM | |

**NTM EBITDA lease adjustment logic:**
- Check each company's GAAP standard via `CIQ(ticker, "IQ_GAAP_BS")`
- If toggle ON **and** company is NOT IFRS → add TTM `IQ_LEASE_ADJUSTMENT_EBITDA` to NTM EBITDA
- Rationale: IFRS 16 capitalizes leases (expense below EBITDA line as D&A + interest on ROU assets). US GAAP operating lease expense flows through EBITDA. Adding TTM lease adj to US GAAP NTM EBITDA normalizes cross-border comparisons.
- Mark adjusted companies with `*adj` in output

### Excluding Leases (`--mode excluding-leases`)

Matches the "Excluding Leases" tab in the firm's comp set template:

| Column | Mnemonic | Period | Notes |
|---|---|---|---|
| Company Name | SP_COMPANY_NAME | — | |
| Market Cap | SP_MARKETCAP | as-of date | |
| Total Debt | IQ_TOTAL_DEBT_EXCL_OPER_LEASES | FQ0 | **Excludes** operating leases |
| Oper. Leases | IQ_TOTAL_OPER_LEASES | FQ0 | For reference only |
| Cash | IQ_CASH_ST_INVEST | FQ0 | |
| Net Debt | *(calculated)* | — | = Total Debt excl leases − Cash |
| TEV | IQ_TEV_EXCL_OPER_LEASES | as-of date | **Excludes** operating leases |
| LTM Revenue | IQ_TOTAL_REV | LTM | |
| NTM Revenue | SP_REV_EST | NTM | |
| LTM EBITDA | IQ_EBITDA_EQ_INC_EXCL_OPER_LEASE_ADJ | LTM | **Excludes** lease adjustment |
| NTM EBITDA | SP_EBITDA_EST | NTM | Straight consensus, no adjustment |
| P/BV | IQ_PBV_X | LTM | |

No GAAP check needed — leases excluded from everything uniformly.

## Derived columns (calculated in Python, not from CIQ)

- **EBITDA Margin %** = LTM EBITDA / LTM Revenue
- **EV/Revenue** = TEV / LTM Revenue
- **EV/LTM EBITDA** = TEV / LTM EBITDA
- **EV/NTM EBITDA** = TEV / (lease-adjusted) NTM EBITDA
- **ND/EBITDA** = Net Debt / LTM EBITDA

## Identifier resolution (built-in)

The script automatically resolves ambiguous identifiers in a **single Excel refresh pass** — zero extra cost. For each company, the workbook contains:

1. **Primary row:** SPG formulas using the original identifier
2. **CIQRANGEA lookup:** `=CIQRANGEA(identifier, "IQ_COMPANY_ID_QUICK_MATCH", 1, 1)` — resolves tickers, company names, CUSIPs, ISINs to an IQ ID
3. **Fallback row:** SPG formulas referencing the CIQRANGEA-resolved IQ ID via cell reference

SPG accepts IQ IDs (e.g., `IQ355859`) as identifiers, so if the primary lookup fails but CIQRANGEA resolves, the fallback row provides full data automatically.

After refresh, the script picks the best result per company:
- **Primary resolved** → uses primary data (original ticker worked)
- **Primary failed, fallback resolved** → uses fallback data (CIQRANGEA IQ ID worked). Reports: "Precision Drilling: resolved via CIQRANGEA -> IQ355859 (data OK)"
- **Both failed** → company not recognized. Reports the failure for user to investigate.

### When you (Claude) are unsure about a ticker

**Before building the command**, apply these rules:

1. **Always use `EXCHANGE:TICKER` format** for non-US companies. Canadian tickers MUST have `TSX:` or `TSXV:` prefix.
2. **Baker Hughes** is `NASDAQGS:BKR` (not NYSE).
3. **NASDAQ variants matter**: `NASDAQGS` (Global Select, large-cap), `NasdaqGM` (Global Market, mid-cap), `NasdaqCM` (Capital Market, small-cap). When unsure, try `NASDAQGS` first.
4. **If the user gives company names instead of tickers** (e.g., "Halliburton", "Precision Drilling"), you can either convert to `EXCHANGE:TICKER` format or pass the name as-is. The CIQRANGEA fallback will auto-resolve company names to IQ IDs and pull full data.
5. **If the user gives CUSIPs or ISINs**, include them as-is — CIQRANGEA will resolve them automatically.
6. **If the script reports FAILED tickers** (not resolved by SPG or CIQRANGEA), ask the user to verify the identifier.

### Known ticker gotchas

| Company | Wrong | Correct |
|---|---|---|
| Baker Hughes | NYSE:BKR | NASDAQGS:BKR |
| Canadian tickers | PD, TCW | TSX:PD, TSX:TCW |
| NASDAQ Global Select | NASDAQ:LBRT | NASDAQGS:LBRT |
| Class shares | TSX:CCL | TSX:CCL.B |
| OTC (US-listed foreign) | — | CMDXF, CNSWF (no prefix needed, but may return N/A for some metrics) |

## Execution steps

1. Parse the user's input for tickers, currency, mode, and grouping preferences
2. **Resolve ticker formats** — add exchange prefixes where needed using the rules above
3. Build the `python C:/Claude/pull_comps.py` command with appropriate flags
4. Run with a 5-minute timeout (Excel COM + refresh takes ~40s)
5. Read the console output — check for `IDENTIFIER RESOLUTION REPORT` section
6. **If any tickers resolved via fallback**: note which ones and the IQ IDs used (data is still complete — no re-run needed)
7. **If any tickers FAILED**: tell the user which ones couldn't be resolved and ask them to verify
7. Present results as a formatted **markdown table** in the chat
8. **Always note the currency** in the table header (e.g., "Millions $CAD" or "Millions $USD")
9. **Always note the mode** (Lease Adjusted vs. Excluding Leases)
10. Save CSV to `comps_output.csv` (or user-specified path)

## Ticker format guide

- **U.S. (NYSE):** `NYSE:HAL`, `NYSE:SLB`, `NYSE:NOV`
- **U.S. (NASDAQ):** `NASDAQGS:BKR`, `NASDAQGS:LBRT`, `NasdaqGM:WLDN`
- **Canada (TSX):** `TSX:PD`, `TSX:TCW`, `TSX:CEU`, `TSX:STN`
- **Canada (TSXV):** `TSXV:AEP`
- **Plain tickers** (US-listed, no exchange needed): `DSGX`, `MANH`, `ROP`
- **OTC:** `CMDXF`, `CNSWF` (may return N/A for some metrics)
- **Class shares:** `TSX:CCL.B`, `TSX:BBD.B`, `TSX:RPI.UN`
- **Australian:** Plain ticker works (e.g., `WTC`)

## Defaults summary

| Setting | Default | Override |
|---|---|---|
| Currency | **CAD** | `--currency USD` |
| Mode | **Lease Adjusted** | `--mode excluding-leases` |
| NTM lease adj | **ON** | `--no-lease-adjust-ntm` |
| Date | **Today** | `--date M/D/YYYY` |
| CSV output | `comps_output.csv` | `--csv-output path.csv` |

## Example output header

When presenting the table, always include a header like:

```
Comparable Companies Analysis (Lease Adjusted)
As at 2/28/2026 (Millions $CAD)
NTM EBITDA lease-adjusted for US GAAP reporters
```

Or for excluding leases:

```
Comparable Companies Analysis (Excluding Leases)
As at 2/28/2026 (Millions $USD)
TEV/Debt exclude operating leases
```
