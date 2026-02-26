# capiq-excel-downloader-py

## Project Overview

Python tool that drives Microsoft Excel via COM automation to download data from S&P Capital IQ using the CIQ Excel plugin. **Modernized** to support **Capital IQ Pro** (SPG/SNL formula families) alongside the legacy CIQ plugin.

**Modernization plan:** `C:\Users\rturpin\Downloads\capiq-pro-compatibility-modernization-plan.md`
**Pro plugin docs:** `C:\Claude\Pro_plugin_documentation\` (Tech Guide + User Guide PDFs)

## Architecture

### Package Structure

```
capiq_excel/
  __init__.py          # Public API: download_data, download_data_for_capiq_ids,
                       #   CapiqConfig, FormulaDialect, AddinMode, get_builder, QuerySpec
  main.py              # Orchestrator: create XLSX -> populate via Excel COM -> combine to CSV
                       #   Accepts optional CapiqConfig for dialect/mode selection
  addin.py             # load_capiq(legacy) + load_capiq_addin(detect runtime, load best)
  config.py            # CapiqConfig, FormulaDialect/AddinMode/RefreshScope enums,
                       #   FormulaOptions with to_spg_options_string()
  cli.py               # CLI entry: capiq status|detect-addins|download|doctor
  exceptions.py        # Exception classes (legacy + new taxonomy)
  ids.py               # ID resolution: arbitrary IDs -> CIQ IDs (builder-aware)
  fileops.py           # Failed-file management (move to failed folder)
  combine.py           # Combine per-company XLSX results into single CSV
  runtime/
    addin_detection.py # RuntimeProfile, detect_runtime(), load_best_addin(),
                       #   registry checks (SNL Office keys, CIQ compat toggle)
  formulas/
    __init__.py        # get_builder(dialect) factory
    base.py            # QuerySpec (canonical query), DialectBuilder ABC
    ciq_builder.py     # CiqBuilder: CIQ(), CIQRANGE() (ID lookup via CIQ)
    spg_builder.py     # SpgBuilder: SPG(), SPGRangeV(), SPGTable()
    snl_builder.py     # SnlBuilder: SNLData(), SNLMarkets(), SNLTable(),
                       #   SNLQuery(), SNLDefinition(), SNLConvert()
  refresh/
    engine.py          # refresh_and_wait() with VBA Application.Run for Pro,
                       #   expanded error/pending token detection
  workbook/
    commands.py        # Builder-aware factories (make_*_command) + legacy CIQ functions
    create.py          # Create XLSX workbooks (accepts optional DialectBuilder)
    wait.py            # Legacy cell-A2 polling (used when no config provided)
    populate/
      main.py          # Open XLSX in Excel, refresh (legacy or Pro), save results
      extract.py       # Extract data from formula cells (financial + market)
      replace.py       # Write aligned DataFrame back to worksheet
  downloader/
    tools.py           # Batch file processing with retry/timeout/Excel restart
                       #   (threads config through for Pro-aware detection)
    timeout.py         # ThreadPool-based timeout wrapper
  tools/
    dates.py           # Date helpers (pandas freq compat: Q->QE, Y->YE)
    ext_pandas.py      # CSV append utilities, date parsing, DataFrame helpers
```

### Data Flow

1. **Config** (`CapiqConfig.from_env()` or explicit) -> resolve dialect
2. **Build formulas** (`get_builder(dialect)` -> `make_*_command(builder)`)
3. **Create XLSX** (`create.py` fills cells with dialect-appropriate formulas)
4. **Populate** (Open in Excel via COM -> plugin evaluates -> refresh engine waits)
5. **Combine** (Extract results into CSV)

### Formula Dialect Differences

| Feature | CIQ (Legacy) | SPG (Pro) | SNL (Pro) |
|---|---|---|---|
| Single value | `=CIQ(id, metric)` | `=SPG(id, metric, period, opts)` | `=SNLData(dataset, id, field, key)` |
| Range (formulas) | `=CIQRANGE(id, metric, period...)` | `=SPGRangeV(id, metric, begin, end, opts)` | `=SNLMarkets(id, field, key, start, end)` |
| Range (values) | `=CIQRANGEV(id, metric, period...)` | | |
| Range (across) | `=CIQRANGEA(id, metric, ...)` | | |
| Table | _(none)_ | `=SPGTable(ids, metrics, periods, opts)` | `=SNLTable(dataset, ids, fields, keys, opts)` |
| ID lookup | `=CIQ(search, "IQ_COMPANY_ID")` | `=SPG(search, field)` | `=SNLData(1, search, field)` |
| Period syntax | `IQ_FQ-80` (relative) | `FQ-80`, `FY2020`, `FQ12020` | `2013Q2`, `MRQ`, `[MRQ-1]` |
| Options | positional args | `"Curr=USD,Mag=Millions"` | `"Curr=USD,Mag=Millions"` |

### CIQ CIQRANGE Parameter Patterns (CRITICAL)

**The parameter pattern differs by data category. Using the wrong one returns `(Invalid Time Period)` or `(Invalid Period Type)`.** Substitute `CIQRANGEV` for efficiency or `CIQRANGEA` for horizontal expansion.

```
# Market data (pricing, TEV, market cap): date range
=CIQRANGE("DSGX", "IQ_CLOSEPRICE", "2/26/2025", "2/26/2026", , , , , , "Price")
=CIQRANGE("DSGX", "IQ_TEV", "2/26/2025", "2/26/2026", , , , , , "TEV")

# Financial data (income stmt, balance sheet, cash flow): period offset
=CIQRANGE("DSGX", "IQ_EBITDA", IQ_FQ-16, , , , , , , "EBITDA")
=CIQRANGE("DSGX", "IQ_TOTAL_REV", IQ_FY-5, , , , , , , "Revenue")

# Financial dates (for date column alongside financial data):
=CIQRANGE("DSGX", "IQ_PERIODDATE_BS", IQ_FQ-16, , , , , , , "Date")

# Trading multiples: period TYPE + date range (relative or absolute)
=CIQRANGEV("DSGX", "IQ_TEV_EBITDA", IQ_LTM, "-1Y", "2/26/2026", , , , "TEV/EBITDA")
=CIQRANGEV("DSGX", "IQ_TEV_TOTAL_REV", IQ_LTM, "2/26/2023", "2/26/2026", , , , "TEV/Revenue")
```

**Period types for multiples:** `IQ_LTM`, `IQ_FY`, `IQ_CY`, `IQ_FQ`, `IQ_CQ`, `IQ_FH`, `IQ_CH`
**Date formats:** Absolute `"M/D/YYYY"` or relative `"-1Y"`, `"-6M"`, `"-90D"`

### Workbook Layout for CIQRANGE

- **Formula goes in Row 1** of each column, data expands downward from Row 2
- The `Label` parameter (last arg) becomes the column header text
- Market data returns ~252 daily points/year; multiples ~198 daily points/year; financial data = N quarters/years

### Pro Refresh Commands (VBA via Application.Run)

- Selected cells: `SNLxlAddin.xla!RefreshActiveCells`
- Entire sheet: `SNLxlAddin.xla!RefreshSheet`
- All sheets: `SNLxlAddin.xla!RefreshWorkbook`

### In-Cell Status Tokens

- **Pending:** `#REFRESH`, `#PEND`
- **Hard errors:** `#ERROR`, `#INVALID COMPANY ID`, `#INVALID METRIC NAME`, `#INVALID FUNCTION PARAMETER`, `#OUTSIDE SUBSCRIPTION`, `KEYERROR`, `DEFUNCT`, `InvalidCurrency`, `InvalidMagnitude`, `InvalidConvMethod`, `#NAME`
- **Legacy errors:** `ciqinactive`, `refresh`

### Registry Keys (Pro)

- Settings: `HKCU\Software\SNL Financial\SNL Office`
- Load: `HKCU\Software\Microsoft\Office\Excel\Addins\SNL.Clients.Office.Excel.ExcelAddIn\LoadBehavior`
- CIQ compat toggle: under SNL Office key (v21.08+)
- Logs: `%LocalAppData%\SPGMI`

## External Dependencies

- `exceldriver` - Excel COM automation
- `processfiles` - Batch file tracking
- `pypiwin32` / `pythoncom` / `win32com` - Windows COM interop
- `openpyxl` - XLSX creation
- `pandas` - Data manipulation
- `xlrd` - Legacy Excel reading

## Conventions

- Python 3.10+ target
- Type hints on all new code
- Dataclasses for config and data structures
- `pytest` for testing; formula output is snapshot-testable (string comparison)
- All COM interaction behind abstraction layers
- Feature flags control dialect/mode; defaults preserve legacy behavior
- Builder-aware command factories (`make_*_command(builder)`) alongside legacy functions

## Commands

```bash
# Run unit tests (no Excel/COM required)
python -m pytest test/test_config.py test/test_formulas.py test/test_smoke.py -v

# Run all tests (requires Windows + Excel + CIQ plugin for integration)
python -m pytest test/ -v

# CLI
capiq status          # Show config from env vars
capiq detect-addins   # Start Excel, detect installed add-ins
capiq download --ids MSFT AAPL --financial-items IQ_TOTAL_REV --dialect spg
capiq doctor          # Check environment health
```

## Key Notes

- Windows-only tool (COM automation with Excel)
- Legacy CIQ add-in name: `S&P Capital IQ Excel Plug-in`
- Pro add-in ProgID: `SNL.Clients.Office.Excel.ExcelAddIn` (display name: `S&P Cap IQ Pro Excel Add-In`)
- Pro UDF/refresh XLA: `SNLxlAddin.xla` (installed at `%ProgramFiles(x86)%\SNL Financial\SNLxl`)
- CIQ compat in Pro: toggleable at install and post-install (Settings UI)
- When CIQ compat is enabled, Pro handles CIQ formula refresh — separate CIQ add-in not needed
- `tools/dates.py` maps `Q->QE`, `Y->YE` for pandas 3.x compatibility
- All public functions accept optional `config: CapiqConfig` — omitting it preserves legacy behavior
- `test_download.py` is an integration test requiring live Excel + CIQ plugin
- **CRITICAL**: Excel must be launched via subprocess (not `Dispatch()`) for UDFs to register — use `exceldriver._start_excel_with_addins_and_attach()`
- In Pro CIQ compat mode: `CIQ()`, `CIQRANGE()`, and `CIQRANGEV()` work; `CIQRANGEA()` does NOT (returns function name as string)
- CIQ builder uses `=CIQ(search,"IQ_COMPANY_ID")` for ID lookup (not CIQRANGEA)
- CIQ compat registry: `DisableCIQUDF = 0` means CIQ is ENABLED (inverted logic)
- SPG formulas require different metric names/ID formats than CIQ — not yet fully mapped
- COM error `-2146826259` = `#NAME?` (UDF not registered); `-2146826273` = `#VALUE!`
- **Pre-calculated multiples** (IQ_TEV_EBITDA, IQ_TEV_TOTAL_REV, etc.) require `IQ_LTM` period type + date range — financial period syntax (`IQ_FQ-N`) returns `(Invalid Time Period)`
- `CIQ()` with period param works for current period only (`IQ_FQ-0`, `IQ_LTM`); historical periods (`IQ_FQ-1` etc.) return `(Invalid Period Type)` for multiples
- `excel_lifecycle.py` provides `launch_excel_isolated()` / `close_session()` for safe non-destructive COM sessions
- Working example scripts: `comp_table.py` (single-value CIQ), `ev_multiples_chart.py` (CIQRANGEV time series)
