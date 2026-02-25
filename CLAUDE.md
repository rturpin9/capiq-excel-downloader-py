# capiq-excel-downloader-py

## Project Overview

Python tool that drives Microsoft Excel via COM automation to download data from S&P Capital IQ using the CIQ Excel plugin. Currently being modernized to support **Capital IQ Pro** (SPG/SNL formula families) alongside the legacy CIQ plugin.

**Modernization plan:** `C:\Users\rturpin\Downloads\capiq-pro-compatibility-modernization-plan.md`

## Architecture

### Current Package Structure

```
capiq_excel/
  __init__.py          # Public API: download_data, download_data_for_capiq_ids
  main.py              # Orchestrator: create XLSX -> populate via Excel COM -> combine to CSV
  addin.py             # Loads legacy "S&P Capital IQ Excel Plug-in" (hardcoded name)
  exceptions.py        # WorkbookClosedException, CapitalIQInactiveException
  ids.py               # ID resolution: arbitrary IDs -> CIQ IDs via CIQRANGEA
  fileops.py           # Failed-file management (move to failed folder)
  combine.py           # Combine per-company XLSX results into single CSV
  workbook/
    commands.py        # Formula generators: CIQRANGE, CIQRANGEA (legacy CIQ only)
    create.py          # Create XLSX workbooks pre-filled with CIQ formulas
    wait.py            # Poll Excel for CIQ result readiness (cell A2 heuristic)
    populate/
      main.py          # Open XLSX in Excel, wait for CIQ refresh, save results
      extract.py       # Extract data from CIQ formula cells (financial + market)
      replace.py       # Write aligned DataFrame back to worksheet
  downloader/
    tools.py           # Batch file processing with retry/timeout/Excel restart
    timeout.py         # ThreadPool-based timeout wrapper
  tools/
    dates.py           # Date helpers for CIQ formula period calculation
    ext_pandas.py      # CSV append utilities, date parsing, DataFrame helpers
```

### External Dependencies (COM/Windows)

- `exceldriver` - Excel COM automation (start/attach/restart Excel, add-in loading, workbook creation, column utilities)
- `processfiles` - File tracking for batch processing (`FileProcessTracker`)
- `pypiwin32` / `pythoncom` / `win32com` - Windows COM interop
- `openpyxl` - XLSX creation (offline, before Excel opens them)
- `pandas` - Data manipulation and CSV I/O
- `xlrd` - Legacy Excel reading

### Data Flow

1. **Create phase** (`workbook/create.py`): Generate XLSX files with CIQ formulas in cells (one per company)
2. **Populate phase** (`downloader/tools.py` -> `workbook/populate/main.py`): Open each XLSX in Excel via COM, CIQ plugin evaluates formulas, wait for results, save
3. **Combine phase** (`combine.py`): Read all populated XLSX files, merge into single CSV output

### Key Patterns

- COM operations require `pythoncom.CoInitialize()` per thread
- Excel is restarted every 500 workbooks to work around memory leaks
- Failed files are moved to a `failed/` subfolder for retry
- Market data items produce unaligned date axes (each cell has its own date in the formula); extraction realigns them via `extract.py`

## Modernization Target (Phased)

### Phase 0: Branching + Feature Flags
- Feature flags: `formula_dialect` (auto|ciq|spg), `addin_mode` (auto|legacy|pro), `refresh_scope`
- Existing behavior preserved under `ciq/legacy` defaults

### Phase 1: Add-In Discovery + Compatibility Layer
- New: `capiq_excel/runtime/addin_detection.py` - detect legacy vs Pro add-ins
- Replace hardcoded `load_capiq()` with `load_best_addin()` with fallback
- Startup diagnostics (Excel version, add-in profile, selected mode)

### Phase 2: Formula Dialect Abstraction
- New: `capiq_excel/formulas/{base,ciq_builder,spg_builder}.py`
- `QuerySpec` canonical input -> dialect-specific formula output
- CIQ builder: `CIQ()`, `CIQRANGE()`, `CIQRANGEA()`
- SPG builder: `SPG()`, `SPGRangeV()`, `SPGTable()`, `SNL*()`

### Phase 3: Refresh Engine
- New: `capiq_excel/refresh/engine.py`
- Replace cell-A2 heuristic with strategy-based completion checks
- Batch sequencing, throttling, dependency-aware refresh

### Phase 4: Config + Settings
- New: `capiq_excel/config.py` - dataclass/pydantic config model
- Registry/env integration for Windows settings
- CLI commands: `capiq status`, `capiq detect-addins`, `capiq download`, `capiq doctor`

### Phase 5: Reliability + Observability
- Structured logging with run/file IDs
- Error taxonomy: `AddinNotFoundError`, `DialectUnsupportedError`, `AuthSessionError`, `RefreshTimeoutError`, `FormulaValidationError`
- Diagnostics bundle export

### Phase 6: Packaging
- Migrate to `pyproject.toml`
- Pin Windows COM dependencies
- CI: unit tests in CI, integration tests gated on Windows+Excel

## Conventions

- Python 3.10+ target (modernization)
- Type hints on all new code
- Dataclasses for config and data structures (pydantic optional)
- `pytest` for testing
- Formula output must be snapshot-testable (string comparison)
- All COM interaction behind abstraction layers
- Feature flags control dialect/mode selection; defaults preserve legacy behavior
- New modules go under clear subdirectories: `runtime/`, `formulas/`, `refresh/`

## Commands

```bash
# Run tests (requires Windows + Excel + CIQ plugin for integration tests)
pytest test/

# Install in development mode
pip install -e .
```

## Important Notes

- This is a Windows-only tool (COM automation with Excel)
- The legacy CIQ add-in name is exactly: `S&P Capital IQ Excel Plug-in`
- CIQ formulas: `CIQ()`, `CIQRANGE()`, `CIQRANGEA()` - use relative periods like `IQ_FQ - 80`
- Market data formulas use absolute date strings instead of relative periods
- The `exceldriver` package provides: `load_addin()`, `_start_excel_with_addins_and_attach()`, `_restart_excel_with_addins_and_attach()`, `get_workbook_and_worksheet()`, `excel_cols()`
- `processfiles.FileProcessTracker` manages which files have been processed in a batch
