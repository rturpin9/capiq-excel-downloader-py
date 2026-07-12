"""
Refresh engine for Capital IQ workbook evaluation.

Replaces the legacy cell-A2 polling heuristic with strategy-based
completion checks, configurable timeouts, and mode-aware refresh behavior.

Supports two refresh strategies:
  - Legacy CIQ: open workbook triggers auto-evaluation, poll cell A2
  - Pro SPG/SNL: explicit VBA Application.Run commands via the Pro add-in

VBA refresh commands (from CIQ Pro documentation):
  - RefreshActiveCells: Application.Run "SNLxlAddin.xla!RefreshActiveCells"
  - RefreshSheet:       Application.Run "SNLxlAddin.xla!RefreshSheet"
  - RefreshWorkbook:    Application.Run "SNLxlAddin.xla!RefreshWorkbook"
"""
from __future__ import annotations

import logging
import time
from typing import Optional

from capiq_excel.config import AddinMode, CapiqConfig, RefreshScope
from capiq_excel.exceptions import RefreshTimeoutError, CapitalIQInactiveException

logger = logging.getLogger(__name__)

# --- In-cell status tokens from the User Guide (p.62-63) ---

# Tokens indicating the plugin is still evaluating (transient/retryable)
PENDING_TOKENS = frozenset({
    "#refresh",     # currently refreshing from servers
    "#pend",        # needs refresh / export in progress
    "fetching",
    "loading",
    "pending",
    "calculating",
})

# Tokens indicating a hard error (not retryable without fixing the cause)
HARD_ERROR_TOKENS = frozenset({
    "#error",                       # general refresh problem
    "#invalid company id",          # bad identifier
    "#invalid metric name",         # bad field/metric
    "#invalid function parameter",  # bad parameter
    "#outside subscription",        # not subscribed
    "keyerror",                     # missing secondary/tertiary key
    "defunct",                      # deprecated field
    "invalidcurrency",              # bad currency in options string
    "invalidmagnitude",             # bad magnitude in options string
    "invalidconvmethod",            # bad conversion method
    "invalidparameter",             # generic invalid parameter
    "#name",                        # add-in not loaded (Excel error)
})

# Tokens from legacy CIQ that indicate plugin-level failure
LEGACY_ERROR_TOKENS = frozenset({
    "ciqinactive",
    "refresh",      # bare "refresh" in legacy mode = stuck
})

ALL_ERROR_TOKENS = HARD_ERROR_TOKENS | LEGACY_ERROR_TOKENS

_VERTICAL_RANGE_FORMULAS = (
    "=ciqrange(",
    "=ciqrangev(",
    "=spgrangev(",
    "=snlmarkets(",
)
_HORIZONTAL_RANGE_FORMULAS = ("=ciqrangea(",)
_COM_NAME_ERROR = -2146826259

# Pro VBA macro names for Application.Run (via SNLxlAddin.xla)
_VBA_REFRESH_SELECTED = "SNLxlAddin.xla!RefreshActiveCells"
_VBA_REFRESH_SHEET = "SNLxlAddin.xla!RefreshSheet"
_VBA_REFRESH_ALL = "SNLxlAddin.xla!RefreshWorkbook"

# Default delay after workbook open to let the add-in initialize (seconds).
# The Pro User Guide VBA example uses 10s; we default to 5 but make it configurable.
DEFAULT_ADDIN_INIT_DELAY = 5.0


def refresh_and_wait(
    excel,
    config: CapiqConfig,
    *,
    addin_mode: AddinMode = AddinMode.AUTO,
    poll_interval: float = 1.0,
    init_delay: float = DEFAULT_ADDIN_INIT_DELAY,
    allow_cell_errors: bool = False,
) -> bool:
    """
    Trigger a refresh and wait for completion.

    Args:
        excel: COM Excel.Application instance.
        config: Runtime configuration.
        addin_mode: Which add-in is active (determines refresh strategy).
        poll_interval: Seconds between readiness checks.
        init_delay: Seconds to wait after triggering refresh before first poll
                    (gives the add-in time to start evaluating).

    Returns:
        True if refresh completed successfully.

    Raises:
        RefreshTimeoutError: If timeout is exceeded.
        CapitalIQInactiveException: If plugin reports an error state.
    """
    timeout = config.retry.timeout_seconds
    scope = config.refresh_scope

    logger.debug("Starting refresh (scope=%s, mode=%s, timeout=%ds)", scope.value, addin_mode.value, timeout)

    # Trigger the appropriate refresh
    _trigger_refresh(excel, scope, addin_mode)

    # Wait for add-in to begin evaluation
    if init_delay > 0:
        time.sleep(init_delay)

    # Poll for completion
    start = time.monotonic()
    while True:
        elapsed = time.monotonic() - start
        if elapsed > timeout:
            raise RefreshTimeoutError(
                f"Refresh did not complete within {timeout}s "
                f"(scope={scope.value}, mode={addin_mode.value})"
            )

        status, detail = _check_readiness(
            excel,
            allow_cell_errors=allow_cell_errors,
        )
        if status == "ready":
            logger.debug("Refresh completed in %.1fs", elapsed)
            return True
        elif status == "error":
            raise CapitalIQInactiveException(
                f"Capital IQ plugin reported an error during refresh: {detail}"
            )
        # status == "pending" -> keep waiting
        time.sleep(poll_interval)


def _trigger_refresh(excel, scope: RefreshScope, addin_mode: AddinMode) -> None:
    """Trigger refresh using the appropriate mechanism.

    Legacy CIQ: opening a workbook with CIQ formulas auto-triggers evaluation.
    Pro SPG/SNL: we invoke VBA Application.Run macros for explicit refresh.
    """
    if addin_mode in (AddinMode.PRO, AddinMode.AUTO):
        _trigger_pro_refresh(excel, scope)
    # Legacy mode: formulas auto-evaluate on workbook open (no-op)


def _trigger_pro_refresh(excel, scope: RefreshScope) -> None:
    """Invoke the Pro add-in VBA refresh command."""
    macro = {
        RefreshScope.SELECTION: _VBA_REFRESH_SELECTED,
        RefreshScope.WORKSHEET: _VBA_REFRESH_SHEET,
        RefreshScope.WORKBOOK: _VBA_REFRESH_ALL,
    }.get(scope, _VBA_REFRESH_SHEET)

    try:
        excel.Run(macro)
        logger.debug("Triggered Pro refresh via %s", macro)
    except Exception:
        # If VBA macro fails (e.g. add-in not ready), log and fall through
        # to polling — the formulas may still auto-evaluate
        logger.warning("Failed to invoke Pro refresh macro %s, falling back to polling", macro, exc_info=True)


def _classify_cell_token(val_lower: str) -> Optional[tuple[str, str]]:
    """Classify a normalized (lowercased, stripped) cell value by status token.

    Pending tokens are checked BEFORE error tokens.  This is required because
    the legacy bare-"refresh" error token is a substring of the pending marker
    "#refresh"; checking errors first would misclassify a cell that is merely
    still refreshing as a hard error and abort the whole refresh.

    Returns ("pending", token) / ("error", token), or None if no token matches.
    """
    for tok in PENDING_TOKENS:
        if tok in val_lower:
            return "pending", tok
    for tok in ALL_ERROR_TOKENS:
        if tok in val_lower:
            return "error", tok
    return None


def _check_readiness(
    excel,
    *,
    allow_cell_errors: bool = False,
) -> tuple[str, Optional[str]]:
    """
    Check whether the active sheet has finished evaluating.

    Batch-scans every formula in the active sheet's used range.  This is
    important for ID workbooks, where each input row has independent formulas;
    checking only the first result can close the workbook while later rows are
    still pending.

    Also checks if formula cells have resolved — a formula cell whose value is
    still None (or a COM #NAME? error code) likely hasn't been evaluated yet.
    A resolved 0 / "" is treated as valid data, not as "unresolved".

    Returns:
        ("ready", None)          - data is present, no error/pending tokens
        ("pending", token)       - still evaluating
        ("error", error_detail)  - plugin reported an error
    """
    try:
        ws = excel.ActiveSheet
    except AttributeError:
        return "error", "workbook_closed"

    try:
        used = ws.UsedRange
        first_row = int(used.Row)
        first_col = int(used.Column)
        row_count = int(used.Rows.Count)
        col_count = int(used.Columns.Count)
        values = _as_matrix(used.Value, row_count, col_count)
        formulas = _as_matrix(used.Formula, row_count, col_count)
    except Exception:
        logger.warning("Could not batch-read active sheet for refresh status", exc_info=True)
        return "pending", "sheet_unreadable"

    found_formula = False
    has_unresolved_formula = False

    def value_at(row: int, col: int):
        local_row = row - first_row
        local_col = col - first_col
        if 0 <= local_row < row_count and 0 <= local_col < col_count:
            return values[local_row][local_col]
        try:
            return ws.Cells(row, col).Value
        except Exception:
            return None

    for row_offset in range(row_count):
        for col_offset in range(col_count):
            formula = formulas[row_offset][col_offset]
            if not isinstance(formula, str) or not formula.startswith("="):
                continue

            found_formula = True
            row = first_row + row_offset
            col = first_col + col_offset
            value = values[row_offset][col_offset]

            status = _classify_formula_value(value)
            if status is not None:
                if status[0] == "error":
                    if not allow_cell_errors:
                        return status
                else:
                    has_unresolved_formula = True

            formula_lower = formula.lower().replace(" ", "")
            spill_value = None
            if formula_lower.startswith(_HORIZONTAL_RANGE_FORMULAS):
                spill_value = value_at(row, col + 1)
            elif formula_lower.startswith(_VERTICAL_RANGE_FORMULAS):
                spill_value = value_at(row + 1, col)

            if spill_value is not None or formula_lower.startswith(
                _HORIZONTAL_RANGE_FORMULAS + _VERTICAL_RANGE_FORMULAS
            ):
                spill_status = _classify_formula_value(spill_value)
                if spill_status is not None:
                    if spill_status[0] == "error":
                        if not allow_cell_errors:
                            return spill_status
                    else:
                        has_unresolved_formula = True

    if not found_formula:
        return "pending", "no_formulas"

    try:
        # xlDone == 0.  Native Excel calculation can finish after cell values
        # first become non-empty, so include it in the completion gate.
        if excel.CalculationState not in (0, None):
            return "pending", "excel_calculating"
    except Exception:
        logger.debug("Could not read Excel.CalculationState", exc_info=True)

    if has_unresolved_formula:
        return "pending", "unresolved_formula"

    return "ready", None


def _classify_formula_value(value) -> Optional[tuple[str, str]]:
    """Classify a formula or spill value without rejecting valid zero/blank strings."""
    if value is None:
        return "pending", "no_value"
    if isinstance(value, int) and value < -2000000000:
        if value == _COM_NAME_ERROR:
            return "pending", "#name"
        return "error", f"excel_error_{value}"
    if isinstance(value, str):
        normalized = value.lower().strip()
        if not normalized:
            return None
        return _classify_cell_token(normalized)
    return None


def _as_matrix(raw, rows: int, cols: int) -> list[list]:
    """Normalize Excel's scalar/tuple Range values to a rectangular matrix."""
    if rows == 1 and cols == 1:
        return [[raw]]
    if not isinstance(raw, (tuple, list)):
        return [[raw for _ in range(cols)] for _ in range(rows)]

    if rows == 1 and len(raw) == cols and not isinstance(raw[0], (tuple, list)):
        return [list(raw)]
    if cols == 1 and len(raw) == rows and not isinstance(raw[0], (tuple, list)):
        return [[item] for item in raw]

    matrix = []
    for row in raw:
        if isinstance(row, (tuple, list)):
            matrix.append(list(row))
        else:
            matrix.append([row])
    return matrix
