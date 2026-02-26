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

        status, detail = _check_readiness(excel)
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


def _check_readiness(excel) -> tuple[str, Optional[str]]:
    """
    Check whether the active sheet has finished evaluating.

    Scans the first two rows across the first several columns for known
    status/error tokens from CIQ and SPG/SNL. This covers both data worksheets
    (formulas in column A) and ID lookup worksheets (formulas in columns B/D).

    Also checks if formula cells have resolved — if a cell has a formula but
    its value is 0/empty/None, the formula likely hasn't been evaluated yet.

    Returns:
        ("ready", None)          - data is present, no error/pending tokens
        ("pending", token)       - still evaluating
        ("error", error_detail)  - plugin reported an error
    """
    try:
        ws = excel.ActiveSheet
    except AttributeError:
        return "error", "workbook_closed"

    found_any_data = False
    has_unresolved_formula = False

    # Scan the first 2 rows x first 8 columns
    for row in (1, 2):
        for col in range(1, 9):
            val = _get_cell_str(excel, col=col, row=row)

            # Check if this cell has a formula that hasn't resolved
            try:
                cell = ws.Cells(row, col)
                if cell.HasFormula:
                    cell_val = cell.Value
                    # COM error codes (e.g. -2146826259 = #NAME?) mean
                    # the UDF hasn't registered yet — still pending
                    if isinstance(cell_val, int) and cell_val < -2000000000:
                        has_unresolved_formula = True
                    # Formula cell with 0, empty, or None = likely unevaluated
                    elif cell_val is None or cell_val == 0 or cell_val == "":
                        has_unresolved_formula = True
            except Exception:
                logger.debug("Could not read formula status from cell(%d, %d)", row, col, exc_info=True)

            if val is None:
                continue

            val_lower = val.lower().strip()
            if not val_lower:
                continue

            found_any_data = True

            # Check for hard error states
            for tok in ALL_ERROR_TOKENS:
                if tok in val_lower:
                    return "error", tok

            # Check for pending states
            for tok in PENDING_TOKENS:
                if tok in val_lower:
                    return "pending", tok

    # If we found no data at all, still pending
    if not found_any_data:
        return "pending", "no_data"

    # If any formula cell in row 2 hasn't resolved, still pending
    if has_unresolved_formula:
        return "pending", "unresolved_formula"

    return "ready", None


def _get_cell_str(excel, col: int, row: int) -> Optional[str]:
    """Safely read a cell value as a string."""
    try:
        val = excel.ActiveSheet.Cells(row, col).Value
        return str(val) if val is not None else None
    except Exception:
        return None
