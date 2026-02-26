class WorkbookClosedException(Exception):
    pass

class CapitalIQInactiveException(Exception):
    pass


# --- New exception taxonomy (Phase 5) ---

class AddinNotFoundError(Exception):
    """No suitable Capital IQ add-in could be detected or loaded."""
    pass

class DialectUnsupportedError(Exception):
    """Requested formula dialect is not available in the current environment."""
    pass

class AuthSessionError(Exception):
    """Capital IQ authentication or session is invalid/expired."""
    pass

class RefreshTimeoutError(Exception):
    """Workbook refresh did not complete within the allowed timeout."""
    pass

class FormulaValidationError(Exception):
    """Generated formula failed validation checks."""
    pass