"""Regression tests for the refresh engine status-token classification.

These are pure-function tests — no Excel/COM required.
"""
from capiq_excel.refresh.engine import (
    _classify_cell_token,
    _check_readiness,
    PENDING_TOKENS,
    ALL_ERROR_TOKENS,
)
from capiq_excel.workbook.wait import _data_in_a2


class TestClassifyCellToken:
    """_classify_cell_token receives an already lowercased+stripped value."""

    def test_hash_refresh_is_pending_not_error(self):
        # Regression: bare "refresh" is a legacy error token AND a substring of
        # the pending marker "#refresh". Pending must be checked first so a
        # still-refreshing cell is not misclassified as a hard error (which
        # previously aborted every config-driven download).
        assert _classify_cell_token("#refresh") == ("pending", "#refresh")

    def test_hash_pend_is_pending(self):
        assert _classify_cell_token("#pend") == ("pending", "#pend")

    def test_bare_refresh_is_still_an_error(self):
        # Legacy "stuck refresh" detection is preserved.
        status, tok = _classify_cell_token("refresh")
        assert status == "error"
        assert tok == "refresh"

    def test_ciqinactive_is_error(self):
        assert _classify_cell_token("ciqinactive")[0] == "error"

    def test_hard_error_tokens(self):
        assert _classify_cell_token("#invalid company id")[0] == "error"
        assert _classify_cell_token("#error")[0] == "error"
        assert _classify_cell_token("#name?")[0] == "error"

    def test_clean_values_unclassified(self):
        assert _classify_cell_token("123.45") is None
        assert _classify_cell_token("descartes systems group") is None

    def test_resolved_zero_is_not_a_token(self):
        # A resolved 0 must not look pending/error to the token classifier.
        assert _classify_cell_token("0") is None

    def test_token_sets_are_disjoint_enough(self):
        # The pending "#refresh"/"#pend" markers must not also live in the
        # error set, otherwise ordering wouldn't save us.
        assert "#refresh" in PENDING_TOKENS
        assert "#refresh" not in ALL_ERROR_TOKENS
        assert "#pend" in PENDING_TOKENS


class _Count:
    def __init__(self, count):
        self.Count = count


class _UsedRange:
    def __init__(self, values, formulas):
        self.Row = 1
        self.Column = 1
        self.Rows = _Count(len(values))
        self.Columns = _Count(len(values[0]))
        self.Value = tuple(tuple(row) for row in values)
        self.Formula = tuple(tuple(row) for row in formulas)


class _Cell:
    def __init__(self, value):
        self.Value = value


class _Sheet:
    def __init__(self, values, formulas):
        self._values = values
        self.UsedRange = _UsedRange(values, formulas)

    def Cells(self, row, col):
        try:
            return _Cell(self._values[row - 1][col - 1])
        except IndexError:
            return _Cell(None)


class _Excel:
    CalculationState = 0

    def __init__(self, values, formulas):
        self.ActiveSheet = _Sheet(values, formulas)


def test_readiness_checks_every_lookup_row():
    formulas = [
        ["Input", "Formula", "Result"],
        ["AAA", '=CIQRANGEA("AAA","ID",1,1)', None],
        ["BBB", '=CIQRANGEA("BBB","ID",1,1)', None],
    ]
    values = [
        ["Input", "Formula", "Result"],
        ["AAA", "CIQRANGEA", "IQ1"],
        ["BBB", "CIQRANGEA", None],
    ]
    assert _check_readiness(_Excel(values, formulas))[0] == "pending"

    values[2][2] = "IQ2"
    assert _check_readiness(_Excel(values, formulas)) == ("ready", None)


def test_readiness_can_tolerate_terminal_lookup_errors():
    formulas = [
        ["Input", "Formula", "Result"],
        ["BAD", '=CIQRANGEA("BAD","ID",1,1)', None],
    ]
    values = [
        ["Input", "Formula", "Result"],
        ["BAD", "CIQRANGEA", "#INVALID COMPANY ID"],
    ]
    excel = _Excel(values, formulas)
    assert _check_readiness(excel)[0] == "error"
    assert _check_readiness(excel, allow_cell_errors=True) == ("ready", None)


def test_legacy_a2_uses_excel_row_column_order():
    class Excel:
        class Sheet:
            @staticmethod
            def Cells(row, col):
                return _Cell("done" if (row, col) == (2, 1) else None)
        ActiveSheet = Sheet()

    assert _data_in_a2(Excel()) is True
