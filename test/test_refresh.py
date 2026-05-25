"""Regression tests for the refresh engine status-token classification.

These are pure-function tests — no Excel/COM required.
"""
from capiq_excel.refresh.engine import (
    _classify_cell_token,
    PENDING_TOKENS,
    ALL_ERROR_TOKENS,
)


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
