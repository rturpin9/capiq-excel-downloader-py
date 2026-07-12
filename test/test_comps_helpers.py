"""Regression tests for comp-table value helpers and group formatting.

Pure-function tests — no Excel/COM required (importing the engine only loads
pythoncom as a module; no COM calls are made here).
"""
import numpy as np

from capiq_excel.engines.comps import (
    safe_div,
    safe_float,
    is_error_value,
    strip_us_exchange_prefix,
    build_formula,
    _spg_options_with_millions,
)
from capiq_excel.cli import _format_comps_markdown, parse_groups


class TestSafeDiv:
    def test_positive(self):
        assert safe_div(10.0, 2.0) == 5.0

    def test_zero_denominator_is_nan(self):
        assert np.isnan(safe_div(10.0, 0.0))

    def test_nan_inputs_are_nan(self):
        assert np.isnan(safe_div(np.nan, 2.0))
        assert np.isnan(safe_div(10.0, np.nan))

    def test_negative_dropped_by_default(self):
        # EV multiples: a negative result is suppressed as not meaningful.
        assert np.isnan(safe_div(-5.0, 2.0))

    def test_negative_kept_when_allowed(self):
        # Regression: Net Debt / EBITDA for a net-cash company (negative net
        # debt) is legitimately negative and must not be blanked.
        assert safe_div(-5.0, 2.0, allow_negative=True) == -2.5

    def test_zero_result_kept_when_allowed(self):
        assert safe_div(0.0, 2.0, allow_negative=True) == 0.0


class TestErrorTokens:
    def test_nm_exact_is_error(self):
        assert is_error_value("NM") is True
        assert np.isnan(safe_float("NM"))

    def test_nm_substring_in_name_is_not_error(self):
        # Regression: "NMI Holdings" / "Lanmark" contain "NM" but are valid
        # company names — must not be flagged as errors.
        assert is_error_value("NMI Holdings") is False
        assert is_error_value("Lanmark Inc") is False

    def test_hard_errors_still_flagged(self):
        assert is_error_value("#ERROR") is True
        assert is_error_value("#INVALID COMPANY ID") is True
        assert is_error_value("(Invalid Time Period)") is True
        assert is_error_value("DEFUNCT") is True

    def test_numeric_values_not_errors(self):
        assert is_error_value(123.4) is False
        assert is_error_value(0) is False
        assert is_error_value(-2146826259) is True
        assert safe_float("1,234.5") == 1234.5


class TestStripUsExchangePrefix:
    def test_us_prefix_stripped(self):
        assert strip_us_exchange_prefix("NYSE:HAL") == "HAL"
        assert strip_us_exchange_prefix("NASDAQGS:BKR") == "BKR"

    def test_non_us_prefix_preserved(self):
        assert strip_us_exchange_prefix("TSX:PD") == "TSX:PD"
        assert strip_us_exchange_prefix("LSE:BP") == "LSE:BP"

    def test_plain_ticker_unchanged(self):
        assert strip_us_exchange_prefix("DSGX") == "DSGX"


class TestMagnitudeOptions:
    def test_spg_options_appends_millions(self):
        assert _spg_options_with_millions("Options: Curr=USD") == "Options: Curr=USD,Mag=Millions"

    def test_spg_options_idempotent(self):
        # Don't double-append if a magnitude is already present.
        assert _spg_options_with_millions("Options: Curr=CAD,Mag=Millions") == \
            "Options: Curr=CAD,Mag=Millions"

    def test_spg_formula_carries_millions(self):
        # SP_ mnemonic -> SPG and must pin Mag=Millions (SPG defaults to thousands).
        f = build_formula("AAPL", "SP_MARKETCAP", "market", "1/1/2026", "Options: Curr=USD")
        assert f.startswith("=SPG(")
        assert "Mag=Millions" in f

    def test_ciq_formula_omits_millions(self):
        # IQ_ mnemonic -> CIQ and must NOT get the longer string, which overflows
        # CIQ's market currency parameter ("(Parameter Length Limit Exceeded)").
        f = build_formula("AAPL", "IQ_TEV", "market", "1/1/2026", "Options: Curr=USD")
        assert f.startswith("=CIQ(")
        assert "Mag=Millions" not in f


class TestGroupsFiltering:
    def _data(self):
        # Stored Ticker mirrors comps.build_workbook: US prefixes stripped,
        # non-US prefixes preserved.
        return {
            "mode": "lease-adjusted", "currency": "USD", "date": "1/1/2026",
            "companies": [
                {"Company": "Halliburton Company", "Ticker": "HAL", "Mkt Cap": 40000.0},
                {"Company": "SLB N.V.", "Ticker": "SLB", "Mkt Cap": 100000.0},
                {"Company": "Precision Drilling", "Ticker": "TSX:PD", "Mkt Cap": 1000.0},
            ],
        }

    def test_us_prefixed_group_includes_companies(self):
        # Regression: group args use EXCHANGE:TICKER but stored Ticker is
        # stripped, so the US section was silently dropped.
        groups = parse_groups(["U.S.: NYSE:HAL NYSE:SLB", "Canada: TSX:PD"])
        md = _format_comps_markdown(
            self._data(), "lease-adjusted", "USD", "1/1/2026", True, groups, []
        )
        assert "### U.S." in md
        assert "Halliburton Company" in md
        assert "SLB N.V." in md
        assert "### Canada" in md
        assert "Precision Drilling" in md
