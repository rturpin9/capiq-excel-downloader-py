"""Tests for config module."""
import os
from capiq_excel.config import (
    CapiqConfig, FormulaDialect, AddinMode, RefreshScope, FormulaOptions
)


class TestCapiqConfigDefaults:
    def test_defaults_are_legacy_compatible(self):
        cfg = CapiqConfig()
        assert cfg.formula_dialect == FormulaDialect.AUTO
        assert cfg.addin_mode == AddinMode.AUTO
        assert cfg.refresh_scope == RefreshScope.WORKSHEET
        assert cfg.freq == "Q"
        assert cfg.num_periods == 80

    def test_retry_defaults(self):
        cfg = CapiqConfig()
        assert cfg.retry.max_retries == 3
        assert cfg.retry.timeout_seconds == 240
        assert cfg.retry.restart_interval == 500


class TestCapiqConfigFromEnv:
    def test_reads_dialect_from_env(self, monkeypatch):
        monkeypatch.setenv("CAPIQ_FORMULA_DIALECT", "spg")
        cfg = CapiqConfig.from_env()
        assert cfg.formula_dialect == FormulaDialect.SPG

    def test_reads_mode_from_env(self, monkeypatch):
        monkeypatch.setenv("CAPIQ_ADDIN_MODE", "legacy")
        cfg = CapiqConfig.from_env()
        assert cfg.addin_mode == AddinMode.LEGACY

    def test_defaults_when_env_not_set(self):
        for key in ("CAPIQ_FORMULA_DIALECT", "CAPIQ_ADDIN_MODE", "CAPIQ_REFRESH_SCOPE"):
            os.environ.pop(key, None)
        cfg = CapiqConfig.from_env()
        assert cfg.formula_dialect == FormulaDialect.AUTO
        assert cfg.addin_mode == AddinMode.AUTO

    def test_reads_formula_options_and_retry_delay(self, monkeypatch):
        monkeypatch.setenv("CAPIQ_CURRENCY", "USD")
        monkeypatch.setenv("CAPIQ_MAGNITUDE", "Millions")
        monkeypatch.setenv("CAPIQ_RETRY_DELAY", "7")
        cfg = CapiqConfig.from_env()
        assert cfg.formula_options.to_spg_options_string() == "Curr=USD,Mag=Millions"
        assert cfg.retry.retry_delay_seconds == 7


class TestResolveDialect:
    def test_auto_with_ciq_compat(self):
        cfg = CapiqConfig(formula_dialect=FormulaDialect.AUTO)
        assert cfg.resolve_dialect(ciq_compat_available=True) == FormulaDialect.CIQ

    def test_auto_without_ciq_compat(self):
        cfg = CapiqConfig(formula_dialect=FormulaDialect.AUTO)
        assert cfg.resolve_dialect(ciq_compat_available=False) == FormulaDialect.SPG

    def test_explicit_dialect_ignores_compat(self):
        cfg = CapiqConfig(formula_dialect=FormulaDialect.SPG)
        assert cfg.resolve_dialect(ciq_compat_available=True) == FormulaDialect.SPG


class TestFormulaOptions:
    def test_empty_options_returns_none(self):
        opts = FormulaOptions()
        assert opts.to_spg_options_string() is None

    def test_single_option(self):
        opts = FormulaOptions(currency="USD")
        assert opts.to_spg_options_string() == "Curr=USD"

    def test_multiple_options(self):
        opts = FormulaOptions(currency="EUR", magnitude="Millions", conversion_method="Recommended")
        result = opts.to_spg_options_string()
        assert result == "Curr=EUR,Mag=Millions,ConvMethod=Recommended"

    def test_all_options(self):
        opts = FormulaOptions(
            currency="GBP",
            magnitude="Billions",
            conversion_method="MRSpot",
            null_display="NA",
            terminology="en-GB",
        )
        result = opts.to_spg_options_string()
        assert "Curr=GBP" in result
        assert "Mag=Billions" in result
        assert "ConvMethod=MRSpot" in result
        assert "NullValue=NA" in result
        assert "Terminology=en-GB" in result

    def test_reported_currency(self):
        opts = FormulaOptions(currency="RC")
        assert opts.to_spg_options_string() == "Curr=RC"

    def test_terminology_only(self):
        opts = FormulaOptions(terminology="ja-JP")
        assert opts.to_spg_options_string() == "Terminology=ja-JP"
