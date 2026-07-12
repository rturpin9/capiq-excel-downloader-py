"""
Smoke tests for the modernized capiq-excel pipeline.

Tests the full integration path without requiring COM/Excel.
Validates that config -> dialect resolution -> builder -> command factories
all wire together correctly.
"""
import pytest
from capiq_excel.config import CapiqConfig, FormulaDialect
from capiq_excel.formulas import get_builder
from capiq_excel.formulas.ciq_builder import CiqBuilder
from capiq_excel.formulas.spg_builder import SpgBuilder
from capiq_excel.workbook.commands import (
    make_financial_command,
    make_market_command,
    make_holdings_command,
    make_id_command,
    make_name_command,
    financial_data_command,
    id_command,
    name_command,
)


# ============================================================
# get_builder factory
# ============================================================

class TestGetBuilder:
    def test_ciq_dialect_returns_ciq_builder(self):
        builder = get_builder(FormulaDialect.CIQ)
        assert isinstance(builder, CiqBuilder)
        assert builder.dialect_name() == "ciq"

    def test_spg_dialect_returns_spg_builder(self):
        builder = get_builder(FormulaDialect.SPG)
        assert isinstance(builder, SpgBuilder)
        assert builder.dialect_name() == "spg"

    def test_auto_dialect_raises(self):
        with pytest.raises(ValueError, match="resolved"):
            get_builder(FormulaDialect.AUTO)


# ============================================================
# Builder-aware command factories
# ============================================================

class TestMakeFinancialCommand:
    def test_ciq_builder_matches_legacy(self):
        builder = get_builder(FormulaDialect.CIQ)
        cmd = make_financial_command(builder)
        result = cmd("IQ21835", "IQ_TOTAL_REV", freq="Q", num_periods=80, data_item_label="Sales")
        # Should produce a CIQRANGE formula
        assert result.startswith("=CIQRANGE(")
        assert '"IQ21835"' in result
        assert '"IQ_TOTAL_REV"' in result
        assert '"Sales"' in result

    def test_spg_builder_produces_spg_range(self):
        builder = get_builder(FormulaDialect.SPG)
        cmd = make_financial_command(builder)
        result = cmd("IQ21835", "IQ_TOTAL_REV", freq="Q", num_periods=80)
        assert result.startswith("=SPGRangeV(")
        assert "FQ-80" in result
        assert "FQ0" in result


class TestMakeMarketCommand:
    def test_ciq_builder_market(self):
        builder = get_builder(FormulaDialect.CIQ)
        cmd = make_market_command(builder)
        result = cmd("IQ21835", "IQ_CLOSEPRICE", freq="Q", num_periods=20)
        assert result.startswith("=CIQRANGE(")
        # Market data should have date strings in CIQ mode
        assert "IQ_CLOSEPRICE" in result

    def test_spg_builder_market(self):
        builder = get_builder(FormulaDialect.SPG)
        cmd = make_market_command(builder)
        result = cmd("IQ21835", "IQ_CLOSEPRICE", freq="Q", num_periods=20)
        assert result.startswith("=SPGRangeV(")


class TestMakeHoldingsCommand:
    def test_ciq_builder_holdings(self):
        builder = get_builder(FormulaDialect.CIQ)
        cmd = make_holdings_command(builder)
        result = cmd("IQ21835", "IQ_HOLDER_NAME", "03/15/2024", data_item_label="Holder")
        assert result.startswith("=CIQRANGE(")
        assert '"Holder"' in result

    def test_spg_builder_holdings(self):
        builder = get_builder(FormulaDialect.SPG)
        cmd = make_holdings_command(builder)
        result = cmd("IQ21835", "IQ_HOLDER_NAME", "03/15/2024")
        assert result.startswith("=SPGRangeV(")


class TestMakeIdCommand:
    def test_ciq_builder_id(self):
        builder = get_builder(FormulaDialect.CIQ)
        cmd = make_id_command(builder)
        result = cmd("MSFT")
        assert result == '=CIQRANGEA("MSFT","IQ_COMPANY_ID_QUICK_MATCH",1,1)'

    def test_spg_builder_id(self):
        builder = get_builder(FormulaDialect.SPG)
        cmd = make_id_command(builder)
        result = cmd("MSFT")
        assert result == '=SPG("MSFT", "IQ_COMPANY_ID_QUICK_MATCH")'


class TestMakeNameCommand:
    def test_ciq_builder_name(self):
        builder = get_builder(FormulaDialect.CIQ)
        cmd = make_name_command(builder)
        result = cmd("AAPL")
        assert result == '=CIQRANGEA("AAPL","IQ_COMPANY_NAME_QUICK_MATCH",1,1)'

    def test_spg_builder_name(self):
        builder = get_builder(FormulaDialect.SPG)
        cmd = make_name_command(builder)
        result = cmd("AAPL")
        assert result == '=SPG("AAPL", "IQ_COMPANY_NAME_QUICK_MATCH")'


# ============================================================
# Legacy command backward compatibility
# ============================================================

class TestLegacyCommandsUnchanged:
    def test_financial_data_command(self):
        result = financial_data_command("IQ21835", "IQ_TOTAL_REV", freq="Q", num_periods=80, data_item_label="Sales")
        assert result == '=CIQRANGE("IQ21835", "IQ_TOTAL_REV", IQ_FQ - 80, , , , , , , "Sales")'

    def test_id_command(self):
        result = id_command("MSFT")
        assert result == '=CIQRANGEA("MSFT","IQ_COMPANY_ID_QUICK_MATCH",1,1)'

    def test_name_command(self):
        result = name_command("AAPL")
        assert result == '=CIQRANGEA("AAPL","IQ_COMPANY_NAME_QUICK_MATCH",1,1)'


# ============================================================
# Config -> Builder -> Command end-to-end
# ============================================================

class TestEndToEndConfigToCommand:
    def test_default_config_produces_ciq(self):
        config = CapiqConfig()
        dialect = config.resolve_dialect()
        builder = get_builder(dialect)
        cmd = make_financial_command(builder)
        result = cmd("IQ21835", "IQ_TOTAL_REV")
        assert result.startswith("=CIQRANGE(")

    def test_spg_config_produces_spg(self):
        config = CapiqConfig(formula_dialect=FormulaDialect.SPG)
        dialect = config.resolve_dialect()
        builder = get_builder(dialect)
        cmd = make_financial_command(builder)
        result = cmd("IQ21835", "IQ_TOTAL_REV")
        assert result.startswith("=SPGRangeV(")

    def test_auto_no_ciq_compat_produces_spg(self):
        config = CapiqConfig(formula_dialect=FormulaDialect.AUTO)
        dialect = config.resolve_dialect(ciq_compat_available=False)
        builder = get_builder(dialect)
        cmd = make_financial_command(builder)
        result = cmd("IQ21835", "IQ_TOTAL_REV")
        assert result.startswith("=SPGRangeV(")

    def test_env_config_to_builder(self, monkeypatch):
        monkeypatch.setenv("CAPIQ_FORMULA_DIALECT", "spg")
        config = CapiqConfig.from_env()
        dialect = config.resolve_dialect()
        builder = get_builder(dialect)
        assert builder.dialect_name() == "spg"


# ============================================================
# Refresh engine token coverage
# ============================================================

class TestRefreshTokens:
    def test_pending_tokens_lowercase(self):
        from capiq_excel.refresh.engine import PENDING_TOKENS
        for tok in PENDING_TOKENS:
            assert tok == tok.lower()

    def test_error_tokens_lowercase(self):
        from capiq_excel.refresh.engine import HARD_ERROR_TOKENS, LEGACY_ERROR_TOKENS
        for tok in HARD_ERROR_TOKENS | LEGACY_ERROR_TOKENS:
            assert tok == tok.lower()

    def test_all_error_tokens_is_union(self):
        from capiq_excel.refresh.engine import ALL_ERROR_TOKENS, HARD_ERROR_TOKENS, LEGACY_ERROR_TOKENS
        assert ALL_ERROR_TOKENS == HARD_ERROR_TOKENS | LEGACY_ERROR_TOKENS


# ============================================================
# CLI module importability
# ============================================================

class TestCLIImport:
    def test_cli_module_imports(self):
        from capiq_excel.cli import main
        assert callable(main)

    def test_cli_status_command(self):
        from capiq_excel.cli import main
        # --help or status should not raise
        result = main(["status"])
        assert result == 0
