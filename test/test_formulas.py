"""Snapshot tests for formula generation across all dialects."""
import pytest
from capiq_excel.config import FormulaOptions
from capiq_excel.formulas.base import QuerySpec
from capiq_excel.formulas.ciq_builder import CiqBuilder
from capiq_excel.formulas.spg_builder import SpgBuilder
from capiq_excel.formulas.snl_builder import SnlBuilder


@pytest.fixture
def ciq():
    return CiqBuilder()


@pytest.fixture
def spg():
    return SpgBuilder()


@pytest.fixture
def snl():
    return SnlBuilder()


# ============================================================
# CIQ Builder
# ============================================================

class TestCiqBuilderFinancial:
    def test_financial_range_quarterly(self, ciq):
        spec = QuerySpec(
            identifiers=["IQ21835"],
            metric="IQ_TOTAL_REV",
            metric_type="financial",
            frequency="Q",
            num_periods=80,
            label="Sales",
        )
        result = ciq.build_range(spec)
        assert result == '=CIQRANGE("IQ21835", "IQ_TOTAL_REV", IQ_FQ - 80, , , , , , "Sales")'

    def test_financial_range_yearly(self, ciq):
        spec = QuerySpec(
            identifiers=["IQ21835"],
            metric="IQ_TOTAL_REV",
            metric_type="financial",
            frequency="Y",
            num_periods=20,
        )
        result = ciq.build_range(spec)
        assert result == '=CIQRANGE("IQ21835", "IQ_TOTAL_REV", IQ_FY - 20, , , , , , "IQ_TOTAL_REV")'

    def test_financial_range_default_label(self, ciq):
        spec = QuerySpec(
            identifiers=["IQ21835"],
            metric="IQ_COST_REV",
            metric_type="financial",
        )
        result = ciq.build_range(spec)
        assert '"IQ_COST_REV")' in result


class TestCiqBuilderMarket:
    def test_market_range_with_dates(self, ciq):
        spec = QuerySpec(
            identifiers=["IQ21835"],
            metric="IQ_FLOAT_PERCENT",
            metric_type="market",
            begin_date="01/01/2020",
            end_date="12/31/2024",
            label="Float %",
        )
        result = ciq.build_range(spec)
        assert result == '=CIQRANGE("IQ21835", "IQ_FLOAT_PERCENT", "01/01/2020", "12/31/2024", , , , , "Float %")'


class TestCiqBuilderIdLookup:
    def test_id_lookup(self, ciq):
        result = ciq.build_identifier_lookup("MSFT")
        assert result == '=CIQ("MSFT","IQ_COMPANY_ID")'

    def test_name_lookup(self, ciq):
        result = ciq.build_identifier_lookup("AAPL", "IQ_COMPANY_NAME_QUICK_MATCH")
        assert result == '=CIQ("AAPL","IQ_COMPANY_NAME")'


class TestCiqBuilderTable:
    def test_table_falls_back_to_range(self, ciq):
        spec = QuerySpec(
            identifiers=["IQ21835"],
            metric="IQ_TOTAL_REV",
            metric_type="financial",
            frequency="Q",
            num_periods=80,
            label="Sales",
        )
        assert ciq.build_table(spec) == ciq.build_range(spec)


# ============================================================
# SPG Builder
# ============================================================

class TestSpgBuilderSingleValue:
    def test_basic_single_value(self, spg):
        spec = QuerySpec(identifiers=["IQ21835"], metric="IQ_TOTAL_REV")
        result = spg.build_single_value(spec)
        assert result == '=SPG("IQ21835", "IQ_TOTAL_REV", "FQ0")'

    def test_single_value_with_period(self, spg):
        spec = QuerySpec(identifiers=["SPGI"], metric="SP_TOTAL_REV", period="FY2020")
        result = spg.build_single_value(spec)
        assert result == '=SPG("SPGI", "SP_TOTAL_REV", "FY2020")'

    def test_single_value_with_options(self, spg):
        spec = QuerySpec(
            identifiers=["SPGI"],
            metric="SP_TOTAL_REV",
            period="FY0",
            options=FormulaOptions(currency="USD", magnitude="Millions"),
        )
        result = spg.build_single_value(spec)
        assert "Curr=USD,Mag=Millions" in result
        assert result.startswith("=SPG(")


class TestSpgBuilderRange:
    def test_financial_range_uses_period_codes(self, spg):
        spec = QuerySpec(
            identifiers=["IQ21835"],
            metric="IQ_TOTAL_REV",
            metric_type="financial",
            frequency="Q",
            num_periods=80,
        )
        result = spg.build_range(spec)
        # Should use FQ-80 and FQ0, NOT date strings
        assert "FQ-80" in result
        assert "FQ0" in result
        assert result.startswith("=SPGRangeV(")

    def test_financial_range_yearly(self, spg):
        spec = QuerySpec(
            identifiers=["IQ21835"],
            metric="IQ_TOTAL_REV",
            metric_type="financial",
            frequency="Y",
            num_periods=20,
        )
        result = spg.build_range(spec)
        assert "FY-20" in result
        assert "FY0" in result

    def test_range_with_explicit_periods(self, spg):
        spec = QuerySpec(
            identifiers=["SPGI"],
            metric="SP_TOTAL_REV",
            begin_date="FQ12017",
            end_date="FQ42020",
        )
        result = spg.build_range(spec)
        assert '"FQ12017"' in result
        assert '"FQ42020"' in result

    def test_range_with_options(self, spg):
        spec = QuerySpec(
            identifiers=["SPGI"],
            metric="SP_TOTAL_REV",
            frequency="Y",
            num_periods=5,
            options=FormulaOptions(currency="EUR", magnitude="Billions"),
        )
        result = spg.build_range(spec)
        assert "Curr=EUR,Mag=Billions" in result


class TestSpgBuilderTable:
    def test_table_formula(self, spg):
        spec = QuerySpec(
            identifiers=["IQ21835"],
            metric="IQ_TOTAL_REV",
            period="FQ0",
        )
        result = spg.build_table(spec)
        assert result.startswith("=SPGTable(")
        assert '"IQ21835"' in result
        assert '"FQ0"' in result

    def test_table_with_options(self, spg):
        spec = QuerySpec(
            identifiers=["SPGI"],
            metric="SP_TOTAL_REV",
            period="FY2020",
            options=FormulaOptions(currency="USD"),
        )
        result = spg.build_table(spec)
        assert "Curr=USD" in result


class TestSpgBuilderIdLookup:
    def test_id_lookup(self, spg):
        result = spg.build_identifier_lookup("MSFT")
        assert result == '=SPG("MSFT", "IQ_COMPANY_ID_QUICK_MATCH")'


# ============================================================
# SNL Builder
# ============================================================

class TestSnlBuilderSingleValue:
    def test_basic_snldata(self, snl):
        spec = QuerySpec(
            identifiers=["4023623"],
            metric="SNL_TOTAL_ASSETS",
            dataset=1,
            secondary_key="2020Q4",
        )
        result = snl.build_single_value(spec)
        assert result == '=SNLData(1, "4023623", "SNL_TOTAL_ASSETS", "2020Q4")'

    def test_snldata_with_tertiary(self, snl):
        spec = QuerySpec(
            identifiers=["4023623"],
            metric="SNL_TOTAL_ASSETS",
            dataset=1,
            secondary_key="2020Q4",
            tertiary_key="Originally Reported",
        )
        result = snl.build_single_value(spec)
        assert '"Originally Reported"' in result

    def test_snldata_with_options(self, snl):
        spec = QuerySpec(
            identifiers=["4023623"],
            metric="SNL_TOTAL_ASSETS",
            dataset=1,
            secondary_key="MRY",
            options=FormulaOptions(currency="USD", magnitude="Millions"),
        )
        result = snl.build_single_value(spec)
        assert "Curr=USD,Mag=Millions" in result

    def test_snldata_default_dataset(self, snl):
        spec = QuerySpec(identifiers=["4023623"], metric="SNL_TOTAL_ASSETS")
        result = snl.build_single_value(spec)
        assert result.startswith("=SNLData(1,")


class TestSnlBuilderRange:
    def test_market_data_range(self, snl):
        spec = QuerySpec(
            identifiers=["SPGI"],
            metric="SNL_CLOSE_PRICE",
            metric_type="market",
            begin_date="01/01/2020",
            end_date="12/31/2024",
        )
        result = snl.build_range(spec)
        assert result.startswith("=SNLMarkets(")
        assert '"01/01/2020"' in result
        assert '"12/31/2024"' in result

    def test_non_market_falls_back_to_snldata(self, snl):
        spec = QuerySpec(
            identifiers=["4023623"],
            metric="SNL_TOTAL_ASSETS",
            metric_type="financial",
            dataset=1,
            secondary_key="MRY",
        )
        result = snl.build_range(spec)
        assert result.startswith("=SNLData(")


class TestSnlBuilderTable:
    def test_snltable(self, snl):
        spec = QuerySpec(
            identifiers=["4023623"],
            metric="SNL_TOTAL_ASSETS",
            dataset=1,
            secondary_key="MRY",
        )
        result = snl.build_table(spec)
        assert result.startswith("=SNLTable(1,")
        assert '"MRY"' in result

    def test_snltable_with_options(self, snl):
        spec = QuerySpec(
            identifiers=["4023623"],
            metric="SNL_TOTAL_ASSETS",
            dataset=266637,
            secondary_key="MRQ",
            options=FormulaOptions(currency="GBP"),
        )
        result = snl.build_table(spec)
        assert "=SNLTable(266637," in result
        assert "Curr=GBP" in result


class TestSnlBuilderIdLookup:
    def test_id_lookup(self, snl):
        result = snl.build_identifier_lookup("MSFT")
        assert result == '=SNLData(1, "MSFT", "IQ_COMPANY_ID_QUICK_MATCH")'


class TestSnlBuilderSpecialFunctions:
    def test_query(self, snl):
        result = snl.build_query("My Saved Screen")
        assert result == '=SNLQuery("My Saved Screen")'

    def test_query_with_fields(self, snl):
        result = snl.build_query("My Screen", field_range="A1:A5")
        assert result == '=SNLQuery("My Screen", A1:A5)'

    def test_definition(self, snl):
        result = snl.build_definition(1, "SNL_TOTAL_ASSETS")
        assert result == '=SNLDefinition(1, "SNL_TOTAL_ASSETS")'

    def test_definition_with_terminology(self, snl):
        result = snl.build_definition(1, "SNL_TOTAL_ASSETS", terminology="ja-JP")
        assert result == '=SNLDefinition(1, "SNL_TOTAL_ASSETS", "ja-JP")'

    def test_convert(self, snl):
        result = snl.build_convert(1, "A1:A10", 17)
        assert result == "=SNLConvert(1, A1:A10, 17)"


# ============================================================
# Cross-Dialect Parity
# ============================================================

class TestDialectParity:
    """Ensure all builders accept the same QuerySpec inputs."""

    def test_same_spec_all_three_dialects(self, ciq, spg, snl):
        spec = QuerySpec(
            identifiers=["IQ21835"],
            metric="IQ_TOTAL_REV",
            metric_type="financial",
            frequency="Q",
            num_periods=10,
            label="Revenue",
        )
        ciq_result = ciq.build_range(spec)
        spg_result = spg.build_range(spec)
        snl_result = snl.build_range(spec)  # non-market falls back to SNLData

        # Each dialect produces a valid formula starting with its prefix
        assert ciq_result.startswith("=CIQ")
        assert spg_result.startswith("=SPG")
        assert snl_result.startswith("=SNL")

    def test_all_builders_have_table_method(self, ciq, spg, snl):
        spec = QuerySpec(
            identifiers=["IQ21835"],
            metric="IQ_TOTAL_REV",
        )
        # All should return strings without error
        assert isinstance(ciq.build_table(spec), str)
        assert isinstance(spg.build_table(spec), str)
        assert isinstance(snl.build_table(spec), str)

    def test_dialect_names(self, ciq, spg, snl):
        assert ciq.dialect_name() == "ciq"
        assert spg.dialect_name() == "spg"
        assert snl.dialect_name() == "snl"
