"""
Legacy CIQ formula builder.

Generates CIQ(), CIQRANGE(), and CIQRANGEA() formulas matching
the legacy S&P Capital IQ Excel Plug-in syntax.
"""
from __future__ import annotations

from capiq_excel.config import MetricType
from capiq_excel.formulas.base import DialectBuilder, QuerySpec
from capiq_excel.tools.dates import freq_and_periods_to_begin_date_str, today_as_str


class CiqBuilder(DialectBuilder):
    """Emits legacy CIQ-family formulas."""

    def identifier_lookup_spills_right(self) -> bool:
        return True

    def dialect_name(self) -> str:
        return "ciq"

    def build_single_value(self, spec: QuerySpec) -> str:
        self._validate_spec(spec)
        return f'=CIQ("{spec.identifiers[0]}", "{spec.metric}")'

    def build_range(self, spec: QuerySpec) -> str:
        self._validate_spec(spec)
        label = spec.label or spec.metric

        if spec.metric_type == MetricType.MARKET:
            return self._build_market_range(spec, label)
        elif spec.metric_type == MetricType.OWNERSHIP:
            return self._build_holdings_range(spec, label)
        else:
            return self._build_financial_range(spec, label)

    def build_table(self, spec: QuerySpec) -> str:
        """CIQ does not have a native table function. Falls back to build_range."""
        return self.build_range(spec)

    def build_identifier_lookup(self, search_str: str, field_name: str = "IQ_COMPANY_ID_QUICK_MATCH") -> str:
        # CIQRANGEA handles all identifier types (tickers, names, CUSIPs, ISINs).
        # In Pro CIQ compat mode, the formula cell shows the function name as a
        # string, but the actual result spills into the adjacent cell to the right.
        # The workbook layout accounts for this with a blank spill column.
        return f'=CIQRANGEA("{search_str}","{field_name}",1,1)'

    # --- private helpers ---

    def _build_financial_range(self, spec: QuerySpec, label: str) -> str:
        freq_char = spec.frequency
        return (
            f'=CIQRANGE("{spec.identifiers[0]}", "{spec.metric}", '
            f'IQ_F{freq_char} - {spec.num_periods}, , , , , , , "{label}")'
        )

    def _build_market_range(self, spec: QuerySpec, label: str) -> str:
        begin = spec.begin_date or freq_and_periods_to_begin_date_str(spec.frequency, spec.num_periods)
        end = spec.end_date or today_as_str()
        return (
            f'=CIQRANGE("{spec.identifiers[0]}", "{spec.metric}", '
            f'"{begin}", "{end}", , , , , , "{label}")'
        )

    def _build_holdings_range(self, spec: QuerySpec, label: str) -> str:
        date_str = spec.begin_date or today_as_str()
        return (
            f'=CIQRANGE("{spec.identifiers[0]}", "{spec.metric}", '
            f'1, 50000, "{date_str}", , , , "{label}")'
        )
