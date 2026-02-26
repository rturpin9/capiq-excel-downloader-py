"""
Legacy CIQ formula builder.

Generates CIQ(), CIQRANGE(), and CIQRANGEA() formulas matching
the legacy S&P Capital IQ Excel Plug-in syntax.
"""
from __future__ import annotations

from capiq_excel.formulas.base import DialectBuilder, QuerySpec
from capiq_excel.tools.dates import freq_and_periods_to_begin_date_str, today_as_str


class CiqBuilder(DialectBuilder):
    """Emits legacy CIQ-family formulas."""

    def dialect_name(self) -> str:
        return "ciq"

    def build_single_value(self, spec: QuerySpec) -> str:
        if not spec.identifiers:
            raise ValueError("QuerySpec.identifiers must not be empty")
        return f'=CIQ("{spec.identifiers[0]}", "{spec.metric}")'

    def build_range(self, spec: QuerySpec) -> str:
        if not spec.identifiers:
            raise ValueError("QuerySpec.identifiers must not be empty")
        label = spec.label or spec.metric

        if spec.metric_type == "market":
            return self._build_market_range(spec, label)
        elif spec.metric_type == "ownership":
            return self._build_holdings_range(spec, label)
        else:
            return self._build_financial_range(spec, label)

    def build_table(self, spec: QuerySpec) -> str:
        """CIQ does not have a native table function. Falls back to build_range."""
        return self.build_range(spec)

    def build_identifier_lookup(self, search_str: str, field_name: str = "IQ_COMPANY_ID_QUICK_MATCH") -> str:
        # CIQ() works in both legacy and Pro compat mode.
        # CIQRANGEA does NOT work in Pro compat mode (returns function name as string).
        # Map the _QUICK_MATCH field names to their CIQ() equivalents.
        ciq_field = field_name
        if field_name == "IQ_COMPANY_ID_QUICK_MATCH":
            ciq_field = "IQ_COMPANY_ID"
        elif field_name == "IQ_COMPANY_NAME_QUICK_MATCH":
            ciq_field = "IQ_COMPANY_NAME"
        return f'=CIQ("{search_str}","{ciq_field}")'

    # --- private helpers ---

    def _build_financial_range(self, spec: QuerySpec, label: str) -> str:
        freq_char = spec.frequency
        return (
            f'=CIQRANGE("{spec.identifiers[0]}", "{spec.metric}", '
            f'IQ_F{freq_char} - {spec.num_periods}, , , , , , "{label}")'
        )

    def _build_market_range(self, spec: QuerySpec, label: str) -> str:
        begin = spec.begin_date or freq_and_periods_to_begin_date_str(spec.frequency, spec.num_periods)
        end = spec.end_date or today_as_str()
        return (
            f'=CIQRANGE("{spec.identifiers[0]}", "{spec.metric}", '
            f'"{begin}", "{end}", , , , , "{label}")'
        )

    def _build_holdings_range(self, spec: QuerySpec, label: str) -> str:
        date_str = spec.begin_date or today_as_str()
        return (
            f'=CIQRANGE("{spec.identifiers[0]}", "{spec.metric}", '
            f'1, 50000, "{date_str}", , , , "{label}")'
        )
