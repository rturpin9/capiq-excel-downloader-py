"""
Capital IQ Pro SPG formula builder.

Generates SPG(), SPGRangeV(), and SPGTable() formulas for the
Capital IQ Pro Office add-in.

SPG period syntax uses native period codes:
  - Relative fiscal:   FY0, FY-1, FQ0, FQ-1, FH0, LTM, YTD, NTM
  - Relative calendar:  CY0, CY-1, CQ0, CQ-1
  - Absolute fiscal:   FY2020, FQ12020, FH12020, LTM22017, YTD22017
  - Absolute calendar:  CY2020, CQ12020

Options are passed as a key=value string:
  "Curr=USD,Mag=Millions,ConvMethod=Recommended"
"""
from __future__ import annotations

from capiq_excel.config import MetricType
from capiq_excel.formulas.base import DialectBuilder, QuerySpec
from capiq_excel.tools.dates import freq_and_periods_to_begin_date_str, today_as_str


# Map our canonical frequency codes to SPG period prefix
_FREQ_TO_PERIOD_PREFIX = {
    "Q": "FQ",
    "Y": "FY",
}


class SpgBuilder(DialectBuilder):
    """Emits Capital IQ Pro SPG-family formulas."""

    def dialect_name(self) -> str:
        return "spg"

    def build_single_value(self, spec: QuerySpec) -> str:
        """=SPG(Identifier, Metric, Period, AsOfDate, Options)"""
        self._validate_spec(spec)
        identifier = spec.identifiers[0]
        period = spec.period or _default_relative_period(spec.frequency)
        opts = spec.options.to_spg_options_string()

        parts = [
            f'"{identifier}"',
            f'"{spec.metric}"',
            f'"{period}"',
        ]
        # Add options at the end if present
        if opts:
            parts.append("")  # skip AsOfDate
            parts.append(f'"{opts}"')

        return f'=SPG({", ".join(parts)})'

    def build_range(self, spec: QuerySpec) -> str:
        """=SPGRangeV(Identifier, Metric, BeginPeriod, EndPeriod, Options)"""
        self._validate_spec(spec)
        identifier = spec.identifiers[0]
        opts = spec.options.to_spg_options_string()

        if spec.metric_type == MetricType.MARKET:
            begin = spec.begin_date or freq_and_periods_to_begin_date_str(
                spec.frequency,
                spec.num_periods,
            )
            end = spec.end_date or today_as_str()
        else:
            begin = spec.begin_date or _begin_period_code(spec.frequency, spec.num_periods)
            end = spec.end_date or _end_period_code(spec.frequency)

        parts = [
            f'"{identifier}"',
            f'"{spec.metric}"',
            f'"{begin}"',
            f'"{end}"',
        ]
        if opts:
            parts.append(f'"{opts}"')

        return f'=SPGRangeV({", ".join(parts)})'

    def build_table(self, spec: QuerySpec) -> str:
        """=SPGTable(IdentifierRange, MetricRange, PeriodRange, Options)

        SPGTable fills a region from a single cell with data for multiple
        identifiers x fields x periods. The cell ranges for identifiers,
        metrics, and periods are typically cell references in real use.
        For formula generation, we emit the literal values.
        """
        self._validate_spec(spec)
        identifier = spec.identifiers[0]
        period = spec.period or _default_relative_period(spec.frequency)
        opts = spec.options.to_spg_options_string()

        parts = [
            f'"{identifier}"',
            f'"{spec.metric}"',
            f'"{period}"',
        ]
        if opts:
            parts.append(f'"{opts}"')

        return f'=SPGTable({", ".join(parts)})'

    def build_identifier_lookup(self, search_str: str, field_name: str = "IQ_COMPANY_ID_QUICK_MATCH") -> str:
        """SPG identifier lookup — same as single-value with ID metric."""
        return f'=SPG("{search_str}", "{field_name}")'


# --- period code helpers ---

def _get_period_prefix(frequency: str) -> str:
    """Return the SPG period prefix for a canonical frequency code.

    Raises ValueError for unsupported frequencies.
    """
    prefix = _FREQ_TO_PERIOD_PREFIX.get(frequency)
    if prefix is None:
        raise ValueError(
            f"SPG builder does not support frequency '{frequency}'. "
            f"Supported: {', '.join(sorted(_FREQ_TO_PERIOD_PREFIX.keys()))}"
        )
    return prefix


def _default_relative_period(frequency: str) -> str:
    """Return the 'latest' relative period code for a frequency.

    e.g. Q -> FQ0, Y -> FY0
    """
    return f"{_get_period_prefix(frequency)}0"


def _begin_period_code(frequency: str, num_periods: int) -> str:
    """Generate a relative begin-period code.

    e.g. frequency=Q, num_periods=80 -> "FQ-80"
         frequency=Y, num_periods=20 -> "FY-20"
    """
    return f"{_get_period_prefix(frequency)}-{num_periods}"


def _end_period_code(frequency: str) -> str:
    """Generate the 'current' end-period code.

    e.g. Q -> FQ0, Y -> FY0
    """
    return _default_relative_period(frequency)
