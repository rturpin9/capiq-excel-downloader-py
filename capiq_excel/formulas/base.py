"""
Formula builder interface and canonical query specification.

All formula generation goes through QuerySpec -> DialectBuilder -> Excel formula string.
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Optional

from capiq_excel.config import FormulaOptions, MetricType


@dataclass
class QuerySpec:
    """Canonical, dialect-agnostic description of a data query.

    This is the single input shape that all dialect builders consume.
    Fields map to the union of CIQ, SPG, and SNL function parameters.
    """
    identifiers: list[str]
    metric: str
    metric_type: MetricType = MetricType.FINANCIAL

    # Period specification — builders interpret these per-dialect
    period: Optional[str] = None      # dialect-native period code (e.g. "FQ0", "IQ_FQ - 80", "2013Q2")
    frequency: str = "Q"             # "Q" | "Y" | "M" | "D" | "W"
    num_periods: int = 80
    begin_date: Optional[str] = None  # absolute date string (mm/dd/yyyy) or period code
    end_date: Optional[str] = None

    # SNL-specific keys (SNLData, SNLMarkets, SNLTable)
    dataset: Optional[int] = None         # SNL dataset number (e.g. 1=Companies Classic, 266637=Companies)
    secondary_key: Optional[str] = None   # e.g. fiscal period for financials
    tertiary_key: Optional[str] = None    # e.g. reporting basis

    # Ownership-specific
    start_rank: Optional[int] = None
    end_rank: Optional[int] = None

    # Output control
    label: Optional[str] = None
    options: FormulaOptions = field(default_factory=FormulaOptions)


class DialectBuilder(ABC):
    """Abstract base for dialect-specific formula generators."""

    def identifier_lookup_spills_right(self) -> bool:
        """Whether identifier lookup results land one cell to the right.

        CIQRANGEA uses a horizontal spill layout.  SPG and SNL lookup
        functions return their value in the formula cell itself.
        """
        return False

    @staticmethod
    def _validate_spec(spec: QuerySpec) -> None:
        """Validate that a QuerySpec has required fields."""
        if not spec.identifiers:
            raise ValueError("QuerySpec.identifiers must not be empty")

    @abstractmethod
    def build_single_value(self, spec: QuerySpec) -> str:
        """Build a single-value formula (e.g., CIQ, SPG, SNLData)."""
        ...

    @abstractmethod
    def build_range(self, spec: QuerySpec) -> str:
        """Build a range formula (e.g., CIQRANGE, SPGRangeV, SNLMarkets)."""
        ...

    @abstractmethod
    def build_table(self, spec: QuerySpec) -> str:
        """Build a table formula (e.g., SPGTable, SNLTable).

        Table functions fill a region from a single cell with multiple
        identifiers and fields.
        """
        ...

    @abstractmethod
    def build_identifier_lookup(self, search_str: str, field_name: str = "IQ_COMPANY_ID_QUICK_MATCH") -> str:
        """Build an identifier lookup formula."""
        ...

    @abstractmethod
    def dialect_name(self) -> str:
        """Return the dialect identifier (e.g., 'ciq', 'spg', 'snl')."""
        ...
