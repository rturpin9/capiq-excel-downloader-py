"""
Formula builder factory.

Use get_builder() to obtain the correct DialectBuilder for the resolved dialect.
"""
from __future__ import annotations

from capiq_excel.config import FormulaDialect
from capiq_excel.formulas.base import DialectBuilder, QuerySpec
from capiq_excel.formulas.ciq_builder import CiqBuilder
from capiq_excel.formulas.spg_builder import SpgBuilder
from capiq_excel.formulas.snl_builder import SnlBuilder

_BUILDERS = {
    FormulaDialect.CIQ: CiqBuilder,
    FormulaDialect.SPG: SpgBuilder,
    FormulaDialect.SNL: SnlBuilder,
}


def get_builder(dialect: FormulaDialect) -> DialectBuilder:
    """Return a DialectBuilder instance for the given dialect.

    Raises ValueError if dialect is AUTO (must be resolved first).
    """
    if dialect == FormulaDialect.AUTO:
        raise ValueError(
            "Dialect must be resolved before getting a builder. "
            "Use config.resolve_dialect() first."
        )
    cls = _BUILDERS.get(dialect)
    if cls is None:
        raise ValueError(f"No builder registered for dialect: {dialect.value}")
    return cls()
