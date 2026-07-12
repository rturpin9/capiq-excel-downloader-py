"""
A tool to drive Excel using the Capital IQ plugin to download Capital IQ data. Useful for downloading
large data sets from Capital IQ.

Supports both the legacy CIQ plugin and Capital IQ Pro (SPG/SNL).
"""
from capiq_excel.main import (
    download_data,
    download_data_for_capiq_ids
)
from capiq_excel.config import (
    CapiqConfig,
    FormulaDialect,
    AddinMode,
    RefreshScope,
    MetricType,
    FormulaOptions,
)
from capiq_excel.formulas import get_builder
from capiq_excel.formulas.base import QuerySpec, DialectBuilder

__all__ = [
    "download_data",
    "download_data_for_capiq_ids",
    "CapiqConfig",
    "FormulaDialect",
    "AddinMode",
    "RefreshScope",
    "MetricType",
    "FormulaOptions",
    "get_builder",
    "QuerySpec",
    "DialectBuilder",
]
