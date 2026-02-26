"""
Legacy SNL formula builder for Capital IQ Pro.

Generates SNLData(), SNLMarkets(), SNLTable(), SNLQuery(), SNLConvert(),
and SNLDefinition() formulas. These are the SNL-heritage functions that
coexist with SPG functions in the Capital IQ Pro add-in.

SNL period syntax uses a different format from SPG:
  - Absolute:  2013Q2, 2012Y, 2014H1, 2012T2 (YTD), 2013L3 (LTM)
  - Relative:  MRQ, MRY, MRH, YTD, LTM, [MRQ-1], [MRH-2]

SNL datasets are identified by numeric IDs:
  - 1 = Companies (Classic), 266637 = Companies, 5 = M&A, etc.
"""
from __future__ import annotations

from typing import Optional

from capiq_excel.formulas.base import DialectBuilder, QuerySpec


class SnlBuilder(DialectBuilder):
    """Emits SNL-family formulas for the Capital IQ Pro add-in."""

    def dialect_name(self) -> str:
        return "snl"

    def build_single_value(self, spec: QuerySpec) -> str:
        """=SNLData(DataSet, SNLID, FieldID, SecondaryKey, TertiaryKey, Options)"""
        if not spec.identifiers:
            raise ValueError("QuerySpec.identifiers must not be empty")
        dataset = spec.dataset or 1
        identifier = spec.identifiers[0]
        secondary = spec.secondary_key or ""
        tertiary = spec.tertiary_key or ""
        opts = spec.options.to_spg_options_string()

        parts = [str(dataset), f'"{identifier}"', f'"{spec.metric}"']

        if secondary:
            parts.append(f'"{secondary}"')
        else:
            parts.append("")

        if tertiary:
            parts.append(f'"{tertiary}"')
        elif opts:
            parts.append("")  # placeholder so options land in right position

        if opts:
            parts.append(f'"{opts}"')

        return f'=SNLData({", ".join(parts)})'

    def build_range(self, spec: QuerySpec) -> str:
        """=SNLMarkets(Display, Entity, FieldID, SecondaryKey, TertiaryKey, StartDate, EndDate, Options)

        SNLMarkets is specifically for market/pricing data with date ranges.
        For non-market data, falls back to SNLData.
        """
        if not spec.identifiers:
            raise ValueError("QuerySpec.identifiers must not be empty")
        if spec.metric_type != "market":
            return self.build_single_value(spec)

        identifier = spec.identifiers[0]
        start = spec.begin_date or ""
        end = spec.end_date or ""
        secondary = spec.secondary_key or ""
        tertiary = spec.tertiary_key or ""
        opts = spec.options.to_spg_options_string()

        parts = [
            f'"{identifier}"',
            f'"{spec.metric}"',
        ]

        if secondary:
            parts.append(f'"{secondary}"')
        else:
            parts.append("")

        if tertiary:
            parts.append(f'"{tertiary}"')
        else:
            parts.append("")

        parts.append(f'"{start}"')
        if end:
            parts.append(f'"{end}"')

        if opts:
            # Pad to the options position if end was omitted
            if not end:
                parts.append("")
            parts.append(f'"{opts}"')

        return f'=SNLMarkets({", ".join(parts)})'

    def build_table(self, spec: QuerySpec) -> str:
        """=SNLTable(DataSet, SNLIDRange, FieldIDRange, SecondaryKeyRange, Options)

        In real use, ID/Field/Key ranges are cell references. For formula
        generation we emit literal values.
        """
        if not spec.identifiers:
            raise ValueError("QuerySpec.identifiers must not be empty")
        dataset = spec.dataset or 1
        identifier = spec.identifiers[0]
        secondary = spec.secondary_key or ""
        opts = spec.options.to_spg_options_string()

        parts = [str(dataset), f'"{identifier}"', f'"{spec.metric}"']

        if secondary:
            parts.append(f'"{secondary}"')
        elif opts:
            parts.append("")

        if opts:
            parts.append(f'"{opts}"')

        return f'=SNLTable({", ".join(parts)})'

    def build_identifier_lookup(self, search_str: str, field_name: str = "IQ_COMPANY_ID_QUICK_MATCH") -> str:
        """SNL identifier lookup via SNLData against Companies dataset."""
        return f'=SNLData(1, "{search_str}", "{field_name}")'

    def build_query(self, query_name: str, field_range: Optional[str] = None,
                    opts: Optional[str] = None) -> str:
        """=SNLQuery(QueryName, FieldIDRange, Options)

        Reruns a saved screen/project by name.
        """
        parts = [f'"{query_name}"']
        if field_range:
            parts.append(field_range)  # cell reference, not quoted
        elif opts:
            parts.append("")
        if opts:
            parts.append(f'"{opts}"')
        return f'=SNLQuery({", ".join(parts)})'

    def build_definition(self, dataset: int, field_id: str,
                         terminology: Optional[str] = None) -> str:
        """=SNLDefinition(DataSet, FieldID, Terminology)"""
        parts = [str(dataset), f'"{field_id}"']
        if terminology:
            parts.append(f'"{terminology}"')
        return f'=SNLDefinition({", ".join(parts)})'

    def build_convert(self, from_dataset: int, id_range: str,
                      to_dataset: int) -> str:
        """=SNLConvert(FromDataSet, SNLIDRange, ToDataSet)

        Converts entities from one dataset to another (e.g. companies -> branches).
        id_range is typically a cell reference.
        """
        return f'=SNLConvert({from_dataset}, {id_range}, {to_dataset})'
