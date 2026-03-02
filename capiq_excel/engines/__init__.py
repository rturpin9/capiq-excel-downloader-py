"""High-level engines for comp tables, charts, and ID lookup.

Each engine builds an Excel workbook with CIQ/SPG formulas, launches Excel
via COM, waits for the add-in to refresh, extracts results, and returns
structured data.
"""
