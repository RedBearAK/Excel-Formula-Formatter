#!/usr/bin/env python3
"""
Excel Formula Formatter Package
File: excel_formula_formatter/__init__.py

A package for transforming Excel formulas into readable, syntax-highlighted
formats with reversible round-trip transformations.

The legacy ExcelFormulaFormatter class was removed in the consolidation to
a single modular architecture. All functionality is provided by
ModularExcelFormatter and its four modes:
    j = JavaScript, a = Annotated Excel, p = Plain Excel, c = Compact Excel
"""

from excel_formula_formatter.version_info import get_version_dotted, get_version_string
from excel_formula_formatter.modular_excel_formatter import (
    ModularExcelFormatter,
    auto_format_with_mode,
    detect_current_mode,
    safe_mode_switch,
)


# Package metadata
__version__ = get_version_dotted()
__author__ = "Excel Formula Tools"
__description__ = "Transform Excel formulas to readable syntax for better editing"

__all__ = [
    "ModularExcelFormatter",
    "auto_format_with_mode",
    "detect_current_mode",
    "get_version_string",
    "safe_mode_switch",
]

# End of file #
