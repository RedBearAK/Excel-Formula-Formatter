#!/usr/bin/env python3
"""
Excel Formula Formatter Package
File: excel_formula_formatter/__init__.py

A package for transforming Excel formulas into JavaScript-like syntax 
with proper indentation and syntax highlighting support.
"""

from excel_formula_formatter.excel_formula_formatter import ExcelFormulaFormatter
from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


# Package metadata
__version__ = "20250815.0"
__author__ = "Excel Formula Tools"
__description__ = "Transform Excel formulas to JavaScript-like syntax for better editing"

# Main exports - recommend using ModularExcelFormatter for new code
__all__ = [
    "ExcelFormulaFormatter",  # Legacy formatter
    "ModularExcelFormatter",  # Recommended for new code
]

# End of file #
