#!/usr/bin/env python3
"""
Excel Formula Formatter Package
File: excel_formula_formatter/__init__.py

A package for transforming Excel formulas into JavaScript-like syntax 
with proper indentation and syntax highlighting support.
"""

from excel_formula_formatter.excel_formula_formatter import ExcelFormulaFormatter


# Package metadata
__version__ = "1.0.0"
__author__ = "Excel Formula Tools"
__description__ = "Transform Excel formulas to JavaScript-like syntax for better editing"

# Main exports
__all__ = [
    "ExcelFormulaFormatter",
]

# End of file #
