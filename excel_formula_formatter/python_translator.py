#!/usr/bin/env python3
"""
Python syntax translator for Excel formulas.
File: excel_formula_formatter/python_translator.py
"""

import re

from excel_formula_formatter.syntax_translator_base import SyntaxTranslatorBase
from excel_formula_formatter.excel_formula_patterns import cell_ref_all_rgx


class PythonTranslator(SyntaxTranslatorBase):
    """Translates Excel formulas to Python-like syntax."""
    
    def get_language_name(self) -> str:
        return "Python"
    
    def get_file_extension(self) -> str:
        return ".py" 
    
    def format_header_comment(self) -> str:
        return "# Excel Formula (Python syntax for highlighting)"
    
    def format_section_comment(self, comment: str) -> str:
        return f"# {comment}"
    
    def format_function_call(self, function_name: str) -> str:
        # Convert Excel functions to Pythonic names
        python_functions = {
            'SUM': 'sum',
            'IF': 'if_else',
            'AND': 'all',
            'OR': 'any',
            'NOT': 'not',
            'CONCATENATE': 'join',
            'LEN': 'len',
            'COUNT': 'count',
            'AVERAGE': 'mean',
            'MAX': 'max',
            'MIN': 'min',
            'LET': 'let'
        }
        return python_functions.get(function_name.upper(), function_name.lower())
    
    def format_cell_reference(self, cell_ref: str) -> str:
        # Convert cell references to Python-like syntax
        if ':' in cell_ref:
            # Range: A1:A10 → sheet["A1":"A10"] 
            return f'sheet["{cell_ref}"]'
        else:
            # Single cell: A1 → sheet["A1"]
            return f'sheet["{cell_ref}"]'
    
    def format_string_literal(self, string_val: str) -> str:
        return string_val  # Keep as-is
    
    def format_number(self, number_val: str) -> str:
        return number_val
    
    def format_operator(self, operator: str) -> str:
        # Convert operators to Python equivalents
        python_operators = {
            '<>': ' != ',
            '=': ' == ',  # Excel = means comparison, not assignment
            '&': ' + ',   # Excel & is string concatenation
        }
        
        if operator in python_operators:
            return python_operators[operator]
        elif operator in ['>=', '<=', '>', '<', '+', '-', '*', '/']:
            return f' {operator} '
        else:
            return f' {operator} '
    
    def format_punctuation(self, punct: str) -> str:
        return punct
    
    def reverse_parse_line(self, line: str) -> str:
        """Convert Python-like line back to Excel syntax."""
        # Remove comments
        line = re.sub(r'\s*#.*$', '', line)
        
        # Convert Python functions back to Excel
        excel_functions = {
            'sum': 'SUM',
            'if_else': 'IF', 
            'all': 'AND',
            'any': 'OR',
            'not': 'NOT',
            'join': 'CONCATENATE',
            'len': 'LEN',
            'count': 'COUNT',
            'mean': 'AVERAGE',
            'max': 'MAX',
            'min': 'MIN',
            'let': 'LET'
        }
        
        for py_func, excel_func in excel_functions.items():
            line = re.sub(rf'\b{py_func}\b', excel_func, line)
        
        return line.strip()
    
    def reverse_parse_cell_reference(self, py_text: str) -> str:
        """Convert sheet["A1"] back to A1."""
        # Pattern: sheet["A1:A10"] or sheet["A1"] 
        def convert_cell_ref(match):
            full_match = match.group(0)
            cell_part = match.group(1)  # The quoted part
            return cell_part  # Return just the cell reference
        
        # Match sheet["..."] patterns and extract the cell reference
        result = re.sub(r'sheet\["([^"]+)"\]', convert_cell_ref, py_text)
        return result
    
    def reverse_parse_operator(self, py_text: str) -> str:
        """Convert Python operators back to Excel."""
        # Reverse the operator mapping
        result = py_text
        result = re.sub(r'\s*!=\s*', '<>', result)
        result = re.sub(r'\s*==\s*', '=', result)  # Python == back to Excel =
        # Note: Be careful with + operator - could be addition or concatenation
        return result

# End of file #
