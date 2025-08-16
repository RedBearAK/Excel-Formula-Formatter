#!/usr/bin/env python3
"""
JavaScript syntax translator for Excel formulas.
File: excel_formula_formatter/javascript_translator.py
"""

import re

from excel_formula_formatter.syntax_translator_base import SyntaxTranslatorBase
from excel_formula_formatter.excel_formula_patterns import (
    cell_ref_all_rgx
)

# Define specific patterns needed for JavaScript translator
js_not_equal_rgx = re.compile(r'!=')


class JavaScriptTranslator(SyntaxTranslatorBase):
    """Translates Excel formulas to JavaScript-like syntax."""
    
    def get_language_name(self) -> str:
        return "JavaScript"
    
    def get_file_extension(self) -> str:
        return ".js"
    
    def format_header_comment(self) -> str:
        return "// Excel Formula (JavaScript syntax for highlighting)"
    
    def format_section_comment(self, comment: str) -> str:
        return f"// {comment}"
    
    def format_function_call(self, function_name: str) -> str:
        return function_name
    
    def format_cell_reference(self, cell_ref: str) -> str:
        # Quote cell references for string highlighting
        return f'"{cell_ref}"'
    
    def format_string_literal(self, string_val: str) -> str:
        return string_val  # Already quoted
    
    def format_number(self, number_val: str) -> str:
        return number_val
    
    def format_operator(self, operator: str) -> str:
        # Convert Excel <> to JavaScript !=
        if operator == '<>':
            return ' != '
        elif operator in ['>=', '<=', '=']:
            return f' {operator} '
        else:
            return f' {operator} '
    
    def format_punctuation(self, punct: str) -> str:
        # Add spacing around function parentheses for readability (same as Excel modes)
        if punct == '(':
            return '( '
        elif punct == ')':
            return ' )'
        else:
            return punct
    
    def reverse_parse_line(self, line: str) -> str:
        """Convert JavaScript-like line back to Excel syntax."""
        # Simple comment removal for JavaScript mode
        comment_pos = line.find('//')
        if comment_pos >= 0:
            return line[:comment_pos].strip()
        return line.strip()
    
    def reverse_parse_cell_reference(self, js_text: str) -> str:
        """Remove quotes from cell references."""
        def unquote_cell_ref(match):
            quoted_text = match.group(0)
            inner_text = quoted_text[1:-1]  # Remove quotes
            # Check if inner text is a cell reference
            if cell_ref_all_rgx.match(inner_text):
                return inner_text
            else:
                return quoted_text  # Keep quotes if not a cell reference
        
        return re.sub(r'"[^"]*"', unquote_cell_ref, js_text)
    
    def reverse_parse_operator(self, js_text: str) -> str:
        """Convert != back to <>."""
        return js_not_equal_rgx.sub('<>', js_text)

# End of file #
