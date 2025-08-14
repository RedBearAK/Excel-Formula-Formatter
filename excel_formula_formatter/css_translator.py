#!/usr/bin/env python3
"""
CSS syntax translator for Excel formulas - a creative alternative approach.
File: excel_formula_formatter/css_translator.py
"""

import re

from excel_formula_formatter.syntax_translator_base import SyntaxTranslatorBase


class CSSTranslator(SyntaxTranslatorBase):
    """Translates Excel formulas to CSS-like syntax for unique highlighting."""
    
    def get_language_name(self) -> str:
        return "CSS"
    
    def get_file_extension(self) -> str:
        return ".css"
    
    def format_header_comment(self) -> str:
        return "/* Excel Formula (CSS syntax for highlighting) */"
    
    def format_section_comment(self, comment: str) -> str:
        return f"/* {comment} */"
    
    def format_function_call(self, function_name: str) -> str:
        # Functions become CSS classes (yellow/orange highlighting)
        return f".{function_name}"
    
    def format_cell_reference(self, cell_ref: str) -> str:
        # Cell references become CSS IDs (blue highlighting)
        # Replace special chars that aren't valid in CSS IDs
        clean_ref = cell_ref.replace(':', '_to_').replace('!', '_sheet_').replace('$', 'abs_')
        return f"#{clean_ref}"
    
    def format_string_literal(self, string_val: str) -> str:
        # String literals as CSS content values
        return f"content: {string_val}"
    
    def format_number(self, number_val: str) -> str:
        # Numbers as CSS property values
        return f"value: {number_val}"
    
    def format_operator(self, operator: str) -> str:
        # Operators as CSS pseudo-selectors or attribute selectors
        css_operators = {
            '>': ' > ',       # Child combinator 
            '<': ' < ',       # Custom (not valid CSS, but visually clear)
            '>=': ' >= ',     # Custom
            '<=': ' <= ',     # Custom  
            '<>': ' != ',     # Custom
            '=': ' = ',       # Custom
            '+': ' + ',       # Adjacent sibling combinator
            '-': ' - ',       # Custom
            '*': ' * ',       # Universal selector (modified)
            '/': ' / ',       # Custom
            '&': ' & ',       # Custom (concatenation)
        }
        
        return css_operators.get(operator, f' {operator} ')
    
    def format_punctuation(self, punct: str) -> str:
        if punct == '(':
            return ' {'
        elif punct == ')':
            return ' }'
        elif punct == ',':
            return ';\n'  # CSS property separator
        else:
            return punct
    
    def reverse_parse_line(self, line: str) -> str:
        """Convert CSS-like line back to Excel syntax."""
        # Remove CSS comments
        line = re.sub(r'/\*.*?\*/', '', line)
        line = line.strip()
        
        if not line or line.endswith(';'):
            return ""
        
        return line
    
    def reverse_parse_cell_reference(self, css_text: str) -> str:
        """Convert #A1_to_A10 back to A1:A10."""
        def convert_css_id(match):
            css_id = match.group(1)  # Everything after #
            # Reverse the transformations
            excel_ref = css_id.replace('_to_', ':').replace('_sheet_', '!').replace('abs_', '$')
            return excel_ref
        
        result = re.sub(r'#([A-Za-z0-9_]+)', convert_css_id, css_text)
        return result
    
    def reverse_parse_operator(self, css_text: str) -> str:
        """Convert CSS operators back to Excel."""
        # Reverse operator mappings
        result = css_text
        result = re.sub(r'\s*!=\s*', '<>', result)
        result = re.sub(r'\s*>=\s*', '>=', result)
        result = re.sub(r'\s*<=\s*', '<=', result)
        result = re.sub(r'\s*=\s*', '=', result)
        result = re.sub(r'\s*>\s*', '>', result)
        result = re.sub(r'\s*<\s*', '<', result)
        
        return result
    
    def reverse_parse_function(self, css_text: str) -> str:
        """Convert .SUM back to SUM."""
        return re.sub(r'\.([A-Z][A-Z0-9_]*)', r'\1', css_text)
    
    def reverse_parse_values(self, css_text: str) -> str:
        """Convert CSS properties back to Excel values."""
        # Convert content: "text" back to "text"
        result = re.sub(r'content:\s*("[^"]*")', r'\1', css_text)
        # Convert value: 123 back to 123
        result = re.sub(r'value:\s*([0-9.]+)', r'\1', result)
        return result

# End of file #
