#!/usr/bin/env python3
"""
Compact Excel syntax translator for maximum space efficiency.
File: excel_formula_formatter/compact_excel_translator.py
"""

from excel_formula_formatter.syntax_translator_base import SyntaxTranslatorBase


class CompactExcelTranslator(SyntaxTranslatorBase):
    """Compact Excel translator - pure Excel syntax with minimal spacing for maximum efficiency."""
    
    def get_language_name(self) -> str:
        return "Excel (Compact)"
    
    def get_file_extension(self) -> str:
        return ".txt"
    
    def format_header_comment(self) -> str:
        return ""  # NO header comment in compact mode
    
    def format_section_comment(self, comment: str) -> str:
        return ""  # NO section comments in compact mode
    
    def format_function_call(self, function_name: str) -> str:
        return function_name  # Keep original case and format
    
    def format_cell_reference(self, cell_ref: str) -> str:
        return cell_ref  # Keep unquoted
    
    def format_string_literal(self, string_val: str) -> str:
        return string_val  # Already quoted - preserve internal spaces
    
    def format_number(self, number_val: str) -> str:
        return number_val
    
    def format_operator(self, operator: str) -> str:
        # NO spaces around operators for maximum compactness
        return operator
    
    def format_punctuation(self, punct: str) -> str:
        # NO spacing around any punctuation for maximum compactness
        return punct
    
    def format_comma(self) -> str:
        """Format comma with no trailing space for compact mode."""
        return ","
    
    def reverse_parse_line(self, line: str) -> str:
        """Compact mode should not have comments, but clean line just in case."""
        return line.strip()
    
    def reverse_parse_cell_reference(self, text: str) -> str:
        """No conversion needed for compact format."""
        return text
    
    def reverse_parse_operator(self, text: str) -> str:
        """No conversion needed for compact format."""
        return text

# End of file #
