#!/usr/bin/env python3
"""
JavaScript syntax translator for Excel formulas.
File: excel_formula_formatter/javascript_translator.py

Cell references are marked with backticks (`A1`) rather than double
quotes. Backticks read as template literals to JavaScript highlighters,
so cell refs still highlight as strings, but the marker cannot collide
with genuine Excel string literals such as "A1" - the ambiguity that
made the old double-quote scheme corrupt formulas on unfold.
"""

from excel_formula_formatter.syntax_translator_base import SyntaxTranslatorBase
from excel_formula_formatter.excel_formula_patterns import (
    backtick_span_rgx,
    js_not_equal_rgx,
)


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
        # Backtick-quote cell references for string-like highlighting
        # without colliding with genuine "..." string literals
        return f'`{cell_ref}`'

    def format_string_literal(self, string_val: str) -> str:
        return string_val  # Already quoted

    def format_number(self, number_val: str) -> str:
        return number_val

    def format_operator(self, operator: str) -> str:
        # Convert Excel <> to JavaScript !=
        if operator == '<>':
            return ' != '
        return f' {operator} '

    def format_punctuation(self, punct: str) -> str:
        # Add spacing around function parentheses for readability
        if punct == '(':
            return '( '
        if punct == ')':
            return ' )'
        return punct

    def reverse_parse_line(self, line: str) -> str:
        """No line-level work needed: comment removal is handled
        centrally (string-aware) before reverse parsing begins."""
        return line.strip()

    def reverse_parse_cell_reference(self, js_text: str) -> str:
        """Strip backtick markers, restoring bare cell references."""
        def unquote_cell_ref(match):
            return match.group(0)[1:-1]

        return backtick_span_rgx.sub(unquote_cell_ref, js_text)

    def reverse_parse_operator(self, js_text: str) -> str:
        """Convert != back to <>."""
        return js_not_equal_rgx.sub('<>', js_text)

# End of file #
