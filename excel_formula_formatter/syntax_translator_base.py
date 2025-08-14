#!/usr/bin/env python3
"""
Base class for Excel formula syntax translators.
File: excel_formula_formatter/syntax_translator_base.py
"""

from abc import ABC, abstractmethod


class SyntaxTranslatorBase(ABC):
    """Base class for translating Excel tokens to target language syntax."""
    
    def __init__(self, indent_size: int = 4):
        self.indent_size = indent_size
    
    @abstractmethod
    def get_language_name(self) -> str:
        """Return the name of the target language."""
        pass
    
    @abstractmethod
    def get_file_extension(self) -> str:
        """Return the file extension for the target language."""
        pass
    
    @abstractmethod
    def format_header_comment(self) -> str:
        """Return the header comment for the formatted output."""
        pass
    
    @abstractmethod
    def format_section_comment(self, comment: str) -> str:
        """Format a section comment."""
        pass
    
    @abstractmethod
    def format_function_call(self, function_name: str) -> str:
        """Format a function name."""
        pass
    
    @abstractmethod
    def format_cell_reference(self, cell_ref: str) -> str:
        """Format a cell reference."""
        pass
    
    @abstractmethod
    def format_string_literal(self, string_val: str) -> str:
        """Format a string literal."""
        pass
    
    @abstractmethod
    def format_number(self, number_val: str) -> str:
        """Format a number."""
        pass
    
    @abstractmethod
    def format_operator(self, operator: str) -> str:
        """Format an operator, converting if necessary."""
        pass
    
    @abstractmethod
    def format_punctuation(self, punct: str) -> str:
        """Format punctuation."""
        pass
    
    @abstractmethod
    def reverse_parse_line(self, line: str) -> str:
        """Convert a line back to Excel syntax."""
        pass
    
    @abstractmethod 
    def reverse_parse_cell_reference(self, formatted_ref: str) -> str:
        """Convert formatted cell reference back to Excel."""
        pass
    
    @abstractmethod
    def reverse_parse_operator(self, formatted_op: str) -> str:
        """Convert formatted operator back to Excel."""
        pass
    
    def indent(self, depth: int) -> str:
        """Generate indentation string."""
        return " " * (depth * self.indent_size)
    
    def get_function_comment(self, function_name: str) -> str:
        """Get descriptive comment for common Excel functions."""
        comments = {
            'SUM': 'Sum values',
            'IF': 'Conditional logic', 
            'VLOOKUP': 'Vertical lookup',
            'HLOOKUP': 'Horizontal lookup',
            'INDEX': 'Index lookup',
            'MATCH': 'Find position',
            'SUMIF': 'Conditional sum',
            'SUMIFS': 'Multiple criteria sum',
            'COUNTIF': 'Conditional count',
            'COUNTIFS': 'Multiple criteria count',
            'CONCATENATE': 'Text concatenation',
            'TEXTJOIN': 'Join text with delimiter',
            'LET': 'Variable assignments',
            'LAMBDA': 'Function definition',
            'AND': 'Logical AND',
            'OR': 'Logical OR',
            'NOT': 'Logical NOT',
            'FILTER': 'Filter array',
            'SORT': 'Sort array',
            'UNIQUE': 'Unique values',
            'XLOOKUP': 'Extended lookup',
            'XMATCH': 'Extended match',
            'IFS': 'Multiple conditions'
        }
        return comments.get(function_name.upper(), '')

# End of file #
