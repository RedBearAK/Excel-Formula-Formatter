#!/usr/bin/env python3
"""
Modular Excel formula formatter supporting multiple target languages.
File: excel_formula_formatter/modular_excel_formatter.py
"""

import sys
import re

from excel_formula_formatter.excel_formula_patterns import (
    cell_ref_all_rgx, excel_functions_rgx, string_literal_rgx, 
    number_rgx, comment_line_rgx, inline_comment_rgx, 
    whitespace_newline_rgx, leading_trailing_space_rgx
)

from excel_formula_formatter.syntax_translator_base import SyntaxTranslatorBase
from excel_formula_formatter.javascript_translator import JavaScriptTranslator
from excel_formula_formatter.python_translator import PythonTranslator
from excel_formula_formatter.css_translator import CSSTranslator


class ModularExcelFormatter:
    """Excel formula formatter with pluggable syntax translators."""
    
    def __init__(self, translator: SyntaxTranslatorBase):
        self.translator = translator
        
    @classmethod
    def create_javascript_formatter(cls):
        """Create formatter with JavaScript translator."""
        return cls(JavaScriptTranslator())
    
    @classmethod  
    def create_python_formatter(cls):
        """Create formatter with Python translator."""
        return cls(PythonTranslator())
    
    @classmethod
    def create_css_formatter(cls):
        """Create formatter with CSS translator.""" 
        return cls(CSSTranslator())
    
    def fold_formula(self, formula: str) -> str:
        """Transform Excel formula using the configured translator."""
        if not formula or not formula.strip():
            return ""
            
        # Handle array formulas and regular formulas
        clean_formula = formula.strip()
        is_array_formula = False
        
        if clean_formula.startswith('{=') and clean_formula.endswith('}'):
            # Array formula: {=SUM(...)}
            is_array_formula = True
            clean_formula = clean_formula[2:-1]  # Remove {= and }
        elif clean_formula.startswith('='):
            # Regular formula: =SUM(...)
            clean_formula = clean_formula[1:]  # Remove =
            
        # Parse Excel tokens
        tokens = self._parse_excel_tokens(clean_formula)
        
        # Format using the translator
        formatted_lines = self._format_tokens_with_translator(tokens)
        
        # Add array formula markers if needed
        if is_array_formula:
            formatted_lines.insert(1, '{=')  # After header comment
            formatted_lines.append('}')
        
        return '\n'.join(formatted_lines)
    
    def unfold_formula(self, formatted_text: str) -> str:
        """Transform formatted text back to Excel formula."""
        if not formatted_text or not formatted_text.strip():
            return ""
            
        # Remove all comments first
        no_comments = comment_line_rgx.sub('', formatted_text)
        no_comments = inline_comment_rgx.sub('', no_comments)
        
        # Remove CSS-style comments if CSS translator
        if isinstance(self.translator, CSSTranslator):
            no_comments = re.sub(r'/\*.*?\*/', '', no_comments, flags=re.DOTALL)
        
        # Flatten to single line
        single_line = whitespace_newline_rgx.sub(' ', no_comments)
        single_line = leading_trailing_space_rgx.sub('', single_line)
        
        if not single_line:
            return ""
            
        # Use translator-specific reverse parsing
        excel_formula = self._reverse_parse_with_translator(single_line)
        
        # Add leading = if not present
        if not excel_formula.startswith('='):
            excel_formula = '=' + excel_formula
            
        return excel_formula
    
    def _parse_excel_tokens(self, formula: str) -> list:
        """Parse Excel formula into tokens with type information."""
        tokens = []
        i = 0
        length = len(formula)
        
        while i < length:
            # Skip whitespace
            if formula[i].isspace():
                i += 1
                continue
                
            # Check for string literals first (quoted text)
            if formula[i] == '"':
                end_quote = formula.find('"', i + 1)
                if end_quote == -1:
                    end_quote = length  # Unclosed quote, take rest
                token_text = formula[i:end_quote + 1]
                tokens.append(('string', token_text))
                i = end_quote + 1
                continue
            
            # Check for cell references (including ranges and sheet references)
            cell_match = cell_ref_all_rgx.match(formula, i)
            if cell_match:
                token_text = cell_match.group(0)
                tokens.append(('cell_ref', token_text))
                i = cell_match.end()
                continue
            
            # Check for two-character operators
            if i < length - 1:
                two_char = formula[i:i+2]
                if two_char in ['<>', '>=', '<=']:
                    tokens.append(('operator', two_char))
                    i += 2
                    continue
            
            # Check for single character operators and punctuation
            if formula[i] in '+-*/=<>(),[]:;!&%^':
                tokens.append(('punctuation', formula[i]))
                i += 1
                continue
            
            # Collect word/number/identifier
            start = i
            while i < length and not formula[i].isspace() and formula[i] not in '+-*/=<>(),[]:;!&%^"':
                i += 1
            
            if start < i:
                token_text = formula[start:i]
                token_type = self._classify_token(token_text)
                tokens.append((token_type, token_text))
                
        return tokens
    
    def _classify_token(self, token: str) -> str:
        """Classify a token by type."""
        if excel_functions_rgx.match(token):
            return 'function'
        elif cell_ref_all_rgx.match(token):
            return 'cell_ref'
        elif number_rgx.match(token):
            return 'number'
        elif token in ['<>', '>=', '<=', '==', '!=']:
            return 'operator'
        else:
            return 'identifier'
    
    def _format_tokens_with_translator(self, tokens: list) -> list:
        """Convert tokens using the configured translator."""
        lines = []
        current_line = ""
        depth = 0
        depth_stack = []  # Track whether each depth level is inline or multiline
        
        # Add header comment
        lines.append(self.translator.format_header_comment())
        
        i = 0
        while i < len(tokens):
            token_type, token_text = tokens[i]
            
            if token_type == 'punctuation' and token_text == '(':
                formatted_punct = self.translator.format_punctuation(token_text)
                current_line += formatted_punct
                
                # Look ahead to determine if this should be inline (simplified for modular)
                should_inline = self._should_keep_simple_inline(tokens, i)
                
                if should_inline:
                    # Keep this function call inline
                    depth_stack.append('inline')
                    i += 1
                    paren_depth = 1
                    inline_content = ""
                    
                    # Collect everything until matching closing parenthesis
                    while i < len(tokens) and paren_depth > 0:
                        t_type, t_text = tokens[i]
                        
                        if t_type == 'punctuation' and t_text == '(':
                            paren_depth += 1
                            inline_content += self.translator.format_punctuation(t_text)
                        elif t_type == 'punctuation' and t_text == ')':
                            paren_depth -= 1
                            if paren_depth > 0:
                                inline_content += self.translator.format_punctuation(t_text)
                        elif t_type == 'cell_ref':
                            inline_content += self.translator.format_cell_reference(t_text)
                        elif t_type == 'operator':
                            inline_content += self.translator.format_operator(t_text).strip()
                        elif t_type == 'string':
                            inline_content += self.translator.format_string_literal(t_text)
                        elif t_type == 'punctuation' and t_text == ',':
                            inline_content += ', '
                        elif t_type == 'number':
                            inline_content += self.translator.format_number(t_text)
                        else:
                            inline_content += t_text
                        
                        i += 1
                    
                    # Add the inline content and closing paren
                    current_line += inline_content + self.translator.format_punctuation(')')
                    depth_stack.pop()
                    i -= 1  # Back up one since the loop will increment
                else:
                    # Multi-line function call
                    depth_stack.append('multiline')
                    self._add_line_if_not_empty(lines, current_line, depth)
                    current_line = ""
                    depth += 1
                
            elif token_type == 'punctuation' and token_text == ')':
                # Check if this closing paren belongs to an inline or multiline context
                if depth_stack and depth_stack[-1] == 'inline':
                    # This shouldn't happen with our new logic
                    formatted_punct = self.translator.format_punctuation(token_text)
                    current_line += formatted_punct
                    depth_stack.pop()
                else:
                    # This is a multi-line function closing
                    self._add_line_if_not_empty(lines, current_line, depth)
                    depth = max(0, depth - 1)
                    
                    # Look ahead to see if there's a comma after this closing paren
                    next_token_is_comma = (i + 1 < len(tokens) and 
                                         tokens[i + 1][0] == 'punctuation' and 
                                         tokens[i + 1][1] == ',')
                    
                    formatted_punct = self.translator.format_punctuation(token_text)
                    if next_token_is_comma:
                        # Put the closing paren and comma together
                        comma_punct = self.translator.format_punctuation(',')
                        lines.append(self.translator.indent(depth) + formatted_punct + comma_punct.strip())
                        i += 1  # Skip the comma token since we've handled it
                    else:
                        lines.append(self.translator.indent(depth) + formatted_punct)
                    
                    current_line = ""
                    if depth_stack:
                        depth_stack.pop()
                
            elif token_type == 'punctuation' and token_text == ',':
                # Only handle commas that aren't already handled above
                formatted_punct = self.translator.format_punctuation(token_text)
                current_line += formatted_punct
                
                # Check if we're in a multiline context
                if depth > 0:
                    # Multi-line context - break line
                    self._add_line_if_not_empty(lines, current_line, depth)
                    current_line = ""
                else:
                    # Top level - just add space for some translators
                    if not formatted_punct.endswith('\n'):
                        current_line += ' '
                
            elif token_type == 'cell_ref':
                formatted_ref = self.translator.format_cell_reference(token_text)
                current_line += formatted_ref
                
            elif token_type == 'operator':
                formatted_op = self.translator.format_operator(token_text)
                current_line += formatted_op
                
            elif token_type == 'function':
                # Add comment for function sections only for complex functions
                if current_line.strip():
                    self._add_line_if_not_empty(lines, current_line, depth)
                    current_line = ""
                
                # Look ahead to see if this will be inline
                if not self._should_keep_simple_inline(tokens, i + 1):
                    comment = self.translator.get_function_comment(token_text)
                    if comment:
                        section_comment = self.translator.format_section_comment(comment)
                        lines.append(self.translator.indent(depth) + section_comment)
                
                formatted_func = self.translator.format_function_call(token_text)
                current_line += formatted_func
                
            elif token_type == 'string':
                formatted_str = self.translator.format_string_literal(token_text)
                current_line += formatted_str
                
            elif token_type == 'number':
                formatted_num = self.translator.format_number(token_text)
                current_line += formatted_num
                
            else:
                current_line += token_text
                
            i += 1
        
        # Add any remaining content
        self._add_line_if_not_empty(lines, current_line, depth)
        
        return lines
    
    def _should_keep_simple_inline(self, tokens: list, paren_index: int) -> bool:
        """Simplified inline detection for modular formatter."""
        if paren_index >= len(tokens) or tokens[paren_index][1] != '(':
            return False
        
        # Count arguments and complexity (simplified version)
        i = paren_index + 1
        paren_depth = 1
        arg_count = 0
        has_nested_functions = False
        
        while i < len(tokens) and paren_depth > 0 and i < paren_index + 10:  # Limit lookahead
            token_type, token_text = tokens[i]
            
            if token_type == 'punctuation' and token_text == '(':
                paren_depth += 1
                if paren_depth > 1:
                    has_nested_functions = True
            elif token_type == 'punctuation' and token_text == ')':
                paren_depth -= 1
            elif token_type == 'punctuation' and token_text == ',' and paren_depth == 1:
                arg_count += 1
            
            i += 1
        
        # Simple criteria: 1-2 args, no nesting, reasonable lookahead
        return arg_count <= 1 and not has_nested_functions
    
    def _add_line_if_not_empty(self, lines: list, line: str, depth: int):
        """Add line to list if it has content."""
        stripped = line.strip()
        if stripped:
            lines.append(self.translator.indent(depth) + stripped)
    
    def _reverse_parse_with_translator(self, formatted_text: str) -> str:
        """Use translator-specific reverse parsing."""
        result = formatted_text
        
        # Apply translator-specific reverse transformations
        result = self.translator.reverse_parse_cell_reference(result)
        result = self.translator.reverse_parse_operator(result)
        
        # Apply any additional translator-specific reverse parsing
        if hasattr(self.translator, 'reverse_parse_function'):
            result = self.translator.reverse_parse_function(result)
        if hasattr(self.translator, 'reverse_parse_values'):
            result = self.translator.reverse_parse_values(result)
        
        # Clean up spaces and syntax
        result = re.sub(r'\s*([+\-*/=<>!,()])\s*', r'\1', result)
        result = re.sub(r'\s*(<>|>=|<=)\s*', r'\1', result)
        result = re.sub(r'\s+', ' ', result)
        
        return result.strip()


def auto_format_with_translator(input_text: str, translator_name: str = 'javascript') -> str:
    """Automatically determine whether to fold or unfold using specified translator."""
    if not input_text or not input_text.strip():
        return ""
    
    # Create formatter with specified translator
    if translator_name == 'javascript':
        formatter = ModularExcelFormatter.create_javascript_formatter()
    elif translator_name == 'python':
        formatter = ModularExcelFormatter.create_python_formatter()
    elif translator_name == 'css':
        formatter = ModularExcelFormatter.create_css_formatter()
    else:
        formatter = ModularExcelFormatter.create_javascript_formatter()  # Default
    
    lines = input_text.strip().split('\n')
    
    # Single line - likely needs folding
    if len(lines) == 1:
        line = lines[0].strip()
        # Check if it looks like an Excel formula
        if line.startswith('=') or line.startswith('{='):
            return formatter.fold_formula(line)
        else:
            return input_text  # Not an Excel formula, return as-is
    
    # Multi-line - check if it's already folded or needs unfolding  
    else:
        # Look for folded indicators (comments, indentation)
        has_js_comments = any('//' in line for line in lines)
        has_py_comments = any('#' in line for line in lines)
        has_css_comments = any('/*' in line for line in lines)
        has_indentation = any(line.startswith('    ') or line.startswith('\t') for line in lines)
        has_excel_header = any('Excel Formula' in line for line in lines)
        
        if has_js_comments or has_py_comments or has_css_comments or has_indentation or has_excel_header:
            # Appears to be folded - unfold it
            return formatter.unfold_formula(input_text)
        else:
            # Multi-line but no folding indicators - manual line breaks?
            # Try to fold it as if it were a single line
            single_line = ' '.join(line.strip() for line in lines if line.strip())
            if single_line.startswith('=') or single_line.startswith('{='):
                return formatter.fold_formula(single_line)
            else:
                return input_text  # Can't determine format, return as-is


def main():
    """Main function for command line usage.""" 
    if len(sys.argv) < 2:
        print("Usage: python modular_excel_formatter.py <translator> [fold|unfold|auto]", file=sys.stderr)
        print("  translator: javascript, python, css", file=sys.stderr)
        print("  operation: fold, unfold, or auto (default)", file=sys.stderr)
        return 1
    
    translator_name = sys.argv[1].lower()
    operation = sys.argv[2].lower() if len(sys.argv) > 2 else 'auto'  # Default to auto
    
    if translator_name not in ['javascript', 'python', 'css']:
        print(f"Unknown translator: {translator_name}", file=sys.stderr)
        print("Available translators: javascript, python, css", file=sys.stderr)
        return 1
    
    if translator_name == 'javascript':
        formatter = ModularExcelFormatter.create_javascript_formatter()
    elif translator_name == 'python':
        formatter = ModularExcelFormatter.create_python_formatter()
    elif translator_name == 'css':
        formatter = ModularExcelFormatter.create_css_formatter()
    
    input_text = sys.stdin.read()
    if not input_text.strip():
        return 0
    
    try:
        if operation == 'fold':
            result = formatter.fold_formula(input_text)
        elif operation == 'unfold':
            result = formatter.unfold_formula(input_text)
        elif operation == 'auto':
            result = auto_format_with_translator(input_text, translator_name)
        else:
            print(f"Unknown operation: {operation}", file=sys.stderr)
            return 1
        
        print(result)
        return 0
        
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

# End of file #
