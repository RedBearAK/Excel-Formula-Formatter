#!/usr/bin/env python3
"""
Excel formula formatter with reversible JavaScript-like syntax transformation.
File: excel_formula_formatter.py
"""

import sys
import re

from excel_formula_formatter.excel_formula_patterns import (
    cell_ref_all_rgx, excel_functions_rgx, string_literal_rgx, 
    number_rgx, excel_not_equal_rgx, js_not_equal_rgx,
    comment_line_rgx, inline_comment_rgx, whitespace_newline_rgx,
    leading_trailing_space_rgx
)


class ExcelFormulaFormatter:
    def __init__(self, indent_size: int = 4):
        self.indent_size = indent_size
        
    def fold_formula(self, formula: str) -> str:
        """Transform Excel formula to JavaScript-like syntax with indentation."""
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
            
        # Parse and transform tokens
        tokens = self._parse_excel_tokens(clean_formula)
        formatted_lines = self._format_tokens_to_js(tokens)
        
        # Add array formula markers if needed
        if is_array_formula:
            formatted_lines.insert(1, '{=')  # After header comment
            formatted_lines.append('}')
        
        return '\n'.join(formatted_lines)
    
    def unfold_formula(self, js_like_text: str) -> str:
        """Transform JavaScript-like syntax back to Excel formula."""
        if not js_like_text or not js_like_text.strip():
            return ""
            
        # Check if this was an array formula
        lines = js_like_text.strip().split('\n')
        is_array_formula = False
        
        # Look for array formula markers
        if len(lines) >= 3 and lines[1].strip() == '{=' and lines[-1].strip() == '}':
            is_array_formula = True
            # Remove the {= and } lines
            js_like_text = '\n'.join(lines[2:-1])  # Skip header comment, {=, and }
        
        # Remove all comments first
        no_comments = comment_line_rgx.sub('', js_like_text)
        no_comments = inline_comment_rgx.sub('', no_comments)
        
        # Flatten to single line
        single_line = whitespace_newline_rgx.sub(' ', no_comments)
        single_line = leading_trailing_space_rgx.sub('', single_line)
        
        if not single_line:
            return ""
            
        # Transform back to Excel syntax
        excel_formula = self._transform_js_to_excel(single_line)
        
        # Add appropriate prefix
        if is_array_formula:
            if not excel_formula.startswith('{='):
                excel_formula = '{=' + excel_formula + '}'
        else:
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
            # Look ahead to see if this could be a cell reference
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
    
    def _format_tokens_to_js(self, tokens: list) -> list:
        """Convert tokens to JavaScript-like formatted lines with smart folding."""
        lines = []
        current_line = ""
        depth = 0
        depth_stack = []  # Track whether each depth level is inline or multiline
        
        # Add header comment
        lines.append("// Excel Formula (JavaScript syntax for highlighting)")
        
        i = 0
        while i < len(tokens):
            token_type, token_text = tokens[i]
            
            if token_type == 'punctuation' and token_text == '(':
                current_line += token_text
                
                # Look ahead to determine if this should be a simple one-liner
                if self._should_keep_inline(tokens, i):
                    # Keep this function call inline - don't change depth or break line
                    depth_stack.append('inline')
                    i += 1
                    paren_depth = 1
                    inline_content = ""
                    
                    # Collect everything until matching closing parenthesis
                    while i < len(tokens) and paren_depth > 0:
                        t_type, t_text = tokens[i]
                        
                        if t_type == 'punctuation' and t_text == '(':
                            paren_depth += 1
                            inline_content += t_text
                        elif t_type == 'punctuation' and t_text == ')':
                            paren_depth -= 1
                            if paren_depth > 0:  # Don't include the final closing paren yet
                                inline_content += t_text
                        elif t_type == 'cell_ref':
                            inline_content += f'"{t_text}"'
                        elif t_type == 'operator':
                            inline_content += self._format_operator_inline(t_text)
                        elif t_type == 'string':
                            inline_content += t_text
                        elif t_type == 'punctuation' and t_text == ',':
                            inline_content += ', '
                        else:
                            inline_content += t_text
                        
                        i += 1
                    
                    # Add the inline content and closing paren
                    current_line += inline_content + ')'
                    depth_stack.pop()  # Remove the inline marker
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
                    # This shouldn't happen with our new logic, but just in case
                    current_line += token_text
                    depth_stack.pop()
                else:
                    # This is a multi-line function closing
                    # Look ahead to see if there's a comma after this closing paren
                    next_token_is_comma = (i + 1 < len(tokens) and 
                                         tokens[i + 1][0] == 'punctuation' and 
                                         tokens[i + 1][1] == ',')
                    
                    # Handle any content on current line first
                    if current_line.strip():
                        lines.append(self._indent(depth) + current_line.strip())
                    
                    depth = max(0, depth - 1)
                    
                    # Add closing paren with or without comma
                    if next_token_is_comma:
                        lines.append(self._indent(depth) + token_text + ',')
                        i += 1  # Skip the comma token since we've handled it
                    else:
                        lines.append(self._indent(depth) + token_text)
                    
                    current_line = ""
                    if depth_stack:
                        depth_stack.pop()
                
            elif token_type == 'punctuation' and token_text == ',':
                # Only handle commas that aren't already handled above
                current_line += token_text
                
                # Check if we're in a multiline context (depth > 0 and not all inline)
                if depth > 0:
                    # Multi-line context - break line
                    self._add_line_if_not_empty(lines, current_line, depth)
                    current_line = ""
                else:
                    # Top level or inline context - just add space
                    current_line += ' '
                
            elif token_type == 'cell_ref':
                # Quote cell references for string highlighting
                current_line += f'"{token_text}"'
                
            elif token_type == 'operator' and token_text == '<>':
                # Convert Excel <> to JavaScript !=
                current_line += ' != '
                
            elif token_type == 'operator':
                current_line += f' {token_text} '
                
            elif token_type == 'function':
                # Add comment for function sections only for complex functions
                if current_line.strip():
                    self._add_line_if_not_empty(lines, current_line, depth)
                    current_line = ""
                
                # Look ahead to see if this will be inline or multiline
                if not self._should_keep_inline(tokens, i + 1):  # +1 because next token should be '('
                    comment = self._get_function_comment(token_text)
                    if comment:
                        lines.append(self._indent(depth) + f"// {comment}")
                
                current_line += token_text
                
            else:
                current_line += token_text
                
            i += 1
        
        # Add any remaining content
        self._add_line_if_not_empty(lines, current_line, depth)
        
        return lines
    
    def _should_keep_inline(self, tokens: list, paren_index: int) -> bool:
        """Determine if a function call should be kept inline."""
        if paren_index >= len(tokens) or tokens[paren_index][1] != '(':
            return False
        
        # Look at the function name before the parenthesis
        function_name = ""
        if paren_index > 0:
            prev_token = tokens[paren_index - 1]
            if prev_token[0] == 'function':
                function_name = prev_token[1].upper()
        
        # Count arguments and complexity within parentheses
        i = paren_index + 1
        paren_depth = 1
        arg_count = 0
        has_nested_functions = False
        total_length = 0
        
        while i < len(tokens) and paren_depth > 0:
            token_type, token_text = tokens[i]
            
            if token_type == 'punctuation' and token_text == '(':
                paren_depth += 1
                if paren_depth > 1:  # Nested function call
                    has_nested_functions = True
            elif token_type == 'punctuation' and token_text == ')':
                paren_depth -= 1
            elif token_type == 'punctuation' and token_text == ',' and paren_depth == 1:
                arg_count += 1
            
            total_length += len(token_text)
            i += 1
        
        arg_count += 1  # Add 1 for the last argument (no comma after it)
        
        # Criteria for keeping inline:
        # 1. Simple functions (LEN, TRIM, etc.) with 1 argument
        simple_functions = {'LEN', 'TRIM', 'UPPER', 'LOWER', 'ABS', 'INT', 'ROUND'}
        if function_name in simple_functions and arg_count == 1 and not has_nested_functions:
            return True
        
        # 2. Any function with 1 simple argument and reasonable length
        if arg_count == 1 and not has_nested_functions and total_length < 30:
            return True
        
        # 3. Functions with 2 simple arguments and short length
        if arg_count == 2 and not has_nested_functions and total_length < 20:
            return True
        
        return False
    
    def _format_operator_inline(self, operator: str) -> str:
        """Format operator for inline use (no extra spaces)."""
        if operator == '<>':
            return ' != '
        elif operator in ['>=', '<=', '>', '<', '=']:
            return f' {operator} '
        else:
            return operator
    
    def _add_line_if_not_empty(self, lines: list, line: str, depth: int):
        """Add line to list if it has content."""
        stripped = line.strip()
        if stripped:
            lines.append(self._indent(depth) + stripped)
    
    def _indent(self, depth: int) -> str:
        """Generate indentation string."""
        return " " * (depth * self.indent_size)
    
    def _get_function_comment(self, function_name: str) -> str:
        """Get descriptive comment for function."""
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
            'LET': 'Variable assignments',
            'AND': 'Logical AND',
            'OR': 'Logical OR',
            'NOT': 'Logical NOT',
            'IFS': 'Multiple conditions'
        }
        return comments.get(function_name.upper(), '')
    
    def _transform_js_to_excel(self, js_text: str) -> str:
        """Transform JavaScript-like syntax back to Excel."""
        # Remove quotes from cell references - need to be more careful here
        # First, find all quoted strings that match cell reference patterns
        def unquote_cell_ref(match):
            quoted_text = match.group(0)
            # Remove outer quotes
            inner_text = quoted_text[1:-1]
            # Check if the inner text is actually a cell reference
            if cell_ref_all_rgx.match(inner_text):
                return inner_text
            else:
                return quoted_text  # Keep quotes if not a cell reference
        
        # Apply unquoting to quoted cell references
        result = re.sub(r'"[^"]*"', unquote_cell_ref, js_text)
        
        # Convert != back to <>
        result = js_not_equal_rgx.sub('<>', result)
        
        # Clean up extra spaces around operators - but be more careful
        # Don't remove spaces inside string literals
        
        # First, temporarily replace string literals with placeholders
        string_parts = []
        def replace_string(match):
            string_parts.append(match.group(0))
            return f"__STRING_{len(string_parts)-1}__"
        
        result = re.sub(r'"[^"]*"', replace_string, result)
        
        # Now clean up operators (avoiding the string placeholders)
        result = re.sub(r'\s*([+\-*/=<>!,()])\s*', r'\1', result)
        result = re.sub(r'\s*(<>|>=|<=)\s*', r'\1', result)
        
        # Restore string literals
        for i, string_literal in enumerate(string_parts):
            result = result.replace(f"__STRING_{i}__", string_literal)
        
        # Clean up any remaining multiple spaces
        result = re.sub(r'\s+', ' ', result)
        
        return result.strip()


def auto_format(input_text: str) -> str:
    """Automatically determine whether to fold or unfold based on input format."""
    if not input_text or not input_text.strip():
        return ""
    
    lines = input_text.strip().split('\n')
    
    # Single line - likely needs folding
    if len(lines) == 1:
        line = lines[0].strip()
        # Check if it looks like an Excel formula
        if line.startswith('=') or line.startswith('{='):
            formatter = ExcelFormulaFormatter()
            return formatter.fold_formula(line)
        else:
            return input_text  # Not an Excel formula, return as-is
    
    # Multi-line - check if it's already folded or needs unfolding
    else:
        # Look for folded indicators (comments, indentation)
        has_comments = any('//' in line for line in lines)
        has_indentation = any(line.startswith('    ') or line.startswith('\t') for line in lines)
        has_excel_header = any('Excel Formula' in line for line in lines)
        
        if has_comments or has_indentation or has_excel_header:
            # Appears to be folded - unfold it
            formatter = ExcelFormulaFormatter()
            return formatter.unfold_formula(input_text)
        else:
            # Multi-line but no folding indicators - manual line breaks?
            # Try to fold it as if it were a single line
            single_line = ' '.join(line.strip() for line in lines if line.strip())
            if single_line.startswith('=') or single_line.startswith('{='):
                formatter = ExcelFormulaFormatter()
                return formatter.fold_formula(single_line)
            else:
                return input_text  # Can't determine format, return as-is


def main():
    """Main function for command line usage."""
    if len(sys.argv) > 1:
        operation = sys.argv[1].lower()
        if operation not in ['fold', 'unfold', 'auto']:
            print("Usage: python excel_formula_formatter.py [fold|unfold|auto]", file=sys.stderr)
            print("  fold: Excel formula → JavaScript-like syntax", file=sys.stderr)
            print("  unfold: JavaScript-like syntax → Excel formula", file=sys.stderr)
            print("  auto: Automatically detect format and toggle", file=sys.stderr)
            return 1
    else:
        operation = 'auto'  # Default to auto-detection
    
    input_text = sys.stdin.read()
    if not input_text.strip():
        return 0
    
    formatter = ExcelFormulaFormatter(indent_size=4)
    
    try:
        if operation == 'fold':
            result = formatter.fold_formula(input_text)
        elif operation == 'unfold':
            result = formatter.unfold_formula(input_text)
        elif operation == 'auto':
            result = auto_format(input_text)
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
