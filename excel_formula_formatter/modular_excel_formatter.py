#!/usr/bin/env python3
"""
Modular Excel formula formatter supporting multiple target languages.
Complete rewrite with true function isolation using recursive descent parsing.
File: excel_formula_formatter/modular_excel_formatter.py
"""

import sys
import re

from excel_formula_formatter.excel_formula_patterns import (
    cell_ref_all_rgx, excel_functions_rgx, string_literal_rgx, 
    number_rgx, comment_line_rgx, inline_comment_rgx, 
    whitespace_newline_rgx, leading_trailing_space_rgx,
    line_comment_removal_rgx, space_cleanup_operators_rgx,
    space_cleanup_multi_char_rgx, space_cleanup_whitespace_rgx
)

from excel_formula_formatter.syntax_translator_base import SyntaxTranslatorBase
from excel_formula_formatter.javascript_translator import JavaScriptTranslator


class PlainTranslator(SyntaxTranslatorBase):
    """Plain Excel translator that preserves original syntax with smart indenting only."""
    
    def get_language_name(self) -> str:
        return "Excel"
    
    def get_file_extension(self) -> str:
        return ".txt"
    
    def format_header_comment(self) -> str:
        return "// Excel Formula (plain Excel syntax with smart indenting)"
    
    def format_section_comment(self, comment: str) -> str:
        return f"// {comment}"
    
    def format_function_call(self, function_name: str) -> str:
        return function_name  # Keep original case and format
    
    def format_cell_reference(self, cell_ref: str) -> str:
        return cell_ref  # Keep unquoted
    
    def format_string_literal(self, string_val: str) -> str:
        return string_val  # Already quoted
    
    def format_number(self, number_val: str) -> str:
        return number_val
    
    def format_operator(self, operator: str) -> str:
        # Keep Excel operators as-is, just add spacing
        return f' {operator} '
    
    def format_punctuation(self, punct: str) -> str:
        return punct
    
    def reverse_parse_line(self, line: str) -> str:
        """Remove comments and return clean line."""
        # Remove comments using the imported pattern
        line = line_comment_removal_rgx.sub('', line)
        return line.strip()
    
    def reverse_parse_cell_reference(self, text: str) -> str:
        """No conversion needed for plain format."""
        return text
    
    def reverse_parse_operator(self, text: str) -> str:
        """No conversion needed for plain format."""
        return text


class ModularExcelFormatter:
    """Excel formula formatter with pluggable syntax translators and true function isolation."""
    
    def __init__(self, translator: SyntaxTranslatorBase):
        self.translator = translator
        
    @classmethod
    def create_javascript_formatter(cls):
        """Create formatter with JavaScript translator."""
        return cls(JavaScriptTranslator())
    
    @classmethod  
    def create_plain_formatter(cls):
        """Create formatter with Plain Excel translator (smart indenting only)."""
        return cls(PlainTranslator())
    
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
        
        # Format using the translator with true isolation
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
            
            # Check for single character operators (separate from punctuation)
            if formula[i] in '+-*/=<>&':
                tokens.append(('operator', formula[i]))
                i += 1
                continue
            
            # Check for punctuation
            if formula[i] in '(),[]:;!%^':
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
        """Convert tokens using the configured translator with TRUE function isolation."""
        lines = []
        lines.append(self.translator.format_header_comment())
        
        # Process tokens with isolated function handling
        processed_lines = self._process_token_sequence(tokens, base_depth=0)
        lines.extend(processed_lines)
        
        return lines

    def _process_token_sequence(self, tokens: list, base_depth: int) -> list:
        """Process a sequence of tokens with proper function isolation."""
        lines = []
        current_line = ""
        i = 0
        
        # Don't filter tokens here - let _tokens_to_string handle it
        while i < len(tokens):
            token_type, token_text = tokens[i]
            
            if token_type == 'function':
                # Handle function calls with complete isolation
                func_name = token_text.upper()
                
                # Find the function's argument tokens (between parentheses)
                if i + 1 < len(tokens) and tokens[i + 1][1] == '(':
                    arg_tokens, end_index = self._extract_function_arguments(tokens, i + 1)
                    
                    # Process this function in complete isolation
                    if func_name in ['IFS', 'SWITCH']:
                        func_lines = self._process_ifs_function(token_text, arg_tokens, base_depth)
                    elif func_name == 'LET':
                        func_lines = self._process_let_function(token_text, arg_tokens, base_depth)
                    elif func_name in ['AND', 'OR']:
                        func_lines = self._process_logical_function(token_text, arg_tokens, base_depth)
                    else:
                        func_lines = self._process_generic_function(token_text, arg_tokens, base_depth)
                    
                    # Add the function content
                    if current_line.strip():
                        lines.append(self.translator.indent(base_depth) + current_line.strip())
                        current_line = ""
                    
                    lines.extend(func_lines)
                    i = end_index - 1  # Point to position that will be incremented
                else:
                    # Function without parentheses - treat as identifier
                    current_line += self.translator.format_function_call(token_text)
            
            elif token_type == 'cell_ref':
                current_line += self.translator.format_cell_reference(token_text)
            elif token_type == 'string':
                current_line += self.translator.format_string_literal(token_text)
            elif token_type == 'number':
                current_line += self.translator.format_number(token_text)
            elif token_type == 'operator':
                current_line += self.translator.format_operator(token_text)
            elif token_type == 'punctuation' and token_text == ',':
                # Top-level comma - just add space
                current_line += self.translator.format_punctuation(token_text) + " "
            elif token_type == 'punctuation':
                current_line += self.translator.format_punctuation(token_text)
            else:
                current_line += token_text
            
            i += 1
        
        # Add any remaining content
        if current_line.strip():
            lines.append(self.translator.indent(base_depth) + current_line.strip())
        
        return lines

    def _extract_function_arguments(self, tokens: list, paren_start: int) -> tuple:
        """Extract tokens between matching parentheses."""
        if tokens[paren_start][1] != '(':
            return [], paren_start
        
        arg_tokens = []
        depth = 1
        i = paren_start + 1
        
        while i < len(tokens) and depth > 0:
            token_type, token_text = tokens[i]
            
            if token_text == '(':
                depth += 1
            elif token_text == ')':
                depth -= 1
                
            if depth > 0:  # Don't include the final closing paren
                arg_tokens.append((token_type, token_text))
            
            i += 1
        
        return arg_tokens, i

    def _process_ifs_function(self, func_name: str, arg_tokens: list, base_depth: int) -> list:
        """Process IFS/SWITCH function in complete isolation."""
        lines = []
        
        # Add function comment only if we're not already inside an IFS/SWITCH
        comment = self.translator.get_function_comment(func_name)
        if comment:
            lines.append(self.translator.indent(base_depth) + self.translator.format_section_comment(comment))
        
        # Function header
        lines.append(self.translator.indent(base_depth) + self.translator.format_function_call(func_name) + self.translator.format_punctuation('('))
        
        # Split arguments by top-level commas
        argument_groups = self._split_by_top_level_commas(arg_tokens)
        
        # Add initial separator only if we have arguments
        if argument_groups:
            lines.append(self.translator.indent(base_depth + 1) + self.translator.format_section_comment("── CASE/RESULT PAIR ──"))
        
        for arg_index, arg_group in enumerate(argument_groups):
            # Add separator before each condition (even arguments > 1, i.e., arguments 2, 4, 6...)
            if arg_index > 1 and arg_index % 2 == 0:
                lines.append("")  # Blank line
                lines.append(self.translator.indent(base_depth + 1) + self.translator.format_section_comment("── CASE/RESULT PAIR ──"))
            
            # Process this argument group
            arg_lines = self._process_token_sequence(arg_group, base_depth + 1)
            
            # Add comma if not last argument
            if arg_index < len(argument_groups) - 1:
                if arg_lines:
                    arg_lines[-1] += self.translator.format_punctuation(',')
                else:
                    lines.append(self.translator.indent(base_depth + 1) + self.translator.format_punctuation(','))
            
            lines.extend(arg_lines)
        
        # Closing paren
        lines.append(self.translator.indent(base_depth) + self.translator.format_punctuation(')'))
        
        return lines

    def _process_let_function(self, func_name: str, arg_tokens: list, base_depth: int) -> list:
        """Process LET function in complete isolation."""
        lines = []
        
        # Add function comment
        comment = self.translator.get_function_comment(func_name)
        if comment:
            lines.append(self.translator.indent(base_depth) + self.translator.format_section_comment(comment))
        
        # Function header
        lines.append(self.translator.indent(base_depth) + self.translator.format_function_call(func_name) + self.translator.format_punctuation('('))
        
        # Split arguments by top-level commas
        argument_groups = self._split_by_top_level_commas(arg_tokens)
        
        i = 0
        while i < len(argument_groups):
            # LET pairs: keep variable name and value on same line
            if i % 2 == 0 and i + 1 < len(argument_groups):
                # Process variable name (should be simple identifier)
                var_name = self._tokens_to_string(argument_groups[i]).strip()
                
                # Process value (could be complex expression)
                value_str = self._tokens_to_string(argument_groups[i + 1]).strip()
                
                # Combine on same line: variable, value,
                combined_line = self.translator.indent(base_depth + 1) + var_name + self.translator.format_punctuation(',') + " " + value_str
                
                # Add comma if not the last pair (check if this isn't the final expression)
                if i + 2 < len(argument_groups):
                    combined_line += self.translator.format_punctuation(',')
                
                lines.append(combined_line)
                i += 2  # Skip both variable and value
            else:
                # Final expression (not a pair) - should be the last argument
                final_expr_lines = self._process_token_sequence(argument_groups[i], base_depth + 1)
                lines.extend(final_expr_lines)
                i += 1
        
        # Closing paren
        lines.append(self.translator.indent(base_depth) + self.translator.format_punctuation(')'))
        
        return lines

    def _process_logical_function(self, func_name: str, arg_tokens: list, base_depth: int) -> list:
        """Process AND/OR functions with natural wrapping."""
        lines = []
        
        # Split arguments by top-level commas
        argument_groups = self._split_by_top_level_commas(arg_tokens)
        
        # Filter out any empty argument groups
        argument_groups = [group for group in argument_groups if group]
        
        if not argument_groups:
            # Empty function call
            func_str = self.translator.format_function_call(func_name) + self.translator.format_punctuation('(') + self.translator.format_punctuation(')')
            lines.append(self.translator.indent(base_depth) + func_str)
            return lines
        
        # Convert argument groups to strings
        arg_strings = []
        for group in argument_groups:
            arg_str = self._tokens_to_string(group).strip()
            if arg_str:
                arg_strings.append(arg_str)
        
        if not arg_strings:
            # No valid arguments
            func_str = self.translator.format_function_call(func_name) + self.translator.format_punctuation('(') + self.translator.format_punctuation(')')
            lines.append(self.translator.indent(base_depth) + func_str)
            return lines
        
        # Try to fit everything on one line first
        single_line_content = ", ".join(arg_strings)
        single_line = (self.translator.format_function_call(func_name) + 
                      self.translator.format_punctuation('(') + "  " + 
                      single_line_content + "  " + 
                      self.translator.format_punctuation(')'))
        
        # Check if single line fits (account for indentation)
        indented_single_line = self.translator.indent(base_depth) + single_line
        if len(indented_single_line) <= 79:
            lines.append(indented_single_line)
            return lines
        
        # Multi-line with natural wrapping - pack as many arguments per line as fit
        lines.append(self.translator.indent(base_depth) + 
                    self.translator.format_function_call(func_name) + 
                    self.translator.format_punctuation('(') + "  " + arg_strings[0])
        
        # Natural wrapping for remaining arguments
        current_line = self.translator.indent(base_depth + 1)
        line_arg_count = 0
        
        for arg_index in range(1, len(arg_strings)):
            arg_str = arg_strings[arg_index]
            
            # Check if adding this argument would make the line too long (account for indentation)
            test_line = current_line
            if line_arg_count > 0:
                test_line += ", "
            test_line += arg_str
            
            if len(test_line) <= 75 and line_arg_count < 3:  # Max 3 args per continuation line
                # Add to current line
                if line_arg_count > 0:
                    current_line += ", "
                current_line += arg_str
                line_arg_count += 1
            else:
                # Finish current line and start new one
                if line_arg_count > 0:
                    current_line += ","
                    lines.append(current_line)
                current_line = self.translator.indent(base_depth + 1) + arg_str
                line_arg_count = 1
        
        # Finish the last line and add closing paren
        if line_arg_count > 0:
            current_line += "  " + self.translator.format_punctuation(')')
            lines.append(current_line)
        else:
            # This shouldn't happen, but just in case
            lines.append(self.translator.indent(base_depth) + "  " + self.translator.format_punctuation(')'))
        
        return lines

    def _process_generic_function(self, func_name: str, arg_tokens: list, base_depth: int) -> list:
        """Process generic functions."""
        lines = []
        
        # Simple inline check
        total_length = len(self._tokens_to_string(arg_tokens))
        
        if total_length < 30:
            # Keep inline
            func_str = self.translator.format_function_call(func_name) + self.translator.format_punctuation('(')
            func_str += self._tokens_to_string(arg_tokens)
            func_str += self.translator.format_punctuation(')')
            lines.append(self.translator.indent(base_depth) + func_str)
        else:
            # Multi-line
            lines.append(self.translator.indent(base_depth) + self.translator.format_function_call(func_name) + self.translator.format_punctuation('('))
            
            # Process arguments
            argument_groups = self._split_by_top_level_commas(arg_tokens)
            for arg_index, arg_group in enumerate(argument_groups):
                arg_lines = self._process_token_sequence(arg_group, base_depth + 1)
                
                if arg_index < len(argument_groups) - 1:
                    if arg_lines:
                        arg_lines[-1] += self.translator.format_punctuation(',')
                
                lines.extend(arg_lines)
            
            lines.append(self.translator.indent(base_depth) + self.translator.format_punctuation(')'))
        
        return lines

    def _split_by_top_level_commas(self, tokens: list) -> list:
        """Split tokens by commas that are at the top level (depth 0)."""
        if not tokens:
            return []
            
        groups = []
        current_group = []
        depth = 0
        
        for token_type, token_text in tokens:
            if token_text == '(':
                depth += 1
                current_group.append((token_type, token_text))
            elif token_text == ')':
                depth -= 1
                current_group.append((token_type, token_text))
            elif token_text == ',' and depth == 0:
                # Top-level comma - start new group only if current group has content
                if current_group:
                    groups.append(current_group)
                    current_group = []
                # Don't add the comma to either group - it's just a separator
            else:
                # Add all other tokens including whitespace (we'll filter later if needed)
                current_group.append((token_type, token_text))
        
        # Add the last group only if it has content
        if current_group:
            groups.append(current_group)
        
        # Filter out completely empty groups, but preserve groups with meaningful content
        filtered_groups = []
        for group in groups:
            # Check if group has any non-whitespace content
            has_content = any(token_text.strip() for token_type, token_text in group)
            if has_content:
                filtered_groups.append(group)
        
        return filtered_groups

    def _tokens_to_string(self, tokens: list) -> str:
        """Convert token sequence to formatted string."""
        result = ""
        for i, (token_type, token_text) in enumerate(tokens):
            if token_type == 'cell_ref':
                result += self.translator.format_cell_reference(token_text)
            elif token_type == 'string':
                result += self.translator.format_string_literal(token_text)
            elif token_type == 'number':
                result += self.translator.format_number(token_text)
            elif token_type == 'operator':
                result += self.translator.format_operator(token_text)
            elif token_type == 'function':
                result += self.translator.format_function_call(token_text)
            elif token_type == 'punctuation':
                if token_text == ',':
                    result += ', '  # Always add space after comma
                else:
                    result += self.translator.format_punctuation(token_text)
            else:
                result += token_text
        
        return result.strip()
    
    def _reverse_parse_with_translator(self, formatted_text: str) -> str:
        """Use translator-specific reverse parsing."""
        result = formatted_text
        
        # Apply translator-specific reverse transformations
        result = self.translator.reverse_parse_cell_reference(result)
        result = self.translator.reverse_parse_operator(result)
        
        # Apply line-level reverse parsing if available
        if hasattr(self.translator, 'reverse_parse_line'):
            result = self.translator.reverse_parse_line(result)
        
        # Apply any additional translator-specific reverse parsing
        if hasattr(self.translator, 'reverse_parse_function'):
            result = self.translator.reverse_parse_function(result)
        if hasattr(self.translator, 'reverse_parse_values'):
            result = self.translator.reverse_parse_values(result)
        
        # Clean up spaces and syntax
        result = space_cleanup_operators_rgx.sub(r'\1', result)
        result = space_cleanup_multi_char_rgx.sub(r'\1', result)
        result = space_cleanup_whitespace_rgx.sub(' ', result)
        
        return result.strip()


def detect_current_mode(text: str) -> str:
    """Detect what formatter mode the text is currently in."""
    if not text or not text.strip():
        return 'unknown'
    
    lines = text.strip().split('\n')
    
    # Single line is likely unfolded (plain Excel)
    if len(lines) == 1:
        return 'plain'
    
    # Check for mode indicators
    text_content = '\n'.join(lines)
    
    if '//' in text_content and 'JavaScript syntax' in text_content:
        return 'javascript'
    elif '//' in text_content and 'plain Excel syntax' in text_content:
        return 'plain'  
    elif any(line.startswith('    ') or line.startswith('\t') for line in lines):
        # Has indentation but no clear mode indicators - assume javascript as most common
        return 'javascript'
    
    return 'plain'  # Default to plain if uncertain


def safe_mode_switch(text: str, current_mode: str, target_mode: str, should_refold: bool = False) -> str:
    """Safely switch between formatter modes by unfolding first."""
    if not text or not text.strip():
        return ""
    
    if current_mode == target_mode:
        return text  # No change needed
    
    # Step 1: Unfold using current mode formatter
    if current_mode == 'javascript':
        current_formatter = ModularExcelFormatter.create_javascript_formatter()
    elif current_mode == 'plain':
        current_formatter = ModularExcelFormatter.create_plain_formatter()
    else:
        # Unknown mode - try to auto-detect and unfold
        return auto_format_with_translator(text, target_mode)
    
    # Unfold to get clean Excel formula
    unfolded = current_formatter.unfold_formula(text)
    
    # Step 2: If requested, fold using target mode formatter
    if should_refold:
        if target_mode == 'javascript':
            target_formatter = ModularExcelFormatter.create_javascript_formatter()
        elif target_mode == 'plain':
            target_formatter = ModularExcelFormatter.create_plain_formatter()
        else:
            # Invalid target mode, return unfolded
            return unfolded
        
        return target_formatter.fold_formula(unfolded)
    else:
        # Just return unfolded - user can re-fold manually if desired
        return unfolded


def auto_format_with_translator(input_text: str, translator_name: str = 'javascript') -> str:
    """Automatically determine whether to fold or unfold using specified translator."""
    if not input_text or not input_text.strip():
        return ""
    
    # Create formatter with specified translator
    if translator_name == 'javascript':
        formatter = ModularExcelFormatter.create_javascript_formatter()
    elif translator_name == 'plain':
        formatter = ModularExcelFormatter.create_plain_formatter()
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
        has_js_comments = any('//' in line and 'JavaScript syntax' in line for line in lines)
        has_plain_comments = any('//' in line and 'plain Excel syntax' in line for line in lines)
        has_indentation = any(line.startswith('    ') or line.startswith('\t') for line in lines)
        
        if has_js_comments or has_plain_comments or has_indentation:
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
        print("  translator: javascript, plain", file=sys.stderr)
        print("  operation: fold, unfold, or auto (default)", file=sys.stderr)
        print("  javascript: JavaScript-like syntax for highlighting", file=sys.stderr)
        print("  plain: Pure Excel syntax with smart indenting", file=sys.stderr)
        return 1
    
    translator_name = sys.argv[1].lower()
    operation = sys.argv[2].lower() if len(sys.argv) > 2 else 'auto'  # Default to auto
    
    if translator_name not in ['javascript', 'plain']:
        print(f"Unknown translator: {translator_name}", file=sys.stderr)
        print("Available translators: javascript, plain", file=sys.stderr)
        return 1
    
    if translator_name == 'javascript':
        formatter = ModularExcelFormatter.create_javascript_formatter()
    elif translator_name == 'plain':
        formatter = ModularExcelFormatter.create_plain_formatter()
    
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
