#!/usr/bin/env python3
"""
Modular Excel formula formatter supporting four distinct target modes.
Complete rewrite with four modes: 
JavaScript (j), 
Annotated Excel (a), 
Plain Excel (p), 
Compact Excel (c).
File: excel_formula_formatter/modular_excel_formatter.py
"""

import sys
import re

from excel_formula_formatter.formula_text_guards import (
    find_opaque_end,
    protect_opaque_spans,
    restore_opaque_spans,
    scan_string_literal,
)
from excel_formula_formatter.excel_formula_patterns import (
    cell_ref_all_rgx,
    number_rgx,
    whitespace_newline_rgx,
    leading_trailing_space_rgx,
    paren_leading_space_rgx,
    paren_trailing_space_rgx,
    multiple_spaces_rgx,
    comma_spacing_rgx,
    operator_spacing_rgx,
    multi_char_operator_spacing_rgx,
    comment_line_detection_rgx
)

# LET name/value pairs longer than this (in characters) get their
# value expression expanded across multiple lines when folding
let_value_wrap_threshold = 100

from excel_formula_formatter.syntax_translator_base import SyntaxTranslatorBase
from excel_formula_formatter.javascript_translator import JavaScriptTranslator
from excel_formula_formatter.compact_excel_translator import CompactExcelTranslator


class AnnotatedExcelTranslator(SyntaxTranslatorBase):
    """Annotated Excel translator that preserves Excel syntax with helpful comments."""
    
    def get_language_name(self) -> str:
        return "Excel (Annotated)"
    
    def get_file_extension(self) -> str:
        return ".txt"
    
    def format_header_comment(self) -> str:
        return "// Excel Formula (annotated Excel syntax with smart indenting)"
    
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
        # Add minimal spacing around operators for readability in annotated mode
        return f' {operator} '
    
    def format_punctuation(self, punct: str) -> str:
        # Add spacing around function parentheses for readability
        if punct == '(':
            return '( '
        elif punct == ')':
            return ' )'
        else:
            return punct
    
    def reverse_parse_line(self, line: str) -> str:
        """Remove comments safely without consuming commas."""
        # Look for // that are NOT immediately preceded by a comma
        comment_pos = -1
        for i, char in enumerate(line):
            if char == '/' and i + 1 < len(line) and line[i + 1] == '/':
                # Found //, check if it's safe to remove
                # Look back to see if there's a comma without intervening non-space chars
                safe_to_remove = True
                look_back = i - 1
                while look_back >= 0 and line[look_back].isspace():
                    look_back -= 1
                
                # If the last non-space character is a comma, don't remove the comment
                if look_back >= 0 and line[look_back] == ',':
                    safe_to_remove = False
                
                if safe_to_remove:
                    comment_pos = i
                    break
        
        if comment_pos >= 0:
            return line[:comment_pos].rstrip()
        else:
            return line.strip()
    
    def reverse_parse_cell_reference(self, text: str) -> str:
        """No conversion needed for annotated format."""
        return text
    
    def reverse_parse_operator(self, text: str) -> str:
        """No conversion needed for annotated format."""
        return text


class PlainExcelTranslator(SyntaxTranslatorBase):
    """Plain Excel translator - pure Excel syntax with smart indenting, NO comments."""
    
    def get_language_name(self) -> str:
        return "Excel (Plain)"
    
    def get_file_extension(self) -> str:
        return ".txt"
    
    def format_header_comment(self) -> str:
        return ""  # NO header comment in plain mode
    
    def format_section_comment(self, comment: str) -> str:
        return ""  # NO section comments in plain mode
    
    def format_function_call(self, function_name: str) -> str:
        return function_name  # Keep original case and format
    
    def format_cell_reference(self, cell_ref: str) -> str:
        return cell_ref  # Keep unquoted
    
    def format_string_literal(self, string_val: str) -> str:
        return string_val  # Already quoted
    
    def format_number(self, number_val: str) -> str:
        return number_val
    
    def format_operator(self, operator: str) -> str:
        # Add spaces around operators for readability (like Excel Advanced Formula Environment)
        return f' {operator} '
    
    def format_punctuation(self, punct: str) -> str:
        # Add spacing around function parentheses for readability
        if punct == '(':
            return '( '
        elif punct == ')':
            return ' )'
        else:
            return punct
    
    def reverse_parse_line(self, line: str) -> str:
        """Plain mode should not have comments, but clean line just in case."""
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
        """Create formatter with JavaScript translator (mode 'j')."""
        return cls(JavaScriptTranslator())
    
    @classmethod  
    def create_annotated_formatter(cls):
        """Create formatter with Annotated Excel translator (mode 'a')."""
        return cls(AnnotatedExcelTranslator())
    
    @classmethod  
    def create_plain_formatter(cls):
        """Create formatter with Plain Excel translator (mode 'p')."""
        return cls(PlainExcelTranslator())
    
    @classmethod  
    def create_compact_formatter(cls):
        """Create formatter with Compact Excel translator (mode 'c')."""
        return cls(CompactExcelTranslator())
    
    @classmethod
    def create_formatter_by_mode(cls, mode: str):
        """Create formatter by mode code, accepting single letters
        (j, a, p, c) or full names (javascript, annotated, plain,
        compact) as the CLI does."""
        mode = mode.lower().strip()
        long_names = {
            'javascript': 'j',
            'annotated': 'a',
            'plain': 'p',
            'compact': 'c',
        }
        mode = long_names.get(mode, mode)
        if mode == 'j':
            return cls.create_javascript_formatter()
        elif mode == 'a':
            return cls.create_annotated_formatter()
        elif mode == 'p':
            return cls.create_plain_formatter()
        elif mode == 'c':
            return cls.create_compact_formatter()
        else:
            raise ValueError(f"Unknown mode '{mode}'. Use 'j' (JavaScript), 'a' (Annotated Excel), 'p' (Plain Excel), or 'c' (Compact Excel)")
    
    def get_mode_code(self) -> str:
        """Get the single letter mode code for this formatter."""
        if isinstance(self.translator, JavaScriptTranslator):
            return 'j'
        elif isinstance(self.translator, AnnotatedExcelTranslator):
            return 'a'
        elif isinstance(self.translator, PlainExcelTranslator):
            return 'p'
        elif isinstance(self.translator, CompactExcelTranslator):
            return 'c'
        else:
            return 'unknown'
    
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
            # Find where to insert {= (after header comment if present)
            insert_index = 1 if formatted_lines and formatted_lines[0].startswith('//') else 0
            formatted_lines.insert(insert_index, '{=')
            formatted_lines.append('}')
        else:
            # Preserve the leading = on the first content line so folded
            # output still reads as a formula
            for idx, line in enumerate(formatted_lines):
                stripped = line.lstrip()
                if stripped and not stripped.startswith('//'):
                    leading_ws = line[:len(line) - len(stripped)]
                    formatted_lines[idx] = leading_ws + '=' + stripped
                    break
        
        # Filter out empty lines for plain and compact modes
        if isinstance(self.translator, (PlainExcelTranslator, CompactExcelTranslator)):
            formatted_lines = [line for line in formatted_lines if line.strip()]
        
        return '\n'.join(formatted_lines)
    
    def unfold_formula(self, formatted_text: str) -> str:
        """Transform formatted text back to Excel formula."""
        if not formatted_text or not formatted_text.strip():
            return ""
            
        # Check if this was an array formula
        lines = formatted_text.strip().split('\n')
        is_array_formula = False
        
        # Look for array formula markers
        for i, line in enumerate(lines):
            if line.strip() == '{=':
                is_array_formula = True
                # Remove the {= line and find matching }
                lines = lines[:i] + lines[i+1:]
                # Find and remove the closing }
                for j in range(len(lines)-1, -1, -1):
                    if lines[j].strip() == '}':
                        lines = lines[:j] + lines[j+1:]
                        break
                break
        
        # Rejoin after removing array markers
        formatted_text = '\n'.join(lines)
        
        # SAFE comment removal that preserves commas
        no_comments = self._safe_remove_comments(formatted_text)
        
        # Flatten to single line
        single_line = whitespace_newline_rgx.sub(' ', no_comments)
        single_line = leading_trailing_space_rgx.sub('', single_line)
        
        if not single_line:
            return ""
            
        # Use translator-specific reverse parsing
        excel_formula = self._reverse_parse_with_translator(single_line)
        
        # Add appropriate prefix
        if is_array_formula:
            if not excel_formula.startswith('{='):
                excel_formula = '{=' + excel_formula + '}'
        else:
            if not excel_formula.startswith('='):
                excel_formula = '=' + excel_formula
            
        return excel_formula
    
    def _safe_remove_comments(self, text: str) -> str:
        """Safely remove comments without touching commas or string contents.

        // sequences inside string literals (e.g. "http://x.com") are
        NOT comments: the scan jumps over quoted spans entirely. The
        comma guard from the original implementation is retained so a
        // directly after a trailing comma is left alone.
        """
        lines = text.split('\n')
        cleaned_lines = []

        for line in lines:
            # Skip lines that are entirely comments
            if comment_line_detection_rgx.match(line):
                continue

            comment_pos = -1
            i = 0
            length = len(line)
            while i < length:
                # Jump over string literals so their contents are opaque
                if line[i] == '"':
                    end = scan_string_literal(line, i)
                    if end < 0:
                        break  # Unclosed on this line: no comment scan past it
                    i = end
                    continue

                if line[i] == '/' and i + 1 < length and line[i + 1] == '/':
                    # Found // outside any string; apply the comma guard
                    look_back = i - 1
                    while look_back >= 0 and line[look_back].isspace():
                        look_back -= 1

                    if not (look_back >= 0 and line[look_back] == ','):
                        comment_pos = i
                        break

                i += 1

            if comment_pos >= 0:
                cleaned_line = line[:comment_pos].rstrip()
            else:
                cleaned_line = line

            if cleaned_line.strip():  # Only add non-empty lines
                cleaned_lines.append(cleaned_line)

        return '\n'.join(cleaned_lines)
    
    def _parse_excel_tokens(self, formula: str) -> list:
        """Parse Excel formula into tokens with type information.

        Opaque segments (strings with "" escapes, 'quoted sheet' refs,
        [structured references], {array constants}) are scanned as
        single atomic tokens so their interior can never be split,
        re-spaced, or mistaken for argument separators.
        """
        tokens = []
        i = 0
        length = len(formula)

        while i < length:
            # Skip whitespace between tokens
            if formula[i].isspace():
                i += 1
                continue

            # String literals (with Excel's doubled "" quote escaping)
            if formula[i] == '"':
                end = scan_string_literal(formula, i)
                if end < 0:
                    raise ValueError(
                        f"Unclosed string literal starting at position {i}: "
                        f"...{formula[i:i + 30]}"
                    )
                tokens.append(('string', formula[i:end]))
                i = end
                continue

            # Quoted sheet names: 'My Sheet'!A1 or 'My Sheet'!Name
            if formula[i] == "'":
                end = find_opaque_end(formula, i)
                if end < 0:
                    raise ValueError(
                        f"Unclosed sheet name quote at position {i}: "
                        f"...{formula[i:i + 30]}"
                    )
                token_text = formula[i:end]
                # Consume the ! and the reference that follows the sheet name
                if end < length and formula[end] == '!':
                    ref_match = cell_ref_all_rgx.match(formula, end + 1)
                    if ref_match:
                        token_text = formula[i:ref_match.end()]
                        i = ref_match.end()
                    else:
                        # Named range on the sheet: collect the name
                        j = end + 1
                        while j < length and (formula[j].isalnum() or formula[j] in '_.'):
                            j += 1
                        token_text = formula[i:j]
                        i = j
                else:
                    i = end
                tokens.append(('cell_ref', token_text))
                continue

            # Structured references: [Column], [@[Col]], [[#All],[Col]]
            if formula[i] == '[':
                end = find_opaque_end(formula, i)
                if end < 0:
                    raise ValueError(
                        f"Unclosed structured reference bracket at position {i}: "
                        f"...{formula[i:i + 30]}"
                    )
                tokens.append(('structured_ref', formula[i:end]))
                i = end
                continue

            # Array constants: {1,2;3,4}
            if formula[i] == '{':
                end = find_opaque_end(formula, i)
                if end < 0:
                    raise ValueError(
                        f"Unclosed array constant brace at position {i}: "
                        f"...{formula[i:i + 30]}"
                    )
                tokens.append(('array_literal', formula[i:end]))
                i = end
                continue

            # Cell references (absolute $A$1, ranges, unquoted sheet refs)
            cell_match = cell_ref_all_rgx.match(formula, i)
            if cell_match:
                tokens.append(('cell_ref', cell_match.group(0)))
                i = cell_match.end()
                continue

            # Numbers, including scientific notation (1.5E+10) and .5
            if formula[i].isdigit() or (
                formula[i] == '.' and i + 1 < length and formula[i + 1].isdigit()
            ):
                number_match = number_rgx.match(formula, i)
                if number_match:
                    tokens.append(('number', number_match.group(0)))
                    i = number_match.end()
                    continue

            # Two-character operators
            if i < length - 1:
                two_char = formula[i:i + 2]
                if two_char in ['<>', '>=', '<=']:
                    tokens.append(('operator', two_char))
                    i += 2
                    continue

            # Single character operators
            if formula[i] in '+-*/=<>&':
                tokens.append(('operator', formula[i]))
                i += 1
                continue

            # Punctuation (brackets/braces handled above as atomic tokens)
            if formula[i] in '(),:;!%^':
                tokens.append(('punctuation', formula[i]))
                i += 1
                continue

            # Collect word/identifier
            start = i
            while (
                i < length
                and not formula[i].isspace()
                and formula[i] not in '+-*/=<>(),[]{}:;!&%^"\''
            ):
                i += 1

            if start < i:
                token_text = formula[start:i]
                tokens.append((self._classify_token(token_text), token_text))
            else:
                # Unrecognized character: keep it verbatim so nothing is lost
                tokens.append(('identifier', formula[i]))
                i += 1

        return self._promote_function_tokens(tokens)

    def _promote_function_tokens(self, tokens: list) -> list:
        """Classify identifiers followed by ( as functions via lookahead.

        This replaces the hardcoded function-name list: any identifier
        immediately followed by an opening parenthesis is a function
        call (IFERROR, TEXTSPLIT, LAMBDA-variable invocations, ...),
        and anything else is not, even if its name matches a known
        function.
        """
        promoted = []
        for index, (token_type, token_text) in enumerate(tokens):
            if token_type in ('identifier', 'function'):
                next_is_paren = (
                    index + 1 < len(tokens) and tokens[index + 1][1] == '('
                )
                token_type = 'function' if next_is_paren else 'identifier'
            promoted.append((token_type, token_text))
        return promoted

    
    def _classify_token(self, token: str) -> str:
        """Classify a word token by type using exact (full) matches.

        Function classification is NOT done here: it happens by paren
        lookahead in _promote_function_tokens, so unlisted functions
        work and function-named identifiers don't misclassify.
        """
        if cell_ref_all_rgx.fullmatch(token):
            return 'cell_ref'
        if number_rgx.fullmatch(token):
            return 'number'
        return 'identifier'
    
    def _format_tokens_with_translator(self, tokens: list) -> list:
        """Convert tokens using the configured translator with TRUE function isolation."""
        lines = []
        
        # Add header comment if translator provides one
        header = self.translator.format_header_comment()
        if header:
            lines.append(header)
        
        # Process tokens with isolated function handling
        processed_lines = self._process_token_sequence(tokens, base_depth=0)
        lines.extend(processed_lines)
        
        return lines

    def _process_token_sequence(self, tokens: list, base_depth: int) -> list:
        """Process a sequence of tokens with proper function isolation."""
        lines = []
        current_line = ""
        i = 0
        
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
                    else:
                        # All other functions (including AND, OR) use simple generic processing
                        func_lines = self._process_generic_function(token_text, arg_tokens, base_depth)
                    
                    # Single-line function renders stay inline with the
                    # surrounding expression (e.g. INDIRECT("O1")="...")
                    if len(func_lines) == 1:
                        current_line += func_lines[0].strip()
                    else:
                        # Multi-line function: flush and emit its lines
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
                # Top-level comma - handle spacing based on translator
                if isinstance(self.translator, CompactExcelTranslator):
                    current_line += self.translator.format_punctuation(token_text)  # No space
                else:
                    current_line += self.translator.format_punctuation(token_text) + " "  # Add space
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
        
        # Add function comment only if translator supports it
        comment = self.translator.get_function_comment(func_name)
        if comment:
            comment_line = self.translator.format_section_comment(comment)
            if comment_line:  # Only add if translator returns non-empty comment
                lines.append(self.translator.indent(base_depth) + comment_line)
        
        # Function header
        lines.append(self.translator.indent(base_depth) + 
                    self.translator.format_function_call(func_name) + 
                    self.translator.format_punctuation('('))
        
        # Split arguments by top-level commas
        argument_groups = self._split_by_top_level_commas(arg_tokens)
        
        # Add initial separator only if we have arguments and translator supports comments
        if argument_groups:
            separator = self.translator.format_section_comment("── CASE/RESULT PAIR ──")
            if separator:  # Only add if translator returns non-empty separator
                lines.append(self.translator.indent(base_depth + 1) + separator)
        
        for arg_index, arg_group in enumerate(argument_groups):
            # Add separator before each condition (even arguments > 1)
            if arg_index > 1 and arg_index % 2 == 0:
                separator = self.translator.format_section_comment("── CASE/RESULT PAIR ──")
                if separator:  # Only add if translator returns non-empty separator
                    lines.append("")  # Blank line
                    lines.append(self.translator.indent(base_depth + 1) + separator)
            
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
        
        # Add function comment if translator supports it
        comment = self.translator.get_function_comment(func_name)
        if comment:
            comment_line = self.translator.format_section_comment(comment)
            if comment_line:
                lines.append(self.translator.indent(base_depth) + comment_line)
        
        # Function header
        lines.append(self.translator.indent(base_depth) + 
                    self.translator.format_function_call(func_name) + 
                    self.translator.format_punctuation('('))
        
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
                if isinstance(self.translator, CompactExcelTranslator):
                    combined_line = (self.translator.indent(base_depth + 1) + var_name + 
                                   self.translator.format_punctuation(',') + value_str)
                else:
                    combined_line = (self.translator.indent(base_depth + 1) + var_name + 
                                   self.translator.format_punctuation(',') + " " + value_str)
                
                is_last_pair = i + 2 >= len(argument_groups)
                
                if len(combined_line) <= let_value_wrap_threshold:
                    # Add comma if not the last pair
                    if not is_last_pair:
                        combined_line += self.translator.format_punctuation(',')
                    lines.append(combined_line)
                else:
                    # Value too long for one line: expand it with normal
                    # function processing and hang it off the variable name
                    value_lines = self._process_token_sequence(
                        argument_groups[i + 1], base_depth + 1
                    )
                    if value_lines:
                        first_value = value_lines[0].lstrip()
                        head = (self.translator.indent(base_depth + 1) + var_name +
                                self.translator.format_punctuation(','))
                        if not isinstance(self.translator, CompactExcelTranslator):
                            head += " "
                        lines.append(head + first_value)
                        lines.extend(value_lines[1:])
                        if not is_last_pair:
                            lines[-1] += self.translator.format_punctuation(',')
                    else:
                        if not is_last_pair:
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

    def _process_generic_function(self, func_name: str, arg_tokens: list, base_depth: int) -> list:
        """Process all functions with simple, consistent formatting."""
        lines = []
        
        # Split arguments by top-level commas
        argument_groups = self._split_by_top_level_commas(arg_tokens)
        
        # Filter out any empty argument groups
        argument_groups = [group for group in argument_groups if group]
        
        if not argument_groups:
            # Empty function call
            func_str = (self.translator.format_function_call(func_name) + 
                       self.translator.format_punctuation('(') + 
                       self.translator.format_punctuation(')'))
            lines.append(self.translator.indent(base_depth) + func_str)
            return lines
        
        # Check for simple inline case: one argument with simple content
        if len(argument_groups) == 1:
            arg_str = self._tokens_to_string(argument_groups[0]).strip()
            
            # Simple inline criteria: one argument, reasonable length, no nested functions
            has_nested_functions = any(token_type == 'function' for token_type, _ in argument_groups[0])
            total_length = len(func_name) + len(arg_str) + 2  # +2 for parentheses
            
            if not has_nested_functions and total_length <= 40:
                # Keep inline
                func_str = (self.translator.format_function_call(func_name) + 
                           self.translator.format_punctuation('(') + 
                           arg_str + 
                           self.translator.format_punctuation(')'))
                lines.append(self.translator.indent(base_depth) + func_str)
                return lines
        
        # Multi-line formatting: one argument per line
        lines.append(self.translator.indent(base_depth) + 
                    self.translator.format_function_call(func_name) + 
                    self.translator.format_punctuation('('))
        
        # Process each argument on its own line
        for arg_index, arg_group in enumerate(argument_groups):
            arg_lines = self._process_token_sequence(arg_group, base_depth + 1)
            
            # Add comma if not last argument
            if arg_index < len(argument_groups) - 1:
                if arg_lines:
                    arg_lines[-1] += self.translator.format_punctuation(',')
            
            lines.extend(arg_lines)
        
        # Closing parenthesis
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
                    if isinstance(self.translator, CompactExcelTranslator):
                        result += ','  # No space after comma in compact mode
                    else:
                        result += ', '  # Always add space after comma in non-compact modes
                else:
                    result += self.translator.format_punctuation(token_text)
            else:
                result += token_text
        
        return result.strip()
    
    def _reverse_parse_with_translator(self, formatted_text: str) -> str:
        """Use translator-specific reverse parsing.

        All spacing cleanups operate on text whose opaque spans (string
        literals, quoted sheet names, structured references, array
        constants) have been replaced with placeholders, so cleanup can
        never alter their interior. Spans are restored at the end.
        """
        protected, spans = protect_opaque_spans(formatted_text)

        # Apply translator-specific reverse transformations
        result = self.translator.reverse_parse_cell_reference(protected)
        result = self.translator.reverse_parse_operator(result)

        # Apply line-level reverse parsing if available
        if hasattr(self.translator, 'reverse_parse_line'):
            result = self.translator.reverse_parse_line(result)

        # Remove extra spaces around parentheses that were added for formatting
        result = paren_leading_space_rgx.sub('(', result)
        result = paren_trailing_space_rgx.sub(')', result)

        # Normalize multiple spaces to single spaces
        result = multiple_spaces_rgx.sub(' ', result)

        # Clean up comma spacing - add space after comma, none before
        result = comma_spacing_rgx.sub(', ', result)

        # Excel modes (except compact) keep operator spacing for readability
        if isinstance(self.translator, (AnnotatedExcelTranslator, PlainExcelTranslator)):
            pass  # Don't strip operator spacing for Excel modes
        else:
            # Compact and JavaScript modes remove operator/comma spacing.
            # String literals are already protected, so no inner content
            # can be affected here.
            result = operator_spacing_rgx.sub(r'\1', result)
            result = multi_char_operator_spacing_rgx.sub(r'\1', result)

        result = restore_opaque_spans(result, spans)
        return result.strip()


def find_structural_issues(formula_text: str) -> list:
    """Detect token adjacencies that indicate missing commas/operators.

    Excel has no juxtaposition operator: two value tokens side by side
    (e.g. a number directly followed by an identifier, or ')' followed
    by a name) always indicate a lost separator - except two adjacent
    references, which can be Excel's space-as-intersection operator and
    are flagged as possibly intentional.

    Returns a list of issue dicts with 'left', 'right', and 'message'
    keys. An empty list means no structural problems were found.
    """
    if not formula_text or not formula_text.strip():
        return []

    clean = formula_text.strip()
    if clean.startswith('{=') and clean.endswith('}'):
        clean = clean[2:-1]
    elif clean.startswith('='):
        clean = clean[1:]

    probe_formatter = ModularExcelFormatter.create_plain_formatter()
    try:
        tokens = probe_formatter._parse_excel_tokens(clean)
    except ValueError as e:
        return [{'left': '', 'right': '', 'message': str(e)}]

    value_end_types = {
        'cell_ref', 'number', 'string', 'identifier',
        'structured_ref', 'array_literal',
    }
    value_start_types = {
        'cell_ref', 'number', 'string', 'identifier', 'function',
        'structured_ref', 'array_literal',
    }

    def ends_value(token_type, token_text):
        return token_type in value_end_types or token_text == ')'

    def starts_value(token_type, token_text):
        return token_type in value_start_types or token_text == '('

    issues = []
    for index in range(len(tokens) - 1):
        left_type, left_text = tokens[index]
        right_type, right_text = tokens[index + 1]

        # A function token followed by its ( is a call, not an adjacency
        if left_type == 'function' and right_text == '(':
            continue

        # An identifier followed by a bracketed segment is a structured
        # table reference (Table1[Column]) - legal adjacency
        if left_type == 'identifier' and right_type == 'structured_ref':
            continue

        if ends_value(left_type, left_text) and starts_value(right_type, right_text):
            if left_type == 'cell_ref' and right_type == 'cell_ref':
                message = (
                    f"'{left_text}' directly followed by '{right_text}' - "
                    f"missing comma/operator, or intentional range "
                    f"intersection (space operator)?"
                )
            else:
                message = (
                    f"'{left_text}' directly followed by '{right_text}' - "
                    f"missing comma or operator?"
                )
            issues.append({
                'left': left_text,
                'right': right_text,
                'message': message,
            })

    # LET pair-shape validation: name slots (even positions, except the
    # final expression) must be single plain identifiers, and the
    # argument count must be odd. A single lost comma shifts every
    # downstream pair, which adjacency checks cannot see because the
    # shifted text is grammatically clean - but the shape check can.
    for index, (token_type, token_text) in enumerate(tokens):
        if token_type != 'function' or token_text.upper() != 'LET':
            continue
        if index + 1 >= len(tokens) or tokens[index + 1][1] != '(':
            continue

        arg_tokens, _end = probe_formatter._extract_function_arguments(
            tokens, index + 1
        )
        groups = probe_formatter._split_by_top_level_commas(arg_tokens)

        if len(groups) % 2 == 0:
            issues.append({
                'left': token_text, 'right': '',
                'message': (
                    f"LET has an even number of arguments "
                    f"({len(groups)}) - a comma is likely missing"
                ),
            })

        for group_index in range(0, len(groups) - 1, 2):
            group = groups[group_index]
            content = [tok for tok in group if tok[1].strip()]
            is_plain_name = (
                len(content) == 1 and content[0][0] == 'identifier'
            )
            if not is_plain_name:
                rendered = probe_formatter._tokens_to_string(group).strip()
                issues.append({
                    'left': rendered, 'right': '',
                    'message': (
                        f"LET name slot {group_index // 2 + 1} holds "
                        f"'{rendered[:40]}' instead of a variable name - "
                        f"pairing shifted by a missing comma earlier?"
                    ),
                })
                break  # One shift report per LET; the rest cascade from it

    return issues



def detect_current_mode(text: str) -> str:
    """Detect what formatter mode the text is currently in."""
    if not text or not text.strip():
        return 'unknown'
    
    lines = text.strip().split('\n')
    
    # Single line is likely unfolded (plain Excel)
    if len(lines) == 1:
        return 'p'  # Plain
    
    # Check for mode indicators
    text_content = '\n'.join(lines)
    
    # Check for explicit mode indicators in comments first
    if '//' in text_content:
        if 'JavaScript syntax' in text_content:
            return 'j'  # JavaScript
        elif 'annotated Excel syntax' in text_content:
            return 'a'  # Annotated Excel
        elif 'plain Excel syntax' in text_content or 'Excel Formula' in text_content:
            return 'a'  # Probably annotated (backward compatibility)
        else:
            # Has comments but no explicit mode indicator
            # Backtick cell reference markers are unique to JavaScript mode
            if '`' in text_content:
                return 'j'  # Has backtick cell refs, JavaScript
            else:
                return 'a'  # Has comments but no backticks, likely Annotated
    
    # No comments found - check indentation patterns
    has_indentation = any(line.startswith('    ') or line.startswith('\t') for line in lines)
    
    if has_indentation:
        # Has indentation but NO comments: Plain or Compact Excel mode.
        # Distinguish by the paren spacing the translators emit: plain
        # writes "FUNC( " and " )", compact writes "FUNC(" and ")".
        # Opaque spans are protected so string contents can't confuse
        # the heuristic.
        protected_content, _spans = protect_opaque_spans(text_content)
        
        if '( ' in protected_content or ' )' in protected_content:
            return 'p'  # Plain mode (readable paren spacing)
        
        # Spaced operators or mid-line comma spacing also indicate plain
        has_spaced_commas = ', ' in protected_content
        has_spaced_operators = (
            ' = ' in protected_content
            or ' > ' in protected_content
            or ' < ' in protected_content
        )
        if has_spaced_commas or has_spaced_operators:
            return 'p'  # Plain mode (readable spacing)
        
        return 'c'  # Compact mode (minimal spacing)
    
    return 'p'  # Default to plain if uncertain


def safe_mode_switch(text: str, current_mode: str, target_mode: str, should_refold: bool = True) -> str:
    """Safely switch between formatter modes by unfolding first."""
    if not text or not text.strip():
        return ""
    
    if current_mode == target_mode:
        return text  # No change needed
    
    try:
        # Step 1: Unfold using current mode formatter
        current_formatter = ModularExcelFormatter.create_formatter_by_mode(current_mode)
        unfolded = current_formatter.unfold_formula(text)
        
        # Step 2: If requested, fold using target mode formatter
        if should_refold:
            target_formatter = ModularExcelFormatter.create_formatter_by_mode(target_mode)
            return target_formatter.fold_formula(unfolded)
        else:
            # Just return unfolded
            return unfolded
            
    except Exception:
        # If mode switching fails, return original text
        return text


def auto_format_with_mode(input_text: str, mode: str = 'j') -> str:
    """Automatically determine whether to fold or unfold using specified mode."""
    if not input_text or not input_text.strip():
        return ""
    
    # Create formatter with specified mode
    formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
    
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
        has_comments = any('//' in line for line in lines)
        has_indentation = any(line.startswith('    ') or line.startswith('\t') for line in lines)
        
        if has_comments or has_indentation:
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
        print("Usage: python modular_excel_formatter.py <mode> [fold|unfold|auto]", file=sys.stderr)
        print("  mode: j (JavaScript), a (Annotated Excel), p (Plain Excel), c (Compact Excel)", file=sys.stderr)
        print("  operation: fold, unfold, or auto (default)", file=sys.stderr)
        print("  j: JavaScript-like syntax for highlighting", file=sys.stderr)
        print("  a: Excel syntax with helpful comments", file=sys.stderr)
        print("  p: Pure Excel syntax with smart indenting only", file=sys.stderr)
        print("  c: Compact Excel syntax with minimal spacing (for 8K limit)", file=sys.stderr)
        return 1
    
    mode = sys.argv[1].lower().strip()
    operation = sys.argv[2].lower() if len(sys.argv) > 2 else 'auto'  # Default to auto
    
    if mode not in ['j', 'a', 'p', 'c']:
        print(f"Unknown mode: {mode}", file=sys.stderr)
        print("Available modes: j (JavaScript), a (Annotated Excel), p (Plain Excel), c (Compact Excel)", file=sys.stderr)
        return 1
    
    try:
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    
    input_text = sys.stdin.read()
    if not input_text.strip():
        return 0
    
    try:
        if operation == 'fold':
            result = formatter.fold_formula(input_text)
            excel_side = input_text
        elif operation == 'unfold':
            result = formatter.unfold_formula(input_text)
            excel_side = result
        elif operation == 'auto':
            result = auto_format_with_mode(input_text, mode)
            was_single_line = len(input_text.strip().split('\n')) == 1
            excel_side = input_text if was_single_line else result
        else:
            print(f"Unknown operation: {operation}", file=sys.stderr)
            return 1
        
        for issue in find_structural_issues(excel_side):
            print(f"warning: {issue['message']}", file=sys.stderr)
        
        print(result)
        return 0
        
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

# End of file #
