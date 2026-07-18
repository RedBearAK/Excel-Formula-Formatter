#!/usr/bin/env python3
"""
Excel formula regex patterns.
File: excel_formula_formatter/excel_formula_patterns.py

Tokenizer-integrity revision:
- Cell reference patterns support absolute refs ($A$1) and limit column
  letters to the real Excel maximum of three, so all-caps identifiers
  like RATE1 no longer collide with cell references.
- Number pattern supports scientific notation (1.5E+10) and bare
  leading-dot decimals (.5).
- Backtick pattern supports JavaScript-mode cell reference marking,
  which replaced double-quoted cell refs to remove ambiguity with
  genuine string literals.
- The excel_functions_rgx list is retained for reference and any legacy
  consumers, but classification now uses paren lookahead in the
  tokenizer, so unlisted functions (IFERROR, TEXTSPLIT, ...) work too.

NOTE: Opaque-span handling (strings with "" escapes, 'sheet names',
[structured refs], {array constants}) is done by character scanners in
formula_text_guards.py because nesting and doubled-character escapes do
not express safely as regex patterns.
"""

import re


# Cell reference patterns
# Column letters limited to 1-3 (Excel max is XFD). Optional $ before
# column and/or row for absolute references. Optional unquoted sheet
# prefix. No leading \b: it cannot assert before $, so the tokenizer and
# classifier are responsible for calling these at token boundaries
# (position match in the tokenizer, fullmatch in classification).
cell_ref_simple_rgx = re.compile(r'\$?[A-Z]{1,3}\$?\d+\b')
cell_ref_range_rgx = re.compile(r'\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+\b')
cell_ref_sheet_rgx = re.compile(r'[A-Za-z0-9_]+!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?\b')
cell_ref_all_rgx = re.compile(r'(?:[A-Za-z0-9_]+!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?\b')

# Excel function names (common ones). Kept for reference/legacy use;
# the tokenizer classifies any identifier followed by ( as a function.
excel_functions_rgx = re.compile(r'\b(?:SUM|IF|VLOOKUP|HLOOKUP|INDEX|MATCH|SUMIF|SUMIFS|COUNTIF|COUNTIFS|AVERAGEIF|AVERAGEIFS|LEN|MID|LEFT|RIGHT|FIND|SEARCH|SUBSTITUTE|CONCATENATE|TEXT|VALUE|DATE|TODAY|NOW|YEAR|MONTH|DAY|WEEKDAY|WORKDAY|NETWORKDAYS|PMT|PV|FV|RATE|NPER|NPV|IRR|AND|OR|NOT|ISERROR|ISBLANK|ISNUMBER|ISTEXT|CHOOSE|INDIRECT|OFFSET|ROW|COLUMN|ROWS|COLUMNS|COUNTA|COUNT|MAX|MIN|AVERAGE|MEDIAN|MODE|STDEV|VAR|ROUND|ROUNDUP|ROUNDDOWN|INT|ABS|SQRT|POWER|EXP|LN|LOG|LOG10|SIN|COS|TAN|ASIN|ACOS|ATAN|PI|RAND|RANDBETWEEN|LET|LAMBDA|MAP|FILTER|SORT|UNIQUE|SEQUENCE|XLOOKUP|XMATCH|IFS|SWITCH|TEXTJOIN|CONCAT|IFERROR|IFNA|TEXTSPLIT|TEXTBEFORE|TEXTAFTER|BYROW|BYCOL|SCAN|REDUCE|MAKEARRAY|VSTACK|HSTACK|TAKE|DROP|TOCOL|TOROW|WRAPROWS|WRAPCOLS)\b', re.IGNORECASE)

# Number patterns: integers, decimals, bare-dot decimals, scientific
# notation. Used with .match() at a scan position in the tokenizer and
# with .fullmatch() in classification.
number_rgx = re.compile(r'(?:\d+(?:\.\d+)?|\.\d+)(?:[Ee][+-]?\d+)?')

# String literal WITHOUT escape support - legacy consumers only.
# New code must use scan_string_literal() from formula_text_guards.
string_literal_rgx = re.compile(r'"[^"]*"')

# Excel operators that need translation
excel_not_equal_rgx = re.compile(r'<>')
js_not_equal_rgx = re.compile(r'!=')

# JavaScript-mode cell reference marker: backtick-quoted spans.
# Backticks cannot appear in Excel formulas, so this is unambiguous,
# unlike the old double-quoted cell refs which collided with genuine
# string literals such as "A1".
backtick_span_rgx = re.compile(r'`[^`]*`')

# Comment patterns for safe removal during unfold
# Only match lines that START with comments (after optional whitespace)
comment_line_rgx = re.compile(r'^\s*(?://|#).*$', re.MULTILINE)

# Inline comment removal that preserves commas
inline_comment_rgx = re.compile(r'(?<![,])\s*//.*$', re.MULTILINE)

# More conservative inline comment removal requiring 2+ spaces before //
inline_comment_safe_rgx = re.compile(r'\s{2,}//.*$', re.MULTILINE)

# Specific pattern for line-level comment removal in translators
line_comment_removal_rgx = re.compile(r'(?<![,])\s*//.*$')

# Cleanup patterns for reverse parsing
space_cleanup_operators_rgx = re.compile(r'\s*([+\-*/=<>!&,()])\s*')
space_cleanup_multi_char_rgx = re.compile(r'\s*(<>|>=|<=)\s*')
space_cleanup_whitespace_rgx = re.compile(r'\s+')

# Whitespace and newline patterns
whitespace_newline_rgx = re.compile(r'\s*\r?\n\s*')
leading_trailing_space_rgx = re.compile(r'^\s+|\s+$')

# Cleanup patterns for reverse parsing.
# These are ONLY safe to apply to text whose opaque spans have been
# protected via formula_text_guards.protect_opaque_spans().
paren_leading_space_rgx = re.compile(r'\(\s+')
paren_trailing_space_rgx = re.compile(r'\s+\)')
multiple_spaces_rgx = re.compile(r'\s+')
comma_spacing_rgx = re.compile(r'\s*,\s*')

# String literal protection - legacy consumers only, does not handle
# doubled-quote escapes. New code uses formula_text_guards.
string_literal_protection_rgx = re.compile(r'"[^"]*"')

# Operator spacing cleanup patterns
operator_spacing_rgx = re.compile(r'\s*([+\-*/=<>!,()])\s*')
multi_char_operator_spacing_rgx = re.compile(r'\s*(<>|>=|<=|!=)\s*')

# Comment detection pattern
comment_line_detection_rgx = re.compile(r'^\s*(?://|#)')

# End of file #
