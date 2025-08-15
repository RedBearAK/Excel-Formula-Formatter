#!/usr/bin/env python3
"""
Excel formula regex patterns module.
File: excel_formula_formatter/excel_formula_patterns.py
"""

import re


# Cell reference patterns
cell_ref_simple_rgx = re.compile(r'\b[A-Z]+\$?\d+\b')
cell_ref_range_rgx = re.compile(r'\b[A-Z]+\$?\d+:[A-Z]+\$?\d+\b')
cell_ref_sheet_rgx = re.compile(r'\b[A-Za-z0-9_]+![A-Z]+\$?\d+(?::[A-Z]+\$?\d+)?\b')
cell_ref_all_rgx = re.compile(r'\b(?:[A-Za-z0-9_]+!)?[A-Z]+\$?\d+(?::[A-Z]+\$?\d+)?\b')

# Excel function names (common ones)
excel_functions_rgx = re.compile(r'\b(?:SUM|IF|VLOOKUP|HLOOKUP|INDEX|MATCH|SUMIF|SUMIFS|COUNTIF|COUNTIFS|AVERAGEIF|AVERAGEIFS|LEN|MID|LEFT|RIGHT|FIND|SEARCH|SUBSTITUTE|CONCATENATE|TEXT|VALUE|DATE|TODAY|NOW|YEAR|MONTH|DAY|WEEKDAY|WORKDAY|NETWORKDAYS|PMT|PV|FV|RATE|NPER|NPV|IRR|AND|OR|NOT|ISERROR|ISBLANK|ISNUMBER|ISTEXT|CHOOSE|INDIRECT|OFFSET|ROW|COLUMN|ROWS|COLUMNS|COUNTA|COUNT|MAX|MIN|AVERAGE|MEDIAN|MODE|STDEV|VAR|ROUND|ROUNDUP|ROUNDDOWN|INT|ABS|SQRT|POWER|EXP|LN|LOG|LOG10|SIN|COS|TAN|ASIN|ACOS|ATAN|PI|RAND|RANDBETWEEN|LET|LAMBDA|MAP|FILTER|SORT|UNIQUE|SEQUENCE|XLOOKUP|XMATCH|IFS|SWITCH|TEXTJOIN|CONCAT)\b', re.IGNORECASE)

# String literals in Excel (double quotes)
string_literal_rgx = re.compile(r'"[^"]*"')

# Number patterns
number_rgx = re.compile(r'\b\d+(?:\.\d+)?\b')

# Excel operators that need translation
excel_not_equal_rgx = re.compile(r'<>')
js_not_equal_rgx = re.compile(r'!=')

# Comment patterns for removal during unfold - support both // and # comments
comment_line_rgx = re.compile(r'^\s*(?://|#).*$', re.MULTILINE)
inline_comment_rgx = re.compile(r'\s*(?://|#).*$')

# Specific pattern for line-level comment removal in translators
line_comment_removal_rgx = re.compile(r'\s*//.*$')

# Cleanup patterns for reverse parsing
space_cleanup_operators_rgx = re.compile(r'\s*([+\-*/=<>!&,()])\s*')
space_cleanup_multi_char_rgx = re.compile(r'\s*(<>|>=|<=)\s*')
space_cleanup_whitespace_rgx = re.compile(r'\s+')

# Whitespace and newline patterns
whitespace_newline_rgx = re.compile(r'\s*\r?\n\s*')
leading_trailing_space_rgx = re.compile(r'^\s+|\s+$')

# End of file #
