#!/usr/bin/env python3
"""
Character-level scanners for opaque segments of Excel formulas.
File: excel_formula_formatter/formula_text_guards.py

Opaque segments are spans whose interior must never be altered by
formatting or cleanup passes:

    "string literals"       (with Excel's doubled "" quote escaping)
    'quoted sheet names'    (with doubled '' apostrophe escaping)
    [structured references] (with nesting, e.g. Table1[[#All],[Col]])
    {array constants}       (e.g. {1,2;3,4})

These are deliberately plain character scanners rather than regexes:
nesting and doubled-character escapes are awkward or impossible to
express safely as patterns, and per project convention regex patterns
live in excel_formula_patterns.py anyway.

All scan_* functions take (text, start) where text[start] is the opening
character, and return the index just past the closing character, or -1
if the segment is unclosed.
"""


def scan_string_literal(text: str, start: int) -> int:
    """Scan a double-quoted string with doubled-quote escaping."""
    if start >= len(text) or text[start] != '"':
        return -1

    i = start + 1
    length = len(text)
    while i < length:
        if text[i] != '"':
            i += 1
            continue
        # Doubled quote is an escaped quote, keep scanning
        if i + 1 < length and text[i + 1] == '"':
            i += 2
            continue
        return i + 1  # Closing quote

    return -1  # Unclosed


def scan_single_quoted(text: str, start: int) -> int:
    """Scan a single-quoted span (sheet name) with doubled-quote escaping."""
    if start >= len(text) or text[start] != "'":
        return -1

    i = start + 1
    length = len(text)
    while i < length:
        if text[i] != "'":
            i += 1
            continue
        if i + 1 < length and text[i + 1] == "'":
            i += 2
            continue
        return i + 1

    return -1


def scan_bracketed(text: str, start: int) -> int:
    """Scan a [bracketed] span with nesting support."""
    if start >= len(text) or text[start] != '[':
        return -1

    depth = 0
    i = start
    length = len(text)
    while i < length:
        if text[i] == '[':
            depth += 1
        elif text[i] == ']':
            depth -= 1
            if depth == 0:
                return i + 1
        i += 1

    return -1


def scan_braced(text: str, start: int) -> int:
    """Scan a {braced} span (array constant) with nesting support."""
    if start >= len(text) or text[start] != '{':
        return -1

    depth = 0
    i = start
    length = len(text)
    while i < length:
        if text[i] == '{':
            depth += 1
        elif text[i] == '}':
            depth -= 1
            if depth == 0:
                return i + 1
        i += 1

    return -1


def find_opaque_end(text: str, start: int) -> int:
    """Dispatch to the right scanner for the opening char, or return -1."""
    opener = text[start]
    if opener == '"':
        return scan_string_literal(text, start)
    if opener == "'":
        return scan_single_quoted(text, start)
    if opener == '[':
        return scan_bracketed(text, start)
    if opener == '{':
        return scan_braced(text, start)
    return -1


def protect_opaque_spans(text: str) -> tuple:
    """Replace opaque spans with placeholders so cleanup passes can't touch them.

    Returns (protected_text, spans) where spans is the ordered list of
    original span texts. Placeholders use NUL delimiters, which cannot
    appear in Excel formulas or our formatted output.
    """
    spans = []
    out_parts = []
    i = 0
    length = len(text)

    while i < length:
        if text[i] in '"\'[{':
            end = find_opaque_end(text, i)
            if end > 0:
                out_parts.append(f"\x00{len(spans)}\x00")
                spans.append(text[i:end])
                i = end
                continue
        out_parts.append(text[i])
        i += 1

    return ''.join(out_parts), spans


def restore_opaque_spans(text: str, spans: list) -> str:
    """Restore placeholders created by protect_opaque_spans."""
    result = text
    for index, span in enumerate(spans):
        result = result.replace(f"\x00{index}\x00", span)
    return result


def normalize_formula_for_comparison(formula: str) -> str:
    """Normalize a formula for round-trip comparison.

    Strips the leading = or {= marker and removes whitespace, but ONLY
    outside opaque spans: spaces inside string literals, quoted sheet
    names, structured references, and array constants are significant
    and must survive normalization so tests cannot mask corruption.
    """
    normalized = formula.strip()
    if normalized.startswith('{=') and normalized.endswith('}'):
        normalized = normalized[2:-1]
    elif normalized.startswith('='):
        normalized = normalized[1:]

    protected, spans = protect_opaque_spans(normalized)
    protected = protected.replace(' ', '').replace('\t', '').replace('\n', '')
    return restore_opaque_spans(protected, spans)

# End of file #
