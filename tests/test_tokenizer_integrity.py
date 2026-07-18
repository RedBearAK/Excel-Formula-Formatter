#!/usr/bin/env python3
"""
Tokenizer integrity tests: the corruption cases that motivated the
tokenizer rewrite must round-trip exactly, and malformed input must
raise clear errors instead of silently swallowing content.
File: tests/test_tokenizer_integrity.py
"""

import sys

from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.formula_text_guards import normalize_formula_for_comparison
from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


all_modes = ['j', 'a', 'p', 'c']


def round_trip_all_modes(original: str) -> bool:
    """Fold and unfold a formula in every mode with strict comparison."""
    overall = True

    for mode in all_modes:
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
        unfolded = formatter.unfold_formula(formatter.fold_formula(original))

        success = (
            normalize_formula_for_comparison(original)
            == normalize_formula_for_comparison(unfolded)
        )
        overall = overall and success

        status = "OK  " if success else "FAIL"
        print(f"  [{mode}] {status}  {unfolded}")

    return overall


def test_escaped_quotes_in_strings():
    """Excel's doubled-quote escaping must survive round trips."""
    cases = [
        '=IF(A1>0,"He said ""Hello""","Goodbye")',
        '=SUBSTITUTE(A1,"""","")',
    ]
    overall = True
    for original in cases:
        print(f"Original: {original}")
        overall = round_trip_all_modes(original) and overall
    print(f"Overall success: {overall}")
    return overall


def test_string_interior_preserved():
    """Spacing, commas, parens, and // inside strings must not change."""
    cases = [
        '=IF(A1=1,"a ,b  c","x")',
        '=IF(A1=1,"see http://x.com","y")',
        '=IF(A1=1,"(gap )","z")',
        '=IF(B1="A1","cell name!","other")',
        '=IF(A1<>"x <> y",A1&" & "&B1,"")',
    ]
    overall = True
    for original in cases:
        print(f"Original: {original}")
        overall = round_trip_all_modes(original) and overall
    print(f"Overall success: {overall}")
    return overall


def test_structured_references():
    """Bracketed references, including spaces and nesting, are atomic."""
    cases = [
        '=SUM(Table1[Sales Amount])',
        '=SUM(Table1[[#All],[Col Name]],Table1[@[Col Two]])',
    ]
    overall = True
    for original in cases:
        print(f"Original: {original}")
        overall = round_trip_all_modes(original) and overall
    print(f"Overall success: {overall}")
    return overall


def test_quoted_sheet_names():
    """Single-quoted sheet names with spaces and '' escapes survive."""
    cases = [
        "='Q1 Data'!A1+'It''s Sheet'!B2",
        "='My Sheet'!SomeName+1",
    ]
    overall = True
    for original in cases:
        print(f"Original: {original}")
        overall = round_trip_all_modes(original) and overall
    print(f"Overall success: {overall}")
    return overall


def test_absolute_refs_and_numbers():
    """Absolute references and scientific notation tokenize correctly."""
    cases = [
        '=SUM($A$1:$B$10)+$C2+D$3',
        '=A1*1.5E+10+2.3E-5+.5',
    ]
    overall = True
    for original in cases:
        print(f"Original: {original}")
        overall = round_trip_all_modes(original) and overall
    print(f"Overall success: {overall}")
    return overall


def test_array_constants():
    """Array constants are atomic: inner commas are not arg separators."""
    original = '=SUM({1,2;3,4})+INDEX({"a","b c"},1)'
    print(f"Original: {original}")
    success = round_trip_all_modes(original)
    print(f"Overall success: {success}")
    return success


def test_unlisted_functions_fold():
    """Functions absent from any hardcoded list still fold as functions."""
    formatter = ModularExcelFormatter.create_plain_formatter()
    original = '=IFERROR(TEXTSPLIT(A1,",",";",TRUE),"none")'

    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")

    # A recognized function gets multi-line treatment: args indented
    folded_lines = folded.split('\n')
    got_multiline = len(folded_lines) > 2
    args_indented = any(line.startswith('    ') for line in folded_lines)

    round_trip_success = (
        normalize_formula_for_comparison(original)
        == normalize_formula_for_comparison(unfolded)
    )

    success = got_multiline and args_indented and round_trip_success
    print(f"Multi-line function formatting: {got_multiline}")
    print(f"Arguments indented: {args_indented}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_let_variable_no_cell_collision():
    """All-caps LET variables with digits are not treated as cell refs."""
    original = '=LET(RATE1,0.05,BASE2,100,BASE2*(1+RATE1))'
    print(f"Original: {original}")
    success = round_trip_all_modes(original)

    # In JavaScript mode, RATE1 must NOT be backtick-marked as a cell ref
    js_formatter = ModularExcelFormatter.create_javascript_formatter()
    js_folded = js_formatter.fold_formula(original)
    no_false_cell_marking = '`RATE1`' not in js_folded and '`BASE2`' not in js_folded

    success = success and no_false_cell_marking
    print(f"No false cell marking in JS mode: {no_false_cell_marking}")
    print(f"Overall success: {success}")
    return success


def test_malformed_input_raises():
    """Unclosed strings, brackets, and braces raise ValueError."""
    formatter = ModularExcelFormatter.create_plain_formatter()
    bad_formulas = [
        '=IF(A1="unclosed,B1,C1)',
        '=SUM(Table1[Unclosed',
        '=SUM({1,2;3',
    ]

    overall = True
    for bad in bad_formulas:
        try:
            formatter.fold_formula(bad)
            print(f"  FAIL no error raised: {bad}")
            overall = False
        except ValueError as e:
            print(f"  OK   ValueError: {str(e)[:60]}")

    print(f"Overall success: {overall}")
    return overall


def main():
    """Run all tokenizer integrity tests."""
    print("Tokenizer Integrity Tests")
    print("=" * 60)
    print()

    tests = [
        ("Escaped Quotes in Strings", test_escaped_quotes_in_strings),
        ("String Interior Preserved", test_string_interior_preserved),
        ("Structured References", test_structured_references),
        ("Quoted Sheet Names", test_quoted_sheet_names),
        ("Absolute Refs and Numbers", test_absolute_refs_and_numbers),
        ("Array Constants", test_array_constants),
        ("Unlisted Functions Fold", test_unlisted_functions_fold),
        ("LET Variable No Cell Collision", test_let_variable_no_cell_collision),
        ("Malformed Input Raises", test_malformed_input_raises),
    ]

    results = []
    for test_name, test_func in tests:
        print(f"Running {test_name} test...")
        print("-" * 40)
        try:
            success = test_func()
            results.append(success)
            print(f"{'✓' if success else '✗'} {test_name}: {'PASS' if success else 'FAIL'}")
        except Exception as e:
            print(f"✗ {test_name}: ERROR - {e}")
            results.append(False)
        print()

    passed = sum(results)
    total = len(results)

    print("=" * 60)
    print(f"Tokenizer Integrity Results: {passed}/{total} tests passed")

    if passed == total:
        print("🎉 All tokenizer integrity tests passed!")
        return 0
    else:
        print("❌ Some integrity tests failed. Check the output above for details.")
        return 1


if __name__ == "__main__":
    sys.exit(main())

# End of file #
