#!/usr/bin/env python3
"""
LET function formatting tests for the modular Excel formatter.
File: tests/test_let_formatting.py

Ported to the modular formatter after the legacy formatter was removed.
Verifies that LET variable name/value pairs stay together on one line.
"""

import sys

from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.formula_text_guards import normalize_formula_for_comparison
from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


def normalize_formula(formula: str) -> str:
    """Strict normalization: spaces inside strings, sheet quotes, and
    brackets are preserved so corruption cannot hide."""
    return normalize_formula_for_comparison(formula)


def test_simple_let_formula():
    """Test basic LET formula with variable name/value pairs on same line."""
    formatter = ModularExcelFormatter.create_annotated_formatter()
    original = '=LET(x,A1,y,B1,x+y)'

    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)

    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()

    # Variable name and value should share a line in annotated mode
    lines = folded.split('\n')
    expected_pattern_found = any('x, A1' in line for line in lines)

    round_trip_success = normalize_formula(original) == normalize_formula(unfolded)

    success = expected_pattern_found and round_trip_success
    print(f"Variable pairs on same line: {expected_pattern_found}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_complex_let_formula():
    """Test complex LET formula with multiple variables and nested functions."""
    formatter = ModularExcelFormatter.create_annotated_formatter()
    original = '=LET(has_Paid_Date,LEN(AM2)>0,has_Price,LEN(AZ2)>=1,has_Invoice_Date,LEN(AE2)>1,has_Title_Date,LEN(AI2)>1,has_WBL_Date,LEN(AJ2)>1,IF(AND(has_Paid_Date,has_Price),1,0))'

    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)

    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()

    # Each boolean variable should stay paired with its LEN() expression
    lines = folded.split('\n')
    variable_pairs_correct = sum(
        1 for line in lines
        if 'has_Paid_Date, LEN(' in line
        or 'has_Price, LEN(' in line
        or 'has_Invoice_Date, LEN(' in line
        or 'has_Title_Date, LEN(' in line
        or 'has_WBL_Date, LEN(' in line
    )
    pairs_found = variable_pairs_correct == 5

    round_trip_success = normalize_formula(original) == normalize_formula(unfolded)

    success = pairs_found and round_trip_success
    print(f"Variable pairs found: {variable_pairs_correct}/5")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_let_javascript_mode():
    """Test LET function with JavaScript modular formatter."""
    formatter = ModularExcelFormatter.create_javascript_formatter()
    original = '=LET(x,A1+B1,y,C1*D1,x/y)'

    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)

    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()

    # JavaScript mode backtick-marks cell references but keeps pairs together
    lines = folded.split('\n')
    found_x_pair = any('x, `A1` + `B1`' in line for line in lines)
    found_y_pair = any('y, `C1` * `D1`' in line for line in lines)
    pairs_found = found_x_pair and found_y_pair

    round_trip_success = normalize_formula(original) == normalize_formula(unfolded)

    success = pairs_found and round_trip_success
    print(f"Found x pair: {found_x_pair}")
    print(f"Found y pair: {found_y_pair}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_nested_let_formula():
    """Test LET formula with nested functions in values."""
    formatter = ModularExcelFormatter.create_annotated_formatter()
    original = '=LET(sum_a,SUM(A1:A10),avg_b,AVERAGE(B1:B10),result,IF(sum_a>avg_b,"High","Low"),result)'

    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)

    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()

    # Complex expressions should stay with their variable names
    lines = folded.split('\n')
    found_sum_pair = any('sum_a, SUM(' in line for line in lines)
    found_avg_pair = any('avg_b, AVERAGE(' in line for line in lines)
    pairs_found = found_sum_pair and found_avg_pair

    round_trip_success = normalize_formula(original) == normalize_formula(unfolded)

    success = pairs_found and round_trip_success
    print(f"Found sum_a pair: {found_sum_pair}")
    print(f"Found avg_b pair: {found_avg_pair}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def main():
    """Run all LET formatting tests."""
    print("LET Function Formatting Tests (Modular)")
    print("=" * 60)
    print()

    tests = [
        ("Simple LET Formula", test_simple_let_formula),
        ("Complex LET Formula", test_complex_let_formula),
        ("LET JavaScript Mode", test_let_javascript_mode),
        ("Nested LET Formula", test_nested_let_formula),
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
    print(f"LET Formatting Test Results: {passed}/{total} tests passed")

    if passed == total:
        print("🎉 All LET formatting tests passed!")
        return 0
    else:
        print("❌ Some LET formatting tests failed. Check the output above for details.")
        return 1


if __name__ == "__main__":
    sys.exit(main())

# End of file #
