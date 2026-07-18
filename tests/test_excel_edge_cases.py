#!/usr/bin/env python3
"""
Edge-case round-trip tests for the modular Excel formatter.
File: tests/test_excel_edge_cases.py

Ported to the modular formatter after the legacy formatter was removed.
Each edge case is round-tripped in Plain (p) and JavaScript (j) modes.
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


def round_trip_check(original: str) -> bool:
    """Round-trip a formula through Plain and JavaScript modes."""
    overall = True

    for mode in ['p', 'j']:
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
        folded = formatter.fold_formula(original)
        unfolded = formatter.unfold_formula(folded)

        # Shared normalizer strips {= and } array markers itself
        orig_cmp = normalize_formula(original)
        rt_cmp = normalize_formula(unfolded)

        success = orig_cmp == rt_cmp
        overall = overall and success

        status = "OK  " if success else "FAIL"
        print(f"  [{mode}] {status}  unfolded: {unfolded}")
        if not success:
            print(f"        folded was:\n{folded}")

    return overall


def test_array_formulas():
    """Test array formulas with curly braces."""
    original = '{=SUM(A1:A10*B1:B10)}'
    print(f"Original: {original}")
    success = round_trip_check(original)
    print(f"Round-trip success: {success}")
    return success


def test_structured_references():
    """Test Excel table structured references."""
    original = '=SUM(Table1[Column1],Table1[@[Column2]])'
    print(f"Original: {original}")
    success = round_trip_check(original)
    print(f"Round-trip success: {success}")
    return success


def test_quoted_sheet_names():
    """Test sheet names with spaces that need quotes."""
    original = "='My Sheet'!A1+'Another Sheet'!B2"
    print(f"Original: {original}")
    success = round_trip_check(original)
    print(f"Round-trip success: {success}")
    return success


def test_nested_quotes():
    """Test strings with escaped quotes inside."""
    original = '=IF(A1>0,"He said ""Hello""","Goodbye")'
    print(f"Original: {original}")
    success = round_trip_check(original)
    print(f"Round-trip success: {success}")
    return success


def test_percentages_and_scientific():
    """Test percentage and scientific notation."""
    original = '=IF(A1>50%,A1*1.5E+10,A1/2.3E-5)'
    print(f"Original: {original}")
    success = round_trip_check(original)
    print(f"Round-trip success: {success}")
    return success


def test_named_ranges():
    """Test named ranges and defined names."""
    original = '=SUM(MyData,Data.Revenue)+AVERAGE(MyRange)'
    print(f"Original: {original}")
    success = round_trip_check(original)
    print(f"Round-trip success: {success}")
    return success


def test_complex_dynamic_arrays():
    """Test modern Excel dynamic array functions."""
    original = '=FILTER(SORT(UNIQUE(A1:A100)),LEN(UNIQUE(A1:A100))>3)'
    print(f"Original: {original}")
    success = round_trip_check(original)
    print(f"Round-trip success: {success}")
    return success


def test_very_long_formula():
    """Test handling of deeply nested formulas."""
    original = '=IF(A1>0,IF(B1>0,IF(C1>0,SUM(A1:A10,B1:B10,C1:C10),AVERAGE(A1:A10)),MAX(B1:B10)),MIN(A1:A10))'
    print(f"Original: {original}")
    success = round_trip_check(original)
    print(f"Round-trip success: {success}")
    return success


def main():
    """Run all edge case tests."""
    print("Excel Formatter Edge Case Tests (Modular)")
    print("=" * 60)
    print()

    tests = [
        ("Array Formulas", test_array_formulas),
        ("Structured References", test_structured_references),
        ("Quoted Sheet Names", test_quoted_sheet_names),
        ("Nested Quotes", test_nested_quotes),
        ("Percentages and Scientific", test_percentages_and_scientific),
        ("Named Ranges", test_named_ranges),
        ("Complex Dynamic Arrays", test_complex_dynamic_arrays),
        ("Very Long Formula", test_very_long_formula),
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
    print(f"Edge Case Test Results: {passed}/{total} tests passed")

    if passed == total:
        print("🎉 All edge case tests passed!")
        return 0
    else:
        print("❌ Some edge case tests failed. Check the output above for details.")
        return 1


if __name__ == "__main__":
    sys.exit(main())

# End of file #
