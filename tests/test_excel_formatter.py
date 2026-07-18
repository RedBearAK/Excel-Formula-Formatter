#!/usr/bin/env python3
"""
Core round-trip tests for the modular Excel formula formatter.
File: tests/test_excel_formatter.py

Each formula is folded and unfolded in all four modes (j, a, p, c) and
must survive with no loss of syntax. This replaced the old legacy-vs-
modular comparison tests when the legacy formatter was removed.
"""

import sys

from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.formula_text_guards import normalize_formula_for_comparison
from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


all_modes = ['j', 'a', 'p', 'c']


def normalize_formula(formula: str) -> str:
    """Strict normalization: spaces inside strings, sheet quotes, and
    brackets are preserved so corruption cannot hide."""
    return normalize_formula_for_comparison(formula)


def round_trip_all_modes(original: str) -> bool:
    """Fold and unfold a formula in every mode, reporting each result."""
    overall = True

    for mode in all_modes:
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
        folded = formatter.fold_formula(original)
        unfolded = formatter.unfold_formula(folded)

        success = normalize_formula(original) == normalize_formula(unfolded)
        overall = overall and success

        status = "OK  " if success else "FAIL"
        print(f"  [{mode}] {status}  unfolded: {unfolded}")
        if not success:
            print(f"        folded was:\n{folded}")

    return overall


def test_simple_sum():
    """Test basic SUM formula round-trip in all modes."""
    original = "=SUM(A1:A10)"
    print(f"Original: {original}")
    success = round_trip_all_modes(original)
    print(f"Overall success: {success}")
    return success


def test_complex_if_formula():
    """Test complex IF formula with cell references and operators."""
    original = '=IF(A1>10,SUM(B1:B5)*2,AVERAGE(C1:C10)/3)'
    print(f"Original: {original}")
    success = round_trip_all_modes(original)
    print(f"Overall success: {success}")
    return success


def test_nested_functions():
    """Test deeply nested function calls."""
    original = '=IF(AND(A1>0,B1<100),VLOOKUP(C1,Sheet2!D:E,2,FALSE),"Not found")'
    print(f"Original: {original}")
    success = round_trip_all_modes(original)
    print(f"Overall success: {success}")
    return success


def test_string_literals_with_operators():
    """Test that operators inside string literals are preserved."""
    original = '=IF(A1<>"",CONCATENATE("Value: ",A1," <> ",B1),"Empty")'
    print(f"Original: {original}")
    success = round_trip_all_modes(original)
    print(f"Overall success: {success}")
    return success


def test_let_with_ifs():
    """Test LET wrapping an IFS, the core business-formula pattern."""
    original = (
        '=LET(has_Date,LEN(A2)>0,days,N2,'
        'IFS(AND(has_Date,days>0),A2+days,TRUE,""))'
    )
    print(f"Original: {original}")
    success = round_trip_all_modes(original)
    print(f"Overall success: {success}")
    return success


def test_multiple_fold_cycles():
    """Test that repeated fold/unfold cycles do not degrade the formula."""
    original = '=IFS(A1>0,"Positive",A1<0,"Negative",TRUE,"Zero")'
    print(f"Original: {original}")
    overall = True

    for mode in all_modes:
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
        current = original
        for _cycle in range(3):
            current = formatter.unfold_formula(formatter.fold_formula(current))

        success = normalize_formula(original) == normalize_formula(current)
        overall = overall and success
        status = "OK  " if success else "FAIL"
        print(f"  [{mode}] {status}  after 3 cycles: {current}")

    print(f"Overall success: {overall}")
    return overall


def main():
    """Run all core formatter tests."""
    print("Modular Excel Formatter Core Round-Trip Tests")
    print("=" * 60)
    print()

    tests = [
        ("Simple SUM", test_simple_sum),
        ("Complex IF", test_complex_if_formula),
        ("Nested Functions", test_nested_functions),
        ("String Literals with Operators", test_string_literals_with_operators),
        ("LET with IFS", test_let_with_ifs),
        ("Multiple Fold Cycles", test_multiple_fold_cycles),
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
    print(f"Core Test Results: {passed}/{total} tests passed")

    if passed == total:
        print("🎉 All core round-trip tests passed in all four modes!")
        return 0
    else:
        print("❌ Some core tests failed. Check the output above for details.")
        return 1


if __name__ == "__main__":
    sys.exit(main())

# End of file #
