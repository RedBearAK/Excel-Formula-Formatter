#!/usr/bin/env python3
"""
Test IFS case formatting and multi-argument function wrapping.
File: tests/test_ifs_and_wrapping.py

Ported to the modular formatter after the legacy formatter was removed.
Annotated mode (a) carries the CASE/RESULT PAIR separators; Plain mode
(p) must produce no comments at all.
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


def test_and_argument_wrapping():
    """Test AND function with one argument per line in annotated mode."""
    formatter = ModularExcelFormatter.create_annotated_formatter()
    original = '=AND(has_Paid_Date,NOT(has_Price),NOT(has_Invoice_Date),NOT(has_Routing_ETA),NOT(has_Title_Date))'

    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)

    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()

    # Multi-argument AND should place each argument on its own indented line,
    # while simple NOT(x) calls stay inline on their argument lines.
    lines = folded.split('\n')
    not_arg_lines = [ln for ln in lines if ln.strip().startswith('NOT(')]
    args_on_own_lines = len(not_arg_lines) == 4
    args_indented = all(ln.startswith('    ') for ln in not_arg_lines)

    round_trip_success = normalize_formula(original) == normalize_formula(unfolded)

    success = args_on_own_lines and args_indented and round_trip_success
    print(f"NOT arguments each on own line: {args_on_own_lines}")
    print(f"Argument lines indented: {args_indented}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_simple_ifs():
    """Test simple IFS function with case/result pair separators."""
    formatter = ModularExcelFormatter.create_annotated_formatter()
    original = '=IFS(A1>0,"Positive",A1<0,"Negative",TRUE,"Zero")'

    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)

    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()

    lines = folded.split('\n')
    has_pair_separators = any('CASE/RESULT PAIR' in line for line in lines)
    has_blank_lines = '' in lines  # Blank lines between cases

    round_trip_success = normalize_formula(original) == normalize_formula(unfolded)

    success = has_pair_separators and has_blank_lines and round_trip_success
    print(f"Has CASE/RESULT PAIR separators: {has_pair_separators}")
    print(f"Has blank lines: {has_blank_lines}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_complex_ifs_with_and():
    """Test complex IFS with AND conditions in annotated mode."""
    formatter = ModularExcelFormatter.create_annotated_formatter()
    original = '=IFS(AND(has_Paid_Date,NOT(has_Price),has_Invoice_Date),Invoice_with_Days,AND(has_Paid_Date,NOT(has_Price),NOT(has_Invoice_Date)),Routing_plus_21,TRUE,"")'

    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)

    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()

    lines = folded.split('\n')
    has_pair_separators = any('CASE/RESULT PAIR' in line for line in lines)

    # AND functions should not emit their own generic comments
    has_logical_and_comment = any('Logical AND' in line for line in lines)

    round_trip_success = normalize_formula(original) == normalize_formula(unfolded)

    success = has_pair_separators and not has_logical_and_comment and round_trip_success
    print(f"Has case/result pair separators: {has_pair_separators}")
    print(f"No 'Logical AND' comments: {not has_logical_and_comment}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_ifs_plain_mode_no_comments():
    """Test IFS in Plain mode: smart indenting with zero comments."""
    formatter = ModularExcelFormatter.create_plain_formatter()
    original = '=IFS(A1>0,"High",A1<0,"Low",TRUE,"Medium")'

    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)

    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()

    # Plain mode must not contain any comment lines
    lines = folded.split('\n')
    has_no_comments = not any('//' in line for line in lines)

    # Arguments should still be broken onto separate indented lines
    indented_lines = [ln for ln in lines if ln.startswith('    ')]
    has_indented_args = len(indented_lines) >= 6  # 3 cases x 2 (condition, result)

    round_trip_success = normalize_formula(original) == normalize_formula(unfolded)

    success = has_no_comments and has_indented_args and round_trip_success
    print(f"Plain mode has no comments: {has_no_comments}")
    print(f"Arguments on indented lines: {has_indented_args}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_switch_function():
    """Test SWITCH function with case/result pair formatting."""
    formatter = ModularExcelFormatter.create_annotated_formatter()
    original = '=SWITCH(A1,1,"One",2,"Two",3,"Three","Other")'

    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)

    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()

    lines = folded.split('\n')
    has_pair_separators = any('CASE/RESULT PAIR' in line for line in lines)

    round_trip_success = normalize_formula(original) == normalize_formula(unfolded)

    success = has_pair_separators and round_trip_success
    print(f"Has CASE/RESULT PAIR separators: {has_pair_separators}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def main():
    """Run all IFS and wrapping tests."""
    print("IFS Case/Result Pair Formatting and Wrapping Tests")
    print("=" * 60)
    print()

    tests = [
        ("AND Argument Wrapping", test_and_argument_wrapping),
        ("Simple IFS Cases", test_simple_ifs),
        ("Complex IFS with AND", test_complex_ifs_with_and),
        ("IFS Plain Mode (No Comments)", test_ifs_plain_mode_no_comments),
        ("SWITCH Function", test_switch_function),
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
    print(f"IFS/Wrapping Test Results: {passed}/{total} tests passed")

    if passed == total:
        print("🎉 All IFS and wrapping tests passed!")
        return 0
    else:
        print("❌ Some tests failed. Check the output above for details.")
        return 1


if __name__ == "__main__":
    sys.exit(main())

# End of file #
