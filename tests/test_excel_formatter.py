#!/usr/bin/env python3
"""
Test module for Excel formula formatter round-trip transformations.
Tests both original and modular formatters for compatibility.
File: test_excel_formatter.py
"""

import sys

from pathlib import Path

# Add the parent directory (containing the package) to Python path
# This goes up from tests/ to the parent directory containing excel_formula_formatter/
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.excel_formula_formatter import ExcelFormulaFormatter
from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


def normalize_formula(formula: str) -> str:
    """Normalize Excel formula for comparison by removing leading = and spaces."""
    normalized = formula.strip()
    if normalized.startswith('='):
        normalized = normalized[1:]
    return normalized.replace(' ', '')


def test_simple_sum():
    """Test basic SUM formula round-trip with both formatters."""
    original = "=SUM(A1:A10)"
    
    print("Testing with Original Formatter:")
    original_formatter = ExcelFormulaFormatter()
    
    folded_orig = original_formatter.fold_formula(original)
    unfolded_orig = original_formatter.unfold_formula(folded_orig)
    
    print(f"Original: {original}")
    print(f"Folded (Original):\n{folded_orig}")
    print(f"Unfolded (Original): {unfolded_orig}")
    
    original_success = normalize_formula(original) == normalize_formula(unfolded_orig)
    print(f"Original formatter success: {original_success}")
    print()
    
    print("Testing with Modular Formatter:")
    modular_formatter = ModularExcelFormatter.create_javascript_formatter()
    
    folded_mod = modular_formatter.fold_formula(original)
    unfolded_mod = modular_formatter.unfold_formula(folded_mod)
    
    print(f"Folded (Modular):\n{folded_mod}")
    print(f"Unfolded (Modular): {unfolded_mod}")
    
    modular_success = normalize_formula(original) == normalize_formula(unfolded_mod)
    print(f"Modular formatter success: {modular_success}")
    print()
    
    success = original_success and modular_success
    print(f"Overall success: {success}")
    return success


def test_complex_if_formula():
    """Test complex IF formula with cell references and operators."""
    original = '=IF(A1>B$2,SUM(A1:A10)*Sheet1!C1,"")'
    
    print("Testing with Original Formatter:")
    original_formatter = ExcelFormulaFormatter()
    
    folded_orig = original_formatter.fold_formula(original)
    unfolded_orig = original_formatter.unfold_formula(folded_orig)
    
    print(f"Original: {original}")
    print(f"Folded (Original):\n{folded_orig}")
    print(f"Unfolded (Original): {unfolded_orig}")
    
    original_success = normalize_formula(original) == normalize_formula(unfolded_orig)
    print(f"Original formatter success: {original_success}")
    print()
    
    print("Testing with Modular Formatter:")
    modular_formatter = ModularExcelFormatter.create_javascript_formatter()
    
    folded_mod = modular_formatter.fold_formula(original)
    unfolded_mod = modular_formatter.unfold_formula(folded_mod)
    
    print(f"Folded (Modular):\n{folded_mod}")
    print(f"Unfolded (Modular): {unfolded_mod}")
    
    modular_success = normalize_formula(original) == normalize_formula(unfolded_mod)
    print(f"Modular formatter success: {modular_success}")
    print()
    
    success = original_success and modular_success
    print(f"Overall success: {success}")
    return success


def test_not_equal_operator():
    """Test Excel <> operator conversion with both formatters."""
    original = '=IF(A1<>B1,"Different","Same")'
    
    print("Testing with Original Formatter:")
    original_formatter = ExcelFormulaFormatter()
    
    folded_orig = original_formatter.fold_formula(original)
    unfolded_orig = original_formatter.unfold_formula(folded_orig)
    
    print(f"Original: {original}")
    print(f"Folded (Original):\n{folded_orig}")
    print(f"Unfolded (Original): {unfolded_orig}")
    
    # Check that folded version contains != 
    has_js_operator_orig = '!=' in folded_orig
    # Check that unfolded version contains <>
    has_excel_operator_orig = '<>' in unfolded_orig
    
    round_trip_success_orig = normalize_formula(original) == normalize_formula(unfolded_orig)
    original_success = has_js_operator_orig and has_excel_operator_orig and round_trip_success_orig
    
    print(f"Has != in folded: {has_js_operator_orig}")
    print(f"Has <> in unfolded: {has_excel_operator_orig}")
    print(f"Round-trip success: {round_trip_success_orig}")
    print(f"Original formatter success: {original_success}")
    print()
    
    print("Testing with Modular Formatter:")
    modular_formatter = ModularExcelFormatter.create_javascript_formatter()
    
    folded_mod = modular_formatter.fold_formula(original)
    unfolded_mod = modular_formatter.unfold_formula(folded_mod)
    
    print(f"Folded (Modular):\n{folded_mod}")
    print(f"Unfolded (Modular): {unfolded_mod}")
    
    has_js_operator_mod = '!=' in folded_mod
    has_excel_operator_mod = '<>' in unfolded_mod
    
    round_trip_success_mod = normalize_formula(original) == normalize_formula(unfolded_mod)
    modular_success = has_js_operator_mod and has_excel_operator_mod and round_trip_success_mod
    
    print(f"Has != in folded: {has_js_operator_mod}")
    print(f"Has <> in unfolded: {has_excel_operator_mod}")
    print(f"Round-trip success: {round_trip_success_mod}")
    print(f"Modular formatter success: {modular_success}")
    print()
    
    success = original_success and modular_success
    print(f"Overall success: {success}")
    return success


def test_nested_functions():
    """Test deeply nested function calls with both formatters."""
    original = '=SUM(IF(ISERROR(VLOOKUP(A1:A10,B:C,2,FALSE)),0,VLOOKUP(A1:A10,B:C,2,FALSE)))'
    
    print("Testing with Original Formatter:")
    original_formatter = ExcelFormulaFormatter()
    
    folded_orig = original_formatter.fold_formula(original)
    unfolded_orig = original_formatter.unfold_formula(folded_orig)
    
    print(f"Original: {original}")
    print(f"Folded lines (Original): {len(folded_orig.split(chr(10)))}")
    print(f"Unfolded (Original): {unfolded_orig}")
    
    original_success = normalize_formula(original) == normalize_formula(unfolded_orig)
    print(f"Original formatter success: {original_success}")
    print()
    
    print("Testing with Modular Formatter:")
    modular_formatter = ModularExcelFormatter.create_javascript_formatter()
    
    folded_mod = modular_formatter.fold_formula(original)
    unfolded_mod = modular_formatter.unfold_formula(folded_mod)
    
    print(f"Folded lines (Modular): {len(folded_mod.split(chr(10)))}")
    print(f"Unfolded (Modular): {unfolded_mod}")
    
    modular_success = normalize_formula(original) == normalize_formula(unfolded_mod)
    print(f"Modular formatter success: {modular_success}")
    print()
    
    success = original_success and modular_success
    print(f"Overall success: {success}")
    return success


def test_string_literals_with_commas():
    """Test string literals containing commas and spaces."""
    original = '=CONCATENATE("Hello, World!",", How are you?",IF(A1>0," Good"," Bad"))'
    
    print("Testing with both formatters (abbreviated output):")
    
    # Test original
    original_formatter = ExcelFormulaFormatter()
    folded_orig = original_formatter.fold_formula(original)
    unfolded_orig = original_formatter.unfold_formula(folded_orig)
    original_success = normalize_formula(original) == normalize_formula(unfolded_orig)
    
    # Test modular
    modular_formatter = ModularExcelFormatter.create_javascript_formatter()
    folded_mod = modular_formatter.fold_formula(original)
    unfolded_mod = modular_formatter.unfold_formula(folded_mod)
    modular_success = normalize_formula(original) == normalize_formula(unfolded_mod)
    
    print(f"Original: {original}")
    print(f"Original formatter success: {original_success}")
    print(f"Modular formatter success: {modular_success}")
    print()
    
    success = original_success and modular_success
    print(f"Overall success: {success}")
    return success


def test_empty_and_edge_cases():
    """Test edge cases like empty input with both formatters."""
    print("Testing edge cases with both formatters:")
    
    original_formatter = ExcelFormulaFormatter()
    modular_formatter = ModularExcelFormatter.create_javascript_formatter()
    
    # Test empty input
    empty_result_orig = original_formatter.fold_formula("")
    empty_unfolded_orig = original_formatter.unfold_formula(empty_result_orig)
    
    empty_result_mod = modular_formatter.fold_formula("")
    empty_unfolded_mod = modular_formatter.unfold_formula(empty_result_mod)
    
    # Test just equals sign
    equals_result_orig = original_formatter.fold_formula("=")
    equals_unfolded_orig = original_formatter.unfold_formula(equals_result_orig)
    
    equals_result_mod = modular_formatter.fold_formula("=")
    equals_unfolded_mod = modular_formatter.unfold_formula(equals_result_mod)
    
    print("Edge case results:")
    print(f"Empty (Original): '{empty_result_orig}' → '{empty_unfolded_orig}'")
    print(f"Empty (Modular): '{empty_result_mod}' → '{empty_unfolded_mod}'")
    print(f"Equals (Original): '{equals_result_orig}' → '{equals_unfolded_orig}'")
    print(f"Equals (Modular): '{equals_result_mod}' → '{equals_unfolded_mod}'")
    print()
    
    original_success = (empty_result_orig == "" and empty_unfolded_orig == "" and
                       equals_unfolded_orig in ["", "="])
    modular_success = (empty_result_mod == "" and empty_unfolded_mod == "" and
                      equals_unfolded_mod in ["", "="])
    
    success = original_success and modular_success
    print(f"Original formatter edge cases: {original_success}")
    print(f"Modular formatter edge cases: {modular_success}")
    print(f"Overall success: {success}")
    return success


def main():
    """Run all tests and report results."""
    print("Excel Formula Formatter Round-Trip Tests (Original vs Modular)")
    print("=" * 70)
    print()
    
    tests = [
        ("Simple SUM", test_simple_sum),
        ("Complex IF", test_complex_if_formula),
        ("Not Equal Operator", test_not_equal_operator),
        ("Nested Functions", test_nested_functions),
        ("String Literals", test_string_literals_with_commas),
        ("Edge Cases", test_empty_and_edge_cases)
    ]
    
    results = []
    for test_name, test_func in tests:
        print(f"Running {test_name} test...")
        print("-" * 50)
        try:
            success = test_func()
            results.append(success)
            print(f"✓ {test_name}: {'PASS' if success else 'FAIL'}")
        except Exception as e:
            print(f"✗ {test_name}: ERROR - {e}")
            results.append(False)
        print()
    
    # Final summary
    passed = sum(results)
    total = len(results)
    
    print("=" * 70)
    print(f"Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All tests passed! Both formatters working correctly.")
        print("✅ Original and Modular formatters produce consistent results.")
        return 0
    else:
        print("❌ Some tests failed. Check the output above for details.")
        print("⚠️  Consider investigating formatter differences.")
        return 1


if __name__ == "__main__":
    exit(main())

# End of file #
