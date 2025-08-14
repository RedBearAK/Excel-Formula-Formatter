#!/usr/bin/env python3
"""
Test module for Excel formula formatter round-trip transformations.
File: test_excel_formatter.py
"""

import sys

from pathlib import Path

# Add the parent directory (containing the package) to Python path
# This goes up from tests/ to the parent directory containing excel_formula_formatter/
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.excel_formula_formatter import ExcelFormulaFormatter


def test_simple_sum():
    """Test basic SUM formula round-trip."""
    formatter = ExcelFormulaFormatter()
    original = "=SUM(A1:A10)"
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Normalize for comparison (remove leading =)
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    
    success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    print(f"Round-trip success: {success}")
    return success


def test_complex_if_formula():
    """Test complex IF formula with cell references and operators."""
    formatter = ExcelFormulaFormatter()
    original = '=IF(A1>B$2,SUM(A1:A10)*Sheet1!C1,"")'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Normalize for comparison
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    
    success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    print(f"Round-trip success: {success}")
    return success


def test_not_equal_operator():
    """Test Excel <> operator conversion."""
    formatter = ExcelFormulaFormatter()
    original = '=IF(A1<>B1,"Different","Same")'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check that folded version contains != 
    has_js_operator = '!=' in folded
    # Check that unfolded version contains <>
    has_excel_operator = '<>' in unfolded
    
    # Normalize for comparison
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    
    round_trip_success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    success = has_js_operator and has_excel_operator and round_trip_success
    
    print(f"Has != in folded: {has_js_operator}")
    print(f"Has <> in unfolded: {has_excel_operator}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_nested_functions():
    """Test deeply nested function calls."""
    formatter = ExcelFormulaFormatter()
    original = '=SUM(IF(ISERROR(VLOOKUP(A1:A10,B:C,2,FALSE)),0,VLOOKUP(A1:A10,B:C,2,FALSE)))'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Normalize for comparison
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    
    success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    print(f"Round-trip success: {success}")
    return success


def test_string_literals_with_commas():
    """Test string literals containing commas and spaces."""
    formatter = ExcelFormulaFormatter()
    original = '=CONCATENATE("Hello, World!",", How are you?",IF(A1>0," Good"," Bad"))'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Normalize for comparison
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    
    success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    print(f"Round-trip success: {success}")
    return success


def test_empty_and_edge_cases():
    """Test edge cases like empty input."""
    formatter = ExcelFormulaFormatter()
    
    # Test empty input
    empty_result = formatter.fold_formula("")
    empty_unfolded = formatter.unfold_formula(empty_result)
    
    # Test just equals sign
    equals_result = formatter.fold_formula("=")
    equals_unfolded = formatter.unfold_formula(equals_result)
    
    # Test whitespace only
    space_result = formatter.fold_formula("   ")
    space_unfolded = formatter.unfold_formula(space_result)
    
    print("Edge case results:")
    print(f"Empty: '{empty_result}' → '{empty_unfolded}'")
    print(f"Equals: '{equals_result}' → '{equals_unfolded}'")
    print(f"Spaces: '{space_result}' → '{space_unfolded}'")
    print()
    
    success = (empty_result == "" and empty_unfolded == "" and
               equals_unfolded in ["", "="] and
               space_result == "" and space_unfolded == "")
    
    print(f"Edge cases success: {success}")
    return success


def main():
    """Run all tests and report results."""
    print("Excel Formula Formatter Round-Trip Tests")
    print("=" * 50)
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
        print("-" * 30)
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
    
    print("=" * 50)
    print(f"Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All tests passed! Round-trip transformation is working correctly.")
        return 0
    else:
        print("❌ Some tests failed. Check the output above for details.")
        return 1


if __name__ == "__main__":
    exit(main())

# End of file #
