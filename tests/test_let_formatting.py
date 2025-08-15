#!/usr/bin/env python3
"""
Test LET function formatting functionality.
File: tests/test_let_formatting.py
"""

import sys

from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter import ExcelFormulaFormatter
from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


def test_simple_let_formula():
    """Test basic LET formula with variable name/value pairs on same line."""
    formatter = ExcelFormulaFormatter()
    original = '=LET(x,A1,y,B1,x+y)'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check that variable pairs are on same lines
    lines = folded.split('\n')
    
    # Should have: header comment, x,A1 line, y,B1 line, x+y line
    expected_pattern_found = False
    for line in lines:
        if 'x, "A1"' in line:  # Variable name and value on same line
            expected_pattern_found = True
            break
    
    # Check round-trip
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    round_trip_success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    
    success = expected_pattern_found and round_trip_success
    print(f"Variable pairs on same line: {expected_pattern_found}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_complex_let_formula():
    """Test complex LET formula with multiple variables and nested functions."""
    formatter = ExcelFormulaFormatter()
    original = '=LET(has_Paid_Date,LEN(AM2)>0,has_Price,LEN(AZ2)>=1,has_Invoice_Date,LEN(AE2)>1,has_Title_Date,LEN(AI2)>1,has_WBL_Date,LEN(AJ2)>1,IF(AND(has_Paid_Date,has_Price),1,0))'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check that variable pairs are on same lines
    lines = folded.split('\n')
    variable_pairs_correct = 0
    
    for line in lines:
        # Look for variable name, value patterns
        if 'has_Paid_Date, LEN(' in line:
            variable_pairs_correct += 1
        elif 'has_Price, LEN(' in line:
            variable_pairs_correct += 1
        elif 'has_Invoice_Date, LEN(' in line:
            variable_pairs_correct += 1
    
    # Should find at least a few variable pairs on same lines
    pairs_found = variable_pairs_correct >= 2
    
    # Check round-trip
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    round_trip_success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    
    success = pairs_found and round_trip_success
    print(f"Variable pairs found on same lines: {variable_pairs_correct}")
    print(f"Pairs formatting correct: {pairs_found}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_modular_let_javascript():
    """Test LET function with JavaScript modular formatter."""
    formatter = ModularExcelFormatter.create_javascript_formatter()
    original = '=LET(x,A1+B1,y,C1*D1,x/y)'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check for variable pairs on same lines (no spaces around operators in this context)
    lines = folded.split('\n')
    found_x_pair = any('x, "A1"+"B1"' in line for line in lines)
    found_y_pair = any('y, "C1"*"D1"' in line for line in lines)
    
    pairs_found = found_x_pair and found_y_pair
    
    # Check round-trip
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    round_trip_success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    
    success = pairs_found and round_trip_success
    print(f"Found x pair: {found_x_pair}")
    print(f"Found y pair: {found_y_pair}")
    print(f"Pairs formatting correct: {pairs_found}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_modular_let_python():
    """Test LET function with Python modular formatter."""
    formatter = ModularExcelFormatter.create_python_formatter()
    original = '=LET(name,A1,age,B1,CONCATENATE(name," is ",age," years old"))'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check for variable pairs on same lines (with Python syntax)
    lines = folded.split('\n')
    found_name_pair = any('name, sheet["A1"]' in line for line in lines)
    found_age_pair = any('age, sheet["B1"]' in line for line in lines)
    
    pairs_found = found_name_pair and found_age_pair
    
    # Check round-trip
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    round_trip_success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    
    success = pairs_found and round_trip_success
    print(f"Found name pair: {found_name_pair}")
    print(f"Found age pair: {found_age_pair}")
    print(f"Pairs formatting correct: {pairs_found}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_nested_let_formula():
    """Test LET formula with nested functions in values."""
    formatter = ExcelFormulaFormatter()
    original = '=LET(sum_a,SUM(A1:A10),avg_b,AVERAGE(B1:B10),result,IF(sum_a>avg_b,"High","Low"),result)'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check that complex expressions stay with their variable names
    lines = folded.split('\n')
    found_sum_pair = any('sum_a, SUM(' in line for line in lines)
    found_avg_pair = any('avg_b, AVERAGE(' in line for line in lines)
    
    pairs_found = found_sum_pair and found_avg_pair
    
    # Check round-trip
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    round_trip_success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    
    success = pairs_found and round_trip_success
    print(f"Found sum pair: {found_sum_pair}")
    print(f"Found avg pair: {found_avg_pair}")
    print(f"Pairs formatting correct: {pairs_found}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def main():
    """Run all LET function formatting tests."""
    print("LET Function Formatting Tests")
    print("=" * 50)
    print()
    
    tests = [
        ("Simple LET Formula", test_simple_let_formula),
        ("Complex LET Formula", test_complex_let_formula),
        ("Modular LET (JavaScript)", test_modular_let_javascript),
        ("Modular LET (Python)", test_modular_let_python),
        ("Nested LET Formula", test_nested_let_formula)
    ]
    
    results = []
    for test_name, test_func in tests:
        print(f"Running {test_name} test...")
        print("-" * 40)
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
    print(f"LET Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All LET formatting tests passed!")
        print("Variable name/value pairs are now kept on the same line.")
        return 0
    else:
        print("❌ Some LET tests failed. Check the output above for details.")
        return 1


if __name__ == "__main__":
    exit(main())

# End of file #
