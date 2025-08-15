#!/usr/bin/env python3
"""
Test IFS case formatting and natural wrapping functionality.
File: tests/test_ifs_and_wrapping.py
"""

import sys

from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter import ExcelFormulaFormatter
from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


def test_and_natural_wrapping():
    """Test AND function with natural length-based wrapping."""
    formatter = ExcelFormulaFormatter()
    original = '=AND(has_Paid_Date,NOT(has_Price),NOT(has_Invoice_Date),NOT(has_Routing_ETA),NOT(has_Title_Date))'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check for natural wrapping pattern
    lines = folded.split('\n')
    has_natural_wrap = False
    has_proper_spacing = False
    
    for line in lines:
        if 'AND(' in line and 'has_Paid_Date' in line:
            # Should start with 2 spaces after AND(
            if 'AND(  has_Paid_Date' in line:
                has_proper_spacing = True
            # Should wrap naturally when too long
            if len(line) > 60:  # Reasonable length for wrapping
                has_natural_wrap = True
        elif line.strip().endswith(')') and 'NOT(' in line:
            # Continuation line should be indented
            has_natural_wrap = True
    
    # Check round-trip
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    round_trip_success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    
    success = has_proper_spacing and round_trip_success
    print(f"Has proper spacing: {has_proper_spacing}")
    print(f"Has natural wrapping: {has_natural_wrap}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_simple_ifs():
    """Test simple IFS function with case/result pair separators."""
    formatter = ExcelFormulaFormatter()
    original = '=IFS(A1>0,"Positive",A1<0,"Negative",TRUE,"Zero")'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check for case/result pair separators
    lines = folded.split('\n')
    has_pair_separators = any('CASE/RESULT PAIR' in line for line in lines)
    has_blank_lines = '' in lines  # Should have blank lines between cases
    
    # Check round-trip
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    round_trip_success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    
    success = has_pair_separators and has_blank_lines and round_trip_success
    print(f"Has CASE/RESULT PAIR separators: {has_pair_separators}")
    print(f"Has blank lines: {has_blank_lines}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_complex_ifs_with_and():
    """Test complex IFS with AND conditions and natural wrapping."""
    formatter = ExcelFormulaFormatter()
    original = '=IFS(AND(has_Paid_Date,NOT(has_Price),has_Invoice_Date),Invoice_with_Days,AND(has_Paid_Date,NOT(has_Price),NOT(has_Invoice_Date)),Routing_plus_21,TRUE,"")'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check for case structure
    lines = folded.split('\n')
    has_pair_separators = any('CASE/RESULT PAIR' in line for line in lines)
    
    # Check that AND functions use natural wrapping (no generic comments)
    has_logical_and_comment = any('Logical AND' in line for line in lines)
    
    # Check for proper AND spacing
    has_and_spacing = any('AND(  has_Paid_Date' in line for line in lines)
    
    # Check round-trip
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    round_trip_success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    
    success = (has_pair_separators and 
              not has_logical_and_comment and has_and_spacing and round_trip_success)
    
    print(f"Has case/result pair separators: {has_pair_separators}")
    print(f"No 'Logical AND' comments: {not has_logical_and_comment}")
    print(f"Has AND spacing: {has_and_spacing}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_modular_ifs_python():
    """Test IFS with Python modular formatter."""
    formatter = ModularExcelFormatter.create_python_formatter()
    original = '=IFS(A1>0,"High",A1<0,"Low",TRUE,"Medium")'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check for Python-style case/result pair separators (using # comments)
    lines = folded.split('\n')
    has_pair_separators = any('CASE/RESULT PAIR' in line and '#' in line for line in lines)
    
    # Check round-trip
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    round_trip_success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    
    success = has_pair_separators and round_trip_success
    print(f"Has Python case/result pair separators: {has_pair_separators}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def test_switch_function():
    """Test SWITCH function with case/result pair formatting."""
    formatter = ExcelFormulaFormatter()
    original = '=SWITCH(A1,1,"One",2,"Two",3,"Three","Other")'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check for case/result pair separators in SWITCH
    lines = folded.split('\n')
    has_pair_separators = any('CASE/RESULT PAIR' in line for line in lines)
    
    # Check round-trip
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    round_trip_success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    
    success = has_pair_separators and round_trip_success
    print(f"Has CASE/RESULT PAIR separators: {has_pair_separators}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Overall success: {success}")
    return success


def main():
    """Run all IFS and natural wrapping tests."""
    print("IFS Case/Result Pair Formatting and Natural Wrapping Tests")
    print("=" * 60)
    print()
    
    tests = [
        ("AND Natural Wrapping", test_and_natural_wrapping),
        ("Simple IFS Cases", test_simple_ifs),
        ("Complex IFS with AND", test_complex_ifs_with_and),
        ("Modular IFS (Python)", test_modular_ifs_python),
        ("SWITCH Function", test_switch_function)
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
    
    print("=" * 60)
    print(f"Enhancement Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All enhancement tests passed!")
        print("✨ IFS case/result pair formatting and natural wrapping are working correctly.")
        return 0
    else:
        print("❌ Some enhancement tests failed. Check the output above for details.")
        return 1


if __name__ == "__main__":
    exit(main())

# End of file #
