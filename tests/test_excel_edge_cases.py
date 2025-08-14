#!/usr/bin/env python3
"""
Extended edge case tests for Excel formula formatter.
File: tests/test_excel_edge_cases.py
"""

import sys

from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter import ExcelFormulaFormatter


def test_array_formulas():
    """Test array formulas with curly braces."""
    formatter = ExcelFormulaFormatter()
    original = '{=SUM(A1:A10*B1:B10)}'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # For array formulas, compare directly after normalizing spaces
    success = original.replace(' ', '') == unfolded.replace(' ', '')
    print(f"Round-trip success: {success}")
    return success


def test_structured_references():
    """Test Excel table structured references.""" 
    formatter = ExcelFormulaFormatter()
    original = '=SUM(Table1[Column1],Table1[@[Column2]])'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    
    success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    print(f"Round-trip success: {success}")
    return success


def test_quoted_sheet_names():
    """Test sheet names with spaces that need quotes."""
    formatter = ExcelFormulaFormatter()
    original = "='My Sheet'!A1+'Another Sheet'!B2"
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    
    success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    print(f"Round-trip success: {success}")
    return success


def test_nested_quotes():
    """Test strings with escaped quotes inside."""
    formatter = ExcelFormulaFormatter()
    original = '=IF(A1>0,"He said ""Hello""","Goodbye")'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")  
    print(f"Unfolded: {unfolded}")
    print()
    
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    
    success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    print(f"Round-trip success: {success}")
    return success


def test_percentages_and_scientific():
    """Test percentage and scientific notation."""
    formatter = ExcelFormulaFormatter()
    original = '=IF(A1>50%,A1*1.5E+10,A1/2.3E-5)'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    
    success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    print(f"Round-trip success: {success}")
    return success


def test_named_ranges():
    """Test named ranges and defined names."""
    formatter = ExcelFormulaFormatter()
    original = '=SUM(MyData,Data.Revenue)+AVERAGE(MyRange)'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    
    success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    print(f"Round-trip success: {success}")
    return success


def test_complex_dynamic_arrays():
    """Test modern Excel dynamic array functions."""
    formatter = ExcelFormulaFormatter()
    original = '=FILTER(SORT(UNIQUE(A1:A100)),LEN(UNIQUE(A1:A100))>3)'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    original_norm = original[1:] if original.startswith('=') else original
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    
    success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    print(f"Round-trip success: {success}")
    return success


def test_very_long_formula():
    """Test handling of extremely long formulas with real nesting."""
    formatter = ExcelFormulaFormatter()
    
    # Build a properly nested formula that will create multiple lines
    nested_formula = "=IF(A1>0,IF(B1>0,IF(C1>0,SUM(A1:A10,B1:B10,C1:C10),AVERAGE(A1:A10)),MAX(B1:B10)),MIN(A1:A10))"
    
    folded = formatter.fold_formula(nested_formula)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original length: {len(nested_formula)} characters")
    print(f"Folded lines: {len(folded.split(chr(10)))}")  # Use chr(10) instead of escaped newline
    print(f"Folded preview:\n{folded[:200]}...")
    print(f"Unfolded: {unfolded[:100]}..." if len(unfolded) > 100 else f"Unfolded: {unfolded}")
    print()
    
    original_norm = nested_formula[1:] if nested_formula.startswith('=') else nested_formula
    unfolded_norm = unfolded[1:] if unfolded.startswith('=') else unfolded
    
    success = original_norm.replace(' ', '') == unfolded_norm.replace(' ', '')
    print(f"Round-trip success: {success}")
    return success


def main():
    """Run all edge case tests."""
    print("Excel Formula Formatter - Extended Edge Case Tests")
    print("=" * 60)
    print()
    
    tests = [
        ("Array Formulas", test_array_formulas),
        ("Structured References", test_structured_references), 
        ("Quoted Sheet Names", test_quoted_sheet_names),
        ("Nested Quotes", test_nested_quotes),
        ("Percentages & Scientific", test_percentages_and_scientific),
        ("Named Ranges", test_named_ranges),
        ("Dynamic Arrays", test_complex_dynamic_arrays),
        ("Very Long Formula", test_very_long_formula)
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
    print(f"Edge Case Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All edge case tests passed!")
        return 0
    else:
        print("❌ Some edge cases failed. These may need additional parser improvements.")
        return 1


if __name__ == "__main__":
    exit(main())

# End of file #
