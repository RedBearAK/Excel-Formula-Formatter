#!/usr/bin/env python3
"""
Test spacing in unfolded formulas for all three modes.
File: tests/test_spacing_fix.py
"""

import sys

from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


def test_spacing_in_modes():
    """Test that unfolded formulas have appropriate spacing in each mode."""
    original = '=IF(A1<>B1,SUM(A1:A10),"Equal")'
    
    print(f"Original formula: {original}")
    print("=" * 50)
    
    modes = {
        'j': 'JavaScript',
        'a': 'Annotated Excel', 
        'p': 'Plain Excel'
    }
    
    for mode_code, mode_name in modes.items():
        print(f"\n{mode_name} Mode ({mode_code}):")
        print("-" * 30)
        
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode_code)
        
        # Fold then unfold
        folded = formatter.fold_formula(original)
        unfolded = formatter.unfold_formula(folded)
        
        print(f"Folded preview (first 60 chars): {folded[:60]}...")
        print(f"Unfolded: {unfolded}")
        
        # Check spacing characteristics
        has_spaces_around_operators = ' <> ' in unfolded or ' > ' in unfolded
        has_spaces_after_commas = ', ' in unfolded
        has_clean_parentheses = '( ' not in unfolded and ' )' not in unfolded
        
        print(f"  Spaces around operators: {has_spaces_around_operators}")
        print(f"  Spaces after commas: {has_spaces_after_commas}")
        print(f"  Clean parentheses: {has_clean_parentheses}")
        
        # Round trip test
        original_clean = original.replace(' ', '').replace('=', '')
        unfolded_clean = unfolded.replace(' ', '').replace('=', '')
        round_trip_success = original_clean == unfolded_clean
        
        print(f"  Round trip success: {round_trip_success}")
        
        if not round_trip_success:
            print(f"  Original clean: {original_clean}")
            print(f"  Unfolded clean: {unfolded_clean}")


def test_complex_formula_spacing():
    """Test spacing with a more complex formula."""
    original = '=LET(x,A1+B1,y,C1*D1,IF(x>y,"X wins",IF(y>x,"Y wins","Tie")))'
    
    print(f"\nComplex formula test: {original}")
    print("=" * 60)
    
    for mode_code in ['j', 'a', 'p']:
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode_code)
        
        folded = formatter.fold_formula(original)
        unfolded = formatter.unfold_formula(folded)
        
        print(f"\nMode {mode_code} unfolded:")
        print(f"  {unfolded}")
        
        # Check if it has reasonable spacing
        has_comma_spacing = ', ' in unfolded
        print(f"  Has comma spacing: {has_comma_spacing}")


def main():
    """Run spacing tests."""
    print("Excel Formula Formatter - Spacing Test")
    print("Testing that unfolded formulas have appropriate spacing")
    print()
    
    try:
        test_spacing_in_modes()
        test_complex_formula_spacing()
        
        print("\n" + "=" * 60)
        print("✅ Spacing test completed successfully!")
        print("Check the output above to verify spacing looks correct.")
        return 0
        
    except Exception as e:
        print(f"❌ Error during spacing test: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit(main())

# End of file #
