#!/usr/bin/env python3
"""
Simple test for comma loss.
File: tests/simple_comma_loss_test.py
"""

import sys
from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


def test_specific_comma_loss():
    """Test the exact pattern that's losing commas."""
    
    # This pattern is definitely losing commas based on user's output
    problematic_pattern = '=AND(NOT(has_Paid_Date),NOT(has_WBL_Date),NOT(has_Title_Date),has_Routing_ETA,is_End_of_Year_ETA)'
    
    print("Testing Specific Comma Loss Pattern")
    print("=" * 45)
    print(f"Pattern: {problematic_pattern}")
    print(f"Original commas: {problematic_pattern.count(',')}")
    print()
    
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    
    try:
        folded = formatter.fold_formula(problematic_pattern)
        print("Folded result:")
        for i, line in enumerate(folded.split('\n'), 1):
            if line.strip():
                comma_count = line.count(',')
                print(f"  {i:2d}: {line} (commas: {comma_count})")
        
        total_folded_commas = folded.count(',')
        print(f"\nTotal folded commas: {total_folded_commas}")
        
        unfolded = formatter.unfold_formula(folded)
        print(f"Unfolded: {unfolded}")
        print(f"Unfolded commas: {unfolded.count(',')}")
        
        if unfolded.count(',') != problematic_pattern.count(','):
            print("❌ COMMA LOSS CONFIRMED!")
        else:
            print("✅ Commas preserved")
            
    except Exception as e:
        print(f"ERROR: {e}")


def test_progressive_loss():
    """Test if the loss gets worse with multiple cycles."""
    
    pattern = '=AND(NOT(A1),NOT(B1),NOT(C1))'
    
    print(f"\n\nProgressive Loss Test")
    print("=" * 25)
    print(f"Starting with: {pattern}")
    
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    current = pattern
    
    for cycle in range(3):
        try:
            folded = formatter.fold_formula(current)
            unfolded = formatter.unfold_formula(folded)
            
            print(f"Cycle {cycle + 1}: {current.count(',')} → {unfolded.count(',')}")
            
            if unfolded.count(',') != pattern.count(','):
                print(f"  ❌ Lost commas: {unfolded}")
                break
            
            current = unfolded
            
        except Exception as e:
            print(f"  ERROR: {e}")
            break


if __name__ == "__main__":
    test_specific_comma_loss()
    test_progressive_loss()
