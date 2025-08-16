#!/usr/bin/env python3
"""
Test to verify the comma fix works correctly.
File: test_comma_fix_verification.py
"""

import sys
from pathlib import Path

package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


def test_comma_fix():
    """Test that the comma fix resolves the AND function comma loss."""
    
    print("Testing Comma Fix for AND Functions")
    print("=" * 40)
    
    # Test cases that were previously losing commas
    test_cases = [
        '=AND(A1,B1)',
        '=AND(A1,B1,C1)',
        '=AND(has_Paid_Date,NOT(has_Price),has_Invoice_Date)',
        '=AND(NOT(has_Paid_Date),NOT(has_WBL_Date),NOT(has_Title_Date))',
        '=IF(AND(has_Paid_Date,NOT(has_Price)),Result1,Result2)',
    ]
    
    for mode in ['j', 'a', 'p']:
        print(f"\nTesting Mode {mode}:")
        print("-" * 20)
        
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
        
        for i, test_case in enumerate(test_cases, 1):
            original_commas = test_case.count(',')
            
            try:
                # Test multiple fold/unfold cycles
                current_formula = test_case
                
                for cycle in range(3):  # Test 3 cycles
                    folded = formatter.fold_formula(current_formula)
                    unfolded = formatter.unfold_formula(folded)
                    
                    if cycle == 0:  # Show first cycle details
                        folded_commas = folded.count(',')
                        unfolded_commas = unfolded.count(',')
                        
                        status = "✅" if unfolded_commas == original_commas else "❌"
                        print(f"  {i}. {test_case[:50]}{'...' if len(test_case) > 50 else ''}")
                        print(f"     {original_commas} → {folded_commas} → {unfolded_commas} {status}")
                        
                        if unfolded_commas != original_commas:
                            print(f"     Result: {unfolded}")
                    
                    current_formula = unfolded
                
                # Final check after 3 cycles
                final_commas = current_formula.count(',')
                if final_commas != original_commas:
                    print(f"     ❌ After 3 cycles: {original_commas} → {final_commas}")
                
            except Exception as e:
                print(f"  {i}. ERROR: {e}")


def test_specific_user_case():
    """Test the specific case from the user's paste.txt."""
    
    print(f"\n\nTesting User's Specific Case")
    print("=" * 35)
    
    # Simplified version of user's problematic formula
    user_formula = '''=LET(
        has_Paid_Date, LEN(AM2) > 0,
        has_Price, LEN(AZ2) >= 1,
        has_Invoice_Date, LEN(AE2) > 1,
        IFS(
            AND(has_Paid_Date, NOT(has_Price), has_Invoice_Date),
            "Case1",
            TRUE,
            ""
        )
    )'''
    
    # Convert to single line for testing
    single_line = ' '.join(line.strip() for line in user_formula.split('\n') if line.strip())
    original_commas = single_line.count(',')
    
    print(f"Original comma count: {original_commas}")
    
    for mode in ['p']:  # Focus on plain mode since that's where user saw the issue
        print(f"\nMode {mode}:")
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
        
        try:
            # Test progressive corruption
            current_formula = single_line
            
            for cycle in range(5):
                folded = formatter.fold_formula(current_formula)
                unfolded = formatter.unfold_formula(folded)
                current_commas = unfolded.count(',')
                
                print(f"  Cycle {cycle + 1}: {current_commas} commas {'✅' if current_commas == original_commas else '❌'}")
                
                if current_commas != original_commas:
                    print(f"    Lost {original_commas - current_commas} commas!")
                    break
                
                current_formula = unfolded
            
        except Exception as e:
            print(f"  ERROR: {e}")


def main():
    """Run comma fix verification tests."""
    test_comma_fix()
    test_specific_user_case()
    
    print("\n" + "=" * 50)
    print("Comma fix verification complete!")
    print("✅ If all tests show ✅, the comma fix is working.")
    print("❌ If any tests show ❌, there are still comma loss issues.")


if __name__ == "__main__":
    main()

# End of file #
