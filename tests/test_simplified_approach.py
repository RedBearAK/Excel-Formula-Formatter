#!/usr/bin/env python3
"""
Test the simplified function processing approach - no natural wrapping.
File: tests/test_simplified_approach.py
"""

import sys
from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


def test_and_or_functions():
    """Test that AND/OR functions now use simple one-per-line formatting."""
    print("Testing AND/OR Functions with Simplified Approach")
    print("=" * 55)
    
    test_cases = [
        # Simple cases that should stay inline
        '=AND(A1,B1)',
        '=OR(flag1,flag2)', 
        
        # Complex cases that should go multi-line
        '=AND(A1>0,B1<10,C1<>"")',
        '=AND(has_Paid_Date,NOT(has_Price),has_Invoice_Date)',
        '=AND(NOT(has_Paid_Date),NOT(has_WBL_Date),NOT(has_Title_Date),has_Routing_ETA,is_End_of_Year_ETA)',
        
        # Nested cases
        '=IF(AND(A1>0,B1<10),Result1,Result2)',
    ]
    
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    
    for i, test_case in enumerate(test_cases, 1):
        original_commas = test_case.count(',')
        
        print(f"\nTest {i}: {test_case}")
        
        try:
            folded = formatter.fold_formula(test_case)
            unfolded = formatter.unfold_formula(folded)
            final_commas = unfolded.count(',')
            
            print(f"  Commas: {original_commas} → {final_commas}")
            
            if final_commas == original_commas:
                print("  ✅ Commas preserved")
            else:
                print(f"  ❌ Lost {original_commas - final_commas} commas")
                print(f"  Result: {unfolded}")
            
            # Show folded structure for complex cases
            if len(folded.split('\n')) > 3:
                print("  Folded structure:")
                for j, line in enumerate(folded.split('\n'), 1):
                    if line.strip():
                        print(f"    {j:2d}: {line}")
                        
        except Exception as e:
            print(f"  ERROR: {e}")


def test_simple_inline_functions():
    """Test that simple functions stay inline."""
    print(f"\n\nTesting Simple Functions (Should Stay Inline)")
    print("=" * 50)
    
    simple_cases = [
        '=LEN(A1)',
        '=SUM(A1:A10)',
        '=NOT(flag)',
        '=ABS(value)',
        '=UPPER(text)',
    ]
    
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    
    for i, test_case in enumerate(simple_cases, 1):
        try:
            folded = formatter.fold_formula(test_case)
            lines = [line for line in folded.split('\n') if line.strip()]
            
            # Should be 2 lines: header comment + single function line
            is_inline = len(lines) <= 2
            
            print(f"{i}. {test_case}")
            print(f"   Inline: {'✅' if is_inline else '❌'} ({len(lines)} lines)")
            
            if not is_inline:
                print("   Folded:")
                for line in lines:
                    print(f"     {line}")
                    
        except Exception as e:
            print(f"{i}. {test_case} - ERROR: {e}")


def test_recursive_nesting():
    """Test recursive nesting works correctly."""
    print(f"\n\nTesting Recursive Nesting")
    print("=" * 30)
    
    nested_cases = [
        '=AND(OR(A1,B1),OR(C1,D1))',
        '=IF(AND(A1>0,OR(B1,C1)),Result1,Result2)',
        '=LET(x,A1,AND(x>0,NOT(ISERROR(B1)),SUM(C1:C10)>100))',
    ]
    
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    
    for i, test_case in enumerate(nested_cases, 1):
        original_commas = test_case.count(',')
        
        print(f"\nNested {i}: {test_case}")
        
        try:
            folded = formatter.fold_formula(test_case)
            unfolded = formatter.unfold_formula(folded)
            final_commas = unfolded.count(',')
            
            print(f"  Commas: {original_commas} → {final_commas} {'✅' if final_commas == original_commas else '❌'}")
            
            if final_commas != original_commas:
                print(f"  Result: {unfolded}")
            
            # Show indentation structure
            print("  Structure:")
            for j, line in enumerate(folded.split('\n'), 1):
                if line.strip() and not line.strip().startswith('//'):
                    indent_level = (len(line) - len(line.lstrip())) // 4
                    print(f"    {j:2d}: {'  ' * indent_level}→ {line.strip()}")
                    
        except Exception as e:
            print(f"  ERROR: {e}")


def test_user_problematic_pattern():
    """Test the exact pattern that was causing comma loss."""
    print(f"\n\nTesting User's Problematic Pattern")
    print("=" * 40)
    
    # The exact pattern from user's output that was losing commas
    problematic = '=AND(NOT(has_Paid_Date),NOT(has_WBL_Date),NOT(has_Title_Date),has_Routing_ETA,is_End_of_Year_ETA)'
    
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    original_commas = problematic.count(',')
    
    print(f"Pattern: {problematic}")
    print(f"Original commas: {original_commas}")
    
    try:
        # Test multiple cycles to check for progressive loss
        current_formula = problematic
        
        for cycle in range(3):
            folded = formatter.fold_formula(current_formula)
            unfolded = formatter.unfold_formula(folded)
            current_commas = unfolded.count(',')
            
            print(f"Cycle {cycle + 1}: {current_commas} commas {'✅' if current_commas == original_commas else '❌'}")
            
            if current_commas != original_commas:
                print(f"  Lost commas! Result: {unfolded}")
                break
                
            current_formula = unfolded
        
        if current_formula.count(',') == original_commas:
            print("✅ No progressive comma loss!")
        
        # Show the clean structure
        print("\nFolded structure:")
        for i, line in enumerate(folded.split('\n'), 1):
            if line.strip():
                print(f"  {i:2d}: {line}")
                
    except Exception as e:
        print(f"ERROR: {e}")


def main():
    """Run simplified approach tests."""
    print("Testing Simplified Function Processing Approach")
    print("=" * 60)
    print("No more natural wrapping - all functions use simple formatting")
    print()
    
    test_and_or_functions()
    test_simple_inline_functions()
    test_recursive_nesting()
    test_user_problematic_pattern()
    
    print("\n" + "=" * 60)
    print("Simplified approach testing complete!")
    print("All functions now use consistent, predictable formatting.")


if __name__ == "__main__":
    main()

# End of file #
