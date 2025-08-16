#!/usr/bin/env python3
"""
Simple test to trace exactly where commas are lost.
File: simple_comma_trace.py
"""

import sys
from pathlib import Path

# Add package to path  
package_parent = Path(__file__).parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


def trace_comma_loss():
    """Trace comma loss in a specific problematic formula."""
    
    # Use a formula that definitely should have commas
    problem_formula = '=AND(has_Paid_Date,NOT(has_Price),has_Invoice_Date)'
    
    print("Simple Comma Loss Trace")
    print("=" * 30)
    print(f"Testing: {problem_formula}")
    print(f"Original commas: {problem_formula.count(',')}")
    print()
    
    for mode in ['j', 'a', 'p']:
        print(f"Mode {mode}:")
        print("-" * 10)
        
        try:
            formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
            
            # Step by step
            print(f"1. Original: {problem_formula} (commas: {problem_formula.count(',')})")
            
            folded = formatter.fold_formula(problem_formula)
            print(f"2. Folded commas: {folded.count(',')}")
            
            # Show folded result
            folded_lines = folded.split('\n')
            for i, line in enumerate(folded_lines):
                if line.strip():
                    print(f"   Line {i+1}: {line}")
            
            unfolded = formatter.unfold_formula(folded)
            print(f"3. Unfolded: {unfolded} (commas: {unfolded.count(',')})")
            
            if unfolded.count(',') != problem_formula.count(','):
                print(f"   ❌ LOST {problem_formula.count(',') - unfolded.count(',')} COMMAS!")
            else:
                print(f"   ✅ All commas preserved")
            
        except Exception as e:
            print(f"   ERROR: {e}")
        
        print()


def test_simple_cases():
    """Test progressively more complex cases to isolate the issue."""
    
    test_cases = [
        '=A1,B1',  # Just commas
        '=SUM(A1,B1)',  # Simple function with comma
        '=AND(A1,B1)',  # Simple AND
        '=AND(A1,B1,C1)',  # AND with 3 args
        '=AND(A1>0,B1<10)',  # AND with conditions
        '=AND(A1,NOT(B1))',  # AND with nested function
        '=AND(A1,NOT(B1),C1)',  # AND with nested function and 3 args
    ]
    
    print("Progressive Complexity Test")
    print("=" * 35)
    
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')  # Use plain mode
    
    for i, test in enumerate(test_cases, 1):
        original_commas = test.count(',')
        
        try:
            folded = formatter.fold_formula(test)
            unfolded = formatter.unfold_formula(folded)
            final_commas = unfolded.count(',')
            
            status = "✅" if final_commas == original_commas else "❌"
            print(f"{i}. {test}")
            print(f"   {original_commas} → {final_commas} {status}")
            
            if final_commas != original_commas:
                print(f"   Unfolded: {unfolded}")
            
        except Exception as e:
            print(f"{i}. {test}")
            print(f"   ERROR: {e}")
        
        print()


def main():
    """Run comma tracing tests."""
    trace_comma_loss()
    test_simple_cases()


if __name__ == "__main__":
    main()

# End of file #
