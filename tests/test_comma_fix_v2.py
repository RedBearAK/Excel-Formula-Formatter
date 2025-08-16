#!/usr/bin/env python3
"""
Test the v2 comma fix targeting the exact issue user reported.
File: tests/test_comma_fix_v2.py
"""

import sys
from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


def test_exact_user_pattern():
    """Test the exact patterns from user's output that were missing commas."""
    
    # These are the patterns that were losing commas in user's output
    problem_patterns = [
        # Basic case that should trigger line wrapping
        '=AND(NOT(has_Paid_Date),NOT(has_WBL_Date),NOT(has_Title_Date),has_Routing_ETA,is_End_of_Year_ETA)',
        
        # Another pattern from user's output  
        '=AND(NOT(has_Paid_Date),NOT(has_Title_Date),has_Routing_ETA,is_End_of_Year_ETA)',
        
        # Simpler case for debugging
        '=AND(NOT(A1),NOT(B1),NOT(C1),D1,E1)',
    ]
    
    print("Testing Exact User Patterns for Comma Loss")
    print("=" * 50)
    
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    
    for i, pattern in enumerate(problem_patterns, 1):
        print(f"\nPattern {i}: {pattern}")
        original_commas = pattern.count(',')
        print(f"Original commas: {original_commas}")
        
        try:
            folded = formatter.fold_formula(pattern)
            
            print("Folded output:")
            for line_num, line in enumerate(folded.split('\n'), 1):
                if line.strip():
                    print(f"  {line_num:2d}: {line}")
            
            # Check for the specific issue: missing comma after first line
            folded_lines = [line.strip() for line in folded.split('\n') if line.strip()]
            first_content_line = None
            for line in folded_lines:
                if 'AND(' in line:
                    first_content_line = line
                    break
            
            if first_content_line:
                has_comma_after_first_arg = first_content_line.endswith(',')
                print(f"  First line ends with comma: {has_comma_after_first_arg}")
                
                # Look for the pattern that was problematic
                if 'NOT(' in first_content_line and len(folded_lines) > 2:
                    if not has_comma_after_first_arg:
                        print(f"  ❌ MISSING COMMA after first argument!")
                    else:
                        print(f"  ✅ Comma correctly placed after first argument")
            
            # Test round trip
            unfolded = formatter.unfold_formula(folded)
            final_commas = unfolded.count(',')
            
            print(f"Final comma count: {final_commas}")
            
            if final_commas == original_commas:
                print("✅ All commas preserved")
            else:
                print(f"❌ Lost {original_commas - final_commas} commas")
                print(f"Result: {unfolded}")
                
        except Exception as e:
            print(f"ERROR: {e}")


def test_simple_debugging_case():
    """Test with a very simple case to debug step by step."""
    
    print(f"\n\nSimple Debugging Case")
    print("=" * 30)
    
    # Very simple case that should definitely wrap
    simple = '=AND(first_very_long_argument_name,second_very_long_argument_name,third_argument)'
    
    print(f"Testing: {simple}")
    
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    
    try:
        folded = formatter.fold_formula(simple)
        
        print("Folded result:")
        lines = folded.split('\n')
        for i, line in enumerate(lines, 1):
            if line.strip():
                commas_in_line = line.count(',')
                print(f"  {i}: {line.strip()} (commas: {commas_in_line})")
        
        total_commas = folded.count(',')
        original_commas = simple.count(',')
        
        print(f"\nComma check: {original_commas} → {total_commas}")
        
        unfolded = formatter.unfold_formula(folded)
        print(f"Unfolded: {unfolded}")
        
    except Exception as e:
        print(f"ERROR: {e}")


if __name__ == "__main__":
    test_exact_user_pattern()
    test_simple_debugging_case()
