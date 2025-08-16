#!/usr/bin/env python3
"""
Debug tool to trace exactly what happens to commas in AND() functions.
File: debug_and_processing.py
"""

import sys

from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


def debug_and_processing():
    """Debug exactly what happens to AND() function processing."""
    
    # Test cases that should reveal the comma loss pattern
    test_cases = [
        '=AND(A1>0,B1<10)',  # Simple case
        '=AND(A1>0,B1<10,C1<>"")',  # Three arguments
        '=AND(has_Paid_Date,NOT(has_Price),has_Invoice_Date)',  # Like user's example
        '=AND(NOT(has_Paid_Date),NOT(has_WBL_Date),NOT(has_Title_Date))',  # Multiple NOTs
        '=IF(AND(A1>0,B1<10),C1,D1)',  # AND inside IF
    ]
    
    print("Debug AND() Function Token Processing")
    print("=" * 50)
    
    for i, test_case in enumerate(test_cases, 1):
        print(f"\nTest Case {i}: {test_case}")
        print("-" * 30)
        
        original_commas = test_case.count(',')
        print(f"Original comma count: {original_commas}")
        
        for mode in ['j', 'a', 'p']:
            print(f"\nMode {mode}:")
            
            try:
                formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
                
                # Step 1: Parse tokens
                if test_case.startswith('='):
                    clean_formula = test_case[1:]  # Remove =
                else:
                    clean_formula = test_case
                
                tokens = formatter._parse_excel_tokens(clean_formula)
                print(f"  Tokens parsed: {len(tokens)}")
                
                # Count commas in tokens
                comma_tokens = [t for t in tokens if t[1] == ',']
                print(f"  Comma tokens found: {len(comma_tokens)}")
                
                # Step 2: Fold
                folded = formatter.fold_formula(test_case)
                folded_commas = folded.count(',')
                print(f"  Folded comma count: {folded_commas}")
                
                # Step 3: Unfold
                unfolded = formatter.unfold_formula(folded)
                unfolded_commas = unfolded.count(',')
                print(f"  Unfolded comma count: {unfolded_commas}")
                
                print(f"  Final result: {unfolded}")
                
                # Check for loss
                if unfolded_commas != original_commas:
                    print(f"  ❌ COMMA LOSS: {original_commas} → {unfolded_commas}")
                    print(f"     Lost in folding: {original_commas - folded_commas}")
                    print(f"     Lost in unfolding: {folded_commas - unfolded_commas}")
                else:
                    print(f"  ✅ Commas preserved")
                    
            except Exception as e:
                print(f"  ❌ ERROR: {e}")


def debug_token_splitting():
    """Debug the _split_by_top_level_commas function specifically."""
    print("\n\nDebug Token Splitting Function")
    print("=" * 40)
    
    # Create a formatter to access the internal method
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    
    # Test the problematic AND pattern
    test_formula = 'AND(has_Paid_Date,NOT(has_Price),has_Invoice_Date)'
    
    print(f"Testing: {test_formula}")
    
    # Parse tokens
    tokens = formatter._parse_excel_tokens(test_formula)
    print(f"All tokens: {tokens}")
    
    # Find the AND function and extract its arguments
    for i, (token_type, token_text) in enumerate(tokens):
        if token_type == 'function' and token_text.upper() == 'AND':
            if i + 1 < len(tokens) and tokens[i + 1][1] == '(':
                arg_tokens, end_index = formatter._extract_function_arguments(tokens, i + 1)
                print(f"AND arguments tokens: {arg_tokens}")
                
                # Split by commas
                argument_groups = formatter._split_by_top_level_commas(arg_tokens)
                print(f"Split into {len(argument_groups)} groups:")
                
                for j, group in enumerate(argument_groups):
                    group_text = formatter._tokens_to_string(group)
                    print(f"  Group {j + 1}: {group} → '{group_text}'")
                
                break


def debug_natural_wrapping():
    """Debug the natural wrapping logic for AND functions."""
    print("\n\nDebug Natural Wrapping Logic")
    print("=" * 35)
    
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    
    # Test a case that would trigger natural wrapping
    long_and = 'AND(has_Paid_Date,NOT(has_Price),has_Invoice_Date,has_Title_Date,has_WBL_Date)'
    
    print(f"Testing long AND: {long_and}")
    original_commas = long_and.count(',')
    
    try:
        folded = formatter.fold_formula('=' + long_and)
        unfolded = formatter.unfold_formula(folded)
        
        print(f"Original commas: {original_commas}")
        print(f"Folded commas: {folded.count(',')}")
        print(f"Unfolded commas: {unfolded.count(',')}")
        
        print(f"\nFolded result:")
        for line in folded.split('\n'):
            print(f"  {line}")
        
        print(f"\nUnfolded result: {unfolded}")
        
    except Exception as e:
        print(f"ERROR: {e}")


def main():
    """Run all debugging functions."""
    debug_and_processing()
    debug_token_splitting()
    debug_natural_wrapping()
    
    print("\n" + "=" * 60)
    print("Debug complete. Check output above to identify where commas are lost.")


if __name__ == "__main__":
    main()

# End of file #
