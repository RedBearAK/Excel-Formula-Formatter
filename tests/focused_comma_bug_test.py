#!/usr/bin/env python3
"""
Focused test to identify the exact comma bug in natural wrapping logic.
File: focused_comma_bug_test.py
"""

import sys
from pathlib import Path

package_parent = Path(__file__).parent
sys.path.insert(0, str(package_parent))


def manual_test_splitting():
    """Manually test the comma splitting and rejoining logic."""
    print("Manual Test of Comma Splitting Logic")
    print("=" * 45)
    
    # Import the formatter to access internal methods
    from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter
    
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    
    # Test the exact problematic case
    test_text = 'has_Paid_Date,NOT(has_Price),has_Invoice_Date'
    print(f"Testing: {test_text}")
    print(f"Original commas: {test_text.count(',')}")
    
    # Step 1: Parse tokens
    tokens = formatter._parse_excel_tokens(test_text)
    print(f"\nStep 1 - Tokens: {tokens}")
    comma_tokens = [i for i, (t_type, t_text) in enumerate(tokens) if t_text == ',']
    print(f"Comma token positions: {comma_tokens}")
    
    # Step 2: Split by commas
    groups = formatter._split_by_top_level_commas(tokens)
    print(f"\nStep 2 - Split into {len(groups)} groups:")
    for i, group in enumerate(groups):
        print(f"  Group {i+1}: {group}")
    
    # Step 3: Convert back to strings
    arg_strings = []
    for group in groups:
        arg_str = formatter._tokens_to_string(group)
        arg_strings.append(arg_str)
        print(f"  Group → '{arg_str}'")
    
    # Step 4: Rejoin with commas
    rejoined = ", ".join(arg_strings)
    print(f"\nStep 4 - Rejoined: '{rejoined}'")
    print(f"Final commas: {rejoined.count(',')}")
    
    if rejoined.count(',') != test_text.count(','):
        print("❌ COMMA LOSS DETECTED in splitting/rejoining!")
    else:
        print("✅ Comma splitting/rejoining works correctly")


def test_and_natural_wrapping():
    """Test the AND natural wrapping specifically."""
    print("\n\nTesting AND Natural Wrapping")
    print("=" * 35)
    
    from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter
    
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    
    # Test cases that should trigger different wrapping behaviors
    test_cases = [
        # Short - should stay on one line
        'AND(A1,B1)',
        # Medium - might trigger wrapping
        'AND(has_Paid_Date,NOT(has_Price),has_Invoice_Date)',
        # Long - definitely should wrap
        'AND(has_Paid_Date,NOT(has_Price),has_Invoice_Date,has_Title_Date,has_WBL_Date,has_Routing_ETA)',
    ]
    
    for i, test_case in enumerate(test_cases, 1):
        print(f"\nTest {i}: {test_case}")
        original_commas = test_case.count(',')
        
        try:
            # Test just the folding step to see if natural wrapping loses commas
            folded = formatter.fold_formula('=' + test_case)
            folded_commas = folded.count(',')
            
            print(f"  Original commas: {original_commas}")
            print(f"  Folded commas: {folded_commas}")
            
            if folded_commas != original_commas:
                print(f"  ❌ LOST {original_commas - folded_commas} COMMAS IN FOLDING!")
                print(f"  Folded result:")
                for line in folded.split('\n'):
                    if line.strip():
                        print(f"    {line}")
            else:
                print(f"  ✅ Folding preserved commas")
                
            # Now test unfolding
            unfolded = formatter.unfold_formula(folded)
            unfolded_commas = unfolded.count(',')
            
            print(f"  Unfolded commas: {unfolded_commas}")
            if unfolded_commas != original_commas:
                print(f"  ❌ LOST {original_commas - unfolded_commas} COMMAS IN UNFOLDING!")
                print(f"  Unfolded: {unfolded}")
            else:
                print(f"  ✅ Unfolding preserved commas")
                
        except Exception as e:
            print(f"  ERROR: {e}")


def test_step_by_step_and():
    """Step through the AND processing logic manually."""
    print("\n\nStep-by-Step AND Processing")
    print("=" * 35)
    
    from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter
    
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    
    test_formula = '=AND(has_Paid_Date,NOT(has_Price),has_Invoice_Date)'
    print(f"Formula: {test_formula}")
    
    # Parse tokens for the whole formula
    clean_formula = test_formula[1:]  # Remove =
    all_tokens = formatter._parse_excel_tokens(clean_formula)
    
    print(f"All tokens: {all_tokens}")
    
    # Find the AND function
    for i, (token_type, token_text) in enumerate(all_tokens):
        if token_type == 'function' and token_text.upper() == 'AND':
            print(f"\nFound AND at position {i}")
            
            # Extract its arguments
            if i + 1 < len(all_tokens) and all_tokens[i + 1][1] == '(':
                arg_tokens, end_index = formatter._extract_function_arguments(all_tokens, i + 1)
                print(f"AND argument tokens: {arg_tokens}")
                
                # Count commas in original tokens
                original_comma_tokens = [t for t in arg_tokens if t[1] == ',']
                print(f"Original comma tokens in AND: {len(original_comma_tokens)}")
                
                # Process through the logical function handler
                try:
                    func_lines = formatter._process_logical_function('AND', arg_tokens, 0)
                    print(f"Processed function lines:")
                    for line in func_lines:
                        print(f"  '{line}'")
                    
                    # Count commas in result
                    result_text = '\n'.join(func_lines)
                    result_commas = result_text.count(',')
                    print(f"Result comma count: {result_commas}")
                    
                    if result_commas != len(original_comma_tokens):
                        print(f"❌ LOGICAL FUNCTION PROCESSING LOST COMMAS!")
                    else:
                        print(f"✅ Logical function processing preserved commas")
                        
                except Exception as e:
                    print(f"ERROR in logical function processing: {e}")
            break


def main():
    """Run focused comma bug tests."""
    manual_test_splitting()
    test_and_natural_wrapping() 
    test_step_by_step_and()


if __name__ == "__main__":
    main()

# End of file #
