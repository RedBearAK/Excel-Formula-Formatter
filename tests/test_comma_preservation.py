#!/usr/bin/env python3
"""
Test comma preservation in complex formulas, especially nested AND/OR statements.
File: tests/test_comma_preservation.py
"""

import sys

from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


def count_commas(text: str) -> int:
    """Count commas in a text string."""
    return text.count(',')


def test_basic_comma_preservation():
    """Test that basic comma patterns are preserved."""
    test_cases = [
        '=SUM(A1,B1,C1)',
        '=IF(A1>0,B1,C1)',
        '=AND(A1>0,B1<10,C1<>"",D1>=E1)',
        '=OR(A1=B1,C1<>D1,E1>F1)',
        '=IFS(A1>0,"High",A1<0,"Low",TRUE,"Medium")',
        '=LET(x,A1,y,B1,x+y)',
    ]
    
    print("Basic Comma Preservation Test")
    print("=" * 40)
    
    all_success = True
    
    for mode in ['j', 'a', 'p']:
        print(f"\nMode {mode}:")
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
        
        for original in test_cases:
            original_commas = count_commas(original)
            
            try:
                folded = formatter.fold_formula(original)
                unfolded = formatter.unfold_formula(folded)
                unfolded_commas = count_commas(unfolded)
                
                success = original_commas == unfolded_commas
                all_success = all_success and success
                
                if not success:
                    print(f"  ❌ {original}")
                    print(f"     Original commas: {original_commas}, Unfolded commas: {unfolded_commas}")
                    print(f"     Unfolded: {unfolded}")
                else:
                    print(f"  ✅ {original[:30]}{'...' if len(original) > 30 else ''}")
                    
            except Exception as e:
                print(f"  ❌ ERROR with {original}: {e}")
                all_success = False
    
    return all_success


def test_complex_and_patterns():
    """Test complex AND patterns that might lose commas."""
    test_cases = [
        '=AND(A1>0,NOT(B1=""),C1<10)',
        '=AND(has_Paid_Date,NOT(has_Price),has_Invoice_Date)',
        '=AND(NOT(has_Paid_Date),NOT(has_WBL_Date),NOT(has_Title_Date))',
        '=IF(AND(A1>0,B1<>""),SUM(A1:A10),"")',
        '=IFS(AND(A1>0,NOT(B1="")),C1,AND(D1<10,E1>5),F1,TRUE,"")',
    ]
    
    print("\n\nComplex AND Pattern Test")
    print("=" * 40)
    
    all_success = True
    
    for mode in ['j', 'a', 'p']:
        print(f"\nMode {mode}:")
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
        
        for original in test_cases:
            original_commas = count_commas(original)
            
            try:
                folded = formatter.fold_formula(original)
                unfolded = formatter.unfold_formula(folded)
                unfolded_commas = count_commas(unfolded)
                
                success = original_commas == unfolded_commas
                all_success = all_success and success
                
                print(f"  {original}")
                print(f"    Commas: {original_commas} → {unfolded_commas} {'✅' if success else '❌'}")
                if not success:
                    print(f"    Unfolded: {unfolded}")
                    # Show the folded version to see what happened
                    print(f"    Folded preview:")
                    for line in folded.split('\n')[:5]:
                        print(f"      {line}")
                    
            except Exception as e:
                print(f"  ❌ ERROR: {e}")
                all_success = False
    
    return all_success


def test_malformed_formula_handling():
    """Test how the formatter handles malformed formulas (like the paste.txt example)."""
    malformed_cases = [
        # Missing comma between function arguments
        '=AND(A1>0NOT(B1=""))',  # Should be AND(A1>0,NOT(B1=""))
        '=IF(A1>0SUM(A1:A10))',   # Should be IF(A1>0,SUM(A1:A10))
        # Missing comma in nested functions
        '=IFS(AND(A1>0NOT(B1="")),"Good",TRUE,"")',
    ]
    
    print("\n\nMalformed Formula Handling Test")
    print("=" * 45)
    
    for mode in ['p']:  # Test just plain mode since that's where the issue was noticed
        print(f"\nMode {mode}:")
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
        
        for malformed in malformed_cases:
            print(f"  Testing: {malformed}")
            
            try:
                folded = formatter.fold_formula(malformed)
                unfolded = formatter.unfold_formula(folded)
                
                print(f"    Folded successfully: {len(folded.split(chr(10)))} lines")
                print(f"    Unfolded: {unfolded}")
                
                # Check if it's obviously broken
                if unfolded.count('(') != unfolded.count(')'):
                    print(f"    ⚠️  Parentheses mismatch!")
                
            except Exception as e:
                print(f"    ❌ ERROR: {e}")


def test_real_world_pattern():
    """Test a simplified version of the real-world pattern from paste.txt."""
    # A corrected version of part of the complex formula
    real_world = '''=LET(
        has_Paid_Date, LEN(AM2)>0,
        has_Price, LEN(AZ2)>=1,
        has_Invoice_Date, LEN(AE2)>1,
        IFS(
            AND(has_Paid_Date, NOT(has_Price), has_Invoice_Date),
            "Case1",
            AND(has_Paid_Date, NOT(has_Price), NOT(has_Invoice_Date)),
            "Case2",
            TRUE,
            ""
        )
    )'''
    
    print("\n\nReal World Pattern Test")
    print("=" * 35)
    
    # Remove extra whitespace and newlines to make it a single line
    single_line = ' '.join(line.strip() for line in real_world.split('\n') if line.strip())
    original_commas = count_commas(single_line)
    
    print(f"Original formula (simplified): {single_line[:100]}...")
    print(f"Original comma count: {original_commas}")
    
    for mode in ['j', 'a', 'p']:
        print(f"\nMode {mode}:")
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
        
        try:
            folded = formatter.fold_formula(single_line)
            unfolded = formatter.unfold_formula(folded)
            unfolded_commas = count_commas(unfolded)
            
            success = original_commas == unfolded_commas
            print(f"  Comma preservation: {original_commas} → {unfolded_commas} {'✅' if success else '❌'}")
            
            if not success:
                print(f"  Unfolded: {unfolded}")
                
        except Exception as e:
            print(f"  ❌ ERROR: {e}")


def main():
    """Run all comma preservation tests."""
    print("Excel Formula Formatter - Comma Preservation Tests")
    print("=" * 60)
    
    try:
        basic_success = test_basic_comma_preservation()
        complex_success = test_complex_and_patterns()
        test_malformed_formula_handling()  # This is diagnostic, not pass/fail
        test_real_world_pattern()
        
        print("\n" + "=" * 60)
        if basic_success and complex_success:
            print("✅ Comma preservation tests passed!")
            print("✨ Basic and complex patterns preserve commas correctly.")
        else:
            print("❌ Some comma preservation tests failed.")
            print("🔍 Check the output above for specific failures.")
        
        print("\n💡 Note: Some apparent 'failures' may be due to malformed input formulas.")
        print("   Always verify the original formula syntax is correct.")
        
        return 0 if basic_success and complex_success else 1
        
    except Exception as e:
        print(f"❌ Error during comma preservation test: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit(main())

# End of file #
