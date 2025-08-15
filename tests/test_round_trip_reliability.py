#!/usr/bin/env python3
"""
Comprehensive round-trip tests for modular Excel formatter.
File: tests/test_round_trip_reliability.py
"""

import sys

from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.modular_excel_formatter import (
    ModularExcelFormatter, detect_current_mode, safe_mode_switch
)


def normalize_formula(formula: str) -> str:
    """Normalize Excel formula for comparison by removing leading = and spaces."""
    normalized = formula.strip()
    if normalized.startswith('='):
        normalized = normalized[1:]
    return normalized.replace(' ', '')


def test_javascript_simple_round_trip():
    """Test JavaScript translator simple round-trip."""
    formatter = ModularExcelFormatter.create_javascript_formatter()
    original = '=SUM(A1:A10,B1:B10)'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded preview: {folded.split(chr(10))[0]}...")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check round-trip accuracy
    original_norm = normalize_formula(original)
    unfolded_norm = normalize_formula(unfolded)
    
    success = original_norm == unfolded_norm
    print(f"Round-trip success: {success}")
    return success


def test_javascript_complex_round_trip():
    """Test JavaScript translator with complex nested formula."""
    formatter = ModularExcelFormatter.create_javascript_formatter()
    original = '=IF(AND(A1>0,B1<>""),SUM(A1:A10),IF(C1>=D1,MAX(E1:E10),"Error"))'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded lines: {len(folded.split(chr(10)))}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check round-trip accuracy  
    original_norm = normalize_formula(original)
    unfolded_norm = normalize_formula(unfolded)
    
    success = original_norm == unfolded_norm
    print(f"Round-trip success: {success}")
    return success


def test_javascript_idempotency():
    """Test JavaScript translator idempotency (fold→unfold→fold produces same result)."""
    formatter = ModularExcelFormatter.create_javascript_formatter()
    original = '=LET(x,A1+B1,y,C1*D1,IF(x>y,x,y))'
    
    # First cycle
    folded1 = formatter.fold_formula(original)
    unfolded1 = formatter.unfold_formula(folded1)
    
    # Second cycle  
    folded2 = formatter.fold_formula(unfolded1)
    unfolded2 = formatter.unfold_formula(folded2)
    
    print(f"Original: {original}")
    print(f"First unfolded: {unfolded1}")
    print(f"Second unfolded: {unfolded2}")
    print()
    
    # Check that both unfolded results are identical
    unfolded1_norm = normalize_formula(unfolded1)
    unfolded2_norm = normalize_formula(unfolded2)
    original_norm = normalize_formula(original)
    
    round_trip_success = original_norm == unfolded1_norm
    idempotent_success = unfolded1_norm == unfolded2_norm
    
    success = round_trip_success and idempotent_success
    print(f"Round-trip success: {round_trip_success}")
    print(f"Idempotency success: {idempotent_success}")
    print(f"Overall success: {success}")
    return success


def test_plain_simple_round_trip():
    """Test Plain translator simple round-trip."""
    formatter = ModularExcelFormatter.create_plain_formatter()
    original = '=SUM(A1:A10,B1:B10)'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded preview: {folded.split(chr(10))[0]}...")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check round-trip accuracy
    original_norm = normalize_formula(original)
    unfolded_norm = normalize_formula(unfolded)
    
    success = original_norm == unfolded_norm
    print(f"Round-trip success: {success}")
    return success


def test_plain_complex_round_trip():
    """Test Plain translator with complex operators and functions."""
    formatter = ModularExcelFormatter.create_plain_formatter()
    original = '=IF(A1<>B1,CONCATENATE("Different: ",A1," vs ",B1),AND(A1>=0,B1<=100))'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded lines: {len(folded.split(chr(10)))}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check round-trip accuracy
    original_norm = normalize_formula(original)
    unfolded_norm = normalize_formula(unfolded)
    
    success = original_norm == unfolded_norm
    print(f"Round-trip success: {success}")
    return success


def test_plain_idempotency():
    """Test Plain translator idempotency."""
    formatter = ModularExcelFormatter.create_plain_formatter()
    original = '=IF(AND(A1>0,B1<>C1),SUM(A1:A10)&" total",AVERAGE(D1:D10))'
    
    # First cycle
    folded1 = formatter.fold_formula(original)
    unfolded1 = formatter.unfold_formula(folded1)
    
    # Second cycle
    folded2 = formatter.fold_formula(unfolded1)
    unfolded2 = formatter.unfold_formula(folded2)
    
    print(f"Original: {original}")
    print(f"First unfolded: {unfolded1}")
    print(f"Second unfolded: {unfolded2}")
    print()
    
    # Check consistency
    unfolded1_norm = normalize_formula(unfolded1)
    unfolded2_norm = normalize_formula(unfolded2)
    original_norm = normalize_formula(original)
    
    round_trip_success = original_norm == unfolded1_norm
    idempotent_success = unfolded1_norm == unfolded2_norm
    
    success = round_trip_success and idempotent_success
    print(f"Round-trip success: {round_trip_success}")
    print(f"Idempotency success: {idempotent_success}")
    print(f"Overall success: {success}")
    return success


def test_operators_preservation():
    """Test that all Excel operators are preserved correctly."""
    test_cases = [
        ('=A1+B1', 'Addition'),
        ('=A1-B1', 'Subtraction'),
        ('=A1*B1', 'Multiplication'),
        ('=A1/B1', 'Division'),
        ('=A1=B1', 'Equality'),
        ('=A1<>B1', 'Not equal'),
        ('=A1>B1', 'Greater than'),
        ('=A1<B1', 'Less than'),
        ('=A1>=B1', 'Greater equal'),
        ('=A1<=B1', 'Less equal'),
        ('=A1&B1', 'Concatenation')
    ]
    
    print("Testing operator preservation...")
    print()
    
    js_formatter = ModularExcelFormatter.create_javascript_formatter()
    plain_formatter = ModularExcelFormatter.create_plain_formatter()
    
    js_results = []
    plain_results = []
    
    for formula, desc in test_cases:
        # Test JavaScript translator
        js_folded = js_formatter.fold_formula(formula)
        js_unfolded = js_formatter.unfold_formula(js_folded)
        js_success = normalize_formula(formula) == normalize_formula(js_unfolded)
        js_results.append(js_success)
        
        # Test Plain translator
        plain_folded = plain_formatter.fold_formula(formula)
        plain_unfolded = plain_formatter.unfold_formula(plain_folded)
        plain_success = normalize_formula(formula) == normalize_formula(plain_unfolded)
        plain_results.append(plain_success)
        
        print(f"{desc:15} | JS: {'✓' if js_success else '✗'} | Plain: {'✓' if plain_success else '✗'} | {formula}")
    
    print()
    js_passed = sum(js_results)
    plain_passed = sum(plain_results)
    total = len(test_cases)
    
    print(f"JavaScript: {js_passed}/{total} operators preserved")
    print(f"Plain: {plain_passed}/{total} operators preserved")
    
    success = js_passed == total and plain_passed == total
    print(f"All operators preserved: {success}")
    return success


def test_safe_mode_switching():
    """Test safe mode switching functionality."""
    original = '=IF(AND(A1>0,B1<>""),SUM(A1:A10),MAX(B1:B10))'
    
    # Start with JavaScript mode
    js_formatter = ModularExcelFormatter.create_javascript_formatter()
    js_folded = js_formatter.fold_formula(original)
    
    print(f"Original: {original}")
    print(f"JavaScript folded (preview): {js_folded.split(chr(10))[0]}...")
    
    # Detect current mode
    detected_mode = detect_current_mode(js_folded)
    print(f"Detected mode: {detected_mode}")
    
    # Safe switch to plain mode (should unfold first)
    plain_result = safe_mode_switch(js_folded, 'javascript', 'plain', should_refold=False)
    print(f"After safe switch to plain: {plain_result}")
    
    # Safe switch back to javascript with refold
    js_result = safe_mode_switch(plain_result, 'plain', 'javascript', should_refold=True)
    print(f"After safe switch back to JS (folded): {len(js_result.split(chr(10)))} lines")
    
    # Final unfold to verify integrity
    final_result = js_formatter.unfold_formula(js_result)
    print(f"Final result: {final_result}")
    print()
    
    # Check that we maintain formula integrity throughout
    original_norm = normalize_formula(original)
    plain_norm = normalize_formula(plain_result)
    final_norm = normalize_formula(final_result)
    
    mode_detection_success = detected_mode == 'javascript'
    switch_success = original_norm == plain_norm == final_norm
    
    success = mode_detection_success and switch_success
    print(f"Mode detection success: {mode_detection_success}")
    print(f"Safe switching success: {switch_success}")
    print(f"Overall success: {success}")
    return success


def test_complex_nested_formula():
    """Test with the actual complex formula from paste.txt."""
    formatter = ModularExcelFormatter.create_javascript_formatter()
    # Simplified version of the complex formula
    original = '=LET(has_Date,LEN(A1)>0,has_Price,LEN(B1)>=1,IFS(AND(has_Date,NOT(has_Price)),A1+30,TRUE,""))'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded lines: {len(folded.split(chr(10)))}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check round-trip accuracy
    original_norm = normalize_formula(original)
    unfolded_norm = normalize_formula(unfolded)
    
    success = original_norm == unfolded_norm
    print(f"Complex formula round-trip success: {success}")
    return success


def main():
    """Run all round-trip reliability tests."""
    print("Excel Formula Formatter - Round-trip Reliability Tests")
    print("=" * 70)
    print()
    
    tests = [
        ("JavaScript Simple Round-trip", test_javascript_simple_round_trip),
        ("JavaScript Complex Round-trip", test_javascript_complex_round_trip),
        ("JavaScript Idempotency", test_javascript_idempotency),
        ("Plain Simple Round-trip", test_plain_simple_round_trip),
        ("Plain Complex Round-trip", test_plain_complex_round_trip),
        ("Plain Idempotency", test_plain_idempotency),
        ("Operator Preservation", test_operators_preservation),
        ("Safe Mode Switching", test_safe_mode_switching),
        ("Complex Nested Formula", test_complex_nested_formula)
    ]
    
    results = []
    for test_name, test_func in tests:
        print(f"Running {test_name} test...")
        print("-" * 50)
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
    
    print("=" * 70)
    print(f"Round-trip Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All round-trip tests passed!")
        print("✅ JavaScript and Plain translators are reliable for production use.")
        return 0
    else:
        print("❌ Some round-trip tests failed.")
        print("⚠️  Review failed tests before using in production.")
        return 1


if __name__ == "__main__":
    exit(main())

# End of file #
