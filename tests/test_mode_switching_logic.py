#!/usr/bin/env python3
"""
Test mode switching logic to ensure plain mode never retains comments.
File: tests/test_mode_switching_logic.py
"""

import sys
from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.modular_excel_formatter import (
    ModularExcelFormatter, safe_mode_switch, detect_current_mode
)


def test_plain_mode_no_comments():
    """Test that plain mode never produces comments."""
    original = '=IF(A1<>B1,SUM(A1:A10),"Equal")'
    
    print("Testing Plain Mode Comment Elimination")
    print("=" * 45)
    print(f"Original: {original}")
    
    plain_formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    folded = plain_formatter.fold_formula(original)
    
    print(f"Plain mode folded:")
    for i, line in enumerate(folded.split('\n'), 1):
        if line.strip():
            print(f"  {i}: {line}")
    
    has_comments = '//' in folded
    print(f"\nHas comments: {has_comments}")
    
    if has_comments:
        print("❌ FAIL: Plain mode should never have comments")
        return False
    else:
        print("✅ PASS: Plain mode correctly has no comments")
        return True


def test_mode_switching_comment_elimination():
    """Test that switching to plain mode eliminates all comments."""
    original = '=IF(A1<>B1,SUM(A1:A10),"Equal")'
    
    print(f"\n\nTesting Mode Switching Comment Elimination")
    print("=" * 50)
    print(f"Original: {original}")
    
    # Start with JavaScript mode (has comments)
    js_formatter = ModularExcelFormatter.create_formatter_by_mode('j')
    js_folded = js_formatter.fold_formula(original)
    
    print(f"\nJavaScript mode folded (should have comments):")
    js_has_comments = '//' in js_folded
    print(f"  Has comments: {js_has_comments}")
    
    if not js_has_comments:
        print("❌ SETUP FAIL: JavaScript mode should have comments")
        return False
    
    # Switch to plain mode
    plain_result = safe_mode_switch(js_folded, 'j', 'p', should_refold=True)
    
    print(f"\nAfter switching to plain mode:")
    for i, line in enumerate(plain_result.split('\n'), 1):
        if line.strip():
            print(f"  {i}: {line}")
    
    plain_has_comments = '//' in plain_result
    print(f"\nHas comments after switch: {plain_has_comments}")
    
    if plain_has_comments:
        print("❌ FAIL: Switching to plain mode should eliminate all comments")
        return False
    else:
        print("✅ PASS: Switching to plain mode correctly eliminated comments")
        return True


def test_detect_mode_accuracy():
    """Test that mode detection works correctly for each mode."""
    original = '=SUM(A1:A10,B1:B10)'
    
    print(f"\n\nTesting Mode Detection Accuracy")
    print("=" * 40)
    print(f"Original: {original}")
    
    results = {}
    
    # Test each mode
    for mode, mode_name in [('j', 'JavaScript'), ('a', 'Annotated'), ('p', 'Plain')]:
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
        folded = formatter.fold_formula(original)
        detected = detect_current_mode(folded)
        
        print(f"\n{mode_name} mode ({mode}):")
        print(f"  Detected as: {detected}")
        print(f"  Correct: {'✅' if detected == mode else '❌'}")
        
        results[mode] = detected == mode
    
    # Test unfolded formula detection
    unfolded_detected = detect_current_mode(original)
    print(f"\nUnfolded formula:")
    print(f"  Detected as: {unfolded_detected}")
    print(f"  Expected: p (plain)")
    print(f"  Correct: {'✅' if unfolded_detected == 'p' else '❌'}")
    
    results['unfolded'] = unfolded_detected == 'p'
    
    all_correct = all(results.values())
    print(f"\nAll detections correct: {'✅' if all_correct else '❌'}")
    
    return all_correct


def test_edge_case_comment_leakage():
    """Test edge cases where comments might leak into plain mode."""
    print(f"\n\nTesting Edge Case Comment Leakage")
    print("=" * 40)
    
    # Test case 1: JavaScript with complex comments -> Plain
    print(f"Test 1: Complex formula with JavaScript comments -> Plain")
    js_formatter = ModularExcelFormatter.create_formatter_by_mode('j')
    complex_original = '=LET(x,A1+B1,y,C1*D1,IF(x>y,"X wins","Y wins"))'
    print(f"Original: {complex_original}")
    
    js_folded = js_formatter.fold_formula(complex_original)
    js_comment_count = js_folded.count('//')
    print(f"JavaScript folded has {js_comment_count} comment lines")
    
    plain_result = safe_mode_switch(js_folded, 'j', 'p', should_refold=True)
    plain_comment_count = plain_result.count('//')
    print(f"Plain result has {plain_comment_count} comment lines")
    
    test1_passed = (plain_comment_count == 0)
    print(f"Result: {'✅ PASS' if test1_passed else '❌ FAIL - comments leaked!'}")
    
    if not test1_passed:
        print(f"Plain result with comments:\n{plain_result}")
    
    # Test case 2: Annotated with section comments -> Plain  
    print(f"\nTest 2: Annotated with section comments -> Plain")
    annotated_formatter = ModularExcelFormatter.create_formatter_by_mode('a')
    ifs_original = '=IFS(A1>0,"High",A1<0,"Low",TRUE,"Medium")'
    
    annotated_folded = annotated_formatter.fold_formula(ifs_original)
    annotated_comment_count = annotated_folded.count('//')
    print(f"Annotated folded has {annotated_comment_count} comment lines")
    
    plain_result2 = safe_mode_switch(annotated_folded, 'a', 'p', should_refold=True)
    plain_comment_count2 = plain_result2.count('//')
    print(f"Plain result has {plain_comment_count2} comment lines")
    
    test2_passed = (plain_comment_count2 == 0)
    print(f"Result: {'✅ PASS' if test2_passed else '❌ FAIL - comments leaked!'}")
    
    if not test2_passed:
        print(f"Plain result with comments:\n{plain_result2}")
    
    # Test case 3: Direct PlainExcelTranslator test
    print(f"\nTest 3: Direct PlainExcelTranslator behavior")
    plain_formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    
    test_formulas = [
        '=SUM(A1:A10)',
        '=IF(A1>0,"Positive","Negative")',
        '=IFS(A1>0,"High",A1<0,"Low",TRUE,"Medium")',
        '=LET(x,A1,y,B1,x+y)'
    ]
    
    test3_passed = True
    for formula in test_formulas:
        plain_folded = plain_formatter.fold_formula(formula)
        if '//' in plain_folded:
            print(f"❌ Plain formatter generated comments for: {formula}")
            print(f"Result: {plain_folded}")
            test3_passed = False
    
    if test3_passed:
        print("✅ Plain formatter never generates comments")
    
    # Final result - explicitly return boolean
    overall_success = test1_passed and test2_passed and test3_passed
    return overall_success


def test_round_trip_mode_switching():
    """Test round-trip mode switching preserves formula integrity."""
    original = '=IF(AND(A1>0,B1<>""),SUM(A1:A10),MAX(B1:B10))'
    
    print(f"\n\nTesting Round-Trip Mode Switching")
    print("=" * 40)
    print(f"Original: {original}")
    
    # Mode switching sequence: p -> j -> a -> p
    current_text = original
    modes = ['p', 'j', 'a', 'p']
    
    for i in range(len(modes) - 1):
        from_mode = modes[i]
        to_mode = modes[i + 1]
        
        print(f"\nSwitching {from_mode} -> {to_mode}:")
        
        # If starting with unfolded formula, fold it first
        if i == 0 and not ('\n' in current_text or '//' in current_text):
            formatter = ModularExcelFormatter.create_formatter_by_mode(from_mode)
            current_text = formatter.fold_formula(current_text)
            print(f"  Folded with {from_mode} first")
        
        # Perform switch
        switched_text = safe_mode_switch(current_text, from_mode, to_mode, should_refold=True)
        
        # Check for comments in plain mode
        if to_mode == 'p':
            has_comments = '//' in switched_text
            print(f"  Plain mode result has comments: {'YES (❌ FAIL)' if has_comments else 'NO (✅ PASS)'}")
            if has_comments:
                print(f"  Text with comments: {switched_text[:100]}...")
                return False
        
        current_text = switched_text
    
    # Final unfold and check integrity
    final_formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    final_unfolded = final_formatter.unfold_formula(current_text)
    
    # Normalize for comparison
    original_norm = original.replace(' ', '').replace('=', '')
    final_norm = final_unfolded.replace(' ', '').replace('=', '')
    
    integrity_preserved = original_norm == final_norm
    
    print(f"\nFinal unfolded: {final_unfolded}")
    print(f"Formula integrity preserved: {'✅' if integrity_preserved else '❌'}")
    
    if not integrity_preserved:
        print(f"  Original normalized: {original_norm}")
        print(f"  Final normalized: {final_norm}")
    
    return integrity_preserved


def main():
    """Run mode switching logic tests."""
    print("Mode Switching Logic Tests")
    print("=" * 60)
    print("Testing to ensure plain mode never retains comments after mode switches")
    print()
    
    tests = [
        ("Plain Mode No Comments", test_plain_mode_no_comments),
        ("Mode Switch Comment Elimination", test_mode_switching_comment_elimination),
        ("Mode Detection Accuracy", test_detect_mode_accuracy),
        ("Edge Case Comment Leakage", test_edge_case_comment_leakage),
        ("Round-Trip Mode Switching", test_round_trip_mode_switching)
    ]
    
    results = []
    for test_name, test_func in tests:
        print(f"Running {test_name} test...")
        print("-" * 50)
        try:
            success = test_func()
            # Ensure we have a boolean result
            if success is None:
                print(f"⚠️ {test_name}: Function returned None, treating as FAIL")
                success = False
            elif not isinstance(success, bool):
                print(f"⚠️ {test_name}: Function returned {type(success)}, converting to bool")
                success = bool(success)
            
            results.append(success)
            print(f"✓ {test_name}: {'PASS' if success else 'FAIL'}")
        except Exception as e:
            print(f"✗ {test_name}: ERROR - {e}")
            results.append(False)
        print()
    
    # Safe sum that handles any potential None values
    passed = sum(1 for result in results if result is True)
    total = len(results)
    
    print("=" * 60)
    print(f"Mode Switching Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All mode switching tests passed!")
        print("✅ Plain mode correctly eliminates comments after mode switches.")
        print("✅ Mode detection accuracy improved.")
        print("✅ No edge case comment leakage detected.")
        return 0
    else:
        print("❌ Some mode switching tests failed.")
        print("⚠️  Check the output above for specific issues.")
        return 1


if __name__ == "__main__":
    exit(main())

# End of file #
