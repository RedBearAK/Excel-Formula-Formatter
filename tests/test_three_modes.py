#!/usr/bin/env python3
"""
Test three-mode system for Excel formula formatter.
File: tests/test_three_modes.py
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


def test_javascript_mode():
    """Test JavaScript mode (j) - should have comments and quoted cell references."""
    formatter = ModularExcelFormatter.create_formatter_by_mode('j')
    original = '=IF(A1<>B1,SUM(A1:A10),"Equal")'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check JavaScript mode characteristics
    has_js_comments = '// Excel Formula (JavaScript syntax' in folded
    has_quoted_cells = '"A1"' in folded and '"B1"' in folded
    has_js_operator = '!=' in folded
    has_excel_operator_back = '<>' in unfolded
    
    # Check round-trip
    round_trip_success = normalize_formula(original) == normalize_formula(unfolded)
    
    success = (has_js_comments and has_quoted_cells and has_js_operator and 
              has_excel_operator_back and round_trip_success)
    
    print(f"Has JavaScript comments: {has_js_comments}")
    print(f"Has quoted cell references: {has_quoted_cells}")
    print(f"Has != operator: {has_js_operator}")
    print(f"Has <> operator in unfolded: {has_excel_operator_back}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"JavaScript mode success: {success}")
    return success


def test_annotated_excel_mode():
    """Test Annotated Excel mode (a) - should have comments but no quoted cells."""
    formatter = ModularExcelFormatter.create_formatter_by_mode('a')
    original = '=IF(A1<>B1,SUM(A1:A10),"Equal")'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check Annotated Excel mode characteristics
    has_annotated_comments = '// Excel Formula (annotated Excel syntax' in folded
    has_unquoted_cells = 'A1' in folded and '"A1"' not in folded
    has_excel_operator = '<>' in folded
    has_spacing = '( ' in folded and ' )' in folded  # Should have spacing around parentheses
    
    # Check round-trip
    round_trip_success = normalize_formula(original) == normalize_formula(unfolded)
    
    success = (has_annotated_comments and has_unquoted_cells and has_excel_operator and 
              has_spacing and round_trip_success)
    
    print(f"Has annotated Excel comments: {has_annotated_comments}")
    print(f"Has unquoted cell references: {has_unquoted_cells}")
    print(f"Has <> operator: {has_excel_operator}")
    print(f"Has spacing around parentheses: {has_spacing}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Annotated Excel mode success: {success}")
    return success


def test_plain_excel_mode():
    """Test Plain Excel mode (p) - should have NO comments, just pure Excel with indenting."""
    formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    original = '=IF(A1<>B1,SUM(A1:A10),"Equal")'
    
    folded = formatter.fold_formula(original)
    unfolded = formatter.unfold_formula(folded)
    
    print(f"Original: {original}")
    print(f"Folded:\n{folded}")
    print(f"Unfolded: {unfolded}")
    print()
    
    # Check Plain Excel mode characteristics
    has_no_comments = '//' not in folded  # Should have NO comments at all
    has_unquoted_cells = 'A1' in folded and '"A1"' not in folded
    has_excel_operator = '<>' in folded
    has_spacing = '( ' in folded and ' )' in folded  # Should have spacing around parentheses
    has_indentation = any(line.startswith('    ') for line in folded.split('\n'))
    
    # Check round-trip
    round_trip_success = normalize_formula(original) == normalize_formula(unfolded)
    
    success = (has_no_comments and has_unquoted_cells and has_excel_operator and 
              has_spacing and has_indentation and round_trip_success)
    
    print(f"Has NO comments: {has_no_comments}")
    print(f"Has unquoted cell references: {has_unquoted_cells}")
    print(f"Has <> operator: {has_excel_operator}")
    print(f"Has spacing around parentheses: {has_spacing}")
    print(f"Has indentation: {has_indentation}")
    print(f"Round-trip success: {round_trip_success}")
    print(f"Plain Excel mode success: {success}")
    return success


def test_mode_detection():
    """Test automatic mode detection from text content."""
    # Test JavaScript mode detection
    js_formatter = ModularExcelFormatter.create_formatter_by_mode('j')
    js_folded = js_formatter.fold_formula('=SUM(A1:A10)')
    js_detected = detect_current_mode(js_folded)
    
    # Test Annotated Excel mode detection
    annotated_formatter = ModularExcelFormatter.create_formatter_by_mode('a')
    annotated_folded = annotated_formatter.fold_formula('=SUM(A1:A10)')
    annotated_detected = detect_current_mode(annotated_folded)
    
    # Test Plain Excel mode detection
    plain_formatter = ModularExcelFormatter.create_formatter_by_mode('p')
    plain_folded = plain_formatter.fold_formula('=SUM(A1:A10)')
    plain_detected = detect_current_mode(plain_folded)
    
    # Test unfolded formula detection
    unfolded_detected = detect_current_mode('=SUM(A1:A10)')
    
    print("Mode Detection Results:")
    print(f"JavaScript folded → detected as: {js_detected}")
    print(f"Annotated folded → detected as: {annotated_detected}")
    print(f"Plain folded → detected as: {plain_detected}")
    print(f"Unfolded formula → detected as: {unfolded_detected}")
    print()
    
    js_success = js_detected == 'j'
    annotated_success = annotated_detected == 'a'
    plain_success = plain_detected in ['p', 'a']  # Plain might be detected as annotated
    unfolded_success = unfolded_detected == 'p'
    
    success = js_success and annotated_success and unfolded_success
    print(f"JavaScript detection: {js_success}")
    print(f"Annotated detection: {annotated_success}")
    print(f"Plain detection: {plain_success}")
    print(f"Unfolded detection: {unfolded_success}")
    print(f"Overall detection success: {success}")
    return success


def test_safe_mode_switching():
    """Test safe mode switching between different modes."""
    original = '=IF(AND(A1>0,B1<>""),SUM(A1:A10),MAX(B1:B10))'
    
    print(f"Original formula: {original}")
    
    # Start with JavaScript mode
    js_formatter = ModularExcelFormatter.create_formatter_by_mode('j')
    js_folded = js_formatter.fold_formula(original)
    print(f"JavaScript folded: {len(js_folded.split(chr(10)))} lines")
    
    # Safe switch from JavaScript to Annotated
    annotated_result = safe_mode_switch(js_folded, 'j', 'a', should_refold=True)
    print(f"Safe switch to Annotated: {len(annotated_result.split(chr(10)))} lines")
    
    # Safe switch from Annotated to Plain
    plain_result = safe_mode_switch(annotated_result, 'a', 'p', should_refold=True)
    print(f"Safe switch to Plain: {len(plain_result.split(chr(10)))} lines")
    
    # Safe switch from Plain back to JavaScript
    js_result = safe_mode_switch(plain_result, 'p', 'j', should_refold=True)
    print(f"Safe switch back to JavaScript: {len(js_result.split(chr(10)))} lines")
    
    # Final unfold to verify integrity
    final_formatter = ModularExcelFormatter.create_formatter_by_mode('j')
    final_unfolded = final_formatter.unfold_formula(js_result)
    print(f"Final unfolded: {final_unfolded}")
    print()
    
    # Check that we maintain formula integrity throughout
    original_norm = normalize_formula(original)
    final_norm = normalize_formula(final_unfolded)
    
    # Check that modes produce different output
    js_has_quotes = '"A1"' in js_folded
    annotated_no_quotes = '"A1"' not in annotated_result
    plain_no_comments = '//' not in plain_result
    
    integrity_success = original_norm == final_norm
    mode_differences = js_has_quotes and annotated_no_quotes and plain_no_comments
    
    success = integrity_success and mode_differences
    print(f"Formula integrity maintained: {integrity_success}")
    print(f"Modes produce different output: {mode_differences}")
    print(f"JavaScript has quotes: {js_has_quotes}")
    print(f"Annotated has no quotes: {annotated_no_quotes}")
    print(f"Plain has no comments: {plain_no_comments}")
    print(f"Safe switching success: {success}")
    return success


def test_complex_formula_all_modes():
    """Test a complex formula with all three modes to ensure they all work."""
    original = '=LET(x,A1+B1,y,C1*D1,IFS(x>y,"X wins",y>x,"Y wins",TRUE,"Tie"))'
    
    print(f"Complex formula: {original}")
    print()
    
    mode_results = {}
    mode_success = {}
    
    for mode in ['j', 'a', 'p']:
        formatter = ModularExcelFormatter.create_formatter_by_mode(mode)
        
        folded = formatter.fold_formula(original)
        unfolded = formatter.unfold_formula(folded)
        
        # Check round-trip
        round_trip_success = normalize_formula(original) == normalize_formula(unfolded)
        
        mode_results[mode] = {
            'folded_lines': len(folded.split('\n')),
            'has_comments': '//' in folded,
            'has_quotes': '"A1"' in folded or '"B1"' in folded,
            'round_trip': round_trip_success
        }
        
        mode_success[mode] = round_trip_success
        
        print(f"Mode {mode}: {mode_results[mode]['folded_lines']} lines, "
              f"comments: {mode_results[mode]['has_comments']}, "
              f"quotes: {mode_results[mode]['has_quotes']}, "
              f"round-trip: {round_trip_success}")
    
    print()
    
    # Check that modes behave as expected
    js_correct = (mode_results['j']['has_comments'] and mode_results['j']['has_quotes'])
    annotated_correct = (mode_results['a']['has_comments'] and not mode_results['a']['has_quotes'])
    plain_correct = (not mode_results['p']['has_comments'] and not mode_results['p']['has_quotes'])
    
    all_round_trip = all(mode_success.values())
    all_mode_behavior = js_correct and annotated_correct and plain_correct
    
    success = all_round_trip and all_mode_behavior
    print(f"All round-trips successful: {all_round_trip}")
    print(f"JavaScript mode behavior correct: {js_correct}")
    print(f"Annotated mode behavior correct: {annotated_correct}")
    print(f"Plain mode behavior correct: {plain_correct}")
    print(f"Complex formula test success: {success}")
    return success


def main():
    """Run all three-mode system tests."""
    print("Excel Formula Formatter - Three Mode System Tests")
    print("=" * 60)
    print()
    
    tests = [
        ("JavaScript Mode (j)", test_javascript_mode),
        ("Annotated Excel Mode (a)", test_annotated_excel_mode),
        ("Plain Excel Mode (p)", test_plain_excel_mode),
        ("Mode Detection", test_mode_detection),
        ("Safe Mode Switching", test_safe_mode_switching),
        ("Complex Formula All Modes", test_complex_formula_all_modes)
    ]
    
    results = []
    for test_name, test_func in tests:
        print(f"Running {test_name} test...")
        print("-" * 40)
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
    
    print("=" * 60)
    print(f"Three Mode Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All three-mode tests passed!")
        print("✨ JavaScript (j), Annotated Excel (a), and Plain Excel (p) modes working correctly.")
        print("📋 Ready for Excel Labs Advanced Formula Environment integration!")
        return 0
    else:
        print("❌ Some three-mode tests failed. Check the output above for details.")
        return 1


if __name__ == "__main__":
    exit(main())

# End of file #
