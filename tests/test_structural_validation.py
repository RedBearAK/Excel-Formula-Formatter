#!/usr/bin/env python3
"""
Structural validation tests: missing commas and operators between value
tokens must be detected, and clean formulas must produce no warnings.
File: tests/test_structural_validation.py
"""

import sys

from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.modular_excel_formatter import (
    ModularExcelFormatter,
    find_structural_issues,
)


def test_missing_comma_detected():
    """Removing a comma between LET pair lines is flagged after unfold."""
    formatter = ModularExcelFormatter.create_plain_formatter()
    original = '=LET(x,10,y,20,x+y)'

    folded = formatter.fold_formula(original)
    broken = folded.replace('x, 10,', 'x, 10', 1)
    unfolded = formatter.unfold_formula(broken)
    issues = find_structural_issues(unfolded)

    print(f"Broken unfold: {unfolded}")
    for issue in issues:
        print(f"  issue: {issue['message']}")

    detected = len(issues) >= 1
    names_right = detected and issues[0]['left'] == '10' and issues[0]['right'] == 'y'

    success = detected and names_right
    print(f"Issue detected: {detected}")
    print(f"Correct tokens named: {names_right}")
    print(f"Overall success: {success}")
    return success


def test_adjacency_shapes():
    """Various value-value adjacencies are all flagged."""
    cases = [
        ('=SUM(A1:A10)total+1', "paren then identifier"),
        ('=CONCAT("a" "b")', "string then string"),
        ('=IF(A1>0 5,10)', "number after expression"),
        ('=LET(x,1,x "done")', "identifier then string"),
    ]

    overall = True
    for formula, label in cases:
        issues = find_structural_issues(formula)
        found = len(issues) >= 1
        overall = overall and found
        print(f"  {label:28s} {'flagged' if found else 'MISSED'}: {formula}")

    print(f"Overall success: {overall}")
    return overall


def test_intersection_flagged_as_possible():
    """Adjacent references warn but mention the intersection operator."""
    issues = find_structural_issues('=SUM(A1:B10 C5:D20)')

    flagged = len(issues) == 1
    mentions_intersection = flagged and 'intersection' in issues[0]['message']

    if flagged:
        print(f"Message: {issues[0]['message']}")

    success = flagged and mentions_intersection
    print(f"Flagged: {flagged}")
    print(f"Mentions intersection: {mentions_intersection}")
    print(f"Overall success: {success}")
    return success


def test_clean_formulas_produce_no_issues():
    """Valid formulas, including complex production shapes, stay clean."""
    clean_formulas = [
        '=SUM(A1:A10,B1:B10)',
        '=LET(x,A1*2,IFS(x>10,"Big",TRUE,"Small"))',
        '=IF(A1<>"x <> y",A1&" & "&B1,"")',
        '=LET(nz,LAMBDA(n,AND(LEN(n)>0,ISNUMBER(n))),v,INDIRECT("O"&ROW()),'
        'IFS(nz(v),HSTACK(v+5,"tag"),TRUE,HSTACK("","none")))',
        '=SUM(Table1[Sales Amount])+{1,2;3,4}',
        "='Q1 Data'!A1+$B$2*50%",
    ]

    overall = True
    for formula in clean_formulas:
        issues = find_structural_issues(formula)
        clean = len(issues) == 0
        overall = overall and clean
        marker = 'clean' if clean else f'{len(issues)} FALSE POSITIVE(S)'
        print(f"  {marker:24s} {formula[:56]}")
        for issue in issues:
            print(f"      {issue['message']}")

    print(f"Overall success: {overall}")
    return overall


def test_malformed_input_reported_not_raised():
    """Unclosed constructs surface as an issue, not an exception."""
    issues = find_structural_issues('=IF(A1="broken')

    reported = len(issues) == 1 and 'Unclosed' in issues[0]['message']
    print(f"Issue: {issues[0]['message'] if issues else 'none'}")
    print(f"Overall success: {reported}")
    return reported


def test_let_pair_shift_detected():
    """A shifted LET (cascade from a lost comma) is caught by shape checks
    even when the shifted text is grammatically clean."""
    even_issues = find_structural_issues('=LET(x,1,y,2,x+y,z)')
    even_flagged = any('even number' in i['message'] for i in even_issues)

    slot_issues = find_structural_issues('=LET(x,1,2,3,x)')
    slot_flagged = any('name slot' in i['message'] for i in slot_issues)

    # Fully cascaded round-trip: fold, drop a comma, unfold, refold,
    # unfold again - the final text has all commas yet is still caught
    formatter = ModularExcelFormatter.create_plain_formatter()
    original = '=LET(aa,1,bb,2,cc,3,aa+bb+cc)'
    folded = formatter.fold_formula(original)
    broken = folded.replace('aa, 1,', 'aa, 1', 1)
    cascaded = formatter.unfold_formula(
        formatter.fold_formula(formatter.unfold_formula(broken))
    )
    cascade_issues = find_structural_issues(cascaded)
    cascade_flagged = any('LET' in i['message'] for i in cascade_issues)

    print(f"Even-count flagged: {even_flagged}")
    print(f"Name-slot flagged: {slot_flagged}")
    print(f"Cascaded text: {cascaded}")
    print(f"Cascade flagged: {cascade_flagged}")

    success = even_flagged and slot_flagged and cascade_flagged
    print(f"Overall success: {success}")
    return success


def main():
    """Run all structural validation tests."""
    print("Structural Validation Tests")
    print("=" * 60)
    print()

    tests = [
        ("Missing Comma Detected", test_missing_comma_detected),
        ("Adjacency Shapes", test_adjacency_shapes),
        ("Intersection Flagged as Possible", test_intersection_flagged_as_possible),
        ("Clean Formulas No Issues", test_clean_formulas_produce_no_issues),
        ("Malformed Input Reported", test_malformed_input_reported_not_raised),
        ("LET Pair Shift Detected", test_let_pair_shift_detected),
    ]

    results = []
    for test_name, test_func in tests:
        print(f"Running {test_name} test...")
        print("-" * 40)
        try:
            success = test_func()
            results.append(success)
            print(f"{'✓' if success else '✗'} {test_name}: {'PASS' if success else 'FAIL'}")
        except Exception as e:
            print(f"✗ {test_name}: ERROR - {e}")
            results.append(False)
        print()

    passed = sum(results)
    total = len(results)

    print("=" * 60)
    print(f"Structural Validation Results: {passed}/{total} tests passed")

    if passed == total:
        print("🎉 All structural validation tests passed!")
        return 0
    else:
        print("❌ Some validation tests failed. Check the output above for details.")
        return 1


if __name__ == "__main__":
    sys.exit(main())

# End of file #
