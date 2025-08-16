#!/usr/bin/env python3
"""
Diagnostic tool to identify missing commas in Excel formulas.
File: comma_diagnostic_tool.py
"""

import re


def find_missing_commas(formula_text: str) -> list:
    """Find patterns that suggest missing commas."""
    issues = []
    lines = formula_text.split('\n')
    
    patterns_to_check = [
        # Function name directly followed by another function name (missing comma)
        (r'([A-Z_][A-Z0-9_]*)\s*([A-Z_][A-Z0-9_]*)\s*\(', 
         'Function name followed by another function (missing comma?)'),
        
        # Cell reference followed by function name (missing comma)
        (r'([A-Z]+\d+)\s*([A-Z_][A-Z0-9_]*)\s*\(', 
         'Cell reference followed by function (missing comma?)'),
        
        # Closing paren followed by opening paren without comma
        (r'\)\s*([A-Z_][A-Z0-9_]*)\s*\(', 
         'Function call followed by another function (missing comma?)'),
        
        # Variable name followed by function call (common in LET)
        (r'([a-zA-Z_][a-zA-Z0-9_]*)\s*(NOT|AND|OR|IF|SUM|LEN)\s*\(', 
         'Variable followed by function (missing comma?)'),
        
        # Two NOT functions without comma
        (r'NOT\s*\(\s*[^)]+\)\s*NOT\s*\(', 
         'Multiple NOT functions (missing comma?)'),
    ]
    
    for line_num, line in enumerate(lines, 1):
        for pattern, description in patterns_to_check:
            matches = re.finditer(pattern, line, re.IGNORECASE)
            for match in matches:
                issues.append({
                    'line': line_num,
                    'position': match.start(),
                    'text': match.group(0),
                    'description': description,
                    'line_content': line.strip()
                })
    
    return issues


def suggest_fixes(formula_text: str) -> str:
    """Suggest fixes for common missing comma patterns."""
    fixed = formula_text
    
    # Fix common patterns
    fixes = [
        # Variable name followed by function call
        (r'([a-zA-Z_][a-zA-Z0-9_]*)\s*(NOT|AND|OR|IF|SUM|LEN)\s*\(', 
         r'\1, \2('),
        
        # Function followed by another function
        (r'([A-Z_][A-Z0-9_]*)\s*\(\s*[^)]+\)\s*(NOT|AND|OR|IF|SUM|LEN)\s*\(', 
         r'\1(...), \2('),  # Simplified - would need more complex logic for real fixes
    ]
    
    for pattern, replacement in fixes:
        fixed = re.sub(pattern, replacement, fixed, flags=re.IGNORECASE)
    
    return fixed


def main():
    """Analyze the paste.txt content for missing commas."""
    # Read the paste.txt content
    try:
        with open('paste.txt', 'r') as f:
            formula_content = f.read()
    except FileNotFoundError:
        print("❌ paste.txt not found")
        return 1
    
    print("Comma Diagnostic Tool")
    print("=" * 50)
    print(f"Analyzing formula content ({len(formula_content)} characters)")
    print()
    
    # Find potential issues
    issues = find_missing_commas(formula_content)
    
    if issues:
        print(f"Found {len(issues)} potential comma issues:")
        print()
        
        for i, issue in enumerate(issues, 1):
            print(f"{i}. Line {issue['line']}, position {issue['position']}:")
            print(f"   Pattern: {issue['text']}")
            print(f"   Issue: {issue['description']}")
            print(f"   Context: ...{issue['line_content'][max(0, issue['position']-20):issue['position']+40]}...")
            print()
    else:
        print("✅ No obvious missing comma patterns detected")
    
    # Show specific problematic lines from paste.txt
    problematic_lines = [
        'has_Paid_DateNOT( has_Price )',
        'NOT( has_Paid_Date )NOT( has_WBL_Date )',
        'has_Paid_DateNOT( has_Price )',
    ]
    
    print("\nKnown problematic patterns in paste.txt:")
    print("-" * 40)
    
    for line in formula_content.split('\n'):
        for prob in problematic_lines:
            if prob in line.replace(' ', ''):
                print(f"❌ {line.strip()}")
                # Show what it should be
                suggested = line
                suggested = suggested.replace('has_Paid_DateNOT', 'has_Paid_Date, NOT')
                suggested = suggested.replace(')NOT(', '), NOT(')
                if suggested != line:
                    print(f"✅ Should be: {suggested.strip()}")
                print()
    
    return 0


if __name__ == "__main__":
    exit(main())

# End of file #
