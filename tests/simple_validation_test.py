#!/usr/bin/env python3
"""
Simple validation test for the three-mode system.
File: simple_validation_test.py
"""

import sys
from pathlib import Path

# Add package to path  
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter


def main():
    print("Simple Validation Test for Three Modes")
    print("="*50)
    
    original = '=IF(A1<>B1,SUM(A1:A10),"Equal")'
    print(f"Original: {original}")
    print()
    
    modes = {
        'j': 'JavaScript',
        'a': 'Annotated Excel', 
        'p': 'Plain Excel'
    }
    
    for mode_code, mode_name in modes.items():
        print(f"{mode_name} Mode ({mode_code}):")
        print("-" * 25)
        
        try:
            formatter = ModularExcelFormatter.create_formatter_by_mode(mode_code)
            folded = formatter.fold_formula(original)
            unfolded = formatter.unfold_formula(folded)
            
            print(f"Folded has comments: {'//' in folded}")
            print(f"Folded has quotes: {'\"A1\"' in folded}")
            print(f"Unfolded: {unfolded}")
            print(f"Unfolded length: {len(unfolded)}")
            
            # Round trip test
            orig_norm = original.replace(' ', '').replace('=', '')
            unfold_norm = unfolded.replace(' ', '').replace('=', '')
            round_trip_ok = orig_norm == unfold_norm
            print(f"Round trip OK: {round_trip_ok}")
            
            if not round_trip_ok:
                print(f"  Original normalized: {orig_norm}")
                print(f"  Unfolded normalized: {unfold_norm}")
            
        except Exception as e:
            print(f"ERROR: {e}")
        
        print()
    
    print("Validation complete!")


if __name__ == "__main__":
    main()

# End of file #
