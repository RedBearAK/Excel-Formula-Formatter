#!/usr/bin/env python3
"""
Main entry point for Excel Formula Formatter package.
File: excel_formula_formatter/__main__.py

Allows the package to be run as: python -m excel_formula_formatter
"""

import sys

# Route to appropriate formatter based on arguments
if len(sys.argv) >= 2 and sys.argv[1] in ['javascript', 'python', 'css']:
    # Modular version: python -m excel_formula_formatter javascript auto
    from excel_formula_formatter.modular_excel_formatter import main
else:
    # Original version: python -m excel_formula_formatter auto  
    from excel_formula_formatter.excel_formula_formatter import main


if __name__ == "__main__":
    sys.exit(main())

# End of file #
