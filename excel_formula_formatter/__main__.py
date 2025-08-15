#!/usr/bin/env python3
"""
Main entry point for Excel Formula Formatter package.
File: excel_formula_formatter/__main__.py

Launches the enhanced terminal editor for interactive formula editing.
Perfect for VSCode integration via multi-command extensions.
"""

import sys


def show_usage():
    """Display usage information."""
    print("Excel Formula Formatter - Interactive Terminal Editor")
    print()
    print("Usage:")
    print("  python -m excel_formula_formatter          Launch interactive editor (default)")
    print("  python -m excel_formula_formatter --cli    Use command-line mode")
    print("  python -m excel_formula_formatter --help   Show this help")
    print()
    print("Interactive Editor Features:")
    print("  • Paste Excel formulas from clipboard")
    print("  • Toggle between Excel and JavaScript-like syntax")
    print("  • Syntax highlighting with Rich")
    print("  • Multiple output modes (JavaScript, Plain Excel)")
    print("  • Perfect for VSCode integration")
    print()
    print("CLI Mode:")
    print("  echo '=SUM(A1:A10)' | python -m excel_formula_formatter --cli javascript fold")


def main():
    """Launch enhanced terminal editor by default, CLI mode on request."""
    # Check for help request
    if len(sys.argv) > 1 and sys.argv[1] in ['-h', '--help', 'help']:
        show_usage()
        return 0
    
    # Check for CLI mode request
    if len(sys.argv) > 1 and sys.argv[1] == '--cli':
        # Remove --cli from args and route to CLI formatter
        sys.argv = [sys.argv[0]] + sys.argv[2:]  # Remove --cli
        
        if len(sys.argv) >= 2 and sys.argv[1] in ['javascript', 'plain']:
            # Modular CLI version
            from excel_formula_formatter.modular_excel_formatter import main as modular_main
            return modular_main()
        else:
            # Original CLI version for backward compatibility
            from excel_formula_formatter.excel_formula_formatter import main as original_main
            return original_main()
    
    # Default: Launch enhanced terminal editor
    try:
        from excel_formula_formatter.enhanced_simple_excel_editor import main as editor_main
        print("🚀 Launching Excel Formula Interactive Editor...")
        print("💡 Perfect for VSCode integration via multi-command extensions!")
        print()
        editor_main()
        return 0
    except ImportError as e:
        print(f"❌ Could not launch enhanced editor: {e}")
        print("📦 Install required dependency: pip install rich")
        print()
        print("💡 Alternative: Use the basic Rich editor instead:")
        print("   python -c \"from excel_formula_formatter.simple_excel_editor import main; main()\"")
        print()
        print("❌ Cannot launch interactive editor without dependencies.")
        return 1
    except Exception as e:
        print(f"❌ Error launching editor: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())

# End of file #
