#!/usr/bin/env python3
"""
Main entry point for Excel Formula Formatter package with three-mode support.
File: excel_formula_formatter/__main__.py

Launches the enhanced terminal editor for interactive formula editing.
Perfect for VSCode integration via multi-command extensions.
"""

import sys


def show_usage():
    """Display usage information."""
    print("Excel Formula Formatter - Interactive Terminal Editor with Three Modes")
    print()
    print("Usage:")
    print("  python -m excel_formula_formatter          Launch interactive editor (default)")
    print("  python -m excel_formula_formatter --cli    Use command-line mode")
    print("  python -m excel_formula_formatter --help   Show this help")
    print()
    print("Interactive Editor Features:")
    print("  • Three syntax modes: JavaScript (j), Annotated Excel (a), Plain Excel (p)")
    print("  • Safe mode switching while formulas are folded")
    print("  • Paste Excel formulas from clipboard")
    print("  • Toggle between Excel and formatted syntax")
    print("  • Syntax highlighting with Rich")
    print("  • Perfect for VSCode integration and Excel Labs Advanced Formula Environment")
    print()
    print("Three Modes:")
    print("  j = JavaScript   - JavaScript-like syntax with quoted cells \"A1\"")
    print("  a = Annotated    - Excel syntax with helpful comments and spacing")
    print("  p = Plain        - Pure Excel syntax with smart indenting only (no comments)")
    print()
    print("CLI Mode:")
    print("  echo '=SUM(A1:A10)' | python -m excel_formula_formatter --cli j fold")
    print("  echo '=SUM(A1:A10)' | python -m excel_formula_formatter --cli a auto")
    print("  echo '=SUM(A1:A10)' | python -m excel_formula_formatter --cli p fold")


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
        
        if len(sys.argv) >= 2 and sys.argv[1] in ['j', 'a', 'p', 'javascript', 'plain']:
            # New modular CLI version with three modes
            from excel_formula_formatter.modular_excel_formatter import main as modular_main
            return modular_main()
        else:
            # Original CLI version for backward compatibility
            from excel_formula_formatter.excel_formula_formatter import main as original_main
            return original_main()
    
    # Default: Launch enhanced terminal editor
    try:
        from excel_formula_formatter.enhanced_simple_excel_editor import main as editor_main
        print("🚀 Starting Excel Formula Interactive Editor...")
        print("💡 Perfect for VSCode integration and Excel Labs Advanced Formula Environment!")
        print("✨ Default mode: Plain Excel (p) - pure Excel syntax with smart indenting")
        print("📝 Other modes: JavaScript (j), Annotated Excel (a)")
        print("🎯 Starting in edit mode - ready for your formula!")
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