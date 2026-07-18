#!/usr/bin/env python3
"""
Main entry point for Excel Formula Formatter package with four-mode support.
File: excel_formula_formatter/__main__.py

Launches the enhanced terminal editor for interactive formula editing.
Perfect for VSCode integration via multi-command extensions.

All CLI operations are handled by the modular formatter. The legacy
formatter was removed in the single-formatter consolidation. Old-style
invocations without a mode letter (e.g. `--cli fold`) are mapped to
Plain Excel mode automatically for backward compatibility.
"""

import sys

from excel_formula_formatter.version_info import get_version_string


_valid_modes = ['j', 'a', 'p', 'c', 'javascript', 'annotated', 'plain', 'compact']
_valid_operations = ['fold', 'unfold', 'auto']


def show_usage():
    """Display usage information."""
    print("Excel Formula Formatter - Interactive Terminal Editor with Four Modes")
    print(f"Version: {get_version_string()}")
    print()
    print("Usage:")
    print("  python -m excel_formula_formatter            Launch Textual editor (default)")
    print("  python -m excel_formula_formatter --rich     Launch Rich prompt-loop editor")
    print("  python -m excel_formula_formatter --cli      Use command-line mode")
    print("  python -m excel_formula_formatter --version  Show version")
    print("  python -m excel_formula_formatter --help     Show this help")
    print()
    print("Textual Editor (default):")
    print("  • Full-screen scrollable editing with alternate screen buffer")
    print("  • Syntax highlighting in JavaScript mode (tree-sitter)")
    print("  • Ctrl+T toggle, Ctrl+F fold, Ctrl+U unfold, Ctrl+N cycle mode")
    print("  • C copies all, P pastes; hold Shift while dragging to mouse-select text")
    print()
    print("Rich Editor (--rich, or fallback when Textual is unavailable):")
    print("  • Four syntax modes: JavaScript (j), Annotated Excel (a), Plain Excel (p), Compact Excel (c)")
    print("  • Safe mode switching while formulas are folded")
    print("  • Paste Excel formulas from clipboard")
    print("  • Toggle between Excel and formatted syntax")
    print("  • Syntax highlighting with Rich")
    print("  • Perfect for VSCode integration and Excel Labs Advanced Formula Environment")
    print("  • Compact mode for 8K character limit optimization")
    print()
    print("Four Modes:")
    print("  j = JavaScript   - JavaScript-like syntax with quoted cells \"A1\"")
    print("  a = Annotated    - Excel syntax with helpful comments and spacing")
    print("  p = Plain        - Pure Excel syntax with smart indenting only (no comments)")
    print("  c = Compact      - Minimal spacing Excel syntax (for 8K character limit)")
    print()
    print("CLI Mode:")
    print("  echo '=SUM(A1:A10)' | python -m excel_formula_formatter --cli j fold")
    print("  echo '=SUM(A1:A10)' | python -m excel_formula_formatter --cli a auto")
    print("  echo '=SUM(A1:A10)' | python -m excel_formula_formatter --cli p fold")
    print("  echo '=SUM(A1:A10)' | python -m excel_formula_formatter --cli c fold")
    print()
    print("  Omitting the mode letter defaults to Plain Excel mode:")
    print("  echo '=SUM(A1:A10)' | python -m excel_formula_formatter --cli fold")
    print()
    print("Compact Mode Benefits:")
    print("  • Removes unnecessary spaces around operators and commas")
    print("  • Preserves spaces inside string literals")
    print("  • Significantly reduces character count for complex formulas")
    print("  • Essential for formulas approaching Excel's 8,192 character limit")


def route_cli_mode() -> int:
    """Normalize CLI args and hand off to the modular formatter's main()."""
    from excel_formula_formatter.modular_excel_formatter import main as modular_main

    # Strip the --cli flag itself
    cli_args = sys.argv[2:]

    # Old-style invocation with only an operation: default to Plain Excel mode
    if cli_args and cli_args[0].lower() in _valid_operations:
        cli_args = ['p'] + cli_args

    # No args at all: Plain Excel mode with auto operation
    if not cli_args:
        cli_args = ['p', 'auto']

    sys.argv = [sys.argv[0]] + cli_args
    return modular_main()


def main():
    """Launch enhanced terminal editor by default, CLI mode on request."""
    # Check for help request
    if len(sys.argv) > 1 and sys.argv[1] in ['-h', '--help', 'help']:
        show_usage()
        return 0

    # Check for version request
    if len(sys.argv) > 1 and sys.argv[1] in ['-V', '--version', 'version']:
        print(get_version_string())
        return 0

    # Check for CLI mode request
    if len(sys.argv) > 1 and sys.argv[1] == '--cli':
        return route_cli_mode()

    # Explicit request for the Rich prompt-loop editor
    force_rich = len(sys.argv) > 1 and sys.argv[1] == '--rich'

    # Default: Textual full-screen editor, falling back to the Rich editor
    if not force_rich:
        try:
            from excel_formula_formatter.textual_excel_editor import main as textual_main
            textual_main()
            return 0
        except ImportError:
            print("📦 Textual not installed (pip install textual) - using Rich editor instead.")
        except Exception as e:
            print(f"⚠️  Textual editor failed ({e}) - falling back to Rich editor.")

    # Fallback (or forced): enhanced Rich prompt-loop editor
    try:
        from excel_formula_formatter.enhanced_simple_excel_editor import main as editor_main
        print(f"🚀 Starting Excel Formula Interactive Editor... [{get_version_string()}]")
        print("💡 Perfect for VSCode integration and Excel Labs Advanced Formula Environment!")
        print("✨ Default mode: Plain Excel (p) - pure Excel syntax with smart indenting")
        print("📝 Four modes available: JavaScript (j), Annotated Excel (a), Plain Excel (p), Compact Excel (c)")
        print("🎯 Starting in edit mode - ready for your formula!")
        print()
        editor_main()
        return 0
    except ImportError as e:
        print(f"❌ Could not launch enhanced editor: {e}")
        print("📦 Install required dependencies: pip install rich 'textual[syntax]'")
        print()
        print("❌ Cannot launch interactive editor without dependencies.")
        return 1
    except Exception as e:
        print(f"❌ Error launching editor: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())

# End of file #
