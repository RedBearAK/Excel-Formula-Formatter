#!/usr/bin/env python3
"""
Headless tests for the four-mode Textual editor using Textual's pilot.
File: tests/test_textual_editor.py

Verifies boot, fold/unfold in the editor, mode cycling with formula
integrity, highlighting language switching, and error surfacing.
Clipboard actions are not tested here: they need a real display/session.
"""

import sys
import asyncio

from pathlib import Path

# Add the parent directory (containing the package) to Python path
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from excel_formula_formatter.formula_text_guards import normalize_formula_for_comparison
from excel_formula_formatter.textual_excel_editor import TextualExcelEditor, mode_order


test_formula = '=LET(x,A1*2,IFS(x>10,"Big",TRUE,"Small"))'


def test_boot_and_fold():
    """Editor boots and folds a formula via the toggle action."""
    async def run():
        from textual.widgets import TextArea

        app = TextualExcelEditor()
        async with app.run_test(size=(110, 32)) as pilot:
            editor = app.query_one('#editor', TextArea)
            editor.text = test_formula
            app.action_toggle_formula()
            await pilot.pause()
            return editor.text

    folded = asyncio.run(run())
    print(f"Folded output ({len(folded.splitlines())} lines):")
    print(folded)

    got_multiline = len(folded.splitlines()) > 3
    round_trip_ready = 'LET(' in folded and 'IFS(' in folded

    success = got_multiline and round_trip_ready
    print(f"Multi-line fold: {got_multiline}")
    print(f"Structure present: {round_trip_ready}")
    print(f"Overall success: {success}")
    return success


def test_mode_cycle_integrity():
    """Cycling through all four modes preserves the formula exactly."""
    async def run():
        from textual.widgets import TextArea

        app = TextualExcelEditor()
        async with app.run_test(size=(110, 32)) as pilot:
            editor = app.query_one('#editor', TextArea)
            editor.text = test_formula
            app.action_fold_formula()
            await pilot.pause()

            seen_modes = [app.formatter_mode]
            languages = {app.formatter_mode: editor.language}
            doc_classes = {app.formatter_mode: type(editor.document).__name__}
            for _step in range(len(mode_order) - 1):
                app.action_cycle_mode()
                await pilot.pause()
                seen_modes.append(app.formatter_mode)
                languages[app.formatter_mode] = editor.language
                doc_classes[app.formatter_mode] = type(editor.document).__name__

            app.action_unfold_formula()
            await pilot.pause()
            return editor.text, seen_modes, languages, doc_classes

    final_text, seen_modes, languages, doc_classes = asyncio.run(run())
    print(f"Modes visited: {seen_modes}")
    print(f"Languages: {languages}")
    print(f"Final unfolded: {final_text}")

    visited_all = sorted(seen_modes) == sorted(mode_order)
    js_highlighted = languages.get('j') == 'javascript'
    # All modes use the tolerant JavaScript grammar for colorization.
    # The language attribute alone proves nothing (Textual accepts it
    # without tree-sitter and silently skips highlighting), so assert
    # the document class: SyntaxAwareDocument means real highlighting.
    excel_highlighted = all(
        languages.get(m) == 'javascript' for m in ['a', 'p', 'c'] if m in languages
    )
    really_highlighted = all(
        cls == 'SyntaxAwareDocument' for cls in doc_classes.values()
    )
    integrity = (
        normalize_formula_for_comparison(final_text)
        == normalize_formula_for_comparison(test_formula)
    )

    success = (visited_all and js_highlighted and excel_highlighted
               and really_highlighted and integrity)
    print(f"Visited all modes: {visited_all}")
    print(f"JS mode highlighted: {js_highlighted}")
    print(f"Excel modes highlighted (tolerant JS): {excel_highlighted}")
    print(f"Document classes: {doc_classes}")
    print(f"Real highlighting (SyntaxAwareDocument): {really_highlighted}")
    print(f"Formula integrity: {integrity}")
    print(f"Overall success: {success}")
    return success


def test_malformed_input_surfaces_error():
    """Malformed formulas show a status error without crashing."""
    async def run():
        from textual.widgets import Static, TextArea

        app = TextualExcelEditor()
        async with app.run_test(size=(110, 32)) as pilot:
            editor = app.query_one('#editor', TextArea)
            editor.text = '=IF(A1="broken'
            app.action_fold_formula()
            await pilot.pause()
            status_text = str(app.query_one('#status', Static).render())
            return editor.text, status_text

    remaining_text, status_text = asyncio.run(run())
    print(f"Status shown: {status_text}")

    error_surfaced = 'error' in status_text.lower()
    has_ui_mode_prefix = status_text.startswith(('COMMAND', 'EDIT'))
    text_untouched = remaining_text == '=IF(A1="broken'

    success = error_surfaced and has_ui_mode_prefix and text_untouched
    print(f"Error surfaced in status: {error_surfaced}")
    print(f"UI mode prefix present: {has_ui_mode_prefix}")
    print(f"Editor text untouched: {text_untouched}")
    print(f"Overall success: {success}")
    return success


def test_version_in_subtitle():
    """The version string appears in the app subtitle."""
    async def run():
        app = TextualExcelEditor()
        async with app.run_test(size=(110, 32)):
            return app.sub_title

    from excel_formula_formatter.version_info import get_version_string

    sub_title = asyncio.run(run())
    print(f"Subtitle: {sub_title}")

    success = get_version_string() in sub_title
    print(f"Version present: {success}")
    print(f"Overall success: {success}")
    return success


def test_modal_single_letter_keys():
    """Command-mode letters act; edit-mode letters type; Escape returns."""
    async def run():
        from textual.widgets import TextArea

        app = TextualExcelEditor()
        async with app.run_test(size=(110, 32)) as pilot:
            editor = app.query_one('#editor', TextArea)
            started_in_command = not app.in_edit_mode()

            await pilot.press('e')
            edit_after_e = app.in_edit_mode()
            for ch in '=SUM(A1:A10,B1:B10)':
                await pilot.press(ch)
            await pilot.press('t')  # must TYPE, not toggle
            typed_t = editor.text.endswith('t')
            await pilot.press('backspace')

            await pilot.press('escape')
            command_after_escape = not app.in_edit_mode()

            await pilot.press('t')  # must TOGGLE now
            await pilot.pause()
            folded_lines = len(editor.text.splitlines())

            return (started_in_command, edit_after_e, typed_t,
                    command_after_escape, folded_lines)

    (started_in_command, edit_after_e, typed_t,
     command_after_escape, folded_lines) = asyncio.run(run())

    print(f"Starts in COMMAND mode: {started_in_command}")
    print(f"E enters EDIT mode: {edit_after_e}")
    print(f"Letters type in EDIT mode: {typed_t}")
    print(f"Escape returns to COMMAND: {command_after_escape}")
    print(f"T folds in COMMAND mode ({folded_lines} lines): {folded_lines > 2}")

    success = (started_in_command and edit_after_e and typed_t
               and command_after_escape and folded_lines > 2)
    print(f"Overall success: {success}")
    return success


def main():
    """Run all Textual editor tests."""
    print("Textual Editor Tests (headless pilot)")
    print("=" * 60)
    print()

    tests = [
        ("Boot and Fold", test_boot_and_fold),
        ("Mode Cycle Integrity", test_mode_cycle_integrity),
        ("Malformed Input Surfaces Error", test_malformed_input_surfaces_error),
        ("Version in Subtitle", test_version_in_subtitle),
        ("Modal Single-Letter Keys", test_modal_single_letter_keys),
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
    print(f"Textual Editor Test Results: {passed}/{total} tests passed")

    if passed == total:
        print("🎉 All Textual editor tests passed!")
        return 0
    else:
        print("❌ Some Textual editor tests failed. Check the output above for details.")
        return 1


if __name__ == "__main__":
    sys.exit(main())

# End of file #
