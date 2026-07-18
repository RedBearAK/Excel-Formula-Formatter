#!/usr/bin/env python3
"""
Four-mode Textual Excel formula editor with modal single-letter keys.
File: excel_formula_formatter/textual_excel_editor.py

Two UI modes, mirroring the Rich editor's ergonomics:

  COMMAND mode (default): single letters act immediately, exactly like
      the Rich prompt-loop editor -
      T toggle, F fold, U unfold, M cycle formatter mode,
      C copy all, P paste (replace), X clear, E edit text, Q quit
  EDIT mode: keys type into the scrollable TextArea;
      Escape returns to COMMAND mode.

Ctrl+T / Ctrl+F / Ctrl+U / Ctrl+Q also work from either mode.
No button row: the Footer shows the live key bindings instead.

Requires: pip install textual
"""

import sys
from pathlib import Path

# Add package to path if needed
package_parent = Path(__file__).parent.parent
sys.path.insert(0, str(package_parent))

from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.widgets import Footer, Header, Static, TextArea

try:
    from excel_formula_formatter.version_info import get_version_string
    from excel_formula_formatter.clipboard_helpers import (
        get_clipboard_text,
        set_clipboard_text,
    )
    from excel_formula_formatter.modular_excel_formatter import (
        ModularExcelFormatter,
        auto_format_with_mode,
        detect_current_mode,
        find_structural_issues,
        safe_mode_switch,
    )
except ImportError:
    print("Excel Formula Formatter package not found. Please install or check PYTHONPATH.")
    sys.exit(1)


mode_order = ['p', 'j', 'a', 'c']

mode_names = {
    'j': "JavaScript",
    'a': "Annotated Excel",
    'p': "Plain Excel",
    'c': "Compact Excel",
}

# Tree-sitter language per mode. There is no tree-sitter CoffeeScript
# grammar (VSCode's CoffeeScript colorizer is TextMate-based), but the
# JavaScript grammar's error recovery colorizes Excel syntax usefully -
# strings, numbers, and function calls light up even where the parse is
# imperfect - so all modes use it. Set an entry to None to disable.
mode_languages = {
    'j': "javascript",
    'a': "javascript",
    'p': "javascript",
    'c': "javascript",
}

# One-line explanations shown in the status bar when cycling modes.
# Plain and Compact fold almost identically; they differ mainly in the
# UNFOLDED output: readable spacing vs minimal characters.
mode_descriptions = {
    'j': "cell refs backticked for highlighting; unfolds to minimal-space Excel",
    'a': "Excel syntax with // comments; unfolds with readable spacing",
    'p': "pure Excel, no comments; unfolds with readable spacing",
    'c': "pure Excel, no comments; unfolds with MINIMAL spacing (8K limit)",
}


def highlighting_available() -> bool:
    """True when tree-sitter grammars are actually importable.

    Textual accepts language assignments even without tree-sitter
    installed and silently renders unhighlighted, so checking the
    language attribute proves nothing - probe the real dependency.
    """
    try:
        import tree_sitter
        import tree_sitter_javascript
        return True
    except ImportError:
        return False


class TextualExcelEditor(App):
    """Excel Formula Editor with command/edit modes and scrollable text."""

    TITLE = "Excel Formula Editor"
    SUB_TITLE = f"v{get_version_string()}"

    # Do not auto-focus the TextArea: the app starts in COMMAND mode
    AUTO_FOCUS = None

    # The command palette (theme switching etc.) opens with its default
    # Ctrl+P (VSCode/Obsidian convention) or ':' in COMMAND mode; paste
    # is plain P. The check_action guard releases Escape/keys to the
    # palette while it's open.

    CSS = """
    Screen {
        background: $surface;
    }

    .editor-area {
        height: 1fr;
        border: solid $primary;
        margin: 0 1;
    }

    .editor-area:focus {
        border: solid $success;
    }

    .status {
        dock: bottom;
        height: 1;
        background: $accent;
        color: $text;
        content-align: center middle;
    }
    """

    BINDINGS = [
        # Command-mode single letters (inactive while the TextArea has
        # focus, because the TextArea consumes printable keys there)
        Binding("t", "toggle_formula", "Toggle", show=True),
        Binding("f", "fold_formula", "Fold", show=True),
        Binding("u", "unfold_formula", "Unfold", show=True),
        Binding("m", "cycle_mode", "Mode", show=True),
        Binding("c", "copy_all", "Copy", show=True),
        Binding("p", "paste_replace", "Paste", show=True),
        Binding("x", "clear_editor", "Clear", show=False),
        Binding("e", "enter_edit_mode", "Edit", show=True),
        Binding("i", "enter_edit_mode", "Edit", show=False),
        Binding("q", "quit", "Quit", show=True),
        # NOTE: binding any key to command_palette replaces Textual's
        # built-in Ctrl+P for that action, so both keys are declared
        Binding("ctrl+p", "command_palette", "Palette", show=False),
        Binding("colon", "command_palette", "Palette", show=False),
        # Escape leaves EDIT mode even though the TextArea is focused
        Binding("escape", "enter_command_mode", "Command Mode", show=True, priority=True),
        # Ctrl variants work from either mode
        Binding("ctrl+t", "toggle_formula", "Toggle", show=False, priority=True),
        Binding("ctrl+f", "fold_formula", "Fold", show=False, priority=True),
        Binding("ctrl+u", "unfold_formula", "Unfold", show=False, priority=True),
        Binding("ctrl+q", "quit", "Quit", show=False, priority=True),
    ]


    def check_action(self, action: str, parameters) -> bool:
        """Disable app-level command keys while any modal screen is open
        so the modal receives Escape/Enter/letters itself."""
        guarded_actions = {
            'toggle_formula', 'fold_formula', 'unfold_formula',
            'cycle_mode', 'copy_all', 'paste_replace', 'clear_editor',
            'enter_edit_mode', 'enter_command_mode',
        }
        if action in guarded_actions and len(self.screen_stack) > 1:
            return False
        return True

    def __init__(self, start_mode: str = 'p'):
        super().__init__()
        self.formatter_mode = start_mode if start_mode in mode_order else 'p'

    def compose(self) -> ComposeResult:
        """Compose the UI layout."""
        yield Header(show_clock=True)
        yield TextArea.code_editor(
            text="",
            language=mode_languages[self.formatter_mode],
            id="editor",
            classes="editor-area",
        )
        yield Static("Ready", classes="status", id="status")
        yield Footer()

    def on_mount(self) -> None:
        """Start in COMMAND mode with mode info displayed."""
        self.action_enter_command_mode()
        if highlighting_available():
            self.update_status(
                f"COMMAND mode - {mode_names[self.formatter_mode]} "
                f"({self.formatter_mode}) - P pastes a formula, T toggles, E edits"
            )
        else:
            self.update_status(
                "Highlighting OFF - run: pip install 'textual[syntax]' - "
                "P paste, T toggle, E edit"
            )

    # ------------------------------------------------------------------
    # Widget access and display helpers

    def get_editor(self) -> TextArea:
        return self.query_one("#editor", TextArea)

    def in_edit_mode(self) -> bool:
        return self.focused is self.get_editor()

    def update_status(self, message: str) -> None:
        ui_mode = "EDIT" if self.in_edit_mode() else "COMMAND"
        # No square brackets: Static interprets them as Rich markup
        self.query_one("#status", Static).update(f"{ui_mode} · {message}")

    def refresh_mode_display(self) -> None:
        """Sync highlighting language and subtitle with formatter mode."""
        self.get_editor().language = mode_languages[self.formatter_mode]
        self.sub_title = (
            f"v{get_version_string()} - {mode_names[self.formatter_mode]}"
        )

    # ------------------------------------------------------------------
    # UI mode switching

    def action_enter_edit_mode(self) -> None:
        self.get_editor().focus()
        self.update_status("Editing - Escape returns to COMMAND mode")

    def action_enter_command_mode(self) -> None:
        self.set_focus(None)
        self.refresh_mode_display()
        self.update_status(
            "T toggle | F fold | U unfold | M mode | C copy | P paste | E edit | Q quit"
        )


    def structure_warning(self, excel_text: str) -> str:
        """Return a status-line warning for structural issues, or ''."""
        issues = find_structural_issues(excel_text)
        if not issues:
            return ""
        first = issues[0]['message']
        extra = f" (+{len(issues) - 1} more)" if len(issues) > 1 else ""
        return f" | ⚠ {first}{extra}"

    # ------------------------------------------------------------------
    # Formula actions

    def action_fold_formula(self) -> None:
        editor = self.get_editor()
        if not editor.text.strip():
            self.update_status("Nothing to fold")
            return

        formatter = ModularExcelFormatter.create_formatter_by_mode(self.formatter_mode)
        try:
            warning = self.structure_warning(editor.text)
            editor.text = formatter.fold_formula(editor.text)
            editor.move_cursor((0, 0))
            self.update_status(f"Folded ({mode_names[self.formatter_mode]}){warning}")
        except ValueError as e:
            self.update_status(f"Formula error: {e}")
        except Exception as e:
            self.update_status(f"Fold error: {e}")

    def action_unfold_formula(self) -> None:
        editor = self.get_editor()
        if not editor.text.strip():
            self.update_status("Nothing to unfold")
            return

        formatter = ModularExcelFormatter.create_formatter_by_mode(self.formatter_mode)
        try:
            editor.text = formatter.unfold_formula(editor.text)
            editor.move_cursor((0, 0))
            warning = self.structure_warning(editor.text)
            self.update_status(f"Unfolded to Excel formula{warning}")
        except ValueError as e:
            self.update_status(f"Formula error: {e}")
        except Exception as e:
            self.update_status(f"Unfold error: {e}")

    def action_toggle_formula(self) -> None:
        """Auto-detect direction and fold or unfold accordingly."""
        editor = self.get_editor()
        text = editor.text
        if not text.strip():
            self.update_status("Nothing to toggle")
            return

        try:
            was_single_line = len(text.strip().split('\n')) == 1
            editor.text = auto_format_with_mode(text, self.formatter_mode)
            editor.move_cursor((0, 0))
            action = "Folded" if was_single_line else "Unfolded"
            excel_side = text if was_single_line else editor.text
            warning = self.structure_warning(excel_side)
            self.update_status(f"{action} ({mode_names[self.formatter_mode]}){warning}")
        except ValueError as e:
            self.update_status(f"Formula error: {e}")
        except Exception as e:
            self.update_status(f"Toggle error: {e}")

    def action_cycle_mode(self) -> None:
        """Switch to the next mode, safely re-folding folded content."""
        old_mode = self.formatter_mode
        new_mode = mode_order[(mode_order.index(old_mode) + 1) % len(mode_order)]

        editor = self.get_editor()
        text = editor.text

        if text.strip():
            is_folded = len(text.strip().split('\n')) > 1
            detected = detect_current_mode(text) if is_folded else old_mode
            try:
                editor.text = safe_mode_switch(
                    text, detected, new_mode, should_refold=is_folded
                )
                editor.move_cursor((0, 0))
            except Exception as e:
                self.update_status(f"Mode switch error: {e}")
                return

        self.formatter_mode = new_mode
        self.refresh_mode_display()
        self.update_status(
            f"Mode: {mode_names[new_mode]} ({new_mode}) - "
            f"{mode_descriptions[new_mode]}"
        )

    # ------------------------------------------------------------------
    # Clipboard actions

    def action_copy_all(self) -> None:
        """Copy full editor content via OSC 52 and the system clipboard."""
        editor = self.get_editor()
        text = editor.text
        if not text.strip():
            self.update_status("Nothing to copy")
            return

        # Terminal-level copy (works over SSH, needs terminal support)
        osc52_ok = True
        try:
            self.copy_to_clipboard(text)
        except Exception:
            osc52_ok = False

        # System clipboard via session-aware tools
        system_ok = set_clipboard_text(text)

        if system_ok or osc52_ok:
            self.update_status(f"Copied {len(text)} chars to clipboard")
        else:
            self.update_status("Copy failed - no clipboard tool available")

    def action_paste_replace(self) -> None:
        """Replace editor content with the system clipboard."""
        pasted = get_clipboard_text()
        if not pasted:
            self.update_status(
                "Nothing to paste (in EDIT mode, terminal paste inserts at cursor)"
            )
            return

        editor = self.get_editor()
        editor.text = pasted
        editor.move_cursor((0, 0))
        self.update_status(f"Pasted {len(pasted)} chars (replaced content)")

    def action_clear_editor(self) -> None:
        self.get_editor().text = ""
        self.update_status("Cleared")


def main():
    """Run the Textual Excel Formula Editor."""
    app = TextualExcelEditor()
    app.run()


if __name__ == "__main__":
    main()

# End of file #
