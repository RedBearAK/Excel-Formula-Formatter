#!/usr/bin/env python3
"""
Terminal-based Excel Formula Editor with keyboard shortcuts and clickable UI.
File: excel_terminal_editor.py

Requires: pip install textual
"""

import sys
from pathlib import Path

# Add package to path if needed
package_parent = Path(__file__).parent
sys.path.insert(0, str(package_parent))

from textual.app import App, ComposeResult
from textual.containers import Container, Horizontal
from textual.widgets import Button, TextArea, Static, Footer, Header
from textual.binding import Binding

try:
    from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter
except ImportError:
    print("Excel Formula Formatter package not found. Please install or check PYTHONPATH.")
    sys.exit(1)


class ExcelEditor(App):
    """Excel Formula Terminal Editor with fold/unfold functionality."""
    
    CSS = """
    Screen {
        background: $surface;
    }
    
    .title {
        dock: top;
        height: 3;
        background: $primary;
        color: $text;
        content-align: center middle;
        text-style: bold;
    }
    
    .instructions {
        dock: top;
        height: 2;
        background: $boost;
        color: $text;
        content-align: center middle;
        text-style: italic;
    }
    
    .editor-container {
        height: 1fr;
        border: solid $primary;
        margin: 1;
    }
    
    .button-row {
        dock: bottom;
        height: 4;
        background: $surface;
        padding: 1;
    }
    
    Button {
        margin: 0 1;
        min-width: 12;
        height: 1;
    }
    
    TextArea {
        height: 1fr;
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
        Binding("ctrl+f", "fold_formula", "Fold", show=True),
        Binding("ctrl+u", "unfold_formula", "Unfold", show=True), 
        Binding("ctrl+t", "toggle_formula", "Toggle", show=True),
        Binding("ctrl+c", "copy_text", "Copy", show=True),
        Binding("ctrl+v", "paste_text", "Paste", show=True),
        Binding("ctrl+q", "quit", "Quit", show=True),
    ]
    
    def __init__(self):
        super().__init__()
        self.formatter = ModularExcelFormatter.create_javascript_formatter()
        self.current_file = None
        self.status_message = "Ready"
    
    def compose(self) -> ComposeResult:
        """Compose the UI layout."""
        yield Header(show_clock=True)
        
        yield Static("📊 Excel Formula Editor", classes="title")
        yield Static("Paste your Excel formula here or use Ctrl+V", classes="instructions")
        
        with Container(classes="editor-container"):
            yield TextArea(
                text="",
                id="editor"
            )
        
        with Horizontal(classes="button-row"):
            yield Button("📋 Paste", id="paste", variant="default")
            yield Button("📑 Copy", id="copy", variant="default")
            yield Button("🔄 Toggle", id="toggle", variant="success")
            yield Button("🔽 Fold", id="fold", variant="primary")
            yield Button("🔼 Unfold", id="unfold", variant="primary")
            yield Button("❌ Clear", id="clear", variant="warning")
            yield Button("🚪 Quit", id="quit", variant="error")
        
        yield Static(self.status_message, classes="status", id="status")
        yield Footer()
    
    def on_button_pressed(self, event: Button.Pressed) -> None:
        """Handle button clicks."""
        button_id = event.button.id
        
        if button_id == "fold":
            self.action_fold_formula()
        elif button_id == "unfold":
            self.action_unfold_formula()
        elif button_id == "toggle":
            self.action_toggle_formula()
        elif button_id == "copy":
            self.action_copy_text()
        elif button_id == "paste":
            self.action_paste_text()
        elif button_id == "clear":
            self.action_clear_editor()
        elif button_id == "quit":
            self.action_quit()
    
    def action_fold_formula(self) -> None:
        """Fold Excel formula to JavaScript-like syntax."""
        editor = self.query_one("#editor", TextArea)
        text = editor.text
        
        if not text.strip():
            self.update_status("❌ No text to fold")
            return
        
        try:
            folded = self.formatter.fold_formula(text)
            editor.text = folded
            editor.move_cursor((0, 0))
            self.update_status("✅ Formula folded")
        except Exception as e:
            self.update_status(f"❌ Fold error: {str(e)}")
    
    def action_unfold_formula(self) -> None:
        """Unfold JavaScript-like syntax to Excel formula."""
        editor = self.query_one("#editor", TextArea)
        text = editor.text
        
        if not text.strip():
            self.update_status("❌ No text to unfold")
            return
        
        try:
            unfolded = self.formatter.unfold_formula(text)
            editor.text = unfolded
            editor.move_cursor((0, 0))
            self.update_status("✅ Formula unfolded")
        except Exception as e:
            self.update_status(f"❌ Unfold error: {str(e)}")
    
    def action_toggle_formula(self) -> None:
        """Auto-detect and toggle between folded/unfolded."""
        editor = self.query_one("#editor", TextArea)
        text = editor.text
        
        if not text.strip():
            self.update_status("❌ No text to toggle")
            return
        
        try:
            lines = text.strip().split('\n')
            
            # Auto-detect format
            if len(lines) == 1 and (text.startswith('=') or text.startswith('{=')):
                # Single line Excel formula - fold it
                result = self.formatter.fold_formula(text)
                action = "folded"
            elif len(lines) > 1 and ('//' in text or 'Excel Formula' in text):
                # Multi-line with comments - unfold it
                result = self.formatter.unfold_formula(text)
                action = "unfolded"
            else:
                # Try to fold as single line
                single_line = ' '.join(line.strip() for line in lines if line.strip())
                if single_line.startswith('=') or single_line.startswith('{='):
                    result = self.formatter.fold_formula(single_line)
                    action = "folded"
                else:
                    self.update_status("❓ Format not recognized")
                    return
            
            editor.text = result
            editor.move_cursor((0, 0))
            self.update_status(f"✅ Formula {action}")
            
        except Exception as e:
            self.update_status(f"❌ Toggle error: {str(e)}")
    
    def action_copy_text(self) -> None:
        """Copy text to clipboard."""
        editor = self.query_one("#editor", TextArea)
        text = editor.text
        
        if not text.strip():
            self.update_status("❌ No text to copy")
            return
        
        try:
            # Try to copy to clipboard (platform dependent)
            import subprocess
            import platform
            
            system = platform.system()
            if system == "Darwin":  # macOS
                subprocess.run(["pbcopy"], input=text, text=True)
            elif system == "Linux":
                subprocess.run(["xclip", "-selection", "clipboard"], input=text, text=True)
            elif system == "Windows":
                subprocess.run(["clip"], input=text, text=True)
            
            self.update_status("✅ Text copied to clipboard")
        except Exception:
            self.update_status("❌ Copy failed - clipboard not available")
    
    def action_paste_text(self) -> None:
        """Paste text from clipboard."""
        try:
            import subprocess
            import platform
            
            system = platform.system()
            if system == "Darwin":  # macOS
                result = subprocess.run(["pbpaste"], capture_output=True, text=True)
            elif system == "Linux":
                result = subprocess.run(["xclip", "-selection", "clipboard", "-o"], capture_output=True, text=True)
            elif system == "Windows":
                result = subprocess.run(["powershell", "Get-Clipboard"], capture_output=True, text=True)
            
            if result.returncode == 0 and result.stdout.strip():
                editor = self.query_one("#editor", TextArea)
                pasted_text = result.stdout.strip()
                
                # Replace all content with pasted text
                editor.text = pasted_text
                
                # Move cursor to beginning
                editor.move_cursor((0, 0))
                
                self.update_status("✅ Text pasted from clipboard")
            else:
                self.update_status("❌ Nothing to paste")
                
        except Exception:
            self.update_status("❌ Paste failed - clipboard not available")
    
    def action_clear_editor(self) -> None:
        """Clear the editor."""
        editor = self.query_one("#editor", TextArea)
        editor.text = ""
        editor.move_cursor((0, 0))
        self.update_status("✅ Editor cleared")
    
    def update_status(self, message: str) -> None:
        """Update status bar message."""
        status = self.query_one("#status", Static)
        status.update(message)


def main():
    """Run the Excel Formula Terminal Editor."""
    print("🚀 Starting Excel Formula Terminal Editor...")
    print("📋 Keyboard shortcuts:")
    print("   Ctrl+T = Toggle (auto fold/unfold)")
    print("   Ctrl+F = Fold formula")
    print("   Ctrl+U = Unfold formula") 
    print("   Ctrl+C = Copy text")
    print("   Ctrl+V = Paste text")
    print("   Ctrl+Q = Quit")
    print()
    print("📝 Note: This editor provides real-time editing but no syntax highlighting.")
    print("   For beautiful syntax highlighting, use: python simple_excel_editor.py")
    print()
    
    app = ExcelEditor()
    app.run()


if __name__ == "__main__":
    main()

# End of file #
