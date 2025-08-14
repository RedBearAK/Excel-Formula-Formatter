#!/usr/bin/env python3
"""
Simple terminal Excel formula editor using Rich.
File: simple_excel_editor.py

Requires: pip install rich
"""

import sys
from pathlib import Path

# Add package to path if needed  
package_parent = Path(__file__).parent
sys.path.insert(0, str(package_parent))

import subprocess
import platform
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Prompt
from rich.syntax import Syntax
from rich.text import Text
from rich.table import Table
from rich.live import Live
from rich.layout import Layout

try:
    from excel_formula_formatter import ExcelFormulaFormatter
except ImportError:
    print("Excel Formula Formatter package not found. Please install or check PYTHONPATH.")
    sys.exit(1)


class SimpleExcelEditor:
    def __init__(self):
        self.console = Console()
        self.formatter = ExcelFormulaFormatter()
        self.text = ""
        
    def show_header(self):
        """Display header with instructions."""
        header = Table.grid(padding=1)
        header.add_column(style="bold blue")
        header.add_column(style="bold green")
        
        header.add_row("📊 Excel Formula Editor", "Keyboard Shortcuts:")
        header.add_row("", "F = Fold formula")
        header.add_row("", "U = Unfold formula") 
        header.add_row("", "T = Toggle (auto)")
        header.add_row("", "P = Paste from clipboard")
        header.add_row("", "C = Copy to clipboard")
        header.add_row("", "E = Edit text manually")
        header.add_row("", "Q = Quit")
        
        return Panel(header, title="Excel Formula Terminal Editor", border_style="blue")
    
    def display_text(self):
        """Display current text with syntax highlighting."""
        if not self.text.strip():
            content = Text("(No content - press P to paste or E to edit)", style="dim")
            title = "Formula Content"
        else:
            # Detect if it's folded (multiline) or unfolded (single line)
            if '\n' in self.text and '//' in self.text:
                # Folded - use JavaScript highlighting
                content = Syntax(self.text, "javascript", theme="monokai", line_numbers=True)
                title = "Formula Content (JavaScript highlighting)"
            elif '\n' in self.text and '#' in self.text:
                # Python syntax
                content = Syntax(self.text, "python", theme="monokai", line_numbers=True)
                title = "Formula Content (Python highlighting)"
            elif '\n' in self.text and '/*' in self.text:
                # CSS syntax
                content = Syntax(self.text, "css", theme="monokai", line_numbers=True)
                title = "Formula Content (CSS highlighting)"
            else:
                # Unfolded - use plain text  
                content = Syntax(self.text, "text", theme="monokai", line_numbers=False)
                title = "Formula Content (Plain text)"
        
        return Panel(content, title=title, border_style="green")
    
    def get_clipboard_text(self):
        """Get text from clipboard."""
        try:
            system = platform.system()
            if system == "Darwin":  # macOS
                result = subprocess.run(["pbpaste"], capture_output=True, text=True)
            elif system == "Linux":
                result = subprocess.run(["xclip", "-selection", "clipboard", "-o"], capture_output=True, text=True)
            elif system == "Windows":
                result = subprocess.run(["powershell", "Get-Clipboard"], capture_output=True, text=True)
            
            if result.returncode == 0:
                return result.stdout.strip()
        except Exception:
            pass
        return None
    
    def set_clipboard_text(self, text):
        """Set clipboard text."""
        try:
            system = platform.system()
            if system == "Darwin":  # macOS
                subprocess.run(["pbcopy"], input=text, text=True)
            elif system == "Linux":
                subprocess.run(["xclip", "-selection", "clipboard"], input=text, text=True)
            elif system == "Windows":
                subprocess.run(["clip"], input=text, text=True)
            return True
        except Exception:
            return False
    
    def auto_toggle(self):
        """Auto-detect and toggle format."""
        if not self.text.strip():
            return "❌ No text to toggle"
        
        try:
            lines = self.text.strip().split('\n')
            
            if len(lines) == 1 and (self.text.startswith('=') or self.text.startswith('{=')):
                # Single line Excel formula - fold it
                self.text = self.formatter.fold_formula(self.text)
                return "✅ Formula folded"
            elif len(lines) > 1 and ('//' in self.text or 'Excel Formula' in self.text):
                # Multi-line with comments - unfold it
                self.text = self.formatter.unfold_formula(self.text)
                return "✅ Formula unfolded"
            else:
                # Try to fold as single line
                single_line = ' '.join(line.strip() for line in lines if line.strip())
                if single_line.startswith('=') or single_line.startswith('{='):
                    self.text = self.formatter.fold_formula(single_line)
                    return "✅ Formula folded"
                else:
                    return "❓ Format not recognized"
        except Exception as e:
            return f"❌ Error: {str(e)}"
    
    def run(self):
        """Main editor loop."""
        self.console.clear()
        
        while True:
            # Display UI
            self.console.clear()
            self.console.print(self.show_header())
            self.console.print()
            self.console.print(self.display_text())
            self.console.print()
            
            # Get user input
            try:
                choice = Prompt.ask(
                    "[bold yellow]Choose action[/bold yellow]",
                    choices=["f", "u", "t", "p", "c", "e", "q"],
                    default="t"
                ).lower()
                
                if choice == 'q':
                    self.console.print("👋 Goodbye!", style="bold blue")
                    break
                    
                elif choice == 'f':
                    if self.text.strip():
                        try:
                            self.text = self.formatter.fold_formula(self.text)
                            self.console.print("✅ Formula folded", style="green")
                        except Exception as e:
                            self.console.print(f"❌ Fold error: {e}", style="red")
                    else:
                        self.console.print("❌ No text to fold", style="red")
                
                elif choice == 'u':
                    if self.text.strip():
                        try:
                            self.text = self.formatter.unfold_formula(self.text)
                            self.console.print("✅ Formula unfolded", style="green")
                        except Exception as e:
                            self.console.print(f"❌ Unfold error: {e}", style="red")
                    else:
                        self.console.print("❌ No text to unfold", style="red")
                
                elif choice == 't':
                    result = self.auto_toggle()
                    if result.startswith('✅'):
                        self.console.print(result, style="green")
                    elif result.startswith('❓'):
                        self.console.print(result, style="yellow")
                    else:
                        self.console.print(result, style="red")
                
                elif choice == 'p':
                    clipboard_text = self.get_clipboard_text()
                    if clipboard_text:
                        self.text = clipboard_text
                        self.console.print("✅ Text pasted from clipboard", style="green")
                    else:
                        self.console.print("❌ No text in clipboard", style="red")
                
                elif choice == 'c':
                    if self.text.strip():
                        if self.set_clipboard_text(self.text):
                            self.console.print("✅ Text copied to clipboard", style="green")
                        else:
                            self.console.print("❌ Failed to copy to clipboard", style="red")
                    else:
                        self.console.print("❌ No text to copy", style="red")
                
                elif choice == 'e':
                    self.console.print("\n[bold yellow]Enter your Excel formula (press Enter twice to finish):[/bold yellow]")
                    lines = []
                    while True:
                        line = input()
                        if line == "" and lines and lines[-1] == "":
                            break
                        lines.append(line)
                    
                    # Remove the last empty line
                    if lines and lines[-1] == "":
                        lines.pop()
                    
                    self.text = '\n'.join(lines)
                    self.console.print("✅ Text updated", style="green")
                
                # Brief pause to show messages
                if choice != 'q':
                    import time
                    time.sleep(1)
                    
            except KeyboardInterrupt:
                self.console.print("\n👋 Goodbye!", style="bold blue")
                break
            except EOFError:
                break


def main():
    """Run the simple Excel formula editor."""
    editor = SimpleExcelEditor()
    editor.run()


if __name__ == "__main__":
    main()

# End of file #
