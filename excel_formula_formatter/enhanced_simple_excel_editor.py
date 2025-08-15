#!/usr/bin/env python3
"""
Enhanced simple terminal Excel formula editor with language mode switching.
File: enhanced_simple_excel_editor.py

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
from rich.prompt import Prompt, Confirm
from rich.syntax import Syntax
from rich.text import Text
from rich.table import Table
from rich.live import Live
from rich.layout import Layout

try:
    from excel_formula_formatter import ExcelFormulaFormatter
    from excel_formula_formatter.modular_excel_formatter import ModularExcelFormatter
except ImportError:
    print("Excel Formula Formatter package not found. Please install or check PYTHONPATH.")
    sys.exit(1)


class EnhancedSimpleExcelEditor:
    def __init__(self):
        self.console = Console()
        self.text = ""
        self.current_mode = "javascript"  # Default mode
        
        # Available formatters - JavaScript and Plain only
        self.formatters = {
            "javascript": ModularExcelFormatter.create_javascript_formatter(),
            "plain": ModularExcelFormatter.create_plain_formatter()
        }
        

    def show_header(self):
        """Display header with instructions and current mode."""
        header = Table.grid(padding=0)  # padding= adds blank lines between rows, keep at 0
        header.add_column(style="bold blue", min_width=26)  # Set minimum width for first column
        header.add_column(style="bold green", min_width=26)  # Set minimum width for second column
        
        header.add_row("📊 Excel Formula Editor", f"Current Mode: {self.current_mode.title()}")
        header.add_row("Main Commands:", "Mode Commands:")
        header.add_row("T = Toggle (auto)", "M = Change mode")
        header.add_row("F = Fold formula", "L = List modes")
        header.add_row("U = Unfold formula", "")
        header.add_row("P = Paste from clipboard", "Other:")
        header.add_row("C = Copy to clipboard", "E = Edit text manually")
        header.add_row("Q = Quit", "")
        
        return Panel(header, title="Excel Formula Terminal Editor", border_style="blue")

    def show_mode_info(self):
        """Show information about available modes."""
        mode_info = Table.grid(padding=0)  # padding= adds blank lines between rows, keep at 0
        mode_info.add_column(style="bold yellow")
        mode_info.add_column(style="white")
        
        mode_info.add_row("javascript", "JavaScript-like syntax (functions, quoted cells \"A1\")")
        mode_info.add_row("plain", "Pure Excel syntax with smart indenting only")
        
        return Panel(mode_info, title="Available Syntax Modes", border_style="yellow")
    
    def display_text(self):
        """Display current text with appropriate syntax highlighting."""
        if not self.text.strip():
            content = Text("(No content - press P to paste or E to edit)", style="dim")
            title = "Formula Content"
        else:
            # Detect syntax and apply highlighting
            if self.current_mode == "javascript":
                if '\n' in self.text and '//' in self.text:
                    content = Syntax(self.text, "javascript", theme="monokai", line_numbers=True)
                    title = "Formula Content (JavaScript highlighting)"
                else:
                    content = Syntax(self.text, "text", theme="monokai", line_numbers=False)
                    title = "Formula Content (Plain text)"
            else:  # plain mode
                if '\n' in self.text:
                    # Multi-line plain Excel - check for comment indicators
                    if '//' in self.text and 'plain Excel syntax' in self.text:
                        content = Syntax(self.text, "text", theme="monokai", line_numbers=True)
                        title = "Formula Content (Pure Excel - Smart Indented)"
                    else:
                        content = Syntax(self.text, "text", theme="monokai", line_numbers=True)
                        title = "Formula Content (Multi-line text)"
                else:
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
        """Auto-detect and toggle format using current mode formatter."""
        if not self.text.strip():
            return "❌ No text to toggle"
        
        try:
            formatter = self.formatters[self.current_mode]
            lines = self.text.strip().split('\n')
            
            if len(lines) == 1 and (self.text.startswith('=') or self.text.startswith('{=')):
                # Single line Excel formula - fold it
                self.text = formatter.fold_formula(self.text)
                return f"✅ Formula folded ({self.current_mode} syntax)"
            elif len(lines) > 1 and (any(indicator in self.text for indicator in ['//', 'Excel Formula'])):
                # Multi-line with formatting - unfold it
                self.text = formatter.unfold_formula(self.text)
                return "✅ Formula unfolded to Excel"
            else:
                # Try to fold as single line
                single_line = ' '.join(line.strip() for line in lines if line.strip())
                if single_line.startswith('=') or single_line.startswith('{='):
                    self.text = formatter.fold_formula(single_line)
                    return f"✅ Formula folded ({self.current_mode} syntax)"
                else:
                    return "❓ Format not recognized"
        except Exception as e:
            return f"❌ Error: {str(e)}"
    
    def change_mode(self):
        """Change the current syntax mode."""
        self.console.print(self.show_mode_info())
        
        choice = Prompt.ask(
            "\n[bold yellow]Choose syntax mode[/bold yellow]",
            choices=list(self.formatters.keys()),
            default=self.current_mode
        )
        
        old_mode = self.current_mode
        self.current_mode = choice
        return f"✅ Mode changed from {old_mode} to {choice}"
    
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
                    choices=["t", "f", "u", "p", "c", "e", "m", "l", "q"],
                    default="t"
                ).lower()
                
                if choice == 'q':
                    self.console.print("👋 Goodbye!", style="bold blue")
                    break
                    
                elif choice == 'm':
                    result = self.change_mode()
                    self.console.print(result, style="green")
                
                elif choice == 'l':
                    self.console.print(self.show_mode_info())
                    input("\nPress Enter to continue...")
                    continue
                    
                elif choice == 'f':
                    if self.text.strip():
                        try:
                            formatter = self.formatters[self.current_mode]
                            self.text = formatter.fold_formula(self.text)
                            self.console.print(f"✅ Formula folded ({self.current_mode} syntax)", style="green")
                        except Exception as e:
                            self.console.print(f"❌ Fold error: {e}", style="red")
                    else:
                        self.console.print("❌ No text to fold", style="red")
                
                elif choice == 'u':
                    if self.text.strip():
                        try:
                            formatter = self.formatters[self.current_mode]
                            self.text = formatter.unfold_formula(self.text)
                            self.console.print("✅ Formula unfolded to Excel", style="green")
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
                    time.sleep(1.5)
                    
            except KeyboardInterrupt:
                self.console.print("\n👋 Goodbye!", style="bold blue")
                break
            except EOFError:
                break


def main():
    """Run the enhanced Excel formula editor."""
    editor = EnhancedSimpleExcelEditor()
    editor.run()


if __name__ == "__main__":
    main()

# End of file #
