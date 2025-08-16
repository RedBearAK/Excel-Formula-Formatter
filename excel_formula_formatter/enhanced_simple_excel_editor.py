#!/usr/bin/env python3
"""
Enhanced simple terminal Excel formula editor with three-mode system and safe switching.
File: excel_formula_formatter/enhanced_simple_excel_editor.py

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
    from excel_formula_formatter.modular_excel_formatter import (
        ModularExcelFormatter, detect_current_mode, safe_mode_switch, auto_format_with_mode
    )
except ImportError:
    print("Excel Formula Formatter package not found. Please install or check PYTHONPATH.")
    sys.exit(1)


class EnhancedThreeModeExcelEditor:
    def __init__(self):
        self.console = Console()
        self.text = ""
        self.current_mode = "j"  # Default to JavaScript mode
        
        # Available formatters with single letter codes
        self.mode_info = {
            "j": {
                "name": "JavaScript",
                "description": "JavaScript-like syntax with quoted cells \"A1\"",
                "formatter": ModularExcelFormatter.create_javascript_formatter()
            },
            "a": {
                "name": "Annotated Excel", 
                "description": "Excel syntax with helpful comments and spacing",
                "formatter": ModularExcelFormatter.create_annotated_formatter()
            },
            "p": {
                "name": "Plain Excel",
                "description": "Pure Excel syntax with smart indenting only",
                "formatter": ModularExcelFormatter.create_plain_formatter()
            }
        }
        

    def show_header(self):
        """Display header with instructions and current mode."""
        header = Table.grid(padding=0)
        header.add_column(style="bold blue", min_width=26)
        header.add_column(style="bold green", min_width=26)
        
        current_mode_name = self.mode_info[self.current_mode]["name"]
        header.add_row("📊 Excel Formula Editor", f"Current Mode: {current_mode_name} ({self.current_mode})")
        header.add_row("Main Commands:", "Mode Commands:")
        header.add_row("T = Toggle (auto)", "M = Change mode")
        header.add_row("F = Fold formula", "L = List modes")
        header.add_row("U = Unfold formula", "")
        header.add_row("P = Paste from clipboard", "Other:")
        header.add_row("C = Copy to clipboard", "E = Edit text manually")
        header.add_row("Q = Quit", "")
        
        return Panel(header, title="Excel Formula Terminal Editor (Three Modes)", border_style="blue")

    def show_mode_info(self):
        """Show information about available modes."""
        mode_info_table = Table.grid(padding=0)
        mode_info_table.add_column(style="bold yellow", min_width=20)
        mode_info_table.add_column(style="white", min_width=40)
        
        for mode_code, info in self.mode_info.items():
            current_marker = " ← CURRENT" if mode_code == self.current_mode else ""
            mode_info_table.add_row(f"{mode_code} = {info['name']}{current_marker}", info['description'])
        
        return Panel(mode_info_table, title="Available Syntax Modes", border_style="yellow")
    
    def display_text(self):
        """Display current text with appropriate syntax highlighting."""
        if not self.text.strip():
            content = Text("(No content - press P to paste or E to edit)", style="dim")
            title = "Formula Content"
        else:
            # Detect syntax and apply highlighting based on current mode and content
            if self.current_mode == "j":
                if '\n' in self.text and '//' in self.text and 'JavaScript syntax' in self.text:
                    content = Syntax(self.text, "javascript", theme="monokai", line_numbers=True)
                    title = "Formula Content (JavaScript highlighting)"
                else:
                    content = Syntax(self.text, "text", theme="monokai", line_numbers=False)
                    title = "Formula Content (Plain text)"
            elif self.current_mode == "a":
                if '\n' in self.text and '//' in self.text and 'annotated Excel syntax' in self.text:
                    content = Syntax(self.text, "text", theme="monokai", line_numbers=True)
                    title = "Formula Content (Annotated Excel - Smart Indented)"
                else:
                    content = Syntax(self.text, "text", theme="monokai", line_numbers=False)
                    title = "Formula Content (Plain text)"
            else:  # plain mode 'p'
                if '\n' in self.text and not '//' in self.text:
                    content = Syntax(self.text, "text", theme="monokai", line_numbers=True)
                    title = "Formula Content (Plain Excel - Smart Indented)"
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
            result = auto_format_with_mode(self.text, self.current_mode)
            if result != self.text:
                self.text = result
                # Determine what happened
                lines = result.strip().split('\n')
                if len(lines) == 1 and (result.startswith('=') or result.startswith('{=')):
                    return f"✅ Formula unfolded to Excel"
                else:
                    return f"✅ Formula folded ({self.mode_info[self.current_mode]['name']} syntax)"
            else:
                return "❓ Format not recognized or no change needed"
        except Exception as e:
            return f"❌ Error: {str(e)}"
    
    def change_mode(self):
        """Change the current syntax mode with automatic safe switching."""
        if not self.text.strip():
            # No text - just change mode setting
            self.console.print(self.show_mode_info())
            
            choice = Prompt.ask(
                "\n[bold yellow]Choose syntax mode[/bold yellow]",
                choices=list(self.mode_info.keys()),
                default=self.current_mode
            )
            
            old_mode = self.current_mode
            self.current_mode = choice
            return f"✅ Mode changed from {self.mode_info[old_mode]['name']} to {self.mode_info[choice]['name']}"
        
        else:
            # Has text - do safe mode switching
            self.console.print(self.show_mode_info())
            
            target_mode = Prompt.ask(
                "\n[bold yellow]Switch to which mode?[/bold yellow]",
                choices=list(self.mode_info.keys()),
                default=self.current_mode
            )
            
            if target_mode == self.current_mode:
                return "ℹ️ Already in that mode"
            
            try:
                # Detect current mode from text if possible
                detected_mode = detect_current_mode(self.text)
                source_mode = detected_mode if detected_mode != 'unknown' else self.current_mode
                
                # Perform safe mode switch with refolding
                old_mode_name = self.mode_info[self.current_mode]['name']
                new_mode_name = self.mode_info[target_mode]['name']
                
                self.text = safe_mode_switch(self.text, source_mode, target_mode, should_refold=True)
                self.current_mode = target_mode
                
                return f"✅ Safely switched from {old_mode_name} to {new_mode_name}"
            except Exception as e:
                return f"❌ Mode switch failed: {str(e)}"
    
    def safe_mode_switch_with_refold(self):
        """Safely switch modes while preserving folded state."""
        if not self.text.strip():
            return "❌ No text to switch modes"
        
        self.console.print(self.show_mode_info())
        
        target_mode = Prompt.ask(
            "\n[bold yellow]Switch to which mode?[/bold yellow]",
            choices=list(self.mode_info.keys()),
            default=self.current_mode
        )
        
        if target_mode == self.current_mode:
            return "ℹ️ Already in that mode"
        
        try:
            # Detect current mode from text if possible
            detected_mode = detect_current_mode(self.text)
            source_mode = detected_mode if detected_mode != 'unknown' else self.current_mode
            
            # Perform safe mode switch with refolding
            old_mode_name = self.mode_info[self.current_mode]['name']
            new_mode_name = self.mode_info[target_mode]['name']
            
            self.text = safe_mode_switch(self.text, source_mode, target_mode, should_refold=True)
            self.current_mode = target_mode
            
            return f"✅ Safely switched from {old_mode_name} to {new_mode_name}"
        except Exception as e:
            return f"❌ Safe switch failed: {str(e)}"
    
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
                    if result.startswith('✅'):
                        self.console.print(result, style="green")
                    elif result.startswith('ℹ️'):
                        self.console.print(result, style="blue")
                    else:
                        self.console.print(result, style="red")
                
                elif choice == 'l':
                    self.console.print(self.show_mode_info())
                    input("\nPress Enter to continue...")
                    continue
                    
                elif choice == 'f':
                    if self.text.strip():
                        try:
                            formatter = self.mode_info[self.current_mode]["formatter"]
                            self.text = formatter.fold_formula(self.text)
                            mode_name = self.mode_info[self.current_mode]["name"]
                            self.console.print(f"✅ Formula folded ({mode_name} syntax)", style="green")
                        except Exception as e:
                            self.console.print(f"❌ Fold error: {e}", style="red")
                    else:
                        self.console.print("❌ No text to fold", style="red")
                
                elif choice == 'u':
                    if self.text.strip():
                        try:
                            formatter = self.mode_info[self.current_mode]["formatter"]
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
    """Run the enhanced three-mode Excel formula editor."""
    editor = EnhancedThreeModeExcelEditor()
    editor.run()


if __name__ == "__main__":
    main()

# End of file #
