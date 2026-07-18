#!/usr/bin/env python3
"""
Enhanced simple terminal Excel formula editor with four syntax modes.
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

try:
    from excel_formula_formatter.version_info import get_version_string
    from excel_formula_formatter.clipboard_helpers import (
        detect_session_type,
        get_clipboard_text,
        set_clipboard_text,
    )
    from excel_formula_formatter.modular_excel_formatter import (
        ModularExcelFormatter, detect_current_mode, safe_mode_switch, auto_format_with_mode
    )
except ImportError:
    print("Excel Formula Formatter package not found. Please install or check PYTHONPATH.")
    sys.exit(1)


class EnhancedFourModeExcelEditor:
    def __init__(self):
        self.console = Console()
        self.text = ""
        self.current_mode = "p"  # Default to Plain Excel mode
        
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
            },
            "c": {
                "name": "Compact Excel",
                "description": "Minimal spacing Excel syntax (for 8K character limit)",
                "formatter": ModularExcelFormatter.create_compact_formatter()
            }
        }
    
    def show_header(self):
        """Display header with instructions and current mode."""
        header = Table.grid(padding=0)
        header.add_column(style="bold blue", min_width=26)
        header.add_column(style="bold green", min_width=26)
        
        current_mode_name = self.mode_info[self.current_mode]["name"]
        has_content = bool(self.text.strip())
        default_hint = "Default: P (paste)" if not has_content else "Default: T (toggle)"
        
        header.add_row("📊 Excel Formula Editor", f"Current Mode: {current_mode_name} ({self.current_mode})")
        header.add_row("Main Commands:", "Mode Commands:")
        header.add_row("T = Toggle (auto)", "M = Change mode")
        header.add_row("F = Fold formula", "L = List all 4 modes")
        header.add_row("U = Unfold formula", "")
        header.add_row("P = Paste from clipboard", "Other:")
        header.add_row("C = Copy to clipboard", "E = Edit text manually")
        header.add_row("Q = Quit", "Tools = Clipboard diagnostic")
        header.add_row("", default_hint)
        
        panel_title = f"Excel Formula Terminal Editor (4 Modes) — v{get_version_string()}"
        return Panel(header, title=panel_title, border_style="blue")

    def show_mode_info(self):
        """Show information about available modes."""
        mode_info_table = Table.grid(padding=0)
        mode_info_table.add_column(style="bold yellow", min_width=22)
        mode_info_table.add_column(style="white", min_width=45)
        
        for mode_code, info in self.mode_info.items():
            current_marker = " ← CURRENT" if mode_code == self.current_mode else ""
            mode_info_table.add_row(f"{mode_code} = {info['name']}{current_marker}", info['description'])
        
        return Panel(mode_info_table, title="Available Syntax Modes (4 Total)", border_style="yellow")
    
    def display_text(self):
        """Display current text with appropriate syntax highlighting."""
        if not self.text.strip():
            content = Text("(No content yet)\n\n💡 Copy your Excel formula to clipboard, then press P to paste", style="dim")
            title = "Ready for Your Excel Formula"
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
            elif self.current_mode == "c":
                if '\n' in self.text and not '//' in self.text:
                    content = Syntax(self.text, "text", theme="monokai", line_numbers=True)
                    title = "Formula Content (Compact Excel - Minimal Spacing)"
                else:
                    content = Syntax(self.text, "text", theme="monokai", line_numbers=False)
                    title = "Formula Content (Plain text, compact)"
            else:  # plain mode 'p'
                if '\n' in self.text and not '//' in self.text:
                    content = Syntax(self.text, "text", theme="monokai", line_numbers=True)
                    title = "Formula Content (Plain Excel - Smart Indented)"
                else:
                    content = Syntax(self.text, "text", theme="monokai", line_numbers=False)
                    title = "Formula Content (Plain text)"
        
        return Panel(content, title=title, border_style="green")
    
    def detect_package_manager(self):
        """Detect common Linux package managers for better help messages."""
        import subprocess
        
        package_managers = [
            ('dnf', 'sudo dnf install'),
            ('apt', 'sudo apt install'), 
            ('pacman', 'sudo pacman -S'),
            ('zypper', 'sudo zypper install'),
            ('emerge', 'sudo emerge'),
        ]
        
        for pm, install_cmd in package_managers:
            try:
                result = subprocess.run(['which', pm], capture_output=True, text=True)
                if result.returncode == 0:
                    return install_cmd
            except Exception:
                continue
        
        return 'sudo <package_manager> install'  # Generic fallback
    
    def detect_session_type(self):
        """Detect if we're in X11 or Wayland session (shared helper)."""
        return detect_session_type()

    def get_clipboard_text(self):
        """Get text from clipboard via the shared session-aware helper."""
        return get_clipboard_text()

    def set_clipboard_text(self, text):
        """Set clipboard text via the shared session-aware helper."""
        return set_clipboard_text(text)

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
                old_mode_name = self.mode_info[self.current_mode]['name']
                new_mode_name = self.mode_info[target_mode]['name']
                
                self.text = safe_mode_switch(self.text, self.current_mode, target_mode, should_refold=True)
                self.current_mode = target_mode
                
                return f"✅ Safely switched from {old_mode_name} to {new_mode_name}"
            except Exception as e:
                return f"❌ Mode switch failed: {str(e)}"
    
    def robust_text_input(self):
        """Robust text input that handles very long pastes better than input()."""
        self.console.print("\n[bold yellow]Enter your Excel formula:[/bold yellow]")
        self.console.print("[dim]• Type or paste your formula[/dim]")
        self.console.print("[dim]• Press Ctrl+D when finished[/dim]")
        self.console.print("[dim]• Cancel with Ctrl+C[/dim]")
        self.console.print("[dim]• Very long formulas work better with clipboard paste (P command)[/dim]")
        
        try:
            content_lines = []
            
            while True:
                try:
                    line = input()
                    content_lines.append(line)
                    
                    # Check length and warn, but don't interrupt the flow
                    if len(line) > 2000 and len(content_lines) == 1:
                        self.console.print(f"\n[yellow]ℹ️  Long line detected ({len(line)} chars)[/yellow]")
                        self.console.print("[yellow]💡 Tip: For very long formulas, clipboard paste (P command) is more reliable[/yellow]")
                        self.console.print("[yellow]💡 Consider using Compact mode (c) to reduce character count[/yellow]")
                        self.console.print("[dim]Continue typing or press Ctrl+D to finish...[/dim]\n")
                    
                except KeyboardInterrupt:
                    return "❌ Cancelled"
                    
        except EOFError:
            # Ctrl+D pressed - process what we have
            if content_lines:
                content = '\n'.join(content_lines).strip()
                if content:
                    self.text = content
                    
                    # Give feedback about what we received
                    if len(content_lines) == 1 and len(content) > 1000:
                        return f"✅ Long formula received ({len(content)} characters)"
                    elif len(content_lines) > 1:
                        return f"✅ Multi-line content received ({len(content_lines)} lines)"
                    else:
                        return "✅ Text updated"
                else:
                    return "ℹ️ No text entered"
            else:
                return "ℹ️ No text entered"
    
    def debug_clipboard_access(self):
        """Debug clipboard access to see exactly what's failing."""
        debug_info = []
        system = platform.system()
        
        # Test each method individually with detailed feedback
        if system == "Darwin":
            debug_info.append("Platform: macOS")
            try:
                result = subprocess.run(["pbpaste"], capture_output=True, text=True)
                if result.returncode == 0:
                    content_len = len(result.stdout.strip())
                    if content_len > 0:
                        debug_info.append(f"✅ pbpaste: Success ({content_len} chars)")
                    else:
                        debug_info.append("❌ pbpaste: Empty clipboard")
                else:
                    debug_info.append(f"❌ pbpaste: Exit code {result.returncode}")
            except Exception as e:
                debug_info.append(f"❌ pbpaste: Error - {e}")
                
        else:
            session_type = self.detect_session_type()
            debug_info.append(f"Platform: {system}")
            debug_info.append(f"Session type: {session_type}")
            
            # Test pyperclip
            try:
                import pyperclip
                content = pyperclip.paste()
                if content and content.strip():
                    debug_info.append(f"✅ pyperclip: Success ({len(content.strip())} chars)")
                else:
                    debug_info.append("❌ pyperclip: Empty result")
            except ImportError:
                debug_info.append("❌ pyperclip: Not installed")
            except Exception as e:
                debug_info.append(f"❌ pyperclip: Error - {e}")
            
            # Test tkinter
            try:
                import tkinter as tk
                root = tk.Tk()
                root.withdraw()
                content = root.clipboard_get()
                root.destroy()
                if content and content.strip():
                    debug_info.append(f"✅ tkinter: Success ({len(content.strip())} chars)")
                else:
                    debug_info.append("❌ tkinter: Empty result")
            except Exception as e:
                debug_info.append(f"❌ tkinter: Error - {e}")
            
            # Test system tools with multiple clipboard selections
            if system == "Linux":
                # Test multiple clipboard selections
                selections = [
                    ("clipboard", "CLIPBOARD (Ctrl+C/V)"),
                    ("primary", "PRIMARY (select/middle-click)"),
                    ("secondary", "SECONDARY")
                ]
                
                # Session-aware tool prioritization
                if session_type == 'wayland':
                    tools = [
                        (["wl-paste"], "wl-paste", False),  # Wayland doesn't use selections
                        (["xclip", "-selection", "{}", "-o"], "xclip", True),
                        (["xsel", "--{}", "--output"], "xsel", True),
                    ]
                else:
                    tools = [
                        (["xclip", "-selection", "{}", "-o"], "xclip", True),
                        (["xsel", "--{}", "--output"], "xsel", True),
                        (["wl-paste"], "wl-paste", False),
                    ]
                
                for tool_cmd_template, tool_name, uses_selections in tools:
                    if uses_selections:
                        # Test each selection
                        for selection, selection_desc in selections:
                            tool_cmd = []
                            for part in tool_cmd_template:
                                if "{}" in part:
                                    tool_cmd.append(part.format(selection))
                                else:
                                    tool_cmd.append(part)
                            
                            try:
                                result = subprocess.run(tool_cmd, capture_output=True, text=True, timeout=2)
                                if result.returncode == 0 and result.stdout.strip():
                                    content_len = len(result.stdout.strip())
                                    debug_info.append(f"✅ {tool_name} ({selection}): Success ({content_len} chars)")
                                else:
                                    debug_info.append(f"❌ {tool_name} ({selection}): Empty or failed")
                            except FileNotFoundError:
                                debug_info.append(f"❌ {tool_name}: Not installed")
                                break  # Don't test other selections if tool not found
                            except Exception as e:
                                debug_info.append(f"❌ {tool_name} ({selection}): Error - {e}")
                    else:
                        # Tool doesn't use selections (like wl-paste)
                        priority_marker = " (prioritized for Wayland)" if session_type == 'wayland' and tool_name == 'wl-paste' else ""
                        try:
                            result = subprocess.run(tool_cmd_template, capture_output=True, text=True, timeout=2)
                            if result.returncode == 0 and result.stdout.strip():
                                content_len = len(result.stdout.strip())
                                debug_info.append(f"✅ {tool_name}: Success ({content_len} chars){priority_marker}")
                            else:
                                debug_info.append(f"❌ {tool_name}: Empty or failed{priority_marker}")
                        except FileNotFoundError:
                            debug_info.append(f"❌ {tool_name}: Not installed{priority_marker}")
                        except Exception as e:
                            debug_info.append(f"❌ {tool_name}: Error - {e}{priority_marker}")
        
        return debug_info

    def check_clipboard_tools(self):
        """Check which clipboard tools are available (for diagnostics)."""
        available_tools = []
        system = platform.system()
        
        if system == "Darwin":
            # macOS uses native tools directly
            available_tools.append("✅ macOS clipboard (pbcopy/pbpaste) - primary method")
            
        else:
            # Non-macOS: Check Python clipboard libraries first
            try:
                import pyperclip
                available_tools.append("✅ Python pyperclip library - primary method")
            except ImportError:
                available_tools.append("❌ Python pyperclip library (install: pip install pyperclip)")
            
            try:
                import tkinter as tk
                available_tools.append("✅ Python tkinter clipboard (built-in) - secondary method")
            except ImportError:
                available_tools.append("❌ Python tkinter clipboard")
            
            # Check system tools as fallbacks
            if system == "Linux":
                tools_to_check = [
                    ("wl-paste", "Wayland clipboard (wl-clipboard) - fallback"),
                    ("xclip", "X11 clipboard (xclip) - fallback"), 
                    ("xsel", "X11 clipboard (xsel) - fallback"),
                ]
                
                for tool, description in tools_to_check:
                    try:
                        result = subprocess.run(["which", tool], capture_output=True, text=True)
                        if result.returncode == 0:
                            available_tools.append(f"✅ {description}")
                        else:
                            available_tools.append(f"❌ {description}")
                    except Exception:
                        available_tools.append(f"❌ {description}")
            
            elif system == "Windows":
                available_tools.append("✅ Windows clipboard (clip/powershell) - fallback")
        
        return available_tools
    
    def run(self):
        """Main editor loop."""
        self.console.clear()
        
        # Show startup message
        self.console.print("📊 [bold blue]Excel Formula Editor (4 Modes)[/bold blue] - Transform your Excel formulas!")
        self.console.print()
        self.console.print("🚀 [bold yellow]Quick Start:[/bold yellow]")
        self.console.print("   1. Copy your Excel formula to clipboard")
        self.console.print("   2. Press [bold]P[/bold] (or just Enter) to paste")
        self.console.print("   3. Formula will auto-format in your chosen mode")
        self.console.print()
        self.console.print("✨ [bold yellow]Four Syntax Modes:[/bold yellow]")
        self.console.print("   [bold]j[/bold] = JavaScript    - For syntax highlighting with quotes")
        self.console.print("   [bold]a[/bold] = Annotated     - Excel with helpful comments")
        self.console.print("   [bold]p[/bold] = Plain         - Pure Excel with smart indenting")
        self.console.print("   [bold]c[/bold] = Compact       - Minimal spacing (8K character limit)")
        self.console.print()
        self.console.print("[dim]💡 Tips:[/dim]")
        self.console.print("[dim]• P (paste) works best for long formulas[/dim]")
        self.console.print("[dim]• E (edit) lets you paste directly in terminal (always works)[/dim]")
        self.console.print("[dim]• Use Compact mode (c) for very long formulas near 8K limit[/dim]")
        if platform.system() == "Linux":
            session_type = self.detect_session_type()
            install_cmd = self.detect_package_manager()
            if session_type == 'wayland':
                self.console.print(f"[dim]• If P fails on Wayland: {install_cmd} wl-clipboard[/dim]")
            elif session_type == 'x11':
                self.console.print(f"[dim]• If P fails on X11: {install_cmd} xclip[/dim]")
            else:
                self.console.print(f"[dim]• If P fails: {install_cmd} wl-clipboard xclip[/dim]")
        elif platform.system() not in ["Darwin", "Windows"]:
            self.console.print("[dim]• If P fails: pip install pyperclip[/dim]")
        self.console.print()
        
        # Main interaction loop
        while True:
            # Display UI
            self.console.clear()
            self.console.print(self.show_header())
            self.console.print()
            self.console.print(self.display_text())
            self.console.print()
            
            # Dynamic default: "p" if no content, "t" if content exists
            has_content = bool(self.text.strip())
            default_choice = "t" if has_content else "p"
            
            # Get user input
            try:
                choice = Prompt.ask(
                    "[bold yellow]Choose action[/bold yellow]",
                    choices=["t", "f", "u", "p", "c", "e", "m", "l", "tools", "q"],
                    default=default_choice
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
                    self.console.print()
                    self.console.print("[dim]💡 Mode descriptions:[/dim]")
                    self.console.print("[dim]• JavaScript: Best for syntax highlighting in editors[/dim]")
                    self.console.print("[dim]• Annotated: Excel syntax with helpful section comments[/dim]")
                    self.console.print("[dim]• Plain: Clean Excel with smart indenting, no comments[/dim]")
                    self.console.print("[dim]• Compact: Minimal spacing to avoid 8K character limit[/dim]")
                    input("\nPress Enter to continue...")
                    continue
                
                elif choice == 'e':
                    result = self.robust_text_input()
                    if result.startswith('✅'):
                        self.console.print(result, style="green")
                    elif result.startswith('ℹ️'):
                        self.console.print(result, style="blue")
                    else:
                        self.console.print(result, style="red")
                    
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
                    self.console.print("🔍 [dim]Attempting to read clipboard...[/dim]")
                    clipboard_text = self.get_clipboard_text()
                    if clipboard_text:
                        self.text = clipboard_text
                        char_count = len(clipboard_text)
                        if char_count > 6000:
                            self.console.print(f"✅ Very large formula pasted ({char_count} chars)", style="green")
                            self.console.print("💡 [yellow]Consider using Compact mode (c) to reduce character count[/yellow]", style="yellow")
                        elif char_count > 1000:
                            self.console.print(f"✅ Large formula pasted from clipboard ({char_count} chars)", style="green")
                        else:
                            self.console.print("✅ Text pasted from clipboard", style="green")
                    else:
                        self.console.print("❌ No text in clipboard (or clipboard access failed)", style="red")
                        
                        # Show detailed debugging info
                        self.console.print("\n[bold yellow]🔧 Clipboard Debug Info:[/bold yellow]")
                        debug_info = self.debug_clipboard_access()
                        for info_line in debug_info:
                            self.console.print(f"  {info_line}")
                        
                        # Don't auto-clear this error - wait for user input
                        input("\nPress Enter to continue...")
                        continue
                
                elif choice == 'c':
                    if self.text.strip():
                        if self.set_clipboard_text(self.text):
                            char_count = len(self.text)
                            if char_count > 1000:
                                self.console.print(f"✅ Large text copied to clipboard ({char_count} chars)", style="green")
                            else:
                                self.console.print("✅ Text copied to clipboard", style="green")
                        else:
                            self.console.print("❌ Failed to copy to clipboard", style="red")
                    else:
                        self.console.print("❌ No text to copy", style="red")
                
                elif choice == 'tools':
                    self.console.print("\n[bold yellow]📋 Clipboard Tools Diagnostic[/bold yellow]")
                    self.console.print()
                    available = self.check_clipboard_tools()
                    for tool_status in available:
                        self.console.print(f"  {tool_status}")
                    
                    self.console.print()
                    self.console.print("[dim]💡 Recommended approach:[/dim]")
                    
                    if platform.system() == "Darwin":
                        self.console.print("[dim]• macOS: Native pbcopy/pbpaste (should work out of box)[/dim]")
                    elif platform.system() == "Linux":
                        session_type = self.detect_session_type()
                        install_cmd = self.detect_package_manager()
                        self.console.print(f"[dim]• Detected session: {session_type}[/dim]")
                        
                        if session_type == 'wayland':
                            self.console.print(f"[dim]• Best for Wayland: {install_cmd} wl-clipboard[/dim]")
                            self.console.print("[dim]• Alternative: pip install pyperclip[/dim]")
                        elif session_type == 'x11':
                            self.console.print(f"[dim]• Best for X11: {install_cmd} xclip xsel[/dim]")
                            self.console.print("[dim]• Alternative: pip install pyperclip[/dim]")
                        else:
                            self.console.print("[dim]• Best: pip install pyperclip (cross-platform Python library)[/dim]")
                            self.console.print(f"[dim]• Or install both: {install_cmd} wl-clipboard xclip[/dim]")
                    else:
                        self.console.print("[dim]• Best: pip install pyperclip (cross-platform Python library)[/dim]")
                    
                    self.console.print("[dim]• Universal fallback: Use E (edit) and paste directly in terminal[/dim]")
                    
                    input("\nPress Enter to continue...")
                    continue
                
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
    """Run the enhanced four-mode Excel formula editor."""
    editor = EnhancedFourModeExcelEditor()
    editor.run()


if __name__ == "__main__":
    main()

# End of file #
