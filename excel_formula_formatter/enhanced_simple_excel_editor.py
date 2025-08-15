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


class PlainExcelFormatter:
    """Plain Excel formatter that only does smart indenting without syntax translation."""
    
    def __init__(self, indent_size: int = 4):
        self.indent_size = indent_size
    
    def fold_formula(self, formula: str) -> str:
        """Smart indent Excel formula without changing syntax."""
        if not formula or not formula.strip():
            return ""
            
        # Handle array formulas and regular formulas
        clean_formula = formula.strip()
        is_array_formula = False
        
        if clean_formula.startswith('{=') and clean_formula.endswith('}'):
            is_array_formula = True
            clean_formula = clean_formula[2:-1]
        elif clean_formula.startswith('='):
            clean_formula = clean_formula[1:]
            
        # Parse into tokens (similar to main formatter but preserve Excel syntax)
        tokens = self._parse_excel_tokens_pure(clean_formula)
        lines = self._format_tokens_pure_excel(tokens, is_array_formula)
        
        return '\n'.join(lines)
    
    def unfold_formula(self, formatted_text: str) -> str:
        """Convert smart indented formula back to Excel."""
        if not formatted_text or not formatted_text.strip():
            return ""
        
        lines = formatted_text.strip().split('\n')
        
        # Check for array formula marker (if any)
        is_array_formula = any('Array' in line for line in lines[:2])
        
        # Collect content lines (skip any header comments if present)
        content_lines = []
        for line in lines:
            stripped = line.strip()
            if stripped and not stripped.startswith('Excel Formula'):
                content_lines.append(stripped)
        
        # Join with spaces and clean up
        single_line = ' '.join(content_lines)
        
        # Normalize spacing around operators and punctuation
        import re
        single_line = re.sub(r'\s*([(),])\s*', r'\1', single_line)
        single_line = re.sub(r'\s+', ' ', single_line)
        
        # Add appropriate prefix
        if is_array_formula:
            return f'{{={single_line}}}'
        else:
            return f'={single_line}'
    
    def _parse_excel_tokens_pure(self, formula: str) -> list:
        """Parse Excel formula preserving pure Excel syntax."""
        tokens = []
        i = 0
        length = len(formula)
        
        while i < length:
            if formula[i].isspace():
                i += 1
                continue
                
            # Handle string literals
            if formula[i] == '"':
                end_quote = formula.find('"', i + 1)
                if end_quote == -1:
                    end_quote = length
                token_text = formula[i:end_quote + 1]
                tokens.append(('string', token_text))
                i = end_quote + 1
                continue
            
            # Handle cell references and ranges (including sheet references)
            if self._is_cell_ref_start(formula, i):
                cell_ref = self._extract_cell_reference(formula, i)
                tokens.append(('cell_ref', cell_ref))
                i += len(cell_ref)
                continue
            
            # Handle operators
            if i < length - 1:
                two_char = formula[i:i+2]
                if two_char in ['<>', '>=', '<=']:
                    tokens.append(('operator', two_char))
                    i += 2
                    continue
            
            # Handle single character punctuation and operators
            if formula[i] in '+-*/=<>(),[]:;!&%^':
                tokens.append(('punctuation', formula[i]))
                i += 1
                continue
            
            # Handle functions and identifiers
            start = i
            while i < length and not formula[i].isspace() and formula[i] not in '+-*/=<>(),[]:;!&%^"':
                i += 1
            
            if start < i:
                token_text = formula[start:i]
                if self._is_excel_function(token_text):
                    tokens.append(('function', token_text))
                elif token_text.replace('.', '').replace('E', '').replace('+', '').replace('-', '').isdigit():
                    tokens.append(('number', token_text))
                else:
                    tokens.append(('identifier', token_text))
                
        return tokens
    
    def _is_cell_ref_start(self, formula: str, pos: int) -> bool:
        """Check if position starts a cell reference."""
        if pos >= len(formula):
            return False
        
        # Look for pattern like A1, $A$1, Sheet1!A1, etc.
        import re
        remainder = formula[pos:]
        pattern = r'^(?:[A-Za-z0-9_\']+!)?[A-Z]+\$?\d+(?::[A-Z]+\$?\d+)?'
        return bool(re.match(pattern, remainder))
    
    def _extract_cell_reference(self, formula: str, pos: int) -> str:
        """Extract a complete cell reference from the position."""
        import re
        remainder = formula[pos:]
        pattern = r'^(?:[A-Za-z0-9_\']+!)?[A-Z]+\$?\d+(?::[A-Z]+\$?\d+)?'
        match = re.match(pattern, remainder)
        return match.group(0) if match else formula[pos]
    
    def _is_excel_function(self, token: str) -> bool:
        """Check if token is an Excel function."""
        excel_functions = {
            'SUM', 'IF', 'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'SUMIF', 'SUMIFS',
            'COUNTIF', 'COUNTIFS', 'AVERAGEIF', 'AVERAGEIFS', 'LEN', 'MID', 'LEFT',
            'RIGHT', 'FIND', 'SEARCH', 'SUBSTITUTE', 'CONCATENATE', 'TEXT', 'VALUE',
            'DATE', 'TODAY', 'NOW', 'YEAR', 'MONTH', 'DAY', 'DATEVALUE', 'AND', 'OR',
            'NOT', 'ISERROR', 'ISBLANK', 'FILTER', 'SORT', 'UNIQUE'
        }
        return token.upper() in excel_functions
    
    def _format_tokens_pure_excel(self, tokens: list, is_array_formula: bool = False) -> list:
        """Format tokens as pure Excel with smart indenting."""
        lines = []
        current_line = ""
        depth = 0
        function_stack = []  # Track function names for special handling
        let_arg_count = 0  # Track argument position in LET functions
        
        # Don't add header comment for pure Excel mode
        
        i = 0
        while i < len(tokens):
            token_type, token_text = tokens[i]
            
            if token_type == 'punctuation' and token_text == '(':
                current_line += token_text
                
                # Check if should be inline
                if self._should_keep_inline_pure(tokens, i):
                    # Keep inline - collect until matching paren
                    function_stack.append('')  # Not tracking for inline
                    i += 1
                    paren_count = 1
                    while i < len(tokens) and paren_count > 0:
                        t_type, t_text = tokens[i]
                        if t_type == 'punctuation' and t_text == '(':
                            paren_count += 1
                        elif t_type == 'punctuation' and t_text == ')':
                            paren_count -= 1
                        current_line += t_text
                        i += 1
                    function_stack.pop()
                    i -= 1  # Back up one
                else:
                    # Multi-line
                    # Check what function we're entering
                    if i > 0 and tokens[i-1][0] == 'function':
                        func_name = tokens[i-1][1].upper()
                        function_stack.append(func_name)
                        if func_name == 'LET':
                            let_arg_count = 0
                    else:
                        function_stack.append('')
                    
                    if current_line.strip():
                        lines.append(" " * (depth * self.indent_size) + current_line.strip())
                    current_line = ""
                    depth += 1
                    
            elif token_type == 'punctuation' and token_text == ')':
                # Look ahead for comma after closing paren
                next_token_is_comma = (i + 1 < len(tokens) and 
                                     tokens[i + 1][0] == 'punctuation' and 
                                     tokens[i + 1][1] == ',')
                
                # Handle any content on current line first
                if current_line.strip():
                    lines.append(" " * (depth * self.indent_size) + current_line.strip())
                
                depth = max(0, depth - 1)
                
                # Add closing paren with or without comma
                if next_token_is_comma:
                    lines.append(" " * (depth * self.indent_size) + token_text + ',')
                    i += 1  # Skip the comma token
                else:
                    lines.append(" " * (depth * self.indent_size) + token_text)
                current_line = ""
                
                # Pop function stack
                if function_stack:
                    function_stack.pop()
                
            elif token_type == 'punctuation' and token_text == ',':
                # Keep comma attached to what comes before it
                current_line += token_text
                
                # Special handling for LET functions
                if function_stack and function_stack[-1] == 'LET':
                    let_arg_count += 1
                    # In LET: keep pairs together (name, value on same line)
                    # Break line only after every second argument (after the value)
                    if let_arg_count % 2 == 0:  # After value (even numbered args)
                        lines.append(" " * (depth * self.indent_size) + current_line.strip())
                        current_line = ""
                    else:  # After variable name (odd numbered args)
                        current_line += " "
                elif depth > 0:  # Multi-line context
                    lines.append(" " * (depth * self.indent_size) + current_line.strip())
                    current_line = ""
                else:
                    current_line += " "
                    
            elif token_type == 'operator':
                current_line += f" {token_text} "
                
            else:
                # All other tokens (functions, cell refs, strings, numbers) - keep as-is
                current_line += token_text
                
            i += 1
        
        # Add remaining content
        if current_line.strip():
            lines.append(" " * (depth * self.indent_size) + current_line.strip())
        
        return lines
    
    def _should_keep_inline_pure(self, tokens: list, paren_index: int) -> bool:
        """Determine if function should be kept inline (same logic as main formatter)."""
        if paren_index == 0 or paren_index >= len(tokens):
            return False
        
        # Look at previous token for function name
        prev_token = tokens[paren_index - 1]
        if prev_token[0] == 'function':
            function_name = prev_token[1].upper()
            simple_functions = {'LEN', 'TRIM', 'UPPER', 'LOWER', 'ABS', 'INT', 'ROUND', 'DATEVALUE', 'VALUE', 'TEXT'}
            
            # Count complexity
            i = paren_index + 1
            paren_depth = 1
            arg_count = 0
            has_nested = False
            
            while i < len(tokens) and paren_depth > 0 and i < paren_index + 10:
                token_type, token_text = tokens[i]
                if token_type == 'punctuation' and token_text == '(':
                    paren_depth += 1
                    if paren_depth > 1:
                        has_nested = True
                elif token_type == 'punctuation' and token_text == ')':
                    paren_depth -= 1
                elif token_type == 'punctuation' and token_text == ',' and paren_depth == 1:
                    arg_count += 1
                i += 1
            
            # Simple functions with 1 arg and no nesting
            if function_name in simple_functions and arg_count <= 1 and not has_nested:
                return True
        
        return False


class EnhancedSimpleExcelEditor:
    def __init__(self):
        self.console = Console()
        self.text = ""
        self.current_mode = "javascript"  # Default mode
        
        # Available formatters
        self.formatters = {
            "javascript": ModularExcelFormatter.create_javascript_formatter(),
            "python": ModularExcelFormatter.create_python_formatter(), 
            "css": ModularExcelFormatter.create_css_formatter(),
            "plain": PlainExcelFormatter()
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
        mode_info.add_row("python", "Python-like syntax (sheet['A1'], pythonic functions)")
        mode_info.add_row("css", "CSS-like syntax (.functions, #cells)")
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
            elif self.current_mode == "python":
                if '\n' in self.text and '#' in self.text:
                    content = Syntax(self.text, "python", theme="monokai", line_numbers=True)
                    title = "Formula Content (Python highlighting)"
                else:
                    content = Syntax(self.text, "text", theme="monokai", line_numbers=False)
                    title = "Formula Content (Plain text)"
            elif self.current_mode == "css":
                if '\n' in self.text and '/*' in self.text:
                    content = Syntax(self.text, "css", theme="monokai", line_numbers=True)
                    title = "Formula Content (CSS highlighting)"
                else:
                    content = Syntax(self.text, "text", theme="monokai", line_numbers=False)
                    title = "Formula Content (Plain text)"
            else:  # plain mode
                if '\n' in self.text:
                    # Multi-line plain Excel - no comments to detect since we removed them
                    content = Syntax(self.text, "text", theme="monokai", line_numbers=True)
                    title = "Formula Content (Pure Excel - Smart Indented)"
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
            elif len(lines) > 1 and (any(indicator in self.text for indicator in ['//', '#', '/*', 'Excel Formula'])):
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
