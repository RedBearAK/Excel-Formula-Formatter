#!/usr/bin/env python3
"""
Session-aware system clipboard access shared by all editors.
File: excel_formula_formatter/clipboard_helpers.py

Extracted from the (previously triplicated) editor clipboard code.
On Linux, tool selection is prioritized by session type: Wayland
sessions try wl-clipboard first, X11 sessions try xclip/xsel first,
with the other family as fallback for XWayland setups. Session
detection follows the approach borrowed from the Toshy keymapper
project: XDG_SESSION_TYPE, then display environment variables, then a
process scan as last resort.
"""

import os
import platform
import subprocess


def detect_session_type() -> str:
    """Detect whether we're in an X11 or Wayland session on Linux."""
    session_type = os.environ.get("XDG_SESSION_TYPE", "").lower()
    if session_type in ["wayland", "x11"]:
        return session_type

    wayland_display = os.environ.get('WAYLAND_DISPLAY', '')
    if wayland_display and wayland_display.startswith('wayland'):
        return 'wayland'

    if os.environ.get('DISPLAY'):
        return 'x11'

    # Last resort: process check
    try:
        process_check = subprocess.run(['ps', 'ax'], capture_output=True, text=True)
        if process_check.returncode == 0:
            listing = process_check.stdout.lower()
            if 'wayland' in listing:
                return 'wayland'
            if 'xorg' in listing:
                return 'x11'
    except Exception:
        pass

    return 'unknown'


def _linux_paste_tools() -> list:
    """Paste tool candidates ordered for the current session type.

    Each entry is (command_template, uses_selections). Templates with a
    {} placeholder get each X11 selection substituted in turn.
    """
    wayland_first = [
        (["wl-paste"], False),
        (["xclip", "-selection", "{}", "-o"], True),
        (["xsel", "--{}", "--output"], True),
    ]
    x11_first = [
        (["xclip", "-selection", "{}", "-o"], True),
        (["xsel", "--{}", "--output"], True),
        (["wl-paste"], False),
    ]
    return wayland_first if detect_session_type() == 'wayland' else x11_first


def _linux_copy_tools() -> list:
    """Copy tool candidates ordered for the current session type."""
    wayland_first = [
        ["wl-copy"],
        ["xclip", "-selection", "clipboard"],
        ["xsel", "--clipboard", "--input"],
    ]
    x11_first = [
        ["xclip", "-selection", "clipboard"],
        ["xsel", "--clipboard", "--input"],
        ["wl-copy"],
    ]
    return wayland_first if detect_session_type() == 'wayland' else x11_first


def get_clipboard_text():
    """Read the system clipboard. Returns the text, or None on failure."""
    system = platform.system()

    if system == "Darwin":
        try:
            result = subprocess.run(["pbpaste"], capture_output=True, text=True)
            if result.returncode == 0:
                return result.stdout.strip()
        except Exception:
            pass
        return None

    # Non-macOS: try Python clipboard libraries before subprocess tools
    try:
        import pyperclip
        content = pyperclip.paste()
        if content and content.strip():
            return content.strip()
    except ImportError:
        pass
    except Exception:
        pass

    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()  # Hide the window
        content = root.clipboard_get()
        root.destroy()
        if content and content.strip():
            return content.strip()
    except Exception:
        pass

    if system == "Windows":
        try:
            result = subprocess.run(
                ["powershell", "Get-Clipboard"], capture_output=True, text=True
            )
            if result.returncode == 0:
                return result.stdout.strip()
        except Exception:
            pass
        return None

    if system == "Linux":
        # Content might live in PRIMARY rather than CLIPBOARD
        selections = ["clipboard", "primary", "secondary"]

        for tool_template, uses_selections in _linux_paste_tools():
            candidate_commands = []
            if uses_selections:
                for selection in selections:
                    candidate_commands.append(
                        [part.format(selection) if "{}" in part else part
                         for part in tool_template]
                    )
            else:
                candidate_commands.append(tool_template)

            for tool_cmd in candidate_commands:
                try:
                    result = subprocess.run(
                        tool_cmd, capture_output=True, text=True, timeout=2
                    )
                    if result.returncode == 0 and result.stdout.strip():
                        return result.stdout.strip()
                except (subprocess.TimeoutExpired, FileNotFoundError,
                        subprocess.CalledProcessError):
                    continue
                except Exception:
                    continue

    return None


def set_clipboard_text(text: str) -> bool:
    """Write text to the system clipboard. Returns True on success."""
    system = platform.system()

    if system == "Darwin":
        try:
            subprocess.run(["pbcopy"], input=text, text=True, check=True)
            return True
        except Exception:
            return False

    if system == "Windows":
        try:
            subprocess.run(["clip"], input=text, text=True, check=True)
            return True
        except Exception:
            return False

    if system == "Linux":
        for tool_cmd in _linux_copy_tools():
            try:
                subprocess.run(
                    tool_cmd, input=text, text=True, timeout=2, check=True
                )
                return True
            except (subprocess.TimeoutExpired, FileNotFoundError,
                    subprocess.CalledProcessError):
                continue
            except Exception:
                continue

    return False

# End of file #
