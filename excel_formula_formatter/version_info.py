#!/usr/bin/env python3
"""
Version information for the Excel Formula Formatter package.
File: excel_formula_formatter/version_info.py

Single source of truth for the package version. Uses timestamp versioning
(YYYYMMDD) with a zero-padded build number for multiple releases on the
same date. Bump VERSION_BUILD for same-day changes; reset it to 1 when
VERSION_DATE changes.
"""


VERSION_DATE = "20260717"
VERSION_BUILD = 13


def get_version_string() -> str:
    """Return the display version, e.g. '20260717 (build 002)'."""
    return f"{VERSION_DATE} (build {VERSION_BUILD:03d})"


def get_version_dotted() -> str:
    """Return the machine-friendly version, e.g. '20260717.002'."""
    return f"{VERSION_DATE}.{VERSION_BUILD:03d}"

# End of file #
