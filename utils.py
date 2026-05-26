"""Shared utilities: external tool detection and install hints."""

import shutil
from typing import List, Tuple


# ---------------------------------------------------------------------------
# Install hints
# ---------------------------------------------------------------------------

_HINTS = {
    "pandoc": (
        "pandoc\n"
        "  macOS  : brew install pandoc\n"
        "  Ubuntu : apt install pandoc\n"
        "  Windows: winget install pandoc  |  https://pandoc.org/installing.html"
    ),
    "libreoffice": (
        "LibreOffice\n"
        "  macOS  : brew install --cask libreoffice\n"
        "  Ubuntu : apt install libreoffice\n"
        "  Windows: winget install TheDocumentFoundation.LibreOffice  |  https://www.libreoffice.org/"
    ),
    "pyhwp": (
        "pyhwp  (provides hwp5html / hwp5txt)\n"
        "  pip install pyhwp"
    ),
}


# ---------------------------------------------------------------------------
# Tool presence
# ---------------------------------------------------------------------------

def which(cmd: str) -> bool:
    """Return True if cmd is on PATH. Handles LibreOffice's dual binary name."""
    if cmd == "libreoffice":
        return bool(shutil.which("libreoffice") or shutil.which("soffice"))
    return bool(shutil.which(cmd))


def _pandoc() -> bool:
    return which("pandoc")


def _libreoffice() -> bool:
    return which("libreoffice")


def _pyhwp() -> bool:
    return which("hwp5html") or which("hwp5txt")


# ---------------------------------------------------------------------------
# Per-format pre-flight check
# ---------------------------------------------------------------------------

def check_tools(ext: str) -> List[str]:
    """
    Return a list of human-readable error strings for any missing tool
    required to convert the given file extension.  Empty list = all good.
    """
    ext = ext.lower()
    errors: List[str] = []

    if ext == ".docx":
        if not _pandoc():
            errors.append(f"Missing: {_HINTS['pandoc']}")

    elif ext == ".doc":
        if not _libreoffice():
            errors.append(f"Missing: {_HINTS['libreoffice']}")
        if not _pandoc():
            errors.append(f"Missing: {_HINTS['pandoc']}")

    elif ext == ".hwp":
        has_pyhwp = _pyhwp()
        has_libre = _libreoffice() and _pandoc()
        if not has_pyhwp and not has_libre:
            errors.append(
                "No HWP conversion tool found.  Install at least one option:\n\n"
                f"  Option A (recommended): {_HINTS['pyhwp']}\n\n"
                f"  Option B (fallback):    {_HINTS['libreoffice']}\n"
                f"                          {_HINTS['pandoc']}"
            )

    # .hwpx is pure-Python — no external tools needed

    return errors


def print_tool_errors(errors: List[str], fatal: bool = True) -> None:
    import sys

    prefix = "✗ " if fatal else "⚠ "
    for err in errors:
        for line in err.splitlines():
            print(f"  {prefix}{line}", file=sys.stderr)
            prefix = "    "  # indent continuation lines
        prefix = "✗ " if fatal else "⚠ "
