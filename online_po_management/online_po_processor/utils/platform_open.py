"""
utils.platform_open
===================

Cross-platform "open this file with its default app" helper.

Wraps the three OS-specific ways of launching a file:

* Windows: ``os.startfile`` (uses the registered default handler)
* macOS:   ``open`` shell command
* Linux:   ``xdg-open`` shell command

Any failure is surfaced as a Tk messagebox rather than a traceback —
this is called from the GUI's "Open Last Output" button where a silent
failure or a stack trace would both be bad UX.
"""

from __future__ import annotations
import os
import platform
import subprocess
from pathlib import Path
from tkinter import messagebox


def open_file(file_path: Path) -> None:
    """
    Launch ``file_path`` with the OS default application.

    Args:
        file_path: File to open. ``Path`` or ``str`` both work (we
                   ``str()`` it before handing to the OS).

    The function never raises — errors become a user-facing dialog.
    """
    try:
        system = platform.system()
        if system == 'Windows':
            # os.startfile is Windows-only; ignore the linter warning on
            # non-Windows platforms. This branch is never hit there.
            os.startfile(str(file_path))  # type: ignore[attr-defined]
        elif system == 'Darwin':
            subprocess.Popen(['open', str(file_path)])
        else:
            subprocess.Popen(['xdg-open', str(file_path)])
    except Exception as e:  # noqa: BLE001 — user-facing fallback
        messagebox.showerror(
            "Open File Error",
            f"Could not open file:\n{e}",
        )
