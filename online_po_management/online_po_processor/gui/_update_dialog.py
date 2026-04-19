"""
gui._update_dialog
==================

Modal dialog used by the "Update Bundled Files" button.

Prompts the user to choose which bundled file (or both) they want to
replace. Returns the user's choice as a string so the caller can pipe
it into the file-copy flow.
"""

from __future__ import annotations
import tkinter as tk
from pathlib import Path
from typing import Literal, Optional


# Literal set for the choice return value — mirrors the keys used by
# ``OnlinePOApp._update_bundled_files`` when deciding what to copy.
Choice = Literal['master', 'mapping', 'both']


class UpdateDialog:
    """
    Modal Toplevel that asks which bundled file to update.

    Usage::

        dialog = UpdateDialog(parent_root, folder=Path('Calculation Data'))
        choice = dialog.show()   # blocks until user clicks a button
        if choice is None:
            return  # user cancelled
        # choice is 'master' | 'mapping' | 'both'
    """

    def __init__(self, parent: tk.Misc, folder: Path) -> None:
        """
        Args:
            parent: Parent Tk window (used for modality + positioning).
            folder: Bundled-data folder path — displayed to the user so
                    they know where the file will be copied TO.
        """
        self._parent = parent
        self._folder = folder
        self._choice: Optional[Choice] = None

    def show(self) -> Optional[Choice]:
        """
        Display the dialog and block until the user picks or cancels.

        Returns:
            The user's choice, or ``None`` if they cancelled (clicked
            Cancel or closed the window).
        """
        win = tk.Toplevel(self._parent)
        win.title("Update Bundled Files")
        win.geometry("380x200")
        win.resizable(False, False)
        win.transient(self._parent)
        win.grab_set()

        tk.Label(
            win, text="Which file do you want to update?",
            font=("Arial", 11, "bold"),
        ).pack(pady=(15, 8))

        tk.Label(
            win,
            text=f"Files will be copied into:\n{self._folder}",
            font=("Arial", 9), fg='gray', justify='center',
        ).pack(pady=(0, 10))

        # ── Horizontal pair of primary options ─────────────────────────
        btn_row = tk.Frame(win)
        btn_row.pack(pady=4)
        tk.Button(
            btn_row, text="Items Master", width=14,
            command=lambda: self._pick('master', win),
        ).pack(side='left', padx=4)
        tk.Button(
            btn_row, text="Ship-To Mapping", width=14,
            command=lambda: self._pick('mapping', win),
        ).pack(side='left', padx=4)

        # ── Full-width alternatives underneath ─────────────────────────
        tk.Button(
            win, text="Both", width=30,
            command=lambda: self._pick('both', win),
        ).pack(pady=4)
        tk.Button(
            win, text="Cancel", width=30,
            command=win.destroy,
        ).pack(pady=4)

        # Block caller until one of the buttons destroys the window.
        self._parent.wait_window(win)
        return self._choice

    def _pick(self, kind: Choice, win: tk.Toplevel) -> None:
        """Internal — store the choice and close the dialog."""
        self._choice = kind
        win.destroy()
