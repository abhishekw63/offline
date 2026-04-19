"""
gui._file_row
=============

Module-level builder for the "picker row" widget used for the three
file inputs in the main window (Items Master, Ship-to Mapping, PO file).

Why module-level instead of a method? Because the row is a pure UI
construct — it doesn't depend on the parent app's state, only on the
widgets + StringVars passed in. Extracting it here keeps
``app_window.py`` focused on app logic.

Layout
------
::

    [label, 16 chars]  [filename, 28 chars, blue]  [Browse]
                       Updated: 19-Apr-2026 18:41     ← optional sub-row

The sub-row is only drawn when ``ts_var`` is provided. It's bound to a
``StringVar`` so refreshes (after auto-load or "Update Bundled Files")
update the visible text without recreating the widget.
"""

from __future__ import annotations
import tkinter as tk
from typing import Callable, Optional


def build_file_row(
    parent: tk.Widget,
    label: str,
    var: tk.StringVar,
    command: Callable[[], None],
    ts_var: Optional[tk.StringVar] = None,
) -> None:
    """
    Build a file-picker row inside ``parent``.

    Args:
        parent:  Parent Tk widget (typically a Frame).
        label:   Left-side label text (e.g. ``"Items Master:"``).
        var:     ``StringVar`` holding the displayed filename text.
        command: Callback for the Browse button.
        ts_var:  Optional ``StringVar`` for the update-timestamp sub-line.
                 When ``None``, no sub-line is drawn (e.g. the PO picker
                 doesn't have an "updated" concept).
    """
    # Container holds both the main row and the optional sub-line so
    # they stay vertically grouped in the parent frame.
    container = tk.Frame(parent)
    container.pack(fill='x', pady=3)

    row = tk.Frame(container)
    row.pack(fill='x')

    tk.Label(
        row, text=label, font=("Arial", 9), width=16, anchor='w',
    ).pack(side='left')
    tk.Label(
        row, textvariable=var, font=("Arial", 9), fg='blue',
        width=28, anchor='w',
    ).pack(side='left', padx=4)
    tk.Button(
        row, text='Browse', width=8, command=command,
    ).pack(side='right')

    if ts_var is not None:
        # Sub-row indented under the filename by giving it a blank
        # label of the same width as the main label. Keeps the
        # timestamp visually "hanging" under its filename.
        sub_row = tk.Frame(container)
        sub_row.pack(fill='x', anchor='w')
        tk.Label(sub_row, text='', width=16).pack(side='left')
        tk.Label(
            sub_row, textvariable=ts_var, font=("Arial", 8),
            fg='#777777', anchor='w',
        ).pack(side='left', padx=4)
