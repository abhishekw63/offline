"""
app
===

Application bootstrap — the ``main()`` entry point called by the
top-level ``main.py`` launcher.

Handles two things before the GUI comes up:

#. Sets up basic logging config (to stdout).
#. Runs the expiry check — if the build has passed its
   :data:`~online_po_processor.config.constants.EXPIRY_DATE`, shows an
   error dialog and exits before any UI is shown.

Everything else is delegated to
:class:`~online_po_processor.gui.app_window.OnlinePOApp`.
"""

from __future__ import annotations
import logging
import sys
import tkinter as tk
from datetime import datetime
from tkinter import messagebox

from online_po_processor.config.constants import EXPIRY_DATE
from online_po_processor.gui.app_window import OnlinePOApp


# Days-to-expiry threshold below which we nag the user on startup
# (one last push to get the next build out).
_EXPIRY_WARN_DAYS = 7


def check_expiry() -> None:
    """
    Enforce the hard expiry cutoff.

    * If the build date has passed: show an error dialog and exit with
      status 0 (graceful shutdown, not a crash).
    * If within ``_EXPIRY_WARN_DAYS`` of expiry: show a warning dialog
      but continue into the app.

    Uses a withdrawn (invisible) Tk root so the dialogs have a valid
    parent without briefly flashing an empty window.
    """
    expiry = datetime.strptime(EXPIRY_DATE, "%d-%m-%Y").date()
    today = datetime.now().date()

    # ── Expired ─────────────────────────────────────────────────────────
    if today > expiry:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Application Expired",
            f"This application expired on {EXPIRY_DATE}.\n\n"
            f"Please contact the administrator for an updated version.",
        )
        root.destroy()
        sys.exit(0)

    # ── Nearing expiry (warn but continue) ──────────────────────────────
    days_remaining = (expiry - today).days
    if days_remaining <= _EXPIRY_WARN_DAYS:
        root = tk.Tk()
        root.withdraw()
        messagebox.showwarning(
            "Expiration Warning",
            f"⚠️ This application will expire in {days_remaining} day(s).\n\n"
            f"Expiry Date: {EXPIRY_DATE}\n"
            f"Please contact the administrator for renewal.",
        )
        root.destroy()


def _configure_logging() -> None:
    """
    Basic logging config: level=INFO, timestamp + level + message.

    We use ``logging`` throughout the package for diagnostics that
    shouldn't clutter the GUI log panel (e.g. fuzzy-match notes, file
    load counts). The GUI's ``_log()`` is separate — that one is the
    user-facing surface.
    """
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s | %(levelname)s | %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )


def main() -> None:
    """Entry point: configure logging, gate on expiry, launch the GUI."""
    _configure_logging()
    check_expiry()
    OnlinePOApp().run()


if __name__ == '__main__':
    # Allow `python -m online_po_processor.app` for quick launch during
    # development. The canonical launcher is still the top-level main.py.
    main()
