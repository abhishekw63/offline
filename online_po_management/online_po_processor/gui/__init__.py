"""
gui — Tkinter UI.

Public re-exports
-----------------
Only ``OnlinePOApp`` is exposed; the internal widget builders
(``_file_row``, ``_update_dialog``) are implementation details.
"""

from online_po_processor.gui.app_window import OnlinePOApp

__all__ = ['OnlinePOApp']
