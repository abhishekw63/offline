"""
exporter.so_exporter
====================

Orchestrates writing the output workbook.

The class itself is intentionally thin: it decides WHERE to write the
file, creates the workbook, and delegates the HOW of each sheet to the
per-sheet modules in :mod:`online_po_processor.exporter.sheets`. This
keeps each sheet's logic independent and easy to change without
touching the others.

Output location (v1.3.4)
------------------------
The output file is written next to the **input punch file** in an
``output/`` subfolder that's auto-created::

    D:\\PO\\Myntra\\April\\Myntra_Punch_17-04-2026.xlsx      ← input
    D:\\PO\\Myntra\\April\\output\\myntra_so_19-04-2026_*.xlsx ← output

Each month's batch lives next to its inputs instead of piling up in a
single global folder.

If ``result.input_file_path`` is empty (shouldn't happen in normal use)
the exporter falls back to the script directory's ``output_online/``
folder — defensive only, won't trigger during a normal run.

Filename format
---------------
``<marketplace_slug>_so_<DD-MM-YYYY>_<HHMMSS>.xlsx``

The timestamp means repeat runs never clobber prior outputs.
"""

from __future__ import annotations
import logging
from datetime import datetime
from pathlib import Path
from tkinter import messagebox
from typing import Optional

from openpyxl import Workbook

from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter.sheets import (
    headers_sheet, lines_sheet, raw_data_sheet,
    summary_sheet, validation_sheet, warnings_sheet,
)


class SOExporter:
    """
    Write a workbook from a :class:`ProcessingResult`.

    Stateless — instances hold nothing, so reusing the same instance
    across runs is safe (and cheap).
    """

    def export(self, result: ProcessingResult) -> Optional[Path]:
        """
        Render ``result`` to an .xlsx file on disk.

        Args:
            result: Fully-populated result from
                    :meth:`MarketplaceEngine.process`.

        Returns:
            ``Path`` to the saved file on success, ``None`` when there
            were no rows to write (a user-facing warning dialog is
            shown in that case).
        """
        if not result.rows:
            # No rows == nothing to import. Better to tell the user
            # than to silently produce an empty workbook.
            messagebox.showwarning(
                "No Data",
                "No valid rows found.\nNothing to export.",
            )
            return None

        file_path = self._resolve_output_path(result)

        wb = Workbook()
        # Workbook() auto-creates an empty 'Sheet' — remove it before we
        # add our own named sheets so the final book has exactly the
        # tabs we want.
        wb.remove(wb.active)

        # Sheet order matters for the user's reading flow:
        #   Headers (SO)  → ERP import (top tab = what you act on)
        #   Lines (SO)    → ERP import
        #   Summary       → human verification, per-PO
        #   Validation    → human verification, per-item price check
        #   Warnings      → only present if there are issues to fix
        #   Raw Data      → audit trail at the bottom
        headers_sheet.write(wb, result)
        lines_sheet.write(wb, result)
        summary_sheet.write(wb, result)
        validation_sheet.write(wb, result)
        warnings_sheet.write(wb, result)
        raw_data_sheet.write(wb, result)

        wb.save(str(file_path))
        logging.info("Output saved: %s", file_path)
        return file_path

    # ── Internal helpers ──────────────────────────────────────────────

    @staticmethod
    def _resolve_output_path(result: ProcessingResult) -> Path:
        """
        Compute the full output path and ensure the parent folder exists.

        Prefers ``<punch-dir>/output/`` so each batch's output lives
        next to the input it came from.
        """
        if result.input_file_path:
            output_folder = Path(result.input_file_path).parent / 'output'
        else:
            # Defensive fallback — the engine always populates
            # input_file_path, so hitting this branch indicates either
            # a programming error or tests that bypass the engine.
            output_folder = Path('output_online')

        output_folder.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.now().strftime('%d-%m-%Y_%H%M%S')
        marketplace_slug = result.marketplace.lower().replace(' ', '_')
        return output_folder / f'{marketplace_slug}_so_{timestamp}.xlsx'
