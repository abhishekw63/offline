"""
exporter.sheets.warnings_sheet
==============================

Writes the **Warnings** sheet — surfaces every non-fatal issue the
engine hit during processing. The sheet is **only created when there
is at least one warning** — a clean run produces a workbook without
this tab, which is itself a signal to the user.

Column layout (3 columns)::

    1. PO        — source PO (empty for global warnings)
    2. Location  — raw location string (empty for non-location warnings)
    3. Warning   — human-readable message

Warnings come from:

* Unmapped locations (per PO+location, deduped).
* Price mismatches (per item, deduped).
* Unknown GST codes (per code, deduped).
* Missing optional columns (global).
* Rows skipped due to missing data (per PO where feasible).

All warnings use the orange :data:`~._styles.WARN_FILL` header to
distinguish them from ordinary data sheets.
"""

from __future__ import annotations

from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    WARN_FILL, auto_width, data_cell, hdr_cell,
)


_HEADERS = ['PO', 'Location', 'Warning']


def write(wb, result: ProcessingResult) -> None:
    """
    Append the 'Warnings' sheet to ``wb``, but only if there are
    warnings to report. No-op on clean runs.
    """
    if not result.warnings:
        return

    ws = wb.create_sheet('Warnings')

    for col_idx, header in enumerate(_HEADERS, start=1):
        hdr_cell(ws, 1, col_idx, header, fill=WARN_FILL)

    for r, (po, loc, msg) in enumerate(result.warnings, start=2):
        data_cell(ws, r, 1, po)
        data_cell(ws, r, 2, loc)
        data_cell(ws, r, 3, msg)

    auto_width(ws)
