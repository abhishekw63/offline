"""
exporter.sheets.lines_sheet
===========================

Writes the **Lines (SO)** sheet — the SO line rows imported into the
ERP. One row per ordered item.

Column layout (8 columns)::

    1. Document Type   = 'Order'
    2. Document No.    = the PO number (matches Headers (SO) col 2)
    3. Line No.        = 10000, 20000, 30000, ... (resets on new PO)
    4. Type            = 'Item'  (always — no service/charge lines here)
    5. No.             = the resolved Item No
    6. Location Code   = 'PICK'  (always)
    7. Quantity        = order qty
    8. Unit Price      = ''      (left blank — WMS computes downstream)

The 10000-step Line No. is an ERP convention so users can later insert
extra lines between existing ones.
"""

from __future__ import annotations

from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    auto_width, data_cell, hdr_cell,
)


_HEADERS = [
    'Document Type', 'Document No.', 'Line No.', 'Type',
    'No.', 'Location Code', 'Quantity', 'Unit Price',
]

# Step between consecutive Line No. values within a PO.
_LINE_NO_STEP = 10_000


def write(wb, result: ProcessingResult) -> None:
    """
    Append the 'Lines (SO)' sheet to ``wb``.

    Args:
        wb:     openpyxl Workbook to write into.
        result: ProcessingResult — emits one line row per ``result.rows``
                entry, in the engine's processing order.
    """
    ws = wb.create_sheet('Lines (SO)')

    for col_idx, header in enumerate(_HEADERS, start=1):
        hdr_cell(ws, 1, col_idx, header)

    # Track Line No. per PO. We don't pre-group — we rely on the engine
    # emitting a PO's rows contiguously, which is true today (rows are in
    # input-file order, and a PO's lines are always contiguous in punch
    # files).
    current_po = None
    line_no = 0

    for r, so_row in enumerate(result.rows, start=2):
        if so_row.po_number != current_po:
            current_po = so_row.po_number
            line_no = 0

        line_no += _LINE_NO_STEP

        data_cell(ws, r, 1, 'Order')          # Document Type
        data_cell(ws, r, 2, so_row.po_number) # Document No.
        data_cell(ws, r, 3, line_no)          # Line No.
        data_cell(ws, r, 4, 'Item')           # Type
        data_cell(ws, r, 5, so_row.item_no)   # No. (Item No)
        data_cell(ws, r, 6, 'PICK')           # Location Code
        data_cell(ws, r, 7, so_row.qty)       # Quantity
        data_cell(ws, r, 8, '')               # Unit Price — WMS fills it

    auto_width(ws)
