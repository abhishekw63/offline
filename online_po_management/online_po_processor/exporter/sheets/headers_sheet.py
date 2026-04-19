"""
exporter.sheets.headers_sheet
=============================

Writes the **Headers (SO)** sheet — the SO header rows that the ERP
imports to create the document shells. One row per unique PO number.

Column layout (18 columns)::

    1.  Document Type             = 'Order'
    2.  No.                       = the PO number
    3.  Sell-to Customer No.      = from mapping
    4.  Ship-to Code              = from mapping
    5.  Posting Date              = today
    6.  Order Date                = today
    7.  Document Date             = today
    8.  Invoice From Date         = today
    9.  Invoice To Date           = today
    10. External Document No.     = the PO number (same as col 2)
    11. Location Code             = 'PICK'  (always)
    12. Dimension Set ID          = ''
    13. Supply Type               = 'B2B'   (always)
    14. Voucher Narration         = ''
    15. Brand Code (Dimension)    = ''
    16. Channel Code (Dimension)  = ''
    17. Catagory (Dimension)      = ''      (sic — matches ERP's spelling)
    18. Geography Code (Dimension)= ''

The dimension columns (15-18) are blank by design — the ERP fills them
from defaults.
"""

from __future__ import annotations
from datetime import datetime

from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    auto_width, data_cell, hdr_cell,
)


# Column headers in display order (1-based positions match docstring).
_HEADERS = [
    'Document Type', 'No.', 'Sell-to Customer No.', 'Ship-to Code',
    'Posting Date', 'Order Date', 'Document Date',
    'Invoice From Date', 'Invoice To Date',
    'External Document No.', 'Location Code', 'Dimension Set ID',
    'Supply Type', 'Voucher Narration',
    'Brand Code (Dimension)', 'Channel Code (Dimension)',
    'Catagory (Dimension)', 'Geography Code (Dimension)',
]


def write(wb, result: ProcessingResult) -> None:
    """
    Append the 'Headers (SO)' sheet to ``wb``.

    Args:
        wb:     openpyxl Workbook to write into.
        result: ProcessingResult — only ``result.rows`` is consulted (to
                derive the unique PO list).
    """
    ws = wb.create_sheet('Headers (SO)')

    for col_idx, header in enumerate(_HEADERS, start=1):
        hdr_cell(ws, 1, col_idx, header)

    today_str = datetime.now().strftime("%d-%m-%Y")

    # Collect unique POs preserving the order they were processed in.
    # We use a set for O(1) membership check and a parallel list for order.
    seen: set = set()
    unique_po_rows = []
    for so_row in result.rows:
        if so_row.po_number not in seen:
            seen.add(so_row.po_number)
            unique_po_rows.append(so_row)

    # One header row per unique PO. We pull cust_no / ship_to from the
    # FIRST SORow we saw for that PO (all rows of a single PO share the
    # same delivery location, so the values are identical).
    for r, so_row in enumerate(unique_po_rows, start=2):
        data_cell(ws, r, 1, 'Order')              # Document Type
        data_cell(ws, r, 2, so_row.po_number)     # No.
        data_cell(ws, r, 3, so_row.cust_no)       # Sell-to Customer No.
        data_cell(ws, r, 4, so_row.ship_to)       # Ship-to Code
        data_cell(ws, r, 5, today_str)            # Posting Date
        data_cell(ws, r, 6, today_str)            # Order Date
        data_cell(ws, r, 7, today_str)            # Document Date
        data_cell(ws, r, 8, today_str)            # Invoice From Date
        data_cell(ws, r, 9, today_str)            # Invoice To Date
        data_cell(ws, r, 10, so_row.po_number)    # External Document No.
        data_cell(ws, r, 11, 'PICK')              # Location Code
        data_cell(ws, r, 12, '')                  # Dimension Set ID
        data_cell(ws, r, 13, 'B2B')               # Supply Type
        # Columns 14–18 left blank (Voucher Narration + 4 dimension cols).

    auto_width(ws)
