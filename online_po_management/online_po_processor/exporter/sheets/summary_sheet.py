"""
exporter.sheets.summary_sheet
=============================

Writes the **Summary** sheet — a per-PO grouped view for human
verification before the SO is imported.

Column layout (8 columns)::

    1. PO
    2. Location (Raw)      — what the marketplace sent us
    3. Location (Mapped)   — canonical key matched to from Ship-To registry
    4. Cust No
    5. Ship-to
    6. Items               — count of lines on this PO
    7. Total Qty           — sum of quantities across lines
    8. Status              — 'OK' (green) or 'UNMAPPED' (red)

Visual aids
-----------
* **Pale yellow fill on both location cells** when the raw and mapped
  names differ (case-insensitive). Means we used a fuzzy match — worth
  a quick eyeball to confirm we matched to the right Ship-To.
* Status pill: green for OK, red for UNMAPPED.
* TOTAL row at the bottom for Items + Qty.
* Info sub-row: marketplace, margin %, filename, generation timestamp.
* Legend row appears **only when** there's at least one yellow
  highlight — keeps clean runs free of noise.
"""

from __future__ import annotations
from datetime import datetime
from typing import Dict

from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    BOLD_DATA_FONT, INFO_ITALIC_FONT, LEGEND_ITALIC_FONT,
    LOC_MISMATCH_FILL, STATUS_BAD_FILL, STATUS_BAD_FONT,
    STATUS_OK_FILL, STATUS_OK_FONT,
    auto_width, data_cell, hdr_cell,
)


_HEADERS = [
    'PO', 'Location (Raw)', 'Location (Mapped)',
    'Cust No', 'Ship-to', 'Items', 'Total Qty', 'Status',
]

# 1-based column indices for cells we style specially.
_COL_PO = 1
_COL_RAW_LOC = 2
_COL_MAPPED_LOC = 3
_COL_ITEMS = 6
_COL_QTY = 7
_COL_STATUS = 8


def write(wb, result: ProcessingResult) -> None:
    """
    Append the 'Summary' sheet to ``wb``.
    """
    ws = wb.create_sheet('Summary')

    # ── Header row ──────────────────────────────────────────────────────
    for col_idx, header in enumerate(_HEADERS, start=1):
        hdr_cell(ws, 1, col_idx, header)

    # ── Group by PO ─────────────────────────────────────────────────────
    # Every row of a given PO shares location/cust_no/ship_to (guaranteed
    # by the engine — one PO = one delivery location). So we capture those
    # from the first SORow seen for each PO, then accumulate Items + Qty.
    po_groups: Dict[str, dict] = {}
    for so_row in result.rows:
        if so_row.po_number not in po_groups:
            po_groups[so_row.po_number] = {
                'location': so_row.location,
                'mapped_location': so_row.mapped_location,
                'cust_no': so_row.cust_no,
                'ship_to': so_row.ship_to,
                'mapped': so_row.mapped,
                'items': 0,
                'qty': 0,
            }
        po_groups[so_row.po_number]['items'] += 1
        po_groups[so_row.po_number]['qty'] += so_row.qty

    # ── Data rows ───────────────────────────────────────────────────────
    r = 2
    for po, info in po_groups.items():
        status = 'OK' if info['mapped'] else 'UNMAPPED'

        data_cell(ws, r, _COL_PO, po)
        data_cell(ws, r, _COL_RAW_LOC, info['location'])
        data_cell(ws, r, _COL_MAPPED_LOC, info['mapped_location'])
        data_cell(ws, r, 4, info['cust_no'])
        data_cell(ws, r, 5, info['ship_to'])
        data_cell(ws, r, _COL_ITEMS, info['items'])
        data_cell(ws, r, _COL_QTY, info['qty'])
        data_cell(ws, r, _COL_STATUS, status)

        # Yellow highlight when raw ≠ mapped (case-insensitive).
        # Indicates a fuzzy match — worth a human glance.
        raw_norm = (info['location'] or '').strip().lower()
        mapped_norm = (info['mapped_location'] or '').strip().lower()
        if info['mapped'] and raw_norm and mapped_norm and raw_norm != mapped_norm:
            ws.cell(row=r, column=_COL_RAW_LOC).fill = LOC_MISMATCH_FILL
            ws.cell(row=r, column=_COL_MAPPED_LOC).fill = LOC_MISMATCH_FILL

        # Status pill
        status_cell = ws.cell(row=r, column=_COL_STATUS)
        if status == 'OK':
            status_cell.fill = STATUS_OK_FILL
            status_cell.font = STATUS_OK_FONT
        else:
            status_cell.fill = STATUS_BAD_FILL
            status_cell.font = STATUS_BAD_FONT

        r += 1

    # ── Totals row ──────────────────────────────────────────────────────
    total_items = sum(g['items'] for g in po_groups.values())
    total_qty = sum(g['qty'] for g in po_groups.values())

    data_cell(ws, r, _COL_PO, 'TOTAL')
    ws.cell(row=r, column=_COL_PO).font = BOLD_DATA_FONT
    data_cell(ws, r, _COL_ITEMS, total_items)
    ws.cell(row=r, column=_COL_ITEMS).font = BOLD_DATA_FONT
    data_cell(ws, r, _COL_QTY, total_qty)
    ws.cell(row=r, column=_COL_QTY).font = BOLD_DATA_FONT

    # ── Info sub-row ────────────────────────────────────────────────────
    r += 2
    margin_str = f"{int(result.margin_pct * 100)}%"
    info_text = (f"Marketplace: {result.marketplace}  |  "
                 f"Margin: {margin_str}  |  "
                 f"File: {result.input_file}  |  "
                 f"Generated: {datetime.now().strftime('%d-%m-%Y %H:%M')}")
    ws.cell(row=r, column=1, value=info_text).font = INFO_ITALIC_FONT

    # ── Legend row (conditional) ────────────────────────────────────────
    # Only show when at least one yellow highlight exists — otherwise the
    # legend is noise in a clean run.
    any_loc_mismatch = any(
        (g['mapped']
         and (g['location'] or '').strip().lower()
             != (g['mapped_location'] or '').strip().lower()
         and g['location'] and g['mapped_location'])
        for g in po_groups.values()
    )
    if any_loc_mismatch:
        r += 1
        ws.cell(
            row=r, column=1,
            value=("🟨 Yellow = raw and mapped location differ "
                   "(fuzzy match) — please verify."),
        ).font = LEGEND_ITALIC_FONT

    auto_width(ws)
