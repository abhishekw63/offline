"""
exporter.sheets.validation_sheet
================================

Writes the **Validation** sheet — per-item price check with clear
PASS/FAIL status for each line.

Column layout (11 columns, or 14 when HSN check is enabled)::

    1.  PO
    2.  Item No                    — resolved from master
    3.  EAN
    4.  Description                — from Items_March (readable product name)
    5.  MRP                        ─┐
    6.  Landing (m%)                ├ GREEN headers — our calculated values
    7.  GST Code                    │
    8.  Our Cost Price             ─┘
    9.  Marketplace <Label>         — fob_col value from the punch
    10. Difference with <Label>     — fob_price − calc_price
    11. Status                      — OK / MISMATCH / NOT_IN_MASTER / NO_PRICE

    — HSN cross-check columns (v1.6.0, conditional) —
    12. HSN (Marketplace)           — hsn_col value from the punch
    13. HSN (Master)                — HSN/SAC Code from Items_March
    14. HSN Check                   — OK / MISMATCH / NOT_IN_MASTER

``<Label>`` is the marketplace's ``compare_label`` from config (e.g.
"Landing Rate" for Myntra, "Cost" for RK).

Column meaning depends on ``compare_basis``
-------------------------------------------
* ``basis='landing'`` (Myntra, Reliance): Marketplace value is compared
  against "Landing (m%)" (= MRP × m%, pre-GST). Diff is clean (no GST
  rounding).
* ``basis='cost'`` (RK, Blink): Marketplace value is compared against
  "Our Cost Price" (= MRP × m% ÷ GST, post-GST). Diff may have tiny
  rounding noise — we treat ≤ 1 rupee as OK.

HSN cross-check (v1.6.0)
------------------------
When the marketplace has ``hsn_col`` set in its config (currently
Reliance only), the engine compares the punch's HSN against the
master's HSN per row. The three trailing columns appear only when at
least one row has a non-empty ``hsn_check_status`` — otherwise this
sheet keeps its familiar 11-column layout.

Visual cues
-----------
* **Mismatch rows** get a pale-pink fill across the entire row, Status
  cell in bold red.
* **OK rows** get a green status pill only (the bulk of a clean batch,
  so we keep row fill neutral to reduce visual fatigue).
* **NOT_IN_MASTER rows** get a pale-orange fill so these are easy to
  spot and fix by adding the item to Items_March.
* **HSN mismatches** get a red status pill (same bold-red font as
  price mismatches) but don't repaint the full row — the price-diff
  styling stays the primary signal, HSN is a secondary audit flag.

The trailing info row records ``basis=... | Margin: m%`` so someone
reviewing the output three months later can tell at a glance what the
numbers mean.
"""

from __future__ import annotations

import pandas as pd

from online_po_processor.data.models import ProcessingResult
from online_po_processor.exporter._styles import (
    CALC_FILL, HEADER_FILL, INFO_ITALIC_FONT, MISMATCH_FILL,
    MISMATCH_TEXT_FONT, NO_MASTER_FILL, NOT_IN_MASTER_TEXT_FONT,
    STATUS_OK_FILL, STATUS_OK_FONT,
    auto_width, data_cell, hdr_cell,
)


# Calculated column indices (1-based). These get a green header instead
# of the default blue to visually separate "our math" from "their data".
_CALC_COL_INDICES = {5, 6, 7, 8}  # MRP, Landing, GST Code, Our Cost Price


def write(wb, result: ProcessingResult) -> None:
    """
    Append the 'Validation' sheet to ``wb``.
    """
    ws = wb.create_sheet('Validation')

    label = result.compare_label or 'Price'
    margin_pct_int = int(result.margin_pct * 100)

    # v1.6.0: decide up front whether HSN columns belong on this run.
    # Only shown when the engine actually populated an hsn_check_status
    # on at least one row (which happens when the marketplace config
    # has ``hsn_col`` set). This keeps the sheet's width consistent
    # with previous versions for marketplaces that don't opt in.
    has_hsn_check = any(
        so_row.hsn_check_status for so_row in result.rows
    )

    headers = [
        'PO', 'Item No', 'EAN', 'Description', 'MRP',
        f'Landing ({margin_pct_int}%)', 'GST Code',
        'Our Cost Price',
        f'Marketplace {label}',
        f'Difference with {label}',
        'Status',
    ]
    if has_hsn_check:
        headers.extend([
            'HSN (Marketplace)', 'HSN (Master)', 'HSN Check',
        ])

    # ── Header row ──────────────────────────────────────────────────────
    for col_idx, header in enumerate(headers, start=1):
        fill = CALC_FILL if col_idx in _CALC_COL_INDICES else HEADER_FILL
        hdr_cell(ws, 1, col_idx, header, fill=fill)

    n_cols = len(headers)
    status_col = 11   # Price-validation status column index (fixed)
    # HSN columns, when present, occupy 12/13/14.
    hsn_punch_col = 12
    hsn_master_col = 13
    hsn_status_col = 14

    # ── Data rows ───────────────────────────────────────────────────────
    r = 2
    mismatches = 0
    hsn_mismatches = 0
    for so_row in result.rows:
        data_cell(ws, r, 1, so_row.po_number)
        data_cell(ws, r, 2, so_row.item_no)
        data_cell(ws, r, 3, so_row.ean)
        data_cell(ws, r, 4, so_row.description)
        data_cell(ws, r, 5, so_row.mrp,
                   '#,##0.00' if so_row.mrp else None)

        # Landing cost (MRP × margin%) — computed fresh for display so the
        # sheet stays self-consistent even if calc_price was derived
        # differently (e.g. RK uses cost basis, but we still want to show
        # Landing here for reference).
        landing = (float(so_row.mrp) * result.margin_pct
                   if so_row.mrp and not pd.isna(so_row.mrp) else None)
        data_cell(ws, r, 6,
                   round(landing, 2) if landing else '',
                   '#,##0.00')

        data_cell(ws, r, 7, so_row.gst_code)

        # Our Cost Price (naked CP) — always shown regardless of basis.
        data_cell(ws, r, 8,
                   round(so_row.cost_price_ref, 2)
                   if so_row.cost_price_ref else '',
                   '#,##0.00')

        # Marketplace value (fob_col)
        data_cell(ws, r, 9,
                   round(so_row.fob_price, 2) if so_row.fob_price else '',
                   '#,##0.00')

        # Difference (rounded to 2dp — finer is floating-point dust)
        data_cell(ws, r, 10,
                   round(so_row.diffn, 2) if so_row.diffn is not None else '',
                   '#,##0.00')

        data_cell(ws, r, status_col, so_row.validation_status)

        # ── Per-status row styling ──────────────────────────────────────
        if so_row.validation_status == 'MISMATCH':
            mismatches += 1
            for c in range(1, n_cols + 1):
                ws.cell(row=r, column=c).fill = MISMATCH_FILL
            ws.cell(row=r, column=status_col).font = MISMATCH_TEXT_FONT

        elif so_row.validation_status == 'OK':
            ws.cell(row=r, column=status_col).fill = STATUS_OK_FILL
            ws.cell(row=r, column=status_col).font = STATUS_OK_FONT

        elif so_row.validation_status == 'NOT_IN_MASTER':
            for c in range(1, n_cols + 1):
                ws.cell(row=r, column=c).fill = NO_MASTER_FILL
            ws.cell(row=r, column=status_col).font = NOT_IN_MASTER_TEXT_FONT

        # ── HSN columns (v1.6.0, only when marketplace opts in) ─────────
        if has_hsn_check:
            data_cell(ws, r, hsn_punch_col, so_row.hsn_punch)
            data_cell(ws, r, hsn_master_col, so_row.hsn_master)
            data_cell(ws, r, hsn_status_col, so_row.hsn_check_status)

            # Mirror the price-status pill styling on the HSN Check
            # cell so the user can scan the column for red at a glance.
            hsn_cell = ws.cell(row=r, column=hsn_status_col)
            if so_row.hsn_check_status == 'OK':
                hsn_cell.fill = STATUS_OK_FILL
                hsn_cell.font = STATUS_OK_FONT
            elif so_row.hsn_check_status == 'MISMATCH':
                hsn_mismatches += 1
                hsn_cell.font = MISMATCH_TEXT_FONT
            elif so_row.hsn_check_status == 'NOT_IN_MASTER':
                hsn_cell.font = NOT_IN_MASTER_TEXT_FONT

        r += 1

    # ── Footer summary ──────────────────────────────────────────────────
    r += 1
    total = len(result.rows)
    ok_count = sum(1 for so_row in result.rows
                    if so_row.validation_status == 'OK')
    basis_note = (f"basis={result.compare_basis} "
                  f"(compared against '{label}')")
    summary_parts = [
        f"Total: {total} items",
        f"OK: {ok_count}",
        f"Mismatches: {mismatches}",
        f"Margin: {margin_pct_int}%",
        basis_note,
    ]
    if has_hsn_check:
        # Only mention HSN stats when the check ran; otherwise the
        # zero in "HSN mismatches: 0" reads as a claim when really
        # the feature wasn't active.
        summary_parts.insert(3, f"HSN mismatches: {hsn_mismatches}")

    ws.cell(
        row=r, column=1,
        value=" | ".join(summary_parts),
    ).font = INFO_ITALIC_FONT

    auto_width(ws)
    ws.freeze_panes = 'A2'