"""
config.marketplaces
===================

Per-marketplace configuration registry.

Each marketplace produces PO/punch files with its own column layout. To
support a new marketplace, add an entry to ``MARKETPLACE_CONFIGS`` with
the keys documented below — the rest of the pipeline is config-driven and
needs no further code changes for the common cases.

Config schema
-------------

Required for every marketplace:

``party_name`` (str)
    Must match the ``Party`` column in the Ship-To B2B mapping sheet
    exactly (case-insensitive). Used to filter the mapping registry to
    just this marketplace's locations.
``po_col`` (str)
    Column containing the PO/SO number.
``loc_col`` (str)
    Column containing the delivery location (matched against ``Del
    Location`` in the mapping sheet).
``qty_col`` (str)
    Column containing the order quantity.
``item_resolution`` (str, ``'from_column'`` | ``'from_ean'``, default
``'from_column'``)
    How to determine the canonical Item No for each row:

    * ``'from_column'`` — take it directly from ``item_col``. Use when the
      marketplace already provides a pre-resolved Item No.
    * ``'from_ean'`` — look the EAN up in Items_March and use
      ``master_info['item_no']``. Use when the marketplace only provides
      EAN/GTIN (real Myntra and RK files).

``item_col`` (str, required when ``item_resolution='from_column'``)
    Column containing the Item No.
``ean_col`` (str, required when ``item_resolution='from_ean'``)
    Column containing the EAN/GTIN. Also used for price-validation lookups
    even when ``item_resolution='from_column'``.

Optional but recommended:

``fob_col`` (str)
    Column containing the marketplace price to validate against.
``ref_fob_col`` (str)
    Optional second marketplace price column shown only as a *reference*
    Diffn in Raw Data (no effect on OK/MISMATCH status).
``compare_basis`` (str, ``'landing'`` | ``'cost'``, default ``'cost'``)
    What we compare ``fob_col`` against:

    * ``'landing'`` → ``MRP × margin%`` (pre-GST). Used by Myntra because
      its "Landing Price" column is the pre-GST value.
    * ``'cost'`` → ``MRP × margin% ÷ GST`` (post-GST). Used by RK.

``compare_label`` (str, default ``'Price'``)
    Friendly label shown in the Validation sheet. E.g. ``'Landing Rate'``
    yields ``"Marketplace Landing Rate"`` and ``"Difference with Landing
    Rate"`` column headers.
``default_margin`` (int, default 70)
    Default margin % pre-filled in the GUI when this marketplace is
    selected. The user can override per-run via the GUI input.
``price_col`` (str | None)
    Column containing unit price for the SO Lines output (rare — both
    current marketplaces leave it None so the WMS computes it).
``template_headers`` (list[str])
    Full column list used when the user clicks "Download PO Template". If
    omitted, a minimal list is built from the required + validation cols.

PO template colour legend
-------------------------
When the user downloads a template, headers are colour-coded by role:

* **BLUE** (``#1A237E``) — Required. Script fails without these.
* **GREEN** (``#1B5E20``) — Validation. Used for price check & master
  lookup.
* **GREY** (``#9E9E9E``) — Not read by script. Kept only to mirror the
  marketplace's native file format.
"""

from __future__ import annotations
from typing import Any, Dict, List


MARKETPLACE_CONFIGS: Dict[str, Dict[str, Any]] = {
    # ────────────────────────────────────────────────────────────────────
    # MYNTRA
    # Real Myntra files have NO 'Item no' column — only 'GTIN' and
    # 'Vendor Article Number' (both carry the EAN). Item No is resolved
    # from EAN via the Items_March master.
    # ────────────────────────────────────────────────────────────────────
    'Myntra': {
        'party_name': 'Myntra',
        'po_col': 'PO',                                       # [REQUIRED]
        'loc_col': 'Location',                                # [REQUIRED]
        'qty_col': 'Quantity',                                # [REQUIRED]
        'item_resolution': 'from_ean',                        # see schema
        'ean_col': 'Vendor Article Number',                   # [REQUIRED in this mode]
        'price_col': None,                                    # WMS computes
        'fob_col': 'Landing Price',                           # [VALIDATION]
        # ref_fob_col surfaces "what would the diff have been against List
        # price?" alongside the active diff. Reference only — has zero
        # effect on OK/MISMATCH status.
        'ref_fob_col': 'List price(FOB+Transport-Excise)',
        'default_margin': 70,
        'compare_basis': 'landing',                           # MRP × m%
        'compare_label': 'Landing Rate',
        'template_headers': [
            'PO', 'Location', 'SKU Id', 'Style Id', 'SKU Code',
            'HSN Code', 'Brand', 'GTIN', 'Vendor Article Number',
            'Vendor Article Name', 'Size', 'Colour', 'Mrp',
            'Credit Period', 'Margin Type', 'Agreed Margin',
            'Gross Margin', 'Quantity', 'FOB Amount',
            'List price(FOB+Transport-Excise)', 'Landing Price',
            'Estimated Delivery Date',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # RK
    # Real RK files also lack an 'Item no' column — they expose EAN as
    # 'External ID'. Same resolution mechanism as Myntra. Compare basis
    # is 'cost' because RK's Cost column is post-GST, matching our
    # MRP × 70% ÷ 1.18 to the paisa.
    # ────────────────────────────────────────────────────────────────────
    'RK': {
        'party_name': 'RK',
        'po_col': 'PO',                                       # [REQUIRED]
        'loc_col': 'Ship-to location',                        # [REQUIRED]
        'qty_col': 'Accepted quantity',                       # [REQUIRED]
        'item_resolution': 'from_ean',
        'ean_col': 'External ID',                             # [REQUIRED in this mode]
        'price_col': None,
        'fob_col': 'Cost',                                    # [VALIDATION]
        'default_margin': 70,
        'compare_basis': 'cost',                              # MRP × m% ÷ GST
        'compare_label': 'Cost',
        'template_headers': [
            'PO', 'Vendor code', 'Order date', 'Product name',
            'External ID', 'Accepted quantity', 'Ship-to location',
            'Cost', 'Total accepted cost',
        ],
    },

    # ┌─────────────────────────────────────────────────────────────────┐
    # │ ADD NEW MARKETPLACES HERE                                       │
    # │                                                                 │
    # │ 'Bigbasket': {                                                  │
    # │     'party_name': 'Bigbasket',                                  │
    # │     'po_col': 'PO Number',                                      │
    # │     'loc_col': 'Delivery Location',                             │
    # │     'qty_col': 'Qty',                                           │
    # │     'item_resolution': 'from_ean',                              │
    # │     'ean_col': 'EAN',                                           │
    # │     'fob_col': 'Unit Price',                                    │
    # │     'price_col': None,                                          │
    # │     'default_margin': 60,                                       │
    # │     'compare_basis': 'cost',                                    │
    # │     'compare_label': 'Cost',                                    │
    # │     'template_headers': [...],                                  │
    # │ },                                                              │
    # └─────────────────────────────────────────────────────────────────┘
}


# Marketplace names for the GUI dropdown (insertion order = display order).
MARKETPLACE_NAMES: List[str] = list(MARKETPLACE_CONFIGS.keys())
