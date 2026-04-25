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
``po_col`` (str | list[str])
    Column containing the PO/SO number. Use a plain string for the
    common case of a single known column name. Use a list when the
    marketplace's punch file sometimes arrives with variant headers
    — the engine will pick the first list entry that actually
    appears in the file. List order is preference order; e.g.
    ``['PO', 'PO Number']`` means "use 'PO' when present, fall back
    to 'PO Number'". Myntra uses this because its dumps sometimes
    label the column 'PO' and sometimes 'PO Number'.
``loc_col`` (str | list[str])
    Column containing the delivery location (matched against ``Del
    Location`` in the mapping sheet). Supports the same list
    fallback as ``po_col``.
``qty_col`` (str | list[str])
    Column containing the order quantity. Supports list fallback.
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
``amount_col`` (str | dict | None, default omitted)
    Per-row monetary amount source. When configured, the engine reads
    this into ``SORow.amount`` and the email report sums it for the
    headline "Amount" stat.

    Accepted forms:

    1. ``str`` — single column name on the punch file::

           'amount_col': 'total_amount'            # Blink
           'amount_col': 'Total accepted cost'     # RK (qty × cost)

    2. ``{'multiply': [col_a, col_b, ...]}`` — product of columns.
       Used when the marketplace doesn't carry a pre-calculated
       amount but does carry the factors (added v1.5.7)::

           'amount_col': {'multiply': ['Landing Price', 'Quantity']}
           # Myntra — landing value pre-tax

    3. ``{'multiply': [...], 'apply_margin': True}`` — product of
       columns, then multiplied by the run's ``margin_pct`` (v1.6.0).
       Used when one conceptual factor is the derived Landing Cost
       rather than a raw column::

           'amount_col': {'multiply': ['MRP', 'Qty'], 'apply_margin': True}
           # Reliance — Landing × Qty = (MRP × margin%) × Qty

    Omit the key entirely when the marketplace has no usable amount
    source; the email aggregator treats absent ``amount`` as 0.

``hsn_col`` (str, optional, v1.6.0)
    Column on the punch file carrying the HSN/SAC code. When set,
    the engine cross-checks it against the master's ``HSN/SAC Code``
    column and records an ``hsn_check_status`` on each ``SORow``
    (``'OK'``, ``'MISMATCH'``, or ``'NOT_IN_MASTER'``). Mismatches
    produce a deduped per-item warning and surface on the Validation
    sheet. Currently Reliance only. Other marketplaces carry HSN
    columns but we don't cross-check them by default because their
    HSN data isn't contractually authoritative the way Reliance's is.

``source_sheet`` (str, default ``'Sheet1'``)
    Name of the sheet to read inside the punch workbook. Supports
    two match modes:

    * **Exact** — plain sheet name. Reliance uses ``'PO'``; the
      default ``'Sheet1'`` covers Blink/Myntra/RK.
    * **Wildcard prefix** (v1.8.0) — trailing ``*`` matches any
      sheet whose name starts with the given prefix. Zepto uses
      ``'PO_*'`` because its data sheet is literally named
      ``'PO_<random-hex>'`` (different hex every dump).

    Behavior on miss:

    * Exact miss falls back to the first sheet in the workbook and
      logs a warning (kind to files with reordered sheets).
    * Wildcard miss aborts with an explicit error — the wildcard
      implies a specific marketplace's expected sheet shape, so
      silently reading a different sheet would produce garbage.

``header_row`` (int, default ``0``)
    0-indexed row that pandas treats as column headers. Reliance's
    PO sheet has a merged title on row 0 so its real headers are on
    row 1. Most marketplaces leave this at 0.

``pre_process`` (str, optional)
    Name of a marketplace-specific hook the engine runs after
    loading the sheet but before column validation. Values:
    ``'reliance_po_sheet'`` parses row 0's title
    (``"5000466441  BHIWANDI (Reliance)"``) into synthetic
    ``__po__`` and ``__loc__`` columns so Reliance's out-of-band
    PO number and location reach the engine in canonical form.
    Omit when the marketplace's data is already wide-format.

``default_margin`` (int | float, default 70)
    Default margin % pre-filled in the GUI when this marketplace is
    selected. The user can override per-run via the GUI input.
    Typically an integer (70 for Blink/RK/Myntra); Reliance uses
    the decimal ``63.42``.

``case_insensitive_cols`` (bool, default False, v1.8.1)
    When True, the engine's column-name resolution matches ``*_col``
    values against the DataFrame's actual headers **case-insensitively**
    and **whitespace-tolerantly** (surrounding whitespace is ignored,
    internal multi-space runs collapse to one). Protects against
    marketplace dashboards that ship the same semantic column under
    varying casings/spacings across dumps. Currently enabled on:

    * **Myntra** — PO header seen as ``'PO'``, ``'PO Number'``, and
      ``'Po number'`` across recent dumps.
    * **Reliance** — HSN header seen as ``'HSN'`` (older batches)
      and ``'Hsn'`` (newer batches).

    Omit/False for marketplaces with stable, contractual column
    headers (Blink/RK/Zepto today) so typo-in-template errors fail
    loud rather than silently match something unintended.

``override_unit_price`` (bool, default False, v1.9.1)
    When True, the D365 exporter populates Sales Line col H (Unit
    Price) with our computed post-GST Cost Price per row rather than
    leaving it blank. Blank is the normal default because the ERP
    auto-computes unit price from the vendor master's margin
    setting downstream — but when the ERP's recorded margin for the
    vendor doesn't match the marketplace's actual margin, the
    auto-computed cost is wrong and has to be overridden.

    Canonical example: BCPL (Blinkit + BlinkMP share a single vendor
    record in Business Central with 70% margin — Blinkit's rate).
    BlinkMP operates at 75%, so without this override every BlinkMP
    row in the D365 import would post with Blinkit's 70% cost,
    understating BlinkMP's margin. Enabling the flag stamps the
    correct cost (computed via the engine's standard
    ``MRP × margin% ÷ GST-divisor`` formula, per-item GST rate) into
    col H so the ERP records what BlinkMP actually charged us.

    Rows without a computable Cost Price (master lookup failed)
    still emit with col H blank — the engine logs a per-row warning
    and continues. Matches the lenient behavior elsewhere in the
    pipeline for missing-master rows.

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
        # v1.8.1: Myntra's dashboard exports have historically varied
        # the case of column headers between dumps. We've seen at
        # least three casings of the PO column header ('PO', 'PO
        # Number', 'Po number') in a two-week span. Enabling
        # case-insensitive column matching absorbs these variations
        # automatically — the engine's ``_resolve_column_aliases``
        # will find whichever header exists regardless of case when
        # this flag is True.
        'case_insensitive_cols': True,
        # v1.5.5: ``po_col`` is a list because Myntra dumps arrive
        # with either 'PO' or 'PO Number' as the header — list order
        # is the preference, so 'PO' wins when both exist. The engine
        # collapses this to a single string after loading the file
        # (see MarketplaceEngine._resolve_column_aliases).
        'po_col': ['PO', 'PO Number'],                        # [REQUIRED]
        'loc_col': 'Location',                                # [REQUIRED]
        'qty_col': 'Quantity',                                # [REQUIRED]
        'item_resolution': 'from_ean',                        # see schema
        # v1.5.8: EAN source switched from 'Vendor Article Number' to
        # 'GTIN'. Semantically GTIN is the EAN — "Vendor Article
        # Number" is Myntra's field for the seller's internal SKU
        # code, and even though Renee has historically populated it
        # with the EAN (so the two columns matched 237/242 rows in
        # the 17-04 file), relying on that is fragile — Myntra could
        # diverge them any time. GTIN is the semantically correct
        # column. Rows where GTIN is blank or unreadable follow the
        # standard missing-EAN warning flow (NOT_IN_MASTER in the
        # Validation sheet + a row-level warning).
        'ean_col': 'GTIN',                                    # [REQUIRED in this mode]
        'price_col': None,                                    # WMS computes
        'fob_col': 'Landing Price',                           # [VALIDATION]
        # ref_fob_col surfaces "what would the diff have been against List
        # price?" alongside the active diff. Reference only — has zero
        # effect on OK/MISMATCH status.
        'ref_fob_col': 'List price(FOB+Transport-Excise)',
        # v1.5.7: Myntra dumps don't carry a pre-calculated per-row
        # amount column (unlike Blink's 'total_amount' or RK's 'Total
        # accepted cost'), but they do carry the factors. We compute
        # Landing × Qty per row which is the pre-tax landing value
        # — matches what the finance team reconciles against on the
        # Myntra invoice. The multiply-spec form is handled by the
        # engine's ``_extract_amount`` helper.
        'amount_col': {'multiply': ['Landing Price', 'Quantity']},
        'default_margin': 70,
        'compare_basis': 'landing',                           # MRP × m%
        'compare_label': 'Landing Rate',
        'template_headers': [
            'PO/PO Number', 'Location', 'SKU Id', 'Style Id', 'SKU Code',
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
        'amount_col': 'Total accepted cost',                  # qty × cost — for email Amount stat
        'template_headers': [
            'PO', 'Vendor code', 'Order date', 'Product name',
            'External ID', 'Accepted quantity', 'Ship-to location',
            'Cost', 'Total accepted cost',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # Blink  (a.k.a. Blinkit — Blink Commerce Private Limited, "BCPL".
    # The dropdown and party_name both use "Blink" to match what's in
    # the Ship-To B2B registry — mapping lookups filter on exactly that
    # string, so the config must agree with the sheet character-for-
    # character.)
    #
    # Real Blink files carry line data on 'Sheet1' (same as all other
    # marketplaces — see note in engine.marketplace_engine). They may
    # also ship pivot / sidecar sheets (Sheet2, Sheet4) with the user's
    # own reference data; those are intentionally ignored by the engine.
    #
    # Columns on Sheet1 (25 native):
    #   po_number | facility_name | manufacturer_name |
    #   entity_vendor_legal_name | vendor_name | order_date |
    #   appointment_date | expiry_date | po_state | item_id | name |
    #   uom_text | upc | units_ordered | remaining_quantity |
    #   landing_rate | cost_price | margin_percentage | cess_value |
    #   sgst_value | igst_value | cgst_value | tax_value |
    #   total_amount | mrp
    #
    # Key specifics:
    #   • No 'Item no' column → resolve from 'upc' (EAN) via master lookup.
    #   • 'margin_percentage' = 30 is Blink's margin → ours is 70%.
    #   • 'cost_price' is POST-GST (landing ÷ 1.18), compare_basis='cost'.
    #   • 'landing_rate' is pre-GST (MRP × 70%), present but not used for
    #     validation — could be added as ref_fob_col if reference diff is
    #     wanted later.
    #
    # 'Blink RO' (reverse-order / return entity) will be added as a
    # sibling entry once we have a sample file.
    # ────────────────────────────────────────────────────────────────────
    'Blink': {
        'party_name': 'Blink',               # Must match 'Party' in mapping sheet
        'po_col': 'po_number',               # [REQUIRED] int64 (e.g. 1723710027417)
        'loc_col': 'facility_name',          # [REQUIRED] e.g. 'Pune P2 - Feeder Warehouse'
        'qty_col': 'units_ordered',          # [REQUIRED]
        'item_resolution': 'from_ean',
        'ean_col': 'upc',                    # [REQUIRED in this mode]
        'price_col': None,                   # WMS computes
        'fob_col': 'cost_price',             # [VALIDATION] post-GST
        'default_margin': 70,                # 100 - Blink's margin_percentage(30)
        'compare_basis': 'cost',             # cost_price ≈ MRP × 70% ÷ 1.18
        'compare_label': 'Cost',
        'amount_col': 'total_amount',        # incl. taxes — for email Amount stat
        'template_headers': [
            'po_number', 'facility_name', 'manufacturer_name',
            'entity_vendor_legal_name', 'vendor_name',
            'order_date', 'appointment_date', 'expiry_date', 'po_state',
            'item_id', 'name', 'uom_text', 'upc',
            'units_ordered', 'remaining_quantity',
            'landing_rate', 'cost_price', 'margin_percentage',
            'cess_value', 'sgst_value', 'igst_value', 'cgst_value',
            'tax_value', 'total_amount', 'mrp',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # Reliance (v1.6.0)
    # Reliance's raw PO attachment is a 6-sheet workbook with clean data
    # on a sheet literally named 'PO'. The PO number and delivery
    # location don't appear in data columns — they're in a merged-cell
    # title on row 0, formatted like:
    #
    #     5000466441  BHIWANDI (Reliance)
    #
    # The ``pre_process: 'reliance_po_sheet'`` hook parses this title
    # and injects synthetic ``__po__`` / ``__loc__`` columns on every
    # data row so the rest of the engine runs normally.
    #
    # HSN cross-check is enabled (``hsn_col: 'HSN'``) — mismatches
    # between Reliance's HSN and our master's HSN/SAC Code surface on
    # the Validation sheet and in the Warnings sheet. This is critical
    # for tax compliance because a wrong HSN → wrong GST rate → the ERP
    # posts an incorrect tax amount.
    #
    # Margin: Reliance's effective landing% is 63.42 (reverse-engineered
    # from sample data: Cost Price / MRP × 1.18 ≈ 0.6342 consistently).
    # User can override in the GUI on a per-run basis.
    # ────────────────────────────────────────────────────────────────────
    'Reliance': {
        'party_name': 'Reliance',            # Must match 'Party' in mapping sheet
        # v1.8.1: Reliance's dumps have shipped with both 'HSN' and
        # 'Hsn' as the HSN column header across batches (we observed
        # both within a single week of samples). Enabling
        # case-insensitive column matching lets the engine find the
        # right column regardless of casing — the match is done by
        # ``_resolve_column_aliases`` using a lowercase comparison.
        'case_insensitive_cols': True,
        'source_sheet': 'PO',                # NOT 'Sheet1' — special-case for Reliance
        'header_row': 1,                     # Row 0 is the merged title; real headers are on row 1
        'pre_process': 'reliance_po_sheet',  # Engine hook that parses title + injects __po__/__loc__
        'po_col': '__po__',                  # [REQUIRED] synthetic — set by pre_process
        'loc_col': '__loc__',                # [REQUIRED] synthetic — set by pre_process
        'qty_col': 'Qty',                    # [REQUIRED]
        'item_resolution': 'from_ean',
        'ean_col': 'EAN Number',             # [REQUIRED in this mode]
        'price_col': None,                   # WMS computes
        'fob_col': 'Cost Price',             # [VALIDATION] Reliance's stated cost → compare against our MRP × margin%
        'hsn_col': 'HSN',                    # v1.6.0 — enables HSN cross-check against master
        'default_margin': 63.42,             # 63.42% — Reliance's operating margin
        # Reliance's 'Cost Price' column is POST-GST (reverse-engineered:
        # 250 × 0.6342 ÷ 1.18 = 134.36 matches their value exactly), so
        # we compare against our post-GST cost. Same basis RK uses.
        # Note: the GUI shows "Landing Rate" labelling because the
        # ERP/accounts team calls this "Landing Cost" even though
        # technically it's post-GST — we preserve their terminology.
        'compare_basis': 'cost',
        'compare_label': 'Cost',
        # Amount per row = Landing Cost × Qty = (MRP × margin%) × Qty.
        # The apply_margin flag tells the engine to multiply the
        # column-product by the runtime margin_pct, giving us the
        # correct amount even though Landing Cost isn't a column on
        # the punch file (it's derived).
        'amount_col': {'multiply': ['MRP', 'Qty'], 'apply_margin': True},
        'template_headers': [
            'MRP', 'Cost Price', 'Reliance Article Code', 'HSN',
            'EAN Number', 'PRODUCT NAME', 'Qty',
            'Batch Num', 'Mfg Date', 'Exp Date', 'Box No.',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # Zepto (v1.8.0)
    # Zepto's raw PO dump bundles many POs (10-15 typical) into a
    # single data sheet, but that sheet's NAME changes every dump —
    # it's literally 'PO_<random-hex>' like 'PO_64863340b23e6c90' or
    # 'PO_c881cfb0a4fa2ebc'. We use a wildcard ``source_sheet`` to
    # match any such prefix. The engine aborts with a clear error if
    # no matching sheet is found (vs. Reliance's behavior of failing
    # over to the first sheet — that would be wrong here because
    # Zepto workbooks usually have other sheets like 'Sheet1' or
    # 'Skipped' that we must NOT read).
    #
    # Math: same as Blink — ``Unit Base Cost`` is POST-GST, so
    # ``compare_basis='cost'``. ``Total Amount`` is pre-calculated
    # on the punch so we just read it directly.
    #
    # Ignored by design (per user decision):
    #   * 'Sheet1' / 'Sheet2' — user's manual ship-to map (we use
    #     Ship-To B2B.xlsx instead, which is location-keyed and
    #     more stable than per-PO mappings).
    #   * 'Skipped' — user's own bucket for high-diff items they
    #     chose to exclude from the SO. Engine treats it as
    #     invisible.
    #   * HSN column is present but we DON'T cross-check it
    #     (Reliance-only feature).
    # ────────────────────────────────────────────────────────────────────
    'Zepto': {
        'party_name': 'Zepto',           # Must match 'Party' in mapping sheet
        'source_sheet': 'PO_*',          # v1.8.0 — wildcard prefix match
        'po_col': 'PO No.',              # [REQUIRED]
        'loc_col': 'Del Location',       # [REQUIRED] e.g. 'CHN-SS-MH-THIRUVALLUR'
        'qty_col': 'Qty',                # [REQUIRED]
        'item_resolution': 'from_ean',
        'ean_col': 'EAN',                # [REQUIRED in this mode]
        'price_col': None,               # WMS computes
        'fob_col': 'Unit Base Cost',     # [VALIDATION] post-GST (same as Blink's cost_price)
        'default_margin': 70,
        'compare_basis': 'cost',         # Unit Base Cost ≈ MRP × 70% ÷ 1.18
        'compare_label': 'Cost',
        'amount_col': 'Total Amount',    # Zepto's pre-calc: Unit Base Cost × Qty
        'template_headers': [
            'PO No.', 'PO Date', 'Status', 'Vendor Code', 'Vendor Name',
            'PO Amount', 'Del Location', 'Line No', 'SKU', 'SKU Code',
            'SKU Desc', 'Brand', 'EAN', 'HSN',
            'CGST %', 'SGST %', 'IGST %', 'CESS %',
            'MRP/RSP', 'Qty', 'Unit Base Cost', 'Landing Cost',
            'Total Amount',
            'Created By', 'ASN Quantity', 'GRN Quantity', 'PO Expiry Date',
        ],
    },

    # ────────────────────────────────────────────────────────────────────
    # BlinkMP (v1.9.0) — distinct from "Blink" (Blinkit-BCPL) despite
    # similar naming.
    #
    # Blink (Blinkit-BCPL) — the quick-commerce dark stores, handles
    #                        PO number in column 'po_number',
    #                        post-GST cost in 'cost_price', 70% margin.
    # BlinkMP (BCPL-RO)   — the reorder/wholesale channel, PO column
    #                        is simply 'PO', has 'Landing Rate'
    #                        (pre-GST) instead of post-GST cost,
    #                        75% margin.
    #
    # Both ship into BCPL's warehouse network (Lucknow L4, Kundli,
    # Mumbai M10, etc.) which is why the location strings look
    # similar, but they're separate commercial relationships with
    # their own Ship-To B2B rows. Do NOT conflate them.
    # ────────────────────────────────────────────────────────────────────
    'BlinkMP': {
        # v1.9.3: ``party_name`` is 'Blink RO' (BCPL's internal name
        # for the Reorder channel) to match the existing mapping rows
        # in Ship-To B2B.xlsx. The user-facing marketplace name stays
        # 'BlinkMP' (that's the dict key), so the GUI dropdown still
        # shows "BlinkMP" — but behind the scenes the mapping lookup
        # uses the canonical ERP name. Lets the GUI label and the
        # operational spreadsheet stay independently correct.
        'party_name': 'Blink RO',            # Must match 'Party' in mapping sheet
        'po_col': 'PO',                      # [REQUIRED] int64 (e.g. 1735810027046)
        'loc_col': 'Location',               # [REQUIRED] e.g. 'BCPL-Mumbai M10 - Feeder Warehouse'
        'qty_col': 'Quantity',               # [REQUIRED]
        'item_resolution': 'from_ean',
        'ean_col': 'Product UPC',            # [REQUIRED in this mode]
        'price_col': None,                   # WMS computes
        'fob_col': 'Landing Rate',           # [VALIDATION] pre-GST (= MRP × 75%)
        'default_margin': 75,                # 75% margin (higher than Blink's 70%)
        'compare_basis': 'landing',          # Landing Rate IS pre-GST (unlike Blink's post-GST cost_price)
        'compare_label': 'Landing Rate',
        'amount_col': 'Total Amount',        # pre-calculated by BlinkMP: Landing Rate × Qty
        # v1.9.1: Override the D365 Sales Line 'Unit Price' (col H)
        # with our computed post-GST Cost Price. BCPL (Blinkit +
        # BlinkMP) is registered in Business Central with 70% margin
        # because that's Blinkit's rate — but BlinkMP's margin is
        # 75%. Without the override the ERP auto-computes the wrong
        # cost on BlinkMP rows (using the 70% vendor default),
        # creating accounting discrepancies. Stamping our own
        # post-GST Cost Price into col H forces the ERP to record
        # the correct figure.
        #
        # When a row has no computable Cost Price (e.g. master
        # lookup failed), col H is left empty for that specific row
        # and a log warning is emitted — matches the engine's
        # existing lenient handling for NOT_IN_MASTER rows.
        'override_unit_price': True,
        'template_headers': [
            'PO', 'Location', 'Item Code', 'HSN Code', 'Product UPC',
            'Product Description', 'Grammage',
            'CGST %', 'SGST %', 'IGST %', 'CESS %', 'Additional CES',
            'Tax Amount', 'Landing Rate', 'Quantity', 'MRP',
            'Total Amount',
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


# ════════════════════════════════════════════════════════════════════════
# Warehouse codes (v1.9.0)
# ════════════════════════════════════════════════════════════════════════
#
# Purpose
# -------
# The D365 Sales Header (col K) and Sales Line (col F) require a valid
# Business Central location code on every row. Until v1.9.0 we hardcoded
# ``PICK`` (Ahmedabad's warehouse code) because it was the only warehouse
# in use. As of v1.9.0 the user can pick between warehouses at run-time
# via a GUI dropdown next to the Margin input.
#
# The dropdown shows the short friendly code (AHD, BLR); the engine
# substitutes the full ERP code (``PICK``, ``DS_BL_OFF1``) into the D365
# file. Users don't have to remember which cryptic string maps to what
# city — they pick by city name and the translation happens under the
# hood.
#
# Adding a new warehouse
# ----------------------
# 1. Add an entry here, e.g. ``'MUM': 'DS_MUM_WH1'``.
# 2. That's it — the GUI dropdown auto-populates from this dict, and
#    D365 export auto-uses the mapped value when the user selects it.
#
# The FIRST entry is treated as the GUI's default selection. Keep AHD
# at the top unless the operational default changes.
#
# Dict values are the literal strings Business Central expects. Get them
# from the ERP team — typos here silently produce D365 imports the ERP
# will reject.
# ════════════════════════════════════════════════════════════════════════

WAREHOUSE_CODES: Dict[str, str] = {
    'AHD': 'PICK',          # Ahmedabad (default) — PICK warehouse
    'BLR': 'DS_BL_OFF1',    # Bangalore — Bangalore office 1 warehouse
    # Add more warehouses as needed, e.g.:
    #   'MUM': 'DS_MUM_WH1',
    #   'DEL': 'DS_DEL_WH1',
}

WAREHOUSE_DISPLAY_NAMES: List[str] = list(WAREHOUSE_CODES.keys())
DEFAULT_WAREHOUSE: str = WAREHOUSE_DISPLAY_NAMES[0]  # 'AHD'