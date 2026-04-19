"""
data.models
===========

Pure-data classes that flow through the pipeline. No I/O, no business
logic — only field definitions and lightweight container behaviour
provided by ``@dataclass``.

Two types::

    SORow             — one row per ordered item line
    ProcessingResult  — what the engine returns to the exporter

The exporter consumes ``ProcessingResult.rows`` (a list of ``SORow``)
together with marketplace metadata to render the output workbook.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any, List, Optional, Tuple


@dataclass
class SORow:
    """
    Single line item destined for the SO Lines (and contributing to the
    SO Header for its PO).

    Field groups
    ------------
    Identity (always populated):
        ``po_number``, ``location``, ``item_no``, ``qty``

    Mapping resolution (populated by MarketplaceEngine after Ship-To
    lookup):
        ``cust_no``, ``ship_to``, ``mapped``, ``mapped_location``

    Master lookup + price validation (populated when MasterLoader is
    available and the EAN/Item No can be resolved):
        ``ean``, ``description``, ``mrp``, ``gst_code``,
        ``cost_price_ref`` (always the post-GST naked CP),
        ``calc_price`` (what's used for the active diff — depends on
        the marketplace's ``compare_basis``),
        ``fob_price`` (marketplace's price column),
        ``ref_fob_price`` (optional reference marketplace price),
        ``diffn`` = ``fob_price - calc_price``,
        ``ref_diffn`` = ``ref_fob_price - cost_price_ref``,
        ``validation_status`` ∈ ``{'OK', 'MISMATCH', 'NOT_IN_MASTER',
        'NO_PRICE', ''}``

    Raw input pass-through:
        ``unit_price`` — only set when ``price_col`` is configured (rare;
        both current marketplaces leave Unit Price blank for the WMS to
        compute downstream).
    """

    # Identity
    po_number: str
    location: str
    item_no: Any
    qty: int

    # Pass-through input
    unit_price: Optional[float] = None

    # Mapping resolution
    cust_no: str = ''
    ship_to: str = ''
    mapped: bool = False
    mapped_location: str = ''   # canonical mapping key matched to (may
                                # differ from raw `location` due to
                                # case-insensitive or fuzzy match)

    # Master lookup
    ean: str = ''
    description: str = ''       # from Items_March 'Description'

    # Marketplace prices
    fob_price: Optional[float] = None
    ref_fob_price: Optional[float] = None  # purely for reference Diffn

    # Our calculations
    calc_price: Optional[float] = None      # value used for the active diff
                                            # (basis='cost'    → MRP×m%÷GST,
                                            #  basis='landing' → MRP×m%)
    cost_price_ref: Optional[float] = None  # ALWAYS post-GST cost price
                                            # (the "naked CP"), shown for
                                            # reference even when basis
                                            # ≠ 'cost'

    # Diffs
    diffn: Optional[float] = None     # active: fob_price - calc_price
    ref_diffn: Optional[float] = None  # reference: ref_fob_price - cost_price_ref

    # Master attributes (raw, for display)
    mrp: Optional[float] = None
    gst_code: str = ''

    # Validation status
    validation_status: str = ''


@dataclass
class ProcessingResult:
    """
    Output of ``MarketplaceEngine.process()``. Consumed by ``SOExporter``.

    Holds the rendered SORows plus the metadata the exporter needs to
    label sheets correctly (compare_basis affects Validation column
    headers, etc.).
    """

    # Per-row results
    rows: List[SORow] = field(default_factory=list)

    # Warnings to surface in the GUI log + Warnings sheet:
    # tuples of (po, location, message). PO and location may be empty
    # strings for global warnings.
    warnings: List[Tuple[str, str, str]] = field(default_factory=list)

    # Marketplace context
    marketplace: str = ''
    input_file: str = ''            # basename, for display
    input_file_path: str = ''       # full path — used to locate the
                                    # output/ folder next to the input
    margin_pct: float = 0.70        # decimal (0.70 = 70%)
    compare_basis: str = 'cost'     # 'landing' | 'cost'
    compare_label: str = 'Price'    # friendly label shown in Validation

    # Original DataFrame (for the Raw Data sheet)
    raw_df: Any = None
