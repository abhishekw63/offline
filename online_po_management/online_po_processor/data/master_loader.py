"""
data.master_loader
==================

Loads ``Items_March.xlsx`` (the canonical product master) into an
in-memory lookup table indexed by both GTIN/EAN and Item No.

The master is the source of truth for:

* ``MRP`` — used to compute Landing Cost (``MRP × margin%``) and post-GST
  Cost Price (``... ÷ GST divisor``).
* ``GST Group Code`` — drives which divisor we apply.
* ``Description`` — surfaced in the Validation sheet next to the EAN so
  the user can read what each item actually is at a glance.
* ``No.`` (Item No) — the canonical ERP code resolved when the
  marketplace only provides an EAN.

The static helpers ``calc_cost_price`` and ``calc_landing_price`` are
exposed as classmethods so the engine can call them without holding a
loader instance.
"""

from __future__ import annotations
from typing import Dict, Optional

import pandas as pd


class MasterLoader:
    """
    In-memory lookup over the Items Master file.

    Indexes each row by BOTH the stringified GTIN and the stringified Item
    No (``No.``). This means callers can look up by either an EAN
    (Myntra-style) or an Item No (RK-style) using the same ``lookup()``
    method.
    """

    def __init__(self) -> None:
        # Map from key (EAN or Item No, both as str) → entry dict
        # Entry shape: {item_no, mrp, gst_code, description}
        self.master: Dict[str, Dict] = {}

    # ── Loading ────────────────────────────────────────────────────────

    def load(self, filepath: str) -> int:
        """
        Read the master file and rebuild the lookup table.

        Args:
            filepath: Path to ``Items_March.xlsx``.

        Returns:
            Number of rows loaded (also the row count of the input).

        Required columns: ``No.``, ``GTIN``, ``Description``,
        ``GST Group Code``, ``Mrp``.
        """
        df = pd.read_excel(filepath, header=0)

        # Pre-stringify GTIN for use as a dict key. Done once on the whole
        # column so the per-row loop stays cheap.
        df['GTIN_str'] = df['GTIN'].astype(str).str.strip()

        self.master = {}

        for _, r in df.iterrows():
            desc = (str(r.get('Description', ''))
                    if pd.notna(r.get('Description')) else '')
            gst = (str(r['GST Group Code'])
                   if pd.notna(r.get('GST Group Code')) else '')
            mrp = r.get('Mrp')
            item_no = str(r['No.']).strip()

            entry = {
                'item_no': item_no,
                'mrp': mrp,
                'gst_code': gst,
                'description': desc,
            }

            # Index by GTIN. The GTIN is the marketplace-facing identifier,
            # so EAN lookups go through this key.
            self.master[r['GTIN_str']] = entry

            # Also index by item code so a punch file with pre-resolved
            # Item No (RK-style, when ``item_resolution='from_column'``)
            # can find the entry too. Don't overwrite an existing GTIN
            # match — GTIN is more specific.
            if item_no not in self.master:
                self.master[item_no] = entry

        return len(df)

    # ── Lookup ─────────────────────────────────────────────────────────

    def lookup(self, key: str) -> Optional[Dict]:
        """
        Find an entry by EAN or Item No.

        Tries the cleaned key first, then falls back to leading-zero-
        stripped form (EANs sometimes have a leading zero in source data
        but not in the master).

        Args:
            key: Stringified EAN (e.g. ``'8906121642599'``) or Item No
                 (e.g. ``'200074'``).

        Returns:
            ``{item_no, mrp, gst_code, description}`` dict on hit, ``None``
            on miss.
        """
        key_clean = str(key).strip()
        if key_clean in self.master:
            return self.master[key_clean]

        # Some sources include a leading zero on EANs that the master
        # file omits — try the trimmed form before giving up.
        stripped = key_clean.lstrip('0')
        if stripped in self.master:
            return self.master[stripped]

        return None

    # ── Pricing helpers (static) ───────────────────────────────────────
    # These are called on every row by the engine. Kept as static methods
    # so the engine doesn't need a loader instance to compute them, and
    # so they're trivially unit-testable.

    @staticmethod
    def calc_cost_price(mrp, gst_code: str,
                        margin_pct: float) -> Optional[float]:
        """
        Post-GST Cost Price: ``MRP × margin% ÷ GST divisor``.

        The GST divisor depends on the master's ``GST Group Code``:

        =========  =====  =========
        Code       GST    Divisor
        =========  =====  =========
        0-G        0%     ÷ 1.00
        G-3        3%     ÷ 1.03
        G-5(-S)    5%     ÷ 1.05
        G-12(-S)   12%    ÷ 1.12
        G-18(-S)   18%    ÷ 1.18
        Unknown    -      ÷ 1.18 (defaults to 18% with engine warning)
        =========  =====  =========

        Args:
            mrp: Maximum Retail Price (may be ``None`` or NaN).
            gst_code: Tax code from ``Items_March['GST Group Code']``.
            margin_pct: Margin as decimal (e.g. ``0.70`` for 70%).

        Returns:
            Calculated cost price, or ``None`` if MRP is missing.
        """
        if mrp is None or pd.isna(mrp):
            return None

        landing = float(mrp) * margin_pct
        gst = str(gst_code).strip().upper()

        # 0% GST — code variants seen in the wild
        if gst in ('0-G', 'G-0', 'G-0-S', '0', '') or gst == 'NAN':
            return landing
        # 3% GST
        if gst in ('G-3', 'G-3-S'):
            return landing / 1.03
        # 5% GST — accept "5" in code as long as it's not 12 or 18
        if '5' in gst and '18' not in gst and '12' not in gst:
            return landing / 1.05
        # 12% GST
        if '12' in gst:
            return landing / 1.12
        # 18% GST
        if '18' in gst:
            return landing / 1.18
        # Unknown code — default to 18% (engine emits a warning separately)
        return landing / 1.18

    @staticmethod
    def calc_landing_price(mrp,
                           margin_pct: float) -> Optional[float]:
        """
        Pre-GST Landing Rate: ``MRP × margin%``. No GST divisor.

        Used by marketplaces whose price column is itself pre-GST (e.g.
        Myntra's "Landing Price"). Avoiding GST division means the diff
        comes out cleanly to zero on a correctly-priced punch — no
        floating-point rounding noise from ``÷ 1.18``.

        Args:
            mrp: Maximum Retail Price (may be ``None`` or NaN).
            margin_pct: Margin as decimal.

        Returns:
            ``MRP × margin%``, or ``None`` if MRP is missing.
        """
        if mrp is None or pd.isna(mrp):
            return None
        return float(mrp) * margin_pct
