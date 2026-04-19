"""
data.mapping_loader
===================

Loads the Ship-To B2B mapping registry — the master list of marketplace
delivery locations and the ERP codes (Sell-to Customer No. + Ship-to
Code) they map to.

The mapping file has a sheet named ``Ship-To B2B`` with columns::

    Party | Del Location | Cust No | Ship to

The loader is filtered per-marketplace at load time: when the user
selects "Myntra", only rows with ``Party == 'Myntra'`` are kept. This
means the lookup table stays small and a wrong-party match becomes
impossible.

Lookup strategy is three-tier (exact → case-insensitive → fuzzy
substring) so we tolerate small variations in how marketplaces spell
location names. Every successful match returns the canonical mapping key
in ``matched_key`` so the GUI can show the original raw value alongside
what we matched it to (Summary sheet's "Location (Raw)" vs "Location
(Mapped)" columns).
"""

from __future__ import annotations
import logging
from typing import Dict, List, Optional, Tuple

import pandas as pd


class MappingLoader:
    """
    Per-marketplace location → (Cust No, Ship-to) lookup table.

    Loaded once per processing run. The ``load()`` call also accepts a
    ``logs`` accumulator so any column-detection or read errors surface
    in the GUI's Warnings sheet, not just stderr.
    """

    def __init__(self) -> None:
        # location string → {cust_no, ship_to}
        self.mappings: Dict[str, Dict[str, str]] = {}
        self.party_name: str = ''
        self.total_loaded: int = 0

    # ── Loading ────────────────────────────────────────────────────────

    def load(self, filepath: str, party_name: str,
             logs: List[Tuple[str, str, str]]) -> int:
        """
        Read the mapping file and build the per-marketplace lookup.

        Args:
            filepath: Path to the mapping Excel file (e.g.
                      ``Calculation Data/Ship to B2B.xlsx``).
            party_name: Marketplace name to filter by — must match the
                        sheet's ``Party`` column case-insensitively.
            logs: Mutable list. Tuples ``(po, location, message)`` are
                  appended on errors. PO and location are empty strings
                  for global errors (e.g. missing sheet).

        Returns:
            Number of locations loaded for ``party_name``. Zero means
            either the file couldn't be read or the marketplace had no
            entries.
        """
        self.party_name = party_name
        self.mappings = {}

        # Try the canonical sheet first; fall back to the first sheet if
        # someone renamed or split the workbook. Cannot-read errors are
        # logged and we return 0 — the caller will see "no locations
        # loaded" and surface a clear error.
        try:
            df = pd.read_excel(filepath, sheet_name='Ship-To B2B', header=0)
        except ValueError:
            logging.warning("Sheet 'Ship-To B2B' not found, trying first sheet")
            df = pd.read_excel(filepath, header=0)
        except Exception as e:
            logs.append(('', '', f"Cannot read mapping file: {e}"))
            return 0

        # ── Column detection (lenient on naming) ────────────────────────
        # Mapping files vary slightly in header capitalisation and exact
        # phrasing across versions, so we accept a small set of synonyms
        # for each canonical column.
        col_map: Dict[str, str] = {}
        for col in df.columns:
            cl = str(col).strip().lower()
            if cl == 'party':
                col_map['party'] = col
            elif cl in ('del location', 'delivery location', 'location'):
                col_map['location'] = col
            elif cl in ('cust no', 'cust no.', 'customer no', 'sell-to'):
                col_map['cust_no'] = col
            elif cl in ('ship to', 'ship-to', 'ship to code'):
                col_map['ship_to'] = col

        missing = [k for k in ('party', 'location', 'cust_no', 'ship_to')
                   if k not in col_map]
        if missing:
            logs.append(('', '',
                         f"Mapping file missing columns: {', '.join(missing)}. "
                         f"Available: {list(df.columns)}"))
            return 0

        # ── Filter by party + build lookup ──────────────────────────────
        for _, row in df.iterrows():
            party = str(row[col_map['party']]).strip()
            if party.lower() != party_name.lower():
                continue

            location = str(row[col_map['location']]).strip()
            cust_no = (str(row[col_map['cust_no']]).strip()
                       if pd.notna(row[col_map['cust_no']]) else '')
            ship_to = (str(row[col_map['ship_to']]).strip()
                       if pd.notna(row[col_map['ship_to']]) else '')

            # Customer numbers are integers in the ERP but pandas reads
            # them as floats when any cell is empty — strip the trailing
            # '.0' so '20011.0' becomes '20011'.
            if cust_no.endswith('.0'):
                cust_no = cust_no[:-2]

            # Skip rows where location is empty / "nan" (unmapped entries)
            if location and location.lower() != 'nan':
                self.mappings[location] = {
                    'cust_no': cust_no,
                    'ship_to': ship_to,
                }

        self.total_loaded = len(self.mappings)
        logging.info("Mapping: Loaded %d locations for '%s'",
                     self.total_loaded, party_name)
        return self.total_loaded

    # ── Lookup ─────────────────────────────────────────────────────────

    def lookup(self, location: str) -> Optional[Dict[str, str]]:
        """
        Find the ERP codes for a delivery location string.

        Three-tier match::

            1. Exact            (preferred — no ambiguity)
            2. Case-insensitive (handles "Bilaspur" vs "bilaspur")
            3. Substring        (handles "Bilaspur Warehouse - Gurgaon"
                                 vs canonical "Bilaspur")

        On a successful match the returned dict includes ``matched_key``
        — the canonical mapping key actually used. The GUI's Summary
        sheet renders this alongside the raw input value so the user can
        visually verify fuzzy matches.

        Args:
            location: Raw delivery location from the punch file.

        Returns:
            ``{cust_no, ship_to, matched_key}`` on hit, ``None`` on miss.
        """
        if not location:
            return None

        loc_clean = location.strip()

        # 1. Exact match
        if loc_clean in self.mappings:
            return {**self.mappings[loc_clean], 'matched_key': loc_clean}

        # 2. Case-insensitive match
        loc_lower = loc_clean.lower()
        for key, val in self.mappings.items():
            if key.lower() == loc_lower:
                return {**val, 'matched_key': key}

        # 3. Substring match (lossy — log it so a misuse is visible)
        for key, val in self.mappings.items():
            key_lower = key.lower()
            if loc_lower in key_lower or key_lower in loc_lower:
                logging.info("Mapping: Fuzzy match '%s' → '%s'",
                             loc_clean, key)
                return {**val, 'matched_key': key}

        return None
