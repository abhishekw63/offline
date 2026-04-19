"""
engine.marketplace_engine
=========================

Turns a marketplace punch file into a list of ``SORow`` rows ready for
the exporter. This is the heart of the pipeline:

#. Read the punch file (Excel).
#. Validate the columns we need exist (depends on ``item_resolution``).
#. For each row:
   * Parse identity (PO, location, qty).
   * Resolve EAN, then Item No (either from ``item_col`` or by EAN→master
     lookup).
   * Look up MRP / GST / Description from the master.
   * Compute our calculated price (``calc_price``) per the marketplace's
     ``compare_basis`` (landing or cost).
   * Compute the diff against the marketplace's quoted price; flag
     mismatches.
   * Look up the delivery location in the mapping registry.
#. Return a ``ProcessingResult``.

The engine never raises on per-row data problems — it appends to
``result.warnings`` so the GUI can surface them. The only fatal failures
return early after a warning: missing required columns, missing master
when EAN-resolution requires it, etc.
"""

from __future__ import annotations
import logging
import os
from typing import Any, Dict, Optional, Set, Tuple

import pandas as pd

from online_po_processor.data.models import SORow, ProcessingResult
from online_po_processor.data.master_loader import MasterLoader
from online_po_processor.data.mapping_loader import MappingLoader


# Set of GST codes we know how to handle. Anything outside this set
# triggers a warning and falls back to 18% in MasterLoader.
_KNOWN_GST_CODES = frozenset({
    '0-G',
    'G-3', 'G-3-S',
    'G-5', 'G-5-S',
    'G-12', 'G-12-S',
    'G-18', 'G-18-S',
    '',
})


class MarketplaceEngine:
    """
    Apply per-marketplace config rules to a punch file.

    Args:
        mapping: Loaded ``MappingLoader`` for the selected marketplace.
        master:  Loaded ``MasterLoader``, or ``None``. Required when the
                 marketplace's ``item_resolution`` is ``'from_ean'`` —
                 otherwise the engine has no way to derive Item No. When
                 ``None`` and resolution is ``'from_column'``, price
                 validation is silently disabled (rows still pass through).
    """

    # Threshold for flagging price mismatches (rupees). Diffs at or below
    # this are treated as rounding noise → status='OK'. Above this →
    # 'MISMATCH' with a warning row.
    DIFFN_THRESHOLD: float = 1.0

    def __init__(self, mapping: MappingLoader,
                 master: Optional[MasterLoader] = None) -> None:
        self.mapping = mapping
        self.master = master

    # ── Public entry point ─────────────────────────────────────────────

    def process(self, filepath: str, config: Dict[str, Any],
                margin_pct: float = 0.70) -> ProcessingResult:
        """
        Read ``filepath`` and produce a ``ProcessingResult``.

        Args:
            filepath:   Path to the punch/PO Excel file.
            config:     Entry from
                        :data:`~online_po_processor.config.marketplaces.MARKETPLACE_CONFIGS`.
            margin_pct: Margin as decimal (e.g. ``0.70`` for 70%).

        Returns:
            Always returns a result, even if the read failed or all rows
            were skipped. Inspect ``result.warnings`` and
            ``len(result.rows)`` to detect problems.
        """
        result = ProcessingResult(
            marketplace=config['party_name'],
            input_file=os.path.basename(filepath),
            input_file_path=str(filepath),  # used for output/ folder location
            compare_basis=config.get('compare_basis', 'cost'),
            compare_label=config.get('compare_label', 'Price'),
            margin_pct=margin_pct,
        )

        # ── Read file ───────────────────────────────────────────────────
        # v1.4.2: All marketplace punch files carry their data on 'Sheet1'.
        # Other sheets in the workbook are user-side pivots / manual calc
        # / sidecars that must NOT be read by the script. Previously we
        # defaulted to pandas' "first sheet" behavior which silently
        # latched onto Sheet2/Sheet4 when a pivot sheet came first —
        # that's now fixed across the board.
        #
        # If 'Sheet1' doesn't exist (highly unusual), we fall back to the
        # first sheet and log a warning so the user can spot the issue
        # rather than getting a cryptic KeyError downstream.
        try:
            available_sheets = pd.ExcelFile(filepath).sheet_names
        except Exception as e:  # noqa: BLE001
            result.warnings.append((
                '', '', f"Cannot open file: {e}"
            ))
            return result

        if 'Sheet1' in available_sheets:
            sheet_to_read = 'Sheet1'
        else:
            sheet_to_read = available_sheets[0]
            result.warnings.append((
                '', '',
                f"'Sheet1' not found in file — falling back to "
                f"'{sheet_to_read}'. Available sheets: {available_sheets}"
            ))
            logging.warning("'Sheet1' missing; reading '%s' instead",
                             sheet_to_read)

        try:
            df = pd.read_excel(filepath, sheet_name=sheet_to_read, header=0)
        except Exception as e:  # noqa: BLE001 — we want to surface ANY read error
            result.warnings.append((
                '', '',
                f"Cannot read sheet {sheet_to_read!r}: {e}"
            ))
            return result

        logging.info("Read %d rows from %s",
                     len(df), os.path.basename(filepath))
        result.raw_df = df

        # ── Required-column validation ──────────────────────────────────
        if not self._validate_required_columns(df, config, result):
            return result

        # Optional columns — log misses but keep going
        price_col = self._validate_optional_column(df, config, 'price_col')
        self._validate_optional_column(df, config, 'fob_col',
                                        log_warn_to_result=result)
        self._validate_optional_column(df, config, 'ean_col',
                                        log_warn_to_result=result)

        # ── Per-row processing ──────────────────────────────────────────
        warned_keys: Set[Tuple] = set()  # dedupe warnings (e.g. one per PO)
        item_resolution = config.get('item_resolution', 'from_column')
        compare_basis = config.get('compare_basis', 'cost')
        compare_label = config.get('compare_label', 'Price')

        for _, row in df.iterrows():
            so_row = self._process_row(
                row=row,
                df=df,
                config=config,
                item_resolution=item_resolution,
                compare_basis=compare_basis,
                compare_label=compare_label,
                margin_pct=margin_pct,
                price_col=price_col,
                result=result,
                warned_keys=warned_keys,
            )
            if so_row is not None:
                result.rows.append(so_row)

        logging.info("Processed %d items across %d PO(s)",
                     len(result.rows),
                     len({r.po_number for r in result.rows}))
        return result

    # ── Column validation helpers ──────────────────────────────────────

    def _validate_required_columns(self, df: pd.DataFrame,
                                    config: Dict[str, Any],
                                    result: ProcessingResult) -> bool:
        """
        Confirm the required columns exist for this marketplace.

        Required set depends on ``item_resolution``:
          * ``from_column`` → po, loc, item, qty
          * ``from_ean``    → po, loc, ean, qty   (item_col may be absent)

        On failure, appends a warning to ``result`` and returns False.
        Caller should ``return result`` immediately.
        """
        item_resolution = config.get('item_resolution', 'from_column')

        required_cols: Dict[str, str] = {
            'po': config['po_col'],
            'loc': config['loc_col'],
            'qty': config['qty_col'],
        }

        if item_resolution == 'from_ean':
            ean_required = config.get('ean_col')
            if not ean_required:
                result.warnings.append((
                    '', '',
                    "Config error: item_resolution='from_ean' requires ean_col."))
                return False
            required_cols['ean'] = ean_required
        else:  # 'from_column' (default)
            item_required = config.get('item_col')
            if not item_required:
                result.warnings.append((
                    '', '',
                    "Config error: item_resolution='from_column' requires item_col."))
                return False
            required_cols['item'] = item_required

        for _key, col_name in required_cols.items():
            if col_name not in df.columns:
                result.warnings.append((
                    '', '',
                    f"Required column '{col_name}' not found. "
                    f"Available: {list(df.columns)[:15]}..."))
                return False

        return True

    @staticmethod
    def _validate_optional_column(
        df: pd.DataFrame,
        config: Dict[str, Any],
        config_key: str,
        log_warn_to_result: Optional[ProcessingResult] = None,
    ) -> Optional[str]:
        """
        Check if an optional column exists. Returns its name if present,
        ``None`` otherwise. If absent and ``log_warn_to_result`` is given,
        appends a warning row.
        """
        col_name = config.get(config_key)
        if col_name and col_name in df.columns:
            return col_name

        if col_name:
            # Configured but missing — log it
            logging.warning("%s column '%s' not found — skipping",
                            config_key, col_name)
            if log_warn_to_result is not None:
                log_warn_to_result.warnings.append((
                    '', '',
                    f"Column '{col_name}' (config key '{config_key}') not "
                    f"found in file — that feature will be skipped. "
                    f"Available: {list(df.columns)[:10]}..."))
        return None

    # ── Per-row processing ─────────────────────────────────────────────

    def _process_row(
        self,
        row: pd.Series,
        df: pd.DataFrame,
        config: Dict[str, Any],
        item_resolution: str,
        compare_basis: str,
        compare_label: str,
        margin_pct: float,
        price_col: Optional[str],
        result: ProcessingResult,
        warned_keys: Set[Tuple],
    ) -> Optional[SORow]:
        """
        Build a single SORow from one DataFrame row, or return None if
        the row should be skipped.

        Skips on: missing PO, qty ≤ 0, missing item value (mode-dependent).

        Side-effect: appends to ``result.warnings`` for unmappable
        locations, GST code surprises, price mismatches, etc.
        """
        # ── Identity: PO, location, qty ─────────────────────────────────
        po = str(row[config['po_col']]).strip()
        location = (str(row[config['loc_col']]).strip()
                    if pd.notna(row[config['loc_col']]) else '')
        qty_raw = row[config['qty_col']]

        # Skip rows with no PO number
        if po.lower() == 'nan':
            return None

        # Parse quantity early — a zero-qty row contributes nothing,
        # avoid running master/mapping lookups for it.
        try:
            qty = int(float(qty_raw)) if pd.notna(qty_raw) else 0
        except (ValueError, TypeError):
            qty = 0
        if qty <= 0:
            return None

        # ── Extract EAN (needed before item resolution for from_ean) ────
        ean = self._extract_ean(row, df, config)

        # ── Resolve Item No per the marketplace's strategy ──────────────
        item_no = self._resolve_item_no(
            row=row, ean=ean, po=po, config=config,
            item_resolution=item_resolution,
            warned_keys=warned_keys, result=result,
        )
        if item_no is None:
            return None  # already warned inside the helper

        # ── Pull pass-through unit price (rare, both current MPs leave
        # this None so the WMS computes it downstream) ──────────────────
        unit_price = self._extract_float(row, price_col) if price_col else None

        # ── Marketplace prices: active (validation) + optional reference ─
        fob_price = self._extract_float(row, config.get('fob_col'),
                                         only_if_in_df=df)
        ref_fob_price = self._extract_float(row, config.get('ref_fob_col'),
                                             only_if_in_df=df)

        # ── Master lookup + price validation ────────────────────────────
        (mrp, gst_code, description, cost_price_ref, calc_price,
         diffn, ref_diffn, validation_status) = self._validate_against_master(
            ean=ean,
            item_no=item_no,
            po=po,
            margin_pct=margin_pct,
            compare_basis=compare_basis,
            compare_label=compare_label,
            fob_price=fob_price,
            ref_fob_price=ref_fob_price,
            warned_keys=warned_keys,
            result=result,
        )

        # ── Mapping lookup ──────────────────────────────────────────────
        cust_no, ship_to, mapped, mapped_location = self._resolve_mapping(
            location=location, po=po, party_name=config['party_name'],
            warned_keys=warned_keys, result=result,
        )

        return SORow(
            po_number=po,
            location=location,
            item_no=item_no,
            qty=qty,
            unit_price=unit_price,
            cust_no=cust_no,
            ship_to=ship_to,
            mapped=mapped,
            mapped_location=mapped_location,
            ean=ean,
            description=description,
            fob_price=fob_price,
            ref_fob_price=ref_fob_price,
            calc_price=calc_price,
            cost_price_ref=cost_price_ref,
            diffn=diffn,
            ref_diffn=ref_diffn,
            mrp=mrp,
            gst_code=gst_code,
            validation_status=validation_status,
        )

    # ── Row-level extraction helpers ───────────────────────────────────

    @staticmethod
    def _extract_ean(row: pd.Series, df: pd.DataFrame,
                     config: Dict[str, Any]) -> str:
        """
        Read the EAN cell as a clean string, handling the float64 case
        where pandas reads ``8906121642599`` as ``8906121642599.0`` and
        we'd otherwise pass that ``.0`` into a master lookup.
        """
        ean_col = config.get('ean_col')
        if not (ean_col and ean_col in df.columns):
            return ''

        ean_raw = row[ean_col]
        if not pd.notna(ean_raw):
            return ''

        # Numeric EAN — coerce through int() to drop the trailing .0,
        # then str() for the lookup key.
        if isinstance(ean_raw, (int, float)):
            try:
                return str(int(ean_raw))
            except (ValueError, OverflowError):
                return str(ean_raw).strip()
        return str(ean_raw).strip()

    @staticmethod
    def _extract_float(row: pd.Series, col: Optional[str],
                       only_if_in_df: Optional[pd.DataFrame] = None,
                       ) -> Optional[float]:
        """
        Read ``row[col]`` as ``float | None``. Defensive against missing
        column, NaN, or non-numeric strings.
        """
        if not col:
            return None
        if only_if_in_df is not None and col not in only_if_in_df.columns:
            return None
        try:
            v = row[col]
        except KeyError:
            return None
        if not pd.notna(v):
            return None
        try:
            return float(v)
        except (ValueError, TypeError):
            return None

    def _resolve_item_no(self, row: pd.Series, ean: str, po: str,
                          config: Dict[str, Any], item_resolution: str,
                          warned_keys: Set[Tuple],
                          result: ProcessingResult) -> Any:
        """
        Resolve the canonical Item No based on ``item_resolution``.

        ``from_column`` path: read from ``item_col``. NaN → skip row.
        ``from_ean`` path: look the EAN up in the master and use
        ``master_info['item_no']``. Empty EAN → skip with warning.
        EAN not in master → emit row with ``item_no = ean`` so it still
        appears (in NOT_IN_MASTER state in the validation sheet).

        Returns ``None`` if the row should be skipped (warnings already
        appended where appropriate).
        """
        if item_resolution == 'from_ean':
            if not ean:
                key = ('NO_EAN', po)
                if key not in warned_keys:
                    warned_keys.add(key)
                    result.warnings.append((
                        po, '',
                        f"Row skipped: ean_col '{config.get('ean_col')}' is "
                        f"empty for PO {po}. item_resolution='from_ean' "
                        f"requires a non-empty EAN."
                    ))
                return None

            if not self.master:
                key = ('NO_MASTER', 'global')
                if key not in warned_keys:
                    warned_keys.add(key)
                    result.warnings.append((
                        '', '',
                        "Cannot resolve Item No: item_resolution='from_ean' "
                        "requires the Items_March master to be loaded."
                    ))
                return None

            master_info = self.master.lookup(ean)
            if not master_info:
                # Surface the row anyway with EAN as the visible item value
                return ean

            resolved = master_info.get('item_no', '')
            try:
                return int(resolved)
            except (ValueError, TypeError):
                return str(resolved).strip()

        # 'from_column' (default)
        item_raw = row[config['item_col']]
        if pd.isna(item_raw):
            return None
        try:
            return int(item_raw)
        except (ValueError, TypeError):
            return str(item_raw).strip()

    def _validate_against_master(
        self, ean: str, item_no: Any, po: str, margin_pct: float,
        compare_basis: str, compare_label: str,
        fob_price: Optional[float], ref_fob_price: Optional[float],
        warned_keys: Set[Tuple], result: ProcessingResult,
    ) -> Tuple[Optional[float], str, str,
               Optional[float], Optional[float],
               Optional[float], Optional[float], str]:
        """
        Run the master lookup and compute calc/cost/diffs/status.

        Returns (in order)::

            mrp, gst_code, description,
            cost_price_ref, calc_price,
            diffn, ref_diffn, validation_status
        """
        mrp: Optional[float] = None
        gst_code: str = ''
        description: str = ''
        cost_price_ref: Optional[float] = None
        calc_price: Optional[float] = None
        diffn: Optional[float] = None
        ref_diffn: Optional[float] = None
        validation_status: str = ''

        if not self.master:
            return (mrp, gst_code, description, cost_price_ref,
                    calc_price, diffn, ref_diffn, validation_status)

        # Try EAN first, then fall back to Item No.
        master_info = self.master.lookup(ean) if ean else None
        if not master_info:
            master_info = self.master.lookup(str(item_no))

        if not master_info:
            return (mrp, gst_code, description, cost_price_ref,
                    calc_price, diffn, ref_diffn, 'NOT_IN_MASTER')

        mrp = master_info['mrp']
        gst_code = master_info['gst_code']
        description = master_info.get('description', '')

        # Warn on unknown GST code (still computes, defaulting to 18%)
        gst_upper = str(gst_code).strip().upper()
        if gst_upper not in _KNOWN_GST_CODES and gst_upper != 'NAN':
            key = ('GST', gst_upper)
            if key not in warned_keys:
                warned_keys.add(key)
                result.warnings.append((
                    po, str(item_no),
                    f"Unknown GST code '{gst_code}' for Item {item_no} — "
                    f"defaulting to 18%. Please verify in Items_March."
                ))
                logging.warning("Unknown GST code '%s' for Item %s",
                                gst_code, item_no)

        # Always compute the post-GST cost price for the "naked CP"
        # column shown in the Validation sheet.
        cost_price_ref = MasterLoader.calc_cost_price(mrp, gst_code, margin_pct)

        # Reference diff (vs naked CP) — display-only, always vs post-GST
        # because the reference column itself is post-GST (e.g. Myntra's
        # List price).
        if cost_price_ref is not None and ref_fob_price is not None:
            ref_diffn = ref_fob_price - cost_price_ref

        # Pick what we ACTUALLY compare against, based on basis.
        if compare_basis == 'landing':
            calc_price = MasterLoader.calc_landing_price(mrp, margin_pct)
        else:  # 'cost' (default)
            calc_price = cost_price_ref

        # Compute active diff + status
        if calc_price is not None and fob_price is not None:
            diffn = fob_price - calc_price
            if abs(diffn) <= self.DIFFN_THRESHOLD:
                validation_status = 'OK'
            else:
                validation_status = 'MISMATCH'
                key = ('VALIDATION', str(item_no))
                if key not in warned_keys:
                    warned_keys.add(key)
                    result.warnings.append((
                        po, str(item_no),
                        f"{compare_label} mismatch: Item {item_no}, "
                        f"Marketplace={fob_price:.2f}, "
                        f"Calculated={calc_price:.2f}, "
                        f"Diff={diffn:.2f}"
                    ))
        else:
            validation_status = 'NO_PRICE'

        return (mrp, gst_code, description, cost_price_ref,
                calc_price, diffn, ref_diffn, validation_status)

    def _resolve_mapping(self, location: str, po: str, party_name: str,
                          warned_keys: Set[Tuple],
                          result: ProcessingResult,
                          ) -> Tuple[str, str, bool, str]:
        """
        Look up the location in the mapping registry.

        Returns (cust_no, ship_to, mapped_bool, mapped_location_str).
        On miss, appends a warning (deduped per (po, location)) and
        returns blanks plus mapped=False.
        """
        mapping_result = self.mapping.lookup(location)

        if mapping_result:
            return (
                mapping_result['cust_no'],
                mapping_result['ship_to'],
                True,
                mapping_result.get('matched_key', location),
            )

        # Unmapped — warn once per (po, location)
        key = (po, location)
        if key not in warned_keys:
            warned_keys.add(key)
            result.warnings.append((
                po, location,
                f"Location '{location}' not found in mapping for {party_name}. "
                f"Cust No and Ship-to left empty."
            ))
        return ('', '', False, '')