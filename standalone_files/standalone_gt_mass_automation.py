"""
╔═══════════════════════════════════════════════════════════════════════════════╗
║               GT MASS DUMP GENERATOR — v2.0                                  ║
║               Tkinter GUI Desktop Application                                ║
╠═══════════════════════════════════════════════════════════════════════════════╣
║  Author  : Agami AI / Vishal                                                ║
║  Version : 2.0                                                               ║
║  Purpose : Reads SOGTM Excel files (GT-Mass Purchase Orders), extracts      ║
║            meta info (Distributor, City, State, Location) and ordered items, ║
║            generates ERP-importable Headers (SO) + Lines (SO) sheets.        ║
║  Stack   : Python 3.13, Tkinter, pandas, openpyxl                           ║
╚═══════════════════════════════════════════════════════════════════════════════╝

═══════════════════════════════════════════════════════════════════════════════
  ARCHITECTURE
═══════════════════════════════════════════════════════════════════════════════

  ┌─────────────────────────────────────────────────────────┐
  │                   AutomationUI (GUI)                    │
  │  Select Files → Generate Dump → Open Output             │
  └────────────────────┬────────────────────────────────────┘
                       │
                       ▼
  ┌─────────────────────────────────────────────────────────┐
  │              GTMassAutomation (Engine)                   │
  │  For each SOGTM file:                                    │
  │    1. FileReader.read()      → raw DataFrame             │
  │    2. MetaExtractor.extract() → Distributor/City/State/ │
  │                                  Location                │
  │    3. ExcelParser.parse()    → OrderRow list             │
  │    4. SOFormatter.from_filename() → SO/GTM/6325          │
  └────────────────────┬────────────────────────────────────┘
                       │
                       ▼
  ┌─────────────────────────────────────────────────────────┐
  │              DumpExporter (Output)                       │
  │  Sheet 1: Headers (SO)  → ERP SO header import          │
  │  Sheet 2: Lines (SO)    → ERP SO line import            │
  │  Sheet 3: Sales Lines   → SO Number | Item No | Qty     │
  │  Sheet 4: Sales Header  → Grouped by SO with meta       │
  │  Sheet 5: Warnings      → Any issues found              │
  └─────────────────────────────────────────────────────────┘

═══════════════════════════════════════════════════════════════════════════════
  INPUT FILES — SOGTM####.xlsx
═══════════════════════════════════════════════════════════════════════════════

  Each file has a fixed meta header region (rows 1-6) followed by data:

  Row 0: Title (ignored)
  Row 1: "Distributor Name" in col A, value in col B
  Row 2: "DB Code" in col A, value in col B (may be blank)
  Row 3: "BDE Name" in col A, value in col B
  Row 4: "City" in col A, value in col B
  Row 5: "State" in col A, value in col B | "Location" in col G, value in col I
  Row 6: Header row (EAN, BC Code, Category, ..., Order Qty, Tester Qty)
  Row 7+: Data rows

  ⚠ Location label is at col G ("Location"), value at col I (e.g., "AHD")
  ⚠ Not all files have Location filled — empty = leave Location Code blank
  ⚠ BC Code column = Item No for ERP import

═══════════════════════════════════════════════════════════════════════════════
  LOCATION CODE MAPPING
═══════════════════════════════════════════════════════════════════════════════

  The Location value from the file is mapped to an ERP Location Code:

  ┌──────────────┬────────────────┐
  │ File Value   │ Location Code  │
  ├──────────────┼────────────────┤
  │ AHD          │ PICK           │
  │ (anything    │ (raw value     │
  │   else)      │  as-is)        │
  │ (empty)      │ (empty)        │
  └──────────────┴────────────────┘

  To add new mappings, update LOCATION_CODE_MAP dict in this script.

═══════════════════════════════════════════════════════════════════════════════
  OUTPUT — 5 EXCEL SHEETS
═══════════════════════════════════════════════════════════════════════════════

  Sheet 1: 'Headers (SO)' — One row per SO number (ERP import)
      Document Type | No. | Sell-to Customer No. | Ship-to Code |
      5 × date fields | External Document No. | Location Code |
      Dimension Set ID | Supply Type | ...dimension columns

  Sheet 2: 'Lines (SO)' — One row per ordered item (ERP import)
      Document Type | Document No. | Line No. | Type |
      No. | Location Code | Quantity | Unit Price
      Line No. = 10000, 20000, 30000... resets per SO

  Sheet 3: 'Sales Lines' — Simple flat list
      SO Number | Item No | Qty

  Sheet 4: 'Sales Header' — Grouped summary with meta
      SO Number | Qty | Distributor | City | State | Location

  Sheet 5: 'Warnings' — Any parsing issues (only if warnings exist)
      File | Warning

═══════════════════════════════════════════════════════════════════════════════
  EXPIRY SYSTEM
═══════════════════════════════════════════════════════════════════════════════

  The script has a built-in expiry date (EXPIRY_DATE constant).
  - After expiry: shows error popup and exits
  - Within 7 days of expiry: shows warning popup
  - To extend: change EXPIRY_DATE value

═══════════════════════════════════════════════════════════════════════════════
  DEPENDENCIES & RUNNING
═══════════════════════════════════════════════════════════════════════════════

  Requirements:
      pip install pandas openpyxl

  For .xls files (legacy format):
      pip install xlrd

  Run:
      python gt_mass_dump.py
"""

# ═══════════════════════════════════════════════════════════════════════════════
#  IMPORTS
# ═══════════════════════════════════════════════════════════════════════════════

from __future__ import annotations      # Enable forward type references
import os                               # File/path operations
import sys                              # System exit
import platform                         # OS detection for file opener
import time                             # Timing processing duration
import logging                          # Structured logging
import re                               # Regex for SO number extraction
import pandas as pd                     # Excel reading (fast, handles dtypes)
import tkinter as tk                    # GUI framework
from tkinter import filedialog, messagebox  # File dialogs, alert popups
from dataclasses import dataclass, field    # Structured data containers
from pathlib import Path                # Cross-platform path handling
from typing import List, Optional, Tuple, Dict  # Type hints
from datetime import datetime           # Date operations (expiry, timestamps)

from openpyxl import Workbook           # Excel writing for output
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side  # Cell formatting
from openpyxl.utils import get_column_letter  # Column index → letter


# ═══════════════════════════════════════════════════════════════════════════════
#  LOGGING CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)


# ═══════════════════════════════════════════════════════════════════════════════
#  EXPIRY CHECK
# ═══════════════════════════════════════════════════════════════════════════════
# Change this date to extend the application's validity period.
# Format: DD-MM-YYYY

EXPIRY_DATE = "30-04-2026"

def check_expiry():
    """Check if the application has expired. Shows popup and exits if so."""
    expiry = datetime.strptime(EXPIRY_DATE, "%d-%m-%Y").date()
    today = datetime.now().date()

    if today > expiry:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Application Expired",
            f"This application expired on {EXPIRY_DATE}.\n\n"
            f"Please contact the administrator for an updated version."
        )
        root.destroy()
        sys.exit(0)

    days_remaining = (expiry - today).days
    if days_remaining <= 7:
        root = tk.Tk()
        root.withdraw()
        messagebox.showwarning(
            "Expiration Warning",
            f"⚠️ This application will expire in {days_remaining} day(s).\n\n"
            f"Expiry Date: {EXPIRY_DATE}\n\n"
            f"Please contact the administrator for renewal."
        )
        root.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════

# ┌─────────────────────────────────────────────────────────────────────────┐
# │ LOCATION CODE MAPPING                                                   │
# │                                                                         │
# │ Maps the Location value from SOGTM files to ERP Location Code.         │
# │ If a location is not in this map, its raw value is used as-is.         │
# │ If location is empty/blank, Location Code is left empty.               │
# │                                                                         │
# │ To add a new mapping:                                                   │
# │   'LOCATION_VALUE_FROM_FILE': 'ERP_LOCATION_CODE',                     │
# └─────────────────────────────────────────────────────────────────────────┘
LOCATION_CODE_MAP: Dict[str, str] = {
    'AHD': 'PICK',
    'BLR': 'DS_BL_OFF1',
    # Add more mappings here as discovered:
    # 'DEL': 'DS_DL_OFF1',
}

# ┌─────────────────────────────────────────────────────────────────────────┐
# │ STATE / ZONE VALUES                                                     │
# │                                                                         │
# │ These should never appear as a Distributor name.                        │
# │ If they do, it means rows were swapped in the source file.             │
# └─────────────────────────────────────────────────────────────────────────┘
STATE_LIKE_VALUES = {
    "up", "mp", "ap", "hp", "uk", "jk", "wb", "tn", "kl", "ka",
    "gj", "rj", "hr", "pb", "br", "od", "as", "mh", "cg", "jh",
    "north", "south", "east", "west", "central",
    "uttar pradesh", "madhya pradesh", "rajasthan", "punjab",
    "maharashtra", "gujarat", "karnataka", "tamil nadu",
    "haryana", "delhi", "u.p", "u.p.", "m.p", "m.p."
}


# ═══════════════════════════════════════════════════════════════════════════════
#  DATA MODEL
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class OrderRow:
    """Single ordered item extracted from a SOGTM file."""
    so_number: str       # e.g., 'SO/GTM/6325'
    item_no: str         # BC Code from file (e.g., '200163')
    qty: int             # Order Qty (regular orders only)
    tester_qty: int      # Tester Qty (testers only)
    distributor: str     # Distributor Name from meta
    city: str            # City from meta
    state: str           # State from meta
    location: str        # Raw location value from meta (e.g., 'AHD')
    location_code: str   # Mapped ERP location code (e.g., 'PICK')


@dataclass
class ProcessingResult:
    """Aggregated result from processing all selected files."""
    rows: List[OrderRow] = field(default_factory=list)
    failed_files: List[Tuple[str, str]] = field(default_factory=list)    # (filename, reason)
    warned_files: List[Tuple[str, str]] = field(default_factory=list)    # (filename, warning)
    output_path: Optional[Path] = None


# ═══════════════════════════════════════════════════════════════════════════════
#  SO FORMATTER
# ═══════════════════════════════════════════════════════════════════════════════

class SOFormatter:
    """Extracts SO number from filename using pattern: SOGTM####.xlsx → SO/GTM/####"""

    @staticmethod
    def from_filename(filepath: Path) -> Optional[str]:
        match = re.search(r"\d+", filepath.stem)
        if not match:
            logging.warning(f"SO number not found in {filepath}")
            return None
        return f"SO/GTM/{match.group()}"


# ═══════════════════════════════════════════════════════════════════════════════
#  FILE READER
# ═══════════════════════════════════════════════════════════════════════════════
# Reading strategy by extension:
#   .xlsx / .xlsm → openpyxl (built-in with pandas)
#   .xls          → xlrd (requires: pip install xlrd)

class FileReader:
    """Reads Excel files into raw DataFrames (no header)."""

    @staticmethod
    def read(file_path: Path) -> pd.DataFrame:
        """
        Returns a raw DataFrame (no header) for the first sheet.
        Raises RuntimeError with a clear message on failure.
        """
        ext = file_path.suffix.lower()

        # ── .xlsx / .xlsm → openpyxl ──
        if ext in (".xlsx", ".xlsm"):
            try:
                df = pd.read_excel(file_path, header=None, engine="openpyxl")
                logging.info(f"{file_path.name} — read via openpyxl")
                return df
            except Exception as e:
                raise RuntimeError(
                    f"Cannot read '{file_path.name}'.\n"
                    f"The file may be corrupt or password-protected.\n"
                    f"Error: {e}"
                )

        # ── .xls → xlrd ──
        if ext == ".xls":
            try:
                df = pd.read_excel(file_path, header=None, engine="xlrd")
                logging.info(f"{file_path.name} — read via xlrd")
                return df
            except ImportError:
                raise RuntimeError(
                    f"Cannot read '{file_path.name}' — xlrd is not installed.\n\n"
                    f"Fix: open your terminal / command prompt and run:\n"
                    f"    pip install xlrd\n\n"
                    f"Then restart this application and try again."
                )
            except Exception as e:
                raise RuntimeError(
                    f"Cannot read '{file_path.name}'.\n"
                    f"The file may be corrupt or password-protected.\n"
                    f"Error: {e}"
                )

        # ── Unsupported ──
        raise RuntimeError(
            f"Unsupported file format: '{ext}'.\n"
            f"Only .xlsx, .xlsm, and .xls files are supported."
        )


# ═══════════════════════════════════════════════════════════════════════════════
#  META EXTRACTOR
# ═══════════════════════════════════════════════════════════════════════════════
# Scans the meta header rows (above the data table) to extract:
#   - Distributor Name (col A label, col B value)
#   - City (col A label, col B value)
#   - State (col A label, col B value — picks last non-blank if multiple)
#   - Location (col G label "Location", col I value — e.g., "AHD")
#
# Row positions vary slightly across files, so we scan by label matching.

class MetaExtractor:
    """Extracts Distributor, City, State, and Location from meta header rows."""

    @staticmethod
    def extract(raw_df: pd.DataFrame, header_row: int) -> Tuple[dict, List[str]]:
        """
        Scans rows 0..header_row-1 for meta fields.

        Returns:
            (meta_dict, warnings_list)
            meta_dict has keys: distributor, city, state, location, location_code
        """
        distributor = ""
        city = ""
        state_values = []
        location = ""
        warnings = []

        meta_df = raw_df.iloc[:header_row]

        for _, row in meta_df.iterrows():
            # ── Column A/B scanning (Distributor, City, State) ──
            label = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ""
            value = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            if value.lower() in ("nan", ""):
                value = ""

            if label == "distributor name" and not distributor:
                distributor = value
                logging.info(f"Distributor found: '{distributor}'")
            elif label == "city" and not city:
                city = value
                logging.info(f"City found: '{city}'")
            elif label == "state":
                state_values.append(value)

            # ── Column G/I scanning (Location) ──
            # Location label is typically at column index 6 ("Location")
            # Location value is at column index 8 (e.g., "AHD")
            for col_idx in range(min(len(row) - 1, 10)):
                cell_val = str(row.iloc[col_idx]).strip().lower() if pd.notna(row.iloc[col_idx]) else ""
                if cell_val == "location":
                    # Look for value in the next available column(s)
                    for val_idx in range(col_idx + 1, min(col_idx + 3, len(row))):
                        loc_val = row.iloc[val_idx]
                        if pd.notna(loc_val) and str(loc_val).strip() and str(loc_val).strip().lower() != 'nan':
                            location = str(loc_val).strip()
                            logging.info(f"Location found: '{location}'")
                            break

        # Pick last non-blank state (bottom State row is the proper state)
        state = ""
        for s in reversed(state_values):
            if s:
                state = s
                break
        logging.info(f"State found: '{state}'")

        # ── Map location to ERP Location Code ──
        location_code = ""
        if location:
            location_upper = location.upper().strip()
            if location_upper in LOCATION_CODE_MAP:
                location_code = LOCATION_CODE_MAP[location_upper]
                logging.info(f"Location '{location}' → mapped to '{location_code}'")
            else:
                # Use raw value as-is (user will add mapping later)
                location_code = location
                logging.info(f"Location '{location}' → no mapping found, using raw value")

        # ── Validation ──
        if not distributor:
            warnings.append("Distributor Name is blank — label not found or value is empty.")
        if not city:
            warnings.append("City is blank — label not found or value is empty.")
        if not state:
            warnings.append("State is blank — both State rows are empty or missing.")
        if distributor and distributor.strip().lower() in STATE_LIKE_VALUES:
            warnings.append(
                f"Distributor value '{distributor}' looks like a state/zone name. "
                f"Rows may be swapped — please verify manually."
            )

        return {
            "distributor": distributor,
            "city": city,
            "state": state,
            "location": location,
            "location_code": location_code,
        }, warnings


# ═══════════════════════════════════════════════════════════════════════════════
#  EXCEL PARSER
# ═══════════════════════════════════════════════════════════════════════════════
# Reads a single SOGTM file → returns list of OrderRow + warnings.
# Finds the header row by scanning for 'BC Code' + 'Order Qty' columns.

class ExcelParser:
    """Parses a single SOGTM Excel file into OrderRow objects."""

    BC_COLUMN = "bc code"           # Column containing Item No (BC Code)
    QTY_COLUMN = "order qty"        # Column containing order quantity
    TESTER_COLUMN = "tester qty"    # Column containing tester quantity

    def parse(self, file_path: Path) -> Tuple[List[OrderRow], List[str]]:
        """
        Parse one SOGTM file.

        Returns: (rows, warnings)
        Raises RuntimeError if file is unreadable or structure is broken.
        """
        logging.info(f"Reading {file_path.name}")
        warnings = []

        # ── Read raw file ──
        raw_df = FileReader.read(file_path)

        # ── Find header row (contains 'BC Code' and 'Order Qty') ──
        header_row = None
        for i, row in raw_df.iterrows():
            row_values = [str(v).lower() for v in row.values]
            if "bc code" in row_values and any("order qty" in v for v in row_values):
                header_row = i
                break

        if header_row is None:
            raise RuntimeError(
                "Header row not found — could not locate both 'BC Code' and 'Order Qty'. "
                "File format may have changed."
            )

        # ── Extract meta fields (Distributor, City, State, Location) ──
        meta, meta_warnings = MetaExtractor.extract(raw_df, header_row)
        warnings.extend(meta_warnings)

        # ── Build data table from raw DataFrame ──
        df = raw_df.iloc[header_row + 1:].copy()
        df.columns = raw_df.iloc[header_row].values
        df = df.reset_index(drop=True)

        bc_col, qty_col, tester_col = self._detect_columns(df)

        if bc_col is None:
            raise RuntimeError("Column 'BC Code' not found in data table.")
        if qty_col is None:
            raise RuntimeError("Column 'Order Qty' not found in data table.")
        if tester_col is None:
            warnings.append("Column 'Tester Qty' not found — tester quantities will be 0.")

        # ── SO number from filename ──
        so_number = SOFormatter.from_filename(file_path)
        if so_number is None:
            warnings.append(
                "Could not extract SO number from filename. "
                "Filename should contain digits e.g. SOGTM6290.xlsx"
            )
            so_number = "SO/GTM/UNKNOWN"

        # ── Extract ordered rows (order qty > 0 OR tester qty > 0) ──
        rows: List[OrderRow] = []
        for _, row in df.iterrows():
            bc_code = row[bc_col]
            if pd.isna(bc_code):
                continue
            try:
                bc_code = int(bc_code)
            except (ValueError, TypeError):
                continue

            qty = self._clean_qty(row[qty_col])
            tester_qty = self._clean_qty(row[tester_col]) if tester_col is not None else 0

            # Skip rows where both order and tester are zero
            if qty <= 0 and tester_qty <= 0:
                continue

            rows.append(OrderRow(
                so_number=so_number,
                item_no=str(bc_code),
                qty=qty,
                tester_qty=tester_qty,
                distributor=meta["distributor"],
                city=meta["city"],
                state=meta["state"],
                location=meta["location"],
                location_code=meta["location_code"],
            ))

        if not rows:
            warnings.append(
                "No ordered rows found — all Order Qty and Tester Qty values are 0 or blank."
            )

        return rows, warnings

    def _detect_columns(self, df) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        """Find the BC Code, Order Qty, and Tester Qty columns by name matching."""
        bc_col = None
        qty_col = None
        tester_col = None
        for col in df.columns:
            name = str(col).strip().lower()
            if name == self.BC_COLUMN:
                bc_col = col
            if self.QTY_COLUMN in name:
                qty_col = col
            if self.TESTER_COLUMN in name:
                tester_col = col
        return bc_col, qty_col, tester_col

    @staticmethod
    def _clean_qty(value) -> int:
        """Clean and convert a quantity value to int. Returns 0 for invalid."""
        if pd.isna(value):
            return 0
        value = str(value).strip()
        if value in ("", "-"):
            return 0
        value = value.replace(",", "")
        try:
            return int(float(value))
        except ValueError:
            return 0


# ═══════════════════════════════════════════════════════════════════════════════
#  DUMP EXPORTER
# ═══════════════════════════════════════════════════════════════════════════════
# Writes the output Excel with 5 sheets:
#   1. Headers (SO) — ERP SO header import format
#   2. Lines (SO)   — ERP SO line import format
#   3. Sales Lines  — Simple flat list
#   4. Sales Header — Grouped summary
#   5. Warnings     — Parsing issues (only if any)

class DumpExporter:
    """Writes the output Excel file with ERP import sheets."""

    # ── Excel formatting constants ──
    HEADER_FILL = PatternFill('solid', fgColor='1A237E')
    HEADER_FONT = Font(bold=True, color='FFFFFF', name='Aptos Display', size=11)
    DATA_FONT = Font(name='Aptos Display', size=11)
    THIN_SIDE = Side(style='thin', color='CCCCCC')
    BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

    def _hdr_cell(self, ws, row, col, value):
        """Create a formatted header cell."""
        cell = ws.cell(row=row, column=col, value=value)
        cell.font = self.HEADER_FONT
        cell.fill = self.HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = self.BORDER
        return cell

    def _data_cell(self, ws, row, col, value, fmt=None):
        """Create a formatted data cell."""
        cell = ws.cell(row=row, column=col, value=value)
        cell.font = self.DATA_FONT
        cell.border = self.BORDER
        if fmt:
            cell.number_format = fmt
        return cell

    def _auto_width(self, ws, max_w=50):
        """Auto-fit column widths based on content."""
        for col in ws.columns:
            letter = col[0].column_letter
            w = max((len(str(c.value or '')) for c in col), default=8)
            ws.column_dimensions[letter].width = min(w + 3, max_w)

    def export(self, result: ProcessingResult) -> Optional[Path]:
        """
        Write output Excel with all 5 sheets.
        Returns the output path on success, or None if nothing to export.
        """
        # ── Show popup for failed files ──
        if result.failed_files:
            msg = "The following files could NOT be read and were skipped:\n\n"
            for fname, reason in result.failed_files:
                msg += f"  • {fname}\n    Reason: {reason}\n\n"
            msg += "Please fix these files and re-process them."
            messagebox.showerror("Files Failed to Read", msg)

        # ── If no rows, stop ──
        if not result.rows:
            messagebox.showwarning(
                "No Data",
                "No valid rows found across all selected files.\nNothing to export."
            )
            return None

        # ── Prepare output file ──
        output_folder = Path("output")
        output_folder.mkdir(exist_ok=True)
        today = datetime.now().strftime("%d-%m-%Y_%H%M%S")
        file_path = output_folder / f"gt_mass_dump_{today}.xlsx"

        # ── Create workbook ──
        wb = Workbook()
        wb.remove(wb.active)

        self._write_headers_so(wb, result)
        self._write_lines_so(wb, result)
        self._write_sales_lines(wb, result)
        self._write_sales_header(wb, result)
        self._write_warnings(wb, result)

        wb.save(str(file_path))
        return file_path

    def _write_headers_so(self, wb, result: ProcessingResult):
        """
        Sheet 1: 'Headers (SO)' — One row per unique SO number.

        Columns:
            Document Type | No. | Sell-to Customer No. | Ship-to Code |
            Posting Date | Order Date | Document Date |
            Invoice From Date | Invoice To Date |
            External Document No. | Location Code | Dimension Set ID |
            Supply Type | Voucher Narration |
            Brand Code | Channel Code | Catagory | Geography Code
        """
        ws = wb.create_sheet('Headers (SO)')
        headers = [
            'Document Type', 'No.', 'Sell-to Customer No.', 'Ship-to Code',
            'Posting Date', 'Order Date', 'Document Date',
            'Invoice From Date', 'Invoice To Date',
            'External Document No.', 'Location Code', 'Dimension Set ID',
            'Supply Type', 'Voucher Narration',
            'Brand Code (Dimension)', 'Channel Code (Dimension)',
            'Catagory (Dimension)', 'Geography Code (Dimension)'
        ]
        for c, h in enumerate(headers, 1):
            self._hdr_cell(ws, 1, c, h)

        today_str = datetime.now().strftime("%d-%m-%Y")

        # Collect unique SO numbers while preserving order
        seen = set()
        unique_sos = []
        for row in result.rows:
            if row.so_number not in seen:
                seen.add(row.so_number)
                unique_sos.append(row)

        r = 2
        for row in unique_sos:
            self._data_cell(ws, r, 1, 'Order')              # Document Type
            self._data_cell(ws, r, 2, row.so_number)         # No.
            self._data_cell(ws, r, 3, '')                    # Sell-to Customer No. (empty — manual lookup)
            self._data_cell(ws, r, 4, '')                    # Ship-to Code (empty)
            self._data_cell(ws, r, 5, today_str)             # Posting Date
            self._data_cell(ws, r, 6, today_str)             # Order Date
            self._data_cell(ws, r, 7, today_str)             # Document Date
            self._data_cell(ws, r, 8, today_str)             # Invoice From Date
            self._data_cell(ws, r, 9, today_str)             # Invoice To Date
            self._data_cell(ws, r, 10, row.so_number)        # External Document No.
            self._data_cell(ws, r, 11, row.location_code)    # Location Code (mapped)
            self._data_cell(ws, r, 12, '')                   # Dimension Set ID
            self._data_cell(ws, r, 13, 'B2B')                # Supply Type
            # Columns 14-18: empty dimension columns
            r += 1

        self._auto_width(ws)

    def _write_lines_so(self, wb, result: ProcessingResult):
        """
        Sheet 2: 'Lines (SO)' — One row per ordered item.
        Line No. increments by 10000, resets per new SO number.

        Columns:
            Document Type | Document No. | Line No. | Type |
            No. | Location Code | Quantity | Unit Price
        """
        ws = wb.create_sheet('Lines (SO)')
        headers = ['Document Type', 'Document No.', 'Line No.', 'Type',
                   'No.', 'Location Code', 'Quantity', 'Unit Price']
        for c, h in enumerate(headers, 1):
            self._hdr_cell(ws, 1, c, h)

        r = 2
        current_so = None
        line_no = 0

        for row in result.rows:
            # Reset line counter on new SO number
            if row.so_number != current_so:
                current_so = row.so_number
                line_no = 0

            line_no += 10000

            self._data_cell(ws, r, 1, 'Order')              # Document Type
            self._data_cell(ws, r, 2, row.so_number)         # Document No.
            self._data_cell(ws, r, 3, line_no)               # Line No.
            self._data_cell(ws, r, 4, 'Item')                # Type
            self._data_cell(ws, r, 5, row.item_no)           # No. (BC Code)
            self._data_cell(ws, r, 6, row.location_code)     # Location Code (mapped)
            self._data_cell(ws, r, 7, row.qty)               # Quantity (order qty only — ERP import)
            self._data_cell(ws, r, 8, '')                    # Unit Price (empty — ERP fetches)
            r += 1

        self._auto_width(ws)

    def _write_sales_lines(self, wb, result: ProcessingResult):
        """Sheet 3: 'Sales Lines' — Simple flat list of all ordered items."""
        ws = wb.create_sheet('Sales Lines')
        headers = ['SO Number', 'Item No', 'Qty']
        for c, h in enumerate(headers, 1):
            self._hdr_cell(ws, 1, c, h)

        for r, row in enumerate(result.rows, 2):
            self._data_cell(ws, r, 1, row.so_number)
            self._data_cell(ws, r, 2, row.item_no)
            self._data_cell(ws, r, 3, row.qty)

        self._auto_width(ws)

    def _write_sales_header(self, wb, result: ProcessingResult):
        """Sheet 4: 'Sales Header' — Grouped by SO with Order Qty, Tester Qty, Total."""
        ws = wb.create_sheet('Sales Header')
        headers = ['SO Number', 'Order Qty', 'Tester Qty', 'Total Qty',
                   'Distributor', 'City', 'State', 'Location']
        for c, h in enumerate(headers, 1):
            self._hdr_cell(ws, 1, c, h)

        # Group by SO number
        so_groups: Dict[str, dict] = {}
        for row in result.rows:
            if row.so_number not in so_groups:
                so_groups[row.so_number] = {
                    'order_qty': 0,
                    'tester_qty': 0,
                    'distributor': row.distributor,
                    'city': row.city,
                    'state': row.state,
                    'location': row.location,
                }
            so_groups[row.so_number]['order_qty'] += row.qty
            so_groups[row.so_number]['tester_qty'] += row.tester_qty

        r = 2
        for so_num, info in so_groups.items():
            order_qty = info['order_qty']
            tester_qty = info['tester_qty']
            total_qty = order_qty + tester_qty
            self._data_cell(ws, r, 1, so_num)
            self._data_cell(ws, r, 2, order_qty)
            self._data_cell(ws, r, 3, tester_qty)
            self._data_cell(ws, r, 4, total_qty)
            self._data_cell(ws, r, 5, info['distributor'])
            self._data_cell(ws, r, 6, info['city'])
            self._data_cell(ws, r, 7, info['state'])
            self._data_cell(ws, r, 8, info['location'])
            r += 1

        self._auto_width(ws)

    def _write_warnings(self, wb, result: ProcessingResult):
        """Sheet 5: 'Warnings' — Only created if warnings exist."""
        if not result.warned_files:
            return

        ws = wb.create_sheet('Warnings')
        headers = ['File', 'Warning']
        for c, h in enumerate(headers, 1):
            self._hdr_cell(ws, 1, c, h)

        for r, (fname, warning) in enumerate(result.warned_files, 2):
            self._data_cell(ws, r, 1, fname)
            self._data_cell(ws, r, 2, warning)

        self._auto_width(ws)


# ═══════════════════════════════════════════════════════════════════════════════
#  FILE OPENER (cross-platform)
# ═══════════════════════════════════════════════════════════════════════════════

def open_file(file_path: Path):
    """Opens the file using the OS default application."""
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(str(file_path))
        elif system == "Darwin":
            import subprocess as sp
            sp.Popen(["open", str(file_path)])
        else:
            import subprocess as sp
            sp.Popen(["xdg-open", str(file_path)])
    except Exception as e:
        messagebox.showerror("Open File Error", f"Could not open file:\n{e}")


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN AUTOMATION ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

class GTMassAutomation:
    """Orchestrates file parsing and export."""

    def __init__(self):
        self.parser = ExcelParser()
        self.exporter = DumpExporter()

    def process_files(self, files: List[Path]) -> ProcessingResult:
        """Process all selected files and return aggregated result."""
        result = ProcessingResult()

        for file in files:
            fname = file.name
            try:
                rows, warnings = self.parser.parse(file)
                result.rows.extend(rows)
                for w in warnings:
                    result.warned_files.append((fname, w))
                    logging.warning(f"{fname}: {w}")
            except RuntimeError as e:
                result.failed_files.append((fname, str(e)))
                logging.error(f"{fname} FAILED: {e}")
            except Exception as e:
                result.failed_files.append((fname, f"Unexpected error: {e}"))
                logging.error(f"{fname} UNEXPECTED ERROR: {e}")

        logging.info(
            f"Processing complete — "
            f"{len(result.rows)} rows | "
            f"{len(result.failed_files)} failed | "
            f"{len(result.warned_files)} warnings"
        )
        return result


# ═══════════════════════════════════════════════════════════════════════════════
#  TKINTER UI
# ═══════════════════════════════════════════════════════════════════════════════

class AutomationUI:
    """Simple Tkinter GUI for file selection and dump generation."""

    def __init__(self, automation: GTMassAutomation):
        self.automation = automation
        self.files: List[Path] = []
        self.last_output_path: Optional[Path] = None

        self.root = tk.Tk()
        self.root.title("GT Mass Dump Generator v2")
        self.root.geometry("460x380")
        self.root.resizable(False, False)

        # ── Title ──
        tk.Label(
            self.root, text="GT Mass Dump Generator",
            font=("Arial", 14, "bold")
        ).pack(pady=10)

        # ── Subtitle ──
        tk.Label(
            self.root, text="SOGTM Files → ERP Import (Headers + Lines)",
            font=("Arial", 9), fg="gray"
        ).pack(pady=0)

        # ── File count ──
        self.label = tk.Label(
            self.root, text="Selected Files: 0", font=("Arial", 10)
        )
        self.label.pack(pady=6)

        # ── Buttons ──
        tk.Button(
            self.root, text="Select Excel Files", width=22,
            command=self.select_files
        ).pack(pady=6)

        tk.Button(
            self.root, text="Generate Dump", width=22,
            command=self.generate_dump
        ).pack(pady=6)

        self.open_button = tk.Button(
            self.root, text="Open Last Output File", width=22,
            state=tk.DISABLED, command=self.open_last_file
        )
        self.open_button.pack(pady=6)

        # ── Status ──
        self.status = tk.Label(
            self.root, text="Status: Waiting", font=("Arial", 10), fg="gray"
        )
        self.status.pack(pady=6)

        # ── Time ──
        self.time_label = tk.Label(
            self.root, text="", font=("Arial", 9), fg="darkgreen"
        )
        self.time_label.pack(pady=2)

    def select_files(self):
        """Open file dialog to select SOGTM Excel files."""
        files = filedialog.askopenfilenames(
            title="Select Sales Order Files",
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
        )
        self.files = [Path(f) for f in files]
        self.label.config(text=f"Selected Files: {len(self.files)}")
        self.time_label.config(text="")
        self.status.config(text="Status: Files selected", fg="gray")

    def generate_dump(self):
        """Process files and generate the output dump."""
        if not self.files:
            messagebox.showwarning("Warning", "Please select files first.")
            return

        start_time = time.time()
        self.status.config(text="Status: Processing files...", fg="blue")
        self.time_label.config(text="")
        self.root.update()

        # ── Process ──
        result = self.automation.process_files(self.files)
        output_path = self.automation.exporter.export(result)

        # ── Timer ──
        elapsed = time.time() - start_time
        elapsed_str = f"{elapsed:.2f} seconds"

        failed = len(result.failed_files)
        warned = len(result.warned_files)
        rows = len(result.rows)
        sos = len(set(r.so_number for r in result.rows)) if result.rows else 0

        if output_path:
            self.last_output_path = output_path
            self.open_button.config(state=tk.NORMAL)

            if failed > 0 or warned > 0:
                self.status.config(
                    text=f"Done — {rows} rows | {failed} failed | {warned} warning(s)",
                    fg="orange"
                )
            else:
                self.status.config(
                    text=f"Done — {rows} rows across {sos} SO(s)",
                    fg="darkgreen"
                )

            self.time_label.config(text=f"⏱  Time taken: {elapsed_str}")

            warn_note = f"\n⚠️  {warned} warning(s) — check 'Warnings' sheet." if warned else ""
            fail_note = f"\n❌  {failed} file(s) failed — see error popup." if failed else ""

            answer = messagebox.askyesno(
                "Dump Generated",
                f"Dump generated successfully!\n\n"
                f"File   : {output_path.name}\n"
                f"Rows   : {rows}\n"
                f"SO(s)  : {sos}\n"
                f"Time   : {elapsed_str}"
                f"{warn_note}{fail_note}\n\n"
                f"Do you want to open the output file?"
            )
            if answer:
                open_file(output_path)
        else:
            self.status.config(text="Status: No data to export", fg="red")
            self.time_label.config(text=f"⏱  Time taken: {elapsed_str}")

    def open_last_file(self):
        """Open the last generated output file."""
        if self.last_output_path and self.last_output_path.exists():
            open_file(self.last_output_path)
        else:
            messagebox.showwarning(
                "File Not Found",
                "The output file no longer exists.\nPlease generate a new dump."
            )

    def run(self):
        """Start the Tkinter main loop."""
        self.root.mainloop()


# ═══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    check_expiry()
    automation = GTMassAutomation()
    ui = AutomationUI(automation)
    ui.run()


if __name__ == "__main__":
    main()