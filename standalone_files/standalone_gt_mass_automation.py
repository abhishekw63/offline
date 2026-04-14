"""
╔═══════════════════════════════════════════════════════════════════════════════╗
║               GT MASS DUMP GENERATOR — v2.1                                  ║
║               Tkinter GUI Desktop Application                                ║
╠═══════════════════════════════════════════════════════════════════════════════╣
║  Author  : Agami AI / Vishal                                                ║
║  Version : 2.1 (SO Number from file's PO Number field)                      ║
║  Purpose : Reads GT-Mass / Everyday PO Excel files from distributors,       ║
║            extracts meta info (SO Number, Distributor, City, State,          ║
║            Location) and ordered items (BC Code, Qty, Tester Qty),          ║
║            generates ERP-importable Sales Order sheets.                      ║
║  Stack   : Python 3.13, Tkinter, pandas, openpyxl                           ║
╚═══════════════════════════════════════════════════════════════════════════════╝

═══════════════════════════════════════════════════════════════════════════════
  ARCHITECTURE
═══════════════════════════════════════════════════════════════════════════════

  ┌─────────────────────────────────────────────────────────────────┐
  │                     AutomationUI (GUI)                          │
  │  User selects Excel files → clicks Generate → opens output      │
  └────────────────────────┬────────────────────────────────────────┘
                           │ passes file list
                           ▼
  ┌─────────────────────────────────────────────────────────────────┐
  │                GTMassAutomation (Engine)                         │
  │  Loops through each file:                                        │
  │    1. ExcelParser.parse()                                        │
  │       a. FileReader.read()        → raw DataFrame (no header)   │
  │       b. Find header row          → scans for 'BC Code'         │
  │       c. MetaExtractor.extract()  → SO#, Distributor, City,     │
  │                                      State, Location            │
  │       d. SO number resolution     → file first, filename backup │
  │       e. Extract ordered items    → OrderRow list               │
  │    2. Collect all rows + warnings across files                   │
  └────────────────────────┬────────────────────────────────────────┘
                           │ ProcessingResult
                           ▼
  ┌─────────────────────────────────────────────────────────────────┐
  │                DumpExporter (Output)                             │
  │  Creates output Excel with 5 sheets:                             │
  │    Sheet 1: Headers (SO)   → ERP Sales Order header import      │
  │    Sheet 2: Lines (SO)     → ERP Sales Order line import        │
  │    Sheet 3: Sales Lines    → Simple flat: SO | Item No | Qty    │
  │    Sheet 4: Sales Header   → Grouped summary with meta info     │
  │    Sheet 5: Warnings       → Parsing issues (if any)            │
  └─────────────────────────────────────────────────────────────────┘

═══════════════════════════════════════════════════════════════════════════════
  INPUT FILE STRUCTURE
═══════════════════════════════════════════════════════════════════════════════

  Each GT-Mass / Everyday PO file has this layout:

  ┌─────────────────────────────────────────────────────────────────────────┐
  │ ROW │ COL A               │ COL B                │ COL G        │ COL I │
  ├─────┼─────────────────────┼──────────────────────┼──────────────┼───────┤
  │  0  │ Title (ignored)     │                      │              │       │
  │  1  │ "Distributor Name"  │ "Classic Enterprises" │ "ASM"       │ name  │
  │  2  │ "DB Code"           │ 20084                │ "RSM"        │ name  │
  │  3  │ "BDE Name"          │ "Annamalai"          │ "PO Number" │SO/GTM/│
  │  4  │ "City"              │ "Chennai"            │ "Date of PO" │ date │
  │  5  │ "State"             │ "Tamilnadu"          │ "Location"   │ "AHD" │
  │  6  │ EAN │ BC Code │ ... │ Order Qty │ Tester Qty │ (header row)       │
  │  7+ │ data │ data  │ ... │ data      │ data       │ (data rows)         │
  └─────────────────────────────────────────────────────────────────────────┘

  KEY META FIELDS:
    Row 3, Col I  → SO/GTM Number (PO Number field) — PRIMARY source for SO#
    Row 5, Col I  → Location (e.g., "AHD", "BLR")
    Row 1, Col B  → Distributor Name
    Row 4, Col B  → City
    Row 5, Col B  → State

  ⚠ Some files have blank meta (Distributor, City, State empty) — these
    still work because SO Number and Location are in the right-side columns.

═══════════════════════════════════════════════════════════════════════════════
  SO NUMBER RESOLUTION PRIORITY
═══════════════════════════════════════════════════════════════════════════════

  1. File's PO Number field (Row 3, Col I)
     → e.g., "SO/GTM/6448" → used directly ✓
     → This is the PRIMARY and preferred source

  2. Filename digits (FALLBACK — legacy support)
     → e.g., "SOGTM6325.xlsx" → "SO/GTM/6325"
     → Used only if PO Number field is blank
     → Logs a WARNING asking team to fill PO Number field

  3. "SO/GTM/UNKNOWN" (LAST RESORT)
     → Used if both file and filename have no SO number
     → Logs a WARNING — must be fixed manually

═══════════════════════════════════════════════════════════════════════════════
  LOCATION CODE MAPPING
═══════════════════════════════════════════════════════════════════════════════

  The raw Location value from the file is mapped to an ERP Location Code:

  ┌──────────────┬────────────────┬──────────────────────────────────┐
  │ File Value   │ ERP Code       │ Notes                            │
  ├──────────────┼────────────────┼──────────────────────────────────┤
  │ AHD          │ PICK           │ Ahmedabad warehouse              │
  │ BLR          │ DS_BL_OFF1     │ Bangalore dispatch office        │
  │ (other)      │ (raw value)    │ Used as-is until mapping added   │
  │ (empty)      │ (empty)        │ Left blank in output             │
  └──────────────┴────────────────┴──────────────────────────────────┘

  To add new mappings: update LOCATION_CODE_MAP dict below.

═══════════════════════════════════════════════════════════════════════════════
  OUTPUT — 5 EXCEL SHEETS
═══════════════════════════════════════════════════════════════════════════════

  Sheet 1: 'Headers (SO)' — One row per unique SO number
      Document Type = 'Order' | No. = SO number | 5 × date = today |
      Location Code = mapped | Supply Type = 'B2B'
      ⚠ Sell-to Customer No. and Ship-to Code left EMPTY (manual)

  Sheet 2: 'Lines (SO)' — One row per ordered item
      Document No. = SO number | Line No. = 10000, 20000... (resets per SO) |
      Type = 'Item' | No. = BC Code | Quantity = Order Qty |
      Unit Price = EMPTY (ERP fetches from item card)

  Sheet 3: 'Sales Lines' — Flat reference list
      SO Number | Item No | Qty

  Sheet 4: 'Sales Header' — Grouped summary for cross-verification
      SO Number | Order Qty | Tester Qty | Total Qty |
      Distributor | City | State | Location

  Sheet 5: 'Warnings' — Only created if warnings exist
      File | Warning

═══════════════════════════════════════════════════════════════════════════════
  EXPIRY SYSTEM
═══════════════════════════════════════════════════════════════════════════════

  Built-in license expiry (EXPIRY_DATE constant).
  - After expiry: error popup → exits
  - Within 7 days: warning popup → continues
  - To extend: change EXPIRY_DATE value

═══════════════════════════════════════════════════════════════════════════════
  DEPENDENCIES & RUNNING
═══════════════════════════════════════════════════════════════════════════════

  Requirements:
      pip install pandas openpyxl

  For legacy .xls files:
      pip install xlrd

  Run:
      python gt_mass_dump.py
"""

# ═══════════════════════════════════════════════════════════════════════════════
#  IMPORTS
# ═══════════════════════════════════════════════════════════════════════════════

from __future__ import annotations      # Enable forward type references (e.g., List[OrderRow])
import os                               # File/path operations (os.startfile for Windows)
import sys                              # sys.exit() for expiry check
import platform                         # platform.system() to detect OS for file opener
import time                             # time.time() to measure processing duration
import logging                          # Structured logging to console
import re                               # Regex for SO number extraction from filename
import pandas as pd                     # Excel reading into DataFrames
import tkinter as tk                    # GUI framework (bundled with Python)
from tkinter import filedialog, messagebox  # File chooser dialog + alert popups
from dataclasses import dataclass, field    # @dataclass for clean data containers
from pathlib import Path                # Cross-platform file path handling
from typing import List, Optional, Tuple, Dict  # Type annotations for clarity
from datetime import datetime           # Date parsing (expiry) + timestamp formatting

from openpyxl import Workbook           # Excel writing for output workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side  # Cell formatting
from openpyxl.utils import get_column_letter  # Convert column index (1) → letter ('A')


# ═══════════════════════════════════════════════════════════════════════════════
#  LOGGING CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════
# All log messages go to the console (stdout).
# Format: "2026-04-14 10:30:45 | INFO | Message here"
# Change level to logging.DEBUG for verbose debugging.

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)


# ═══════════════════════════════════════════════════════════════════════════════
#  EXPIRY CHECK
# ═══════════════════════════════════════════════════════════════════════════════
# Built-in license expiry for the desktop tool.
# Change EXPIRY_DATE to extend validity.
# Format: DD-MM-YYYY (e.g., "30-06-2026")

EXPIRY_DATE = "30-06-2026"


def check_expiry():
    """
    Check if the application has expired.

    Behavior:
      - Past expiry: show error popup → exit application
      - Within 7 days of expiry: show warning popup → continue normally
      - More than 7 days: do nothing

    Called once at application startup (in main()).
    """
    # Parse the expiry date string into a date object
    expiry = datetime.strptime(EXPIRY_DATE, "%d-%m-%Y").date()
    today = datetime.now().date()

    # ── Expired: block application ──
    if today > expiry:
        root = tk.Tk()
        root.withdraw()  # Hide the empty Tk window (only show the message box)
        messagebox.showerror(
            "Application Expired",
            f"This application expired on {EXPIRY_DATE}.\n\n"
            f"Please contact the administrator for an updated version."
        )
        root.destroy()
        sys.exit(0)  # Hard exit — no processing allowed

    # ── Expiring soon: warn but allow ──
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
# │ Maps the raw Location value from GT-Mass files to ERP Location Codes.  │
# │                                                                         │
# │ How it works:                                                           │
# │   - File says "AHD" → script outputs "PICK" in the Location Code col  │
# │   - File says "BLR" → script outputs "DS_BL_OFF1"                     │
# │   - File says anything NOT in this map → raw value used as-is          │
# │   - File has no Location → Location Code is left empty                 │
# │                                                                         │
# │ To add a new mapping:                                                   │
# │   1. Find out what value appears in the file (e.g., "DEL")            │
# │   2. Find out the ERP Location Code (e.g., "DS_DL_OFF1")              │
# │   3. Add: 'DEL': 'DS_DL_OFF1',                                        │
# └─────────────────────────────────────────────────────────────────────────┘
LOCATION_CODE_MAP: Dict[str, str] = {
    'AHD': 'PICK',           # Ahmedabad → PICK warehouse
    'BLR': 'DS_BL_OFF1',    # Bangalore → Dispatch office
    # Add more mappings here as new locations are discovered:
    # 'DEL': 'DS_DL_OFF1',  # Delhi
    # 'MUM': 'DS_MUM01',    # Mumbai
}

# ┌─────────────────────────────────────────────────────────────────────────┐
# │ STATE / ZONE DETECTION                                                  │
# │                                                                         │
# │ Safety check: if a value that looks like a state name appears in the   │
# │ Distributor Name field, it probably means the rows in the source file  │
# │ are swapped (someone put the state where the distributor should be).   │
# │                                                                         │
# │ When detected: logs a warning → shows in Warnings sheet.               │
# └─────────────────────────────────────────────────────────────────────────┘
STATE_LIKE_VALUES = {
    # Two-letter state codes
    "up", "mp", "ap", "hp", "uk", "jk", "wb", "tn", "kl", "ka",
    "gj", "rj", "hr", "pb", "br", "od", "as", "mh", "cg", "jh",
    # Zone names (sometimes used instead of state)
    "north", "south", "east", "west", "central",
    # Full state names (lowercase for matching)
    "uttar pradesh", "madhya pradesh", "rajasthan", "punjab",
    "maharashtra", "gujarat", "karnataka", "tamil nadu",
    "haryana", "delhi", "u.p", "u.p.", "m.p", "m.p."
}


# ═══════════════════════════════════════════════════════════════════════════════
#  DATA MODEL
# ═══════════════════════════════════════════════════════════════════════════════
# These dataclasses hold the structured data extracted from the input files.
# They are passed between the parser, engine, and exporter.

@dataclass
class OrderRow:
    """
    Single ordered item extracted from a GT-Mass file.

    Each row represents one product line in the purchase order.
    One file can produce multiple OrderRow objects (one per SKU with qty > 0).

    Fields:
        so_number     : SO/GTM number (e.g., 'SO/GTM/6448') — from file or filename
        item_no       : BC Code / Item No for ERP (e.g., '200163')
        ean           : EAN barcode from the file (e.g., '8904473104307')
        category      : Product category from the file (e.g., 'Eye', 'FACE')
        description   : Article Description from the file (e.g., 'RENEE PURE BROWN KAJAL...')
        qty           : Order Qty — regular stock quantity
        tester_qty    : Tester Qty — tester/sample quantity (separate from order)
        distributor   : Distributor Name from the file's meta header
        city          : City from the file's meta header
        state         : State from the file's meta header
        location      : Raw location value (e.g., 'AHD', 'BLR') from file
        location_code : Mapped ERP Location Code (e.g., 'PICK', 'DS_BL_OFF1')
    """
    so_number: str
    item_no: str
    ean: str
    category: str
    description: str
    qty: int
    tester_qty: int
    distributor: str
    city: str
    state: str
    location: str
    location_code: str


@dataclass
class ProcessingResult:
    """
    Aggregated result from processing all selected files.

    The GTMassAutomation engine populates this as it processes each file,
    then the DumpExporter uses it to write the output Excel.

    Fields:
        rows         : All OrderRow objects across all files (main data)
        failed_files : Files that could not be read at all → (filename, error_message)
        warned_files : Files that processed but had warnings → (filename, warning_message)
        output_path  : Path to the generated output file (set after export)
    """
    rows: List[OrderRow] = field(default_factory=list)
    failed_files: List[Tuple[str, str]] = field(default_factory=list)
    warned_files: List[Tuple[str, str]] = field(default_factory=list)
    output_path: Optional[Path] = None


# ═══════════════════════════════════════════════════════════════════════════════
#  SO FORMATTER (Filename Fallback)
# ═══════════════════════════════════════════════════════════════════════════════
# This class provides the FALLBACK method for extracting SO numbers from
# filenames. It's used ONLY when the file's internal PO Number field is empty.
#
# Pattern: looks for any sequence of digits in the filename stem.
#   "SOGTM6325.xlsx"  → "SO/GTM/6325"  (correct)
#   "GT-Mass_PO_Format_April_26.xlsx" → "SO/GTM/26" (not ideal — hence fallback)
#
# The PRIMARY source for SO numbers is now the file's PO Number field
# (Row 3, Column I), which is extracted by MetaExtractor.

class SOFormatter:
    """Extracts SO number from filename as a fallback when PO Number field is empty."""

    @staticmethod
    def from_filename(filepath: Path) -> Optional[str]:
        """
        Extract SO number from filename digits.

        Example: "SOGTM6325.xlsx" → stem "SOGTM6325" → digits "6325" → "SO/GTM/6325"

        Args:
            filepath: Path to the Excel file

        Returns:
            "SO/GTM/####" string, or None if no digits found in filename
        """
        # Search for first sequence of digits in the filename (without extension)
        match = re.search(r"\d+", filepath.stem)
        if not match:
            logging.warning(f"SO number not found in filename: {filepath.name}")
            return None
        # Format as SO/GTM/#### (standard ERP format for GT-Mass orders)
        return f"SO/GTM/{match.group()}"


# ═══════════════════════════════════════════════════════════════════════════════
#  FILE READER
# ═══════════════════════════════════════════════════════════════════════════════
# Reads Excel files into raw DataFrames with NO header (header=None).
# This is because the actual data header is NOT in row 0 — it's buried
# below the meta rows (typically row 6). The caller finds the header row.
#
# Reading strategy by file extension:
#   .xlsx / .xlsm  → openpyxl engine (default, bundled with pandas)
#   .xls           → xlrd engine (legacy format, requires: pip install xlrd)

class FileReader:
    """Reads Excel files into raw DataFrames (no header row assumed)."""

    @staticmethod
    def read(file_path: Path) -> pd.DataFrame:
        """
        Read an Excel file and return ALL rows as a raw DataFrame.

        The DataFrame has integer column indices (0, 1, 2...) because
        header=None — the actual column names come later when we identify
        the header row.

        Args:
            file_path: Path to the .xlsx/.xlsm/.xls file

        Returns:
            pd.DataFrame with all rows, no header

        Raises:
            RuntimeError: if file cannot be read (corrupt, password-protected,
                          unsupported format, missing xlrd for .xls)
        """
        ext = file_path.suffix.lower()  # Get extension: '.xlsx', '.xls', etc.

        # ── Modern Excel format (.xlsx / .xlsm) → openpyxl engine ──
        if ext in (".xlsx", ".xlsm"):
            try:
                df = pd.read_excel(file_path, header=None, engine="openpyxl")
                logging.info(f"{file_path.name} — read via openpyxl ({len(df)} rows)")
                return df
            except Exception as e:
                raise RuntimeError(
                    f"Cannot read '{file_path.name}'.\n"
                    f"The file may be corrupt or password-protected.\n"
                    f"Error: {e}"
                )

        # ── Legacy Excel format (.xls) → xlrd engine ──
        if ext == ".xls":
            try:
                df = pd.read_excel(file_path, header=None, engine="xlrd")
                logging.info(f"{file_path.name} — read via xlrd ({len(df)} rows)")
                return df
            except ImportError:
                # xlrd not installed — give clear installation instructions
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

        # ── Unsupported format (.csv, .pdf, etc.) ──
        raise RuntimeError(
            f"Unsupported file format: '{ext}'.\n"
            f"Only .xlsx, .xlsm, and .xls files are supported."
        )


# ═══════════════════════════════════════════════════════════════════════════════
#  META EXTRACTOR
# ═══════════════════════════════════════════════════════════════════════════════
# Scans the meta header region (rows above the data table) to extract:
#
#   LEFT SIDE (Col A = label, Col B = value):
#     - Distributor Name  (Row 1 typically)
#     - City              (Row 4 typically)
#     - State             (Row 5 typically — picks last non-blank if multiple)
#
#   RIGHT SIDE (Col G = label, Col I = value):
#     - PO Number → SO/GTM number  (Row 3, e.g., "SO/GTM/6448")
#     - Location                    (Row 5, e.g., "AHD", "BLR")
#
# Row positions vary slightly across files, so we SCAN by label matching
# rather than hardcoding row indices. This makes the code resilient to
# minor layout changes.

class MetaExtractor:
    """Extracts meta fields (SO#, Distributor, City, State, Location) from header rows."""

    @staticmethod
    def extract(raw_df: pd.DataFrame, header_row: int) -> Tuple[dict, List[str]]:
        """
        Scan rows 0 through header_row-1 for meta field labels and values.

        Scanning approach:
          - For each row in the meta region:
            1. Check Col A for labels ("Distributor Name", "City", "State")
            2. Read corresponding value from Col B
            3. Scan Cols 0-9 for right-side labels ("PO Number", "Location")
            4. Read their values from the next 1-2 columns

        Args:
            raw_df    : Full DataFrame (no header) as read by FileReader
            header_row: Row index where the data header (BC Code, Order Qty...) starts

        Returns:
            Tuple of:
              - meta_dict with keys: distributor, city, state, location,
                location_code, so_number
              - warnings list (empty if all fields found)
        """
        # ── Initialize empty meta fields ──
        distributor = ""        # Distributor company name
        city = ""               # City of the distributor
        state_values = []       # Collect ALL state values (pick last non-blank later)
        location = ""           # Raw location code from file (e.g., "AHD")
        so_number = ""          # SO/GTM number from PO Number field
        warnings = []           # Accumulate any issues found

        # ── Only scan the meta rows (above the header row) ──
        meta_df = raw_df.iloc[:header_row]

        # Iterate via values for performance (avoiding overhead of Series creation per row)
        for row_vals in meta_df.values:
            # ── Column A/B scanning (Distributor, City, State) ──
            label = str(row_vals[0]).strip().lower() if pd.notna(row_vals[0]) else ""
            value = str(row_vals[1]).strip() if pd.notna(row_vals[1]) else ""
            if value.lower() in ("nan", ""):
                value = ""

            # Match labels and extract values (first match wins for distributor/city)
            if label == "distributor name" and not distributor:
                distributor = value
                logging.info(f"Distributor found: '{distributor}'")
            elif label == "city" and not city:
                city = value
                logging.info(f"City found: '{city}'")
            elif label == "state":
                # Collect ALL state values — some files have state in multiple rows
                # We'll pick the last non-blank one later
                state_values.append(value)

            # ── Column G/I scanning (Location) ──
            # Location label is typically at column index 6 ("Location")
            # Location value is at column index 8 (e.g., "AHD")
            for col_idx in range(min(len(row_vals) - 1, 10)):
                cell_val = str(row_vals[col_idx]).strip().lower() if pd.notna(row_vals[col_idx]) else ""
                if cell_val == "location":
                    # Look for value in the next available column(s)
                    for val_idx in range(col_idx + 1, min(col_idx + 3, len(row_vals))):
                        loc_val = row_vals[val_idx]
                        if pd.notna(loc_val) and str(loc_val).strip() and str(loc_val).strip().lower() != 'nan':
                            location = str(loc_val).strip()
                            logging.info(f"Location found: '{location}'")
                            break  # Stop at first non-blank value

                # ── "PO Number" label → read SO/GTM number from next 1-2 columns ──
                elif cell_val == "po number" and not so_number:
                    for val_idx in range(col_idx + 1, min(col_idx + 3, len(row))):
                        po_val = row.iloc[val_idx]
                        if pd.notna(po_val) and str(po_val).strip() and str(po_val).strip().lower() != 'nan':
                            so_number = str(po_val).strip()
                            logging.info(f"SO Number found in file: '{so_number}'")
                            break  # Stop at first non-blank value

        # ── Resolve State: pick last non-blank value ──
        # Some files have "State" label in multiple rows. The bottom one
        # (closer to the data) is usually the correct state.
        state = ""
        for s in reversed(state_values):
            if s:
                state = s
                break
        logging.info(f"State found: '{state}'")

        # ── Map raw Location to ERP Location Code ──
        # e.g., "AHD" → "PICK", "BLR" → "DS_BL_OFF1"
        location_code = ""
        if location:
            location_upper = location.upper().strip()
            if location_upper in LOCATION_CODE_MAP:
                # Known mapping exists → use the mapped ERP code
                location_code = LOCATION_CODE_MAP[location_upper]
                logging.info(f"Location '{location}' → mapped to '{location_code}'")
            else:
                # Unknown location → use raw value as-is (user can add mapping later)
                location_code = location
                logging.info(f"Location '{location}' → no mapping found, using raw value")

        # ── Validation: warn about missing or suspicious meta fields ──
        if not distributor:
            warnings.append("Distributor Name is blank — label not found or value is empty.")
        if not city:
            warnings.append("City is blank — label not found or value is empty.")
        if not state:
            warnings.append("State is blank — both State rows are empty or missing.")

        # Safety check: if Distributor value looks like a state name,
        # the source file probably has swapped rows
        if distributor and distributor.strip().lower() in STATE_LIKE_VALUES:
            warnings.append(
                f"Distributor value '{distributor}' looks like a state/zone name. "
                f"Rows may be swapped in the source file — please verify manually."
            )

        # ── Return all extracted meta + any warnings ──
        return {
            "distributor": distributor,
            "city": city,
            "state": state,
            "location": location,
            "location_code": location_code,
            "so_number": so_number,
        }, warnings


# ═══════════════════════════════════════════════════════════════════════════════
#  EXCEL PARSER
# ═══════════════════════════════════════════════════════════════════════════════
# The main parsing logic for a single GT-Mass / Everyday PO file.
#
# Processing flow:
#   1. Read raw file (FileReader.read)
#   2. Find the header row by scanning for 'BC Code' + 'Order Qty'
#   3. Extract meta fields from rows above the header (MetaExtractor.extract)
#   4. Resolve SO number (file PO Number → filename fallback → UNKNOWN)
#   5. Build data table from rows below the header
#   6. Extract ordered items (BC Code > 0 AND qty > 0 or tester > 0)
#   7. Return list of OrderRow objects + any warnings

class ExcelParser:
    """Parses a single GT-Mass / Everyday PO Excel file into OrderRow objects."""

    # Column name constants (lowercase for case-insensitive matching)
    BC_COLUMN = "bc code"           # Column containing Item No (BC Code from ERP)
    QTY_COLUMN = "order qty"        # Column containing regular order quantity
    TESTER_COLUMN = "tester qty"    # Column containing tester/sample quantity

    def parse(self, file_path: Path) -> Tuple[List[OrderRow], List[str]]:
        """
        Parse one GT-Mass / Everyday PO file.

        Full processing pipeline:
          1. Read raw Excel → DataFrame with no header
          2. Scan for header row (contains both 'BC Code' and 'Order Qty')
          3. Extract meta fields (SO#, Distributor, City, State, Location)
          4. Resolve SO number (file first, filename fallback)
          5. Build data table from rows below header
          6. Detect column positions (BC Code, Order Qty, Tester Qty)
          7. Loop through data rows → create OrderRow for each valid item
          8. Return (rows, warnings)

        Args:
            file_path: Path to the Excel file

        Returns:
            Tuple of (list_of_OrderRow, list_of_warning_strings)

        Raises:
            RuntimeError: if file is unreadable or has broken structure
                          (no header row, no BC Code column, etc.)
        """
        logging.info(f"Parsing: {file_path.name}")
        warnings = []

        # ── Step 1: Read the raw Excel file into a DataFrame ──
        # Returns ALL rows with integer column indices (0, 1, 2...)
        # No header assumed — meta rows and data rows are all mixed in
        raw_df = FileReader.read(file_path)

        # ── Step 2: Find the header row ──
        # The header row is the one that contains BOTH 'BC Code' AND 'Order Qty'.
        # It's typically Row 6, but we scan to be safe.
        header_row = None
        # Efficiently scan using enumeration on values to avoid Series generation overhead
        for i, row_vals in enumerate(raw_df.values):
            row_values = [str(v).lower() for v in row_vals]
            if "bc code" in row_values and any("order qty" in v for v in row_values):
                header_row = i
                break

        # If header row not found, the file format is broken → abort
        if header_row is None:
            raise RuntimeError(
                "Header row not found — could not locate both 'BC Code' and 'Order Qty'. "
                "File format may have changed."
            )

        # ── Step 3: Extract meta fields from rows ABOVE the header ──
        # MetaExtractor scans rows 0..header_row-1 for labels like
        # "Distributor Name", "City", "State", "PO Number", "Location"
        meta, meta_warnings = MetaExtractor.extract(raw_df, header_row)
        warnings.extend(meta_warnings)  # Add any meta warnings to our list

        # ── Step 4: Resolve SO number ──
        # Priority: file's PO Number field > filename digits > UNKNOWN
        so_number = meta.get("so_number", "")
        if so_number:
            # BEST CASE: SO number found inside the file (PO Number field)
            logging.info(f"Using SO number from file: '{so_number}'")
        else:
            # FALLBACK: try extracting from filename (legacy SOGTM#### pattern)
            so_number = SOFormatter.from_filename(file_path)
            if so_number:
                logging.info(f"SO number not in file — using filename: '{so_number}'")
                warnings.append(
                    f"SO number not found inside file (PO Number field is empty). "
                    f"Using filename-based SO: '{so_number}'. "
                    f"Please ask the team to fill in PO Number field."
                )
            else:
                # LAST RESORT: no SO number anywhere
                so_number = "SO/GTM/UNKNOWN"
                warnings.append(
                    "SO number not found in file or filename. "
                    "Using 'SO/GTM/UNKNOWN' — please fix manually."
                )

        # ── Step 5: Build the data table from rows BELOW the header ──
        # Take all rows after the header row
        df = raw_df.iloc[header_row + 1:].copy()
        # Set column names from the header row values
        df.columns = raw_df.iloc[header_row].values
        # Reset index so it starts from 0
        df = df.reset_index(drop=True)

        # ── Step 6: Detect column positions ──
        # Find which columns are BC Code, Order Qty, Tester Qty, EAN, Category, Description
        bc_col, qty_col, tester_col, ean_col, cat_col, desc_col = self._detect_columns(df)

        # BC Code is mandatory — can't create SO lines without item numbers
        if bc_col is None:
            raise RuntimeError("Column 'BC Code' not found in data table.")
        # Order Qty is mandatory — need to know how many to order
        if qty_col is None:
            raise RuntimeError("Column 'Order Qty' not found in data table.")
        # Tester Qty is optional — some files don't have it
        if tester_col is None:
            warnings.append("Column 'Tester Qty' not found — tester quantities will be 0.")

        # ── Step 7: Extract ordered items ──
        # Loop through each data row and create OrderRow for items with qty > 0
        rows: List[OrderRow] = []

        # Get indices for faster access over values
        bc_idx = df.columns.get_loc(bc_col)
        qty_idx = df.columns.get_loc(qty_col)
        tester_idx = df.columns.get_loc(tester_col) if tester_col is not None else None

        # Iterate via values for performance (avoiding overhead of Series creation per row)
        for row_vals in df.values:
            bc_code = row_vals[bc_idx]
            if pd.isna(bc_code):
                continue
            # BC Code must be numeric (e.g., 200163). Skip non-numeric values
            # (catches "Total", "Grand Total", category headers, etc.)
            try:
                bc_code = int(bc_code)
            except (ValueError, TypeError):
                # If BC Code is not numeric or convertible, skip the row
                continue

            qty = self._clean_qty(row_vals[qty_idx])
            tester_qty = self._clean_qty(row_vals[tester_idx]) if tester_idx is not None else 0

            # Skip rows where BOTH order and tester are zero/blank
            # (no point creating a line for zero-quantity items)
            if qty <= 0 and tester_qty <= 0:
                continue

            # ── Read EAN, Category, Description (optional — for Sales Lines reference) ──
            ean_val = ''
            if ean_idx is not None and pd.notna(row_vals[ean_idx]):
                ean_raw = row_vals[ean_idx]
                # EAN is numeric in Excel — convert to string without '.0'
                ean_val = str(int(ean_raw)) if isinstance(ean_raw, (int, float)) else str(ean_raw).strip()

            cat_val = ''
            if cat_idx is not None and pd.notna(row_vals[cat_idx]):
                cat_val = str(row_vals[cat_idx]).strip()

            desc_val = ''
            if desc_idx is not None and pd.notna(row_vals[desc_idx]):
                desc_val = str(row_vals[desc_idx]).strip()

            # ── Create OrderRow with all fields ──
            rows.append(OrderRow(
                so_number=so_number,
                item_no=str(bc_code),               # Store as string for ERP compatibility
                ean=ean_val,                         # EAN barcode from file
                category=cat_val,                    # Product category (Eye, FACE, etc.)
                description=desc_val,                # Article Description
                qty=qty,
                tester_qty=tester_qty,
                distributor=meta["distributor"],
                city=meta["city"],
                state=meta["state"],
                location=meta["location"],           # Raw value (e.g., "AHD")
                location_code=meta["location_code"],  # Mapped value (e.g., "PICK")
            ))

        # Warn if no items found (file exists but all quantities are zero)
        if not rows:
            warnings.append(
                "No ordered rows found — all Order Qty and Tester Qty values are 0 or blank."
            )

        return rows, warnings

    def _detect_columns(self, df) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        """
        Find the BC Code, Order Qty, and Tester Qty columns by name matching.

        Scans all column names in the data table header and matches:
          - Exact match for 'bc code'
          - Substring match for 'order qty' (handles variants like 'Order Qty (Retail)')
          - Substring match for 'tester qty'
          - Exact match for 'ean'
          - Exact match for 'category'
          - Substring match for 'article description' or 'description'

        Args:
            df: DataFrame with column names set from the header row

        Returns:
            Tuple of (bc_col, qty_col, tester_col, ean_col, cat_col, desc_col)
            Any can be None if not found.
        """
        bc_col = None
        qty_col = None
        tester_col = None
        ean_col = None
        cat_col = None
        desc_col = None
        for col in df.columns:
            name = str(col).strip().lower()
            if name == self.BC_COLUMN:                    # Exact: "bc code"
                bc_col = col
            if self.QTY_COLUMN in name:                   # Substring: "order qty"
                qty_col = col
            if self.TESTER_COLUMN in name:                # Substring: "tester qty"
                tester_col = col
            if name == 'ean' and ean_col is None:         # Exact: "ean"
                ean_col = col
            if name == 'category' and cat_col is None:    # Exact: "category"
                cat_col = col
            if 'article description' in name or (name == 'description' and desc_col is None):
                desc_col = col                            # Substring: "article description"
        return bc_col, qty_col, tester_col, ean_col, cat_col, desc_col

    @staticmethod
    def _clean_qty(value) -> int:
        """
        Clean and convert a quantity cell value to int.

        Handles common messy values from Excel:
          - NaN / None → 0
          - Empty string / "-" → 0
          - "1,234" → 1234 (remove commas)
          - "12.0" → 12 (float to int)
          - "abc" → 0 (non-numeric)

        Args:
            value: Raw cell value from DataFrame

        Returns:
            int quantity, or 0 if value is invalid/empty
        """
        if pd.isna(value):
            return 0
        value = str(value).strip()
        if value in ("", "-"):
            return 0
        value = value.replace(",", "")  # Remove thousands separator
        try:
            return int(float(value))    # float() first handles "12.0", then int()
        except ValueError:
            return 0


# ═══════════════════════════════════════════════════════════════════════════════
#  DUMP EXPORTER
# ═══════════════════════════════════════════════════════════════════════════════
# Writes the final output Excel workbook with 5 sheets.
#
# Sheet structure:
#   1. Headers (SO)  — One row per unique SO → ERP Sales Order header import
#   2. Lines (SO)    — One row per item → ERP Sales Order line import
#   3. Sales Lines   — Simple flat reference list (SO | Item | Qty)
#   4. Sales Header  — Grouped summary per SO with Order/Tester/Total qty
#   5. Warnings      — Only created if any warnings/issues were found
#
# All sheets use consistent formatting:
#   - Headers: navy blue background, white bold text (Aptos Display 11pt)
#   - Data: light grey borders, Aptos Display 11pt font
#   - Auto column widths based on content

class DumpExporter:
    """Writes the output Excel file with ERP import sheets."""

    # ── Shared Excel formatting constants ──
    HEADER_FILL = PatternFill('solid', fgColor='1A237E')     # Navy blue header background
    HEADER_FONT = Font(bold=True, color='FFFFFF', name='Aptos Display', size=11)  # White bold header text
    DATA_FONT = Font(name='Aptos Display', size=11)          # Standard data cell font
    THIN_SIDE = Side(style='thin', color='CCCCCC')           # Light grey border line
    BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)  # All-sides border

    def _hdr_cell(self, ws, row, col, value):
        """
        Create a formatted header cell (navy blue background, white bold text).

        Used for row 1 (column headers) in every sheet.

        Args:
            ws   : openpyxl worksheet
            row  : row number (1-indexed)
            col  : column number (1-indexed)
            value: cell content (typically column header name)
        """
        cell = ws.cell(row=row, column=col, value=value)
        cell.font = self.HEADER_FONT
        cell.fill = self.HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = self.BORDER
        return cell

    def _data_cell(self, ws, row, col, value, fmt=None):
        """
        Create a formatted data cell (standard font, light border).

        Used for all data rows (row 2+) in every sheet.

        Args:
            ws   : openpyxl worksheet
            row  : row number (1-indexed)
            col  : column number (1-indexed)
            value: cell content
            fmt  : optional Excel number format string (e.g., '#,##0')
        """
        cell = ws.cell(row=row, column=col, value=value)
        cell.font = self.DATA_FONT
        cell.border = self.BORDER
        if fmt:
            cell.number_format = fmt
        return cell

    def _auto_width(self, ws, max_w=50):
        """
        Auto-fit column widths based on the longest value in each column.

        Scans all cells in each column, finds the max character length,
        adds 3 for padding. Caps at max_w to prevent extremely wide columns.

        Args:
            ws    : openpyxl worksheet
            max_w : maximum allowed column width (default 50 characters)
        """
        for col in ws.columns:
            letter = col[0].column_letter  # Get column letter (A, B, C...)
            # Find the longest cell value in this column
            w = max((len(str(c.value or '')) for c in col), default=8)
            # Set width with padding, capped at max
            ws.column_dimensions[letter].width = min(w + 3, max_w)

    def export(self, result: ProcessingResult) -> Optional[Path]:
        """
        Write the complete output Excel workbook.

        Creates an 'output' folder in the current directory if it doesn't exist,
        generates a timestamped filename, writes all 5 sheets, saves the file.

        Args:
            result: ProcessingResult with all rows, warnings, failed files

        Returns:
            Path to the generated file, or None if no data to export
        """
        # ── Show error popup for any files that failed to read ──
        if result.failed_files:
            msg = "The following files could NOT be read and were skipped:\n\n"
            for fname, reason in result.failed_files:
                msg += f"  • {fname}\n    Reason: {reason}\n\n"
            msg += "Please fix these files and re-process them."
            messagebox.showerror("Files Failed to Read", msg)

        # ── No data at all? Nothing to export ──
        if not result.rows:
            messagebox.showwarning(
                "No Data",
                "No valid rows found across all selected files.\nNothing to export."
            )
            return None

        # ── Prepare output file path ──
        output_folder = Path("output")
        output_folder.mkdir(exist_ok=True)  # Create 'output' folder if it doesn't exist
        today = datetime.now().strftime("%d-%m-%Y_%H%M%S")  # Timestamp for unique filename
        file_path = output_folder / f"gt_mass_dump_{today}.xlsx"

        # ── Create workbook and write all sheets ──
        wb = Workbook()
        wb.remove(wb.active)  # Remove the default empty sheet ("Sheet")

        self._write_headers_so(wb, result)     # Sheet 1: ERP SO header import
        self._write_lines_so(wb, result)       # Sheet 2: ERP SO line import
        self._write_sales_lines(wb, result)    # Sheet 3: Detailed reference list
        self._write_sales_header(wb, result)   # Sheet 4: Grouped summary per SO
        self._write_sku_summary(wb, result)    # Sheet 5: SKU-level pivot (overall demand)
        self._write_warnings(wb, result)       # Sheet 6: Warnings (if any exist)

        wb.save(str(file_path))
        logging.info(f"Output saved: {file_path}")
        return file_path

    def _write_headers_so(self, wb, result: ProcessingResult):
        """
        Sheet 1: 'Headers (SO)' — One row per unique SO number.

        This sheet is imported into Business Central ERP as Sales Order headers.
        Each row creates one Sales Order in the system.

        Columns:
            Document Type          = 'Order' (always — tells ERP this is a Sales Order)
            No.                    = SO/GTM number (e.g., 'SO/GTM/6448')
            Sell-to Customer No.   = EMPTY (filled manually in ERP — distributor code)
            Ship-to Code           = EMPTY (filled manually in ERP)
            Posting Date           = today's date (DD-MM-YYYY)
            Order Date             = today's date
            Document Date          = today's date
            Invoice From Date      = today's date
            Invoice To Date        = today's date
            External Document No.  = same as No. (for cross-reference between systems)
            Location Code          = mapped from file (e.g., 'PICK', 'DS_BL_OFF1')
            Dimension Set ID       = EMPTY (ERP auto-fills)
            Supply Type            = 'B2B' (always — Business-to-Business)
            Columns 14-18          = EMPTY dimension placeholders (Brand, Channel, etc.)
        """
        ws = wb.create_sheet('Headers (SO)')

        # Define all column headers matching the ERP import template
        headers = [
            'Document Type', 'No.', 'Sell-to Customer No.', 'Ship-to Code',
            'Posting Date', 'Order Date', 'Document Date',
            'Invoice From Date', 'Invoice To Date',
            'External Document No.', 'Location Code', 'Dimension Set ID',
            'Supply Type', 'Voucher Narration',
            'Brand Code (Dimension)', 'Channel Code (Dimension)',
            'Catagory (Dimension)', 'Geography Code (Dimension)'
        ]
        # Write header row (row 1)
        for c, h in enumerate(headers, 1):
            self._hdr_cell(ws, 1, c, h)

        # Today's date formatted for ERP (DD-MM-YYYY)
        today_str = datetime.now().strftime("%d-%m-%Y")

        # ── Collect unique SO numbers while preserving file order ──
        # If multiple files have the same SO number (shouldn't happen), we deduplicate.
        # We keep the first row's meta data for each unique SO.
        seen = set()
        unique_sos = []
        for row in result.rows:
            if row.so_number not in seen:
                seen.add(row.so_number)
                unique_sos.append(row)  # Keep the first row's meta for this SO

        # ── Write one data row per unique SO number ──
        r = 2  # Start at row 2 (row 1 is the header)
        for row in unique_sos:
            self._data_cell(ws, r, 1, 'Order')              # Document Type = always 'Order'
            self._data_cell(ws, r, 2, row.so_number)         # No. = SO/GTM number
            self._data_cell(ws, r, 3, '')                    # Sell-to Customer No. = EMPTY (manual)
            self._data_cell(ws, r, 4, '')                    # Ship-to Code = EMPTY (manual)
            self._data_cell(ws, r, 5, today_str)             # Posting Date = today
            self._data_cell(ws, r, 6, today_str)             # Order Date = today
            self._data_cell(ws, r, 7, today_str)             # Document Date = today
            self._data_cell(ws, r, 8, today_str)             # Invoice From Date = today
            self._data_cell(ws, r, 9, today_str)             # Invoice To Date = today
            self._data_cell(ws, r, 10, row.so_number)        # External Document No. = same as No.
            self._data_cell(ws, r, 11, row.location_code)    # Location Code = mapped (e.g., 'PICK')
            self._data_cell(ws, r, 12, '')                   # Dimension Set ID = EMPTY
            self._data_cell(ws, r, 13, 'B2B')                # Supply Type = always 'B2B'
            # Columns 14-18: dimension placeholders left empty for ERP to fill
            r += 1

        self._auto_width(ws)

    def _write_lines_so(self, wb, result: ProcessingResult):
        """
        Sheet 2: 'Lines (SO)' — One row per ordered item.

        This sheet is imported into Business Central ERP as Sales Order lines.
        Each row adds one item line to its parent Sales Order.

        Line No. increments by 10000 and RESETS per new SO number:
            SO/GTM/6448: Line 10000, 20000, 30000...
            SO/GTM/6449: Line 10000, 20000, 30000... (reset)

        Columns:
            Document Type  = 'Order' (always)
            Document No.   = SO/GTM number (must match a header's No. value)
            Line No.       = 10000, 20000, 30000... (ERP sequential line ID)
            Type           = 'Item' (always — we're ordering stock items)
            No.            = BC Code / Item No from ERP master data
            Location Code  = mapped from file (same as header)
            Quantity       = Order Qty ONLY (tester qty tracked in Sales Header sheet)
            Unit Price     = EMPTY (ERP auto-fetches from item card pricing)
        """
        ws = wb.create_sheet('Lines (SO)')
        headers = ['Document Type', 'Document No.', 'Line No.', 'Type',
                   'No.', 'Location Code', 'Quantity', 'Unit Price']
        for c, h in enumerate(headers, 1):
            self._hdr_cell(ws, 1, c, h)

        r = 2              # Current output row (starts at 2, after header)
        current_so = None  # Track which SO we're on for line number reset
        line_no = 0        # Line number counter (resets per SO)

        for row in result.rows:
            # ── Reset line counter when SO number changes ──
            # Each Sales Order has its own sequence: 10000, 20000, 30000...
            if row.so_number != current_so:
                current_so = row.so_number
                line_no = 0  # Reset to 0 (will become 10000 after increment below)

            line_no += 10000  # Increment by 10K (Business Central ERP convention)

            self._data_cell(ws, r, 1, 'Order')              # Document Type
            self._data_cell(ws, r, 2, row.so_number)         # Document No. (links to header)
            self._data_cell(ws, r, 3, line_no)               # Line No. (10K increments)
            self._data_cell(ws, r, 4, 'Item')                # Type = always 'Item'
            self._data_cell(ws, r, 5, row.item_no)           # No. = BC Code / Item No
            self._data_cell(ws, r, 6, row.location_code)     # Location Code (mapped)
            self._data_cell(ws, r, 7, row.qty)               # Quantity = Order Qty ONLY
            self._data_cell(ws, r, 8, '')                    # Unit Price = EMPTY (ERP handles)
            r += 1

        self._auto_width(ws)

    def _write_sales_lines(self, wb, result: ProcessingResult):
        """
        Sheet 3: 'Sales Lines' — Detailed flat list of all ordered items.

        This is a reference sheet for verifying what was extracted from each file.
        Includes product details (EAN, Category, Description) for easy identification.
        NOT used for ERP import — purely for human review and cross-checking.

        Columns: SO Number | EAN | BC Code | Category | Article Description | Order Qty | Tester Qty
        """
        ws = wb.create_sheet('Sales Lines')
        headers = ['SO Number', 'EAN', 'BC Code', 'Category', 'Article Description',
                   'Order Qty', 'Tester Qty']
        for c, h in enumerate(headers, 1):
            self._hdr_cell(ws, 1, c, h)

        # Write one row per OrderRow (no grouping, no deduplication)
        for r, row in enumerate(result.rows, 2):
            self._data_cell(ws, r, 1, row.so_number)     # SO/GTM number
            self._data_cell(ws, r, 2, row.ean)            # EAN barcode
            self._data_cell(ws, r, 3, row.item_no)        # BC Code / Item No
            self._data_cell(ws, r, 4, row.category)       # Product category (Eye, FACE, etc.)
            self._data_cell(ws, r, 5, row.description)    # Article Description
            self._data_cell(ws, r, 6, row.qty)            # Order Qty
            self._data_cell(ws, r, 7, row.tester_qty)     # Tester Qty

        self._auto_width(ws)

    def _write_sales_header(self, wb, result: ProcessingResult):
        """
        Sheet 4: 'Sales Header' — Grouped summary per SO number.

        Shows aggregated totals for cross-verification against source files:
          - Order Qty:  total regular stock quantity across all items in this SO
          - Tester Qty: total tester/sample quantity across all items in this SO
          - Total Qty:  Order + Tester (for quick sanity check against file total)
          - Meta info:  Distributor, City, State, Location (from the file)

        One row per unique SO number. Used to verify before ERP import.
        """
        ws = wb.create_sheet('Sales Header')
        headers = ['SO Number', 'Order Qty', 'Tester Qty', 'Total Qty',
                   'Distributor', 'City', 'State', 'Location']
        for c, h in enumerate(headers, 1):
            self._hdr_cell(ws, 1, c, h)

        # ── Group rows by SO number ──
        # Aggregate quantities and keep the first row's meta fields for each SO
        so_groups: Dict[str, dict] = {}
        for row in result.rows:
            if row.so_number not in so_groups:
                # First time seeing this SO — initialize with meta from first row
                so_groups[row.so_number] = {
                    'order_qty': 0,
                    'tester_qty': 0,
                    'distributor': row.distributor,
                    'city': row.city,
                    'state': row.state,
                    'location': row.location,
                }
            # Accumulate quantities across all items belonging to this SO
            so_groups[row.so_number]['order_qty'] += row.qty
            so_groups[row.so_number]['tester_qty'] += row.tester_qty

        # ── Write one summary row per SO ──
        r = 2
        for so_num, info in so_groups.items():
            order_qty = info['order_qty']
            tester_qty = info['tester_qty']
            total_qty = order_qty + tester_qty   # Combined total for verification

            self._data_cell(ws, r, 1, so_num)            # SO Number
            self._data_cell(ws, r, 2, order_qty)          # Total Order Qty for this SO
            self._data_cell(ws, r, 3, tester_qty)         # Total Tester Qty for this SO
            self._data_cell(ws, r, 4, total_qty)          # Combined Total
            self._data_cell(ws, r, 5, info['distributor']) # Distributor Name
            self._data_cell(ws, r, 6, info['city'])        # City
            self._data_cell(ws, r, 7, info['state'])       # State
            self._data_cell(ws, r, 8, info['location'])    # Raw Location value
            r += 1

        self._auto_width(ws)

    def _write_sku_summary(self, wb, result: ProcessingResult):
        """
        Sheet 5: 'SKU Summary' — Pivot view of overall demand per item across ALL SOs.

        Aggregates quantities across all files/distributors to show total demand
        for each unique BC Code. Useful for warehouse/production planning.

        Example output:
            BC Code | Description                              | Category | Order Qty | Tester Qty | Total Qty
            200555  | RENEE Midnight Kajal Kohl Pencil 1.5 Gm | Eye      | 1000      | 50         | 1050
            200113  | RENEE Cover Up Hair Powder- Black 4g     | Hair     | 500       | 25         | 525

        Sorted by Total Qty descending (highest demand items at top).
        """
        ws = wb.create_sheet('SKU Summary')
        headers = ['BC Code', 'Description', 'Category', 'Order Qty', 'Tester Qty', 'Total Qty']
        for c, h in enumerate(headers, 1):
            self._hdr_cell(ws, 1, c, h)

        # ── Aggregate quantities per unique BC Code (item_no) ──
        # Keep the first description/category seen for each item
        sku_groups: Dict[str, dict] = {}
        for row in result.rows:
            if row.item_no not in sku_groups:
                sku_groups[row.item_no] = {
                    'description': row.description,    # First description found for this item
                    'category': row.category,          # First category found for this item
                    'order_qty': 0,                    # Accumulated order qty across all SOs
                    'tester_qty': 0,                   # Accumulated tester qty across all SOs
                }
            sku_groups[row.item_no]['order_qty'] += row.qty
            sku_groups[row.item_no]['tester_qty'] += row.tester_qty
            # Update description/category if previously blank (some files may have it, others not)
            if not sku_groups[row.item_no]['description'] and row.description:
                sku_groups[row.item_no]['description'] = row.description
            if not sku_groups[row.item_no]['category'] and row.category:
                sku_groups[row.item_no]['category'] = row.category

        # ── Sort by total qty descending (highest demand first) ──
        sorted_skus = sorted(
            sku_groups.items(),
            key=lambda x: x[1]['order_qty'] + x[1]['tester_qty'],
            reverse=True  # Highest total at top
        )

        # ── Write data rows ──
        r = 2
        grand_order = 0    # Grand total order qty
        grand_tester = 0   # Grand total tester qty
        for item_no, info in sorted_skus:
            total = info['order_qty'] + info['tester_qty']
            grand_order += info['order_qty']
            grand_tester += info['tester_qty']

            self._data_cell(ws, r, 1, item_no)              # BC Code
            self._data_cell(ws, r, 2, info['description'])   # Article Description
            self._data_cell(ws, r, 3, info['category'])      # Category
            self._data_cell(ws, r, 4, info['order_qty'])     # Total Order Qty across all SOs
            self._data_cell(ws, r, 5, info['tester_qty'])    # Total Tester Qty across all SOs
            self._data_cell(ws, r, 6, total)                 # Combined Total
            r += 1

        # ── Grand totals row (bold) ──
        grand_total = grand_order + grand_tester
        bold_font = Font(name='Aptos Display', size=11, bold=True)
        ws.cell(row=r, column=1, value='GRAND TOTAL').font = bold_font
        ws.cell(row=r, column=2, value=f'{len(sorted_skus)} unique SKUs').font = bold_font
        ws.cell(row=r, column=4, value=grand_order).font = bold_font
        ws.cell(row=r, column=5, value=grand_tester).font = bold_font
        ws.cell(row=r, column=6, value=grand_total).font = bold_font
        # Apply borders to totals row
        for c in range(1, 7):
            ws.cell(row=r, column=c).border = self.BORDER

        self._auto_width(ws)

    def _write_warnings(self, wb, result: ProcessingResult):
        """
        Sheet 6: 'Warnings' — Only created if any warnings/issues exist.

        Lists all warnings from all files so the user can review and fix.

        Common warnings:
          - "Distributor Name is blank — label not found or value is empty."
          - "SO number not found inside file (PO Number field is empty)."
          - "Column 'Tester Qty' not found — tester quantities will be 0."
          - "Distributor value 'South' looks like a state/zone name."

        Each warning shows which file it came from for easy tracking.
        """
        # Don't create the sheet at all if there are no warnings
        # (keeps the output clean when everything goes well)
        if not result.warned_files:
            return

        ws = wb.create_sheet('Warnings')
        headers = ['File', 'Warning']
        for c, h in enumerate(headers, 1):
            self._hdr_cell(ws, 1, c, h)

        for r, (fname, warning) in enumerate(result.warned_files, 2):
            self._data_cell(ws, r, 1, fname)    # Source filename that generated this warning
            self._data_cell(ws, r, 2, warning)   # Warning message describing the issue

        self._auto_width(ws)


# ═══════════════════════════════════════════════════════════════════════════════
#  FILE OPENER (cross-platform)
# ═══════════════════════════════════════════════════════════════════════════════
# Opens a file using the OS default application.
# Windows: os.startfile() — opens Excel for .xlsx files
# macOS:   'open' command
# Linux:   'xdg-open' command

def open_file(file_path: Path):
    """
    Open a file using the operating system's default application.

    On Windows: opens Excel (or whatever is associated with .xlsx)
    On macOS:   uses the 'open' command
    On Linux:   uses the 'xdg-open' command

    Shows an error popup if the file cannot be opened.

    Args:
        file_path: Path to the file to open
    """
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(str(file_path))           # Windows-specific file opener
        elif system == "Darwin":
            import subprocess as sp
            sp.Popen(["open", str(file_path)])     # macOS file opener
        else:
            import subprocess as sp
            sp.Popen(["xdg-open", str(file_path)]) # Linux file opener
    except Exception as e:
        messagebox.showerror("Open File Error", f"Could not open file:\n{e}")


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN AUTOMATION ENGINE
# ═══════════════════════════════════════════════════════════════════════════════
# Orchestrates the entire processing pipeline:
#   1. Takes a list of file paths from the GUI
#   2. Parses each file using ExcelParser
#   3. Collects all rows, warnings, and failures into ProcessingResult
#   4. The result is then passed to DumpExporter by the GUI

class GTMassAutomation:
    """Orchestrates file parsing and export for GT-Mass PO files."""

    def __init__(self):
        self.parser = ExcelParser()      # Parses individual files into OrderRows
        self.exporter = DumpExporter()   # Writes the output Excel workbook

    def process_files(self, files: List[Path]) -> ProcessingResult:
        """
        Process all selected files and return aggregated result.

        For each file:
          - Try to parse it → collect rows + warnings
          - If parsing fails (RuntimeError) → add to failed_files list
          - If unexpected error occurs → add to failed_files with error details

        Error handling strategy:
          - RuntimeError = known/expected failures (bad format, missing columns)
          - Exception = unexpected bugs (logged for debugging)
          - One bad file doesn't stop processing of other files

        Args:
            files: List of Path objects pointing to Excel files

        Returns:
            ProcessingResult containing all rows, warnings, and failures
        """
        result = ProcessingResult()

        for file in files:
            fname = file.name
            try:
                # Parse the file → get list of OrderRow + warnings
                rows, warnings = self.parser.parse(file)

                # Add all extracted rows to the combined result
                result.rows.extend(rows)

                # Add any warnings with the filename for context
                for w in warnings:
                    result.warned_files.append((fname, w))
                    logging.warning(f"{fname}: {w}")

            except RuntimeError as e:
                # Known error (file unreadable, missing columns, bad format)
                result.failed_files.append((fname, str(e)))
                logging.error(f"{fname} FAILED: {e}")

            except Exception as e:
                # Unexpected error (bug in code, weird data, etc.)
                # Log with full error for debugging
                result.failed_files.append((fname, f"Unexpected error: {e}"))
                logging.error(f"{fname} UNEXPECTED ERROR: {e}")

        # Log processing summary
        logging.info(
            f"Processing complete — "
            f"{len(result.rows)} rows | "
            f"{len(set(r.so_number for r in result.rows))} SO(s) | "
            f"{len(result.failed_files)} failed | "
            f"{len(result.warned_files)} warnings"
        )
        return result


# ═══════════════════════════════════════════════════════════════════════════════
#  TKINTER UI
# ═══════════════════════════════════════════════════════════════════════════════
# Simple desktop GUI with three buttons:
#   1. Select Excel Files  → opens file chooser dialog
#   2. Generate Dump       → processes files and creates output
#   3. Open Last Output    → opens the most recently generated file
#
# Also shows:
#   - Selected file count
#   - Processing status (waiting / processing / done / error)
#   - Time taken for processing

class AutomationUI:
    """Simple Tkinter GUI for file selection and dump generation."""

    def __init__(self, automation: GTMassAutomation):
        """
        Initialize the GUI window and all its widgets.

        Args:
            automation: GTMassAutomation engine instance for processing files
        """
        self.automation = automation                   # Reference to the processing engine
        self.files: List[Path] = []                    # Currently selected file paths
        self.last_output_path: Optional[Path] = None   # Path to most recent output file

        # ── Create the main application window ──
        self.root = tk.Tk()
        self.root.title("GT Mass Dump Generator v2.1")
        self.root.geometry("460x420")       # Fixed window size (width x height)
        self.root.resizable(False, False)   # Prevent window resizing

        # ── Title label (large bold text at top) ──
        tk.Label(
            self.root, text="GT Mass Dump Generator",
            font=("Arial", 14, "bold")
        ).pack(pady=10)

        # ── Subtitle (grey descriptive text below title) ──
        tk.Label(
            self.root, text="GT-Mass / Everyday PO Files → ERP Import (Headers + Lines)",
            font=("Arial", 9), fg="gray"
        ).pack(pady=0)

        # ── File count display (updates when files are selected) ──
        self.label = tk.Label(
            self.root, text="Selected Files: 0", font=("Arial", 10)
        )
        self.label.pack(pady=6)

        # ── Button: Select Excel Files ──
        tk.Button(
            self.root, text="Select Excel Files", width=22,
            command=self.select_files
        ).pack(pady=6)

        # ── Button: Generate Dump (main action) ──
        tk.Button(
            self.root, text="Generate Dump", width=22,
            command=self.generate_dump
        ).pack(pady=6)

        # ── Button: Open Last Output (disabled until first successful generation) ──
        self.open_button = tk.Button(
            self.root, text="Open Last Output File", width=22,
            state=tk.DISABLED,   # Starts disabled — enabled after first export
            command=self.open_last_file
        )
        self.open_button.pack(pady=6)

        # ── Button: Download PO Template (shows required format) ──
        tk.Button(
            self.root, text="📋 Download PO Template", width=22,
            command=self._download_template
        ).pack(pady=6)

        # ── Status label (shows current state: waiting / processing / done) ──
        self.status = tk.Label(
            self.root, text="Status: Waiting", font=("Arial", 10), fg="gray"
        )
        self.status.pack(pady=6)

        # ── Time label (shows processing duration after generation) ──
        self.time_label = tk.Label(
            self.root, text="", font=("Arial", 9), fg="darkgreen"
        )
        self.time_label.pack(pady=2)

    def select_files(self):
        """
        Open file chooser dialog to select GT-Mass / Everyday PO Excel files.

        Accepts .xlsx, .xls, and .xlsm files.
        Updates the file count label and resets status after selection.
        """
        # Open native OS file dialog (allows multi-select)
        files = filedialog.askopenfilenames(
            title="Select Sales Order Files",
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
        )
        # Convert selected file paths from strings to Path objects
        self.files = [Path(f) for f in files]
        # Update the display to show how many files were selected
        self.label.config(text=f"Selected Files: {len(self.files)}")
        self.time_label.config(text="")  # Clear previous time display
        self.status.config(text="Status: Files selected", fg="gray")

    def generate_dump(self):
        """
        Process all selected files and generate the output Excel dump.

        Flow:
          1. Validate that files are selected
          2. Start timer
          3. Process files via GTMassAutomation engine
          4. Export results via DumpExporter
          5. Display results summary (rows, SOs, warnings, failures)
          6. Offer to open the output file
        """
        # Guard: make sure files are selected before processing
        if not self.files:
            messagebox.showwarning("Warning", "Please select files first.")
            return

        # ── Start processing ──
        start_time = time.time()
        self.status.config(text="Status: Processing files...", fg="blue")
        self.time_label.config(text="")
        self.root.update()  # Force GUI refresh (prevents freeze during processing)

        # ── Process all selected files ──
        result = self.automation.process_files(self.files)
        # ── Export to Excel ──
        output_path = self.automation.exporter.export(result)

        # ── Calculate elapsed time ──
        elapsed = time.time() - start_time
        elapsed_str = f"{elapsed:.2f} seconds"

        # ── Collect stats for display ──
        failed = len(result.failed_files)    # Number of files that couldn't be read
        warned = len(result.warned_files)    # Number of warning messages
        rows = len(result.rows)              # Total OrderRow objects extracted
        sos = len(set(r.so_number for r in result.rows)) if result.rows else 0  # Unique SO count

        if output_path:
            # ── Success: output file was generated ──
            self.last_output_path = output_path
            self.open_button.config(state=tk.NORMAL)  # Enable "Open Last Output" button

            # Set status text and color based on whether there were issues
            if failed > 0 or warned > 0:
                self.status.config(
                    text=f"Done — {rows} rows | {failed} failed | {warned} warning(s)",
                    fg="orange"  # Orange = completed but with issues
                )
            else:
                self.status.config(
                    text=f"Done — {rows} rows across {sos} SO(s)",
                    fg="darkgreen"  # Green = clean success, no issues
                )

            self.time_label.config(text=f"⏱  Time taken: {elapsed_str}")

            # ── Build summary popup message ──
            warn_note = f"\n⚠️  {warned} warning(s) — check 'Warnings' sheet." if warned else ""
            fail_note = f"\n❌  {failed} file(s) failed — see error popup." if failed else ""

            # Ask user if they want to open the output file
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
            # ── No output: all files failed or had no data ──
            self.status.config(text="Status: No data to export", fg="red")
            self.time_label.config(text=f"⏱  Time taken: {elapsed_str}")

    def open_last_file(self):
        """
        Open the most recently generated output file.

        Shows a warning popup if the file has been deleted or moved since generation.
        """
        if self.last_output_path and self.last_output_path.exists():
            open_file(self.last_output_path)
        else:
            messagebox.showwarning(
                "File Not Found",
                "The output file no longer exists.\nPlease generate a new dump."
            )

    def _download_template(self):
        """
        Generate and save a blank GT-Mass PO template Excel file.

        The template matches the exact layout the script expects:
          - Meta header region (rows 1-6) with labels and placeholder values
          - Data header row (row 7) with all expected column names
          - A sample data row (row 8) showing the format

        The team can fill this template to ensure their PO files are
        always compatible with the script. No more format guessing.

        Template structure:
          Row 1: Title
          Row 2: Distributor Name | (value) | ... | ASM | (value)
          Row 3: DB Code          | (value) | ... | RSM | (value)
          Row 4: BDE Name         | (value) | ... | PO Number | SO/GTM/####
          Row 5: City             | (value) | ... | Date of PO | (date)
          Row 6: State            | (value) | ... | Location   | AHD/BLR
          Row 7: EAN | BC Code | Category | Article Description | ... | Order Qty | Tester Qty
          Row 8: (sample data)
        """
        save_path = filedialog.asksaveasfilename(
            title="Save GT-Mass PO Template",
            defaultextension=".xlsx",
            initialfile="GT-Mass_PO_Template.xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

            wb = Workbook()
            ws = wb.active
            ws.title = 'PO Template'

            # ── Formatting ──
            title_font = Font(name='Aptos Display', size=14, bold=True, color='1A237E')
            label_font = Font(name='Aptos Display', size=11, bold=True)
            value_font = Font(name='Aptos Display', size=11, color='0000CC')
            note_font = Font(name='Aptos Display', size=10, italic=True, color='FF6600')
            hdr_fill = PatternFill('solid', fgColor='1A237E')
            hdr_font = Font(bold=True, color='FFFFFF', name='Aptos Display', size=11)
            sample_font = Font(name='Aptos Display', size=11, color='888888', italic=True)
            meta_fill = PatternFill('solid', fgColor='E3F2FD')  # Light blue for meta region

            # ── Row 1: Title ──
            ws.cell(row=1, column=1, value='Purchase Order GT-Mass (Template)').font = title_font

            # ── Row 2: Distributor Name ──
            ws.cell(row=2, column=1, value='Distributor Name').font = label_font
            ws.cell(row=2, column=1).fill = meta_fill
            ws.cell(row=2, column=2, value='<Enter Distributor Name>').font = value_font
            ws.cell(row=2, column=7, value='ASM').font = label_font
            ws.cell(row=2, column=7).fill = meta_fill
            ws.cell(row=2, column=9, value='<ASM Name>').font = value_font

            # ── Row 3: DB Code ──
            ws.cell(row=3, column=1, value='DB Code').font = label_font
            ws.cell(row=3, column=1).fill = meta_fill
            ws.cell(row=3, column=2, value='<DB Code>').font = value_font
            ws.cell(row=3, column=7, value='RSM').font = label_font
            ws.cell(row=3, column=7).fill = meta_fill
            ws.cell(row=3, column=9, value='<RSM Name>').font = value_font

            # ── Row 4: BDE Name + PO Number (CRITICAL) ──
            ws.cell(row=4, column=1, value='BDE Name').font = label_font
            ws.cell(row=4, column=1).fill = meta_fill
            ws.cell(row=4, column=2, value='<BDE Name>').font = value_font
            ws.cell(row=4, column=7, value='PO Number').font = Font(name='Aptos Display', size=11, bold=True, color='D32F2F')
            ws.cell(row=4, column=7).fill = PatternFill('solid', fgColor='FFCDD2')  # Red highlight — critical field
            ws.cell(row=4, column=9, value='SO/GTM/0000').font = Font(name='Aptos Display', size=11, bold=True, color='D32F2F')
            ws.cell(row=4, column=9).fill = PatternFill('solid', fgColor='FFCDD2')

            # ── Row 5: City + Date of PO ──
            ws.cell(row=5, column=1, value='City').font = label_font
            ws.cell(row=5, column=1).fill = meta_fill
            ws.cell(row=5, column=2, value='<City>').font = value_font
            ws.cell(row=5, column=7, value='Date of PO').font = label_font
            ws.cell(row=5, column=7).fill = meta_fill
            ws.cell(row=5, column=9, value='DD.MM.YYYY').font = value_font

            # ── Row 6: State + Location (CRITICAL) ──
            ws.cell(row=6, column=1, value='State').font = label_font
            ws.cell(row=6, column=1).fill = meta_fill
            ws.cell(row=6, column=2, value='<State>').font = value_font
            ws.cell(row=6, column=7, value='Location').font = Font(name='Aptos Display', size=11, bold=True, color='D32F2F')
            ws.cell(row=6, column=7).fill = PatternFill('solid', fgColor='FFCDD2')
            ws.cell(row=6, column=9, value='AHD').font = Font(name='Aptos Display', size=11, bold=True, color='D32F2F')
            ws.cell(row=6, column=9).fill = PatternFill('solid', fgColor='FFCDD2')

            # ── Row 7: Data header row ──
            # These column names must match EXACTLY what the script expects
            data_headers = [
                'EAN', 'BC Code', 'Category', 'Article Description ',
                'Nail Paint Shade Number ', 'Product Classification',
                'HSN Code\n8 Digit', 'MRP', 'Retiler Margin',
                'Trade & Display Scheme', 'Ullage', 'QPS',
                'Qty In Case', 'Rate @ RLP', 'Amount @ RLP',
                'Order Qty', 'Order Amount', 'Tester Qty'
            ]
            # Critical columns that the script reads — highlighted in RED
            critical_cols = {'EAN', 'BC Code', 'Order Qty', 'Tester Qty'}
            critical_fill = PatternFill('solid', fgColor='D32F2F')  # Red background
            critical_font = Font(bold=True, color='FFFFFF', name='Aptos Display', size=11)

            for c, h in enumerate(data_headers, 1):
                cell = ws.cell(row=7, column=c, value=h)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                if h.strip() in critical_cols:
                    # RED highlight for critical columns the script parses
                    cell.font = critical_font
                    cell.fill = critical_fill
                else:
                    # Standard navy blue for non-critical columns
                    cell.font = hdr_font
                    cell.fill = hdr_fill

            # ── Row 8: Sample data row (grey italic — delete before use) ──
            sample_data = [
                8904473104307, 201238, 'Eye',
                'RENEE PURE BROWN KAJAL PEN WITH SHARPENER, 0.35GM',
                '-', 'Cosmetics', 33049990, 199, 1.2,
                '16.67% on RLP', '1.66 % on RLP', '4.81% on RLP',
                '', '', '', 72, '', 6
            ]
            for c, v in enumerate(sample_data, 1):
                cell = ws.cell(row=8, column=c, value=v)
                cell.font = sample_font

            # ── Row 10: Instructions note ──
            ws.cell(row=10, column=1,
                    value='⚠ INSTRUCTIONS:').font = Font(name='Aptos Display', size=11, bold=True, color='D32F2F')
            instructions = [
                '1. Fill PO Number (Row 4, Col I) with SO/GTM/#### — this is MANDATORY.',
                '2. Fill Location (Row 6, Col I) with AHD or BLR — this determines warehouse.',
                '3. Fill Distributor Name, City, State in the meta rows above.',
                '4. Add product data starting from Row 8. Delete the sample row first.',
                '5. BC Code must be numeric (Item No from ERP). Order Qty and Tester Qty must be numbers.',
                '6. Columns highlighted in RED are critical — the script reads SO Number and Location from them.',
                '7. Save as .xlsx and load into the GT Mass Dump Generator.',
            ]
            for i, instruction in enumerate(instructions):
                ws.cell(row=11 + i, column=1, value=instruction).font = note_font

            # ── Column widths ──
            widths = {
                'A': 16, 'B': 12, 'C': 12, 'D': 50, 'E': 12, 'F': 18,
                'G': 14, 'H': 8, 'I': 14, 'J': 20, 'K': 16, 'L': 14,
                'M': 12, 'N': 12, 'O': 14, 'P': 12, 'Q': 14, 'R': 12
            }
            for col_letter, w in widths.items():
                ws.column_dimensions[col_letter].width = w

            # Freeze at row 8 (data starts here)
            ws.freeze_panes = 'A8'

            wb.save(save_path)
            logging.info(f"Template saved: {save_path}")
            messagebox.showinfo(
                "Template Saved",
                f"GT-Mass PO template saved to:\n{save_path}\n\n"
                f"CRITICAL fields (highlighted in RED):\n"
                f"  • PO Number (Row 4, Col I) → SO/GTM/####\n"
                f"  • Location (Row 6, Col I) → AHD or BLR\n\n"
                f"Fill meta rows (Distributor, City, State),\n"
                f"add product data from Row 8 onwards,\n"
                f"delete the sample row, and save as .xlsx."
            )
        except Exception as e:
            logging.error(f"Template save failed: {e}")
            messagebox.showerror("Error", f"Failed to save template:\n{e}")

    def run(self):
        """
        Start the Tkinter event loop.

        This call BLOCKS until the user closes the window.
        All GUI interactions (button clicks, file dialogs) happen within this loop.
        """
        self.root.mainloop()


# ═══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════
# Application startup sequence:
#   1. check_expiry()       → block if expired, warn if expiring soon
#   2. GTMassAutomation()   → create the processing engine
#   3. AutomationUI()       → create the GUI window (with engine reference)
#   4. ui.run()             → start the event loop (shows window, blocks here)

def main():
    """Application entry point — called when script is run directly."""
    check_expiry()                          # Step 1: License/expiry check
    automation = GTMassAutomation()         # Step 2: Create processing engine
    ui = AutomationUI(automation)           # Step 3: Create GUI (passes engine)
    ui.run()                                # Step 4: Start event loop (blocks)


if __name__ == "__main__":
    main()