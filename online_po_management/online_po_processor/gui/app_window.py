"""
gui.app_window
==============

Main Tkinter window — ``OnlinePOApp``.

Layout (520×620 px)::

    ┌─────────────────────────────────────────┐
    │         Online PO Processor              │  ← title
    │  Marketplace PO → ERP Sales Order Import │
    │                                          │
    │ Marketplace: [Myntra ▼]  Margin: [70]%   │  ← mkt + margin row
    │                                          │
    │ ┌─ Input Files ─────────────────────┐   │
    │ │ Items Master:    ✓ Items March... │   │
    │ │                  Updated: …        │   │
    │ │ Ship-To Mapping: ✓ Ship to B2B... │   │
    │ │                  Updated: …        │   │
    │ │ Marketplace PO:  Not selected      │   │
    │ └────────────────────────────────────┘   │
    │                                          │
    │        [▶ Generate SO]                   │  ← actions
    │        [📂 Open Last Output]             │
    │        [📋 Download PO Template]         │
    │        [📁 Update Bundled Files]         │
    │                                          │
    │ Status: ...                              │
    │ ┌─ Log ──────────────────────────────┐   │
    │ │ [time] message                     │   │
    │ └────────────────────────────────────┘   │
    └─────────────────────────────────────────┘

Responsibilities
----------------
* Wire the UI and state together.
* Auto-load bundled master/mapping on startup.
* Route user actions to the engine/exporter/template-writer.
* Surface progress in the Log panel and Status line.

The class is intentionally "procedural inside a class" — it holds Tk
widget references plus a few StringVars and path strings. Business
logic lives in ``engine`` / ``exporter`` modules; this file is the
thin layer on top.
"""

from __future__ import annotations
import logging
import os
import shutil
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from online_po_processor.config.constants import (
    BUNDLED_DATA_FOLDER, BUNDLED_MAPPING_NAME, BUNDLED_MASTER_NAME,
)
from online_po_processor.config.marketplaces import (
    MARKETPLACE_CONFIGS, MARKETPLACE_NAMES,
)
from online_po_processor.config.paths import (
    get_bundled_data_folder, get_bundled_mapping_path,
    get_bundled_master_path, get_update_timestamp, record_update,
)
from online_po_processor.data.mapping_loader import MappingLoader
from online_po_processor.data.master_loader import MasterLoader
from online_po_processor.engine.marketplace_engine import MarketplaceEngine
from online_po_processor.exporter.so_exporter import SOExporter
from online_po_processor.gui._file_row import build_file_row
from online_po_processor.gui._update_dialog import UpdateDialog
from online_po_processor.utils.platform_open import open_file


class OnlinePOApp:
    """GUI for Online Marketplace PO → SO generation."""

    # ── Construction ───────────────────────────────────────────────────

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Online PO Processor — Marketplace SO Generator")
        self.root.geometry("520x620")
        self.root.resizable(False, False)

        # ── File paths (None until picked or auto-loaded) ───────────────
        self.master_path: Optional[str] = None
        self.mapping_path: Optional[str] = None
        self.po_path: Optional[str] = None

        # ── Output tracking ─────────────────────────────────────────────
        self.last_output: Optional[Path] = None

        # Track whether master/mapping came from the bundled folder (vs
        # user-picked). Used so the GUI can show "(auto-loaded)" and the
        # "Update Bundled Files" flow knows what's in use.
        self.master_is_bundled: bool = False
        self.mapping_is_bundled: bool = False

        # ── Engine-side state ───────────────────────────────────────────
        # MappingLoader is held on the app because it gets re-loaded each
        # run (when marketplace changes). MasterLoader is created fresh
        # in generate() — no state to carry between runs.
        self.mapping_loader = MappingLoader()
        self.exporter = SOExporter()

        # ── Widget references populated by _build_ui ────────────────────
        self.marketplace_var: tk.StringVar
        self.marketplace_dropdown: ttk.Combobox
        self.margin_var: tk.StringVar
        self.margin_entry: tk.Entry
        self.master_var: tk.StringVar
        self.master_ts_var: tk.StringVar
        self.mapping_var: tk.StringVar
        self.mapping_ts_var: tk.StringVar
        self.po_var: tk.StringVar
        self.open_btn: tk.Button
        self.status_var: tk.StringVar
        self.status_label: tk.Label
        self.log_text: tk.Text

        self._build_ui()

        # Auto-load AFTER the UI exists, so we can log and update
        # picker labels in one go.
        self._auto_load_bundled_files()

    # ── UI construction ────────────────────────────────────────────────

    def _build_ui(self) -> None:
        """Build the Tk widget tree."""

        # ── Title ───────────────────────────────────────────────────────
        tk.Label(
            self.root, text="Online PO Processor",
            font=("Arial", 14, "bold"),
        ).pack(pady=(12, 2))

        tk.Label(
            self.root, text="Marketplace PO → ERP Sales Order Import",
            font=("Arial", 9), fg='gray',
        ).pack(pady=(0, 10))

        # ── Marketplace selector + Margin input ─────────────────────────
        mkt_frame = tk.Frame(self.root)
        mkt_frame.pack(fill='x', padx=20, pady=(0, 8))

        tk.Label(
            mkt_frame, text="Marketplace:", font=("Arial", 10, "bold"),
        ).pack(side='left')

        self.marketplace_var = tk.StringVar(
            value=MARKETPLACE_NAMES[0] if MARKETPLACE_NAMES else ''
        )
        self.marketplace_dropdown = ttk.Combobox(
            mkt_frame, textvariable=self.marketplace_var,
            values=MARKETPLACE_NAMES, state='readonly', width=20,
        )
        self.marketplace_dropdown.pack(side='left', padx=8)
        self.marketplace_dropdown.bind(
            '<<ComboboxSelected>>', self._on_marketplace_change,
        )

        # Margin % — user can override per run (pre-filled from config)
        tk.Label(
            mkt_frame, text="Margin:", font=("Arial", 10, "bold"),
        ).pack(side='left', padx=(12, 0))
        self.margin_var = tk.StringVar(value=str(self._get_default_margin()))
        self.margin_entry = tk.Entry(
            mkt_frame, textvariable=self.margin_var, width=5,
            font=("Arial", 10), justify='center',
        )
        self.margin_entry.pack(side='left', padx=4)
        tk.Label(mkt_frame, text="%", font=("Arial", 10)).pack(side='left')
        tk.Label(
            mkt_frame, text="(Landing Cost)", font=("Arial", 8), fg='gray',
        ).pack(side='left', padx=4)

        # ── File selectors ──────────────────────────────────────────────
        files_frame = tk.LabelFrame(
            self.root, text="Input Files", font=("Arial", 10, "bold"),
            padx=10, pady=8,
        )
        files_frame.pack(fill='x', padx=20, pady=(0, 8))

        # Items Master (with timestamp sub-line)
        self.master_var = tk.StringVar(value="Not selected")
        self.master_ts_var = tk.StringVar(value="")
        build_file_row(
            files_frame, "Items Master:", self.master_var,
            self._select_master, ts_var=self.master_ts_var,
        )

        # Ship-To Mapping (with timestamp sub-line)
        self.mapping_var = tk.StringVar(value="Not selected")
        self.mapping_ts_var = tk.StringVar(value="")
        build_file_row(
            files_frame, "Ship-To Mapping:", self.mapping_var,
            self._select_mapping, ts_var=self.mapping_ts_var,
        )

        # Marketplace PO (no timestamp sub-line — per-run input)
        self.po_var = tk.StringVar(value="Not selected")
        build_file_row(
            files_frame, "Marketplace PO:", self.po_var, self._select_po,
        )

        # ── Action buttons ──────────────────────────────────────────────
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=8)

        tk.Button(
            btn_frame, text="▶  Generate SO", width=20,
            font=("Arial", 10, "bold"),
            bg="#00C853", fg='white', command=self.generate,
        ).pack(pady=4)

        self.open_btn = tk.Button(
            btn_frame, text="📂  Open Last Output", width=20,
            state=tk.DISABLED, command=self.open_last,
        )
        self.open_btn.pack(pady=4)

        tk.Button(
            btn_frame, text="📋  Download PO Template", width=20,
            command=self._download_template,
        ).pack(pady=4)

        tk.Button(
            btn_frame, text="📁  Update Bundled Files", width=20,
            command=self._update_bundled_files,
        ).pack(pady=4)

        # ── Status line ─────────────────────────────────────────────────
        self.status_var = tk.StringVar(
            value="Status: Waiting — select files and generate"
        )
        self.status_label = tk.Label(
            self.root, textvariable=self.status_var,
            font=("Arial", 10), fg='gray', wraplength=460,
        )
        self.status_label.pack(pady=6)

        # ── Log panel ───────────────────────────────────────────────────
        log_frame = tk.LabelFrame(self.root, text="Log", font=("Arial", 9))
        log_frame.pack(fill='both', expand=True, padx=20, pady=(0, 12))

        scroll = ttk.Scrollbar(log_frame, orient='vertical')
        scroll.pack(side='right', fill='y')

        self.log_text = tk.Text(
            log_frame, height=6, font=("Consolas", 9),
            state='disabled', wrap='word',
            yscrollcommand=scroll.set,
        )
        self.log_text.pack(fill='both', expand=True)
        scroll.config(command=self.log_text.yview)

    # ── Logging helpers ────────────────────────────────────────────────

    def _log(self, msg: str) -> None:
        """Append a timestamped message to the log panel."""
        self.log_text.config(state='normal')
        ts = time.strftime("%H:%M:%S")
        self.log_text.insert('end', f"[{ts}] {msg}\n")
        self.log_text.see('end')
        self.log_text.config(state='disabled')

    # ── Margin helpers ─────────────────────────────────────────────────

    def _get_default_margin(self) -> int:
        """Default margin % for the currently selected marketplace."""
        mkt = (self.marketplace_var.get()
               if hasattr(self, 'marketplace_var') else '')
        if mkt and mkt in MARKETPLACE_CONFIGS:
            return MARKETPLACE_CONFIGS[mkt].get('default_margin', 70)
        return 70

    def _on_marketplace_change(self, _event=None) -> None:
        """Reset margin to the newly-selected marketplace's default."""
        margin = self._get_default_margin()
        self.margin_var.set(str(margin))
        self._log(f"Marketplace changed to {self.marketplace_var.get()}, "
                  f"margin set to {margin}%")

    def _get_margin(self) -> float:
        """
        Current margin as a decimal (e.g. ``70`` → ``0.70``).

        Falls back to the marketplace default if the input field is
        empty or invalid. Valid range: 1..100 (inclusive).
        """
        try:
            val = float(self.margin_var.get().strip())
            if val <= 0 or val > 100:
                raise ValueError
            return val / 100.0
        except (ValueError, AttributeError):
            default = self._get_default_margin()
            self._log(f"Invalid margin input, using default {default}%")
            return default / 100.0

    # ── Bundled-file handling ─────────────────────────────────────────
    #
    # Items Master and Ship-To Mapping live in ``Calculation Data/``.
    # Startup auto-loads them so the user doesn't re-pick every run.
    # The "Update Bundled Files" button replaces what's in that folder.

    def _auto_load_bundled_files(self) -> None:
        """
        Look for Items Master + Ship-To Mapping in ``Calculation Data/``
        and pre-populate the picker labels if found.

        Does not abort startup on missing files — logs a hint and leaves
        the pickers in their default "Not selected" state.
        """
        master_p = get_bundled_master_path()
        mapping_p = get_bundled_mapping_path()

        if master_p:
            self.master_path = str(master_p)
            self.master_is_bundled = True
            self.master_var.set(f"✓ {master_p.name} (auto-loaded)")
            self._refresh_ts_label(self.master_ts_var, master_p.name)
            self._log(f"Auto-loaded master from "
                      f"{BUNDLED_DATA_FOLDER}/{master_p.name}")
        else:
            self._log(f"No bundled master at "
                      f"{BUNDLED_DATA_FOLDER}/{BUNDLED_MASTER_NAME} "
                      f"— pick one manually or use 'Update Bundled Files'")

        if mapping_p:
            self.mapping_path = str(mapping_p)
            self.mapping_is_bundled = True
            self.mapping_var.set(f"✓ {mapping_p.name} (auto-loaded)")
            self._refresh_ts_label(self.mapping_ts_var, mapping_p.name)
            self._log(f"Auto-loaded mapping from "
                      f"{BUNDLED_DATA_FOLDER}/{mapping_p.name}")
        else:
            self._log(f"No bundled mapping at "
                      f"{BUNDLED_DATA_FOLDER}/{BUNDLED_MAPPING_NAME} "
                      f"— pick one manually or use 'Update Bundled Files'")

    def _update_bundled_files(self) -> None:
        """
        Replace the bundled master and/or mapping in ``Calculation Data/``.

        Workflow:

        1. Ask which file(s) to update via :class:`UpdateDialog`.
        2. For each chosen kind, open a file picker and copy the picked
           file into the bundled folder under the canonical name.
        3. Refresh in-memory paths and picker labels.
        4. Log the outcome.
        """
        target_folder = get_bundled_data_folder(create=True)

        dialog = UpdateDialog(self.root, folder=target_folder)
        choice = dialog.show()
        if choice is None:
            return  # user cancelled

        updated_any = False

        if choice in ('master', 'both'):
            updated_any |= self._do_update_one_bundled(
                kind_label='Items Master',
                source_title='Select new Items Master file to bundle',
                target_path=target_folder / BUNDLED_MASTER_NAME,
                on_success=self._refresh_master_after_update,
            )

        if choice in ('mapping', 'both'):
            updated_any |= self._do_update_one_bundled(
                kind_label='Ship-To Mapping',
                source_title='Select new Ship-To Mapping file to bundle',
                target_path=target_folder / BUNDLED_MAPPING_NAME,
                on_success=self._refresh_mapping_after_update,
            )

        if updated_any:
            messagebox.showinfo(
                "Bundled Files Updated",
                f"Bundled files updated in:\n{target_folder}\n\n"
                f"Future runs will auto-load the new version.",
            )

    def _do_update_one_bundled(self, kind_label: str, source_title: str,
                                target_path: Path, on_success) -> bool:
        """
        Prompt for a source file and copy it to ``target_path``.

        Args:
            kind_label:   Display label (used in log/dialog text).
            source_title: Title of the file-picker dialog.
            target_path:  Destination (e.g.
                          ``Calculation Data/Items March.xlsx``).
            on_success:   Callback to refresh GUI state after a
                          successful copy.

        Returns:
            True if a copy was performed, False if the user cancelled
            or the copy failed.
        """
        src = filedialog.askopenfilename(
            title=source_title,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not src:
            self._log(f"Update cancelled for {kind_label}")
            return False

        try:
            shutil.copy2(src, str(target_path))
            # Stamp history BEFORE refresh so the sub-line shows the new
            # timestamp immediately.
            record_update(target_path.name)
            self._log(f"Bundled {kind_label} updated → {target_path}")
            on_success()
            return True
        except Exception as e:  # noqa: BLE001 — surface ANY copy error
            self._log(f"ERROR copying {kind_label}: {e}")
            messagebox.showerror(
                "Update Failed",
                f"Could not copy {kind_label}:\n{e}",
            )
            return False

    def _refresh_ts_label(self, ts_var: tk.StringVar, filename: str) -> None:
        """
        Refresh a timestamp StringVar from the in-app update history.

        Sets to ``"Updated: <date>"`` when there's a record, empty
        string otherwise (which renders as a blank sub-line).
        """
        ts = get_update_timestamp(filename)
        ts_var.set(f"Updated: {ts}" if ts else "")

    def _refresh_master_after_update(self) -> None:
        """Re-point in-memory master to the freshly bundled file."""
        p = get_bundled_master_path()
        if p:
            self.master_path = str(p)
            self.master_is_bundled = True
            self.master_var.set(f"✓ {p.name} (auto-loaded)")
            self._refresh_ts_label(self.master_ts_var, p.name)

    def _refresh_mapping_after_update(self) -> None:
        """Re-point in-memory mapping to the freshly bundled file."""
        p = get_bundled_mapping_path()
        if p:
            self.mapping_path = str(p)
            self.mapping_is_bundled = True
            self.mapping_var.set(f"✓ {p.name} (auto-loaded)")
            self._refresh_ts_label(self.mapping_ts_var, p.name)

    # ── Manual file pickers ────────────────────────────────────────────

    def _select_master(self) -> None:
        """
        Manually pick an Items Master file.

        Marks master as user-picked (not bundled); the bundled file in
        ``Calculation Data/`` is NOT touched — use "Update Bundled Files"
        for that.
        """
        path = filedialog.askopenfilename(
            title="Select Items Master file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.master_path = path
            self.master_is_bundled = False
            self.master_var.set(os.path.basename(path))
            # Clear the bundled timestamp — manual picks aren't tracked.
            self.master_ts_var.set("")
            self._log(f"Master (manual override): {os.path.basename(path)}")

    def _select_mapping(self) -> None:
        """
        Manually pick a Ship-To B2B mapping file. Bundled file untouched.
        """
        path = filedialog.askopenfilename(
            title="Select Mapping File (Ship-To B2B)",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.mapping_path = path
            self.mapping_is_bundled = False
            self.mapping_var.set(os.path.basename(path))
            self.mapping_ts_var.set("")
            self._log(f"Mapping (manual override): {os.path.basename(path)}")

    def _select_po(self) -> None:
        """Pick the marketplace PO/punch file for this run."""
        path = filedialog.askopenfilename(
            title="Select Marketplace PO File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.po_path = path
            self.po_var.set(os.path.basename(path))
            self._log(f"PO file: {os.path.basename(path)}")

    # ── Main processing flow ───────────────────────────────────────────

    def generate(self) -> None:
        """Main action: load mapping → parse PO → generate output."""
        marketplace = self.marketplace_var.get()
        if not marketplace or marketplace not in MARKETPLACE_CONFIGS:
            messagebox.showwarning(
                "No Marketplace", "Please select a marketplace.",
            )
            return

        if not self.mapping_path:
            messagebox.showwarning(
                "No Mapping", "Please select the Ship-To mapping file.",
            )
            return

        if not self.po_path:
            messagebox.showwarning(
                "No PO File", "Please select the marketplace PO file.",
            )
            return

        config = MARKETPLACE_CONFIGS[marketplace]
        margin_pct = self._get_margin()
        start_time = time.time()

        self.status_var.set("Processing...")
        self.status_label.config(fg='blue')
        self.root.update()

        self._log(f"Marketplace: {marketplace} | "
                  f"Margin: {int(margin_pct * 100)}%")

        # ── Load mapping for this marketplace ───────────────────────────
        self._log(f"Loading mapping for '{marketplace}'...")
        warnings: List[Tuple[str, str, str]] = []
        loc_count = self.mapping_loader.load(
            self.mapping_path, config['party_name'], warnings,
        )

        if loc_count == 0:
            self._log("ERROR: No mapping locations found!")
            for _, _, msg in warnings:
                self._log(f"  {msg}")
            self.status_var.set("Failed — mapping load error")
            self.status_label.config(fg='red')
            return

        self._log(f"Loaded {loc_count} locations for {marketplace}")

        # ── Load Items_March (master) ───────────────────────────────────
        master_loader: Optional[MasterLoader] = None
        if self.master_path:
            self._log("Loading Items_March for validation...")
            master_loader = MasterLoader()
            try:
                item_count = master_loader.load(self.master_path)
                self._log(f"Loaded {item_count:,} items from master")
            except Exception as e:  # noqa: BLE001
                self._log(f"WARNING: Master load failed: {e} "
                          f"— skipping validation")
                master_loader = None

        # ── Engine run ──────────────────────────────────────────────────
        self._log(f"Processing {os.path.basename(self.po_path)}...")
        engine = MarketplaceEngine(self.mapping_loader, master=master_loader)
        result = engine.process(self.po_path, config, margin_pct=margin_pct)
        result.margin_pct = margin_pct  # redundant but explicit

        if not result.rows:
            self._log("ERROR: No valid rows extracted!")
            for _, _, msg in result.warnings:
                self._log(f"  WARNING: {msg}")
            self.status_var.set("Failed — no data extracted")
            self.status_label.config(fg='red')
            return

        # ── Log summary ─────────────────────────────────────────────────
        unique_pos = {r.po_number for r in result.rows}
        total_qty = sum(r.qty for r in result.rows)

        self._log(f"Extracted: {len(result.rows)} items, "
                  f"{len(unique_pos)} PO(s), {total_qty} total qty")
        if result.warnings:
            self._log(f"Warnings: {len(result.warnings)}")
            for po, _loc, msg in result.warnings[:5]:
                self._log(f"  [{po}] {msg}")
            if len(result.warnings) > 5:
                self._log(f"  ... and {len(result.warnings) - 5} more "
                          f"(see Warnings sheet)")

        # ── Export ──────────────────────────────────────────────────────
        self._log("Writing output...")
        output_path = self.exporter.export(result)

        elapsed = time.time() - start_time

        if output_path:
            self.last_output = output_path
            self.open_btn.config(state=tk.NORMAL)

            status_msg = (f"Done — {len(result.rows)} items, "
                          f"{len(unique_pos)} PO(s), "
                          f"{total_qty} qty | {elapsed:.2f}s")
            if result.warnings:
                status_msg += f" | {len(result.warnings)} warning(s)"
                self.status_label.config(fg='orange')
            else:
                self.status_label.config(fg='darkgreen')

            self.status_var.set(status_msg)
            self._log(f"Saved: {output_path}")

            answer = messagebox.askyesno(
                "SO Generated",
                f"Sales Order generated successfully!\n\n"
                f"Marketplace : {marketplace}\n"
                f"PO(s)       : {len(unique_pos)}\n"
                f"Items       : {len(result.rows)}\n"
                f"Total Qty   : {total_qty}\n"
                f"Warnings    : {len(result.warnings)}\n"
                f"Time        : {elapsed:.2f}s\n\n"
                f"Do you want to open the output file?",
            )
            if answer:
                open_file(output_path)
        else:
            self.status_var.set("Failed — no output generated")
            self.status_label.config(fg='red')

    def open_last(self) -> None:
        """Open the last generated output file in the default app."""
        if self.last_output and self.last_output.exists():
            open_file(self.last_output)
        else:
            messagebox.showwarning("Not Found", "Output file not found.")

    # ── PO template download ──────────────────────────────────────────

    def _download_template(self) -> None:
        """
        Generate a blank PO template for the selected marketplace.

        Headers are colour-coded so the user knows which columns to
        actually fill in:

        * **BLUE** (``#1A237E``) — Required. Script fails without these.
        * **GREEN** (``#1B5E20``) — Validation. Used for price check +
          master lookup.
        * **GREY** (``#9E9E9E``) — Not read by script. Kept only to
          mirror the marketplace's native file format.

        Includes a 3-row legend below the header explaining each colour
        plus a final orange instruction line.

        The required set adapts to ``item_resolution``:
          * ``from_column`` → item_col is the required identifier.
          * ``from_ean`` → ean_col is the required identifier
            (promoted from GREEN to BLUE).
        """
        marketplace = self.marketplace_var.get()
        if not marketplace or marketplace not in MARKETPLACE_CONFIGS:
            messagebox.showwarning(
                "No Marketplace", "Please select a marketplace first.",
            )
            return

        config = MARKETPLACE_CONFIGS[marketplace]

        save_path = filedialog.asksaveasfilename(
            title=f"Save {marketplace} PO Template",
            defaultextension=".xlsx",
            initialfile=f"{marketplace}_PO_Template.xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not save_path:
            return

        try:
            self._write_template_workbook(save_path, marketplace, config)
            self._log(f"{marketplace} template saved → {save_path}")
            messagebox.showinfo(
                "Template Saved",
                f"{marketplace} PO template saved to:\n{save_path}\n\n"
                f"Header colours:\n"
                f"  • Blue  = Required (must fill)\n"
                f"  • Green = Validation (recommended)\n"
                f"  • Grey  = Not read by script",
            )
        except Exception as e:  # noqa: BLE001
            self._log(f"Template save failed: {e}")
            messagebox.showerror(
                "Error", f"Failed to save template:\n{e}",
            )

    @staticmethod
    def _write_template_workbook(save_path: str, marketplace: str,
                                   config: dict) -> None:
        """
        Actually build + save the PO template workbook.

        Split out of :meth:`_download_template` so the wrapping
        save-dialog / success-messagebox flow stays short and readable.
        """
        wb = Workbook()
        ws = wb.active
        ws.title = f'{marketplace} PO'

        # ── Determine required vs validation vs unused cols ─────────────
        item_resolution = config.get('item_resolution', 'from_column')

        required_cols = {
            config['po_col'], config['loc_col'], config['qty_col'],
        }
        if item_resolution == 'from_ean':
            # EAN is the required identifier when Item No is resolved
            # from it — promote from GREEN to BLUE.
            if config.get('ean_col'):
                required_cols.add(config['ean_col'])
        else:  # 'from_column'
            if config.get('item_col'):
                required_cols.add(config['item_col'])

        validation_cols = set()
        if config.get('ean_col') and config['ean_col'] not in required_cols:
            validation_cols.add(config['ean_col'])
        if config.get('fob_col'):
            validation_cols.add(config['fob_col'])

        # ── Build column list ───────────────────────────────────────────
        # Prefer the full marketplace template_headers; fall back to a
        # minimal list built from required + validation columns.
        headers = config.get('template_headers')
        if not headers:
            headers = [config['po_col'], config['loc_col']]
            if (item_resolution == 'from_column'
                    and config.get('item_col')):
                headers.append(config['item_col'])
            headers.append(config['qty_col'])
            if config.get('ean_col') and config['ean_col'] not in headers:
                headers.append(config['ean_col'])
            if config.get('fob_col') and config['fob_col'] not in headers:
                headers.append(config['fob_col'])

        # ── Styles per role ─────────────────────────────────────────────
        required_fill = PatternFill('solid', fgColor='1A237E')   # blue
        validation_fill = PatternFill('solid', fgColor='1B5E20')  # green
        unused_fill = PatternFill('solid', fgColor='9E9E9E')       # grey

        hdr_font_white = Font(bold=True, color='FFFFFF',
                               name='Aptos Display', size=11)
        hdr_font_dim = Font(bold=True, color='EEEEEE',
                             name='Aptos Display', size=11, italic=True)

        # ── Header row (colour-coded) ───────────────────────────────────
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=c, value=h)
            if h in required_cols:
                cell.fill = required_fill
                cell.font = hdr_font_white
            elif h in validation_cols:
                cell.fill = validation_fill
                cell.font = hdr_font_white
            else:
                cell.fill = unused_fill
                cell.font = hdr_font_dim
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[get_column_letter(c)].width = max(
                len(h) + 4, 12,
            )

        # ── Legend rows (3-5) ───────────────────────────────────────────
        legend_row = 3
        legend_items = [
            ('1A237E', 'FFFFFF', 'REQUIRED',
             f'Script fails without these — fill them in: '
             f'{", ".join(sorted(required_cols))}'),
            ('1B5E20', 'FFFFFF', 'VALIDATION',
             f'Used for price check & master lookup: '
             f'{", ".join(sorted(validation_cols)) or "(none)"}'),
            ('9E9E9E', 'FFFFFF', 'NOT READ',
             'Optional — kept only to match marketplace file format; '
             'can stay blank'),
        ]
        for fg, fc, label, desc in legend_items:
            tag = ws.cell(row=legend_row, column=1, value=label)
            tag.fill = PatternFill('solid', fgColor=fg)
            tag.font = Font(bold=True, color=fc,
                             name='Aptos Display', size=10)
            tag.alignment = Alignment(horizontal='center')

            desc_cell = ws.cell(row=legend_row, column=2, value=desc)
            desc_cell.font = Font(name='Aptos Display', size=10,
                                    color='333333', italic=True)
            ws.merge_cells(start_row=legend_row, start_column=2,
                            end_row=legend_row,
                            end_column=min(8, len(headers)))
            legend_row += 1

        # ── Final orange instruction row ────────────────────────────────
        ws.cell(
            row=legend_row + 1, column=1,
            value=(f'← {marketplace} PO template. Fill data rows below '
                   f'the header. Only the BLUE & GREEN columns are read '
                   f'by the script.'),
        ).font = Font(name='Aptos Display', size=10,
                      color='FF6600', italic=True)

        ws.freeze_panes = 'A2'
        wb.save(save_path)

    # ── Run the app ────────────────────────────────────────────────────

    def run(self) -> None:
        """Start the Tkinter main loop. Blocks until the window closes."""
        self.root.mainloop()
