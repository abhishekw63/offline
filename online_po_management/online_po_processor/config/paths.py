"""
config.paths
============

Filesystem path helpers and the in-app update-history JSON sidecar.

This module is the **single source of truth** for "where do bundled files
live and when were they last updated by the user?" Other modules import
from here rather than constructing paths themselves, so a future move of
the bundled folder (or a switch to per-user data dirs) only touches this
file.
"""

from __future__ import annotations
import json
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

from online_po_processor.config.constants import (
    BUNDLED_DATA_FOLDER,
    BUNDLED_MASTER_NAME,
    BUNDLED_MAPPING_NAME,
    UPDATE_HISTORY_FILE,
)


# ────────────────────────────────────────────────────────────────────────────
#  Script-relative roots
# ────────────────────────────────────────────────────────────────────────────

def _script_dir() -> Path:
    """
    Return the directory the application is running from.

    Handles three cases:
      * Frozen PyInstaller exe → directory of ``sys.executable``
      * Normal ``python main.py`` run → directory of the launcher script
        (resolved by walking up from this module's file location)
      * Interactive / fallback → current working directory

    Note: in the legacy single-file version this used ``Path(__file__).parent``
    of the script itself. Now that the package lives one level deep, we walk
    up to find the project root (the directory CONTAINING the package).
    """
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent

    try:
        # /home/.../<project>/online_po_processor/config/paths.py
        # parents[0] = config
        # parents[1] = online_po_processor
        # parents[2] = <project root>      ← what we want
        return Path(__file__).resolve().parents[2]
    except (NameError, IndexError):
        return Path.cwd()


# ────────────────────────────────────────────────────────────────────────────
#  Bundled-file resolution
# ────────────────────────────────────────────────────────────────────────────
#
# Layout the GUI looks for on startup::
#
#     <project root>/
#         main.py
#         online_po_processor/...
#         Calculation Data/
#             Items March.xlsx     ← auto-loaded as Items Master
#             Ship to B2B.xlsx     ← auto-loaded as Ship-To Mapping
#
# If files are missing the GUI falls back to manual file pickers and shows
# a hint in the log panel. Nothing here raises on a missing file — the
# caller decides whether absence is fatal.

def get_bundled_master_path() -> Optional[Path]:
    """
    Path to the bundled Items Master, or ``None`` if not present.
    """
    p = _script_dir() / BUNDLED_DATA_FOLDER / BUNDLED_MASTER_NAME
    return p if p.exists() else None


def get_bundled_mapping_path() -> Optional[Path]:
    """
    Path to the bundled Ship-To Mapping, or ``None`` if not present.
    """
    p = _script_dir() / BUNDLED_DATA_FOLDER / BUNDLED_MAPPING_NAME
    return p if p.exists() else None


def get_bundled_data_folder(create: bool = False) -> Path:
    """
    Path to the ``Calculation Data/`` folder.

    Args:
        create: If True, create the folder if it doesn't exist. Used by the
                "Update Bundled Files" flow so the user can drop the first
                file in even before the folder exists on disk.
    """
    folder = _script_dir() / BUNDLED_DATA_FOLDER
    if create:
        folder.mkdir(parents=True, exist_ok=True)
    return folder


# ────────────────────────────────────────────────────────────────────────────
#  In-app update history (JSON sidecar)
# ────────────────────────────────────────────────────────────────────────────
#
# Format (Calculation Data/.update_history.json):
#
#     {
#         "Items March.xlsx": "2026-04-19T18:41:32",
#         "Ship to B2B.xlsx": "2026-04-17T11:22:05"
#     }
#
# Why this rather than file mtime?
#   * Tracks ONLY explicit user updates via the GUI — exactly what the user
#     asked for.
#   * Survives unrelated saves (someone opens the file in Excel without
#     modifying it; mtime would change but we don't want to lie about an
#     update having happened).

def _history_path() -> Path:
    """Path to the JSON sidecar (may not exist yet)."""
    return _script_dir() / BUNDLED_DATA_FOLDER / UPDATE_HISTORY_FILE


def load_update_history() -> Dict[str, str]:
    """
    Load the in-app update history.

    Returns:
        ``{filename: ISO-timestamp}`` mapping. Empty dict if the sidecar
        doesn't exist or is corrupt — never raises (history is best-effort
        metadata, must not block real operations).
    """
    p = _history_path()
    if not p.exists():
        return {}
    try:
        with open(p, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, OSError) as e:
        logging.warning("Could not read update history at %s: %s", p, e)
        return {}


def record_update(filename: str) -> None:
    """
    Stamp ``filename`` (e.g. ``'Items March.xlsx'``) as just-updated NOW.

    Creates the ``Calculation Data/`` folder if needed, then merges the new
    timestamp into the existing JSON. Failures are logged but never raised
    — losing a timestamp shouldn't block a real file copy.
    """
    folder = get_bundled_data_folder(create=True)
    history = load_update_history()
    history[filename] = datetime.now().isoformat(timespec='seconds')
    try:
        with open(folder / UPDATE_HISTORY_FILE, 'w', encoding='utf-8') as f:
            json.dump(history, f, indent=2)
    except OSError as e:
        logging.warning("Could not write update history: %s", e)


def get_update_timestamp(filename: str) -> Optional[str]:
    """
    Recorded update timestamp for ``filename`` formatted for display
    (e.g. ``'19-Apr-2026 18:41'``).

    Returns:
        Formatted string, or ``None`` if no record exists or the stored
        value couldn't be parsed.
    """
    history = load_update_history()
    iso = history.get(filename)
    if not iso:
        return None
    try:
        dt = datetime.fromisoformat(iso)
        return dt.strftime('%d-%b-%Y %H:%M')
    except ValueError:
        return None
