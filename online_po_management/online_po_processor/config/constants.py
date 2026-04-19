"""
config.constants
================

Pure constants used across the package. No business logic, no imports from
sibling modules — kept as a leaf so any other module can depend on it
without risk of import cycles.
"""

from __future__ import annotations

# ── Application expiry (hard-coded build-time deadline) ─────────────────────
# When this date passes, the GUI shows a popup and exits before doing anything
# else. Bumped manually each release cycle. The check itself lives in
# online_po_processor.app.check_expiry().
EXPIRY_DATE: str = "30-06-2026"


# ── Bundled-file folder + filenames ─────────────────────────────────────────
# The Items Master and Ship-To Mapping rarely change; instead of forcing the
# user to pick them on every run, the GUI looks for them in this folder next
# to the script. See online_po_processor.config.paths for the resolution
# helpers.

BUNDLED_DATA_FOLDER: str = "Calculation Data"
BUNDLED_MASTER_NAME: str = "Items March.xlsx"
BUNDLED_MAPPING_NAME: str = "Ship to B2B.xlsx"


# ── In-app update history (JSON sidecar) ────────────────────────────────────
# Tracks WHEN the user last clicked "Update Bundled Files" for each tracked
# file. Lives inside Calculation Data/ as a hidden file. Used by the GUI to
# show the small "Updated: 19-Apr-2026 18:41" sub-line under each picker row.
UPDATE_HISTORY_FILE: str = ".update_history.json"
