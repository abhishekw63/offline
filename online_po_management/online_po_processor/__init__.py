"""
online_po_processor
===================

Marketplace PO/punch file → ERP-importable Sales Order generator.

This package replaces the single-file ``standalone_po_processing.py`` script
(now retained as ``legacy_standalone.py`` for fallback). Logic is identical
through v1.4.0 — only the file layout changed.

Layout overview
---------------
::

    online_po_processor/
        config/      → constants, marketplace registry, paths/history helpers
        data/        → pure-data classes (SORow, ProcessingResult) and loaders
        engine/      → MarketplaceEngine — turns a punch file into result rows
        exporter/    → SOExporter + D365Exporter + per-sheet writers
        emailer/     → HTML report builder + SMTP sender
        gui/         → Tkinter UI (OnlinePOApp + dialogs)
        utils/       → cross-platform helpers (open_file)
        app.py       → bootstrap: expiry check + main() entry point

Quick start
-----------
The intended entry point is the top-level ``main.py`` in the project root::

    python main.py

That file does nothing more than::

    from online_po_processor.app import main
    main()

Public re-exports
-----------------
The most commonly imported names are exposed at package level for
convenience and to mirror what the legacy single-file module exported:

    >>> from online_po_processor import (
    ...     OnlinePOApp,
    ...     MARKETPLACE_CONFIGS,
    ...     SORow,
    ...     ProcessingResult,
    ... )
"""

__version__ = "1.9.3"
__all__ = [
    "__version__",
    # Re-exports for code that used to ``import standalone_po_processing as opp``
    "OnlinePOApp",
    "MARKETPLACE_CONFIGS",
    "MARKETPLACE_NAMES",
    "WAREHOUSE_CODES",
    "WAREHOUSE_DISPLAY_NAMES",
    "DEFAULT_WAREHOUSE",
    "SORow",
    "ProcessingResult",
    "MasterLoader",
    "MappingLoader",
    "MarketplaceEngine",
    "SOExporter",
    "D365Exporter",
    "EmailBuilder",
    "EmailSender",
    "get_email_config",
    "main",
]

# --- public re-exports (kept thin; all real code lives in submodules) -------
from online_po_processor.config.email_config import get_email_config
from online_po_processor.config.marketplaces import (
    DEFAULT_WAREHOUSE,
    MARKETPLACE_CONFIGS,
    MARKETPLACE_NAMES,
    WAREHOUSE_CODES,
    WAREHOUSE_DISPLAY_NAMES,
)
from online_po_processor.data.mapping_loader import MappingLoader
from online_po_processor.data.master_loader import MasterLoader
from online_po_processor.data.models import ProcessingResult, SORow
from online_po_processor.emailer import EmailBuilder, EmailSender
from online_po_processor.engine.marketplace_engine import MarketplaceEngine
from online_po_processor.exporter.d365_exporter import D365Exporter
from online_po_processor.exporter.so_exporter import SOExporter
from online_po_processor.gui.app_window import OnlinePOApp
from online_po_processor.app import main