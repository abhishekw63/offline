"""
main.py — top-level launcher for the Online PO Processor.

Intended usage::

    cd D:\\PO tracking\\Automation\\warehouse\\online_po_management
    python main.py

Or, once bundled with PyInstaller::

    online_po_processor.exe

This file is intentionally kept tiny — all real logic lives in the
``online_po_processor`` package. Keeping the launcher here means:

* PyInstaller has a clean entry point to target.
* The package itself is importable as a library (e.g. for tests or
  batch-processing scripts) without running the GUI.
* Anyone opening the project sees ``main.py`` at the top level and
  knows exactly where to start reading.

If you'd rather run the package directly without this shim, the
equivalent invocation is::

    python -m online_po_processor
"""

from online_po_processor.app import main


if __name__ == '__main__':
    main()
