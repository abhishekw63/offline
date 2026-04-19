"""
Enables ``python -m online_po_processor`` as an alternate launch form
alongside the top-level ``main.py``. Both delegate to the same
:func:`~online_po_processor.app.main`.
"""

from online_po_processor.app import main


if __name__ == '__main__':
    main()
