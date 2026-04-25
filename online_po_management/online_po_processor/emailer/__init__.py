"""
emailer — HTML email composition and SMTP delivery for online-PO reports.

Two modules
-----------
``email_builder``
    Pure transformation: ``ProcessingResult`` → HTML string + subject.
    No I/O. Fully testable without SMTP credentials.

``email_sender``
    Network layer: takes a pre-built HTML string and ships it via
    Gmail SMTP. Catches and surfaces specific SMTP/network errors so
    the GUI can show actionable messages.

The split follows the Single Responsibility Principle — mirrors the
``EmailBuilder`` + ``EmailSender`` split in GT Mass Dump.
"""

from online_po_processor.emailer.email_builder import EmailBuilder
from online_po_processor.emailer.email_sender import EmailSender

__all__ = ['EmailBuilder', 'EmailSender']