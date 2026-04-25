"""
exporter — ``ProcessingResult`` → Excel output.

Two exporters, different destinations:

``SOExporter``
    Produces the main reference workbook with Headers, Lines, Summary,
    Validation, Raw Data, Warnings sheets. Used as the user-facing
    output for every "Generate SO" run.

``D365Exporter``
    Fills a Dynamics 365 sample package template with the same rows
    so the Headers + Lines tabs can be imported directly into the
    D365 Data Management framework. Triggered by the "Export D365
    Package" button on the GUI.
"""

from online_po_processor.exporter.d365_exporter import D365Exporter
from online_po_processor.exporter.so_exporter import SOExporter

__all__ = ['SOExporter', 'D365Exporter']