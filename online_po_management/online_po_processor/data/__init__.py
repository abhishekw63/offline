"""
data — pure-data classes and the file loaders that populate them.

Public re-exports
-----------------
``SORow``, ``ProcessingResult`` — dataclasses passed between engine and
exporter.
``MasterLoader``, ``MappingLoader`` — file → in-memory lookup loaders.
"""

from online_po_processor.data.models import SORow, ProcessingResult
from online_po_processor.data.master_loader import MasterLoader
from online_po_processor.data.mapping_loader import MappingLoader

__all__ = [
    'SORow',
    'ProcessingResult',
    'MasterLoader',
    'MappingLoader',
]
