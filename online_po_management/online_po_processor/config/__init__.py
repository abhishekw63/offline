"""
config — constants, marketplace registry, and filesystem path helpers.

Public re-exports
-----------------
Anything imported below can be used as ``from online_po_processor.config
import X`` instead of the deeper ``from .marketplaces import X`` form.
"""

from online_po_processor.config.constants import (
    EXPIRY_DATE,
    BUNDLED_DATA_FOLDER,
    BUNDLED_MASTER_NAME,
    BUNDLED_MAPPING_NAME,
    UPDATE_HISTORY_FILE,
)
from online_po_processor.config.marketplaces import (
    MARKETPLACE_CONFIGS,
    MARKETPLACE_NAMES,
)
from online_po_processor.config.paths import (
    get_bundled_master_path,
    get_bundled_mapping_path,
    get_bundled_data_folder,
    load_update_history,
    record_update,
    get_update_timestamp,
)

__all__ = [
    # constants
    'EXPIRY_DATE',
    'BUNDLED_DATA_FOLDER',
    'BUNDLED_MASTER_NAME',
    'BUNDLED_MAPPING_NAME',
    'UPDATE_HISTORY_FILE',
    # marketplaces
    'MARKETPLACE_CONFIGS',
    'MARKETPLACE_NAMES',
    # paths
    'get_bundled_master_path',
    'get_bundled_mapping_path',
    'get_bundled_data_folder',
    'load_update_history',
    'record_update',
    'get_update_timestamp',
]
