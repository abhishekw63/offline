"""Print version markers from each critical file so we can spot which
files are stale vs current."""
from pathlib import Path

ROOT = Path('online_po_processor')

CHECKS = [
    # (file path, marker string, expected version it was added in)
    ('__init__.py',                            '__version__',         '1.7.0 expected'),
    ('engine/marketplace_engine.py',           'def process_multi',    '1.7.0'),
    ('engine/marketplace_engine.py',           '_preprocess_reliance', '1.6.0'),
    ('engine/marketplace_engine.py',           '_check_hsn',           '1.6.0'),
    ('engine/marketplace_engine.py',           'apply_margin',         '1.6.0'),
    ('exporter/d365_exporter.py',              '_ERP_LOCATION_CODE',   '1.5.9'),
    ('exporter/sheets/raw_data_sheet.py',      'Source',               '1.7.0'),
    ('exporter/sheets/raw_data_sheet.py',      'resolved_config',      '1.5.6'),
    ('exporter/sheets/validation_sheet.py',    'has_hsn_check',        '1.6.0'),
    ('exporter/so_exporter.py',                'input_files_count',    '1.7.0'),
    ('emailer/email_builder.py',               'input_files_count',    '1.7.0'),
    ('data/models.py',                         'source_po',            '1.7.0'),
    ('data/models.py',                         'hsn_check_status',     '1.6.0'),
    ('data/master_loader.py',                  'HSN/SAC Code',         '1.6.0'),
    ('config/marketplaces.py',                 "'Reliance'",           '1.6.0'),
    ('gui/app_window.py',                      'po_paths',             '1.7.0'),
]

print(f'{"File":<45} {"Marker":<25} {"Status":<10} Expected')
print('-' * 100)
for filepath, marker, expected in CHECKS:
    p = ROOT / filepath
    if not p.exists():
        status = 'MISSING'
    else:
        text = p.read_text(encoding='utf-8', errors='replace')
        status = 'OK' if marker in text else 'STALE'
    print(f'{filepath:<45} {marker[:25]:<25} {status:<10} {expected}')

# Also print the first __version__ line from __init__.py
init_path = ROOT / '__init__.py'
if init_path.exists():
    for line in init_path.read_text(encoding='utf-8', errors='replace').splitlines():
        if '__version__' in line and '=' in line:
            print(f'\nVersion in __init__.py: {line.strip()}')
            break