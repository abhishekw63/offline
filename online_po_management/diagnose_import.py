"""Diagnose why 'OnlinePOApp' can't be imported."""
import sys
from pathlib import Path

here = Path(__file__).parent
app_window_path = here / 'online_po_processor' / 'gui' / 'app_window.py'

print(f'File exists: {app_window_path.exists()}')
print(f'File size: {app_window_path.stat().st_size} bytes')

# Check for the class line
text = app_window_path.read_text(encoding='utf-8', errors='replace')
lines = text.splitlines()
print(f'Total lines: {len(lines)}')

# Find class definition
for i, line in enumerate(lines, 1):
    if 'class OnlinePOApp' in line:
        print(f'✓ class OnlinePOApp found at line {i}')
        break
else:
    print('✗ class OnlinePOApp NOT FOUND — paste was truncated!')
    print(f'  Last 3 lines of file:')
    for line in lines[-3:]:
        print(f'    {line!r}')

# Look for import problem
for i, line in enumerate(lines, 1):
    if 'WAREHOUSE_CODES' in line or 'DEFAULT_WAREHOUSE' in line:
        if 'import' in line or 'from' in line:
            print(f'  line {i}: {line.strip()}')

# Try to compile the module directly
import py_compile
try:
    py_compile.compile(str(app_window_path), doraise=True)
    print('✓ File compiles cleanly')
except py_compile.PyCompileError as e:
    print(f'✗ Compile error: {e}')
    sys.exit(1)

# Try to import it (will fail if there's a runtime problem)
sys.path.insert(0, str(here))
try:
    # Clear any cached module
    for key in list(sys.modules):
        if 'online_po_processor' in key:
            del sys.modules[key]
    from online_po_processor.gui.app_window import OnlinePOApp
    print(f'✓ Import succeeded: {OnlinePOApp}')
except Exception as e:
    print(f'✗ Import failed: {type(e).__name__}: {e}')