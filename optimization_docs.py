"""
=============================================================================
             PANDAS DATAFRAME ITERATION OPTIMIZATION
=============================================================================

This document outlines the recent refactoring made to optimize Pandas
DataFrame iteration within our standalone scripts and offline utils.

Target Branches:
- refactor-optimize-pandas-iteration-10618749738622278764
- refactor-pandas-iterrows-optimization

─────────────────────────────────────────────────────────────────────────────
 1. THE PROBLEM: `df.iterrows()`
─────────────────────────────────────────────────────────────────────────────
EARLIER CODE:
```python
for i, row in df.iterrows():
    bc_code = row['bc code']
    qty = row['order qty']
    # ... process data ...
```

WHY IT WAS FLAWED:
Using `.iterrows()` in Pandas is a well-known anti-pattern for large datasets.
When you call `.iterrows()`, Pandas creates a brand new `pd.Series` object
for *every single row* in the DataFrame.
Creating these objects involves significant memory allocation and type-checking
overhead, causing loops to run extremely slowly when processing hundreds or
thousands of rows in large Excel files.

─────────────────────────────────────────────────────────────────────────────
 2. THE SOLUTION: `.values` array iteration
─────────────────────────────────────────────────────────────────────────────
NEW CODE:
```python
# 1. Look up the integer indices of the columns BEFORE the loop starts
bc_idx = df.columns.get_loc('bc code')
qty_idx = df.columns.get_loc('order qty')

# 2. Iterate over the pure NumPy array underlying the DataFrame
for row_vals in df.values:
    bc_code = row_vals[bc_idx]
    qty = row_vals[qty_idx]
    # ... process data ...
```

IMPACT:
By dropping down to `df.values`, we iterate over a pure C-based NumPy array.
This bypasses all Pandas Series creation overhead.
- Speed: This approach is generally 50x to 100x faster than `.iterrows()`.
- Memory: It consumes far less memory during the loop.
- Safety: By pre-calculating indices (`df.columns.get_loc`), we keep the safe,
  name-based lookups but pay the cost only once instead of per-row.

─────────────────────────────────────────────────────────────────────────────
 3. EXCEPTION HANDLING REFACTOR (Bonus)
─────────────────────────────────────────────────────────────────────────────
EARLIER CODE:
```python
try:
    bc_code = int(bc_code)
except:
    continue
```

NEW CODE:
```python
try:
    bc_code = int(bc_code)
except (ValueError, TypeError):
    # If BC Code is not numeric or convertible, skip the row
    continue
```

IMPACT:
A "bare except" (`except:`) is dangerous because it catches EVERYTHING, including
`KeyboardInterrupt` (when a user presses Ctrl+C to stop a frozen script) or
`SystemExit`. By specifying exactly what we expect to fail (`ValueError` and `TypeError`
for bad conversions), we prevent the script from becoming un-killable and make
debugging much easier.
"""
