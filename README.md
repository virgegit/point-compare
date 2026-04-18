# point-compare

Compare two worksheets inside one Excel workbook with CMM / industrial measurement points (`Name`, `X`, `Y`, `Z`, `I`, `J`, `K`).

The comparison creates a **new output file** (original workbook is never modified) with detailed comparison results and proximity analysis.

## Files

| File | Description |
|------|-------------|
| `compare_points_gui.py` | **Main entry point** – Tkinter desktop GUI with workbook browser, sheet selection, tolerance settings, and live summary |
| `compare_points_v3.py` | Batch/config-based Python script for automation |
| `ComparePoints_v3.bas` | VBA module for Excel (legacy, not maintained) |
| `PointsCompare_Template.xlsx` | Sample workbook with measurement points |
| `point_compare/` | Shared library package (schema, core logic, validators, styling, Excel I/O) |

## Quick Start

### GUI (Recommended)

```bat
run_point_compare_gui.bat
```

The GUI walks you through:
1. **Browse workbook** — select `.xlsx` or `.xlsm`
2. **Load sheets** — populate dropdowns with available worksheet names
3. **Select sheets** — choose original and new sheets
4. **Set tolerances** — XYZ (mm), IJK, and close-points distance (mm)
5. **Run** — compare and view summary
6. **Output** — new file created: `original_name_CMP_YYYYMMDD_HHMMSS.xlsx`

### Batch (Python)

```bash
pip install pandas openpyxl numpy
python compare_points_v3.py
```

Edit `CONFIG` dict in the script to set workbook path, sheet names, and tolerances, then run.

## Output

A new Excel file is created with:
- **Original sheet (updated)** — appended result columns and new-only rows
- **`CMP_Overview`** — summary counts by status
- **`CMP_All Results`** — all points with status and diff columns
- **`CMP_Match` / `CMP_Name Changed` / `CMP_Coord Changed` / `CMP_Replaced` / `CMP_Deleted` / `CMP_Added`** — filtered views
- **`CMP_Close Points`** — original/new point pairs within configurable distance (default 10 mm)
- **`CMP_Nearest Points`** — for each original point, its nearest new point (any distance)

The original source sheet is also updated in place with:
- Row 1 becomes a grouped label row (`Original data`, `Comparison status`, `New Data`, `Difference`)
- `STATUS` column showing comparison result
- Appended `Name / X / Y / Z / I / J / K` columns for the new-sheet values
- All `*_diff` columns showing coordinate/name changes

Layout on the original sheet:
- Original source headers move from row 1 to row 2 when needed
- Row 1 labels the sections as `Original data`, `Comparison status`, `New Data`, and `Difference`
- Any active AutoFilter is cleared before writing results; hidden rows are unhidden
- Result columns are appended after the last used source column
- Points that exist only in new sheet are appended as extra rows with `STATUS = NEW`
- `X / Y / Z / I / J / K` display values use `0.000` formatting
- Appended/group/status cells inherit the source sheet's font family
- Status colors: `MATCH` light blue/gray, `NAME_CHANGED/COORD_CHANGED/REPLACED?` yellow, `DELETED` red, `NEW` orange

## Comparison Categories

| Status | Meaning | Color |
|--------|---------|-------|
| `MATCH` | Name and coordinates match within tolerance | Light blue |
| `NAME_CHANGED` | Same coordinates, different name | Yellow |
| `COORD_CHANGED` | Same name, different coordinates → dX/dY/dZ shown | Yellow |
| `REPLACED?` | Original point deleted, replaced by nearby new point (within close-points distance) | Yellow |
| `DELETED` | Point exists only in the original list | Red |
| `ADDED` | Point exists only in the new list | Orange |

On the original sheet, new-only points are shown as appended rows with `STATUS = NEW`.

## Tolerances

- **XYZ tolerance** — coordinate matching threshold (mm)
- **IJK tolerance** — optional I/J/K matching threshold
- **Close-points distance** — threshold for REPLACED? detection (default 10 mm, independent of XYZ tolerance)

## Required Source Columns

Both source sheets must contain these exact column names (case-sensitive):

- `Name`
- `X`
- `Y`
- `Z`
- `I`
- `J`
- `K`

If any required column is missing, the tool shows a warning and stops.

## Diff Columns

Diff columns (`X_diff`, `Y_diff`, `Z_diff`, `I_diff`, `J_diff`, `K_diff`, `NAME_diff`, `DIFF_Fields`) are **highlighted in orange bold** when non-zero.

## Supported Workbook Formats

- `.xlsx`
- `.xlsm`

## Requirements

```
pandas
openpyxl
numpy
```
