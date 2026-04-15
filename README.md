# point-compare

Compare two worksheets inside one Excel workbook with CMM / industrial measurement points (`Name`, `X`, `Y`, `Z`, `I`, `J`, `K`).

## Files

| File | Description |
|------|-------------|
| `compare_points_interactive.py` | **Main script** – interactive CLI, asks for one workbook and two sheet names, writes result sheets back into the same Excel file |
| `compare_points_gui.py` | Desktop GUI for the same workbook-first flow |
| `compare_points_v3.py` | Batch/config-based Python script for the same one-workbook workflow |
| `ComparePoints_v3.bas` | VBA module for Excel (import via Alt+F11 → Insert → Module) |
| `PointsCompare_Template.xlsm` | Excel template / workbook placeholder |

## Quick Start

```bash
pip install pandas openpyxl numpy
python compare_points_interactive.py
```

Or start the interactive tool with the included launcher:

```bat
run_point_compare.bat
```

Or start the desktop GUI:

```bat
run_point_compare_gui.bat
```

The script will guide you through 4 steps:
1. Select the **Excel workbook**
2. Select the **first source sheet**
3. Select the **second source sheet**
4. Set tolerances (XYZ, IJK) and options

Result sheets are written into the same workbook with the `CMP_` prefix.

The original source sheet is also updated in place:
- row 1 becomes a grouped label row
- `STATUS`
- appended `Name / X / Y / Z / I / J / K` columns for the new-sheet values
- all `*_diff` columns

Layout on the original sheet:
- original source headers move from row 1 to row 2 when needed
- row 1 labels the sections as `Original data`, `Comparison status`, `New Data`, and `Difference`
- any active AutoFilter on the original sheet is cleared before writing results
- result columns are appended after the last used source column
- points that exist only in sheet 2 are appended as extra rows with `STATUS = NEW`
- `X / Y / Z / I / J / K` display values use `0.000` formatting on the original sheet and in the report sheets
- appended/group/status cells inherit the source sheet's font family, and the row-1 / status colors follow the reference workbook scheme

## GUI Flow

The GUI lets you:
- browse to one `.xlsx` or `.xlsm` workbook
- load worksheet names into dropdowns
- choose the original and new sheets
- set XYZ / IJK tolerances
- run the comparison and view the summary without using the terminal
- choose whether to also create the separate `CMP_*` report sheets

When `CMP_*` report creation is enabled, the workbook also gets a `CMP_Close Points` sheet listing original/new point pairs whose 3D XYZ distance is within the selected XYZ tolerance. This is meant to highlight likely same points that were shifted slightly.

## Required Source Columns

Both source sheets must contain these exact column names:

- `Name`
- `X`
- `Y`
- `Z`
- `I`
- `J`
- `K`

If any required column is missing on either sheet, the tool shows a warning and stops so the headers can be fixed.

## Comparison Categories

| Status | Meaning |
|--------|---------|
| `MATCH` | Name and coordinates match within tolerance |
| `NAME_CHANGED` | Same coordinates, different name |
| `COORD_CHANGED` | Same name, different coordinates → dX/dY/dZ shown |
| `DELETED` | Point exists only in the original list |
| `ADDED` | Point exists only in the new list |

On the original sheet writeback, list-2-only points are shown as appended rows with `STATUS = NEW`.

## Output Report (Excel)

The report contains 8 sheets:
- **`CMP_Overview`** – summary table with counts per category
- **`CMP_All Results`** – all points with status and diff columns
- **`CMP_Match`** / **`CMP_Name Changed`** / **`CMP_Coord Changed`** / **`CMP_Deleted`** / **`CMP_Added`** – filtered views
- **`CMP_Close Points`** – pairs of original/new points whose 3D XYZ distance is within the selected XYZ tolerance

If you disable report-sheet creation in the GUI, the original sheet still gets the appended result columns and `NEW` rows, but the separate filtered `CMP_*` views are not created.

Diff columns (`X_diff`, `Y_diff`, `Z_diff`, `NAME_diff`, `DIFF_Fields`) are highlighted in **orange bold** when non-zero.

## Supported Workbook Formats

- `.xlsx`
- `.xlsm`

## Requirements

```
pandas
openpyxl
numpy
```
