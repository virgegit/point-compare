# point-compare

Compare two lists of CMM / industrial measurement points (Name, X, Y, Z, I, J, K).

## Files

| File | Description |
|------|-------------|
| `compare_points_interactive.py` | **Main script** – interactive CLI, asks for files and column mapping, outputs Excel report |
| `compare_points_v3.py` | Batch/config-based Python script (no interactive prompts) |
| `ComparePoints_v3.bas` | VBA module for Excel (import via Alt+F11 → Insert → Module) |
| `PointsCompare_Template.xlsm` | Excel template with sample data, config sheet and instructions |

## Quick Start

```bash
pip install pandas openpyxl numpy
python compare_points_interactive.py
```

The script will guide you through 5 steps:
1. Select the **original** file (Excel or CSV)
2. Select the **new** file (Excel or CSV)
3. Map columns to logical fields (Name / X / Y / Z / I / J / K)
4. Set tolerances (XYZ, IJK) and options
5. Choose output report path

## Comparison Categories

| Status | Meaning |
|--------|---------|
| `MATCH` | Name and coordinates match within tolerance |
| `NAME_CHANGED` | Same coordinates, different name |
| `COORD_CHANGED` | Same name, different coordinates → dX/dY/dZ shown |
| `DELETED` | Point exists only in the original list |
| `ADDED` | Point exists only in the new list |

## Output Report (Excel)

The report contains 7 sheets:
- **Overview** – summary table with counts per category
- **All Results** – all points with status and diff columns
- **✔ Match** / **✎ Name Changed** / **⚠ Coord Changed** / **✖ Deleted** / **＋ Added** – filtered views

Diff columns (`X_diff`, `Y_diff`, `Z_diff`, `NAME_diff`, `DIFF_Fields`) are highlighted in **orange bold** when non-zero.

## Supported Input Formats

- `.csv` (auto-detect `,` or `;` separator)
- `.xlsx`, `.xls`, `.xlsm` (sheet selection supported)

## Requirements

```
pandas
openpyxl
numpy
```
