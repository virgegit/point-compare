# point-compare

Compare two lists of CMM / industrial measurement points (Name, X, Y, Z, I, J, K).

## Features

- Compare by **name** and **coordinates** simultaneously
- Detects:
  - ✔ **MATCH** – same name and coordinates (within tolerance)
  - ✎ **NAME_CHANGED** – same coordinates, different name → shows `NAME_diff`
  - ⚠ **COORD_CHANGED** – same name, different coordinates → shows `dX`, `dY`, `dZ`, `dI`, `dJ`, `dK` + `DIFF_Fields`
  - ✖ **DELETED** – exists only in List 1
  - ＋ **ADDED** – exists only in List 2
- Configurable column mapping (any column names)
- Compare two CSV files **or** two sheets of the same Excel file
- Configurable XYZ and IJK tolerances
- Fully configurable colour palette in reports
- Diff cells highlighted in orange for instant visibility

## Files

| File | Description |
|---|---|
| `compare_points_v3.py` | Python script — produces styled Excel report |
| `ComparePoints_v3.bas` | VBA module — runs inside Excel, compares two sheets |
| `PointsCompare_Template.xlsm` | Ready-to-use Excel template with all sheets + VBA instructions |

## Python Usage

Edit the `CONFIG` section at the top of `compare_points_v3.py`:

```python
CONFIG = {
    "file1":  "path/to/file1.csv",   # or .xlsx
    "sheet1": None,                   # sheet name or None
    "file2":  "path/to/file2.csv",
    "sheet2": None,
    "col_map1": {
        "Name": "Label",
        "X": "X", "Y": "Y", "Z": "Z",
        "I": "Vx", "J": "Vy", "K": "Vz",
    },
    "col_map2": {
        "Name": "PointName",
        "X": "X_mm", "Y": "Y_mm", "Z": "Z_mm",
        "I": None, "J": None, "K": None,
    },
    "coord_tol": 0.05,
    "ijk_tol":   0.001,
    "compare_ijk": False,
    "output_file": "output/report.xlsx",
    "palette": {
        "MATCH":         "C6EFCE",
        "NAME_CHANGED":  "FFEB9C",
        "COORD_CHANGED": "FFD7D7",
        "DELETED":       "F4CCCC",
        "ADDED":         "D9EAD3",
        "DIFF_CELL":     "FF6600",
        "HEADER_BG":     "4472C4",
        "HEADER_FG":     "FFFFFF",
        "OVERVIEW_TITLE":"1F3864",
    },
}
```

Run:
```bash
pip install pandas openpyxl numpy
python compare_points_v3.py
```

## Excel VBA Usage

1. Open `PointsCompare_Template.xlsm` (or your own `.xlsm`)
2. Press `Alt + F11` → Insert → Module
3. Paste the full code from `ComparePoints_v3.bas`
4. Close the editor
5. Go to sheet **▶ ЗАПУСК**, right-click the green cell → **Assign Macro → ComparePoints**
6. Click the button — results appear on `CMP_Overview`, `CMP_All Results`, etc.

Settings (sheet names, column mapping, tolerances, colours) are read from the **⚙ Настройки** sheet — no code editing needed.

## Output Sheets

| Sheet | Content |
|---|---|
| `CMP_Overview` | Summary table with counts per status + colour legend |
| `CMP_All Results` | All records with status + diff columns |
| `CMP_Match` | Only matching points |
| `CMP_Name Changed` | Points with same coords but different name |
| `CMP_Coord Changed` | Points with same name but different coords |
| `CMP_Deleted` | Points only in List 1 |
| `CMP_Added` | Points only in List 2 |

## Requirements (Python)

- Python 3.8+
- pandas
- openpyxl
- numpy
