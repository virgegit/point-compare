# AI Handoff — point-compare

## Purpose

`point-compare` compares two worksheets inside one Excel workbook for CMM / industrial measurement workflows.
Supported logical fields are `Name`, `X`, `Y`, `Z`, `I`, `J`, `K`.
The main output is a new Excel file with styled comparison reports and updated original sheet with status buckets:

- `MATCH` — name and coordinates match
- `NAME_CHANGED` — same coordinates, different name
- `COORD_CHANGED` — same name, different coordinates
- `REPLACED?` — original point deleted and replaced by nearby new point
- `DELETED` — point exists only in original
- `ADDED` — point exists only in new sheet

## Current Layout

- `compare_points_gui.py` — **primary entry point**; Tkinter desktop GUI that loads workbook, selects two sheets, runs comparison, displays summary, and creates new output file
- `compare_points_v3.py` — batch/config-driven Python variant for automation
- `ComparePoints_v3.bas` — Excel VBA implementation
- `run_point_compare_gui.bat` — Windows launcher for the GUI
- `PointsCompare_Template.xlsx` — workbook template with sample data
- `point_compare/` — shared library package (schema, core, styles, validators, excel_io)
- `archived/` — legacy interactive CLI (no longer used)
- `README.md` — quick-start and behavior summary

## Current Comparison Logic

All implementations use the shared `point_compare` library:

1. Read two source sheets from one Excel workbook.
2. Require exact source headers `Name/X/Y/Z/I/J/K` on both sheets.
3. Normalize columns into canonical fields via `prepare_sheet()`.
4. Run `compare()` function (core.py):
   - Build name lookup from new dataset
   - Build coordinate buckets using tolerance-quantized keys
   - For each original row: match by name first, then by coordinate key, else mark as `DELETED`
   - Unmatched new rows become `ADDED`
5. Run `link_moved_pairs()` to identify DELETED+ADDED pairs within `close_points_tol`:
   - If a DELETED has a nearby ADDED within tolerance, both are marked as `REPLACED?`
   - Populates `NEW_*` and difference columns on DELETED rows
6. Copy original workbook to **new output file** with timestamp: `filename_CMP_YYYYMMDD_HHMMSS.xlsx`
7. Append comparison results to original sheet:
   - Insert row 1 if needed for group labels
   - Write `STATUS`, `NEW_Name/X/Y/Z/I/J/K`, and `*_diff` columns after source data
   - Append unmatched new points as extra rows with `STATUS = NEW`
8. Create optional `CMP_*` report sheets (when enabled)

**Output File**: New file created; original workbook is never modified.

**Comparison Parameters**:
- `tol` — XYZ coordinate tolerance (mm)
- `ijk_tol` — I/J/K tolerance
- `use_ijk` — include I/J/K in coordinate matching
- `close_points_tol` — distance threshold for REPLACED? linking (default 10 mm, independent of `tol`)
- `create_report_sheets` — enable `CMP_*` report sheet generation

**Report Sheets** (when enabled):
- `CMP_Overview` — summary counts
- `CMP_All Results` — all points with status
- `CMP_Match` / `CMP_Name Changed` / `CMP_Coord Changed` / `CMP_Replaced` / `CMP_Deleted` / `CMP_Added` — filtered views
- `CMP_Close Points` — ALL original/new pairs within `close_points_tol` with status and 3D distance
- `CMP_Nearest Points` — for each original point, its single nearest new point

**Color Scheme** (openpyxl PatternFill):
- `MATCH` — light blue/gray theme
- `NAME_CHANGED`, `COORD_CHANGED`, `REPLACED?` — yellow (unified "changed" indication)
- `DELETED` — red
- `ADDED` / `NEW` — orange

**Original Sheet Writeback**:
- Shifts headers down to row 2, inserts grouped labels in row 1: `Original data`, `Comparison status`, `New Data`, `Difference`
- Preserves existing colors in source area
- Clears active AutoFilter before writing
- Uses source sheet's font family
- Displays coordinates with `0.000` formatting
- Difference columns (`*_diff`) highlighted in orange bold when non-zero

**Edge Cases**:
- Hidden Excel query tables are normalized to regular tables; external data names are stripped
- Sheet names with hidden leading/trailing spaces are resolved via unique trim-matches
- Duplicate names in new dataset: last match wins (name lookup dictionary behavior)

## Library Structure (`point_compare/` package)

- `__init__.py` — package marker
- `schema.py` — constants (DEFAULTS, PALETTE, STATUS_LABEL, STATUS_FILLS, field definitions, report naming)
- `core.py` — comparison logic (`compare()`, `link_moved_pairs()`, `build_close_points_report()`, `build_nearest_points_report()`)
- `validators.py` — input validation (`validate_workbook()`, `list_sheets()`)
- `styles.py` — formatting utilities (`StyleKit`, cell styling functions)
- `excel_io.py` — workbook I/O (`read_sheet()`, `prepare_sheet()`, `run_comparison()`, `write_data_sheet()`, `write_report_sheets()`)

All implementations (GUI, batch, potential future tools) import from this shared library to ensure consistency.

## Known Constraints

- Duplicate point names in the new dataset: last match wins (name lookup dictionary behavior).
- No local automated test fixtures yet.
- VBA implementation (`ComparePoints_v3.bas`) has not been updated to match recent Python changes (REPLACED? status, output-file mode).
- Original workbook is never modified; user must manage old comparison files manually.

## Recent Changes (This Session)

- **Refactored into shared library** (`point_compare/` package) to eliminate code duplication across CLI, GUI, and batch modes
- **Output file mode**: Comparison creates new timestamped file instead of modifying original workbook
- **REPLACED? status**: New status category for DELETED+ADDED pairs within `close_points_tol`; handled by fixed `link_moved_pairs()` logic
- **Fixed index-mapping bug** in `link_moved_pairs()`: now uses ORIG_*/NEW_* values directly from df_all instead of error-prone position math
- **Unified colors**: REPLACED?, COORD_CHANGED, NAME_CHANGED now all use yellow (same visual category)
- **Archived CLI**: Interactive script moved to `archived/` (GUI is now primary user-facing entry point)

## Recommended Next Work

- Add automated test fixtures for all six comparison categories (MATCH, NAME_CHANGED, COORD_CHANGED, REPLACED?, DELETED, ADDED)
- Implement duplicate-name error/warning to surface lossy behavior
- Update VBA implementation to match current Python status categories and output-file behavior
- Add batch-mode tests for concurrent workbook processing
