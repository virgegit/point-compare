# AI Handoff — point-compare

## Purpose

`point-compare` compares two worksheets inside one Excel workbook for CMM / industrial measurement workflows.
Supported logical fields are `Name`, `X`, `Y`, `Z`, `I`, `J`, `K`.
The main output is a styled Excel report with status buckets:

- `MATCH`
- `NAME_CHANGED`
- `COORD_CHANGED`
- `DELETED`
- `ADDED`

## Current Layout

- `compare_points_interactive.py` — primary entry point; interactive CLI that asks for one workbook, two sheet names, tolerances, and writes result sheets back into the same workbook
- `compare_points_gui.py` — Tkinter desktop GUI for the same one-workbook flow
- `compare_points_v3.py` — batch/config-driven Python variant with the same one-workbook flow
- `ComparePoints_v3.bas` — Excel VBA implementation
- `run_point_compare.bat` / `run_point_compare_gui.bat` — Windows launchers that read the root `.env` Python path
- `PointsCompare_Template.xlsm` — workbook/template placeholder distributed with the project
- `README.md` — quick-start and behavior summary

## Current Comparison Logic

All implementations now use the same workbook-first model:

1. Read two source sheets from one Excel workbook.
2. Require exact source headers `Name/X/Y/Z/I/J/K` on both sheets.
3. Normalize those columns into canonical fields.
4. Build a name lookup from the new dataset.
5. Build coordinate buckets from the new dataset using a tolerance-quantized key.
6. For each original row:
   - match by `Name` first
   - if not found, try match by coordinate key
   - otherwise mark as `DELETED`
7. Any unmatched rows in the new dataset become `ADDED`.
8. Write comparison output to new `CMP_*` sheets in that same workbook.

`compare_ijk` is optional and extends the coordinate key with `I/J/K`.

The GUI should stay thin and reuse the same comparison/report-writing path as the CLI to avoid drift between interfaces.
The current Python flow also appends `STATUS`, plain `Name/X/Y/Z/I/J/K` columns for the new-sheet values, and all `*_diff` columns onto the original source sheet after its last used source column. If the source headers start on row 1, the sheet is shifted down so row 1 can hold grouped labels: `Original data`, `Comparison status`, `New Data`, and `Difference`. Existing colors in the original source area should remain untouched.
Rows that exist only in the new dataset are appended onto the original sheet as extra rows with `STATUS = NEW`.
Original-sheet writeback now inherits the source sheet's font family and uses the reference workbook's row-1 / status-column color scheme (`MATCH` light theme, `NAME_CHANGED/COORD_CHANGED` yellow, `DELETED` red, `NEW` orange).
Before writing onto the original sheet, any active AutoFilter is cleared and hidden filtered rows are unhidden.
Coordinate display on the original sheet and report sheets now uses `0.000` formatting for `X/Y/Z/I/J/K` values.
When the optional `CMP_*` report sheets are enabled, the workbook also gets two proximity-review sheets:

- **`CMP_Close Points`** — ALL original/new pairs whose 3D XYZ distance is within the configurable `close_points_tol` (default 10 mm). This tolerance is independent of the XYZ comparison tolerance and is adjustable in both the CLI and GUI.
- **`CMP_Nearest Points`** — For each original point, the SINGLE nearest new point regardless of distance. This helps identify potential shifted duplicates even when they exceed the close-points threshold.

Both sheets share the same column schema (`ORIG_Name/X/Y/Z`, `NEW_Name/X/Y/Z`, `dX/dY/dZ`, `Distance_3D`, `Same_Name`) and formatting.
Some customer workbooks contain hidden Excel query/import tables. Because `openpyxl` does not preserve the full external-query package, the save path now normalizes query tables into regular worksheet tables and strips hidden `ExternalData_*` names from the workbook package after save so Excel will not open the file with a repair warning.

## Observed Constraints

- Project-specific handoff/memory files did not exist before this session; they were created on 2026-04-14 to restore workspace protocol compliance.
- The workspace Python bootstrap now works, and the Python scripts were validated during this session with compile/import checks plus a temporary workbook smoke test.
- In both Python scripts, duplicate names in the new dataset are collapsed by the name lookup dictionary, so the last duplicate wins during name-based matching.
- Duplicate coordinate keys are supported as buckets, but duplicate-name behavior is not explicitly surfaced to the user.
- There are no local automated tests or sample fixtures in the repo yet.

## Recommended Next Work

- Add reproducible sample fixtures and a small automated regression test set for the five result categories.
- Decide whether duplicate point names should be treated as an error, warning, or supported multi-match case.
- If this tool remains user-facing, consider converging the Python and VBA behaviors into one documented reference implementation.

## Session Note

On 2026-04-14 the project was onboarded into the root wiki workflow, then updated to a workbook-first flow with strict source headers and in-place `CMP_*` result sheets.
