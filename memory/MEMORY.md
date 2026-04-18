# MEMORY — point-compare

## Snapshot

- Project type: Excel worksheet comparison utility for CMM measurement points (Name, X, Y, Z, I, J, K)
- Primary runtime: `compare_points_gui.py` (Tkinter GUI)
- Alternate runtimes: `compare_points_v3.py` (batch/config), `ComparePoints_v3.bas` (VBA, legacy)
- Shared library: `point_compare/` package with schema, core logic, validators, styles, excel_io modules
- Main artifact: **New timestamped output file** (`filename_CMP_YYYYMMDD_HHMMSS.xlsx`) with updated original sheet + `CMP_*` report sheets
- Comparison categories: MATCH / NAME_CHANGED / COORD_CHANGED / REPLACED? / DELETED / ADDED

## Durable Notes

- `AI_HANDOFF.md` and `memory/MEMORY.md` were created on 2026-04-14 because the project had no project-level knowledge files yet.
- The Python comparison flow is name-first and coordinate-second; coordinate matching uses tolerance-quantized keys and optional `I/J/K`.
- Source sheets are now expected to live in the same workbook and must expose exact headers `Name`, `X`, `Y`, `Z`, `I`, `J`, `K`.
- Project docs and scripts should use English-only text for prompts, labels, and documentation.
- `run_point_compare.bat` is the Windows launcher for the interactive flow. It reads `D:\projects\TOOLS\.env`, uses the configured `PYTHON`, and forces UTF-8 console mode so the script's box-drawing output works under `cmd.exe`.
- `compare_points_gui.py` is the desktop Tkinter entry point. It reuses the CLI comparison/report logic through a shared `run_comparison(...)` helper in `compare_points_interactive.py`.
- `run_point_compare_gui.bat` launches the GUI with the same root `.env` Python path.
- The Python flow now turns the original sheet into a labeled comparison table: if headers were on row 1, the sheet is shifted down so row 1 can hold `Original data`, `Comparison status`, `New Data`, and `Difference`. Then `STATUS`, appended new-data columns named `Name/X/Y/Z/I/J/K`, and all `*_diff` columns are written after the last used source column.
- Points that exist only in the new sheet are appended onto the original sheet as extra rows with `STATUS = NEW`.
- Existing colors on the original source area are intentionally left unchanged.
- Original-sheet writeback now inherits the source sheet's font family rather than forcing `Calibri`.
- The row-1 group labels and `STATUS` column colors now follow the reference workbook style: `MATCH` light theme, `NAME_CHANGED/COORD_CHANGED` yellow, `DELETED` red, `NEW` orange, with `Comparison status` and `Difference` blue and `New Data` yellow.
- Any active filter on the original sheet is cleared before writeback, and filtered-hidden rows are unhidden.
- `X/Y/Z/I/J/K` display values now use `0.000` number formatting on the original sheet and in the generated report sheets.
- Optional report-sheet creation now includes TWO proximity-review sheets:
  - **`CMP_Close Points`** — ALL original/new pairs within the configurable `close_points_tol` (default 10 mm). Independent of the XYZ comparison tolerance. Adjustable in CLI and GUI.
  - **`CMP_Nearest Points`** — For each original point, the SINGLE nearest new point regardless of distance. Helps identify shifted duplicates beyond the close-points threshold.
- Both proximity sheets share the same columns: `ORIG_Name/X/Y/Z`, `NEW_Name/X/Y/Z`, `dX/dY/dZ`, `Distance_3D`, `Same_Name`.
- Legacy workbooks that still contain the old `NEW_*` appended headers are rewritten in place instead of receiving a second appended output block.
- GUI users can choose whether to also create the `CMP_*` report sheets. If that option is off, the original sheet still includes the appended comparison columns and `NEW` rows, but the separate filtered `CMP_*` views are not created.
- Some real customer workbooks contain hidden query/import tables such as `AB3 FMK ALL ROW`. `openpyxl` drops the matching external-query package parts on save, so point-compare now post-processes the saved `.xlsx` package to remove hidden `ExternalData_*` names and convert `queryTable` definitions into regular worksheet tables, preventing Excel repair warnings on open.
- Excel sheet names may include hidden leading or trailing spaces. The GUI now resolves sheet selections against the exact workbook names, including unique trim-matches such as `New point list `.
- Duplicate names in the new dataset are currently lossy in Python because the name lookup keeps only one row per name.
- The workspace Python bootstrap now works for this project, and the interactive script, GUI import path, launchers, and labeled original-sheet writeback were validated on 2026-04-14.
- On 2026-04-15 the close-points tolerance was decoupled from the XYZ comparison tolerance. A new `close_points_tol` parameter (default 10 mm) is accepted by `run_comparison()`, prompted in the CLI, and exposed as a separate GUI field. The `CMP_Nearest Points` sheet was also added.

## Recent Changes (Session 2026-04-18)

- **Refactored into shared library** (`point_compare/` package) — eliminated 600+ lines of duplication across 3 entry points (interactive CLI, batch, GUI). All implementations now use identical comparison/styling logic
- **Output file mode** — comparison creates new timestamped file (`filename_CMP_YYYYMMDD_HHMMSS.xlsx`) instead of modifying original workbook. Original file is always safe
- **REPLACED? status** — new comparison category for DELETED+ADDED pairs within `close_points_tol`. Distinct from DELETED/ADDED: indicates original point was deleted AND a nearby new point exists as likely replacement
- **Fixed link_moved_pairs() bug** — was using incorrect position mapping through df_orig/df_new, causing many REPLACED? pairs to be missed. Now uses ORIG_*/NEW_* values directly from df_all
- **Unified colors** — REPLACED?, COORD_CHANGED, NAME_CHANGED all use yellow (same visual "changed" category). MATCH light blue, DELETED red, ADDED orange
- **Archived interactive CLI** — moved `compare_points_interactive.py` to `archived/` since GUI is superior (interactive, summary feedback, no terminal needed)
- **Color consistency** — all three statuses (REPLACED?, COORD_CHANGED, NAME_CHANGED) now display same yellow in both PALETTE and STATUS_FILLS

## Open Work

- Add fixtures/tests for all six comparison categories (MATCH, NAME_CHANGED, COORD_CHANGED, REPLACED?, DELETED, ADDED)
- Implement duplicate-name warning/error to surface lossy behavior
- Update VBA implementation to match current Python changes (REPLACED? status, output-file mode, color scheme)
- Add concurrent batch-processing tests for multiple workbooks
