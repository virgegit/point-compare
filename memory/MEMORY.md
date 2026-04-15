# MEMORY — point-compare

## Snapshot

- Project type: Excel worksheet comparison utility for measurement points
- Primary runtime: `compare_points_interactive.py`
- Alternate runtimes: `compare_points_v3.py`, `ComparePoints_v3.bas`
- Main artifact: `CMP_*` sheets written back into the source workbook and grouped by `MATCH / NAME_CHANGED / COORD_CHANGED / DELETED / ADDED`

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

## Open Work

- Add fixtures/tests for the comparison categories and for duplicate-name behavior.
- Decide whether duplicate names should fail fast, warn, or remain last-match-wins.
