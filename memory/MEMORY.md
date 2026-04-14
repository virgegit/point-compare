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
- The Python flow now turns the original sheet into a labeled comparison table: if headers were on row 1, the sheet is shifted down so row 1 can hold `Original data`, `STATUS`, `New Data`, and `Difference`. Then `STATUS`, all `NEW_*`, and all `*_diff` columns are appended after the last used source column.
- Points that exist only in the new sheet are appended onto the original sheet as extra rows with `STATUS = NEW`.
- Existing colors on the original source area are intentionally left unchanged.
- GUI users can choose whether to also create the `CMP_*` report sheets. If that option is off, the original sheet still includes the appended comparison columns and `NEW` rows, but the separate filtered `CMP_*` views are not created.
- Duplicate names in the new dataset are currently lossy in Python because the name lookup keeps only one row per name.
- The workspace Python bootstrap now works for this project, and the interactive script, GUI import path, launchers, and labeled original-sheet writeback were validated on 2026-04-14.

## Open Work

- Add fixtures/tests for the comparison categories and for duplicate-name behavior.
- Decide whether duplicate names should fail fast, warn, or remain last-match-wins.
