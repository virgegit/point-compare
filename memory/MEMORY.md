# MEMORY — point-compare

## Snapshot

- Project type: measurement-point comparison utility
- Primary runtime: `compare_points_interactive.py`
- Alternate runtimes: `compare_points_v3.py`, `ComparePoints_v3.bas`
- Main artifact: Excel comparison report grouped by `MATCH / NAME_CHANGED / COORD_CHANGED / DELETED / ADDED`

## Durable Notes

- `AI_HANDOFF.md` and `memory/MEMORY.md` were created on 2026-04-14 because the project had no project-level knowledge files yet.
- The Python comparison flow is name-first and coordinate-second; coordinate matching uses tolerance-quantized keys and optional `I/J/K`.
- Duplicate names in the new dataset are currently lossy in Python because the name lookup keeps only one row per name.
- Current environment blocker: `py.exe` exists, but no installed Python interpreter was available, so scripts were not executed in this session.

## Open Work

- Add fixtures/tests for the comparison categories and for duplicate-name behavior.
- Clarify the expected product direction: CLI hardening, batch automation, or Excel-first workflow.
