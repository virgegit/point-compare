# AI Handoff — point-compare

## Purpose

`point-compare` compares two lists of measurement points used in CMM / industrial workflows.
Supported logical fields are `Name`, `X`, `Y`, `Z`, `I`, `J`, `K`.
The main output is a styled Excel report with status buckets:

- `MATCH`
- `NAME_CHANGED`
- `COORD_CHANGED`
- `DELETED`
- `ADDED`

## Current Layout

- `compare_points_interactive.py` — primary entry point; interactive CLI that asks for files, sheet names, column mapping, tolerances, and output path
- `compare_points_v3.py` — batch/config-driven Python variant with the same comparison core
- `ComparePoints_v3.bas` — Excel VBA implementation
- `PointsCompare_Template.xlsm` — workbook/template placeholder distributed with the project
- `README.md` — quick-start and behavior summary

## Current Comparison Logic

Both Python implementations use the same matching strategy:

1. Read CSV or Excel input.
2. Normalize source columns into canonical fields `Name/X/Y/Z/I/J/K`.
3. Build a name lookup from the new dataset.
4. Build coordinate buckets from the new dataset using a tolerance-quantized key.
5. For each original row:
   - match by `Name` first
   - if not found, try match by coordinate key
   - otherwise mark as `DELETED`
6. Any unmatched rows in the new dataset become `ADDED`.

`compare_ijk` is optional and extends the coordinate key with `I/J/K`.

## Observed Constraints

- Project-specific handoff/memory files did not exist before this session; they were created on 2026-04-14 to restore workspace protocol compliance.
- The local machine currently exposes `py.exe`, but no installed Python interpreter was available from it during this session, so runtime validation was not possible.
- In both Python scripts, duplicate names in the new dataset are collapsed by the name lookup dictionary, so the last duplicate wins during name-based matching.
- Duplicate coordinate keys are supported as buckets, but duplicate-name behavior is not explicitly surfaced to the user.
- There are no local automated tests or sample fixtures in the repo yet.

## Recommended Next Work

- Add reproducible sample fixtures and a small automated regression test set for the five result categories.
- Decide whether duplicate point names should be treated as an error, warning, or supported multi-match case.
- If this tool remains user-facing, consider converging the Python and VBA behaviors into one documented reference implementation.

## Session Note

On 2026-04-14 the project was onboarded into the root wiki workflow, and its baseline behavior was documented before feature work.
