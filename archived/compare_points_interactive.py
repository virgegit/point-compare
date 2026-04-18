#!/usr/bin/env python3
"""
compare_points_interactive.py
==============================
Interactive script for comparing two worksheets inside one Excel workbook.
Prompts for a workbook, two sheet names, and comparison settings, then
writes the result sheets back into the same Excel file.

Supported input formats:
  - Excel (.xlsx, .xlsm)

Result categories:
  MATCH          - name and coordinates match
  NAME_CHANGED   - same coordinates, different name
  COORD_CHANGED  - same name, different coordinates -> show dX/dY/dZ
  DELETED        - exists only in the original sheet
  ADDED          - exists only in the new sheet
"""

import sys
from pathlib import Path

import pandas as pd

from point_compare.schema import DEFAULTS, STATUS_LABEL
from point_compare.validators import validate_workbook, list_sheets
from point_compare.excel_io import read_sheet, prepare_sheet, run_comparison

# ══════════════════════════════════════════════════════════════════════════════
#  CLI UTILITIES (Interactive-specific)
# ══════════════════════════════════════════════════════════════════════════════

def hr(char="─", n=60):
    print(char * n)

def ask(prompt, default=None, validator=None):
    """Prompt the user and return a string."""
    hint = f" [{default}]" if default is not None else ""
    while True:
        raw = input(f"{prompt}{hint}: ").strip()
        if raw == "" and default is not None:
            return str(default)
        if raw == "" and default is None:
            print("  ⚠  A value is required.")
            continue
        if validator:
            ok, msg = validator(raw)
            if not ok:
                print(f"  ⚠  {msg}")
                continue
        return raw

def ask_yn(prompt, default=True):
    hint = "Y/n" if default else "y/N"
    raw = input(f"{prompt} [{hint}]: ").strip().lower()
    if raw == "":
        return default
    return raw in ("y", "yes")

def ask_float(prompt, default):
    while True:
        raw = input(f"{prompt} [{default}]: ").strip()
        if raw == "":
            return default
        try:
            return float(raw)
        except ValueError:
            print("  ⚠  Enter a numeric value such as 0.05")


# ══════════════════════════════════════════════════════════════════════════════
#  INTERACTIVE FLOW
# ══════════════════════════════════════════════════════════════════════════════

def interactive():
    hr("═")
    print("  MEASUREMENT POINTS COMPARISON")
    hr("═")
    print()

    hr()
    print("  STEP 1 OF 4 - EXCEL WORKBOOK")
    hr()
    workbook_path = ask("  Path to the Excel workbook", validator=validate_workbook)

    sheets = list_sheets(workbook_path)
    if not sheets:
        print("\n  ⚠  Unable to read worksheet names from the Excel file.")
        return None, None

    print(f"\n  Available sheets: {', '.join(str(s) for s in sheets)}")

    print()
    hr()
    print("  STEP 2 OF 4 - ORIGINAL SHEET")
    hr()
    orig_sheet = ask("  Name of the first sheet", default=str(sheets[0]))
    orig_sheet = orig_sheet if orig_sheet in sheets else sheets[0]

    print()
    hr()
    print("  STEP 3 OF 4 - NEW SHEET")
    hr()
    default_new_sheet = str(sheets[1]) if len(sheets) > 1 else str(sheets[0])
    new_sheet = ask("  Name of the second sheet", default=default_new_sheet)
    new_sheet = new_sheet if new_sheet in sheets else default_new_sheet

    print()
    hr()
    print("  Reading Excel sheets ...")
    hr()
    df_orig_raw = read_sheet(workbook_path, orig_sheet)
    df_new_raw = read_sheet(workbook_path, new_sheet)
    print(f"  Sheet 1: {orig_sheet}  |  Rows: {len(df_orig_raw)}  |  Columns: {list(df_orig_raw.columns)}")
    print(f"  Sheet 2: {new_sheet}  |  Rows: {len(df_new_raw)}  |  Columns: {list(df_new_raw.columns)}")

    print()
    hr()
    print("  STEP 4 OF 4 - COMPARISON SETTINGS")
    hr()
    tol     = ask_float("  XYZ tolerance (mm)", DEFAULTS["coord_tol"])
    ijk_tol = ask_float("  IJK tolerance",       DEFAULTS["ijk_tol"])
    use_ijk = ask_yn("  Include I/J/K in coordinate comparison?", DEFAULTS["compare_ijk"])
    close_tol = ask_float("  Close-points 3D tolerance (mm)", DEFAULTS["close_points_tol"])

    print()
    hr()
    print("  Normalizing data ...")
    try:
        result = run_comparison(workbook_path, orig_sheet, new_sheet, tol, ijk_tol, use_ijk, close_points_tol=close_tol)
    except ValueError as exc:
        print(f"  ⚠  {exc}")
        hr("═")
        return None, None

    df_all = result["df_all"]
    meta = result["meta"]

    vc = df_all["Status"].value_counts()
    print()
    hr()
    print("  RESULTS")
    hr()
    for code in ["MATCH","NAME_CHANGED","COORD_CHANGED","MOVED","DELETED","ADDED"]:
        print(f"  {STATUS_LABEL[code]:<40s}: {vc.get(code, 0):>5d}")
    print(f"  {'TOTAL':<40s}: {len(df_all):>5d}")
    print(f"  {'Close 3D pairs (within tolerance)':<40s}: {result['close_points_count']:>5d}")
    print(f"  {'Moved pairs (DELETED+ADDED within tolerance)':<40s}: {result['moved_count']:>5d}")
    hr()

    print(f"\n  ✅ Done. The original sheet was updated and CMP_* sheets were written to {Path(workbook_path).resolve()}")
    hr("═")

    return df_all, workbook_path


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    try:
        interactive()
    except KeyboardInterrupt:
        print("\n\n  Cancelled by user.")
        sys.exit(0)
