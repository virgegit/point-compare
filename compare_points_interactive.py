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
from datetime import datetime

import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# ══════════════════════════════════════════════════════════════════════════════
#  DEFAULT SETTINGS
# ══════════════════════════════════════════════════════════════════════════════
DEFAULTS = {
    "coord_tol":   0.05,   # XYZ tolerance, mm
    "ijk_tol":     0.001,  # IJK tolerance
    "compare_ijk": False,  # include I/J/K in coordinate comparison?
}

PALETTE = {
    "MATCH":         "C6EFCE",
    "NAME_CHANGED":  "FFEB9C",
    "COORD_CHANGED": "FFD7D7",
    "DELETED":       "F4CCCC",
    "ADDED":         "D9EAD3",
    "DIFF_CELL":     "CC3300",
    "HEADER_BG":     "4472C4",
    "HEADER_FG":     "FFFFFF",
    "TITLE":         "1F3864",
}

STATUS_LABEL = {
    "MATCH":         "✔  Match",
    "NAME_CHANGED":  "✎  Name Changed",
    "COORD_CHANGED": "⚠  Coordinates Changed",
    "DELETED":       "✖  Deleted (only in original)",
    "ADDED":         "＋ Added (only in new)",
}

LOGICAL_FIELDS = ["Name", "X", "Y", "Z", "I", "J", "K"]
FONT_NAME = "Calibri"
REPORT_PREFIX = "CMP_"
REPORT_SUFFIXES = [
    "Overview",
    "All Results",
    "Match",
    "Name Changed",
    "Coord Changed",
    "Deleted",
    "Added",
]

GROUP_HEADER_LABELS = {"Original data", "New Data", "Difference", "STATUS"}

ORIGINAL_SHEET_APPEND_MAP = [
    ("Status", "STATUS"),
    ("NEW_Name", "NEW_Name"),
    ("NEW_X", "NEW_X"),
    ("NEW_Y", "NEW_Y"),
    ("NEW_Z", "NEW_Z"),
    ("NEW_I", "NEW_I"),
    ("NEW_J", "NEW_J"),
    ("NEW_K", "NEW_K"),
    ("NAME_diff", "NAME_diff"),
    ("X_diff", "X_diff"),
    ("Y_diff", "Y_diff"),
    ("Z_diff", "Z_diff"),
    ("I_diff", "I_diff"),
    ("J_diff", "J_diff"),
    ("K_diff", "K_diff"),
    ("DIFF_Fields", "DIFF_Fields"),
]

ORIGINAL_SHEET_NUMERIC_COLUMNS = {
    "NEW_X", "NEW_Y", "NEW_Z", "NEW_I", "NEW_J", "NEW_K",
    "X_diff", "Y_diff", "Z_diff", "I_diff", "J_diff", "K_diff",
}

# ══════════════════════════════════════════════════════════════════════════════
#  UTILITIES
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

def validate_workbook(path_str):
    p = Path(path_str)
    if not p.exists():
        return False, f"File not found: {p}"
    if p.suffix.lower() not in (".xlsx", ".xlsm"):
        return False, "Only .xlsx and .xlsm Excel files are supported"
    return True, ""

def list_sheets(path):
    """Return the worksheet names from an Excel file."""
    try:
        xl = pd.ExcelFile(path)
        return xl.sheet_names
    except Exception:
        return []

def read_sheet(path, sheet):
    """Read an Excel sheet into a DataFrame."""
    header_row = detect_header_row(path, sheet)
    return pd.read_excel(path, sheet_name=sheet, header=header_row - 1)

def detect_header_row(path, sheet):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet]
        row1 = ["" if cell.value is None else str(cell.value).strip() for cell in ws[1]]
        row2 = ["" if cell.value is None else str(cell.value).strip() for cell in ws[2]]
    finally:
        wb.close()

    if any(value in GROUP_HEADER_LABELS for value in row1) and all(field in row2 for field in LOGICAL_FIELDS):
        return 2
    return 1


# ══════════════════════════════════════════════════════════════════════════════
#  COMPARISON LOGIC
# ══════════════════════════════════════════════════════════════════════════════

def prepare_sheet(df, sheet_name):
    """Validate required columns and normalize the sheet."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    missing = [field for field in LOGICAL_FIELDS if field not in df.columns]
    if missing:
        missing_str = ", ".join(missing)
        raise ValueError(
            f"Required columns are missing on sheet '{sheet_name}': {missing_str}. "
            "Fix the worksheet headers in the Excel file and run the tool again."
        )

    out = {}
    for field in LOGICAL_FIELDS:
        if field == "Name":
            out[field] = df[field].where(df[field].notna(), "").astype(str).str.strip()
        else:
            out[field] = pd.to_numeric(df[field], errors="coerce")
    prepared = pd.DataFrame(out)
    keep_mask = (prepared["Name"] != "") | prepared[["X", "Y", "Z", "I", "J", "K"]].notna().any(axis=1)
    return prepared.loc[keep_mask].reset_index(drop=True)

def coord_key(row, tol, ijk_tol, use_ijk):
    def rnd(v, t):
        try:
            return int(round(float(v) / t))
        except (TypeError, ValueError):
            return None
    k = (rnd(row["X"], tol), rnd(row["Y"], tol), rnd(row["Z"], tol))
    if use_ijk:
        k += (rnd(row["I"], ijk_tol), rnd(row["J"], ijk_tol), rnd(row["K"], ijk_tol))
    return k

def safe_delta(v1, v2):
    try:
        return round(float(v2) - float(v1), 4)
    except (TypeError, ValueError):
        return np.nan

def fmt_val(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "–"
    if isinstance(v, float):
        return f"{v:.4f}"
    return str(v)

def compare(df_orig, df_new, tol, ijk_tol, use_ijk):
    name_to_new  = {r["Name"]: i for i, r in df_new.iterrows()}
    coord_to_new = {}
    for i, r in df_new.iterrows():
        k = coord_key(r, tol, ijk_tol, use_ijk)
        coord_to_new.setdefault(k, []).append(i)

    matched_new = set()
    results = []

    for _, r1 in df_orig.iterrows():
        n1   = r1["Name"]
        ck1  = coord_key(r1, tol, ijk_tol, use_ijk)
        by_name  = n1 in name_to_new
        by_coord = ck1 in coord_to_new

        row = {f"ORIG_{f}": r1[f] for f in LOGICAL_FIELDS}
        row.update({f"NEW_{f}": np.nan for f in LOGICAL_FIELDS})
        row.update({
            "NAME_diff":  "",
            "X_diff": np.nan, "Y_diff": np.nan, "Z_diff": np.nan,
            "I_diff": np.nan, "J_diff": np.nan, "K_diff": np.nan,
            "DIFF_Fields": "",
        })

        if by_name:
            i2 = name_to_new[n1]
            r2 = df_new.loc[i2]
            matched_new.add(i2)
            row.update({f"NEW_{f}": r2[f] for f in LOGICAL_FIELDS})
            ck2 = coord_key(r2, tol, ijk_tol, use_ijk)
            if ck1 == ck2:
                row["Status"] = "MATCH"
            else:
                row["Status"] = "COORD_CHANGED"
                diffs = []
                for f in ["X","Y","Z","I","J","K"]:
                    d = safe_delta(r1[f], r2[f])
                    row[f"{f}_diff"] = d
                    if not (isinstance(d, float) and np.isnan(d)) and abs(d) > 1e-9:
                        diffs.append(f)
                row["DIFF_Fields"] = ", ".join(diffs)

        elif by_coord:
            found = False
            for i2 in coord_to_new[ck1]:
                if i2 not in matched_new:
                    r2 = df_new.loc[i2]
                    matched_new.add(i2)
                    row.update({f"NEW_{f}": r2[f] for f in LOGICAL_FIELDS})
                    row["Status"]    = "NAME_CHANGED"
                    row["NAME_diff"] = f"{n1}  →  {r2['Name']}"
                    row["DIFF_Fields"] = "Name"
                    found = True
                    break
            if not found:
                row["Status"] = "DELETED"
        else:
            row["Status"] = "DELETED"

        results.append(row)

    for i2, r2 in df_new.iterrows():
        if i2 not in matched_new:
            row = {"Status": "ADDED"}
            row.update({f"ORIG_{f}": np.nan for f in LOGICAL_FIELDS})
            row.update({f"NEW_{f}": r2[f] for f in LOGICAL_FIELDS})
            row.update({
                "NAME_diff": "", "DIFF_Fields": "",
                "X_diff": np.nan, "Y_diff": np.nan, "Z_diff": np.nan,
                "I_diff": np.nan, "J_diff": np.nan, "K_diff": np.nan,
            })
            results.append(row)

    return pd.DataFrame(results)


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL REPORT
# ══════════════════════════════════════════════════════════════════════════════

COL_ORDER = [
    "Status",
    "ORIG_Name","ORIG_X","ORIG_Y","ORIG_Z","ORIG_I","ORIG_J","ORIG_K",
    "NEW_Name", "NEW_X", "NEW_Y", "NEW_Z", "NEW_I", "NEW_J", "NEW_K",
    "NAME_diff","X_diff","Y_diff","Z_diff","I_diff","J_diff","K_diff",
    "DIFF_Fields",
]

DIFF_COLS = {"NAME_diff","X_diff","Y_diff","Z_diff","I_diff","J_diff","K_diff","DIFF_Fields"}

def thin(color="CCCCCC"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def ap(cell, style):
    for k, v in style.items():
        setattr(cell, k, v)

def hdr_style():
    return dict(
        font=Font(name=FONT_NAME, bold=True, color=PALETTE["HEADER_FG"], size=11),
        fill=PatternFill("solid", fgColor=PALETTE["HEADER_BG"]),
        alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
        border=thin(),
    )

def data_style(status, is_diff=False):
    bg = PALETTE.get(status, "FFFFFF")
    if is_diff:
        font = Font(name=FONT_NAME, size=10, bold=True, color=PALETTE["DIFF_CELL"])
    else:
        font = Font(name=FONT_NAME, size=10)
    return dict(
        font=font,
        fill=PatternFill("solid", fgColor=bg),
        alignment=Alignment(horizontal="left", vertical="center"),
        border=thin(),
    )

def original_sheet_header_style():
    return dict(
        font=Font(name=FONT_NAME, bold=True, color="222222", size=11),
        fill=PatternFill("solid", fgColor="EDEDED"),
        alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
        border=thin(),
    )

def original_sheet_group_style():
    return dict(
        font=Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11),
        fill=PatternFill("solid", fgColor="5B9BD5"),
        alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
        border=thin(),
    )

def original_sheet_data_style(is_diff=False):
    font = Font(name=FONT_NAME, size=10, bold=is_diff, color=PALETTE["DIFF_CELL"] if is_diff else "000000")
    return dict(
        font=font,
        alignment=Alignment(horizontal="left", vertical="center"),
        border=thin(),
    )

def write_data_sheet(wb, title, df, table_name):
    ws = wb.create_sheet(title)
    if df.empty:
        ws.cell(1, 1, "No data").font = Font(italic=True)
        return

    cols = [c for c in COL_ORDER if c in df.columns]

    for ci, col in enumerate(cols, 1):
        ap(ws.cell(1, ci, col), hdr_style())
    ws.row_dimensions[1].height = 28

    for ri, (_, row) in enumerate(df.iterrows(), 2):
        status = str(row.get("Status", ""))
        for ci, col in enumerate(cols, 1):
            val = row[col]
            if isinstance(val, float) and np.isnan(val):
                val = ""
            is_diff = (col in DIFF_COLS) and val not in ("", 0, 0.0) \
                      and not (isinstance(val, float) and abs(val) < 1e-9)
            c = ws.cell(ri, ci, val)
            ap(c, data_style(status, is_diff))
            if col not in ("Status", "ORIG_Name", "NEW_Name", "NAME_diff", "DIFF_Fields"):
                c.number_format = "0.0000"
        ws.row_dimensions[ri].height = 18

    for ci, col in enumerate(cols, 1):
        sample = [str(col)] + [str(df.iloc[r, ci-1]) for r in range(min(300, len(df)))]
        w = min(max(len(v) + 2 for v in sample), 46)
        ws.column_dimensions[get_column_letter(ci)].width = w

    ref = f"A1:{get_column_letter(len(cols))}{len(df)+1}"
    t = Table(displayName=table_name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = "B2"

def write_overview(wb, title, df_all, meta):
    ws = wb.create_sheet(title)

    for col, w in {"A": 3, "B": 32, "C": 14, "D": 10, "E": 42}.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells("B2:E2")
    ws["B2"] = "Measurement Points Comparison Report"
    ws["B2"].font = Font(name=FONT_NAME, size=16, bold=True, color=PALETTE["TITLE"])
    ws.row_dimensions[2].height = 28

    ws.merge_cells("B3:E3")
    ws["B3"] = (f"Workbook: {meta['workbook']}   |   "
                f"{datetime.now().strftime('%Y-%m-%d  %H:%M')}")
    ws["B3"].font = Font(name=FONT_NAME, size=10, italic=True, color="888888")
    ws.row_dimensions[3].height = 18

    ws.merge_cells("B4:E4")
    ws["B4"] = (f"Sheet 1: {meta['orig_sheet']} ({meta['n_orig']} rows)   |   "
                f"Sheet 2: {meta['new_sheet']} ({meta['n_new']} rows)   |   "
                f"XYZ tolerance: {meta['tol']} mm   |   "
                f"IJK: {'on' if meta['use_ijk'] else 'off'}")
    ws["B4"].font = Font(name=FONT_NAME, size=10, italic=True, color="AAAAAA")
    ws.row_dimensions[4].height = 16

    hrow = 6
    for ci, h in enumerate(["Status", "Count", "Color", "Description"], 2):
        ap(ws.cell(hrow, ci, h), hdr_style())
    ws.row_dimensions[hrow].height = 24

    vc = df_all["Status"].value_counts()
    for ri, code in enumerate(["MATCH","NAME_CHANGED","COORD_CHANGED","DELETED","ADDED"], hrow+1):
        cnt = vc.get(code, 0)
        bg  = PALETTE.get(code, "FFFFFF")
        ws.row_dimensions[ri].height = 22
        for ci, val in enumerate([code, cnt, "", STATUS_LABEL[code]], 2):
            c = ws.cell(ri, ci, val)
            c.fill   = PatternFill("solid", fgColor=bg)
            c.font   = Font(name=FONT_NAME, size=11, bold=(ci in (2, 3)))
            c.border = thin()
            c.alignment = Alignment(
                horizontal="center" if ci in (2, 3, 4) else "left",
                vertical="center")

    tr = hrow + 6
    ws.row_dimensions[tr].height = 22
    for ci, val in enumerate(["TOTAL", len(df_all), "", "Total records"], 2):
        c = ws.cell(tr, ci, val)
        c.font   = Font(name=FONT_NAME, size=11, bold=True)
        c.border = thin()
        c.alignment = Alignment(
            horizontal="center" if ci in (2, 3, 4) else "left",
            vertical="center")

    lr = tr + 3
    ws.merge_cells(f"B{lr}:E{lr}")
    ws[f"B{lr}"] = "Orange text in diff columns means the value is not zero"
    ws[f"B{lr}"].font = Font(name=FONT_NAME, size=10, italic=True,
                              color=PALETTE["DIFF_CELL"])
    ws.row_dimensions[lr].height = 18

def delete_report_sheets(wb, prefix):
    for suffix in REPORT_SUFFIXES:
        name = f"{prefix}{suffix}"
        if name in wb.sheetnames:
            del wb[name]

def write_report_sheets(wb, df_all, meta, prefix=REPORT_PREFIX):
    delete_report_sheets(wb, prefix)
    write_overview(wb, f"{prefix}Overview", df_all, meta)
    write_data_sheet(wb, f"{prefix}All Results",  df_all,                                    "AllResults")
    write_data_sheet(wb, f"{prefix}Match",        df_all[df_all.Status=="MATCH"].reset_index(drop=True),         "Match")
    write_data_sheet(wb, f"{prefix}Name Changed", df_all[df_all.Status=="NAME_CHANGED"].reset_index(drop=True),  "NameChanged")
    write_data_sheet(wb, f"{prefix}Coord Changed", df_all[df_all.Status=="COORD_CHANGED"].reset_index(drop=True), "CoordChanged")
    write_data_sheet(wb, f"{prefix}Deleted",      df_all[df_all.Status=="DELETED"].reset_index(drop=True),       "Deleted")
    write_data_sheet(wb, f"{prefix}Added",        df_all[df_all.Status=="ADDED"].reset_index(drop=True),         "Added")

def build_report(df_all, meta, workbook_path, prefix=REPORT_PREFIX):
    keep_vba = Path(workbook_path).suffix.lower() == ".xlsm"
    wb = load_workbook(workbook_path, keep_vba=keep_vba)
    write_report_sheets(wb, df_all, meta, prefix=prefix)
    wb.save(workbook_path)

def find_existing_append_block(ws, headers):
    row_values = []
    header_row = get_original_sheet_header_row(ws)
    for col_idx in range(1, ws.max_column + 1):
        value = ws.cell(header_row, col_idx).value
        row_values.append("" if value is None else str(value).strip())

    block_len = len(headers)
    for start_idx in range(0, len(row_values) - block_len + 1):
        if row_values[start_idx:start_idx + block_len] == headers:
            return start_idx + 1
    return None

def get_original_sheet_header_row(ws):
    row1 = ["" if ws.cell(1, col_idx).value is None else str(ws.cell(1, col_idx).value).strip() for col_idx in range(1, ws.max_column + 1)]
    row2 = ["" if ws.cell(2, col_idx).value is None else str(ws.cell(2, col_idx).value).strip() for col_idx in range(1, ws.max_column + 1)]
    if any(value in GROUP_HEADER_LABELS for value in row1) and all(field in row2 for field in LOGICAL_FIELDS):
        return 2
    return 1

def ensure_original_sheet_group_row(ws):
    header_row = get_original_sheet_header_row(ws)
    if header_row == 1:
        ws.insert_rows(1)
        header_row = 2
    return header_row

def clear_existing_output_rows(ws, data_start_row, original_data_rows, source_end_col):
    row_idx = data_start_row + original_data_rows
    while row_idx <= ws.max_row:
        has_original_data = any(
            ws.cell(row_idx, col_idx).value not in (None, "")
            for col_idx in range(1, source_end_col + 1)
        )
        if has_original_data:
            break
        ws.delete_rows(row_idx, 1)

def apply_group_labels(ws, source_end_col, status_col, new_start_col, new_end_col, diff_start_col, diff_end_col):
    for merge_range in list(ws.merged_cells.ranges):
        if merge_range.min_row == 1 and merge_range.max_row == 1:
            ws.unmerge_cells(str(merge_range))

    def write_group(start_col, end_col, label):
        if start_col > end_col:
            return
        if start_col == end_col:
            cell = ws.cell(1, start_col, label)
            ap(cell, original_sheet_group_style())
            return
        ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
        cell = ws.cell(1, start_col, label)
        ap(cell, original_sheet_group_style())

    write_group(1, source_end_col, "Original data")
    cell = ws.cell(1, status_col, "STATUS")
    ap(cell, original_sheet_group_style())
    write_group(new_start_col, new_end_col, "New Data")
    write_group(diff_start_col, diff_end_col, "Difference")

def append_results_to_original_sheet(wb, sheet_name, df_all, n_orig):
    ws = wb[sheet_name]
    header_row = ensure_original_sheet_group_row(ws)
    headers = [header for _, header in ORIGINAL_SHEET_APPEND_MAP]
    start_col = find_existing_append_block(ws, headers)
    source_end_col = start_col - 1 if start_col is not None else ws.max_column
    if start_col is None:
        start_col = source_end_col + 1
    end_col = start_col + len(headers) - 1
    data_start_row = header_row + 1

    for col_idx in range(start_col, end_col + 1):
        for row_idx in range(header_row, ws.max_row + 1):
            cell = ws.cell(row_idx, col_idx)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()
            cell.alignment = Alignment(horizontal="general", vertical="bottom")
            cell.font = Font(name=FONT_NAME, size=10)

    clear_existing_output_rows(ws, data_start_row, n_orig, source_end_col)

    for offset, (_, header) in enumerate(ORIGINAL_SHEET_APPEND_MAP):
        cell = ws.cell(header_row, start_col + offset, header)
        ap(cell, original_sheet_header_style())

    orig_rows = df_all.iloc[:n_orig].reset_index(drop=True)
    for row_idx, (_, result_row) in enumerate(orig_rows.iterrows(), start=data_start_row):
        for offset, (source_col, _) in enumerate(ORIGINAL_SHEET_APPEND_MAP):
            value = result_row[source_col] if source_col in result_row else ""
            if isinstance(value, float) and np.isnan(value):
                value = ""
            is_diff = source_col.endswith("_diff") or source_col == "DIFF_Fields"
            cell = ws.cell(row_idx, start_col + offset, value)
            ap(cell, original_sheet_data_style(is_diff=is_diff and value not in ("", 0, 0.0)))
            if source_col in ORIGINAL_SHEET_NUMERIC_COLUMNS:
                cell.number_format = "0.0000"

    added_rows = df_all[df_all.Status == "ADDED"].reset_index(drop=True)
    append_row_idx = data_start_row + n_orig
    for _, added_row in added_rows.iterrows():
        for offset, (source_col, _) in enumerate(ORIGINAL_SHEET_APPEND_MAP):
            value = ""
            if source_col == "Status":
                value = "NEW"
            elif source_col.startswith("NEW_"):
                value = added_row.get(source_col, "")
                if isinstance(value, float) and np.isnan(value):
                    value = ""
            cell = ws.cell(append_row_idx, start_col + offset, value)
            is_diff = source_col.endswith("_diff") or source_col == "DIFF_Fields"
            ap(cell, original_sheet_data_style(is_diff=is_diff and value not in ("", 0, 0.0)))
            if source_col in ORIGINAL_SHEET_NUMERIC_COLUMNS:
                cell.number_format = "0.0000"
        append_row_idx += 1

    for offset, (_, header) in enumerate(ORIGINAL_SHEET_APPEND_MAP):
        col_idx = start_col + offset
        sample_values = [str(header)]
        sample_values.extend(
            str(ws.cell(row_idx, col_idx).value)
            for row_idx in range(data_start_row, ws.max_row + 1)
            if ws.cell(row_idx, col_idx).value not in (None, "")
        )
        width = min(max((len(value) + 2) for value in sample_values), 24)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    apply_group_labels(
        ws,
        source_end_col=source_end_col,
        status_col=start_col,
        new_start_col=start_col + 1,
        new_end_col=start_col + 7,
        diff_start_col=start_col + 8,
        diff_end_col=end_col,
    )


# ══════════════════════════════════════════════════════════════════════════════
#  SHARED EXECUTION FLOW
# ══════════════════════════════════════════════════════════════════════════════

def run_comparison(
    workbook_path,
    orig_sheet,
    new_sheet,
    tol,
    ijk_tol,
    use_ijk,
    create_report_sheets=True,
):
    """Run the workbook comparison and write report sheets back into the workbook."""
    df_orig_raw = read_sheet(workbook_path, orig_sheet)
    df_new_raw = read_sheet(workbook_path, new_sheet)

    df_orig = prepare_sheet(df_orig_raw, orig_sheet)
    df_new = prepare_sheet(df_new_raw, new_sheet)
    df_all = compare(df_orig, df_new, tol, ijk_tol, use_ijk)

    meta = {
        "workbook": Path(workbook_path).name,
        "orig_sheet": orig_sheet,
        "new_sheet": new_sheet,
        "n_orig": len(df_orig),
        "n_new": len(df_new),
        "tol": tol,
        "ijk_tol": ijk_tol,
        "use_ijk": use_ijk,
    }
    keep_vba = Path(workbook_path).suffix.lower() == ".xlsm"
    wb = load_workbook(workbook_path, keep_vba=keep_vba)
    append_results_to_original_sheet(wb, orig_sheet, df_all, len(df_orig))
    if create_report_sheets:
        write_report_sheets(wb, df_all, meta)
    wb.save(workbook_path)
    return {
        "df_all": df_all,
        "meta": meta,
        "orig_columns": list(df_orig_raw.columns),
        "new_columns": list(df_new_raw.columns),
        "create_report_sheets": create_report_sheets,
    }


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

    print()
    hr()
    print("  Normalizing data ...")
    try:
        result = run_comparison(workbook_path, orig_sheet, new_sheet, tol, ijk_tol, use_ijk)
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
    for code in ["MATCH","NAME_CHANGED","COORD_CHANGED","DELETED","ADDED"]:
        print(f"  {STATUS_LABEL[code]:<40s}: {vc.get(code, 0):>5d}")
    print(f"  {'TOTAL':<40s}: {len(df_all):>5d}")
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
