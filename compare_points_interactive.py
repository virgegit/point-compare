#!/usr/bin/env python3
"""
compare_points_interactive.py
==============================
Интерактивный скрипт сравнения двух списков измерительных точек.
Спрашивает у пользователя файлы и настройки, затем генерирует Excel-отчёт.

Поддерживаемые форматы входных файлов:
  - CSV  (.csv)
  - Excel (.xlsx, .xls, .xlsm)

Категории результата:
  MATCH          – имя и координаты совпадают
  NAME_CHANGED   – те же координаты, другое имя
  COORD_CHANGED  – то же имя, другие координаты  →  показываем dX/dY/dZ
  DELETED        – есть только в оригинальном списке
  ADDED          – есть только в новом списке
"""

import sys
import os
from pathlib import Path
from datetime import datetime

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# ══════════════════════════════════════════════════════════════════════════════
#  НАСТРОЙКИ ПО УМОЛЧАНИЮ  (можно менять здесь или через интерактивный ввод)
# ══════════════════════════════════════════════════════════════════════════════
DEFAULTS = {
    "coord_tol":   0.05,   # допуск XYZ, мм
    "ijk_tol":     0.001,  # допуск IJK
    "compare_ijk": False,  # учитывать IJK при сравнении координат?
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
    "MATCH":         "✔  Совпадение",
    "NAME_CHANGED":  "✎  Изменилось имя",
    "COORD_CHANGED": "⚠  Изменились координаты",
    "DELETED":       "✖  Удалена  (только в оригинале)",
    "ADDED":         "＋ Добавлена (только в новом)",
}

LOGICAL_FIELDS = ["Name", "X", "Y", "Z", "I", "J", "K"]
FONT_NAME = "Calibri"

# ══════════════════════════════════════════════════════════════════════════════
#  УТИЛИТЫ
# ══════════════════════════════════════════════════════════════════════════════

def hr(char="─", n=60):
    print(char * n)

def ask(prompt, default=None, validator=None):
    """Спрашивает пользователя, возвращает строку."""
    hint = f" [{default}]" if default is not None else ""
    while True:
        raw = input(f"{prompt}{hint}: ").strip()
        if raw == "" and default is not None:
            return str(default)
        if raw == "" and default is None:
            print("  ⚠  Значение обязательно.")
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
    return raw in ("y", "yes", "да", "д")

def ask_float(prompt, default):
    while True:
        raw = input(f"{prompt} [{default}]: ").strip()
        if raw == "":
            return default
        try:
            return float(raw)
        except ValueError:
            print("  ⚠  Введите число, например 0.05")

def validate_file(path_str):
    p = Path(path_str)
    if not p.exists():
        return False, f"Файл не найден: {p}"
    if p.suffix.lower() not in (".csv", ".xlsx", ".xls", ".xlsm"):
        return False, "Поддерживаются: .csv, .xlsx, .xls, .xlsm"
    return True, ""

def list_sheets(path):
    """Возвращает список листов для Excel-файла."""
    p = Path(path)
    if p.suffix.lower() == ".csv":
        return []
    try:
        xl = pd.ExcelFile(path)
        return xl.sheet_names
    except Exception:
        return []

def read_file(path, sheet=None):
    """Читает файл, возвращает DataFrame."""
    p = Path(path)
    if p.suffix.lower() == ".csv":
        with open(p, encoding="utf-8-sig", errors="replace") as f:
            sample = f.read(2048)
        sep = ";" if sample.count(";") > sample.count(",") else ","
        return pd.read_csv(p, sep=sep, encoding="utf-8-sig")
    else:
        kw = {"sheet_name": sheet} if sheet else {"sheet_name": 0}
        return pd.read_excel(p, **kw)

def choose_column(df, field_hint, file_label, allow_skip=False):
    """Интерактивно выбирает колонку из DataFrame."""
    cols = list(df.columns)
    hints = {
        "Name": ["name", "label", "pointname", "point", "id", "имя", "название", "обозначение"],
        "X":    ["x", "x_mm", "cx"],
        "Y":    ["y", "y_mm", "cy"],
        "Z":    ["z", "z_mm", "cz"],
        "I":    ["i", "vx", "nx", "ix"],
        "J":    ["j", "vy", "ny", "jy"],
        "K":    ["k", "vz", "nz", "kz"],
    }
    guess = None
    for col in cols:
        if col.lower() in hints.get(field_hint, []):
            guess = col
            break

    print(f"\n  Поле «{field_hint}» для {file_label}:")
    for i, col in enumerate(cols, 1):
        marker = " ← авто" if col == guess else ""
        print(f"    {i:2d}. {col}{marker}")
    if allow_skip:
        print(f"     0. Пропустить (нет этой колонки)")

    hint_str = f" [{guess}]" if guess else (" [0-пропустить]" if allow_skip else "")
    while True:
        raw = input(f"  Выберите номер или имя колонки{hint_str}: ").strip()
        if raw == "" and guess:
            return guess
        if raw == "" and allow_skip:
            return None
        if raw == "0" and allow_skip:
            return None
        try:
            idx = int(raw) - 1
            if 0 <= idx < len(cols):
                return cols[idx]
        except ValueError:
            pass
        if raw in cols:
            return raw
        print("  ⚠  Не найдено. Введите номер из списка выше.")


# ══════════════════════════════════════════════════════════════════════════════
#  ЛОГИКА СРАВНЕНИЯ
# ══════════════════════════════════════════════════════════════════════════════

def normalize(df, col_map):
    """Приводит DataFrame к стандартным полям Name/X/Y/Z/I/J/K."""
    out = {}
    for logical, actual in col_map.items():
        if actual and actual in df.columns:
            if logical == "Name":
                out[logical] = df[actual].astype(str).str.strip()
            else:
                out[logical] = pd.to_numeric(df[actual], errors="coerce")
        else:
            out[logical] = np.nan
    return pd.DataFrame(out)

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
#  EXCEL ОТЧЁТ
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

def write_data_sheet(wb, title, df, table_name):
    ws = wb.create_sheet(title)
    if df.empty:
        ws.cell(1, 1, "Нет данных").font = Font(italic=True)
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

def write_overview(wb, df_all, meta):
    ws = wb.active
    ws.title = "Overview"

    for col, w in {"A": 3, "B": 32, "C": 14, "D": 10, "E": 42}.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells("B2:E2")
    ws["B2"] = "Отчёт: сравнение измерительных точек"
    ws["B2"].font = Font(name=FONT_NAME, size=16, bold=True, color=PALETTE["TITLE"])
    ws.row_dimensions[2].height = 28

    ws.merge_cells("B3:E3")
    ws["B3"] = (f"Оригинал: {meta['orig_file']}   |   "
                f"Новый: {meta['new_file']}   |   "
                f"{datetime.now().strftime('%Y-%m-%d  %H:%M')}")
    ws["B3"].font = Font(name=FONT_NAME, size=10, italic=True, color="888888")
    ws.row_dimensions[3].height = 18

    ws.merge_cells("B4:E4")
    ws["B4"] = (f"Точек в оригинале: {meta['n_orig']}   |   "
                f"Точек в новом: {meta['n_new']}   |   "
                f"Допуск XYZ: {meta['tol']} мм   |   "
                f"IJK: {'вкл.' if meta['use_ijk'] else 'выкл.'}")
    ws["B4"].font = Font(name=FONT_NAME, size=10, italic=True, color="AAAAAA")
    ws.row_dimensions[4].height = 16

    hrow = 6
    for ci, h in enumerate(["Статус", "Кол-во", "Цвет", "Описание"], 2):
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
    for ci, val in enumerate(["TOTAL", len(df_all), "", "Всего записей"], 2):
        c = ws.cell(tr, ci, val)
        c.font   = Font(name=FONT_NAME, size=11, bold=True)
        c.border = thin()
        c.alignment = Alignment(
            horizontal="center" if ci in (2, 3, 4) else "left",
            vertical="center")

    lr = tr + 3
    ws.merge_cells(f"B{lr}:E{lr}")
    ws[f"B{lr}"] = "Оранжевый текст в diff-колонках = значение отличается от нуля"
    ws[f"B{lr}"].font = Font(name=FONT_NAME, size=10, italic=True,
                              color=PALETTE["DIFF_CELL"])
    ws.row_dimensions[lr].height = 18

def build_report(df_all, meta, out_path):
    wb = Workbook()
    write_overview(wb, df_all, meta)
    write_data_sheet(wb, "All Results",     df_all,                                    "AllResults")
    write_data_sheet(wb, "✔ Match",         df_all[df_all.Status=="MATCH"].reset_index(drop=True),         "Match")
    write_data_sheet(wb, "✎ Name Changed",  df_all[df_all.Status=="NAME_CHANGED"].reset_index(drop=True),  "NameChanged")
    write_data_sheet(wb, "⚠ Coord Changed", df_all[df_all.Status=="COORD_CHANGED"].reset_index(drop=True), "CoordChanged")
    write_data_sheet(wb, "✖ Deleted",       df_all[df_all.Status=="DELETED"].reset_index(drop=True),       "Deleted")
    write_data_sheet(wb, "＋ Added",         df_all[df_all.Status=="ADDED"].reset_index(drop=True),         "Added")
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


# ══════════════════════════════════════════════════════════════════════════════
#  ИНТЕРАКТИВНЫЙ ДИАЛОГ
# ══════════════════════════════════════════════════════════════════════════════

def interactive():
    hr("═")
    print("  СРАВНЕНИЕ ИЗМЕРИТЕЛЬНЫХ ТОЧЕК")
    hr("═")
    print()

    hr()
    print("  ШАГ 1 из 5 — ОРИГИНАЛЬНЫЙ ФАЙЛ")
    hr()
    orig_path = ask("  Путь к оригинальному файлу", validator=validate_file)

    orig_sheet = None
    orig_sheets = list_sheets(orig_path)
    if orig_sheets:
        print(f"\n  Листы в файле: {', '.join(str(s) for s in orig_sheets)}")
        orig_sheet = ask("  Имя листа", default=str(orig_sheets[0]))
        orig_sheet = orig_sheet if orig_sheet in orig_sheets else orig_sheets[0]

    print("\n  Читаю оригинальный файл …")
    df_orig_raw = read_file(orig_path, orig_sheet)
    df_orig_raw.columns = [str(c).strip() for c in df_orig_raw.columns]
    print(f"  Загружено строк: {len(df_orig_raw)}  |  Колонки: {list(df_orig_raw.columns)}")

    print()
    hr()
    print("  ШАГ 2 из 5 — НОВЫЙ ФАЙЛ")
    hr()
    new_path = ask("  Путь к новому файлу", validator=validate_file)

    new_sheet = None
    new_sheets = list_sheets(new_path)
    if new_sheets:
        print(f"\n  Листы в файле: {', '.join(str(s) for s in new_sheets)}")
        new_sheet = ask("  Имя листа", default=str(new_sheets[0]))
        new_sheet = new_sheet if new_sheet in new_sheets else new_sheets[0]

    print("\n  Читаю новый файл …")
    df_new_raw = read_file(new_path, new_sheet)
    df_new_raw.columns = [str(c).strip() for c in df_new_raw.columns]
    print(f"  Загружено строк: {len(df_new_raw)}  |  Колонки: {list(df_new_raw.columns)}")

    print()
    hr()
    print("  ШАГ 3 из 5 — МАППИНГ КОЛОНОК")
    hr()
    print("  Укажите, какие колонки соответствуют полям Name/X/Y/Z/I/J/K")
    print("  (I, J, K можно пропустить если их нет)")

    col_map_orig = {}
    col_map_new  = {}
    for field in LOGICAL_FIELDS:
        allow_skip = field in ("I","J","K")
        col_map_orig[field] = choose_column(df_orig_raw, field, "ОРИГИНАЛ", allow_skip)
        col_map_new[field]  = choose_column(df_new_raw,  field, "НОВЫЙ",    allow_skip)

    print()
    hr()
    print("  ШАГ 4 из 5 — НАСТРОЙКИ СРАВНЕНИЯ")
    hr()
    tol     = ask_float("  Допуск XYZ (мм)", DEFAULTS["coord_tol"])
    ijk_tol = ask_float("  Допуск IJK",       DEFAULTS["ijk_tol"])
    use_ijk = ask_yn("  Учитывать I/J/K при сравнении координат?", DEFAULTS["compare_ijk"])

    print()
    hr()
    print("  ШАГ 5 из 5 — ВЫХОДНОЙ ФАЙЛ")
    hr()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    default_out = f"comparison_report_{ts}.xlsx"
    out_path = ask("  Путь к файлу отчёта (.xlsx)", default=default_out)
    if not out_path.lower().endswith(".xlsx"):
        out_path += ".xlsx"

    print()
    hr()
    print("  Нормализую данные …")
    df_orig = normalize(df_orig_raw, col_map_orig)
    df_new  = normalize(df_new_raw,  col_map_new)

    print("  Сравниваю точки …")
    df_all = compare(df_orig, df_new, tol, ijk_tol, use_ijk)

    vc = df_all["Status"].value_counts()
    print()
    hr()
    print("  РЕЗУЛЬТАТЫ")
    hr()
    for code in ["MATCH","NAME_CHANGED","COORD_CHANGED","DELETED","ADDED"]:
        print(f"  {STATUS_LABEL[code]:<40s}: {vc.get(code, 0):>5d}")
    print(f"  {'ИТОГО':<40s}: {len(df_all):>5d}")
    hr()

    meta = {
        "orig_file": Path(orig_path).name,
        "new_file":  Path(new_path).name,
        "n_orig":    len(df_orig),
        "n_new":     len(df_new),
        "tol":       tol,
        "ijk_tol":   ijk_tol,
        "use_ijk":   use_ijk,
    }
    print(f"\n  Генерирую отчёт: {out_path} …")
    build_report(df_all, meta, out_path)
    print(f"  ✅ Готово!  →  {Path(out_path).resolve()}")
    hr("═")

    return df_all, out_path


# ══════════════════════════════════════════════════════════════════════════════
#  ТОЧКА ВХОДА
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    try:
        interactive()
    except KeyboardInterrupt:
        print("\n\n  Прервано пользователем.")
        sys.exit(0)
