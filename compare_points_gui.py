#!/usr/bin/env python3
"""
compare_points_gui.py
=====================
Desktop GUI for comparing two worksheets inside one Excel workbook.
Uses the same comparison and report-writing logic as the interactive CLI.
"""

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from compare_points_interactive import (
    DEFAULTS,
    STATUS_LABEL,
    list_sheets,
    run_comparison,
    validate_workbook,
)


class PointCompareApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Point Compare")
        self.root.geometry("860x700")
        self.root.minsize(760, 620)

        self.workbook_var = tk.StringVar()
        self.orig_sheet_var = tk.StringVar()
        self.new_sheet_var = tk.StringVar()
        self.coord_tol_var = tk.StringVar(value=str(DEFAULTS["coord_tol"]))
        self.ijk_tol_var = tk.StringVar(value=str(DEFAULTS["ijk_tol"]))
        self.compare_ijk_var = tk.BooleanVar(value=DEFAULTS["compare_ijk"])
        self.create_report_sheets_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Select a workbook to begin.")
        self.summary_vars = {
            code: tk.StringVar(value="0")
            for code in ["MATCH", "NAME_CHANGED", "COORD_CHANGED", "DELETED", "ADDED", "TOTAL"]
        }

        self.sheet_names = []
        self._build_ui()

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        outer = ttk.Frame(self.root, padding=16)
        outer.grid(sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(4, weight=1)

        title = ttk.Label(outer, text="Measurement Points Comparison", font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, sticky="w")

        subtitle = ttk.Label(
            outer,
            text="Compare two sheets in one workbook and write CMP_* result sheets back into the same file.",
        )
        subtitle.grid(row=1, column=0, sticky="w", pady=(4, 14))

        self._build_workbook_frame(outer)
        self._build_settings_frame(outer)
        self._build_results_frame(outer)
        self._build_log_frame(outer)

    def _build_workbook_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Workbook", padding=12)
        frame.grid(row=2, column=0, sticky="ew")
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, weight=0)
        frame.columnconfigure(3, weight=0)

        ttk.Label(frame, text="Excel workbook").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
        entry = ttk.Entry(frame, textvariable=self.workbook_var)
        entry.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        entry.bind("<FocusOut>", lambda _event: self.refresh_sheets())

        browse_btn = ttk.Button(frame, text="Browse...", command=self.choose_workbook)
        browse_btn.grid(row=0, column=2, sticky="ew", padx=(8, 8), pady=(0, 8))

        refresh_btn = ttk.Button(frame, text="Load sheets", command=self.refresh_sheets)
        refresh_btn.grid(row=0, column=3, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Original sheet").grid(row=1, column=0, sticky="w", padx=(0, 8))
        self.orig_sheet_combo = ttk.Combobox(
            frame,
            textvariable=self.orig_sheet_var,
            state="readonly",
            values=[],
        )
        self.orig_sheet_combo.grid(row=1, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="New sheet").grid(row=2, column=0, sticky="w", padx=(0, 8))
        self.new_sheet_combo = ttk.Combobox(
            frame,
            textvariable=self.new_sheet_var,
            state="readonly",
            values=[],
        )
        self.new_sheet_combo.grid(row=2, column=1, sticky="ew")

        ttk.Label(
            frame,
            text="Required headers on both sheets: Name, X, Y, Z, I, J, K",
        ).grid(row=3, column=0, columnspan=4, sticky="w", pady=(10, 0))

    def _build_settings_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Comparison Settings", padding=12)
        frame.grid(row=3, column=0, sticky="ew", pady=(12, 12))
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="XYZ tolerance (mm)").grid(row=0, column=0, sticky="w", padx=(0, 12))
        ttk.Entry(frame, textvariable=self.coord_tol_var, width=16).grid(row=0, column=1, sticky="w")

        ttk.Label(frame, text="IJK tolerance").grid(row=1, column=0, sticky="w", padx=(0, 12), pady=(8, 0))
        ttk.Entry(frame, textvariable=self.ijk_tol_var, width=16).grid(row=1, column=1, sticky="w", pady=(8, 0))

        ttk.Checkbutton(
            frame,
            text="Include I/J/K in coordinate comparison",
            variable=self.compare_ijk_var,
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))

        ttk.Checkbutton(
            frame,
            text="Also create CMP_* report sheets",
            variable=self.create_report_sheets_var,
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))

        actions = ttk.Frame(frame)
        actions.grid(row=0, column=2, rowspan=4, sticky="e")
        self.run_button = ttk.Button(actions, text="Run Comparison", command=self.run_compare)
        self.run_button.grid(row=0, column=0, sticky="ew")

        ttk.Label(frame, textvariable=self.status_var, foreground="#1f4e79").grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(12, 0)
        )

    def _build_results_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Summary", padding=12)
        frame.grid(row=4, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        summary_grid = ttk.Frame(frame)
        summary_grid.grid(row=0, column=0, sticky="ew")
        for index in range(3):
            summary_grid.columnconfigure(index, weight=1)

        labels = ["MATCH", "NAME_CHANGED", "COORD_CHANGED", "DELETED", "ADDED", "TOTAL"]
        for index, code in enumerate(labels):
            cell = ttk.Frame(summary_grid, padding=8)
            cell.grid(row=index // 3, column=index % 3, sticky="ew")
            ttk.Label(cell, text=code, font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
            ttk.Label(cell, textvariable=self.summary_vars[code], font=("Segoe UI", 14)).grid(
                row=1, column=0, sticky="w"
            )

        self.result_box = ScrolledText(frame, height=14, wrap="word", state="disabled")
        self.result_box.grid(row=1, column=0, sticky="nsew", pady=(12, 0))

    def _build_log_frame(self, parent):
        footer = ttk.Frame(parent)
        footer.grid(row=5, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        ttk.Label(
            footer,
            text="Tip: existing CMP_* sheets in the workbook will be replaced when you run a new comparison.",
        ).grid(row=0, column=0, sticky="w")

    def choose_workbook(self):
        path = filedialog.askopenfilename(
            title="Select workbook",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not path:
            return
        self.workbook_var.set(path)
        self.refresh_sheets()

    def refresh_sheets(self):
        workbook_path = self.workbook_var.get().strip()
        if not workbook_path:
            return

        ok, message = validate_workbook(workbook_path)
        if not ok:
            self.status_var.set(message)
            self._set_sheet_names([])
            return

        sheets = list_sheets(workbook_path)
        if not sheets:
            self.status_var.set("Unable to read sheet names from the workbook.")
            self._set_sheet_names([])
            return

        self._set_sheet_names(sheets)
        self.status_var.set(f"Loaded {len(sheets)} sheet(s) from {Path(workbook_path).name}.")

    def _set_sheet_names(self, sheets):
        self.sheet_names = list(sheets)
        self.orig_sheet_combo["values"] = self.sheet_names
        self.new_sheet_combo["values"] = self.sheet_names

        if not self.sheet_names:
            self.orig_sheet_var.set("")
            self.new_sheet_var.set("")
            return

        if self.orig_sheet_var.get() not in self.sheet_names:
            self.orig_sheet_var.set(self.sheet_names[0])

        default_new = self.sheet_names[1] if len(self.sheet_names) > 1 else self.sheet_names[0]
        if self.new_sheet_var.get() not in self.sheet_names:
            self.new_sheet_var.set(default_new)

    def run_compare(self):
        workbook_path = self.workbook_var.get().strip()
        orig_sheet = self.orig_sheet_var.get().strip()
        new_sheet = self.new_sheet_var.get().strip()

        ok, message = validate_workbook(workbook_path)
        if not ok:
            messagebox.showerror("Invalid workbook", message)
            return

        if not self.sheet_names:
            self.refresh_sheets()
        if not self.sheet_names:
            messagebox.showerror("Missing sheets", "Load a workbook with readable sheet names first.")
            return

        if orig_sheet not in self.sheet_names or new_sheet not in self.sheet_names:
            messagebox.showerror("Invalid sheet selection", "Choose both source sheets from the workbook list.")
            return

        try:
            coord_tol = float(self.coord_tol_var.get().strip())
            ijk_tol = float(self.ijk_tol_var.get().strip())
        except ValueError:
            messagebox.showerror("Invalid tolerance", "Enter numeric values for XYZ and IJK tolerance.")
            return

        if coord_tol <= 0 or ijk_tol <= 0:
            messagebox.showerror("Invalid tolerance", "Tolerance values must be greater than zero.")
            return

        self.run_button.configure(state="disabled")
        self.status_var.set("Running comparison and writing CMP_* sheets...")
        self.root.update_idletasks()

        try:
            result = run_comparison(
                workbook_path=workbook_path,
                orig_sheet=orig_sheet,
                new_sheet=new_sheet,
                tol=coord_tol,
                ijk_tol=ijk_tol,
                use_ijk=self.compare_ijk_var.get(),
                create_report_sheets=self.create_report_sheets_var.get(),
            )
        except Exception as exc:
            self.status_var.set("Comparison failed.")
            messagebox.showerror("Comparison failed", str(exc))
            return
        finally:
            self.run_button.configure(state="normal")

        self._update_summary(result)
        if result["create_report_sheets"]:
            self.status_var.set(f"Finished. The original sheet and CMP_* sheets were updated in {Path(workbook_path).name}.")
        else:
            self.status_var.set(f"Finished. The original sheet was updated in {Path(workbook_path).name}.")

        if result["create_report_sheets"]:
            messagebox.showinfo("Done", "Comparison finished. The original sheet and new CMP_* sheets were updated.")
        else:
            messagebox.showinfo("Done", "Comparison finished. The original sheet was updated.")

    def _update_summary(self, result):
        df_all = result["df_all"]
        meta = result["meta"]
        counts = df_all["Status"].value_counts()
        for code in ["MATCH", "NAME_CHANGED", "COORD_CHANGED", "DELETED", "ADDED"]:
            self.summary_vars[code].set(str(int(counts.get(code, 0))))
        self.summary_vars["TOTAL"].set(str(len(df_all)))

        lines = [
            f"Workbook: {meta['workbook']}",
            f"Original sheet: {meta['orig_sheet']} ({meta['n_orig']} rows)",
            f"New sheet: {meta['new_sheet']} ({meta['n_new']} rows)",
            f"XYZ tolerance: {meta['tol']}",
            f"IJK tolerance: {meta['ijk_tol']}",
            f"Compare IJK: {'Yes' if meta['use_ijk'] else 'No'}",
            f"Create CMP_* report sheets: {'Yes' if result['create_report_sheets'] else 'No'}",
            "",
            "Source columns detected:",
            f"  {meta['orig_sheet']}: {', '.join(str(col) for col in result['orig_columns'])}",
            f"  {meta['new_sheet']}: {', '.join(str(col) for col in result['new_columns'])}",
            "",
            "Original sheet updates:",
            "  Row 1 contains grouped labels: Original data / STATUS / New Data / Difference.",
            "  STATUS + NEW_* + *_diff columns appended after the last used source column.",
            "  New-only points from sheet 2 are appended as extra rows with STATUS = NEW.",
            "  Existing colors in the original source area are left unchanged.",
            "",
            "Status counts:",
        ]
        for code in ["MATCH", "NAME_CHANGED", "COORD_CHANGED", "DELETED", "ADDED"]:
            lines.append(f"  {STATUS_LABEL[code]}: {counts.get(code, 0)}")
        lines.append(f"  TOTAL: {len(df_all)}")
        lines.append("")
        lines.append("Result sheets written:")
        if result["create_report_sheets"]:
            lines.append("  CMP_Overview")
            lines.append("  CMP_All Results")
            lines.append("  CMP_Match")
            lines.append("  CMP_Name Changed")
            lines.append("  CMP_Coord Changed")
            lines.append("  CMP_Deleted")
            lines.append("  CMP_Added")
        else:
            lines.append("  Not created. Only the original sheet was updated.")

        self.result_box.configure(state="normal")
        self.result_box.delete("1.0", tk.END)
        self.result_box.insert("1.0", "\n".join(lines))
        self.result_box.configure(state="disabled")


def main():
    root = tk.Tk()
    PointCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
