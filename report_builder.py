#!/usr/bin/env python3
"""
Report Builder - Desktop GUI Application
------------------------------------------
- Load a messy Excel export, tell it where the real header row is.
- Pick columns to keep, change data types, filter rows by condition,
  delete individual rows by clicking (like Excel), sort by clicking
  column headers or via the Sort tab.
- Save "Output Templates" (target column structures you use often) to
  a local library, then auto-fill one from a loaded sheet with
  automatic column matching + a review screen to fix any guesses.
- Bulk Process: pick many source files at once, map them all against
  one saved template, and either combine everything into one working
  table (which still goes through the same filter/sort/delete tools)
  or export one output file per source file in one go.
- Export the working table to Excel and/or PDF.

Requirements:
    pip install pandas openpyxl reportlab

Run:
    python report_builder.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import os
import re
import json
import glob
import difflib
import traceback

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


DTYPE_OPTIONS = ["Text", "Integer", "Decimal", "Date", "Yes/No (Boolean)"]
OPERATORS = [
    "Equals", "Not Equals", "Contains", "Does Not Contain",
    "Greater Than", "Less Than", "Greater or Equal", "Less or Equal",
    "Is Empty", "Is Not Empty",
]

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(SCRIPT_DIR, "output_templates")
os.makedirs(TEMPLATES_DIR, exist_ok=True)


# ============================================================ pure helpers
# (no tkinter dependency - kept separate so the mapping/parsing logic can
#  be tested and reused by both the single-file and bulk-process flows)

def normalize_col(name):
    return re.sub(r"[^A-Z0-9]", "", str(name).upper())


def guess_header_row(raw, n_check=20):
    n_check = min(n_check, len(raw))
    best_idx, best_score = 0, -1
    for i in range(n_check):
        non_null = raw.iloc[i].notna().sum()
        below = raw.iloc[i + 1].notna().sum() if i + 1 < len(raw) else 0
        score = non_null + (0.3 * min(non_null, below))
        if score > best_score:
            best_score = score
            best_idx = i
    return best_idx


def dedupe_columns(columns):
    seen = {}
    out = []
    for c in columns:
        if c in seen:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            out.append(c)
    return out


def build_df_from_raw(raw, header_idx):
    """Turn a header-less raw sheet into a proper DataFrame using the
    given row index as the column header, dropping everything above it
    and any fully-empty rows below it."""
    header_vals = raw.iloc[header_idx].tolist()
    columns = [str(v).strip() if not pd.isna(v) else f"Column{i+1}" for i, v in enumerate(header_vals)]
    columns = dedupe_columns(columns)
    df = raw.iloc[header_idx + 1:].copy()
    df.columns = columns
    df = df.reset_index(drop=True)
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def guess_mapping(template_cols, source_cols):
    """For each template column, guess the best-matching source column:
    exact normalized match first, then fuzzy match, else None."""
    norm_source = {normalize_col(s): s for s in source_cols}
    norm_source_keys = list(norm_source.keys())
    mapping = {}
    for tcol in template_cols:
        nt = normalize_col(tcol)
        if nt in norm_source:
            mapping[tcol] = norm_source[nt]
            continue
        candidates = difflib.get_close_matches(nt, norm_source_keys, n=1, cutoff=0.6)
        mapping[tcol] = norm_source[candidates[0]] if candidates else None
    return mapping


def apply_mapping(df_source, template_cols, mapping):
    """Build a new DataFrame with template_cols as columns, pulling data
    from df_source per the mapping dict (template_col -> source_col or None)."""
    new_df = pd.DataFrame()
    n = len(df_source)
    for tcol in template_cols:
        scol = mapping.get(tcol)
        if scol and scol in df_source.columns:
            new_df[tcol] = df_source[scol].values
        else:
            new_df[tcol] = [""] * n
    return new_df


# ================================================================ dialogs

class HeaderRowDialog(tk.Toplevel):
    """Lets the user look at the top of a raw sheet and pick which row
    is the real column-header row."""

    def __init__(self, parent, raw_df, guessed_idx, title="Select the header row"):
        super().__init__(parent)
        self.title(title)
        self.geometry("900x420")
        self.result = None

        ttk.Label(
            self, padding=8,
            text="Click the row that contains the column names.\n"
                 "Rows above it (titles, addresses, filter summaries) will be skipped.",
        ).pack(anchor="w")

        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True, padx=8, pady=4)

        n_preview = min(20, len(raw_df))
        n_cols_preview = min(8, raw_df.shape[1])
        columns = ["Row"] + [f"Col{i+1}" for i in range(n_cols_preview)]

        self.tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse", height=n_preview)
        for c in columns:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=100 if c != "Row" else 50, anchor="w")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        for i in range(n_preview):
            row_vals = raw_df.iloc[i, :n_cols_preview].tolist()
            row_vals = ["" if pd.isna(v) else str(v) for v in row_vals]
            self.tree.insert("", "end", iid=str(i), values=[i] + row_vals)

        if guessed_idx is not None and 0 <= guessed_idx < n_preview:
            self.tree.selection_set(str(guessed_idx))
            self.tree.see(str(guessed_idx))

        btn_row = ttk.Frame(self, padding=8)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="Use Selected Row as Header", command=self._confirm).pack(side="left")
        ttk.Button(btn_row, text="Cancel (use row 1)", command=self._cancel).pack(side="left", padx=8)

        self.transient(parent)
        self.grab_set()

    def _confirm(self):
        sel = self.tree.selection()
        self.result = int(sel[0]) if sel else 0
        self.destroy()

    def _cancel(self):
        self.result = 0
        self.destroy()


class MappingDialog(tk.Toplevel):
    """Review/fix the auto-guessed column mapping between a template's
    columns and a source sheet's columns before it's applied."""

    def __init__(self, parent, template_cols, source_cols, guess):
        super().__init__(parent)
        self.title("Review Column Mapping")
        self.geometry("650x500")
        self.result = None

        ttk.Label(
            self, padding=8,
            text="For each output column, pick which source column supplies the data.\n"
                 "Best guesses are pre-filled - fix any that are wrong, or leave blank.",
        ).pack(anchor="w")

        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="top", fill="both", expand=True, padx=8)
        scrollbar.pack(side="right", fill="y")

        header = ttk.Frame(inner)
        header.pack(fill="x")
        ttk.Label(header, text="Output Column", width=28).pack(side="left")
        ttk.Label(header, text="Source Column").pack(side="left")

        options = ["-- Leave Blank --"] + list(source_cols)
        self.vars = {}
        for tcol in template_cols:
            row = ttk.Frame(inner)
            row.pack(fill="x", pady=1)
            ttk.Label(row, text=str(tcol), width=28).pack(side="left")
            var = tk.StringVar(value=guess.get(tcol) or "-- Leave Blank --")
            cb = ttk.Combobox(row, textvariable=var, state="readonly", values=options, width=35)
            cb.pack(side="left")
            self.vars[tcol] = var

        btn_row = ttk.Frame(self, padding=8)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="Apply Mapping", command=self._confirm).pack(side="left")
        ttk.Button(btn_row, text="Cancel", command=self._cancel).pack(side="left", padx=8)

        self.transient(parent)
        self.grab_set()

    def _confirm(self):
        mapping = {}
        for tcol, var in self.vars.items():
            v = var.get()
            mapping[tcol] = None if v == "-- Leave Blank --" else v
        self.result = mapping
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# =================================================================== app

class ReportBuilderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Report Builder")
        self.root.geometry("1300x900")

        self.df_original = None
        self.df_processed = None
        self.file_path = None

        self.column_vars = {}
        self.dtype_vars = {}
        self.filters = []
        self.sort_keys = []

        self._sort_col = None
        self._sort_asc = True

        self.bulk_files = []          # list of paths
        self.bulk_mapping = None      # dict template_col -> source_col
        self.bulk_template_cols = None

        self.undo_stack = []
        self.redo_stack = []
        self.UNDO_LIMIT = 25

        self._build_ui()
        self._refresh_templates_list()

        self.root.bind_all("<Control-z>", lambda e: self.undo())
        self.root.bind_all("<Control-y>", lambda e: self.redo())
        self.root.bind_all("<Control-Shift-Z>", lambda e: self.redo())

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Upload Excel File", command=self.load_excel).pack(side="left")
        ttk.Button(top, text="Change Header Row...", command=self.change_header_row).pack(side="left", padx=6)
        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(top, text="\u21b6 Undo (Ctrl+Z)", command=self.undo).pack(side="left")
        ttk.Button(top, text="\u21b7 Redo (Ctrl+Y)", command=self.redo).pack(side="left", padx=4)
        self.file_label = ttk.Label(top, text="No file loaded", foreground="gray")
        self.file_label.pack(side="left", padx=10)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=False, padx=8, pady=4)

        self.tab_columns = ttk.Frame(self.notebook)
        self.tab_filters = ttk.Frame(self.notebook)
        self.tab_sort = ttk.Frame(self.notebook)
        self.tab_templates = ttk.Frame(self.notebook)
        self.tab_bulk = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_columns, text="Columns && Data Types")
        self.notebook.add(self.tab_filters, text="Filters (by condition)")
        self.notebook.add(self.tab_sort, text="Sort")
        self.notebook.add(self.tab_templates, text="Output Templates")
        self.notebook.add(self.tab_bulk, text="Bulk Process")

        self._build_columns_tab()
        self._build_filters_tab()
        self._build_sort_tab()
        self._build_templates_tab()
        self._build_bulk_tab()

        apply_bar = ttk.Frame(self.root, padding=8)
        apply_bar.pack(fill="x")
        ttk.Button(apply_bar, text="Apply && Preview", command=self.apply_all).pack(side="left")
        self.status_label = ttk.Label(apply_bar, text="", foreground="green")
        self.status_label.pack(side="left", padx=10)

        preview_frame = ttk.LabelFrame(
            self.root,
            text="Preview - click a column heading to sort, click/Ctrl+click/Shift+click rows then 'Delete Selected Rows'",
            padding=4,
        )
        preview_frame.pack(fill="both", expand=True, padx=8, pady=4)

        row_btns = ttk.Frame(preview_frame)
        row_btns.pack(fill="x", pady=(0, 4))
        ttk.Button(row_btns, text="Delete Selected Rows", command=self.delete_selected_rows).pack(side="left")
        self.row_count_label = ttk.Label(row_btns, text="")
        self.row_count_label.pack(side="left", padx=10)

        tree_container = ttk.Frame(preview_frame)
        tree_container.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_container, show="headings", selectmode="extended")
        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        export_bar = ttk.Frame(self.root, padding=8)
        export_bar.pack(fill="x")
        ttk.Button(export_bar, text="Export to Excel", command=self.export_excel).pack(side="left", padx=4)
        ttk.Button(export_bar, text="Export to PDF", command=self.export_pdf).pack(side="left", padx=4)

    def _build_columns_tab(self):
        container = ttk.Frame(self.tab_columns, padding=6)
        container.pack(fill="both", expand=True)
        ttk.Label(container, text="Select columns to keep and choose a data type for each:").pack(anchor="w")

        canvas = tk.Canvas(container, height=180)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.columns_inner = ttk.Frame(canvas)

        self.columns_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.columns_inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _build_filters_tab(self):
        container = ttk.Frame(self.tab_filters, padding=6)
        container.pack(fill="both", expand=True)

        row = ttk.Frame(container)
        row.pack(fill="x", pady=4)
        ttk.Label(row, text="Column:").pack(side="left")
        self.filter_col_cb = ttk.Combobox(row, state="readonly", width=20)
        self.filter_col_cb.pack(side="left", padx=4)

        ttk.Label(row, text="Operator:").pack(side="left")
        self.filter_op_cb = ttk.Combobox(row, state="readonly", values=OPERATORS, width=18)
        self.filter_op_cb.pack(side="left", padx=4)
        self.filter_op_cb.current(0)

        ttk.Label(row, text="Value:").pack(side="left")
        self.filter_val_entry = ttk.Entry(row, width=20)
        self.filter_val_entry.pack(side="left", padx=4)

        ttk.Button(row, text="Add Filter", command=self.add_filter).pack(side="left", padx=8)

        self.filters_listbox = tk.Listbox(container, height=5)
        self.filters_listbox.pack(fill="x", pady=6)
        ttk.Button(container, text="Remove Selected Filter", command=self.remove_filter).pack(anchor="w")
        ttk.Label(container, text="(These filters are also used by Bulk Process > Export Each File Separately)",
                  foreground="gray").pack(anchor="w", pady=(6, 0))

    def _build_sort_tab(self):
        container = ttk.Frame(self.tab_sort, padding=6)
        container.pack(fill="both", expand=True)
        ttk.Label(container, text="(Tip: click a column heading in the preview table to sort instantly)").pack(anchor="w")

        row = ttk.Frame(container)
        row.pack(fill="x", pady=4)
        ttk.Label(row, text="Column:").pack(side="left")
        self.sort_col_cb = ttk.Combobox(row, state="readonly", width=20)
        self.sort_col_cb.pack(side="left", padx=4)

        ttk.Label(row, text="Order:").pack(side="left")
        self.sort_order_cb = ttk.Combobox(row, state="readonly", values=["Ascending", "Descending"], width=12)
        self.sort_order_cb.pack(side="left", padx=4)
        self.sort_order_cb.current(0)

        ttk.Button(row, text="Add Sort Key", command=self.add_sort_key).pack(side="left", padx=8)

        self.sort_listbox = tk.Listbox(container, height=5)
        self.sort_listbox.pack(fill="x", pady=6)
        ttk.Label(container, text="(Multiple sort keys applied in the order listed - first is primary)").pack(anchor="w")
        ttk.Button(container, text="Remove Selected Sort Key", command=self.remove_sort_key).pack(anchor="w")

    def _build_templates_tab(self):
        container = ttk.Frame(self.tab_templates, padding=6)
        container.pack(fill="both", expand=True)

        ttk.Label(
            container,
            text="Save output structures you use often here, then auto-fill one from whatever sheet you loaded above.",
        ).pack(anchor="w")

        btn_row = ttk.Frame(container)
        btn_row.pack(fill="x", pady=6)
        ttk.Button(btn_row, text="Upload New Template Structure...", command=self.upload_template).pack(side="left")
        ttk.Button(btn_row, text="Delete Selected Template", command=self.delete_template).pack(side="left", padx=6)
        ttk.Button(btn_row, text="Preview Selected Template's Columns", command=self.preview_template_columns).pack(side="left", padx=6)

        self.templates_tree = ttk.Treeview(container, columns=["name", "cols"], show="headings",
                                            selectmode="browse", height=6)
        self.templates_tree.heading("name", text="Template Name")
        self.templates_tree.heading("cols", text="# Columns")
        self.templates_tree.column("name", width=300, anchor="w")
        self.templates_tree.column("cols", width=100, anchor="center")
        self.templates_tree.pack(fill="x", pady=6)

        ttk.Button(
            container, text="Auto-Fill Selected Template From Loaded Data",
            command=self.autofill_template_from_loaded,
        ).pack(anchor="w", pady=(6, 0))
        ttk.Label(
            container,
            text="(This replaces the working data below with the template's columns, matched from your loaded sheet.\n"
                 "You'll get a review screen to fix any column matches before it's applied.)",
            foreground="gray",
        ).pack(anchor="w")

    def _build_bulk_tab(self):
        container = ttk.Frame(self.tab_bulk, padding=6)
        container.pack(fill="both", expand=True)

        ttk.Label(
            container,
            text="Process many source files at once against one saved Output Template.",
        ).pack(anchor="w")

        btn_row = ttk.Frame(container)
        btn_row.pack(fill="x", pady=6)
        ttk.Button(btn_row, text="Select Multiple Excel Files...", command=self.select_bulk_files).pack(side="left")
        ttk.Button(btn_row, text="Remove Selected File", command=self.remove_bulk_file).pack(side="left", padx=6)
        ttk.Button(btn_row, text="Clear All", command=self.clear_bulk_files).pack(side="left", padx=6)

        self.bulk_files_listbox = tk.Listbox(container, height=6)
        self.bulk_files_listbox.pack(fill="x", pady=4)

        template_row = ttk.Frame(container)
        template_row.pack(fill="x", pady=6)
        ttk.Label(template_row, text="Output Template:").pack(side="left")
        self.bulk_template_cb = ttk.Combobox(template_row, state="readonly", width=35)
        self.bulk_template_cb.pack(side="left", padx=6)
        ttk.Button(template_row, text="Refresh List", command=self._refresh_templates_list).pack(side="left")

        ttk.Button(
            container, text="1) Auto-Map Columns (using first file)",
            command=self.bulk_automap,
        ).pack(anchor="w", pady=(8, 2))

        self.bulk_source_col_var = tk.StringVar(value="Add 'Source File' column")
        self.bulk_add_source_col = tk.BooleanVar(value=True)
        ttk.Checkbutton(container, text="Add a 'Source File' column showing which file each row came from",
                         variable=self.bulk_add_source_col).pack(anchor="w")

        action_row = ttk.Frame(container)
        action_row.pack(fill="x", pady=8)
        ttk.Button(
            action_row, text="2a) Combine All Into Working Data (then filter/sort/export as usual)",
            command=self.bulk_combine,
        ).pack(side="left")

        action_row2 = ttk.Frame(container)
        action_row2.pack(fill="x", pady=4)
        self.bulk_export_pdf_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(action_row2, text="Also export a PDF per file", variable=self.bulk_export_pdf_var).pack(side="left")
        ttk.Button(
            action_row2, text="2b) Export Each File Separately (uses current Filters/Sort settings)",
            command=self.bulk_export_separate,
        ).pack(side="left", padx=10)

        self.bulk_log = tk.Text(container, height=8, state="disabled")
        self.bulk_log.pack(fill="both", expand=True, pady=(8, 0))

    def _bulk_log(self, msg):
        self.bulk_log.configure(state="normal")
        self.bulk_log.insert("end", msg + "\n")
        self.bulk_log.see("end")
        self.bulk_log.configure(state="disabled")

    # ------------------------------------------------------------- Loading
    def load_excel(self):
        path = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            raw = pd.read_excel(path, header=None)
        except Exception as e:
            messagebox.showerror("Error loading file", str(e))
            return

        self.file_path = path
        self._raw_df = raw
        guess = guess_header_row(raw)
        self._prompt_header_row(guess)

    def change_header_row(self):
        if not hasattr(self, "_raw_df") or self._raw_df is None:
            messagebox.showwarning("No file", "Upload an Excel file first.")
            return
        guess = guess_header_row(self._raw_df)
        self._prompt_header_row(guess)

    def _prompt_header_row(self, guess):
        dialog = HeaderRowDialog(self.root, self._raw_df, guess)
        self.root.wait_window(dialog)
        header_idx = dialog.result if dialog.result is not None else 0
        self._finalize_load(header_idx)

    def _finalize_load(self, header_idx):
        try:
            df = build_df_from_raw(self._raw_df, header_idx)
        except Exception as e:
            messagebox.showerror("Error parsing header", str(e))
            return
        self._push_undo()
        self._set_working_data(df, label=f"{os.path.basename(self.file_path)}  (header row {header_idx + 1})")
        messagebox.showinfo("Loaded", f"Loaded {len(df)} rows and {len(df.columns)} columns.")

    # ------------------------------------------------------------ Undo/Redo
    def _snapshot(self):
        return {
            "df_original": self.df_original.copy() if self.df_original is not None else None,
            "df_processed": self.df_processed.copy() if self.df_processed is not None else None,
            "file_label": self.file_label.cget("text"),
            "column_selection": {c: v.get() for c, v in self.column_vars.items()},
            "dtype_selection": {c: v.get() for c, v in self.dtype_vars.items()},
            "filters": [dict(f) for f in self.filters],
            "sort_keys": [dict(k) for k in self.sort_keys],
            "sort_col": self._sort_col,
            "sort_asc": self._sort_asc,
        }

    def _push_undo(self):
        """Call this right BEFORE making a change, so the state being
        replaced is saved. Any new action clears the redo stack."""
        if self.df_original is None and self.df_processed is None:
            return
        self.undo_stack.append(self._snapshot())
        if len(self.undo_stack) > self.UNDO_LIMIT:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def _restore_snapshot(self, snap):
        self.df_original = snap["df_original"]
        self.df_processed = snap["df_processed"]
        label = snap["file_label"]
        self.file_label.config(text=label, foreground="gray" if label == "No file loaded" else "black")

        cols = list(self.df_original.columns) if self.df_original is not None else []
        self._populate_column_controls(cols)
        for c, val in snap["column_selection"].items():
            if c in self.column_vars:
                self.column_vars[c].set(val)
        for c, val in snap["dtype_selection"].items():
            if c in self.dtype_vars:
                self.dtype_vars[c].set(val)
        self._populate_combo_values(cols)

        self.filters = [dict(f) for f in snap["filters"]]
        self.filters_listbox.delete(0, "end")
        for f in self.filters:
            lbl = f"{f['column']} {f['operator']} '{f['value']}'" if f["operator"] not in ("Is Empty", "Is Not Empty") \
                else f"{f['column']} {f['operator']}"
            self.filters_listbox.insert("end", lbl)

        self.sort_keys = [dict(k) for k in snap["sort_keys"]]
        self.sort_listbox.delete(0, "end")
        for k in self.sort_keys:
            self.sort_listbox.insert("end", f"{k['column']} - {k['order']}")

        self._sort_col = snap["sort_col"]
        self._sort_asc = snap["sort_asc"]

        if self.df_processed is not None:
            self._render_preview(self.df_processed)
        else:
            self.tree.delete(*self.tree.get_children())
            self.tree["columns"] = []
            self.row_count_label.config(text="")

    def undo(self):
        if not self.undo_stack:
            messagebox.showinfo("Nothing to undo", "No earlier state to go back to.")
            return
        self.redo_stack.append(self._snapshot())
        snap = self.undo_stack.pop()
        self._restore_snapshot(snap)
        self.status_label.config(text="Undid last action.", foreground="blue")

    def redo(self):
        if not self.redo_stack:
            messagebox.showinfo("Nothing to redo", "No undone action to redo.")
            return
        self.undo_stack.append(self._snapshot())
        snap = self.redo_stack.pop()
        self._restore_snapshot(snap)
        self.status_label.config(text="Redid last undone action.", foreground="blue")

    def _set_working_data(self, df, label):
        """Central place that (re)sets the working dataset and refreshes
        every dependent control - used by normal load, template autofill,
        and bulk combine."""
        self.df_original = df
        self.file_label.config(text=label, foreground="black")
        self.status_label.config(text="")
        self._populate_column_controls(df.columns.tolist())
        self._populate_combo_values(df.columns.tolist())
        self.filters = []
        self.sort_keys = []
        self.filters_listbox.delete(0, "end")
        self.sort_listbox.delete(0, "end")
        self.df_processed = df.copy()
        self._sort_col = None
        self._render_preview(self.df_processed)

    def _populate_column_controls(self, columns):
        for widget in self.columns_inner.winfo_children():
            widget.destroy()
        self.column_vars = {}
        self.dtype_vars = {}

        header = ttk.Frame(self.columns_inner)
        header.pack(fill="x")
        ttk.Label(header, text="Keep", width=6).pack(side="left")
        ttk.Label(header, text="Column", width=30).pack(side="left")
        ttk.Label(header, text="Convert to type").pack(side="left")

        for col in columns:
            row = ttk.Frame(self.columns_inner)
            row.pack(fill="x", pady=1)
            var = tk.BooleanVar(value=True)
            ttk.Checkbutton(row, variable=var, width=6).pack(side="left")
            ttk.Label(row, text=str(col), width=30).pack(side="left")
            dtype_var = tk.StringVar(value="Text")
            cb = ttk.Combobox(row, textvariable=dtype_var, state="readonly", values=DTYPE_OPTIONS, width=20)
            cb.pack(side="left")
            self.column_vars[col] = var
            self.dtype_vars[col] = dtype_var

    def _populate_combo_values(self, columns):
        cols = [str(c) for c in columns]
        self.filter_col_cb.configure(values=cols)
        self.sort_col_cb.configure(values=cols)
        if cols:
            self.filter_col_cb.current(0)
            self.sort_col_cb.current(0)

    # ------------------------------------------------------------- Filters
    def add_filter(self):
        if self.df_original is None:
            messagebox.showwarning("No data", "Load an Excel file first.")
            return
        col = self.filter_col_cb.get()
        op = self.filter_op_cb.get()
        val = self.filter_val_entry.get()
        if not col or not op:
            return
        if op not in ("Is Empty", "Is Not Empty") and val == "":
            messagebox.showwarning("Missing value", "Enter a value, or choose 'Is Empty' / 'Is Not Empty'.")
            return
        self.filters.append({"column": col, "operator": op, "value": val})
        label = f"{col} {op} '{val}'" if op not in ("Is Empty", "Is Not Empty") else f"{col} {op}"
        self.filters_listbox.insert("end", label)
        self.filter_val_entry.delete(0, "end")

    def remove_filter(self):
        sel = self.filters_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.filters_listbox.delete(idx)
        del self.filters[idx]

    # --------------------------------------------------------------- Sort
    def add_sort_key(self):
        if self.df_original is None:
            messagebox.showwarning("No data", "Load an Excel file first.")
            return
        col = self.sort_col_cb.get()
        order = self.sort_order_cb.get()
        if not col:
            return
        self.sort_keys.append({"column": col, "order": order})
        self.sort_listbox.insert("end", f"{col} - {order}")

    def remove_sort_key(self):
        sel = self.sort_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.sort_listbox.delete(idx)
        del self.sort_keys[idx]

    def _click_sort(self, col):
        if self.df_processed is None or col not in self.df_processed.columns:
            return
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        try:
            numeric = pd.to_numeric(self.df_processed[col], errors="coerce")
            if numeric.notna().sum() >= 0.7 * len(numeric):
                key = numeric
            else:
                key = self.df_processed[col].astype(str).str.lower()
            order = key.sort_values(ascending=self._sort_asc, kind="mergesort").index
            self.df_processed = self.df_processed.loc[order].reset_index(drop=True)
        except Exception:
            self.df_processed = self.df_processed.sort_values(
                by=col, ascending=self._sort_asc, kind="mergesort"
            ).reset_index(drop=True)
        self._render_preview(self.df_processed)

    # ------------------------------------------------------------- Apply
    def apply_all(self):
        if self.df_original is None:
            messagebox.showwarning("No data", "Load an Excel file first.")
            return
        try:
            df = self._process_dataframe(self.df_original)
            self._push_undo()
            self.df_processed = df
            self._sort_col = None
            self._render_preview(self.df_processed)
            self.status_label.config(
                text=f"Processed: {len(self.df_processed)} rows, {len(self.df_processed.columns)} columns.",
                foreground="green",
            )
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error applying changes", str(e))

    def _process_dataframe(self, df_in):
        """Apply the current column-selection / type-conversion / filters
        / sort settings to any dataframe that has the same columns as the
        currently loaded working data. Used for the main preview and for
        bulk 'export separately'."""
        df = df_in.copy()
        keep_cols = [c for c in df.columns if self.column_vars.get(c) and self.column_vars[c].get()]
        if not keep_cols:
            keep_cols = list(df.columns)
        df = df[keep_cols]

        for col in keep_cols:
            target = self.dtype_vars.get(col)
            if target is not None:
                df[col] = self._convert_column(df[col], target.get())

        for f in self.filters:
            df = self._apply_filter(df, f)

        if self.sort_keys:
            by = [k["column"] for k in self.sort_keys if k["column"] in df.columns]
            ascending = [k["order"] == "Ascending" for k in self.sort_keys if k["column"] in df.columns]
            if by:
                df = df.sort_values(by=by, ascending=ascending)

        return df.reset_index(drop=True)

    def _convert_column(self, series, target):
        try:
            if target == "Text":
                return series.astype(str).replace("nan", "")
            elif target == "Integer":
                return pd.to_numeric(series, errors="coerce").astype("Int64")
            elif target == "Decimal":
                return pd.to_numeric(series, errors="coerce")
            elif target == "Date":
                return pd.to_datetime(series, errors="coerce")
            elif target == "Yes/No (Boolean)":
                return series.astype(str).str.strip().str.lower().map(
                    lambda v: True if v in ("true", "yes", "1", "y") else
                    (False if v in ("false", "no", "0", "n") else None)
                )
        except Exception:
            pass
        return series

    def _apply_filter(self, df, f):
        col, op, val = f["column"], f["operator"], f["value"]
        if col not in df.columns:
            return df
        series = df[col]

        if op == "Is Empty":
            return df[series.isna() | (series.astype(str).str.strip() == "")]
        if op == "Is Not Empty":
            return df[~(series.isna() | (series.astype(str).str.strip() == ""))]

        if op in ("Greater Than", "Less Than", "Greater or Equal", "Less or Equal"):
            try:
                num_series = pd.to_numeric(series, errors="coerce")
                num_val = float(val)
                if op == "Greater Than":
                    return df[num_series > num_val]
                if op == "Less Than":
                    return df[num_series < num_val]
                if op == "Greater or Equal":
                    return df[num_series >= num_val]
                if op == "Less or Equal":
                    return df[num_series <= num_val]
            except Exception:
                s = series.astype(str)
                if op == "Greater Than":
                    return df[s > val]
                if op == "Less Than":
                    return df[s < val]
                if op == "Greater or Equal":
                    return df[s >= val]
                if op == "Less or Equal":
                    return df[s <= val]

        s = series.astype(str)
        if op == "Equals":
            return df[s.str.strip().str.lower() == val.strip().lower()]
        if op == "Not Equals":
            return df[s.str.strip().str.lower() != val.strip().lower()]
        if op == "Contains":
            return df[s.str.contains(val, case=False, na=False)]
        if op == "Does Not Contain":
            return df[~s.str.contains(val, case=False, na=False)]

        return df

    # ------------------------------------------------------- Manual rows
    def delete_selected_rows(self):
        if self.df_processed is None:
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Nothing selected", "Click one or more rows in the preview first "
                                                      "(Ctrl+click or Shift+click for multiple).")
            return
        idx_to_drop = [int(i) for i in sel]
        self._push_undo()
        self.df_processed = self.df_processed.drop(index=idx_to_drop).reset_index(drop=True)
        self._render_preview(self.df_processed)
        self.status_label.config(text=f"Deleted {len(idx_to_drop)} row(s). {len(self.df_processed)} rows remain.",
                                  foreground="green")

    # ------------------------------------------------------------ Preview
    def _render_preview(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            arrow = ""
            if col == self._sort_col:
                arrow = " \u25B2" if self._sort_asc else " \u25BC"
            self.tree.heading(col, text=str(col) + arrow, command=lambda c=col: self._click_sort(c))
            self.tree.column(col, width=120, anchor="w")
        for i, row in df.iterrows():
            values = ["" if pd.isna(v) else str(v) for v in row.tolist()]
            self.tree.insert("", "end", iid=str(i), values=values)
        self.row_count_label.config(text=f"{len(df)} rows shown")

    # -------------------------------------------------- Template library
    def _load_templates(self):
        templates = {}
        for path in glob.glob(os.path.join(TEMPLATES_DIR, "*.json")):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                templates[data["name"]] = data["columns"]
            except Exception:
                continue
        return templates

    def _refresh_templates_list(self):
        self.templates = self._load_templates()
        self.templates_tree.delete(*self.templates_tree.get_children())
        for name, cols in sorted(self.templates.items()):
            self.templates_tree.insert("", "end", iid=name, values=[name, len(cols)])
        names = sorted(self.templates.keys())
        self.bulk_template_cb.configure(values=names)
        if names and not self.bulk_template_cb.get():
            self.bulk_template_cb.current(0)

    def _save_template(self, name, columns):
        safe_name = re.sub(r'[\\/*?:"<>|]', "_", name).strip() or "template"
        path = os.path.join(TEMPLATES_DIR, f"{safe_name}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"name": name, "columns": columns}, f, indent=2)

    def upload_template(self):
        path = filedialog.askopenfilename(
            title="Select an Excel file to use as an output structure",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            raw = pd.read_excel(path, header=None)
        except Exception as e:
            messagebox.showerror("Error loading file", str(e))
            return

        guess = guess_header_row(raw)
        dialog = HeaderRowDialog(self.root, raw, guess, title="Select the header row in this template")
        self.root.wait_window(dialog)
        header_idx = dialog.result if dialog.result is not None else 0
        header_vals = raw.iloc[header_idx].tolist()
        columns = [str(v).strip() if not pd.isna(v) else f"Column{i+1}" for i, v in enumerate(header_vals)]
        columns = dedupe_columns(columns)

        default_name = os.path.splitext(os.path.basename(path))[0]
        name = simpledialog.askstring(
            "Name this template",
            "Save this output structure under what name?\n(You'll pick it from a list next time.)",
            initialvalue=default_name, parent=self.root,
        )
        if not name:
            return
        if name in self._load_templates():
            if not messagebox.askyesno("Overwrite?", f"A template named '{name}' already exists. Overwrite it?"):
                return
        self._save_template(name, columns)
        self._refresh_templates_list()
        messagebox.showinfo("Saved", f"Template '{name}' saved with {len(columns)} columns.")

    def delete_template(self):
        sel = self.templates_tree.selection()
        if not sel:
            messagebox.showwarning("Nothing selected", "Select a template first.")
            return
        name = sel[0]
        if not messagebox.askyesno("Delete template?", f"Delete the template '{name}'? This cannot be undone."):
            return
        safe_name = re.sub(r'[\\/*?:"<>|]', "_", name).strip() or "template"
        path = os.path.join(TEMPLATES_DIR, f"{safe_name}.json")
        try:
            os.remove(path)
        except OSError:
            pass
        self._refresh_templates_list()

    def preview_template_columns(self):
        sel = self.templates_tree.selection()
        if not sel:
            messagebox.showwarning("Nothing selected", "Select a template first.")
            return
        name = sel[0]
        cols = self.templates.get(name, [])
        messagebox.showinfo(f"'{name}' columns", "\n".join(cols) if cols else "(no columns)")

    def autofill_template_from_loaded(self):
        if self.df_original is None:
            messagebox.showwarning("No data loaded", "Upload a source Excel file above first.")
            return
        sel = self.templates_tree.selection()
        if not sel:
            messagebox.showwarning("No template selected", "Select a saved template first.")
            return
        name = sel[0]
        template_cols = self.templates.get(name, [])
        source_cols = list(self.df_original.columns)
        guess = guess_mapping(template_cols, source_cols)

        dialog = MappingDialog(self.root, template_cols, source_cols, guess)
        self.root.wait_window(dialog)
        if dialog.result is None:
            return
        mapping = dialog.result

        new_df = apply_mapping(self.df_original, template_cols, mapping)
        unmatched = [t for t, s in mapping.items() if not s]
        self._push_undo()
        self._set_working_data(new_df, label=f"Template '{name}' applied to {os.path.basename(self.file_path or '')}")
        msg = f"Filled '{name}' ({len(template_cols)} columns) with {len(new_df)} rows."
        if unmatched:
            msg += f"\n\nLeft blank (no match chosen): {', '.join(unmatched)}"
        messagebox.showinfo("Template applied", msg)

    # -------------------------------------------------------- Bulk process
    def select_bulk_files(self):
        paths = filedialog.askopenfilenames(
            title="Select multiple Excel files",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not paths:
            return
        for p in paths:
            if p not in self.bulk_files:
                self.bulk_files.append(p)
                self.bulk_files_listbox.insert("end", os.path.basename(p))
        self._bulk_log(f"{len(paths)} file(s) added. Total: {len(self.bulk_files)}.")

    def remove_bulk_file(self):
        sel = self.bulk_files_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.bulk_files_listbox.delete(idx)
        del self.bulk_files[idx]

    def clear_bulk_files(self):
        self.bulk_files = []
        self.bulk_files_listbox.delete(0, "end")
        self.bulk_mapping = None
        self.bulk_template_cols = None

    def _bulk_selected_template(self):
        name = self.bulk_template_cb.get()
        if not name:
            messagebox.showwarning("No template selected", "Choose an Output Template first (Output Templates tab "
                                                             "to create one if you haven't yet).")
            return None, None
        templates = self._load_templates()
        if name not in templates:
            messagebox.showerror("Template not found", f"'{name}' no longer exists. Refresh the list.")
            return None, None
        return name, templates[name]

    def _read_and_header(self, path):
        raw = pd.read_excel(path, header=None)
        header_idx = guess_header_row(raw)
        df = build_df_from_raw(raw, header_idx)
        return df

    def bulk_automap(self):
        if not self.bulk_files:
            messagebox.showwarning("No files", "Select the source files first.")
            return
        name, template_cols = self._bulk_selected_template()
        if template_cols is None:
            return
        try:
            first_df = self._read_and_header(self.bulk_files[0])
        except Exception as e:
            messagebox.showerror("Error reading file", f"{os.path.basename(self.bulk_files[0])}:\n{e}")
            return

        guess = guess_mapping(template_cols, list(first_df.columns))
        dialog = MappingDialog(self.root, template_cols, list(first_df.columns), guess)
        self.root.wait_window(dialog)
        if dialog.result is None:
            self._bulk_log("Mapping cancelled.")
            return
        self.bulk_mapping = dialog.result
        self.bulk_template_cols = template_cols
        self._bulk_log(f"Mapping set using '{os.path.basename(self.bulk_files[0])}' as reference for template '{name}'. "
                        f"This same mapping will be used for every file in the batch.")

    def bulk_combine(self):
        if not self.bulk_files:
            messagebox.showwarning("No files", "Select the source files first.")
            return
        if not self.bulk_mapping or not self.bulk_template_cols:
            messagebox.showwarning("No mapping yet", "Click '1) Auto-Map Columns' first.")
            return

        combined = []
        errors = []
        for path in self.bulk_files:
            try:
                df = self._read_and_header(path)
                mapped = apply_mapping(df, self.bulk_template_cols, self.bulk_mapping)
                if self.bulk_add_source_col.get():
                    mapped.insert(0, "Source File", os.path.basename(path))
                combined.append(mapped)
                self._bulk_log(f"OK: {os.path.basename(path)} -> {len(mapped)} rows")
            except Exception as e:
                errors.append((path, str(e)))
                self._bulk_log(f"FAILED: {os.path.basename(path)} -> {e}")

        if not combined:
            messagebox.showerror("Nothing combined", "No files could be read successfully.")
            return

        result = pd.concat(combined, ignore_index=True)
        self._push_undo()
        self._set_working_data(result, label=f"Bulk combined: {len(self.bulk_files)} files -> {len(result)} rows")
        msg = f"Combined {len(combined)} of {len(self.bulk_files)} files into {len(result)} rows."
        if errors:
            msg += f"\n\n{len(errors)} file(s) failed - see the log at the bottom of the Bulk Process tab."
        messagebox.showinfo("Combined", msg)

    def bulk_export_separate(self):
        if not self.bulk_files:
            messagebox.showwarning("No files", "Select the source files first.")
            return
        if not self.bulk_mapping or not self.bulk_template_cols:
            messagebox.showwarning("No mapping yet", "Click '1) Auto-Map Columns' first.")
            return
        out_dir = filedialog.askdirectory(title="Choose an output folder")
        if not out_dir:
            return

        # Build column/type controls against the template's columns so
        # _process_dataframe's filters/sort (defined on those column
        # names) apply correctly to each mapped file.
        prev_state = (self.column_vars, self.dtype_vars)
        self._populate_column_controls(self.bulk_template_cols)

        succeeded, failed = 0, 0
        also_pdf = self.bulk_export_pdf_var.get()
        for path in self.bulk_files:
            try:
                df = self._read_and_header(path)
                mapped = apply_mapping(df, self.bulk_template_cols, self.bulk_mapping)
                processed = self._process_dataframe(mapped)
                base = os.path.splitext(os.path.basename(path))[0]
                xlsx_path = os.path.join(out_dir, f"{base}_report.xlsx")
                processed.to_excel(xlsx_path, index=False, engine="openpyxl")
                if also_pdf:
                    pdf_path = os.path.join(out_dir, f"{base}_report.pdf")
                    self._write_pdf(processed, pdf_path)
                self._bulk_log(f"Exported: {os.path.basename(path)} -> {os.path.basename(xlsx_path)}"
                                f"{' + PDF' if also_pdf else ''}")
                succeeded += 1
            except Exception as e:
                self._bulk_log(f"FAILED: {os.path.basename(path)} -> {e}")
                failed += 1

        # restore the column/type controls for the currently loaded working data
        if self.df_original is not None:
            self._populate_column_controls(self.df_original.columns.tolist())
        else:
            self.column_vars, self.dtype_vars = prev_state

        messagebox.showinfo("Bulk export finished", f"{succeeded} file(s) exported to:\n{out_dir}\n\n{failed} failed.")

    # ------------------------------------------------------------- Export
    def _get_export_df(self):
        if self.df_processed is not None:
            return self.df_processed
        if self.df_original is not None:
            return self.df_original
        return None

    def export_excel(self):
        df = self._get_export_df()
        if df is None:
            messagebox.showwarning("No data", "Load an Excel file first.")
            return
        path = filedialog.asksaveasfilename(
            title="Save Excel Report", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        try:
            df.to_excel(path, index=False, engine="openpyxl")
            messagebox.showinfo("Success", f"Excel report saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error saving Excel", str(e))

    def export_pdf(self):
        df = self._get_export_df()
        if df is None:
            messagebox.showwarning("No data", "Load an Excel file first.")
            return
        path = filedialog.asksaveasfilename(
            title="Save PDF Report", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
        )
        if not path:
            return
        try:
            self._write_pdf(df, path)
            messagebox.showinfo("Success", f"PDF report saved to:\n{path}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error saving PDF", str(e))

    def _write_pdf(self, df, path):
        styles = getSampleStyleSheet()
        cell_style = ParagraphStyle("cell", parent=styles["Normal"], fontSize=7, leading=9)
        header_style = ParagraphStyle("header", parent=styles["Normal"], fontSize=8, leading=10,
                                       textColor=colors.white, fontName="Helvetica-Bold")

        num_cols = len(df.columns)
        page_size = landscape(letter) if num_cols > 5 else letter
        doc = SimpleDocTemplate(path, pagesize=page_size,
                                 leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24)

        elements = []
        title_style = ParagraphStyle("title", parent=styles["Heading1"], fontSize=16)
        elements.append(Paragraph("Report", title_style))
        elements.append(Spacer(1, 12))

        header_row = [Paragraph(str(c), header_style) for c in df.columns]
        data = [header_row]
        for _, row in df.iterrows():
            data.append([Paragraph("" if pd.isna(v) else str(v), cell_style) for v in row.tolist()])

        available_width = page_size[0] - 48
        col_width = available_width / max(num_cols, 1)
        table = Table(data, colWidths=[col_width] * num_cols, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2C3E50")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        elements.append(table)
        doc.build(elements)


def main():
    root = tk.Tk()
    app = ReportBuilderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
