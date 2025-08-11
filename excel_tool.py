# rupjyoti_excel_tool_v1_1_full.py
"""
Rupjyoti's Excel Tool — Single-file final
Version 1.1 — 10th Aug 2025
Author: Rupjyoti Sarma — rup.sarma007@gmail.com

Single-file desktop app with:
- Modern UI (customtkinter)
- Excel tools: open, preview, save, remove duplicates by column, compare two files by column (export matches)
- Lookup tab (Exact / Partial) returning full row(s) and export option
- Data analysis: descriptive stats, correlation, quick charts
- Export to Excel
- Email via SMTP (optional — skipped if no credentials)
- WhatsApp Web sending via Selenium (optional)
- Keyboard Shortcuts tab and About tab
- Optional Joker icon (joker.ico / joker.png in same folder)
- Saves basic settings in a JSON config (plain text)
"""

import sys, os, time, json, threading, tempfile, webbrowser, shutil
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# Third-party imports
try:
    import customtkinter as ctk
    import pandas as pd
    import numpy as np
    import matplotlib.pyplot as plt
    import pyperclip
    import pyautogui
    from PIL import Image, ImageTk
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.service import Service as ChromeService
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.common.exceptions import WebDriverException, NoSuchElementException
except Exception as e:
    missing = str(e)
    print("One or more required packages are missing. Install dependencies before running.")
    print("Required packages: customtkinter, pandas, openpyxl, matplotlib, pyperclip, pyautogui, pillow, selenium, webdriver-manager")
    raise

import smtplib
from email.message import EmailMessage

# App constants
APP_NAME = "Rupjyoti's Excel Tool"
VERSION = "1.1"
DATE_STR = "10th Aug 2025"
AUTHOR = "Rupjyoti Sarma"
AUTHOR_EMAIL = "rup.sarma007@gmail.com"
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".rupjyoti_excel_config.json")
PREVIEW_ROWS = 200

# Utility functions
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception as e:
        print("Failed to save config:", e)

def ensure_xlsx(path):
    if not path.lower().endswith(".xlsx"):
        path = path + ".xlsx"
    return path

def save_multiple_sheets(path, sheets_dict):
    path = ensure_xlsx(path)
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for name, df in sheets_dict.items():
                safe = str(name)[:31]
                try:
                    df.to_excel(writer, index=False, sheet_name=safe)
                except Exception:
                    df2 = df.copy()
                    for col in df2.columns:
                        df2[col] = df2[col].astype(str)
                    df2.to_excel(writer, index=False, sheet_name=safe)
        return True, None
    except Exception as e:
        return False, str(e)

def preview_df_to_tree(tree, df, max_rows=PREVIEW_ROWS):
    tree.delete(*tree.get_children())
    cols = list(df.columns)
    tree["columns"] = cols
    tree["show"] = "headings"
    for c in cols:
        tree.heading(c, text=str(c))
        tree.column(c, width=120, anchor="w")
    for idx, row in df.head(max_rows).iterrows():
        vals = ["" if pd.isna(row.get(c)) else str(row.get(c)) for c in cols]
        tree.insert("", "end", iid=str(idx), values=vals)

def try_read_excel(path, sheet_name=0):
    return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")

def open_file_with_system(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform.startswith("darwin"):
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')
    except Exception:
        pass

# GUI setup
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class RupjyotiExcelApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} — v{VERSION}")
        self.geometry("1200x760")
        self.cfg = load_config()
        self.smtp_settings = self.cfg.get("smtp", {})
        self.last_exported = self.cfg.get("last_exported")
        self.df = None
        self.filepath = None
        self.sheet_names = []
        self.lookup_files = {}

        # Try to load icon if exists
        self._load_icon()

        # Menu and shortcuts
        self._create_menu()
        self._bind_shortcuts()

        # Notebook tabs
        self._build_notebook()

        # Load last opened if available
        last = self.cfg.get("last_opened")
        if last and os.path.exists(last):
            try:
                self.load_excel(last)
            except Exception:
                pass

    def _load_icon(self):
        for fname in ("joker.ico","joker.png","joker.jpg","joker.jpeg"):
            if os.path.exists(fname):
                try:
                    if fname.lower().endswith(".ico") and sys.platform.startswith("win"):
                        self.iconbitmap(fname)
                    else:
                        img = Image.open(fname)
                        img = img.resize((64,64), Image.ANTIALIAS)
                        self._icon_img = ImageTk.PhotoImage(img)
                        self.iconphoto(False, self._icon_img)
                except Exception:
                    pass
                break

    def _create_menu(self):
        menubar = tk.Menu(self)
        filem = tk.Menu(menubar, tearoff=0)
        filem.add_command(label="Open (Ctrl+O)", command=self.open_file)
        filem.add_command(label="Save (Ctrl+S)", command=self.save_file)
        filem.add_command(label="Save As...", command=self.save_as)
        filem.add_separator()
        filem.add_command(label="Exit (Ctrl+Q)", command=self.on_close)
        menubar.add_cascade(label="File", menu=filem)

        toolm = tk.Menu(menubar, tearoff=0)
        toolm.add_command(label="Compare Two Files (Ctrl+C)", command=self.open_compare_window)
        toolm.add_command(label="Remove Duplicates by Column (Ctrl+D)", command=self.remove_duplicates_by_column)
        menubar.add_cascade(label="Tools", menu=toolm)

        sendm = tk.Menu(menubar, tearoff=0)
        sendm.add_command(label="Send via Email (Ctrl+E)", command=self.send_email_dialog)
        sendm.add_command(label="Send via WhatsApp Web (Ctrl+W)", command=self.send_whatsapp_web_dialog)
        menubar.add_cascade(label="Send", menu=sendm)

        self.config(menu=menubar)

    def _bind_shortcuts(self):
        self.bind_all("<Control-o>", lambda e: self.open_file())
        self.bind_all("<Control-s>", lambda e: self.save_file())
        self.bind_all("<Control-q>", lambda e: self.on_close())
        self.bind_all("<Control-d>", lambda e: self.remove_duplicates_by_column())
        self.bind_all("<Control-c>", lambda e: self.open_compare_window())
        self.bind_all("<Control-e>", lambda e: self.send_email_dialog())
        self.bind_all("<Control-w>", lambda e: self.send_whatsapp_web_dialog())

    def _build_notebook(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.nb = nb

        self.tab_tools = ttk.Frame(nb)
        self.tab_lookup = ttk.Frame(nb)
        self.tab_analysis = ttk.Frame(nb)
        self.tab_send = ttk.Frame(nb)
        self.tab_keyboard = ttk.Frame(nb)
        self.tab_about = ttk.Frame(nb)

        nb.add(self.tab_tools, text="Excel Tools")
        nb.add(self.tab_lookup, text="Lookup")
        nb.add(self.tab_analysis, text="Data Analysis")
        nb.add(self.tab_send, text="Send File")
        nb.add(self.tab_keyboard, text="Keyboard")
        nb.add(self.tab_about, text="About")

        self._build_tools_tab()
        self._build_lookup_tab()
        self._build_analysis_tab()
        self._build_send_tab()
        self._build_keyboard_tab()
        self._build_about_tab()

    # Tools tab
    def _build_tools_tab(self):
        left = ttk.Frame(self.tab_tools)
        left.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        right = ttk.Frame(self.tab_tools, width=360)
        right.pack(side="right", fill="y", padx=6, pady=6)

        btn_frame = ttk.Frame(left)
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="Open Excel (Ctrl+O)", command=self.open_file).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Save (Ctrl+S)", command=self.save_file).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Save As...", command=self.save_as).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Compare (Ctrl+C)", command=self.open_compare_window).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Export current", command=self.export_current).pack(side="left", padx=4)

        ttk.Label(left, text=f"Preview (first {PREVIEW_ROWS} rows)").pack(anchor="w", pady=(6,0))
        self.tree = ttk.Treeview(left)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", lambda e: self.edit_selected_cell())

        self.status_var = tk.StringVar(value="No file loaded")
        ttk.Label(left, textvariable=self.status_var).pack(fill="x")

        # right panel
        ttk.Label(right, text="Operations").pack(anchor="w")
        ttk.Button(right, text="Edit selected row/cell", command=self.edit_selected_cell).pack(fill="x", pady=3)
        ttk.Button(right, text="Show first 10 rows", command=self.show_first_10).pack(fill="x", pady=3)
        ttk.Button(right, text="Show columns", command=self.show_columns).pack(fill="x", pady=3)
        ttk.Separator(right, orient="horizontal").pack(fill="x", pady=6)

        ttk.Label(right, text="Choose sheet:").pack(anchor="w")
        self.sheet_cb = ttk.Combobox(right, state="readonly")
        self.sheet_cb.pack(fill="x", pady=3)
        ttk.Button(right, text="Reload sheet", command=self.reload_sheet).pack(fill="x", pady=3)
        ttk.Separator(right, orient="horizontal").pack(fill="x", pady=6)

        ttk.Label(right, text="Remove duplicates by column:").pack(anchor="w")
        self.dupe_col_cb = ttk.Combobox(right, state="readonly")
        self.dupe_col_cb.pack(fill="x", pady=3)
        ttk.Button(right, text="Remove duplicates (keep first) (Ctrl+D)", command=self.remove_duplicates_by_column).pack(fill="x", pady=3)
        ttk.Button(right, text="Filter Duplicates by Column & Export", command=self.filter_duplicates_by_column_export).pack(fill="x", pady=3)

        ttk.Separator(right, orient="horizontal").pack(fill="x", pady=6)
        ttk.Button(right, text="Trim spaces (clean)", command=self.trim_spaces).pack(fill="x", pady=3)

    # Lookup tab
    def _build_lookup_tab(self):
        frame = self.tab_lookup
        top = ttk.Frame(frame)
        top.pack(fill="x", padx=8, pady=8)
        ttk.Label(top, text="Choose active file (or load new)").pack(anchor="w")
        self.lookup_file_cb = ttk.Combobox(top, state="readonly")
        self.lookup_file_cb.pack(fill="x", pady=4)
        ttk.Button(top, text="Browse & add file", command=self.lookup_browse_file).pack(anchor="w", pady=4)

        param = ttk.Frame(frame)
        param.pack(fill="x", padx=8, pady=6)
        ttk.Label(param, text="Lookup column:").grid(row=0, column=0, sticky="w")
        self.lookup_col_cb = ttk.Combobox(param, state="readonly")
        self.lookup_col_cb.grid(row=0, column=1, sticky="we", padx=6)
        ttk.Label(param, text="Lookup value:").grid(row=1, column=0, sticky="w", pady=(6,0))
        self.lookup_value_e = ttk.Entry(param)
        self.lookup_value_e.grid(row=1, column=1, sticky="we", padx=6, pady=(6,0))
        ttk.Label(param, text="Match type:").grid(row=2, column=0, sticky="w", pady=(6,0))
        self.lookup_match_cb = ttk.Combobox(param, state="readonly", values=["Exact","Partial"])
        self.lookup_match_cb.grid(row=2, column=1, sticky="we", padx=6, pady=(6,0))
        self.lookup_match_cb.set("Exact")
        ttk.Button(param, text="Find (XLOOKUP)", command=self.lookup_find).grid(row=3, column=0, columnspan=2, pady=8)
        param.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Matched rows (choose to export)").pack(anchor="w", padx=8)
        self.lookup_tree = ttk.Treeview(frame)
        self.lookup_tree.pack(fill="both", expand=True, padx=8, pady=(4,8))
        ttk.Button(frame, text="Export matched rows", command=self.lookup_export).pack(pady=6)

    def lookup_browse_file(self):
        p = filedialog.askopenfilename(title="Select Excel for lookup", filetypes=[("Excel","*.xlsx;*.xls")])
        if not p:
            return
        try:
            xls = pd.ExcelFile(p, engine="openpyxl")
            df = pd.read_excel(p, sheet_name=0, engine="openpyxl")
            self.lookup_files[p] = {"xls": xls, "df": df, "sheets": xls.sheet_names}
            vals = list(self.lookup_files.keys())
            self.lookup_file_cb["values"] = vals
            self.lookup_file_cb.set(p)
            cols = list(df.columns)
            self.lookup_col_cb["values"] = cols
            if cols:
                self.lookup_col_cb.set(cols[0])
            messagebox.showinfo("Loaded", f"Loaded {p}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not load file:\n{e}")

    def lookup_find(self):
        sel_file = self.lookup_file_cb.get()
        col = self.lookup_col_cb.get()
        val = self.lookup_value_e.get()
        if not sel_file or not col:
            messagebox.showerror("Missing", "Choose file and column.")
            return
        info = self.lookup_files.get(sel_file)
        if not info:
            messagebox.showerror("Missing", "File not loaded properly.")
            return
        df = info["df"]
        if self.lookup_match_cb.get() == "Exact":
            mask = df[col].astype(str).str.strip() == str(val).strip()
        else:
            mask = df[col].astype(str).str.contains(str(val), case=False, na=False)
        res = df.loc[mask].reset_index(drop=True)
        if res.empty:
            messagebox.showinfo("No match", "No matching rows found.")
            self.lookup_tree.delete(*self.lookup_tree.get_children())
            return
        preview_df_to_tree(self.lookup_tree, res, max_rows=len(res))
        self.last_lookup_result = res
        messagebox.showinfo("Found", f"{len(res)} matching row(s) found.")

    def lookup_export(self):
        if not hasattr(self, "last_lookup_result") or self.last_lookup_result is None or self.last_lookup_result.empty:
            messagebox.showinfo("No data", "No matched rows to export.")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if not p:
            return
        ok, err = save_multiple_sheets(p, {"LookupResult": self.last_lookup_result})
        if ok:
            messagebox.showinfo("Exported", f"Lookup results exported: {p}")
            self.last_exported = p
            self.cfg["last_exported"] = p
            save_config(self.cfg)
        else:
            messagebox.showerror("Error", err)

    # Analysis tab
    def _build_analysis_tab(self):
        f = self.tab_analysis
        ttk.Label(f, text="Data Analysis Tools").pack(anchor="w", padx=8, pady=6)
        ttk.Button(f, text="Descriptive statistics", command=self.descriptive_stats).pack(fill="x", padx=8, pady=4)
        ttk.Button(f, text="Correlation matrix", command=self.correlation_matrix).pack(fill="x", padx=8, pady=4)
        ttk.Button(f, text="Quick chart", command=self.quick_chart).pack(fill="x", padx=8, pady=4)

    # Send tab
    def _build_send_tab(self):
        f = self.tab_send
        ttk.Label(f, text="Send exported file").pack(anchor="w", padx=8, pady=6)
        ttk.Button(f, text="Send via Email (Ctrl+E)", command=self.send_email_dialog).pack(fill="x", padx=8, pady=4)
        ttk.Button(f, text="Send via WhatsApp Web (Ctrl+W)", command=self.send_whatsapp_web_dialog).pack(fill="x", padx=8, pady=4)
        ttk.Label(f, text="Last exported file:").pack(anchor="w", padx=8, pady=(12,0))
        self.last_export_var = tk.StringVar(value=self.cfg.get("last_exported", "None"))
        ttk.Label(f, textvariable=self.last_export_var).pack(anchor="w", padx=8)

    # Keyboard tab
    def _build_keyboard_tab(self):
        f = self.tab_keyboard
        ttk.Label(f, text="Keyboard Shortcuts", font=("Helvetica", 12, "bold")).pack(anchor="w", padx=8, pady=8)
        shortcuts = [
            ("Open Excel File", "Ctrl + O"),
            ("Save / Export Excel File", "Ctrl + S"),
            ("Remove Duplicates", "Ctrl + D"),
            ("Compare Columns (2 files)", "Ctrl + C"),
            ("Data Analysis tab", "Ctrl + A"),
            ("Send via Email", "Ctrl + E"),
            ("Send via WhatsApp Web", "Ctrl + W"),
            ("Exit App", "Ctrl + Q")
        ]
        tree = ttk.Treeview(f, columns=("Action","Shortcut"), show="headings", height=len(shortcuts))
        tree.heading("Action", text="Action")
        tree.heading("Shortcut", text="Shortcut")
        tree.column("Action", width=480)
        tree.column("Shortcut", width=160)
        tree.pack(padx=8, pady=6)
        for a,s in shortcuts:
            tree.insert("", "end", values=(a,s))

    # About tab
    def _build_about_tab(self):
        f = self.tab_about
        ttk.Label(f, text=APP_NAME, font=("Helvetica", 16, "bold")).pack(anchor="w", padx=8, pady=(8,4))
        ttk.Label(f, text=f"Version: {VERSION}").pack(anchor="w", padx=8)
        ttk.Label(f, text=f"Date: {DATE_STR}").pack(anchor="w", padx=8)
        ttk.Label(f, text=f"Author: {AUTHOR} — {AUTHOR_EMAIL}").pack(anchor="w", padx=8, pady=(0,8))
        ttk.Label(f, text="Credit:").pack(anchor="w", padx=8)
        about_text = "This Excel Tool was created by Rupjyoti Sarma to simplify and speed up office work."
        t = tk.Text(f, wrap="word", height=6)
        t.pack(fill="x", padx=8, pady=6)
        t.insert("1.0", about_text)
        t.config(state="disabled")

    # Core functions: open, save, edit, export, duplicates, compare, analysis, email, whatsapp
    def open_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx;*.xls")])
        if not path:
            return
        try:
            self.load_excel(path)
            self.cfg["last_opened"] = path
            save_config(self.cfg)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{e}")

    def load_excel(self, path):
        xls = pd.ExcelFile(path, engine="openpyxl")
        sheets = xls.sheet_names
        self.sheet_names = sheets
        df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
        self.df = df
        self.filepath = path
        self.status_var.set(f"Loaded: {path} | Rows: {len(df)} | Columns: {len(df.columns)}")
        preview_df_to_tree(self.tree, df)
        self.sheet_cb["values"] = sheets
        if sheets:
            self.sheet_cb.set(sheets[0])
        cols = list(df.columns)
        self.dupe_col_cb["values"] = cols
        if cols:
            self.dupe_col_cb.set(cols[0])
        if path not in self.lookup_files:
            self.lookup_files[path] = {"xls": xls, "df": df, "sheets": sheets}
            vals = list(self.lookup_files.keys())
            self.lookup_file_cb["values"] = vals
            self.lookup_file_cb.set(path)
            self.lookup_col_cb["values"] = cols
            if cols:
                self.lookup_col_cb.set(cols[0])

    def reload_sheet(self):
        if not self.filepath:
            messagebox.showinfo("No file", "Load a file first.")
            return
        selected_sheet = self.sheet_cb.get()
        if not selected_sheet:
            selected_sheet = 0
        try:
            df = pd.read_excel(self.filepath, sheet_name=selected_sheet, engine="openpyxl")
            self.df = df
            preview_df_to_tree(self.tree, df)
            cols = list(df.columns)
            self.dupe_col_cb["values"] = cols
            if cols:
                self.dupe_col_cb.set(cols[0])
            messagebox.showinfo("Loaded", f"Loaded sheet: {selected_sheet}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def save_file(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a file first.")
            return
        if not self.filepath:
            self.save_as()
            return
        try:
            bkdir = os.path.join(os.path.dirname(self.filepath), "rup_backups")
            os.makedirs(bkdir, exist_ok=True)
            shutil.copy2(self.filepath, os.path.join(bkdir, f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{os.path.basename(self.filepath)}"))
        except Exception:
            pass
        try:
            with pd.ExcelWriter(self.filepath, engine="openpyxl") as writer:
                self.df.to_excel(writer, index=False, sheet_name=str(self.sheet_cb.get() or "Sheet1")[:31])
            messagebox.showinfo("Saved", f"Saved to {self.filepath}")
            self.last_exported = self.filepath
            self.cfg["last_exported"] = self.filepath
            save_config(self.cfg)
        except Exception as e:
            messagebox.showerror("Save error", str(e))

    def save_as(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a file first.")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if not p:
            return
        try:
            with pd.ExcelWriter(p, engine="openpyxl") as writer:
                self.df.to_excel(writer, index=False, sheet_name="Sheet1")
            messagebox.showinfo("Saved", f"Saved to {p}")
            self.last_exported = p
            self.cfg["last_exported"] = p
            save_config(self.cfg)
        except Exception as e:
            messagebox.showerror("Save error", str(e))

    def export_current(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a file first.")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if not p:
            return
        ok, err = save_multiple_sheets(p, {"Sheet1": self.df})
        if ok:
            messagebox.showinfo("Exported", f"Exported to {p}")
            self.last_exported = p
            self.cfg["last_exported"] = p
            save_config(self.cfg)
        else:
            messagebox.showerror("Export error", err)

    def edit_selected_cell(self):
        sel = self.tree.focus()
        if not sel:
            messagebox.showinfo("Select row", "Double-click a row in preview to select it.")
            return
        row_idx = int(sel)
        cols = list(self.df.columns)
        dlg = tk.Toplevel(self); dlg.title("Edit cell")
        ttk.Label(dlg, text=f"Row index: {row_idx}").pack(anchor="w", padx=8, pady=(6,0))
        ttk.Label(dlg, text="Choose column:").pack(anchor="w", padx=8)
        col_cb = ttk.Combobox(dlg, values=cols, state="readonly"); col_cb.pack(fill="x", padx=8, pady=4); col_cb.set(cols[0])
        ttk.Label(dlg, text="New value:").pack(anchor="w", padx=8); val_e = ttk.Entry(dlg); val_e.pack(fill="x", padx=8, pady=4)
        def do_save():
            col = col_cb.get(); val = val_e.get()
            try:
                if pd.api.types.is_numeric_dtype(self.df[col].dtype):
                    newv = pd.to_numeric(val)
                elif pd.api.types.is_datetime64_any_dtype(self.df[col].dtype):
                    newv = pd.to_datetime(val)
                else:
                    newv = val
            except Exception:
                newv = val
            self.df.at[row_idx, col] = newv
            preview_df_to_tree(self.tree, self.df)
            dlg.destroy()
        ttk.Button(dlg, text="Save", command=do_save).pack(pady=6)

    def show_first_10(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a file first."); return
        top = tk.Toplevel(self); top.title("First 10 rows"); txt = tk.Text(top, width=120, height=20); txt.pack(fill="both", expand=True); txt.insert("1.0", str(self.df.head(10)))

    def show_columns(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a file first."); return
        messagebox.showinfo("Columns", "\n".join(list(self.df.columns)))

    def remove_duplicates_by_column(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a file first."); return
        col = self.dupe_col_cb.get()
        if not col:
            messagebox.showerror("Choose column", "Select a column from the dropdown."); return
        before = len(self.df)
        self.df = self.df.drop_duplicates(subset=[col]).reset_index(drop=True)
        preview_df_to_tree(self.tree, self.df)
        self.status_var.set(f"Removed duplicates by {col} | Rows: {len(self.df)}")
        messagebox.showinfo("Done", f"Rows before: {before} | after: {len(self.df)}")

    def filter_duplicates_by_column_export(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a file first."); return
        col = self.dupe_col_cb.get()
        if not col:
            messagebox.showerror("Choose column", "Select a column."); return
        cleaned = self.df.drop_duplicates(subset=[col]).reset_index(drop=True)
        p = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if not p:
            return
        ok, err = save_multiple_sheets(p, {"Cleaned": cleaned})
        if ok:
            messagebox.showinfo("Exported", f"Cleaned file exported to {p}"); self.last_exported = p; self.cfg["last_exported"] = p; save_config(self.cfg)
        else:
            messagebox.showerror("Export error", err)

    def trim_spaces(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a file first."); return
        def trim(x):
            if isinstance(x, str):
                return x.strip()
            return x
        self.df = self.df.applymap(trim); preview_df_to_tree(self.tree, self.df); messagebox.showinfo("Done", "Trimmed spaces in string cells.")

    def open_compare_window(self):
        dlg = tk.Toplevel(self); dlg.title("Compare Two Excel Files (by column) -> single .xlsx with 3 sheets"); dlg.geometry("700x380")
        frm = ttk.Frame(dlg, padding=10); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="File 1").grid(row=0, column=0, sticky="w")
        e1 = ttk.Entry(frm, width=60); e1.grid(row=1, column=0, padx=(0,6))
        ttk.Button(frm, text="Browse...", command=lambda: self._browse_and_populate(e1, cb1)).grid(row=1, column=1)
        ttk.Label(frm, text="Choose column (File 1)").grid(row=2, column=0, sticky="w", pady=(8,0))
        cb1 = ttk.Combobox(frm, state="readonly"); cb1.grid(row=3, column=0, sticky="we", padx=(0,6))
        ttk.Label(frm, text="File 2").grid(row=0, column=2, sticky="w")
        e2 = ttk.Entry(frm, width=60); e2.grid(row=1, column=2, padx=(6,0))
        ttk.Button(frm, text="Browse...", command=lambda: self._browse_and_populate(e2, cb2)).grid(row=1, column=3)
        ttk.Label(frm, text="Choose column (File 2)").grid(row=2, column=2, sticky="w", pady=(8,0))
        cb2 = ttk.Combobox(frm, state="readonly"); cb2.grid(row=3, column=2, sticky="we", padx=(6,0))
        def do_compare():
            p1 = e1.get().strip(); p2 = e2.get().strip(); c1 = cb1.get().strip(); c2 = cb2.get().strip()
            if not p1 or not p2 or not c1 or not c2: messagebox.showerror("Missing", "Choose both files and both columns."); return
            dlg.destroy(); self.compare_and_export(p1, c1, p2, c2)
        ttk.Button(frm, text="Compare & Export", command=do_compare).grid(row=4, column=0, columnspan=4, pady=14)
        frm.columnconfigure(0, weight=1); frm.columnconfigure(2, weight=1)

    def _browse_and_populate(self, entry_widget, combobox_widget):
        p = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx;*.xls")])
        if not p:
            return
        entry_widget.delete(0, "end"); entry_widget.insert(0, p)
        try:
            xls = pd.ExcelFile(p, engine="openpyxl"); df0 = pd.read_excel(p, sheet_name=0, engine="openpyxl")
            combobox_widget["values"] = list(df0.columns)
            if len(df0.columns)>0: combobox_widget.set(df0.columns[0])
        except Exception as e:
            messagebox.showerror("Error reading file", str(e))

    def compare_and_export(self, p1, c1, p2, c2):
        try:
            df1 = pd.read_excel(p1, engine="openpyxl"); df2 = pd.read_excel(p2, engine="openpyxl")
        except Exception as e:
            messagebox.showerror("Read error", str(e)); return
        s1 = df1[c1].astype(str).fillna("").str.strip(); s2 = df2[c2].astype(str).fillna("").str.strip()
        try:
            merged = pd.merge(df1, df2, left_on=c1, right_on=c2, how="inner", suffixes=("_file1","_file2"))
        except Exception:
            df1_tmp = df1.copy(); df2_tmp = df2.copy(); df1_tmp["_cmp_key_"] = s1; df2_tmp["_cmp_key_"] = s2
            merged = pd.merge(df1_tmp, df2_tmp, on="_cmp_key_", how="inner", suffixes=("_file1","_file2")); merged.drop(columns=["_cmp_key_"], inplace=True, errors="ignore")
        only1 = df1.loc[~s1.isin(set(s2))].reset_index(drop=True); only2 = df2.loc[~s2.isin(set(s1))].reset_index(drop=True)
        suggested = f"compare_{os.path.splitext(os.path.basename(p1))[0]}__{os.path.splitext(os.path.basename(p2))[0]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        outp = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=suggested)
        if not outp: return
        sheets = {"Matches": merged if not merged.empty else pd.DataFrame(), "Only_in_file1": only1, "Only_in_file2": only2}
        ok, err = save_multiple_sheets(outp, sheets)
        if ok:
            messagebox.showinfo("Exported", f"Comparison exported: {outp}"); self.last_exported = outp; self.cfg["last_exported"] = outp; save_config(self.cfg); self.last_export_var.set(outp)
        else:
            messagebox.showerror("Export error", err)

    def descriptive_stats(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a file first."); return
        numeric = self.df.select_dtypes(include=[np.number])
        if numeric.empty:
            messagebox.showinfo("No numeric columns", "No numeric columns found."); return
        stats = numeric.describe().transpose()
        top = tk.Toplevel(self); top.title("Descriptive Statistics"); txt = tk.Text(top, wrap="none", width=120, height=30); txt.pack(fill="both", expand=True); txt.insert("end", stats.to_string())
        def save_stats(): p = filedialog.asksaveasfilename(defaultextension=".xlsx"); 
        try:
            if p: save_multiple_sheets(p, {"Descriptive": stats.reset_index()}); messagebox.showinfo("Saved", f"Saved to {p}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        ttk.Button(top, text="Save as Excel", command=save_stats).pack(pady=6)

    def correlation_matrix(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a file first."); return
        numeric = self.df.select_dtypes(include=[np.number])
        if numeric.empty:
            messagebox.showinfo("No numeric columns", "No numeric columns found."); return
        corr = numeric.corr()
        top = tk.Toplevel(self); top.title("Correlation Matrix"); txt = tk.Text(top, wrap="none", width=120, height=20); txt.pack(fill="both", expand=True); txt.insert("end", corr.to_string())
        try:
            fig, ax = plt.subplots(figsize=(6,6)); cax = ax.matshow(corr, cmap="RdBu"); fig.colorbar(cax)
            ax.set_xticks(range(len(corr.columns))); ax.set_xticklabels(corr.columns, rotation=90); ax.set_yticks(range(len(corr.columns))); ax.set_yticklabels(corr.columns)
            tmp = os.path.join(tempfile.gettempdir(), "corr_heatmap.png"); fig.tight_layout(); fig.savefig(tmp, bbox_inches="tight"); plt.close(fig)
            img = ImageTk.PhotoImage(Image.open(tmp)); lbl = ttk.Label(top, image=img); lbl.image = img; lbl.pack(pady=6)
        except Exception as e:
            print("Heatmap generation error:", e)

    def quick_chart(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a file first."); return
        cols = list(self.df.columns)
        top = tk.Toplevel(self); top.title("Quick Chart"); ttk.Label(top, text="Chart type (line/bar/pie)").pack(anchor="w", padx=8, pady=(6,0))
        cb_type = ttk.Combobox(top, values=["line","bar","pie"], state="readonly"); cb_type.pack(fill="x", padx=8, pady=4); cb_type.set("line")
        ttk.Label(top, text="X column (optional)").pack(anchor="w", padx=8); cb_x = ttk.Combobox(top, values=cols, state="readonly"); cb_x.pack(fill="x", padx=8, pady=4)
        ttk.Label(top, text="Y column (numeric)").pack(anchor="w", padx=8); cb_y = ttk.Combobox(top, values=cols, state="readonly"); cb_y.pack(fill="x", padx=8, pady=4)
        def make_chart():
            ctype = cb_type.get(); xcol = cb_x.get(); ycol = cb_y.get()
            if not ycol: messagebox.showerror("Missing", "Choose Y column."); return
            try:
                fig, ax = plt.subplots(figsize=(8,4))
                if ctype=="line":
                    if xcol: ax.plot(self.df[xcol], self.df[ycol]); ax.set_xlabel(xcol)
                    else: ax.plot(self.df[ycol])
                    ax.set_ylabel(ycol)
                elif ctype=="bar":
                    if xcol: ax.bar(self.df[xcol].astype(str), self.df[ycol])
                    else: ax.bar(range(len(self.df[ycol])), self.df[ycol])
                elif ctype=="pie":
                    vc = self.df[ycol].value_counts(); ax.pie(vc.values, labels=vc.index.astype(str), autopct="%1.1f%%")
                fig.tight_layout(); tmp = os.path.join(tempfile.gettempdir(), "quick_chart.png"); fig.savefig(tmp, bbox_inches="tight"); plt.close(fig); open_file_with_system(tmp)
            except Exception as e:
                messagebox.showerror("Chart error", str(e))
        ttk.Button(top, text="Create chart", command=make_chart).pack(pady=8)

    def send_email_dialog(self):
        path = getattr(self, "last_exported", None)
        if not path or not os.path.exists(path):
            path = filedialog.askopenfilename(title="Select file to email", filetypes=[("Excel","*.xlsx")])
            if not path:
                return
        dlg = tk.Toplevel(self); dlg.title("Send via Email (SMTP)"); dlg.geometry("480x480"); frm = ttk.Frame(dlg); frm.pack(fill="both", expand=True, padx=8, pady=8)
        ttk.Label(frm, text=f"File: {path}").pack(anchor="w")
        ttk.Label(frm, text="To (comma separated)").pack(anchor="w", pady=(6,0)); to_e = ttk.Entry(frm); to_e.pack(fill="x", pady=4)
        ttk.Label(frm, text="Subject").pack(anchor="w"); subj_e = ttk.Entry(frm); subj_e.pack(fill="x", pady=4); subj_e.insert(0, f"{APP_NAME} - file")
        ttk.Label(frm, text="Message").pack(anchor="w"); msg_txt = tk.Text(frm, height=6); msg_txt.pack(fill="x", pady=4); msg_txt.insert("1.0","Please find the attached file.")
        ttk.Separator(frm).pack(fill="x", pady=6)
        ttk.Label(frm, text="SMTP (host:port) e.g. smtp.gmail.com:587").pack(anchor="w"); smtp_e = ttk.Entry(frm); smtp_e.pack(fill="x", pady=4); smtp_e.insert(0, self.smtp_settings.get("smtp", "smtp.gmail.com:587"))
        ttk.Label(frm, text="Sender email").pack(anchor="w"); sender_e = ttk.Entry(frm); sender_e.pack(fill="x", pady=4); sender_e.insert(0, self.smtp_settings.get("email", AUTHOR_EMAIL))
        ttk.Label(frm, text="Password / app password").pack(anchor="w"); pwd_e = ttk.Entry(frm, show="*"); pwd_e.pack(fill="x", pady=4)
        savecreds_var = tk.BooleanVar(value=self.smtp_settings.get("save", False)); ttk.Checkbutton(frm, text="Save SMTP settings (plain text)", variable=savecreds_var).pack(anchor="w", pady=(6,0))
        def do_send():
            to_list = [t.strip() for t in to_e.get().split(",") if t.strip()]
            if not to_list: messagebox.showerror("Missing", "Enter recipient email(s)."); return
            smtp_info = smtp_e.get().strip(); sender = sender_e.get().strip(); pwd = pwd_e.get().strip(); subject = subj_e.get().strip(); body = msg_txt.get("1.0","end").strip()
            if not smtp_info or not sender or not pwd: 
                if not smtp_info and not sender and not pwd:
                    messagebox.showinfo("Skipped", "No SMTP credentials provided — skipping email send.")
                    dlg.destroy()
                    return
                messagebox.showerror("Missing", "Enter SMTP info, sender and password."); return
            if savecreds_var.get(): self.smtp_settings = {"smtp": smtp_info, "email": sender, "save": True}; self.cfg["smtp"] = self.smtp_settings; save_config(self.cfg)
            dlg.destroy(); threading.Thread(target=self._send_email, args=(smtp_info, sender, pwd, to_list, subject, body, path), daemon=True).start()
        ttk.Button(frm, text="Send Email", command=do_send).pack(pady=8)

    def _send_email(self, smtp_info, sender, password, to_list, subject, body, filepath):
        try:
            host, port = smtp_info.split(":"); port = int(port)
            msg = EmailMessage(); msg["From"] = sender; msg["To"] = ", ".join(to_list); msg["Subject"] = subject; msg.set_content(body)
            with open(filepath, "rb") as f: data = f.read()
            maintype = "application"; subtype = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=os.path.basename(filepath))
            server = smtplib.SMTP(host, port, timeout=30); server.starttls(); server.login(sender, password); server.send_message(msg); server.quit()
            messagebox.showinfo("Email", "Email sent successfully.")
        except Exception as e:
            messagebox.showerror("Email error", f"Failed to send email:\n{e}")

    def send_whatsapp_web_dialog(self):
        path = getattr(self, "last_exported", None)
        if not path or not os.path.exists(path):
            path = filedialog.askopenfilename(title="Select file to send via WhatsApp Web", filetypes=[("Excel","*.xlsx;*.xls")])
            if not path:
                return
        dlg = tk.Toplevel(self); dlg.title("Send via WhatsApp Web (Selenium)"); dlg.geometry("480x220"); frm = ttk.Frame(dlg); frm.pack(fill="both", expand=True, padx=8, pady=8)
        ttk.Label(frm, text=f"File: {path}").pack(anchor="w"); ttk.Label(frm, text="Phone number (international, e.g. 9198xxxxx)").pack(anchor="w", pady=(6,0))
        num_e = ttk.Entry(frm); num_e.pack(fill="x", pady=4); ttk.Label(frm, text="Message (optional)").pack(anchor="w"); msg_e = tk.Text(frm, height=4); msg_e.pack(fill="x", pady=4)
        def do_send(): num = num_e.get().strip(); msg = msg_e.get("1.0","end").strip(); 
        if not num: messagebox.showerror("Missing", "Enter phone number."); return
        dlg.destroy(); threading.Thread(target=self._send_whatsapp_web, args=(num, path, msg), daemon=True).start()
        ttk.Button(frm, text="Open & Send", command=do_send).pack(pady=8)

    def _send_whatsapp_web(self, phone, filepath, message_text=""):
        try:
            options = webdriver.ChromeOptions()
            profile_dir = os.path.join(os.path.expanduser("~"), ".chrome_whatsapp_profile")
            options.add_argument("--user-data-dir=" + profile_dir)
            driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
            url = f"https://web.whatsapp.com/send?phone={phone}&text={message_text}"
            driver.get(url); time.sleep(8)
            try:
                attach_btn = driver.find_element(By.CSS_SELECTOR, "span[data-icon='clip']"); attach_btn.click(); time.sleep(1)
                file_input = driver.find_element(By.XPATH, "//input[@type='file']"); file_input.send_keys(os.path.abspath(filepath)); time.sleep(2)
                send_btn = driver.find_element(By.CSS_SELECTOR, "span[data-icon='send']"); send_btn.click(); time.sleep(2)
                messagebox.showinfo("WhatsApp Web", "File send attempted. Check WhatsApp Web in browser.")
            except NoSuchElementException:
                messagebox.showinfo("WhatsApp Web", "Could not find attach/send UI automatically. The chat is open in your browser; please attach and send manually.")
        except WebDriverException as e:
            messagebox.showerror("WebDriver error", f"Selenium WebDriver error:\n{e}\nMake sure Chrome is installed and compatible with the driver.")
        except Exception as e:
            messagebox.showerror("WhatsApp Web error", f"Failed to send via WhatsApp Web:\n{e}")

    def on_close(self):
        self.cfg["last_opened"] = self.filepath
        self.cfg["last_exported"] = getattr(self, "last_exported", None)
        self.cfg["theme"] = ctk.get_appearance_mode()
        save_config(self.cfg)
        self.destroy()

def main():
    app = RupjyotiExcelApp()
    app.mainloop()

if __name__ == "__main__":
    main()
