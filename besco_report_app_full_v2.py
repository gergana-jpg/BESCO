#!/usr/bin/env python3
"""
BESCO New Members Report — Desktop App
Double-click to run. Pick your Excel file, set dates, generate PDF.
Requirements: pip install pandas openpyxl reportlab
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
import subprocess
from datetime import date, timedelta

# ── Try importing dependencies, guide user if missing ────────────────────────
missing = []
try:
    import pandas as pd
except ImportError:
    missing.append("pandas")
try:
    import openpyxl
except ImportError:
    missing.append("openpyxl")
try:
    import reportlab
except ImportError:
    missing.append("reportlab")

if missing:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Missing libraries",
        f"Please install the required libraries first.\n\n"
        f"Open Terminal and run:\n\n"
        f"pip install {' '.join(missing)}\n\n"
        f"Then reopen this app."
    )
    sys.exit(1)

# ── Now import the generation logic (same file must be alongside) ─────────────
import importlib.util, pathlib

SCRIPT_DIR = pathlib.Path(__file__).parent
GEN_SCRIPT = SCRIPT_DIR / "generate_besco_report_full_v2.py"

if not GEN_SCRIPT.exists():
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Missing file",
        f"generate_besco_report_full_v2.py not found.\n\n"
        f"Make sure both files are in the same folder:\n{SCRIPT_DIR}"
    )
    sys.exit(1)

spec = importlib.util.spec_from_file_location("gen", GEN_SCRIPT)
gen = importlib.util.module_from_spec(spec)
spec.loader.exec_module(gen)


# ── Colours & fonts ──────────────────────────────────────────────────────────
BG         = "#F8F8F6"
PANEL      = "#FFFFFF"
BESCO_BLK  = "#1A1A1A"
ACCENT     = "#2D2D2D"
SUBTLE     = "#888888"
BTN_BG     = "#1A1A1A"
BTN_FG     = "#FFFFFF"
BTN_HOV    = "#444444"
SUCCESS    = "#2E7D32"
ERROR      = "#C62828"
BORDER     = "#E0E0E0"

FONT_BODY  = ("Helvetica Neue", 13)
FONT_SMALL = ("Helvetica Neue", 11)
FONT_LOGO  = ("Helvetica Neue", 26, "bold")
FONT_LABEL = ("Helvetica Neue", 12)
FONT_BTN   = ("Helvetica Neue", 13, "bold")
FONT_MONO  = ("Courier", 11)


class BescoApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BESCO — New Members Report")
        self.configure(bg=BG)
        self.resizable(False, False)

        # Default dates: last 14 days
        today = date.today()
        two_weeks_ago = today - timedelta(days=14)

        self.excel_path = tk.StringVar()
        self.start_year  = tk.StringVar(value=str(two_weeks_ago.year))
        self.start_month = tk.StringVar(value=f"{two_weeks_ago.month:02d}")
        self.start_day   = tk.StringVar(value=f"{two_weeks_ago.day:02d}")
        self.end_year    = tk.StringVar(value=str(today.year))
        self.end_month   = tk.StringVar(value=f"{today.month:02d}")
        self.end_day     = tk.StringVar(value=f"{today.day:02d}")
        self.output_path = tk.StringVar()
        self.status_msg  = tk.StringVar(value="")

        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

    def _build_ui(self):
        outer = tk.Frame(self, bg=BG, padx=32, pady=28)
        outer.pack(fill="both", expand=True)

        # ── Logo ─────────────────────────────────────────────────────────────
        logo_frame = tk.Frame(outer, bg=BG)
        logo_frame.pack(anchor="w", pady=(0, 4))
        tk.Label(logo_frame, text="B", font=("Helvetica Neue", 28, "bold"),
                 bg=BG, fg=BESCO_BLK).pack(side="left")
        tk.Label(logo_frame, text="E", font=("Helvetica Neue", 16, "bold"),
                 bg=BG, fg=BESCO_BLK).pack(side="left", pady=(0, 10))
        tk.Label(logo_frame, text="SCO", font=("Helvetica Neue", 28, "bold"),
                 bg=BG, fg=BESCO_BLK).pack(side="left")

        tk.Label(outer, text="New Members Report Generator",
                 font=("Helvetica Neue", 13), bg=BG, fg=SUBTLE).pack(anchor="w", pady=(0, 20))

        # ── Divider ──────────────────────────────────────────────────────────
        tk.Frame(outer, bg=BORDER, height=1).pack(fill="x", pady=(0, 22))

        # ── Excel file picker ────────────────────────────────────────────────
        self._section(outer, "1  Member Database (Excel)")

        file_row = tk.Frame(outer, bg=BG)
        file_row.pack(fill="x", pady=(6, 16))

        self._path_display = tk.Label(
            file_row,
            textvariable=self.excel_path,
            font=FONT_MONO, bg=PANEL, fg=BESCO_BLK,
            anchor="w", width=38,
            relief="flat", bd=0,
            padx=10, pady=8,
        )
        self._path_display.pack(side="left", fill="x", expand=True)
        self._style_entry(self._path_display)

        tk.Button(
            file_row, text="Browse…", command=self._pick_excel,
            font=FONT_LABEL, bg=PANEL, fg=BESCO_BLK,
            activebackground="#EEEEEE", relief="flat", bd=0,
            padx=14, pady=8, cursor="hand2",
            highlightthickness=1, highlightbackground=BORDER
        ).pack(side="left", padx=(8, 0))

        # ── Date range ───────────────────────────────────────────────────────
        self._section(outer, "2  Reporting Period")

        dates_row = tk.Frame(outer, bg=BG)
        dates_row.pack(fill="x", pady=(6, 16))

        self._date_group(dates_row, "From",
                         self.start_day, self.start_month, self.start_year)
        tk.Label(dates_row, text="→", font=("Helvetica Neue", 18),
                 bg=BG, fg=SUBTLE, padx=12).pack(side="left")
        self._date_group(dates_row, "To",
                         self.end_day, self.end_month, self.end_year)

        # ── Output folder ────────────────────────────────────────────────────
        self._section(outer, "3  Save PDF To")

        out_row = tk.Frame(outer, bg=BG)
        out_row.pack(fill="x", pady=(6, 20))

        self._out_display = tk.Label(
            out_row,
            textvariable=self.output_path,
            font=FONT_MONO, bg=PANEL, fg=SUBTLE,
            anchor="w", width=38,
            relief="flat", bd=0,
            padx=10, pady=8,
        )
        self._out_display.pack(side="left", fill="x", expand=True)
        self._style_entry(self._out_display)

        tk.Button(
            out_row, text="Choose…", command=self._pick_output,
            font=FONT_LABEL, bg=PANEL, fg=BESCO_BLK,
            activebackground="#EEEEEE", relief="flat", bd=0,
            padx=14, pady=8, cursor="hand2",
            highlightthickness=1, highlightbackground=BORDER
        ).pack(side="left", padx=(8, 0))

        # ── Generate button ──────────────────────────────────────────────────
        tk.Frame(outer, bg=BORDER, height=1).pack(fill="x", pady=(4, 20))

        self._gen_btn = tk.Button(
            outer, text="Generate PDF Report",
            command=self._generate,
            font=FONT_BTN, bg=BTN_BG, fg=BTN_FG,
            activebackground=BTN_HOV, activeforeground=BTN_FG,
            relief="flat", bd=0, padx=0, pady=13,
            cursor="hand2"
        )
        self._gen_btn.pack(fill="x")
        self._gen_btn.bind("<Enter>", lambda e: self._gen_btn.config(bg=BTN_HOV))
        self._gen_btn.bind("<Leave>", lambda e: self._gen_btn.config(bg=BTN_BG))

        # ── Status ───────────────────────────────────────────────────────────
        self._status_lbl = tk.Label(
            outer, textvariable=self.status_msg,
            font=FONT_SMALL, bg=BG, fg=SUBTLE,
            wraplength=460, justify="center"
        )
        self._status_lbl.pack(pady=(12, 0))

    def _section(self, parent, text):
        tk.Label(parent, text=text, font=("Helvetica Neue", 11, "bold"),
                 bg=BG, fg=SUBTLE).pack(anchor="w")

    def _style_entry(self, widget):
        widget.config(
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor="#888888"
        )

    def _date_group(self, parent, label, day_var, month_var, year_var):
        grp = tk.Frame(parent, bg=BG)
        grp.pack(side="left")

        tk.Label(grp, text=label, font=FONT_SMALL, bg=BG, fg=SUBTLE).pack(anchor="w")

        row = tk.Frame(grp, bg=BG)
        row.pack()

        months = [f"{m:02d}" for m in range(1, 13)]
        days   = [f"{d:02d}" for d in range(1, 32)]
        years  = [str(y) for y in range(2024, 2031)]

        for var, vals, w in [(day_var, days, 4), (month_var, months, 4), (year_var, years, 6)]:
            sep = "/" if var != year_var else ""
            cb = ttk.Combobox(row, textvariable=var, values=vals,
                              width=w, state="readonly", font=FONT_LABEL)
            cb.pack(side="left", padx=1)
            if sep:
                tk.Label(row, text=sep, font=FONT_LABEL, bg=BG, fg=SUBTLE).pack(side="left")

    def _pick_excel(self):
        path = filedialog.askopenfilename(
            title="Select Member Database",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.excel_path.set(path)
            # Auto-set output folder to same directory
            if not self.output_path.get():
                self.output_path.set(os.path.dirname(path))

    def _pick_output(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_path.set(folder)

    def _parse_date(self, day_var, month_var, year_var, label):
        try:
            return date(int(year_var.get()), int(month_var.get()), int(day_var.get()))
        except ValueError:
            raise ValueError(f"Invalid {label} date — please check day/month/year.")

    def _generate(self):
        # Validate inputs
        if not self.excel_path.get():
            self._set_status("Please select the Excel database file first.", error=True)
            return

        try:
            start_dt = self._parse_date(self.start_day, self.start_month, self.start_year, "start")
            end_dt   = self._parse_date(self.end_day,   self.end_month,   self.end_year,   "end")
        except ValueError as e:
            self._set_status(str(e), error=True)
            return

        if start_dt > end_dt:
            self._set_status("Start date must be before end date.", error=True)
            return

        out_folder = self.output_path.get() or os.path.dirname(self.excel_path.get())
        s = start_dt.strftime("%d_%m_%y")
        e = end_dt.strftime("%d_%m_%y")
        out_file = os.path.join(out_folder, f"BESCO_New_Members_{s}-{e}.pdf")

        # Run in background thread so UI doesn't freeze
        self._gen_btn.config(state="disabled", text="Generating…")
        self._set_status("Reading database and building report…", error=False)

        def run():
            try:
                records = gen.load_members(self.excel_path.get(), start_dt, end_dt)
                if not records:
                    self.after(0, lambda: self._set_status(
                        f"No new members found between {start_dt} and {end_dt}.", error=True))
                    return
                gen.build_pdf(records, start_dt, end_dt, out_file)
                self.after(0, lambda: self._on_success(out_file, len(records)))
            except Exception as ex:
                self.after(0, lambda err=err: self._set_status(f"Error: {err}", error=True))
            finally:
                self.after(0, lambda: self._gen_btn.config(
                    state="normal", text="Generate PDF Report"))

        threading.Thread(target=run, daemon=True).start()

    def _on_success(self, path, count):
        self._set_status(
            f"✓  Done! {count} members included. PDF saved to:\n{path}",
            error=False, success=True
        )
        # Offer to open the PDF
        if messagebox.askyesno("Report ready",
                               f"{count} new members included.\n\nOpen the PDF now?"):
            self._open_file(path)

    def _open_file(self, path):
        if sys.platform == "darwin":
            subprocess.run(["open", path])
        elif sys.platform == "win32":
            os.startfile(path)
        else:
            subprocess.run(["xdg-open", path])

    def _set_status(self, msg, error=False, success=False):
        self.status_msg.set(msg)
        color = ERROR if error else (SUCCESS if success else SUBTLE)
        self._status_lbl.config(fg=color)


if __name__ == "__main__":
    # On macOS, fix Tk path if needed
    if sys.platform == "darwin":
        try:
            from ctypes import cdll
            cdll.LoadLibrary("libtk8.6.dylib")
        except Exception:
            pass

    app = BescoApp()
    app.mainloop()
