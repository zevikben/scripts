#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_conversion_gui.py
======================
Small interactive screen for running the per-customer xlsx -> xls
conversion scripts (convertabirxlsx2ram.py and friends).

The mapping of "customer name" (שם לקוח) -> "which script to run for them"
lives in customers_config.json, next to this file. To support a new
customer, just add one line to that file - no code changes needed:

    {
      "אביר":      "convertabirxlsx2ram.py",
      "לקוח אחר":  "convert_income.py"
    }

You can edit customers_config.json directly with a text editor, or use
the "לקוח חדש" (New customer) button in this screen.

Usage
    python run_conversion_gui.py
(no command-line arguments - everything is entered in the window)

Dependencies
    tkinter (ships with the standard python.org Windows installer)
"""

import json
import os
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# ---------------------------------------------------------------------------
BASE_DIR = r"c:\veka\scr"
CONFIG_NAME = "customers_config.json"


def config_path():
    return os.path.join(BASE_DIR, CONFIG_NAME)


def load_customers():
    """Returns {customer_name: script_filename}, skipping the _comment key."""
    path = config_path()
    if not os.path.exists(path):
        raise FileNotFoundError(f"Config file not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return {k: v for k, v in data.items() if not k.startswith("_")}


def save_customers(customers):
    """Writes {customer_name: script_filename} back to the config file,
    keeping the explanatory _comment entry at the top."""
    path = config_path()
    out = {
        "_comment": (
            "Map each customer name (as it will appear in the dropdown) to "
            "the .py conversion script that should run for them. Add a new "
            "line for every new customer - no code changes needed. Script "
            "names are resolved relative to BASE_DIR (c:\\veka\\scr) unless "
            "given as a full path."
        )
    }
    out.update(customers)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)


def resolve_script_path(script_name):
    return script_name if os.path.isabs(script_name) else os.path.join(BASE_DIR, script_name)


# ---------------------------------------------------------------------------
class ConversionApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("הרצת המרת קובץ לקוח")
        self.geometry("640x520")

        try:
            self.customers = load_customers()
        except Exception as exc:
            messagebox.showerror("שגיאה", f"לא ניתן לטעון את קובץ ההגדרות:\n{exc}")
            self.customers = {}

        self._build_widgets()

    # -- UI ------------------------------------------------------------
    def _build_widgets(self):
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="x", **pad)

        # Customer picker
        ttk.Label(frm, text='שם לקוח:').grid(row=0, column=0, sticky="e")
        self.customer_var = tk.StringVar()
        self.customer_combo = ttk.Combobox(
            frm, textvariable=self.customer_var,
            values=sorted(self.customers.keys()), state="readonly", width=35
        )
        self.customer_combo.grid(row=0, column=1, sticky="w")
        if self.customers:
            self.customer_combo.current(0)
        ttk.Button(frm, text="לקוח חדש...", command=self.add_customer).grid(row=0, column=2, padx=6)

        # Input file
        ttk.Label(frm, text="קובץ קלט (xlsx):").grid(row=1, column=0, sticky="e")
        self.input_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.input_var, width=38).grid(row=1, column=1, sticky="w")
        ttk.Button(frm, text="בחר...", command=self.browse_input).grid(row=1, column=2, padx=6)

        # Output file name
        ttk.Label(frm, text="שם קובץ פלט (xls):").grid(row=2, column=0, sticky="e")
        self.output_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.output_var, width=38).grid(row=2, column=1, sticky="w")

        # Batch number
        ttk.Label(frm, text='מס\' מנה:').grid(row=3, column=0, sticky="e")
        self.batch_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.batch_var, width=38).grid(row=3, column=1, sticky="w")

        ttk.Button(self, text="הרץ המרה", command=self.run_conversion).pack(pady=10)

        ttk.Label(self, text="פלט:").pack(anchor="w", padx=10)
        self.output_box = scrolledtext.ScrolledText(self, height=16, wrap="word")
        self.output_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    # -- actions ---------------------------------------------------------
    def browse_input(self):
        path = filedialog.askopenfilename(
            title="בחר קובץ קלט",
            initialdir=BASE_DIR if os.path.isdir(BASE_DIR) else None,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.input_var.set(path)

    def add_customer(self):
        win = tk.Toplevel(self)
        win.title("לקוח חדש")
        win.geometry("380x160")

        ttk.Label(win, text="שם לקוח:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        name_var = tk.StringVar()
        ttk.Entry(win, textvariable=name_var, width=30).grid(row=0, column=1, pady=10)

        ttk.Label(win, text="שם קובץ סקריפט (.py):").grid(row=1, column=0, padx=10, sticky="e")
        script_var = tk.StringVar()
        ttk.Entry(win, textvariable=script_var, width=30).grid(row=1, column=1)

        def save():
            name = name_var.get().strip()
            script = script_var.get().strip()
            if not name or not script:
                messagebox.showwarning("חסר מידע", "יש למלא שם לקוח ושם סקריפט", parent=win)
                return
            if not script.lower().endswith(".py"):
                script += ".py"
            script_full = resolve_script_path(script)
            if not os.path.exists(script_full):
                if not messagebox.askyesno(
                    "הסקריפט לא נמצא",
                    f"הקובץ {script_full} לא נמצא בדיסק.\nלהוסיף את הלקוח בכל זאת?",
                    parent=win,
                ):
                    return
            self.customers[name] = script
            save_customers(self.customers)
            self.customer_combo["values"] = sorted(self.customers.keys())
            self.customer_var.set(name)
            win.destroy()

        ttk.Button(win, text="שמור", command=save).grid(row=2, column=0, columnspan=2, pady=14)

    def run_conversion(self):
        customer = self.customer_var.get().strip()
        input_path = self.input_var.get().strip()
        output_name = self.output_var.get().strip()
        batch_no = self.batch_var.get().strip()

        if not customer:
            messagebox.showwarning("חסר מידע", "יש לבחור לקוח")
            return
        if customer not in self.customers:
            messagebox.showerror("שגיאה", f"אין סקריפט מוגדר עבור הלקוח '{customer}'")
            return
        if not input_path:
            messagebox.showwarning("חסר מידע", "יש לבחור קובץ קלט")
            return
        if not output_name:
            messagebox.showwarning("חסר מידע", "יש להזין שם קובץ פלט")
            return
        if not batch_no:
            messagebox.showwarning("חסר מידע", "יש להזין מס' מנה")
            return

        script_path = resolve_script_path(self.customers[customer])
        if not os.path.exists(script_path):
            messagebox.showerror("שגיאה", f"קובץ הסקריפט לא נמצא:\n{script_path}")
            return

        cmd = [sys.executable, script_path, input_path, output_name, batch_no]

        self.output_box.delete("1.0", tk.END)
        self.output_box.insert(tk.END, "מריץ: " + " ".join(f'"{c}"' for c in cmd) + "\n\n")
        self.update_idletasks()

        try:
            result = subprocess.run(
                cmd, cwd=BASE_DIR if os.path.isdir(BASE_DIR) else None,
                capture_output=True, text=True,
            )
        except Exception as exc:
            self.output_box.insert(tk.END, f"שגיאה בהרצה: {exc}\n")
            return

        if result.stdout:
            self.output_box.insert(tk.END, result.stdout)
        if result.stderr:
            self.output_box.insert(tk.END, "\n--- שגיאות ---\n" + result.stderr)

        if result.returncode == 0:
            self.output_box.insert(tk.END, "\nהושלם בהצלחה.\n")
            messagebox.showinfo("הושלם", "ההמרה הסתיימה בהצלחה")
        else:
            self.output_box.insert(tk.END, f"\nהסקריפט נכשל (קוד יציאה {result.returncode}).\n")
            messagebox.showerror("שגיאה", "ההמרה נכשלה - ראה פרטים בתיבת הפלט")


if __name__ == "__main__":
    ConversionApp().mainloop()
