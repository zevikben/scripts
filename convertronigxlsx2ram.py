# -*- coding: utf-8 -*-
"""
convertronigxlsx2ram.py

Converts an input .xlsx file into an output .xls file, based on the
"income_base_template.xls" template.

USAGE (run from cmd / PowerShell):
    python convertronigxlsx2ram.py <input.xlsx> <output.xls> <mispar_mana>

Example:
    python convertronigxlsx2ram.py "input_data.xlsx" "output_report.xls" "5"

Notes / assumptions (adjust the CONFIG section below if these don't match
your files):
    - The script's own base directory is C:\\veka\\scr. That's where the
      template file "income_base_template.xls" lives.
    - The input file is read from, and the output file is written to,
      C:\\veka\\datafiles. If you pass file names only (no folder) for
      input_file/output_file, the script looks for/saves them there. You
      can also pass full paths and they will be used as-is.
    - Rows 1-4 of the template are left completely untouched. Row 4 is
      assumed to hold the output column titles, and new data is written
      starting at row 5.
    - Row 1 of the input .xlsx is assumed to hold the input column titles.
"""

import sys
import os
import re
import argparse
from datetime import datetime, date

import openpyxl
import xlrd
import xlwt
from xlutils.copy import copy as xl_copy


# ----------------------------- CONFIG ---------------------------------- #

BASE_DIR = r"C:\veka\scr"          # where the script/template lives
DATA_DIR = r"C:\veka\datafiles"    # where input is read from / output is written to
TEMPLATE_NAME = "income_base_template.xls"

# Row (1-indexed, human-friendly) where headers live
INPUT_HEADER_ROW = 1        # input xlsx
TEMPLATE_HEADER_ROW = 4     # template xls -> data starts on row 5

# Column titles (input file)
COL_IN_MISMACH = "מסמך"
COL_IN_LAKOACH = "לקוח/ה"
COL_IN_TAARICH_YETZIRA = "תאריך יצירה"
COL_IN_TAARICH = "תאריך"
COL_IN_MAAM = 'סכום כולל מע"מ'

# Column titles (output/template file)
COL_OUT_ASMACHTA = "אסמכתא"
COL_OUT_TEAR = "תאור"
COL_OUT_TAARICH = "תאריך"
COL_OUT_TAARICH_EREH = "תאריך ערך"
COL_OUT_DEBIT = "קוד חשבון חובה"
COL_OUT_CREDIT = "קוד חשבון זכות"
COL_OUT_MANA = "מספר מנה"
COL_OUT_SACH_PKUDA = "סך פקודה"
COL_OUT_SUG_PKUDA = "סוג פקודה"

# Fixed values
VAL_DEBIT = 21001
VAL_CREDIT = 71001
VAL_SUG_PKUDA = "חמ"

# Pattern to pull the reference number out of the מסמך column, e.g.
#   "חשבונית מס/קבלה/12345"  ->  "12345"
# Tolerant of extra spaces or a missing trailing slash, e.g.
#   "חשבונית מס/קבלה 12345", "חשבונית מס/קבלה/ 12345"
MISMACH_PATTERN = re.compile(r"חשבונית\s*מס\s*/\s*קבלה\s*/?\s*(\d+)")

# ------------------------------------------------------------------------ #


def resolve_path(path, base_dir):
    """If path isn't absolute, resolve it relative to base_dir."""
    if os.path.isabs(path):
        return path
    return os.path.join(base_dir, path)


def find_column(headers, title):
    """Return the 0-indexed position of `title` inside `headers`, or raise."""
    for idx, val in enumerate(headers):
        if val is not None and str(val).strip() == title:
            return idx
    raise ValueError(f"Column '{title}' not found in headers: {headers}")


def to_excel_serial(value, epoch=date(1899, 12, 30)):
    """
    Convert a python datetime/date (or a string) into an Excel serial
    date number suitable for xlwt writing with a date number format.
    Returns None if the value is empty/unparseable.
    """
    if value is None or value == "":
        return None

    if isinstance(value, datetime):
        dt = value
    elif isinstance(value, date):
        dt = datetime(value.year, value.month, value.day)
    else:
        # Try a few common string formats as a fallback
        text = str(value).strip()
        dt = None
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y %H:%M:%S"):
            try:
                dt = datetime.strptime(text, fmt)
                break
            except ValueError:
                continue
        if dt is None:
            return None

    delta = dt - datetime(epoch.year, epoch.month, epoch.day)
    return delta.days + delta.seconds / 86400.0


def main():
    parser = argparse.ArgumentParser(
        description="Convert income xlsx input into xls output using a template."
    )
    parser.add_argument("input_file", help="Path (or filename in base dir) of the input .xlsx file")
    parser.add_argument("output_file", help="Path (or filename in base dir) of the output .xls file")
    parser.add_argument("mispar_mana", help='Value to write into "מספר מנה" for every row')
    args = parser.parse_args()

    input_path = resolve_path(args.input_file, DATA_DIR)
    output_path = resolve_path(args.output_file, DATA_DIR)
    template_path = os.path.join(BASE_DIR, TEMPLATE_NAME)

    if not os.path.isfile(input_path):
        sys.exit(f"ERROR: input file not found: {input_path}")
    if not os.path.isfile(template_path):
        sys.exit(f"ERROR: template file not found: {template_path}")

    # ---------------- Read input xlsx ---------------- #
    wb_in = openpyxl.load_workbook(input_path, data_only=True)
    ws_in = wb_in.active

    header_cells = ws_in[INPUT_HEADER_ROW]
    headers_in = [c.value for c in header_cells]

    col_mismach = find_column(headers_in, COL_IN_MISMACH)
    col_lakoach = find_column(headers_in, COL_IN_LAKOACH)
    col_taarich_yetzira = find_column(headers_in, COL_IN_TAARICH_YETZIRA)
    col_taarich = find_column(headers_in, COL_IN_TAARICH)
    col_maam = find_column(headers_in, COL_IN_MAAM)

    # ---------------- Open template xls (preserving formatting) -------- #
    rb = xlrd.open_workbook(template_path, formatting_info=True)
    wb_out = xl_copy(rb)
    ws_out = wb_out.get_sheet(0)

    rb_sheet = rb.sheet_by_index(0)
    header_row_idx0 = TEMPLATE_HEADER_ROW - 1  # zero-indexed
    headers_out = [rb_sheet.cell_value(header_row_idx0, c) for c in range(rb_sheet.ncols)]

    col_asmachta = find_column(headers_out, COL_OUT_ASMACHTA)
    col_tear = find_column(headers_out, COL_OUT_TEAR)
    col_out_taarich = find_column(headers_out, COL_OUT_TAARICH)
    col_taarich_ereh = find_column(headers_out, COL_OUT_TAARICH_EREH)
    col_debit = find_column(headers_out, COL_OUT_DEBIT)
    col_credit = find_column(headers_out, COL_OUT_CREDIT)
    col_mana = find_column(headers_out, COL_OUT_MANA)
    col_sach_pkuda = find_column(headers_out, COL_OUT_SACH_PKUDA)
    col_sug_pkuda = find_column(headers_out, COL_OUT_SUG_PKUDA)

    date_style = xlwt.easyxf(num_format_str="DD/MM/YYYY")

    out_row = TEMPLATE_HEADER_ROW  # zero-indexed row right after the header row
    rows_written = 0

    for row in ws_in.iter_rows(min_row=INPUT_HEADER_ROW + 1, values_only=False):
        mismach_val = row[col_mismach].value
        if mismach_val is None or str(mismach_val).strip() == "":
            continue  # skip rows with no text in מסמך

        mismach_str = str(mismach_val)
        match = MISMACH_PATTERN.search(mismach_str)
        if match:
            asmachta_val = match.group(1)
        else:
            # Fallback: exact phrase not found as expected - grab the last
            # run of digits in the cell instead, so we still capture a
            # reference number rather than leaving it blank.
            digit_groups = re.findall(r"\d+", mismach_str)
            asmachta_val = digit_groups[-1] if digit_groups else ""

        lakoach_val = row[col_lakoach].value
        taarich_yetzira_val = row[col_taarich_yetzira].value
        taarich_val = row[col_taarich].value
        maam_val = row[col_maam].value

        # a. אסמכתא
        ws_out.write(out_row, col_asmachta, asmachta_val)
        # b. תאור
        ws_out.write(out_row, col_tear, lakoach_val)
        # c. תאריך  <- תאריך יצירה
        serial = to_excel_serial(taarich_yetzira_val)
        ws_out.write(out_row, col_out_taarich, serial, date_style)
        # d. תאריך ערך <- תאריך
        serial2 = to_excel_serial(taarich_val)
        ws_out.write(out_row, col_taarich_ereh, serial2, date_style)
        # e. קוד חשבון חובה
        ws_out.write(out_row, col_debit, VAL_DEBIT)
        # f. קוד חשבון זכות
        ws_out.write(out_row, col_credit, VAL_CREDIT)
        # g. מספר מנה
        ws_out.write(out_row, col_mana, args.mispar_mana)
        # h. סך פקודה / סוג פקודה
        ws_out.write(out_row, col_sach_pkuda, maam_val)
        ws_out.write(out_row, col_sug_pkuda, VAL_SUG_PKUDA)

        out_row += 1
        rows_written += 1

    wb_out.save(output_path)
    print(f"Done. {rows_written} row(s) written to: {output_path}")


if __name__ == "__main__":
    main()
