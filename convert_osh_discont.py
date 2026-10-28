# -*- coding: utf-8 -*-
"""
convert_osh_discont.py - Convert a bank Osh (current account) .xlsx export into an .xls
file built from the BankTrans.xls template.

Usage (from c:\\veka\\scr):
    python convert_osh_discont.py <input.xlsx> [output.xls]

  - <input.xlsx>  : full path, or a name looked up in c:\\veka\\scr then c:\\veka\\datafiles
  - [output.xls]  : optional. Default = same name as the input with .xls.
                    Always saved in c:\\veka\\datafiles

Requires:  pip install openpyxl xlrd xlwt xlutils
"""
import sys
import os
import re
import datetime

import openpyxl
from openpyxl.utils.datetime import from_excel
import xlrd
import xlwt
from xlutils.copy import copy as xl_copy

# ---------------------------------------------------------------- settings --
BASE_DIR = r"c:\veka\scr"
OUT_DIR = r"c:\veka\datafiles"
TEMPLATE = "BankTrans.xls"          # looked up in BASE_DIR
INPUT_TITLE_ROW = 9                 # titles in the input file are in row 9 (1-based)
INPUT_FIRST_DATA_ROW = 10           # first transaction row in the input (1-based)
OUTPUT_TITLE_ROW = 4                # titles in the template are in row 4 (1-based)
OUTPUT_FIRST_DATA_ROW = 5           # data is written from row 5; rows 1-4 untouched
DESC_MAX = 25                       # max characters copied to "תאור"
DATE_FORMAT = "dd/mm/yyyy"

# Column titles (alternative spellings allowed)
IN_COLS = {
    "date":       ["תאריך"],
    "value_date": ["יום ערך"],
    "ref":        ["אסמכתה", "אסמכתא"],
    "desc":       ["תיאור התנועה"],
    "amount":     ["זכות/חובה  ₪", "זכות/חובה"],
}
OUT_COLS = {
    "date":       ["תאריך"],
    "value_date": ["תאריך ערך"],
    "ref":        ["אסמכתא", "אסמכתה"],
    "desc":       ["תאור", "תיאור"],
    "credit":     ["סך זכות"],
    "debit":      ["סך חובה"],
}
# Output columns given by letter (used instead of searching row 4 for the title)
OUT_FIXED_COLS = {
    "debit":      "J",   # סך חובה
    "value_date": "H",   # תאריך ערך
}
# ---------------------------------------------------------------------------


def col_index(letter):
    """'A' -> 0, 'H' -> 7, 'AA' -> 26"""
    n = 0
    for ch in letter.strip().upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def norm(s):
    """Normalize a header text for comparison."""
    s = "" if s is None else str(s)
    s = s.replace('"', "").replace("'", "").replace("\u200f", "").replace("\u200e", "")
    s = s.replace("₪", "").replace("\u00a0", " ")      # ignore shekel sign / nbsp
    s = re.sub(r"\s*/\s*", "/", s)                       # "זכות / חובה" -> "זכות/חובה"
    return re.sub(r"\s+", " ", s).strip()


def find_header(rows, wanted, max_rows):
    """rows: list of lists of cell values. Returns (row_index, {key: col_index})."""
    for r_idx, row in enumerate(rows[:max_rows]):
        cells = [norm(v) for v in row]
        found = {}
        for key, names in wanted.items():
            for name in names:
                if norm(name) in cells:
                    found[key] = cells.index(norm(name))
                    break
        if len(found) == len(wanted):
            return r_idx, found
    return None, None


def match_titles(row, wanted):
    """Match titles in one row. Returns ({key: col_index}, [missing keys])."""
    cells = [norm(v) for v in row]
    found, missing = {}, []
    for key, names in wanted.items():
        for name in names:
            if norm(name) in cells:
                found[key] = cells.index(norm(name))
                break
        else:
            missing.append(key)
    return found, missing


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime.datetime):
        return v
    if isinstance(v, datetime.date):
        return datetime.datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)):
        try:
            return from_excel(v)
        except Exception:
            return None
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%d.%m.%Y", "%d.%m.%y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def to_number(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").replace("₪", "").replace("\u200f", "").replace("\u200e", "").strip()
    neg = False
    if s.endswith("-"):            # "123.45-" style
        neg, s = True, s[:-1]
    try:
        n = float(s)
    except ValueError:
        return None
    return -n if neg else n


def to_ref(v):
    if v is None or v == "":
        return ""
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    return int(s) if s.isdigit() else s


def main():
    if len(sys.argv) not in (2, 3):
        print("Usage: python convert_osh_discont.py <input.xlsx> [output.xls]")
        sys.exit(1)

    os.chdir(BASE_DIR)

    in_path = sys.argv[1]
    if not os.path.isabs(in_path):
        for d in (BASE_DIR, OUT_DIR):
            cand = os.path.join(d, in_path)
            if os.path.isfile(cand):
                in_path = cand
                break
        else:
            in_path = os.path.join(BASE_DIR, in_path)

    if len(sys.argv) == 3:
        out_name = os.path.basename(sys.argv[2])
    else:
        out_name = os.path.splitext(os.path.basename(in_path))[0]
    if not out_name.lower().endswith(".xls"):
        out_name = os.path.splitext(out_name)[0] + ".xls"
    out_path = os.path.join(OUT_DIR, out_name)

    template_path = os.path.join(BASE_DIR, TEMPLATE)
    for p in (in_path, template_path):
        if not os.path.isfile(p):
            print("File not found: " + p)
            sys.exit(2)

    # ---- read input xlsx
    wb_in = openpyxl.load_workbook(in_path, data_only=True, read_only=True)
    ws_in = wb_in.worksheets[0]
    in_rows = [list(r) for r in ws_in.iter_rows(values_only=True)]
    if len(in_rows) < INPUT_TITLE_ROW:
        print("Input file has no row %d (titles row)" % INPUT_TITLE_ROW)
        sys.exit(3)
    in_title_row = in_rows[INPUT_TITLE_ROW - 1]
    in_map, in_missing = match_titles(in_title_row, IN_COLS)
    if in_missing:
        print("Titles found in input row %d: %s"
              % (INPUT_TITLE_ROW, " | ".join(norm(v) for v in in_title_row if norm(v))))
        print("Missing titles: " + ", ".join(IN_COLS[k][0] for k in in_missing))
        sys.exit(3)

    # ---- open template (keep its formatting)
    rb = xlrd.open_workbook(template_path, formatting_info=True)
    rs = rb.sheet_by_index(0)
    if rs.nrows < OUTPUT_TITLE_ROW:
        print("Template %s has no row %d (titles row)" % (TEMPLATE, OUTPUT_TITLE_ROW))
        sys.exit(4)
    title_row = rs.row_values(OUTPUT_TITLE_ROW - 1)
    search = {k: v for k, v in OUT_COLS.items() if k not in OUT_FIXED_COLS}
    out_map, missing = match_titles(title_row, search)
    for k, letter in OUT_FIXED_COLS.items():
        out_map[k] = col_index(letter)
    if missing:
        print("Titles found in %s row %d: %s"
              % (TEMPLATE, OUTPUT_TITLE_ROW, " | ".join(norm(v) for v in title_row if norm(v))))
        print("Missing titles: " + ", ".join(OUT_COLS[k][0] for k in missing))
        sys.exit(4)

    wb_out = xl_copy(rb)
    ws_out = wb_out.get_sheet(0)
    date_style = xlwt.easyxf(num_format_str=DATE_FORMAT)
    num_style = xlwt.easyxf(num_format_str="#,##0.00")

    out_r = OUTPUT_FIRST_DATA_ROW - 1                # 0-based index of row 5
    first_out_r = out_r
    count = 0

    for row in in_rows[INPUT_FIRST_DATA_ROW - 1:]:
        row = list(row) + [None] * 20
        d = to_date(row[in_map["date"]])
        if d is None:
            if all(v is None or str(v).strip() == "" for v in row):
                break           # end of data
            continue            # skip summary / non-transaction lines

        vd = to_date(row[in_map["value_date"]])
        ref = to_ref(row[in_map["ref"]])
        desc = row[in_map["desc"]]
        desc = "" if desc is None else str(desc).strip()[:DESC_MAX]
        amt = to_number(row[in_map["amount"]])

        ws_out.write(out_r, out_map["date"], d, date_style)
        if vd is not None:
            ws_out.write(out_r, out_map["value_date"], vd, date_style)
        else:
            ws_out.write(out_r, out_map["value_date"], "")
        ws_out.write(out_r, out_map["ref"], ref)
        ws_out.write(out_r, out_map["desc"], desc)
        if amt is not None and amt > 0:
            ws_out.write(out_r, out_map["credit"], amt, num_style)
            ws_out.write(out_r, out_map["debit"], "")
        elif amt is not None and amt < 0:
            ws_out.write(out_r, out_map["debit"], -amt, num_style)
            ws_out.write(out_r, out_map["credit"], "")
        out_r += 1
        count += 1

    # clear any leftover sample rows from the template below the new data
    for r in range(out_r, rs.nrows):
        for c in out_map.values():
            ws_out.write(r, c, "")

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb_out.save(out_path)
    print("Done: %d transactions written to %s (rows %d-%d)"
          % (count, out_path, first_out_r + 1, out_r))


if __name__ == "__main__":
    main()
