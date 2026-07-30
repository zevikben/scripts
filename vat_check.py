"""
VAT Reconciliation Check
=========================
a. Sum Excel column M ("סכום חייב מע"מ") + column P ("מע"מ")
b. Sum PDF rows where "סוג פקודה" == "חמ77", using column "סכום"
c. Compare a vs b:
   - Equal      -> print "מעמ תקין"
   - Not equal  -> for each Excel row with M != 0, look up column A in PDF's
                   "אסמכתא 1":
                   ii.2 - if a matching PDF row's type is "חמ77", compare its
                     "סכום" to (Excel M value + Excel P value).
                     - Equal   -> continue to next line
                     - Not equal -> print in a table: column A number, value
                       in column M, value in PDF file, and the word "maam"
                   ii.3 - if a matching PDF row's type is "חמפ", compare its
                     "סכום" to Excel column N.
                     - Equal   -> continue to next line
                     - Not equal -> print in a table: column A number, value
                       in column N, value in PDF file, and the word "free"

Usage:
    python vat_check.py <excel_path> <pdf_path>
"""

import sys
import re
import openpyxl
import pdfplumber

HEB = re.compile(r'[\u0590-\u05FF]')
TYPES = {'חמ77', 'חמפחדש', 'חמפ', 'חס7'}

COLS = [
    ('amount', 0, 112),
    ('name', 112, 225),
    ('date_value', 225, 275),
    ('date_ref', 275, 320),
    ('ref1', 320, 400),
    ('credit', 400, 445),
    ('debit', 445, 485),
    ('type', 485, 530),
    ('order', 530, 600),
]


def rev(s):
    """Reverse pure-Hebrew tokens (PDF stores RTL glyphs in visual order)."""
    return s[::-1] if HEB.search(s) else s


def classify(x0):
    for name, lo, hi in COLS:
        if lo <= x0 < hi:
            return name
    return None


def parse_amount(raw):
    s = raw.replace(',', '').strip()
    # Parentheses may be visually mirrored by RTL rendering -> just detect presence
    neg = '(' in s or ')' in s
    s = s.strip('()')
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def parse_pdf(pdf_path):
    """Extract PDF journal rows using coordinate-based column detection."""
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
            lines = {}
            for w in words:
                top = round(w['top'], 1)
                lines.setdefault(top, []).append(w)
            for _top, ws in lines.items():
                cols = {}
                for w in ws:
                    col = classify(w['x0'])
                    if col:
                        cols.setdefault(col, []).append(w)
                if 'order' not in cols or 'type' not in cols:
                    continue
                order_txt = ''.join(w['text'] for w in cols['order'])
                if not re.fullmatch(r'\d{3,5}', order_txt):
                    continue
                type_txt = rev(''.join(w['text'] for w in cols['type']))
                if type_txt not in TYPES:
                    continue
                amt_txt = ''.join(w['text'] for w in cols.get('amount', []))
                amount = parse_amount(amt_txt) if amt_txt else None
                ref1_txt = ''.join(w['text'] for w in cols.get('ref1', []))
                name_words = cols.get('name', [])
                has_latin = any(re.search(r'[A-Za-z]', w['text']) for w in name_words)
                name_words_sorted = sorted(name_words, key=lambda w: w['x0'], reverse=not has_latin)
                name = ' '.join(rev(w['text']) for w in name_words_sorted)
                rows.append({
                    'order': order_txt,
                    'type': type_txt,
                    'ref1': ref1_txt,
                    'name': name,
                    'amount': amount,
                })
    return rows


def main(excel_path, pdf_path):
    # --- a. Excel: sum column M + column P ---
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    sum_m = 0.0
    sum_p = 0.0
    excel_rows = []  # (A value, M value, N value, P value)
    for row in ws.iter_rows(min_row=2, values_only=False):
        a_val = row[0].value  # column A
        if a_val is None:
            continue
        m_val = row[12].value or 0  # column M
        n_val = row[13].value or 0  # column N
        p_val = row[15].value or 0  # column P
        sum_m += m_val
        sum_p += p_val
        excel_rows.append((a_val, m_val, n_val, p_val))
    a_total = round(sum_m + sum_p, 2)

    # --- b. PDF: sum "סכום" where "סוג פקודה" == "חמ77" ---
    pdf_rows = parse_pdf(pdf_path)
    b_total = round(sum(r['amount'] for r in pdf_rows if r['type'] == 'חמ77' and r['amount'] is not None), 2)

    print(f"sum of columns M+P at excel file: {a_total:,.2f}")
    print(f"sum in PDF file (order type =חמ77): {b_total:,.2f}")
    print()

    # --- c. Compare ---
    if a_total == b_total:
        print("maam is ok")
        return

    print("The amounts do not match!\n")

    # Build lookup: אסמכתא 1 -> list of PDF rows (there can be multiple)
    pdf_by_ref1 = {}
    for r in pdf_rows:
        pdf_by_ref1.setdefault(r['ref1'], []).append(r)

    mismatches = []
    for a_val, m_val, n_val, p_val in excel_rows:
        if not m_val:  # M == 0 -> skip
            continue
        a_str = str(int(a_val)) if isinstance(a_val, float) and a_val.is_integer() else str(a_val)
        matches = pdf_by_ref1.get(a_str, [])
        mp_val = m_val + p_val

        # --- c.ii.2: rows of type חמ77 -> compare against Excel M+P ("maam") ---
        matches_h77 = [pr for pr in matches if pr['type'] == 'חמ77']
        if not matches_h77:
            # No חמ77 row found in the PDF for this invoice at all
            mismatches.append((a_val, m_val, None, 'maam'))
        else:
            for pr in matches_h77:
                if round(pr['amount'], 2) == round(mp_val, 2):
                    continue  # equal -> fine, move to next
                mismatches.append((a_val, m_val, pr['amount'], 'maam'))

        # --- c.ii.3: rows of type חמפ -> compare against Excel N ("free") ---
        matches_hemp = [pr for pr in matches if pr['type'] == 'חמפ']
        for pr in matches_hemp:
            if round(pr['amount'], 2) == round(n_val, 2):
                continue  # equal -> fine, move to next
            mismatches.append((a_val, n_val, pr['amount'], 'free'))

    # --- Print mismatch table ---
    print(f"{'column A':<12}{'value (excel)':<16}{'value in PDF':<16}{'note':<8}")
    print("-" * 52)
    for a_val, excel_val, pdf_val, note in mismatches:
        pdf_str = f"{pdf_val:.2f}" if pdf_val is not None else "NOT FOUND"
        print(f"{a_val:<12}{excel_val:<16.2f}{pdf_str:<16}{note:<8}")

    if not mismatches:
        print("(No mismatched rows were found after the detailed inspection)")


if __name__ == '__main__':
    excel_path = sys.argv[1] if len(sys.argv) > 1 else '/mnt/user-data/uploads/incomes.xlsx'
    pdf_path = sys.argv[2] if len(sys.argv) > 2 else '/mnt/user-data/uploads/itai2m.pdf'
    main(excel_path, pdf_path)
