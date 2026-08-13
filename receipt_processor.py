# -*- coding: utf-8 -*-
r"""
receipt_processor.py
=====================

Processes Hebrew PDF receipts into an Excel (xlsx) output file, using a
customer-specific template, and (optionally) cross-references an order-type
table to fill accounting codes.

BASE DIRECTORIES
-----------------
Customer data lives under C:\veka\<customer>\ (BASE_DIR below).
The template xlsx is shared and lives under C:\veka\scr\ (TEMPLATE_DIR below).

NOTE ON GOOGLE DRIVE
---------------------
This script reads PDFs from a LOCAL folder (the "receipt directory" parameter,
exactly as specified in the requirements). It does not call the Google Drive
API. Sync/download the Drive folder locally first (Drive for Desktop, or a
manual export) and place the PDFs in the customer's receipts folder. If you
want live Drive API access instead, that needs a separate OAuth setup on
your machine and can be added as an extra step.

FOLDER LAYOUT
--------------
    C:\veka\scr\
        income_base_template.xlsx      <- shared template (rows 1-4 = header/meta, untouched)
        receipt_processor.py

    C:\veka\<customer name>\
        receipts\                      <- input PDFs live here (see RECEIPTS_SUBDIR below)
        commandtype.xlsx                <- cross-reference table (סוג פקודה pass)
        company_db.json                 <- auto-built, created on first run
        receipts_output.xlsx            <- output file (written here, same folder)

USAGE
-----
Step 1 - extract data from PDFs into a new xlsx (leaves 'סוג פקודה' blank).
Only the customer name is required; the receipts folder, template,
commandtype file and output all resolve automatically inside the customer's
directory:

    python receipt_processor.py extract --customer "Acme Ltd" --order-no 1234

    (Optional: --output to pick a different output filename, --receipts-subdir
    if the receipts folder isn't named "receipts".)

    After this runs, open the output file, fill in 'סוג פקודה' for each row
    by hand / using the company_db.json notes, then run step 2.

Step 2 - once 'סוג פקודה' has been filled in manually, cross-reference
commandtype.xlsx to fill 'קוד חשבון חובה' and 'קוד חשבון זכות':

    python receipt_processor.py apply-xref --customer "Acme Ltd"

    (Optional: --output if you used a non-default output filename in step 1.)

REQUIREMENTS
------------
    pip install openpyxl pdfplumber pymupdf pytesseract

OCR (SCANNED RECEIPTS)
-----------------------
Some receipt PDFs are scans/photos with no embedded text layer at all - the
label/number is only visible as pixels, not as searchable text. For those
files this script automatically falls back to OCR (Tesseract).

OCR on thermal-receipt fonts is NOT perfectly reliable - it can misread a
digit (e.g. reading a '6' as a '0'). Because this data feeds accounting,
every row produced via OCR is flagged in a separate "ocr_review.log" file
in the customer directory so you know exactly which rows must be checked
against the original PDF before trusting them - they are NOT silently
treated the same as rows read from a real text layer.

To use OCR you additionally need, on the Windows machine running this
script (these are NOT pip packages, they are separate installs):

  1. Tesseract OCR engine for Windows:
     https://github.com/UB-Mannheim/tesseract/wiki
     During install, make sure "Hebrew" is selected under additional
     language data (or download heb.traineddata afterwards into the
     Tesseract "tessdata" folder).
  2. If the installer doesn't add Tesseract to your PATH, set
     TESSERACT_CMD below to the full path of tesseract.exe, e.g.:
         TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
  3. pip install pymupdf pytesseract   (in addition to openpyxl/pdfplumber)

If Tesseract isn't installed/found, scanned PDFs will simply produce a
"could not OCR" error for that file (extraction of normal text-layer PDFs
is unaffected either way).
"""

import argparse
import glob
import json
import os
import re
import sys
from datetime import datetime

# Force UTF-8 output so Hebrew text prints correctly regardless of the
# Windows console's active codepage (still recommend `chcp 65001` and a
# Hebrew-capable font like Consolas in cmd.exe, or use Windows Terminal).
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font, Alignment
except ImportError:
    print("openpyxl is required: pip install openpyxl")
    raise

try:
    import fitz  # PyMuPDF - used to rasterize scanned pages for OCR
except ImportError:
    fitz = None

try:
    import pytesseract
    from PIL import Image, ImageOps
except ImportError:
    pytesseract = None
    Image = None
    ImageOps = None

try:
    from google.cloud import vision as gvision
except ImportError:
    gvision = None

try:
    import xlrd  # reads legacy .xls files (commadtype.xls) - openpyxl can't
except ImportError:
    xlrd = None


# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

BASE_DIR = r"C:\veka"
TEMPLATE_DIR = r"C:\veka\scr"

# If Tesseract isn't on your Windows PATH, set the full path here, e.g.:
#   TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if TESSERACT_CMD and pytesseract is not None:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

OCR_LANG = "heb+eng"
OCR_ZOOM = 3  # render scale multiplier before OCR (higher = sharper but slower)

# Google Vision handles lower-resolution images far better than Tesseract
# does, so it doesn't need as much upscaling - and using a lower zoom here
# avoids exceeding Vision's 40MB request size limit on long thermal receipts.
GOOGLE_VISION_ZOOM = 2
GOOGLE_VISION_MAX_BYTES = 15 * 1024 * 1024  # stay well under the 40MB API cap

# Which OCR engine to use for scanned (no-text-layer) receipts:
#   "google_vision" - much more accurate on this thermal-receipt font, but
#                      costs money past the free tier and needs the setup
#                      below. Falls back to Tesseract automatically if it
#                      fails for any reason (missing credentials, network
#                      error, etc.) so a run never hard-fails because of it.
#   "tesseract"      - free, fully offline, less accurate on this font.
OCR_ENGINE = "google_vision"

# Path to the service account JSON key downloaded from Google Cloud Console.
GOOGLE_VISION_CREDENTIALS_PATH = r"C:\veka\scr\receipt-ocr-001-8d342ab9ae2e.json"

TEMPLATE_FILENAME = "income_base_template.xlsx"
COMMANDTYPE_FILENAME = "commadtype.xls"
COMPANY_DB_FILENAME = "company_db.json"
RECEIPTS_SUBDIR_DEFAULT = "receipts"
OUTPUT_FILENAME_DEFAULT = "receipts_output.xlsx"

# Header row inside the template where column titles live. Rows 1-4 are
# reserved/untouched per spec; the header titles are assumed to be on row 4.
# If your template's header row is different, change this.
HEADER_ROW = 4
FIRST_DATA_ROW = HEADER_ROW + 1

# Column titles expected in the template (must match exactly, including nikud-
# free Hebrew as typed in the spec).
COL_BATCH_NO = "מספר מנה"
COL_ASMACHTA = "אסמכתא"
COL_DATE = "תאריך"
COL_VALUE_DATE = "תאריך ערך"
COL_AMOUNT = "סך פקודה"
COL_ORDER_TYPE = "סוג פקודה"
COL_DESCRIPTION = "תאור"  # first line of the receipt (used as a human-readable
                          # identifier when company-name detection isn't
                          # reliable). Optional column - filled if present in
                          # the template, skipped otherwise.
COL_DEBIT_CODE = "קוד חשבון חובה"
COL_CREDIT_CODE = "קוד חשבון זכות"

# commandtype.xlsx expected columns
CT_CODE = "קוד"
CT_DEBIT = "חשבון חובה"
CT_CREDIT = "חשבון זכות"
CT_TYPE = "סוג פקודה"          # matched against the category name found via cross.xlsx
CT_DETAILS = "פרטים לפקודה"     # appended to תאור, separated by "-"

# Shared cross-reference file (keyword -> category), same customer-agnostic
# location as the template.
CROSS_REFERENCE_PATH = r"C:\veka\scr\cross.xlsx"

# Flexible fragment matching any quote/gershayim style used in Hebrew
# receipts (', ׳, '', ״, ", “”) so label variants match regardless of which
# quote characters a given receipt printer used.
_Q = r"[\"'׳״]{0,2}"

# Receipt-number (אסמכתא) labels, in priority order (a-m). The FIRST label
# found anywhere in the receipt wins; the number is taken from the SAME LINE
# as that label (not just nearby characters).
ASMACHTA_LABELS = [
    r"חשבונית\s*/\s*קבלה",                    # a. חשבונית / קבלה
    r"חשבונית\s*מס\s*/\s*קבלה",                # b. חשבונית מס/קבלה
    r"חשבונית\s*מס\s*/\s*קבלה",                # c. חשבונית מס / קבלה (same pattern as b, \s* covers spacing)
    r"מספר\s*עסקה",                            # d. מספר עסקה
    r"חשבונית\s*מספר",                         # e. חשבונית מספר
    rf"חש{_Q}\s*מס\s*/\s*קבלה",                # f. חש' מס / קבלה
    r"חשבונית\s*מס\s*קבלה",                    # g. חשבונית מס קבלה
    r"(?<![\u0590-\u05FF])עסקה",               # h. עסקה (word boundary - don't match inside "העסקה")
    rf"סה{_Q}כ\s*כולל\s*מע{_Q}מ",              # i. סה'כ כולל מע''מ
    rf"כולל\s*מע{_Q}מ",                        # j. כולל מע''מ
    r"חן\s*מס\s*/\s*קבלה",                     # k. חן מס/קבלה
    r"ת\.?\s*משלוח",                           # l. ת. משלוח
]
AMOUNT_LABELS = [
    r"ל\s*[תח]\s*שלום",                        # a. לתשלום (tolerate OCR inserting a space anywhere, or misreading ת as ח)
    rf"סה{_Q}כ\s*כולל",                        # b. סה"כ כולל
    r"מחיר\s*כולל",                            # c. מחיר כולל
    r"סכום\s*התשלום",                          # d. סכום התשלום
    rf"סה{_Q}כ",                               # e. סה"כ (broadest - kept last)
]
DATE_LABELS = [
    r"תאריך",
]

DATE_PATTERN = re.compile(r"\b(\d{1,2})[./](\d{1,2})[./](\d{2,4})\b")
NUMBER_PATTERN = re.compile(r"[\d][\d,]*\.\d{2}|\d[\d,]*")
DECIMAL_AMOUNT_PATTERN = re.compile(r"\d[\d,]*\.\d{2}")  # strict money format (##.##)
ASMACHTA_NUMBER_PATTERN = re.compile(r"\d+(?:/\d+)+|\d{3,}")  # keep slash-joined groups together, e.g. "205/27/58940"
# A line made up entirely of digits/whitespace/punctuation (no letters at
# all, Hebrew or Latin) - matches handwritten annotation numbers scribbled
# above the printed receipt header, e.g. "3174." or "3/75".
NON_TEXT_LINE_PATTERN = re.compile(r"^[\d\s./\\\-:,]+$")
# Standalone stamp/marker words that commonly appear near the top of an
# Israeli receipt (e.g. "מקור" = "original", "קוד" = "code" - a stray table
# header) but are never the company name - skip these too when looking for
# the first real header line.
NOISE_HEADER_WORDS = {"מקור", "העתק", "קוד", "תעודה"}
# Lines matching these should NEVER be mistaken for a receipt number by the
# positional OCR fallback below - they commonly appear as "label: number"
# near the top of a receipt (structurally identical to a real receipt-number
# line) but are always something else: a company tax ID, business license
# number, or phone number.
FALLBACK_BLACKLIST_LABELS = [
    r"ח\.?\s*פ",           # ח.פ - company tax/registration ID
    r"עוסק\s*מורשה",       # business license number
    r"טלפון",
    r"טל\.",
    r"נייד",               # mobile phone
]


# --------------------------------------------------------------------------
# Company "known structures" database
# --------------------------------------------------------------------------

class CompanyDB:
    """
    Stores what we've learned about each company's receipt layout so we
    don't have to re-derive it from scratch every time we see that company
    again. Each entry can optionally override the generic label lists above
    if the generic patterns fail for that company - just edit the JSON file
    by hand and add "asmachta_labels" / "amount_labels" / "date_labels".
    """

    def __init__(self, path):
        self.path = path
        self.data = {}
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                self.data = json.load(f)

    def save(self):
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def identify(self, header_text):
        """Try to match header_text against a known company by substring."""
        header_text_norm = _norm(header_text)
        for name, entry in self.data.items():
            for kw in entry.get("keywords", [name]):
                if _norm(kw) in header_text_norm:
                    return name
        return None

    def register_new(self, header_text, source_file):
        """
        Create a new entry for an unrecognized company. Uses the first
        non-empty header line as both the display name and the primary
        keyword. Returns the assigned name.
        """
        lines = [l.strip() for l in header_text.splitlines() if l.strip()]
        name = lines[0] if lines else f"UNKNOWN_{source_file}"
        # avoid collisions
        base_name = name
        i = 2
        while name in self.data:
            name = f"{base_name} ({i})"
            i += 1
        self.data[name] = {
            "keywords": [base_name],
            "first_seen_file": source_file,
            "asmachta_labels": None,   # None => use generic defaults
            "amount_labels": None,
            "date_labels": None,
        }
        return name


def _norm(s):
    return re.sub(r"\s+", " ", s or "").strip()


# --------------------------------------------------------------------------
# PDF text extraction
# --------------------------------------------------------------------------

def extract_text(pdf_path):
    """
    Returns (text, used_ocr). Tries the PDF's embedded text layer first;
    if there is none (scanned/photographed receipt), falls back to OCR.
    """
    if pdfplumber is None:
        raise RuntimeError("pdfplumber is required: pip install pdfplumber")

    text_parts = []
    has_any_chars = False
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            if len(page.chars) > 0:
                has_any_chars = True
            t = page.extract_text() or ""
            text_parts.append(t)
    text = "\n".join(text_parts)

    if has_any_chars and text.strip():
        return text, False

    # No embedded text layer at all -> this is a scanned/photographed
    # receipt. Fall back to OCR.
    ocr_text = ocr_extract_text(pdf_path)
    return ocr_text, True


def ocr_extract_text(pdf_path):
    """
    Dispatches to the configured OCR engine (OCR_ENGINE). If Google Vision
    is selected but fails for any reason (missing/invalid credentials,
    network error, package not installed, etc.), automatically falls back
    to Tesseract so a run never hard-fails just because of the paid engine.
    """
    if OCR_ENGINE == "google_vision":
        try:
            return google_vision_ocr_text(pdf_path)
        except Exception as e:
            print(f"Note: Google Vision OCR failed ({e}); falling back to Tesseract for this file.")
            return tesseract_ocr_text(pdf_path)
    return tesseract_ocr_text(pdf_path)


def google_vision_ocr_text(pdf_path):
    if gvision is None:
        raise RuntimeError(
            "google-cloud-vision is not installed: pip install google-cloud-vision"
        )
    if not GOOGLE_VISION_CREDENTIALS_PATH or not os.path.isfile(GOOGLE_VISION_CREDENTIALS_PATH):
        raise RuntimeError(
            f"Google Vision credentials file not found: {GOOGLE_VISION_CREDENTIALS_PATH}"
        )
    if fitz is None:
        raise RuntimeError(
            "PyMuPDF is required to rasterize scanned pages: pip install pymupdf"
        )

    client = gvision.ImageAnnotatorClient.from_service_account_file(GOOGLE_VISION_CREDENTIALS_PATH)

    text_parts = []
    doc = fitz.open(pdf_path)
    try:
        matrix = fitz.Matrix(GOOGLE_VISION_ZOOM, GOOGLE_VISION_ZOOM)
        for page in doc:
            pix = page.get_pixmap(matrix=matrix)
            image_bytes = pix.tobytes("jpg", jpg_quality=90)
            # If still too large (very long receipt), progressively shrink
            # the zoom until it fits under Vision's request-size limit.
            zoom = GOOGLE_VISION_ZOOM
            while len(image_bytes) > GOOGLE_VISION_MAX_BYTES and zoom > 0.5:
                zoom = zoom * 0.7
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                image_bytes = pix.tobytes("jpg", jpg_quality=85)
            image = gvision.Image(content=image_bytes)
            image_context = gvision.ImageContext(language_hints=["he", "en"])
            response = client.document_text_detection(image=image, image_context=image_context)
            if response.error.message:
                raise RuntimeError(f"Google Vision API error: {response.error.message}")
            text_parts.append(response.full_text_annotation.text or "")
    finally:
        doc.close()
    return "\n".join(text_parts)


def tesseract_ocr_text(pdf_path):
    if fitz is None:
        raise RuntimeError(
            "This PDF has no text layer (it's a scan) and PyMuPDF is not "
            "installed for OCR rendering: pip install pymupdf"
        )
    if pytesseract is None:
        raise RuntimeError(
            "This PDF has no text layer (it's a scan) and pytesseract is not "
            "installed: pip install pytesseract pillow"
        )

    text_parts = []
    doc = fitz.open(pdf_path)
    try:
        matrix = fitz.Matrix(OCR_ZOOM, OCR_ZOOM)
        for page in doc:
            pix = page.get_pixmap(matrix=matrix)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            img = img.convert("L")
            img = ImageOps.autocontrast(img)
            try:
                page_text = pytesseract.image_to_string(img, lang=OCR_LANG, config="--psm 6")
            except pytesseract.TesseractNotFoundError:
                raise RuntimeError(
                    "Tesseract OCR engine not found. Install it from "
                    "https://github.com/UB-Mannheim/tesseract/wiki (with the "
                    "Hebrew language data) and, if needed, set TESSERACT_CMD "
                    "near the top of this script to its full path."
                )
            text_parts.append(page_text)
    finally:
        doc.close()
    return "\n".join(text_parts)


def get_header_text(full_text, n_lines=6):
    lines = [l for l in full_text.splitlines() if l.strip()]
    return "\n".join(lines[:n_lines])


# --------------------------------------------------------------------------
# Field extraction helpers
# --------------------------------------------------------------------------

def _find_after_label(text, label_patterns, value_pattern, window=40):
    """
    Search for each label pattern in text; if found, look at the following
    `window` characters (and, as a fallback, the same line) for a match of
    value_pattern. Returns the first match found, or None.
    """
    for label in label_patterns:
        for m in re.finditer(label, text):
            start = m.end()
            snippet = text[start:start + window]
            vm = value_pattern.search(snippet)
            if vm:
                return vm.group(0)
            # fallback: rest of the same line
            line_end = text.find("\n", start)
            line_end = line_end if line_end != -1 else len(text)
            line_snippet = text[start:line_end]
            vm = value_pattern.search(line_snippet)
            if vm:
                return vm.group(0)
    return None


def _strip_brackets(line):
    """
    Remove parenthesized content (e.g. "(8 items)") so numbers inside
    brackets - like an item count - aren't mistaken for the real value.
    If the line has an OPEN "(" with no matching close on this same line
    (OCR sometimes drops the closing paren or splits it across lines),
    everything from that "(" onward is dropped too, since whatever is in
    there can't be trusted.
    """
    stripped = re.sub(r"\([^)]*\)", "", line)
    if stripped != line:
        return stripped
    idx = line.find("(")
    if idx != -1:
        return line[:idx]
    return line


def extract_asmachta(text, labels):
    """
    Search `labels` in priority order. For the FIRST label pattern found
    anywhere in the receipt, take the number from the SAME LINE as that
    label (ignoring any number inside parentheses). Returns None if none
    of the labels are found.
    """
    lines = text.splitlines()
    for label in labels:
        compiled = re.compile(label)
        for line in lines:
            if compiled.search(line):
                m = ASMACHTA_NUMBER_PATTERN.search(_strip_brackets(line))
                if m:
                    return m.group(0)
                # label matched but no number on that line - keep looking
                # at other occurrences of the same label before moving on
        # (if we reach here, this label either wasn't found or had no
        # number on any of its lines - fall through to the next label)
    return None


def extract_asmachta_ocr_fallback(text, date_labels):
    """
    Last-resort fallback used ONLY for OCR'd (scanned) receipts where none of
    the 12 label words survived OCR intact - which happens because OCR can
    read digits far more reliably than it reads Hebrew label text on a
    degraded thermal-receipt scan. Takes the longest digit run (6+ digits)
    found in the "header zone" - the lines before the date line/item table -
    on the theory that the receipt/transaction number is normally the
    longest number printed in that zone. This is a heuristic, not a labeled
    match, so callers must flag any result from this function for manual
    review.

    To avoid mistaking a product barcode for the receipt number:
      - the header zone stops at the first date line OR the first line that
        looks like an item/barcode row (starts directly with a long digit
        run, e.g. a product code) - whichever comes first, since that marks
        the start of the itemized table.
      - within that zone, only digit runs that do NOT sit at the very start
        of their line are considered - a real receipt-number line normally
        has Hebrew text (and often a colon) before the digits, whereas a
        barcode/item-code line starts directly with the digits.
    """
    lines = text.splitlines()
    header_end = len(lines)
    for i, line in enumerate(lines):
        stripped = line.strip()
        if DATE_PATTERN.search(line) or re.match(r"^\d{6,}", stripped):
            header_end = i
            break
    header_lines = lines[:header_end] if header_end > 0 else lines[:10]

    blacklist_compiled = [re.compile(p) for p in FALLBACK_BLACKLIST_LABELS]

    candidates = []
    for line in header_lines:
        stripped = line.strip()
        if any(p.search(stripped) for p in blacklist_compiled):
            continue  # this line is a tax ID / phone / license number, never a receipt number
        for m in re.finditer(r"\d{6,}", stripped):
            if m.start() == 0:
                continue  # digit run at line start looks like a barcode/item code, not a labeled number
            candidates.append(m.group(0))
    if not candidates:
        return None
    return max(candidates, key=len)


def extract_date(text, labels):
    # Prefer a date found near a "תאריך" label; else first date anywhere.
    near = _find_after_label(text, labels, DATE_PATTERN, window=25)
    if near:
        return near
    m = DATE_PATTERN.search(text)
    return m.group(0) if m else None


def normalize_date(raw):
    """Normalize dd/mm/yy(yy) into dd/mm/yyyy where possible."""
    if not raw:
        return raw
    m = DATE_PATTERN.match(raw)
    if not m:
        return raw
    d, mo, y = m.groups()
    if len(y) == 2:
        y = "20" + y
    try:
        dt = datetime(int(y), int(mo), int(d))
        return dt.strftime("%d/%m/%Y")
    except ValueError:
        return raw


AMOUNT_LOOKAHEAD_LINES = 10


def extract_amount(text, labels):
    """
    Search `labels` in priority order (a-e). For the FIRST label pattern
    found anywhere in the receipt, take the amount from the SAME LINE as
    that label (ignoring any number inside parentheses, e.g. an item
    count like "(8 items)").

    If the label's line has no usable number (e.g. an item-count bracket
    that never closes on that line - OCR sometimes splits a receipt's
    "label column" and "value column" into separate line groups), look at
    the next few lines for the real total: skip any line that starts with
    "-" (a discount/negative line) and require a strict ##.## money format
    so an unrelated integer (a card/reference serial number, an item
    count) doesn't get mistaken for the amount.
    """
    lines = text.splitlines()
    for label in labels:
        compiled = re.compile(label)
        for i, line in enumerate(lines):
            if not compiled.search(line):
                continue
            m = NUMBER_PATTERN.search(_strip_brackets(line))
            if m:
                return m.group(0).replace(",", "")
            # nothing usable on the label's own line - look ahead
            for j in range(i + 1, min(i + 1 + AMOUNT_LOOKAHEAD_LINES, len(lines))):
                nxt = lines[j].strip()
                if not nxt or nxt.startswith("-"):
                    continue
                m2 = DECIMAL_AMOUNT_PATTERN.search(_strip_brackets(nxt))
                if m2:
                    return m2.group(0).replace(",", "")
    return None


# --------------------------------------------------------------------------
# Per-file processing
# --------------------------------------------------------------------------

def process_receipt(pdf_path, company_db, debug_dir=None):
    """
    Returns (row_dict, hard_error_msg_or_None).

    row_dict is returned even when some fields are missing - each of
    asmachta / date / amount will be None if it couldn't be found, and
    row_dict["missing_fields"] lists their Hebrew names, so the caller can
    still write a partial row (with blanks) instead of dropping the file
    entirely. hard_error_msg is only set when the file couldn't be read at
    all (corrupt PDF, missing OCR engine, etc.) - in that case row_dict is
    None and there is nothing to write.

    If debug_dir is given, the raw extracted text (exactly what the
    label-matching regex sees) is written to debug_dir/<filename>.txt -
    useful for diagnosing why a field was or wasn't found.
    """
    filename = os.path.basename(pdf_path)
    try:
        text, used_ocr = extract_text(pdf_path)
    except Exception as e:
        return None, f'error – could not read file "{filename}": {e}'

    if debug_dir:
        try:
            debug_path = os.path.join(debug_dir, filename + ".txt")
            with open(debug_path, "w", encoding="utf-8") as f:
                f.write(f"used_ocr: {used_ocr}\n\n")
                f.write(text)
        except Exception:
            pass  # debug dump is best-effort, never let it break processing

    header_text = get_header_text(text)
    header_lines = [l.strip() for l in header_text.splitlines() if l.strip()]
    # Some receipts have a handwritten annotation number scribbled above the
    # printed header (e.g. "3174."), which OCR reads as the first line -
    # skip any leading line that's purely digits/punctuation (no letters at
    # all) and use the first line that actually contains text.
    first_line = next(
        (l for l in header_lines
         if not NON_TEXT_LINE_PATTERN.fullmatch(l) and l not in NOISE_HEADER_WORDS),
        None,
    )
    if first_line is None:
        first_line = header_lines[0] if header_lines else ""

    company = company_db.identify(header_text)
    if company is None:
        company = company_db.register_new(header_text, filename)

    entry = company_db.data.get(company, {})
    asmachta_labels = entry.get("asmachta_labels") or ASMACHTA_LABELS
    amount_labels = entry.get("amount_labels") or AMOUNT_LABELS
    date_labels = entry.get("date_labels") or DATE_LABELS

    missing_fields = []

    asmachta = extract_asmachta(text, asmachta_labels)
    asmachta_is_fallback = False
    if not asmachta and used_ocr:
        # Labeled extraction failed on a scanned receipt - try the
        # positional OCR fallback before giving up.
        asmachta = extract_asmachta_ocr_fallback(text, date_labels)
        asmachta_is_fallback = asmachta is not None
    if not asmachta:
        missing_fields.append("אסמכתא")

    date_val = None
    raw_date = extract_date(text, date_labels)
    if not raw_date:
        missing_fields.append("תאריך")
    else:
        date_val = normalize_date(raw_date)

    amount = extract_amount(text, amount_labels)
    if not amount:
        missing_fields.append("סך פקודה")

    return {
        "company": company,
        "asmachta": asmachta,
        "date": date_val,
        "amount": amount,
        "used_ocr": used_ocr,
        "asmachta_is_fallback": asmachta_is_fallback,
        "filename": filename,
        "description": first_line,
        "missing_fields": missing_fields,
    }, None



# --------------------------------------------------------------------------
# Excel handling
# --------------------------------------------------------------------------

# Excel's standard "Bad" cell style colors - light red fill, dark red text.
ERROR_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
ERROR_FONT = Font(color="9C0006")


def highlight_row(ws, row_idx, col_map):
    """Marks every mapped column in this row with the red 'error' style, so
    problem rows are immediately visible when the file is opened in Excel."""
    for col_idx in col_map.values():
        cell = ws.cell(row=row_idx, column=col_idx)
        cell.fill = ERROR_FILL
        cell.font = ERROR_FONT


LEFT_ALIGN = Alignment(horizontal="left")


RIGHT_ALIGN = Alignment(horizontal="right")
# Columns that should stay right-aligned instead of the default forced-left
# alignment applied to everything else.
RIGHT_ALIGNED_COLUMNS = {"תאור", "סוג פקודה"}


def set_cell(ws, row, col, value, column_title=None):
    """Writes a value and explicitly forces alignment, so numeric-looking
    values (codes, amounts) don't end up inconsistently aligned by Excel's
    default number formatting. Left-aligned by default; pass column_title
    for a column in RIGHT_ALIGNED_COLUMNS to keep it right-aligned instead."""
    cell = ws.cell(row=row, column=col, value=value)
    cell.alignment = RIGHT_ALIGN if column_title in RIGHT_ALIGNED_COLUMNS else LEFT_ALIGN
    return cell


def find_columns(ws, header_row=HEADER_ROW):
    """Map expected Hebrew column titles to their column index in header_row."""
    col_map = {}
    for cell in ws[header_row]:
        if cell.value is not None:
            col_map[_norm(str(cell.value))] = cell.column

    required = [
        COL_BATCH_NO, COL_ASMACHTA, COL_DATE, COL_VALUE_DATE,
        COL_AMOUNT, COL_ORDER_TYPE,
    ]
    missing = [c for c in required if c not in col_map]
    if missing:
        raise RuntimeError(
            f"Template is missing expected column header(s) on row {header_row}: "
            f"{', '.join(missing)}. Check HEADER_ROW / column titles at the top "
            f"of receipt_processor.py."
        )
    return col_map


def cmd_extract(args):
    customer_dir = os.path.join(BASE_DIR, args.customer)
    if not os.path.isdir(customer_dir):
        sys.exit(f"Customer directory not found: {customer_dir}")

    template_path = os.path.join(TEMPLATE_DIR, TEMPLATE_FILENAME)
    db_path = os.path.join(customer_dir, COMPANY_DB_FILENAME)
    receipts_dir = os.path.join(customer_dir, args.receipts_subdir)
    output_path = os.path.join(customer_dir, args.output)

    if not os.path.isfile(template_path):
        sys.exit(f"Template not found: {template_path}")
    if not os.path.isdir(receipts_dir):
        sys.exit(f"Receipts directory not found: {receipts_dir}")

    wb = load_workbook(template_path)
    ws = wb.active
    col_map = find_columns(ws)
    if COL_DESCRIPTION not in col_map:
        print(f'Note: template has no "{COL_DESCRIPTION}" column - skipping that field.')

    company_db = CompanyDB(db_path)

    debug_dir = None
    if getattr(args, "debug_ocr", False):
        debug_dir = os.path.join(customer_dir, "ocr_debug")
        os.makedirs(debug_dir, exist_ok=True)
        print(f"Debug mode: raw OCR/text output will be saved to {debug_dir}")

    pdf_files = sorted(glob.glob(os.path.join(receipts_dir, "*.pdf")))
    if not pdf_files:
        sys.exit(f"No PDF files found in {receipts_dir}")

    row_idx = FIRST_DATA_ROW
    # advance past any existing data rows so we always append
    while ws.cell(row=row_idx, column=col_map[COL_BATCH_NO]).value not in (None, ""):
        row_idx += 1

    errors = []
    ocr_rows = []
    partial_rows = []
    processed = 0

    for pdf_path in pdf_files:
        row, hard_err = process_receipt(pdf_path, company_db, debug_dir=debug_dir)
        if hard_err:
            # file couldn't be read at all - nothing to write
            errors.append(hard_err)
            print(hard_err)
            continue

        set_cell(ws, row_idx, col_map[COL_BATCH_NO], args.order_no)
        set_cell(ws, row_idx, col_map[COL_ASMACHTA], row["asmachta"])
        set_cell(ws, row_idx, col_map[COL_DATE], row["date"])
        set_cell(ws, row_idx, col_map[COL_VALUE_DATE], row["date"])
        set_cell(ws, row_idx, col_map[COL_AMOUNT], row["amount"])
        if COL_DESCRIPTION in col_map:
            set_cell(ws, row_idx, col_map[COL_DESCRIPTION], row["description"], column_title=COL_DESCRIPTION)
        # COL_ORDER_TYPE intentionally left blank - filled offline (step 7)

        if row["missing_fields"]:
            fields_str = ", ".join(row["missing_fields"])
            msg = (
                f'row {row_idx}: file "{row["filename"]}" - missing field(s): '
                f'{fields_str} (other fields were written, please fill these in manually)'
            )
            partial_rows.append(msg)
            print(f"⚠ {msg}")
            highlight_row(ws, row_idx, col_map)

        if row["used_ocr"]:
            fallback_note = (
                " [[FALLBACK: no label matched - אסמכתא taken positionally, verify carefully]]"
                if row.get("asmachta_is_fallback") else ""
            )
            ocr_rows.append(
                f'row {row_idx}: file "{row["filename"]}" -> אסמכתא={row["asmachta"]}, '
                f'תאריך={row["date"]}, סך פקודה={row["amount"]}{fallback_note}'
            )

        row_idx += 1
        processed += 1

    wb.save(output_path)
    company_db.save()

    print(f"\nDone. {processed} receipt(s) written, {len(errors)} unreadable file(s), "
          f"{len(partial_rows)} row(s) with missing field(s).")
    print(f"Output: {output_path}")
    print(f"Company database: {db_path}")

    # Always refresh these three log files every run - overwrite with fresh
    # content, or delete a stale one from a previous run if there's nothing
    # to report this time. Without this, an old log with outdated (already
    # fixed) errors could linger and be mistaken for current results.
    partial_log = os.path.join(customer_dir, "missing_fields.log")
    if partial_rows:
        with open(partial_log, "w", encoding="utf-8") as f:
            f.write(
                "These rows were written to the output file but are missing one "
                "or more fields - please fill them in directly in the xlsx:\n\n"
            )
            f.write("\n".join(partial_rows))
        print(f"⚠ {len(partial_rows)} row(s) need manual completion: {partial_log}")
    elif os.path.isfile(partial_log):
        os.remove(partial_log)

    ocr_log = os.path.join(customer_dir, "ocr_review.log")
    if ocr_rows:
        with open(ocr_log, "w", encoding="utf-8") as f:
            f.write(
                "These rows were read via OCR (scanned/photographed receipts, "
                "no text layer). OCR can misread digits - please verify each "
                "row below against its original PDF before relying on it:\n\n"
            )
            f.write("\n".join(ocr_rows))
        print(f"⚠ {len(ocr_rows)} row(s) were read via OCR - please review: {ocr_log}")
    elif os.path.isfile(ocr_log):
        os.remove(ocr_log)

    err_log = os.path.join(customer_dir, "errors.log")
    if errors:
        with open(err_log, "w", encoding="utf-8") as f:
            f.write("\n".join(errors))
        print(f"Error log: {err_log}")
    elif os.path.isfile(err_log):
        os.remove(err_log)


def _clean_xls_value(v):
    """xlrd (legacy .xls) always returns numbers as float, e.g. 101 -> 101.0.
    Convert whole-number floats back to a clean string so codes match
    consistently regardless of whether the source file is .xls or .xlsx."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return v


def read_commandtype_grid(path):
    """
    Reads commandtype (.xls or .xlsx) into a uniform grid: a list of rows,
    each row a list of cell values, row 0 being the header row. Handles
    both formats transparently since openpyxl can't open legacy .xls files.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        if xlrd is None:
            raise RuntimeError(
                f'"{path}" is a legacy .xls file - reading it requires: pip install xlrd'
            )
        book = xlrd.open_workbook(path)
        sheet = book.sheet_by_index(0)
        return [
            [_clean_xls_value(sheet.cell_value(r, c)) for c in range(sheet.ncols)]
            for r in range(sheet.nrows)
        ]
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    return [[cell.value for cell in row] for row in ws.iter_rows()]


def cmd_apply_xref(args):
    customer_dir = os.path.join(BASE_DIR, args.customer)
    if not os.path.isdir(customer_dir):
        sys.exit(f"Customer directory not found: {customer_dir}")

    commandtype_path = os.path.join(customer_dir, COMMANDTYPE_FILENAME)
    xls_path = os.path.join(customer_dir, args.output)

    if not os.path.isfile(commandtype_path):
        sys.exit(f"commandtype file not found: {commandtype_path}")
    if not os.path.isfile(xls_path):
        sys.exit(f"Target xls file not found: {xls_path}")

    # Load cross-reference table: קוד -> (חשבון חובה, חשבון זכות)
    grid = read_commandtype_grid(commandtype_path)
    if not grid:
        sys.exit(f"commandtype file is empty: {commandtype_path}")
    ct_header = {_norm(str(v)): i for i, v in enumerate(grid[0]) if v not in (None, "")}
    for required in (CT_CODE, CT_DEBIT, CT_CREDIT):
        if required not in ct_header:
            sys.exit(f"{os.path.basename(commandtype_path)} is missing column '{required}' on row 1")

    def _cell(row, col_idx):
        return row[col_idx] if col_idx is not None and col_idx < len(row) else None

    xref = {}
    for row in grid[1:]:
        code = _cell(row, ct_header[CT_CODE])
        if code in (None, ""):
            continue
        xref[str(code).strip()] = {
            "debit": _cell(row, ct_header[CT_DEBIT]),
            "credit": _cell(row, ct_header[CT_CREDIT]),
        }

    wb = load_workbook(xls_path)
    ws = wb.active
    col_map = find_columns(ws)
    for required in (COL_DEBIT_CODE, COL_CREDIT_CODE):
        if required not in col_map:
            sys.exit(f"Output xls is missing column '{required}' on row {HEADER_ROW}")

    updated, unmatched = 0, []
    row_idx = FIRST_DATA_ROW
    while True:
        order_type = ws.cell(row=row_idx, column=col_map[COL_ORDER_TYPE]).value
        batch_no = ws.cell(row=row_idx, column=col_map[COL_BATCH_NO]).value
        if order_type in (None, "") and batch_no in (None, ""):
            break  # end of data
        if order_type not in (None, ""):
            key = str(order_type).strip()
            if key in xref:
                set_cell(ws, row_idx, col_map[COL_DEBIT_CODE], xref[key]["debit"])
                set_cell(ws, row_idx, col_map[COL_CREDIT_CODE], xref[key]["credit"])
                updated += 1
            else:
                unmatched.append((row_idx, key))
        row_idx += 1

    wb.save(xls_path)
    print(f"Updated {updated} row(s) with account codes.")
    if unmatched:
        print("Rows with 'סוג פקודה' not found in commandtype.xlsx:")
        for r, k in unmatched:
            print(f"  row {r}: '{k}'")


def load_cross_reference(path):
    """
    Reads cross.xlsx: row 1 = category names (one per column), every column
    (including A) holds that category's keyword list in the rows below its
    header. Every non-empty cell below row 1 is a keyword that maps to that
    column's header (category name).

    Returns a list of (keyword, category) pairs in reading order (column by
    column, left to right; top to bottom within each column) - used so that
    matching always checks keywords in a stable, predictable order.
    """
    wb = load_workbook(path)
    ws = wb.active
    pairs = []
    for col in range(1, ws.max_column + 1):  # all columns, including A
        title = ws.cell(row=1, column=col).value
        if title in (None, ""):
            continue
        title = str(title).strip()
        for r in range(2, ws.max_row + 1):
            word = ws.cell(row=r, column=col).value
            if word in (None, ""):
                continue
            pairs.append((str(word).strip(), title))
    return pairs


def find_category_for_text(text, cross_ref_pairs):
    """
    Returns the category name for the LONGEST keyword found as a substring
    of `text` (most-specific match wins - e.g. "חצי חינם" beats a shorter
    "חינם" that also happens to match), or None if no keyword matches at
    all. Ties (same length) go to whichever appears first in cross_ref_pairs.
    """
    if not text:
        return None
    best_category = None
    best_len = -1
    for word, category in cross_ref_pairs:
        if word and word in text and len(word) > best_len:
            best_category = category
            best_len = len(word)
    return best_category


def load_commandtype_by_type(path):
    """
    Reads commandtype (.xls or .xlsx) keyed by the CT_TYPE ("סוג פקודה")
    column - the category name, as opposed to cmd_apply_xref which keys by
    CT_CODE. If a type name appears more than once, the first row wins.

    Returns {type_name: {"code":..., "details":..., "debit":..., "credit":...}}
    """
    grid = read_commandtype_grid(path)
    if not grid:
        raise RuntimeError(f"commandtype file is empty: {path}")
    header = {_norm(str(v)): i for i, v in enumerate(grid[0]) if v not in (None, "")}
    for required in (CT_TYPE, CT_CODE, CT_DEBIT, CT_CREDIT):
        if required not in header:
            raise RuntimeError(f"{os.path.basename(path)} is missing column '{required}' on row 1")
    details_col = header.get(CT_DETAILS)  # optional column

    def _cell(row, col_idx):
        return row[col_idx] if col_idx is not None and col_idx < len(row) else None

    by_type = {}
    for row in grid[1:]:
        type_name = _cell(row, header[CT_TYPE])
        if type_name in (None, ""):
            continue
        type_name = str(type_name).strip()
        if type_name in by_type:
            continue  # first row for this type wins
        by_type[type_name] = {
            "code": _cell(row, header[CT_CODE]),
            "details": _cell(row, details_col),
            "debit": _cell(row, header[CT_DEBIT]),
            "credit": _cell(row, header[CT_CREDIT]),
        }
    return by_type


def cmd_classify(args):
    customer_dir = os.path.join(BASE_DIR, args.customer)
    if not os.path.isdir(customer_dir):
        sys.exit(f"Customer directory not found: {customer_dir}")

    output_path = os.path.join(customer_dir, args.output)
    commandtype_path = os.path.join(customer_dir, COMMANDTYPE_FILENAME)

    if not os.path.isfile(output_path):
        sys.exit(f"Output file not found: {output_path}. Run 'extract' first.")
    if not os.path.isfile(CROSS_REFERENCE_PATH):
        sys.exit(f"Cross-reference file not found: {CROSS_REFERENCE_PATH}")
    if not os.path.isfile(commandtype_path):
        sys.exit(f"commandtype file not found: {commandtype_path}")

    cross_ref_pairs = load_cross_reference(CROSS_REFERENCE_PATH)
    try:
        commandtype_by_type = load_commandtype_by_type(commandtype_path)
    except RuntimeError as e:
        sys.exit(str(e))

    wb = load_workbook(output_path)
    ws = wb.active
    col_map = find_columns(ws)
    if COL_DESCRIPTION not in col_map:
        sys.exit(f'Output file has no "{COL_DESCRIPTION}" column - cannot classify without it.')
    for required in (COL_ORDER_TYPE, COL_DEBIT_CODE, COL_CREDIT_CODE):
        if required not in col_map:
            sys.exit(f"Output file is missing column '{required}' on row {HEADER_ROW}")

    matched = 0
    skipped = 0
    errors = []
    row_idx = FIRST_DATA_ROW
    while True:
        desc_cell_value = ws.cell(row=row_idx, column=col_map[COL_DESCRIPTION]).value
        batch_no = ws.cell(row=row_idx, column=col_map[COL_BATCH_NO]).value
        if desc_cell_value in (None, "") and batch_no in (None, ""):
            break  # end of data

        existing_type = ws.cell(row=row_idx, column=col_map[COL_ORDER_TYPE]).value
        if existing_type not in (None, ""):
            # Already classified in a previous run - skip it so re-running
            # this command doesn't re-append the details suffix to תאור
            # again (which is what caused the earlier triple-concatenation).
            skipped += 1
            row_idx += 1
            continue

        desc_text = str(desc_cell_value or "")
        category = find_category_for_text(desc_text, cross_ref_pairs)

        if category is None:
            msg = f'row {row_idx}: no matching word found in cross.xlsx for תאור: "{desc_text}"'
            errors.append(msg)
            print(f"⚠ {msg}")
            highlight_row(ws, row_idx, col_map)
            row_idx += 1
            continue

        type_row = commandtype_by_type.get(category)
        if type_row is None:
            msg = (f'row {row_idx}: matched category "{category}" via cross.xlsx, but it was '
                   f'not found in commandtype.xlsx column "{CT_TYPE}"')
            errors.append(msg)
            print(f"⚠ {msg}")
            highlight_row(ws, row_idx, col_map)
            row_idx += 1
            continue

        set_cell(ws, row_idx, col_map[COL_ORDER_TYPE], type_row["code"], column_title=COL_ORDER_TYPE)
        if type_row["details"] not in (None, ""):
            new_desc = f'{desc_text}-{type_row["details"]}'
            set_cell(ws, row_idx, col_map[COL_DESCRIPTION], new_desc, column_title=COL_DESCRIPTION)
        set_cell(ws, row_idx, col_map[COL_DEBIT_CODE], type_row["debit"])
        set_cell(ws, row_idx, col_map[COL_CREDIT_CODE], type_row["credit"])
        matched += 1
        row_idx += 1

    wb.save(output_path)
    print(f"\nClassified {matched} row(s). Skipped {skipped} row(s) already classified in a previous run.")

    err_log = os.path.join(customer_dir, "classify_errors.log")
    if errors:
        with open(err_log, "w", encoding="utf-8") as f:
            f.write("\n".join(errors))
        print(f"⚠ {len(errors)} row(s) could not be classified: {err_log}")
    elif os.path.isfile(err_log):
        os.remove(err_log)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def build_parser():
    p = argparse.ArgumentParser(description="Hebrew receipt PDF -> Excel processor")
    sub = p.add_subparsers(dest="command", required=True)

    p_ext = sub.add_parser("extract", help="Extract fields from PDFs into a new xlsx")
    p_ext.add_argument("--customer", required=True,
                        help="Customer name (subfolder under BASE_DIR, e.g. C:\\veka\\<customer>)")
    p_ext.add_argument("--order-no", required=True, help='מספר מנה value to stamp on every row')
    p_ext.add_argument("--receipts-subdir", default=RECEIPTS_SUBDIR_DEFAULT,
                        help=f"Name of the receipts subfolder inside the customer directory "
                             f"(default: '{RECEIPTS_SUBDIR_DEFAULT}')")
    p_ext.add_argument("--output", default=OUTPUT_FILENAME_DEFAULT,
                        help=f"Output xlsx filename, written inside the customer directory "
                             f"(default: '{OUTPUT_FILENAME_DEFAULT}')")
    p_ext.add_argument("--debug-ocr", action="store_true",
                        help="Save the raw extracted text for every receipt to "
                             "<customer>\\ocr_debug\\<filename>.txt, for diagnosing "
                             "why a field was or wasn't found")
    p_ext.set_defaults(func=cmd_extract)

    p_xref = sub.add_parser("apply-xref", help="Fill account codes from commandtype.xlsx (step 8)")
    p_xref.add_argument("--customer", required=True,
                         help="Customer name (subfolder under BASE_DIR)")
    p_xref.add_argument("--output", default=OUTPUT_FILENAME_DEFAULT,
                         help=f"xlsx filename (produced by 'extract', with סוג פקודה already filled "
                              f"in manually) inside the customer directory "
                              f"(default: '{OUTPUT_FILENAME_DEFAULT}')")
    p_xref.set_defaults(func=cmd_apply_xref)

    p_classify = sub.add_parser(
        "classify",
        help="Auto-fill סוג פקודה + account codes by matching תאור against cross.xlsx keywords",
    )
    p_classify.add_argument("--customer", required=True,
                             help="Customer name (subfolder under BASE_DIR)")
    p_classify.add_argument("--output", default=OUTPUT_FILENAME_DEFAULT,
                             help=f"xlsx filename produced by 'extract', inside the customer "
                                  f"directory (default: '{OUTPUT_FILENAME_DEFAULT}')")
    p_classify.set_defaults(func=cmd_classify)

    return p


def main():
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
