# -*- coding: utf-8 -*-
"""
receipt_upload_server.py
==========================

A tiny local web app that lets you scan a receipt on your phone and have it
land directly in the correct customer's receipts folder - no Google Drive,
no phone gallery, no manual copy step.

HOW IT WORKS
------------
Run this on your Windows machine (same computer as receipt_processor.py).
It starts a small local web server. On your phone (connected to the SAME
WiFi network as the computer), open a browser and go to the address printed
when you start the server (something like http://192.168.1.23:5000).

You'll see a page with:
  1. A dropdown to pick the customer (auto-detected from folders under
     C:\veka)
  2. A "Take Photo" button that opens your phone's camera directly
  3. After snapping a photo, it uploads immediately and converts it to a
     PDF, saved straight into that customer's "receipts" folder - ready
     for receipt_processor.py's "extract" command.

Because the browser's camera capture is used directly (not your Photos
app), the photo is not saved to your phone's gallery in most cases - see
the note in the chat reply for the couple of phone/browser combinations
where a copy may still be kept by the OS.

SETUP
-----
    pip install flask pillow

RUN
---
    python receipt_upload_server.py

Then, on your phone (same WiFi), open the URL printed in the console.

STOPPING
--------
Press Ctrl+C in the console window running this script.
"""

import io
import os
import socket
import sys
import uuid
from datetime import datetime

from flask import Flask, request, jsonify, Response
from PIL import Image, ImageOps
import cv2
import numpy as np

# --------------------------------------------------------------------------
# Configuration - matches receipt_processor.py's layout
# --------------------------------------------------------------------------

BASE_DIR = r"C:\veka"
RECEIPTS_SUBDIR = "receipts"
PORT = 5000

# Folder names under BASE_DIR that are NOT customers (shared/system folders)
NON_CUSTOMER_DIRS = {"scr"}

app = Flask(__name__)


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def list_customers():
    if not os.path.isdir(BASE_DIR):
        return []
    names = []
    for entry in sorted(os.listdir(BASE_DIR)):
        full = os.path.join(BASE_DIR, entry)
        if os.path.isdir(full) and entry not in NON_CUSTOMER_DIRS and not entry.startswith("."):
            names.append(entry)
    return names


def get_local_ip():
    """Best-effort local network IP so we can print a phone-friendly URL."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        return s.getsockname()[0]
    except Exception:
        return "127.0.0.1"
    finally:
        s.close()


# --------------------------------------------------------------------------
# "Scanner" image processing - detect the receipt's edges, flatten the
# perspective so it looks straight-on (like a real scanner), then enhance
# contrast so it reads like a clean scan instead of a photo.
# --------------------------------------------------------------------------

def _order_points(pts):
    rect = np.zeros((4, 2), dtype="float32")
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]      # top-left
    rect[2] = pts[np.argmax(s)]      # bottom-right
    diff = np.diff(pts, axis=1)
    rect[1] = pts[np.argmin(diff)]   # top-right
    rect[3] = pts[np.argmax(diff)]   # bottom-left
    return rect


def _four_point_transform(image, pts):
    rect = _order_points(pts)
    (tl, tr, br, bl) = rect
    width = max(int(np.linalg.norm(br - bl)), int(np.linalg.norm(tr - tl)))
    height = max(int(np.linalg.norm(tr - br)), int(np.linalg.norm(tl - bl)))
    dst = np.array([[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]], dtype="float32")
    M = cv2.getPerspectiveTransform(rect, dst)
    return cv2.warpPerspective(image, M, (width, height))


def _four_corners_from_contours(contours, img_area, min_area_fraction):
    """Tries to find a 4-point polygon among the largest contours."""
    contours = sorted(contours, key=cv2.contourArea, reverse=True)[:10]
    for c in contours:
        peri = cv2.arcLength(c, True)
        approx = cv2.approxPolyDP(c, 0.02 * peri, True)
        if len(approx) == 4 and cv2.contourArea(approx) > min_area_fraction * img_area:
            return approx.reshape(4, 2)
    return None


def _find_document_contour(image):
    """
    Returns the 4 corner points of the largest document-like quadrilateral
    in the image (scaled back to original resolution), or None if nothing
    confident enough was found. Tries a few different strategies since real
    photos vary a lot in lighting/background/shadows - a single fixed-
    threshold approach is too brittle for that.
    """
    ratio = 800.0 / image.shape[0]
    small = cv2.resize(image, (int(image.shape[1] * ratio), 800))
    gray = cv2.cvtColor(small, cv2.COLOR_BGR2GRAY)
    img_area = small.shape[0] * small.shape[1]
    blur = cv2.GaussianBlur(gray, (5, 5), 0)

    # Strategy 1: Canny edge detection at a couple of threshold pairs
    # (different lighting/contrast needs different sensitivity).
    for low, high in [(50, 150), (30, 100), (75, 200)]:
        edged = cv2.Canny(blur, low, high)
        edged = cv2.dilate(edged, None, iterations=2)
        contours, _ = cv2.findContours(edged.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        corners = _four_corners_from_contours(contours, img_area, 0.1)
        if corners is not None:
            return (corners / ratio).astype("float32")

    # Strategy 2: adaptive threshold - handles uneven lighting/shadows
    # better than a single fixed Canny setting.
    thresh = cv2.adaptiveThreshold(blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                    cv2.THRESH_BINARY_INV, 11, 2)
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, np.ones((9, 9), np.uint8))
    contours, _ = cv2.findContours(thresh.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    corners = _four_corners_from_contours(contours, img_area, 0.1)
    if corners is not None:
        return (corners / ratio).astype("float32")

    # Strategy 3: fall back to the minimum-area rotated bounding box of the
    # single largest contour found by either method above, as long as it's
    # a substantial chunk of the frame - better than giving up entirely on
    # a receipt that's roughly rectangular but has a slightly irregular
    # edge (creases, torn corner, etc).
    all_contours = []
    for source in (cv2.Canny(blur, 50, 150), thresh):
        cs, _ = cv2.findContours(source.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        all_contours.extend(cs)
    if all_contours:
        largest = max(all_contours, key=cv2.contourArea)
        if cv2.contourArea(largest) > 0.15 * img_area:
            rect = cv2.minAreaRect(largest)
            box = cv2.boxPoints(rect)
            return (box / ratio).astype("float32")

    return None


def scan_style_image(pil_image):
    """
    Takes a PIL image straight from the phone camera and returns a PIL image
    that looks like it came from a document scanner: detects the receipt's
    edges and flattens the perspective if a confident match is found, then
    applies adaptive contrast enhancement either way. Falls back gracefully
    (contrast-enhance only, no crop) if edge detection isn't confident -
    better to keep the whole photo than risk cropping into the receipt.
    """
    cv_img = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)

    # Cap resolution early - phone photos can be 10+ megapixels, which is
    # far more than needed for OCR and makes both processing and the final
    # PDF unnecessarily slow/large. 2000px on the long side is still sharp
    # enough for text extraction.
    h, w = cv_img.shape[:2]
    max_dim = max(h, w)
    if max_dim > 2000:
        scale = 2000 / max_dim
        cv_img = cv2.resize(cv_img, (int(w * scale), int(h * scale)))

    contour = _find_document_contour(cv_img)
    if contour is not None:
        cv_img = _four_point_transform(cv_img, contour)
        print("Scanner: document edges detected - cropped and straightened.")
    else:
        print("Scanner: could not confidently detect document edges - "
              "applied contrast enhancement only, kept full photo framing.")

    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)

    return Image.fromarray(enhanced).convert("RGB")


# --------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------

PAGE_HTML = """<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>סריקת קבלות</title>
<style>
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, "Segoe UI", Arial, sans-serif;
    background: #f4f5f7;
    margin: 0;
    padding: 20px 16px 40px;
    color: #1a1a1a;
  }
  h1 { font-size: 20px; margin: 0 0 20px; text-align: center; }
  .card {
    background: white;
    border-radius: 14px;
    padding: 20px;
    max-width: 420px;
    margin: 0 auto 16px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
  }
  label { display: block; font-size: 14px; color: #555; margin-bottom: 6px; }
  select {
    width: 100%;
    padding: 12px;
    font-size: 16px;
    border-radius: 10px;
    border: 1px solid #ccc;
    margin-bottom: 16px;
    background: white;
  }
  .capture-btn {
    display: block;
    width: 100%;
    padding: 18px;
    font-size: 17px;
    font-weight: 600;
    border: none;
    border-radius: 12px;
    background: #2563eb;
    color: white;
    cursor: pointer;
    text-align: center;
  }
  .capture-btn:active { background: #1d4ed8; }
  .capture-btn:disabled { background: #93c5fd; }
  input[type=file] { display: none; }
  #status {
    max-width: 420px;
    margin: 0 auto;
    text-align: center;
    font-size: 15px;
    min-height: 24px;
  }
  .ok { color: #15803d; font-weight: 600; }
  .err { color: #b91c1c; font-weight: 600; }
  .spinner {
    display: inline-block;
    width: 16px; height: 16px;
    border: 2px solid #cbd5e1;
    border-top-color: #2563eb;
    border-radius: 50%;
    animation: spin 0.7s linear infinite;
    vertical-align: middle;
    margin-left: 8px;
  }
  @keyframes spin { to { transform: rotate(360deg); } }
  #recent {
    max-width: 420px;
    margin: 20px auto 0;
    font-size: 13px;
    color: #666;
  }
  #recent div { padding: 4px 0; }
</style>
</head>
<body>
  <h1>📄 סריקת קבלות</h1>

  <div class="card">
    <label for="customer">לקוח</label>
    <select id="customer"></select>
    <div id="selectedLabel" style="display:none; font-size:13px; color:#2563eb; font-weight:600; margin:-8px 0 14px;"></div>

    <button class="capture-btn" id="captureBtn" disabled>📷 צלם קבלה</button>
    <input type="file" id="fileInput" accept="image/*" capture="environment">
  </div>

  <div id="status"></div>
  <div id="recent"></div>

<script>
const customerSelect = document.getElementById('customer');
const captureBtn = document.getElementById('captureBtn');
const fileInput = document.getElementById('fileInput');
const statusEl = document.getElementById('status');
const recentEl = document.getElementById('recent');
const selectedLabel = document.getElementById('selectedLabel');
let recentList = [];

async function loadCustomers() {
  const res = await fetch('/api/customers');
  const data = await res.json();
  customerSelect.innerHTML = '';

  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = '-- בחר לקוח --';
  placeholder.disabled = true;
  placeholder.selected = true;
  customerSelect.appendChild(placeholder);

  if (data.customers.length === 0) {
    captureBtn.disabled = true;
    statusEl.innerHTML = '<span class="err">לא נמצאו תיקיות לקוחות תחת ' + 'C:\\veka' + '</span>';
    return;
  }
  data.customers.forEach(name => {
    const opt = document.createElement('option');
    opt.value = name;
    opt.textContent = name;
    customerSelect.appendChild(opt);
  });

  captureBtn.disabled = true;  // stays disabled until a real customer is picked
  const saved = localStorage.getItem('lastCustomer');
  if (saved && data.customers.includes(saved)) {
    customerSelect.value = saved;
    captureBtn.disabled = false;
  }
  updateSelectedLabel();
}

function updateSelectedLabel() {
  if (customerSelect.value) {
    selectedLabel.textContent = 'סורק עבור: ' + customerSelect.value;
    selectedLabel.style.display = 'block';
  } else {
    selectedLabel.style.display = 'none';
  }
}

customerSelect.addEventListener('change', () => {
  captureBtn.disabled = !customerSelect.value;
  updateSelectedLabel();
});

captureBtn.addEventListener('click', () => fileInput.click());

function resizeImageForUpload(file, maxDim, quality) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = () => {
      URL.revokeObjectURL(url);
      let { width, height } = img;
      const scale = Math.min(1, maxDim / Math.max(width, height));
      width = Math.round(width * scale);
      height = Math.round(height * scale);
      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      canvas.getContext('2d').drawImage(img, 0, 0, width, height);
      canvas.toBlob(blob => {
        if (blob) resolve(blob);
        else reject(new Error('canvas toBlob failed'));
      }, 'image/jpeg', quality);
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('image load failed')); };
    img.src = url;
  });
}

fileInput.addEventListener('change', async () => {
  const file = fileInput.files[0];
  if (!file) return;
  const customer = customerSelect.value;
  localStorage.setItem('lastCustomer', customer);

  captureBtn.disabled = true;
  statusEl.innerHTML = 'מכין תמונה<span class="spinner"></span>';

  let uploadFile = file;
  try {
    // Shrink client-side before upload - phone photos can be 5-10+ MB,
    // which is slow to send over WiFi and is more resolution than needed
    // for text extraction anyway. This is the main fix for the upload
    // timing out on slower connections.
    uploadFile = await resizeImageForUpload(file, 2000, 0.85);
  } catch (e) {
    uploadFile = file;  // fall back to the original if resizing fails for any reason
  }

  statusEl.innerHTML = 'מעלה ושומר<span class="spinner"></span>';

  const form = new FormData();
  form.append('customer', customer);
  form.append('photo', uploadFile, 'photo.jpg');

  try {
    const res = await fetch('/api/upload', { method: 'POST', body: form });
    const data = await res.json();
    if (res.ok) {
      statusEl.innerHTML = '<span class="ok">✓ נשמר: ' + data.filename + '</span>';
      recentList.unshift(customer + ' / ' + data.filename);
      recentList = recentList.slice(0, 5);
      recentEl.innerHTML = recentList.map(r => '<div>✓ ' + r + '</div>').join('');
    } else {
      statusEl.innerHTML = '<span class="err">שגיאה: ' + (data.error || 'unknown') + '</span>';
    }
  } catch (e) {
    statusEl.innerHTML = '<span class="err">שגיאת רשת: ' + e + '</span>';
  } finally {
    captureBtn.disabled = false;
    fileInput.value = '';
  }
});

loadCustomers();
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return Response(PAGE_HTML, mimetype="text/html")


@app.route("/api/customers")
def api_customers():
    return jsonify({"customers": list_customers()})


@app.route("/api/upload", methods=["POST"])
def api_upload():
    customer = request.form.get("customer", "").strip()
    if not customer:
        return jsonify({"error": "No customer selected"}), 400
    if customer not in list_customers():
        return jsonify({"error": f'Customer folder "{customer}" not found under {BASE_DIR}'}), 400

    photo = request.files.get("photo")
    if not photo or photo.filename == "":
        return jsonify({"error": "No photo received"}), 400

    receipts_dir = os.path.join(BASE_DIR, customer, RECEIPTS_SUBDIR)
    os.makedirs(receipts_dir, exist_ok=True)

    # Convert the captured image straight to PDF so it's immediately
    # compatible with receipt_processor.py's "extract" command (which only
    # looks for *.pdf files).
    try:
        img = Image.open(io.BytesIO(photo.read()))
        img = ImageOps.exif_transpose(img)  # apply the phone's rotation metadata to the actual pixels
        img = img.convert("RGB")
    except Exception as e:
        return jsonify({"error": f"Could not read image: {e}"}), 400

    # Give it a "scanned" look: detect the receipt's edges, flatten the
    # perspective, and boost contrast - if anything goes wrong here, fall
    # back to the original photo rather than losing the receipt entirely.
    try:
        img = scan_style_image(img)
    except Exception as e:
        print(f"Note: scanner processing failed ({e}); saving the original photo instead.")

    filename = f"scan_{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:6]}.pdf"
    out_path = os.path.join(receipts_dir, filename)
    try:
        img.save(out_path, "PDF")
    except Exception as e:
        return jsonify({"error": f"Could not save PDF: {e}"}), 500

    return jsonify({"filename": filename, "path": out_path})


# --------------------------------------------------------------------------

if __name__ == "__main__":
    ip = get_local_ip()
    print("=" * 60)
    print("Receipt scanning server is starting.")
    print(f"On your phone (same WiFi), open:   http://{ip}:{PORT}")
    print("Press Ctrl+C to stop.")
    print("=" * 60)
    app.run(host="0.0.0.0", port=PORT, debug=False, threaded=True)
