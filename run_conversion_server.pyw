#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_conversion_server.pyw
==========================
Browser-based version of the conversion launcher. Double-click this file
(it runs with pythonw.exe, so no black CMD console window appears) and it:

    1. starts a small local web server on http://127.0.0.1:8765/
    2. opens that page in your default browser automatically
    3. shows a form: לקוח / קובץ קלט / שם קובץ פלט / מס' מנה
    4. on submit, runs the correct script for that customer (same as the
       Tkinter version did) and shows the output on the page

The customer -> script mapping lives in customers_config.json, right next
to this file. Add a new customer either by editing that file directly, or
via the "לקוח חדש" link in the page.

To stop the server: click "כבה שרת" in the page, or close it from Task
Manager (look for pythonw.exe / python.exe).

If something breaks and you want to see error messages in a console
window, run this same file as .py instead of .pyw:
    python run_conversion_server.pyw
(rename the extension, or just run it explicitly with `python`).

Dependencies: none beyond the Python standard library.
"""

import html
import json
import os
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs

# ---------------------------------------------------------------------------
BASE_DIR = r"c:\veka\scr"
DATA_DIR = r"c:\veka\datafiles"
CONFIG_NAME = "customers_config.json"
PORT = 8765

PAGE_STYLE = """
<style>
  body { font-family: Segoe UI, Arial, sans-serif; direction: rtl; text-align: right;
         max-width: 640px; margin: 40px auto; padding: 0 16px; color: #222; }
  h1 { font-size: 20px; }
  label { display: block; margin-top: 14px; font-weight: 600; }
  select, input[type=text] { width: 100%; padding: 8px; margin-top: 4px;
        box-sizing: border-box; font-size: 14px; }
  button, input[type=submit] { margin-top: 20px; padding: 10px 18px; font-size: 14px;
        cursor: pointer; border: none; border-radius: 4px; background: #2563eb; color: white; }
  button.secondary, a.secondary { background: #6b7280; color: white; padding: 10px 18px;
        border-radius: 4px; display: inline-block; }
  a.secondary:link, a.secondary:visited, a.secondary:hover, a.secondary:active { color: white; }
  pre { background: #f3f4f6; padding: 12px; border-radius: 4px; white-space: pre-wrap;
        text-align: left; direction: ltr; font-size: 13px; }
  .ok { color: #15803d; font-weight: 600; }
  .err { color: #b91c1c; font-weight: 600; }
  .row { display: flex; gap: 10px; margin-top: 16px; }
  a { color: #2563eb; text-decoration: none; }
</style>
"""


# ---------------------------------------------------------------------------
def config_path():
    return os.path.join(BASE_DIR, CONFIG_NAME)


def load_customers():
    path = config_path()
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return {k: v for k, v in data.items() if not k.startswith("_")}


def save_customers(customers):
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


def list_xlsx_files():
    if not os.path.isdir(DATA_DIR):
        return []
    return sorted(f for f in os.listdir(DATA_DIR) if f.lower().endswith(".xlsx"))


# ---------------------------------------------------------------------------
def render_home(message=""):
    customers = load_customers()
    customer_options = "".join(
        f'<option value="{html.escape(name)}">{html.escape(name)}</option>'
        for name in sorted(customers.keys())
    )
    xlsx_files = list_xlsx_files()
    file_options = "".join(
        f'<option value="{html.escape(f)}">{html.escape(f)}</option>' for f in xlsx_files
    )

    return f"""<!DOCTYPE html>
<html lang="he"><head><meta charset="utf-8"><title>הרצת המרת קובץ לקוח</title>{PAGE_STYLE}</head>
<body>
  <h1>הרצת המרת קובץ לקוח</h1>
  {f'<p>{message}</p>' if message else ''}
  <form method="post" action="/run">
    <label>שם לקוח</label>
    <select name="customer" required>{customer_options}</select>

    <label>קובץ קלט (מתוך {html.escape(DATA_DIR)})</label>
    <select name="input_file">{file_options}</select>
    <label style="font-weight:400;margin-top:6px;">או נתיב מלא ידני (עוקף את הבחירה למעלה):</label>
    <input type="text" name="input_file_manual" placeholder="c:\\veka\\datafiles\\myfile.xlsx">

    <label>שם קובץ פלט (xls)</label>
    <input type="text" name="output_name" required placeholder="journal_2026_01.xls">

    <label>מס' מנה</label>
    <input type="text" name="batch_no" required placeholder="12">

    <div class="row">
      <input type="submit" value="הרץ המרה">
      <a class="secondary" href="/add_customer" style="align-self:center;">+ לקוח חדש</a>
      <a class="secondary" href="/shutdown" style="align-self:center;">כבה שרת</a>
    </div>
  </form>
</body></html>"""


def render_add_customer(message=""):
    return f"""<!DOCTYPE html>
<html lang="he"><head><meta charset="utf-8"><title>לקוח חדש</title>{PAGE_STYLE}</head>
<body>
  <h1>הוספת לקוח חדש</h1>
  {f'<p class="err">{message}</p>' if message else ''}
  <form method="post" action="/add_customer">
    <label>שם לקוח</label>
    <input type="text" name="name" required>

    <label>שם קובץ סקריפט (.py)</label>
    <input type="text" name="script" required placeholder="convertabirxlsx2ram.py">

    <input type="submit" value="שמור">
  </form>
  <p style="margin-top:20px;"><a href="/">&larr; חזרה</a></p>
</body></html>"""


def render_result(cmd, returncode, stdout, stderr):
    status = '<p class="ok">הושלם בהצלחה.</p>' if returncode == 0 \
        else f'<p class="err">הסקריפט נכשל (קוד יציאה {returncode}).</p>'
    output_text = (stdout or "") + (("\n--- שגיאות ---\n" + stderr) if stderr else "")
    return f"""<!DOCTYPE html>
<html lang="he"><head><meta charset="utf-8"><title>תוצאת הרצה</title>{PAGE_STYLE}</head>
<body>
  <h1>תוצאת הרצה</h1>
  <p style="direction:ltr; text-align:left; font-size:12px; color:#555;">{html.escape(' '.join(cmd))}</p>
  {status}
  <pre>{html.escape(output_text) if output_text.strip() else '(אין פלט)'}</pre>
  <p><a href="/">&larr; חזרה להרצה נוספת</a></p>
</body></html>"""


# ---------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # keep quiet, no console anyway when run as .pyw

    def _send_html(self, body, status=200):
        encoded = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def _read_form(self):
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length).decode("utf-8")
        parsed = parse_qs(body)
        return {k: v[0] for k, v in parsed.items()}

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self._send_html(render_home())
        elif self.path == "/add_customer":
            self._send_html(render_add_customer())
        elif self.path == "/shutdown":
            self._send_html(
                "<html><body style='font-family:Segoe UI;direction:rtl;text-align:right;'>"
                "<h2>השרת נסגר. ניתן לסגור את הכרטיסייה הזו.</h2></body></html>"
            )
            threading.Thread(target=self.server.shutdown, daemon=True).start()
        else:
            self._send_html("<h1>404</h1>", status=404)

    def do_POST(self):
        if self.path == "/run":
            self._handle_run()
        elif self.path == "/add_customer":
            self._handle_add_customer()
        else:
            self._send_html("<h1>404</h1>", status=404)

    def _handle_run(self):
        form = self._read_form()
        customer = (form.get("customer") or "").strip()
        input_manual = (form.get("input_file_manual") or "").strip()
        input_selected = (form.get("input_file") or "").strip()
        output_name = (form.get("output_name") or "").strip()
        batch_no = (form.get("batch_no") or "").strip()

        input_file = input_manual or input_selected

        customers = load_customers()
        if not customer or customer not in customers:
            self._send_html(render_home(f'<span class="err">יש לבחור לקוח תקין</span>'))
            return
        if not input_file:
            self._send_html(render_home('<span class="err">יש לבחור או להזין קובץ קלט</span>'))
            return
        if not output_name or not batch_no:
            self._send_html(render_home('<span class="err">יש למלא שם קובץ פלט ומס\' מנה</span>'))
            return

        # resolve input path against DATA_DIR if it's not already a full path
        input_path = input_file if os.path.isabs(input_file) else os.path.join(DATA_DIR, input_file)
        script_path = resolve_script_path(customers[customer])

        if not os.path.exists(script_path):
            self._send_html(render_home(f'<span class="err">קובץ הסקריפט לא נמצא: {html.escape(script_path)}</span>'))
            return
        if not os.path.exists(input_path):
            self._send_html(render_home(f'<span class="err">קובץ הקלט לא נמצא: {html.escape(input_path)}</span>'))
            return

        cmd = [sys.executable, script_path, input_path, output_name, batch_no]
        try:
            result = subprocess.run(
                cmd, cwd=BASE_DIR if os.path.isdir(BASE_DIR) else None,
                capture_output=True, text=True,
            )
            self._send_html(render_result(cmd, result.returncode, result.stdout, result.stderr))
        except Exception as exc:
            self._send_html(render_result(cmd, -1, "", str(exc)))

    def _handle_add_customer(self):
        form = self._read_form()
        name = (form.get("name") or "").strip()
        script = (form.get("script") or "").strip()
        if not name or not script:
            self._send_html(render_add_customer("יש למלא שם לקוח ושם סקריפט"))
            return
        if not script.lower().endswith(".py"):
            script += ".py"
        customers = load_customers()
        customers[name] = script
        save_customers(customers)
        self._send_html(render_home(f'לקוח "{html.escape(name)}" נוסף בהצלחה.'))


def main():
    server = HTTPServer(("127.0.0.1", PORT), Handler)
    url = f"http://127.0.0.1:{PORT}/"
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    server.serve_forever()


if __name__ == "__main__":
    main()
