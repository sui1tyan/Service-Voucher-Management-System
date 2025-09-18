#!/usr/bin/env python3
# Service Voucher Management System (Monolith)
# Deps: customtkinter, reportlab, qrcode, pillow
# Optional deps for Excel export: pandas, openpyxl
# Run:  python main.py

import os
import sys
import sqlite3
import webbrowser
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import customtkinter as ctk
import qrcode
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader

# ---- Optional: Excel export ----
try:
    import pandas as pd  # type: ignore
    _HAS_PANDAS = True
except Exception:
    _HAS_PANDAS = False

# ---------- Wrapped text helpers ----------
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet
_styles = getSampleStyleSheet()
_styleN = _styles["Normal"]

def draw_wrapped(c, text, x, y, w, h, fontsize=10, bold=False, leading=None):
    style = _styleN.clone('wrap')
    style.fontName = "Helvetica-Bold" if bold else "Helvetica"
    style.fontSize = fontsize
    style.leading = leading if leading else fontsize + 2
    para = Paragraph((text or "-").replace("\n", "<br/>"), style)
    w_used, h_used = para.wrap(w, h)
    para.drawOn(c, x, y + h - h_used)
    return h_used

def draw_wrapped_top(c, text, x, top_y, w, fontsize=10, bold=False, leading=None):
    style = _styleN.clone('wrapTop')
    style.fontName = "Helvetica-Bold" if bold else "Helvetica"
    style.fontSize = fontsize
    style.leading = leading if leading else fontsize + 2
    para = Paragraph((text or "-").replace("\n", "<br/>"), style)
    w_used, h_used = para.wrap(w, 1000*mm)
    para.drawOn(c, x, top_y - h_used)
    return h_used

# ------------------ Paths/Config ------------------
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

DB_FILE   = os.path.join(APP_DIR, "vouchers.db")
PDF_DIR   = os.path.join(APP_DIR, "pdfs")
os.makedirs(PDF_DIR, exist_ok=True)

SHOP_NAME = "TONY.COM"
SHOP_ADDR = "TB4318, Lot 5, Block 31, Fajar Complex  91000 Tawau Sabah, Malaysia"
SHOP_TEL  = "Tel : 089-763778, H/P: 0168260533"
LOGO_PATH = ""  # optional: path to logo image (png/jpg). leave blank to skip

DEFAULT_BASE_VID = 41000  # used only at first run

# ------------------ Date helpers (DD-MM-YYYY for UI/PDF) ------------------
def _to_ui_date(dt: datetime) -> str:
    return dt.strftime("%d-%m-%Y")

def _to_ui_datetime_str(iso_str: str) -> str:
    try:
        dt = datetime.strptime(iso_str, "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%d-%m-%Y %H:%M:%S")
    except Exception:
        return iso_str

def _from_ui_date_to_sqldate(s: str):
    s = (s or "").strip()
    if not s:
        return None
    try:
        d = datetime.strptime(s, "%d-%m-%Y")
        return d.strftime("%Y-%m-%d")
    except Exception:
        return None

# ------------------ DB ------------------
BASE_SCHEMA = """
CREATE TABLE IF NOT EXISTS vouchers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    voucher_id TEXT UNIQUE,
    created_at TEXT,
    customer_name TEXT,
    contact_number TEXT,
    units INTEGER DEFAULT 1,
    particulars TEXT,
    problem TEXT,
    staff_name TEXT,
    status TEXT DEFAULT 'Pending',
    recipient TEXT,
    remark TEXT,
    pdf_path TEXT
);
CREATE TABLE IF NOT EXISTS staffs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE
);
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT
);
"""

def _column_exists(cur, table, column):
    cur.execute(f"PRAGMA table_info({table})")
    return any(row[1] == column for row in cur.fetchall())

def _get_setting(cur, key, default=None):
    cur.execute("SELECT value FROM settings WHERE key=?", (key,))
    row = cur.fetchone()
    if row and row[0] is not None:
        return row[0]
    if default is not None:
        cur.execute("INSERT OR IGNORE INTO settings(key,value) VALUES (?,?)", (key, str(default)))
        return str(default)
    return None

def _set_setting(cur, key, value):
    cur.execute("INSERT OR REPLACE INTO settings(key,value) VALUES (?,?)", (key, str(value)))

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.executescript(BASE_SCHEMA)
    # defensive: add any missing columns
    for col, typ in [
        ("voucher_id","TEXT"), ("created_at","TEXT"), ("customer_name","TEXT"),
        ("contact_number","TEXT"), ("units","INTEGER"), ("particulars","TEXT"),
        ("problem","TEXT"), ("staff_name","TEXT"), ("status","TEXT"),
        ("recipient","TEXT"), ("remark","TEXT"), ("pdf_path","TEXT")
    ]:
        if not _column_exists(cur, "vouchers", col):
            cur.execute(f"ALTER TABLE vouchers ADD COLUMN {col} {typ}")
    try:
        cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_vouchers_vid ON vouchers(voucher_id)")
        cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_staffs_name ON staffs(name)")
    except Exception:
        pass
    _get_setting(cur, "base_vid", DEFAULT_BASE_VID)
    conn.commit(); conn.close()

def _read_base_vid():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT value FROM settings WHERE key='base_vid'")
    row = cur.fetchone()
    conn.close()
    return int(row[0]) if row and row[0] else DEFAULT_BASE_VID

def next_voucher_id():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT MAX(CAST(voucher_id AS INTEGER)) FROM vouchers")
    row = cur.fetchone()
    if not row or row[0] is None:
        base = _get_setting(cur, "base_vid", DEFAULT_BASE_VID)
        conn.commit(); conn.close()
        return str(base)
    nxt = str(int(row[0]) + 1)
    conn.close()
    return nxt

# ---- Recipient ops ----
def list_staffs():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT name FROM staffs ORDER BY name COLLATE NOCASE ASC")
    rows = [r[0] for r in cur.fetchall()]
    conn.close()
    return rows

def add_staff(name: str):
    name = (name or "").strip()
    if not name:
        return False
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    try:
        cur.execute("INSERT OR IGNORE INTO staffs (name) VALUES (?)", (name,))
        conn.commit()
    finally:
        conn.close()
    return True

def delete_staff(name: str):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("DELETE FROM staffs WHERE name = ?", (name,))
    conn.commit()
    conn.close()

# ------------------ PDF (helpers) ------------------
def _draw_header(c, left, right, top_y, voucher_id):
    # Logo (optional) on right; shop text on left
    y = top_y
    LOGO_W = 28*mm
    LOGO_H = 18*mm
    logo_drawn = False
    if LOGO_PATH and os.path.exists(LOGO_PATH):
        try:
            c.drawImage(LOGO_PATH, right - LOGO_W, y - LOGO_H, LOGO_W, LOGO_H, preserveAspectRatio=True, mask='auto')
            logo_drawn = True
        except Exception:
            logo_drawn = False

    c.setFont("Helvetica-Bold", 14); c.drawString(left, y, SHOP_NAME)
    c.setFont("Helvetica", 9.2)
    c.drawString(left, y - 5.0*mm, SHOP_ADDR)
    c.drawString(left, y - 9.0*mm, SHOP_TEL)

    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString((left+right)/2, y - 16.0*mm, "SERVICE VOUCHER")
    c.drawRightString(right, y - 16.0*mm, f"No : {voucher_id}")
    return y - 16.0*mm  # returns the baseline y used by title

def _draw_datetime_row(c, left, right, base_y, created_at):
    try:
        dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
        date_str = dt.strftime("%d-%m-%Y")
        time_str = dt.strftime("%H:%M:%S")
    except Exception:
        date_str = created_at[:10]; time_str = created_at[11:19]
    c.setFont("Helvetica", 10)
    c.drawString(left, base_y - 8.0*mm, "Date :")
    c.drawString(left + 18*mm, base_y - 8.0*mm, date_str)
    c.drawRightString(right - 27*mm, base_y - 8.0*mm, "Time In :")
    c.drawRightString(right, base_y - 8.0*mm, time_str)

def _draw_main_table(c, left, right, top_table, customer_name, particulars, units, contact_number, problem):
    qty_col_w    = 20*mm
    left_col_w   = 74*mm
    middle_col_w = (right - left) - left_col_w - qty_col_w
    name_col_x = left + left_col_w
    qty_col_x  = right - qty_col_w
    row1_h = 20*mm
    row2_h = 20*mm
    bottom_table = top_table - (row1_h + row2_h)
    mid_y = top_table - row1_h
    pad = 3*mm

    # Frame
    c.rect(left, bottom_table, right-left, (row1_h + row2_h), stroke=1, fill=0)
    c.line(name_col_x, top_table, name_col_x, bottom_table)
    c.line(qty_col_x, top_table, qty_col_x, mid_y)
    c.line(left, mid_y, right, mid_y)

    # Row 1 Left: CUSTOMER NAME
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left + pad, top_table - pad - 8, "CUSTOMER NAME")
    draw_wrapped(c, customer_name, left + pad, mid_y + pad,
                 w=left_col_w - 2*pad, h=row1_h - 2*pad - 10, fontsize=10)

    # Row 1 Middle: PARTICULARS
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(name_col_x + pad, top_table - pad - 8, "PARTICULARS")
    draw_wrapped(c, particulars or "-", name_col_x + pad, mid_y + pad,
                 w=middle_col_w - 2*pad, h=row1_h - 2*pad - 10, fontsize=10)

    # Row 1 Right: QTY
    c.setFont("Helvetica-Bold", 10.4)
    c.drawCentredString(qty_col_x + qty_col_w/2, top_table - pad - 8, "QTY")
    c.setFont("Helvetica", 11)
    c.drawCentredString(qty_col_x + qty_col_w/2, mid_y + (row1_h/2) - 3, str(max(1, int(units or 1))))

    # Row 2 Left: TEL
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left + pad, mid_y - pad - 8, "TEL")
    draw_wrapped(c, contact_number, left + pad, bottom_table + pad,
                 w=left_col_w - 2*pad, h=row2_h - 2*pad - 10, fontsize=10)

    # Row 2 Middle+Right: PROBLEM
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(name_col_x + pad, mid_y - pad - 8, "PROBLEM")
    draw_wrapped(c, problem or "-", name_col_x + pad, bottom_table + pad,
                 w=(middle_col_w + qty_col_w) - 2*pad, h=row2_h - 2*pad - 10, fontsize=10)

    return bottom_table, left_col_w, name_col_x

def _draw_policies_and_signatures(c, left, right, bottom_table, left_col_w, recipient, voucher_id, customer_name, contact_number, date_str):
    # Acknowledgement sentence (right half)
    ack_text = ("WE HEREBY CONFIRMED THAT THE MACHINE WAS SERVICE AND "
                "REPAIRED SATISFACTORILY")
    ack_left   = left + left_col_w + 10*mm
    ack_right  = right - 6*mm
    ack_width  = max(20*mm, ack_right - ack_left)
    ack_top_y  = bottom_table - 5*mm
    draw_wrapped_top(c, ack_text, ack_left, ack_top_y, ack_width, fontsize=9, bold=True, leading=11)

    # Recipient + line
    y_rec = bottom_table - 9*mm
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left, y_rec, "RECIPIENT :")
    label_w = c.stringWidth("RECIPIENT :", "Helvetica-Bold", 10.4)
    line_x0 = left + label_w + 6
    line_x1 = left + left_col_w - 2*mm
    line_y  = y_rec - 3*mm
    c.line(line_x0, line_y, line_x1, line_y)
    if recipient:
        c.setFont("Helvetica", 9)
        c.drawString(line_x0 + 1*mm, line_y + 2.2*mm, recipient)

    # Policies (left half)
    policies_top = y_rec - 7*mm
    policies_w   = left_col_w - 1.5*mm

    p1 = "Kindly collect your goods within <font color='red' size='9'>60 days</font> from date of sending for repair."
    used_h = draw_wrapped_top(c, p1, left, policies_top, policies_w, fontsize=6.5, leading=10)
    y_cursor = policies_top - used_h - 2
    used_h = draw_wrapped_top(c, "A) We do not hold ourselves responsible for any loss or damage.",
                              left, y_cursor, policies_w, fontsize=6.5, leading=10)
    y_cursor -= used_h - 1
    used_h = draw_wrapped_top(c, "B) We reserve our right to sell off the goods to cover our cost and loss.",
                              left, y_cursor, policies_w, fontsize=6.5, leading=10)
    y_cursor -= used_h + 2
    p4 = ("MINIMUM <font color='red' size='9'><b>RM60.00</b></font> WILL BE CHARGED ON TROUBLESHOOTING, "
          "INSPECTION AND SERVICE ON ALL KIND OF HARDWARE AND SOFTWARE.")
    used_h = draw_wrapped_top(c, p4, left, y_cursor, policies_w, fontsize=8, leading=10)
    y_cursor -= used_h - 1
    used_h = draw_wrapped_top(c, "PLEASE BRING ALONG THIS SERVICE VOUCHER TO COLLECT YOUR GOODS",
                              left, y_cursor, policies_w, fontsize=8, leading=10)
    y_cursor -= used_h - 1
    used_h = draw_wrapped_top(c, "NO ATTENTION GIVEN WITHOUT SERVICE VOUCHER",
                              left, y_cursor, policies_w, fontsize=8, leading=10)
    policies_bottom = y_cursor - used_h

    # QR — anchor near bottom-right safely above signature lines
    qr_size = 20*mm
    try:
        qr_data = f"Voucher:{voucher_id}|Name:{customer_name}|Tel:{contact_number}|Date:{date_str}"
        qr_img  = qrcode.make(qr_data)
        margin = 10*mm
        qr_x = right - qr_size
        qr_y = max(policies_bottom + 3*mm, margin + qr_size)  # ensure not below bottom margin
        c.drawImage(ImageReader(qr_img), qr_x, qr_y - qr_size, qr_size, qr_size)
    except Exception:
        pass

    # Signatures
    SIG_LINE_W = 45*mm
    SIG_GAP    = 6*mm
    y_sig = max(policies_bottom + 4*mm, (A4[1]/2) - 20*mm)  # avoid too high/low

    sig_left_start = right - (2*SIG_LINE_W + SIG_GAP)
    c.line(sig_left_start, y_sig, sig_left_start + SIG_LINE_W, y_sig)
    c.setFont("Helvetica", 8.8)
    c.drawString(sig_left_start, y_sig - 3.6*mm, "CUSTOMER SIGNATURE")

    right_line_x0 = sig_left_start + SIG_LINE_W + SIG_GAP
    c.line(right_line_x0, y_sig, right_line_x0 + SIG_LINE_W, y_sig)
    c.drawString(right_line_x0, y_sig - 3.6*mm, "DATE COLLECTED")

def _draw_voucher(c, width, height, voucher_id, customer_name, contact_number,
                  units, particulars, problem, staff_name, created_at, recipient):
    # Margins
    left   = 12*mm
    right  = width - 12*mm
    top_y  = height - 15*mm

    # Header + Title
    title_baseline = _draw_header(c, left, right, top_y, voucher_id)
    _draw_datetime_row(c, left, right, title_baseline, created_at)

    # Main table
    top_table = title_baseline - 12*mm
    bottom_table, left_col_w, name_col_x = _draw_main_table(
        c, left, right, top_table, customer_name, particulars, units, contact_number, problem
    )

    # Date string for QR composition
    try:
        dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
        date_str = dt.strftime("%d-%m-%Y")
    except Exception:
        date_str = created_at[:10]

    # Policies + QR + signatures
    _draw_policies_and_signatures(
        c, left, right, bottom_table, left_col_w, recipient,
        voucher_id, customer_name, contact_number, date_str
    )

def generate_pdf(voucher_id, customer_name, contact_number, units,
                 particulars, problem, staff_name, status, created_at, recipient):
    filename = os.path.join(PDF_DIR, f"voucher_{voucher_id}.pdf")
    c = rl_canvas.Canvas(filename, pagesize=A4)
    width, height = A4
    _draw_voucher(c, width, height, voucher_id, customer_name, contact_number,
                  units, particulars, problem, staff_name, created_at, recipient)
    c.showPage(); c.save()
    return filename

# ------------------ DB ops ------------------
def add_voucher(customer_name, contact_number, units, particulars, problem, staff_name, recipient="", remark=""):
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()

    voucher_id = next_voucher_id()
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    status = "Pending"

    pdf_path = generate_pdf(
        voucher_id, customer_name, contact_number, units,
        particulars, problem, staff_name, status, created_at, recipient
    )

    cur.execute("""
        INSERT INTO vouchers (voucher_id, created_at, customer_name, contact_number, units,
                              particulars, problem, staff_name, status, recipient, remark, pdf_path)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
    """, (voucher_id, created_at, customer_name, contact_number, units,
          particulars, problem, staff_name, status, recipient, remark, pdf_path))

    conn.commit(); conn.close()
    return voucher_id, pdf_path

def mark_completed(voucher_id):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("UPDATE vouchers SET status='Completed' WHERE voucher_id = ?", (voucher_id,))
    conn.commit()
    conn.close()

def mark_deleted(voucher_id):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("UPDATE vouchers SET status='Deleted' WHERE voucher_id = ?", (voucher_id,))
    conn.commit()
    conn.close()

def _build_search_sql(filters):
    sql = ("SELECT voucher_id, created_at, customer_name, contact_number, units, "
           "recipient, status, remark, pdf_path "
           "FROM vouchers WHERE 1=1")
    params = []
    if filters.get("voucher_id"):
        sql += " AND voucher_id LIKE ?"; params.append(f"%{filters['voucher_id']}%")
    if filters.get("name"):
        sql += " AND customer_name LIKE ? COLLATE NOCASE"; params.append(f"%{filters['name']}%")
    if filters.get("contact"):
        sql += " AND contact_number LIKE ?"; params.append(f"%{filters['contact']}%")

    df = filters.get("date_from")
    dt = filters.get("date_to")
    if df:
        sql += " AND created_at >= ?"; params.append(df + " 00:00:00")
    if dt:
        sql += " AND created_at <= ?"; params.append(dt + " 23:59:59")

    if filters.get("status") and filters["status"] != "All":
        sql += " AND status = ?"; params.append(filters["status"])

    sql += " ORDER BY CAST(voucher_id AS INTEGER) DESC"
    return sql, params

def search_vouchers(filters):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    sql, params = _build_search_sql(filters)
    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()
    return rows

def modify_base_vid(new_base: int):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT MIN(CAST(voucher_id AS INTEGER)) FROM vouchers")
    row = cur.fetchone()
    if not row or row[0] is None:
        _set_setting(cur, "base_vid", new_base)
        conn.commit(); conn.close()
        return 0
    current_min = int(row[0])
    delta = int(new_base) - current_min
    if delta == 0:
        _set_setting(cur, "base_vid", new_base)
        conn.commit(); conn.close()
        return 0

    order = "DESC" if delta > 0 else "ASC"
    cur.execute(f"""
        SELECT voucher_id, created_at, customer_name, contact_number, units,
               particulars, problem, staff_name, status, recipient, remark, pdf_path
        FROM vouchers
        ORDER BY CAST(voucher_id AS INTEGER) {order}
    """)
    rows = cur.fetchall()

    for (vid, created_at, customer_name, contact_number, units,
         particulars, problem, staff_name, status, recipient, remark, old_pdf) in rows:
        old_id = int(vid)
        new_id = old_id + delta
        try:
            if old_pdf and os.path.exists(old_pdf):
                os.remove(old_pdf)
        except Exception:
            pass
        cur.execute("UPDATE vouchers SET voucher_id=? WHERE voucher_id=?", (str(new_id), str(old_id)))
        new_pdf = generate_pdf(
            str(new_id), customer_name, contact_number, int(units or 1),
            particulars, problem, staff_name, status, created_at, recipient
        )
        cur.execute("UPDATE vouchers SET pdf_path=? WHERE voucher_id=?", (new_pdf, str(new_id)))

    _set_setting(cur, "base_vid", new_base)
    conn.commit(); conn.close()
    return delta

# ------------------ UI ------------------
class VoucherApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Service Voucher Management System")
        self.geometry("1200x720")
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.minsize(980, 560)

        # Always open fullscreen / maximized
        self.after(50, self._go_fullscreen)

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # ---- Filters ----
        filt = ctk.CTkFrame(self)
        filt.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        for i in range(12):
            filt.grid_columnconfigure(i, weight=0)
        filt.grid_columnconfigure(11, weight=1)

        today_ui = _to_ui_date(datetime.now())
        self.f_voucher = ctk.CTkEntry(filt, width=130, placeholder_text="VoucherID");    self.f_voucher.grid(row=0, column=0, padx=5)
        self.f_name    = ctk.CTkEntry(filt, width=220, placeholder_text="Customer Name"); self.f_name.grid(row=0, column=1, padx=5)
        self.f_contact = ctk.CTkEntry(filt, width=180, placeholder_text="Contact Number"); self.f_contact.grid(row=0, column=2, padx=5)
        self.f_from    = ctk.CTkEntry(filt, width=160, placeholder_text="Date From (DD-MM-YYYY)"); self.f_from.grid(row=0, column=3, padx=5)
        self.f_to      = ctk.CTkEntry(filt, width=160, placeholder_text="Date To (DD-MM-YYYY)");   self.f_to.grid(row=0, column=4, padx=5)
        # default: show everything until today
        self.f_to.insert(0, today_ui)

        self.f_status  = ctk.CTkOptionMenu(filt, values=["All","Pending","Completed","Deleted"], width=140); self.f_status.grid(row=0, column=5, padx=(8,5))
        self.f_status.set("All")
        self.btn_search = ctk.CTkButton(filt, text="Search", command=self.perform_search, width=100); self.btn_search.grid(row=0, column=6, padx=5)
        self.btn_reset  = ctk.CTkButton(filt, text="Reset",  command=self.reset_filters, width=80);   self.btn_reset.grid(row=0, column=7, padx=5)

        # ---- Export buttons ----
        exp_frame = ctk.CTkFrame(filt)
        exp_frame.grid(row=0, column=8, padx=10)
        ctk.CTkButton(exp_frame, text="Export CSV", command=self.export_csv, width=110).pack(side="left", padx=5)
        ctk.CTkButton(exp_frame, text="Export Excel", command=self.export_excel, width=120).pack(side="left", padx=5)

        # ---- Table (Remark moved to far-right; Recipient before Status) ----
        table_frame = ctk.CTkFrame(self)
        table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 8))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_frame,
            columns=("VoucherID","Date","Customer","Contact","Units","Recipient","Status","Remark","PDF"),
            show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew")

        table_vbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        table_hbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=table_vbar.set, xscrollcommand=table_hbar.set)
        table_vbar.grid(row=0, column=1, sticky="ns")
        table_hbar.grid(row=1, column=0, sticky="ew")

        self.col_weights = {
            "VoucherID":1, "Date":2, "Customer":2, "Contact":2,
            "Units":1, "Recipient":2, "Status":1, "Remark":3
        }
        for col in ("VoucherID","Date","Customer","Contact","Units","Recipient","Status","Remark"):
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="w", stretch=True, width=80)
        self.tree.heading("PDF", text="PDF")
        self.tree.column("PDF", width=0, anchor="w", stretch=False)

        def _autosize_columns(event=None):
            total_weight = sum(self.col_weights.values())
            usable = max(self.tree.winfo_width()-40, 300)
            for col, wt in self.col_weights.items():
                self.tree.column(col, width=int(usable * wt/total_weight))
        table_frame.bind("<Configure>", _autosize_columns)

        # Row highlighting by status
        self.tree.tag_configure("Pending", background="#FFF4B3", foreground="#333333")
        self.tree.tag_configure("Completed", background="#CDEEC8", foreground="#223322")
        self.tree.tag_configure("Deleted", background="#F8D7DA", foreground="#6A1B1A")

        self.tree.bind("<Double-1>", lambda e: self.open_pdf())

        # ---- Actions ----
        bar = ctk.CTkFrame(self)
        bar.grid(row=2, column=0, sticky="ew", padx=10, pady=(0,10))
        for i in range(14): bar.grid_columnconfigure(i, weight=0)

        ctk.CTkButton(bar, text="Add Voucher",        command=self.add_voucher_ui,    width=120).grid(row=0, column=0, padx=6, pady=8)
        ctk.CTkButton(bar, text="Edit Selected",      command=self.edit_selected,     width=120).grid(row=0, column=1, padx=6, pady=8)
        ctk.CTkButton(bar, text="Mark Completed",     command=self.mark_selected_completed, width=140).grid(row=0, column=2, padx=6, pady=8)
        ctk.CTkButton(bar, text="Mark Deleted",       command=self.mark_selected_deleted,   width=130, fg_color="#d9534f", hover_color="#c9302c").grid(row=0, column=3, padx=6, pady=8)
        ctk.CTkButton(bar, text="Open PDF",           command=self.open_pdf,          width=110).grid(row=0, column=4, padx=6, pady=8)
        ctk.CTkButton(bar, text="Open PDF Folder",    command=self.open_pdf_folder,   width=150).grid(row=0, column=5, padx=6, pady=8)
        ctk.CTkButton(bar, text="Manage Recipients",  command=self.manage_staffs_ui,  width=160).grid(row=0, column=6, padx=6, pady=8)
        ctk.CTkButton(bar, text="Modify Base VID",    command=self.modify_base_vid_ui, width=150).grid(row=0, column=7, padx=6, pady=8)

        self.perform_search()

    def _go_fullscreen(self):
        """Open maximized/fullscreen depending on platform."""
        try:
            if sys.platform.startswith("win") or sys.platform.startswith("linux"):
                self.state("zoomed")
            else:
                self.attributes("-fullscreen", True)
        except Exception:
            self.attributes("-fullscreen", True)

    # ---- Filters helpers ----
    def _get_filters(self):
        return {
            "voucher_id": self.f_voucher.get().strip(),
            "name": self.f_name.get().strip(),
            "contact": self.f_contact.get().strip(),
            "date_from": _from_ui_date_to_sqldate(self.f_from.get().strip()),
            "date_to": _from_ui_date_to_sqldate(self.f_to.get().strip()),
            "status": self.f_status.get(),
        }

    def reset_filters(self):
        for e in (self.f_voucher, self.f_name, self.f_contact, self.f_from, self.f_to):
            e.delete(0, "end")
        today_ui = _to_ui_date(datetime.now())
        self.f_to.insert(0, today_ui)
        self.f_status.set("All")
        self.perform_search()

    def perform_search(self):
        rows = search_vouchers(self._get_filters())
        self.tree.delete(*self.tree.get_children())
        # rows order: (vid, created_at, customer, contact, units, recipient, status, remark, pdf)
        for (vid, created_at, customer, contact, units, recipient, status, remark, pdf) in rows:
            self.tree.insert("", "end", values=(
                vid,
                _to_ui_datetime_str(created_at),
                customer, contact, units,
                recipient or "", status, remark or "", pdf
            ), tags=(status,))
        self.tree.update_idletasks()

    # ---- Create Voucher ----
    def add_voucher_ui(self):
        top = ctk.CTkToplevel(self)
        top.title("Create Voucher")
        top.geometry("900x740")
        top.grab_set()

        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=12, pady=12)
        frm.grid_columnconfigure(1, weight=1)

        WIDE = 540
        r = 0
        ctk.CTkLabel(frm, text="Customer Name", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        e_name = ctk.CTkEntry(frm, width=WIDE); e_name.grid(row=r, column=1, sticky="ew", padx=10, pady=(0,8)); r+=1

        ctk.CTkLabel(frm, text="Contact Number", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        e_contact = ctk.CTkEntry(frm, width=WIDE); e_contact.grid(row=r, column=1, sticky="ew", padx=10, pady=(0,8)); r+=1

        ctk.CTkLabel(frm, text="No. of Units", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        e_units = ctk.CTkEntry(frm, width=120); e_units.insert(0,"1")
        e_units.grid(row=r, column=1, sticky="w", padx=10, pady=(0,8)); r+=1

        # Text with scrollbar helper
        def mk_text(parent, width, height):
            wrap = ctk.CTkFrame(parent)
            wrap.grid_columnconfigure(0, weight=1)
            txt = tk.Text(wrap, width=width, height=height)
            sb  = ttk.Scrollbar(wrap, orient="vertical", command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            txt.grid(row=0, column=0, sticky="nsew")
            sb.grid(row=0, column=1, sticky="ns")
            return wrap, txt

        ctk.CTkLabel(frm, text="Particulars", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        part_wrap, t_part = mk_text(frm, width=66, height=6); part_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=(0,8)); r+=1

        ctk.CTkLabel(frm, text="Problem", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        prob_wrap, t_prob = mk_text(frm, width=66, height=5); prob_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=(0,8)); r+=1

        ctk.CTkLabel(frm, text="Recipient", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        staff_values = list_staffs() or [""]
        e_recipient = ctk.CTkComboBox(frm, values=staff_values, width=WIDE)
        if staff_values and staff_values[0] != "":
            e_recipient.set(staff_values[0])
        e_recipient.grid(row=r, column=1, sticky="w", padx=10, pady=(0,8)); r+=1

        # Remark
        ctk.CTkLabel(frm, text="Remark", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        rem_wrap, t_remark = mk_text(frm, width=66, height=4)
        rem_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=(0,8)); r+=1

        btns = ctk.CTkFrame(top); btns.pack(fill="x", padx=12, pady=(0,12))

        def save():
            name = e_name.get().strip()
            contact = e_contact.get().strip()
            # Validate units positive int
            try:
                units = int((e_units.get() or "1").strip())
                if units <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid", "Units must be a positive integer."); return
            particulars = t_part.get("1.0","end").strip()
            problem     = t_prob.get("1.0","end").strip()
            recipient   = e_recipient.get().strip()
            staff_name  = recipient
            remark      = t_remark.get("1.0","end").strip()

            if not name or not contact:
                messagebox.showerror("Missing", "Customer name and contact are required."); return
            voucher_id, pdf_path = add_voucher(name, contact, units, particulars, problem, staff_name, recipient, remark)
            messagebox.showinfo("Saved", f"Voucher {voucher_id} created.")
            try: webbrowser.open_new(os.path.abspath(pdf_path))
            except Exception: pass
            top.destroy(); self.perform_search()

        ctk.CTkButton(btns, text="Save & Open PDF", command=save, width=180).pack(side="right")

    # ---- Manage Recipients ----
    def manage_staffs_ui(self):
        top = ctk.CTkToplevel(self)
        top.title("Manage Recipients")
        top.geometry("680x480")
        top.grab_set()

        root = ctk.CTkFrame(top)
        root.pack(fill="both", expand=True, padx=14, pady=14)
        root.grid_rowconfigure(2, weight=1)
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=0)

        entry = ctk.CTkEntry(root, placeholder_text="New recipient name")
        entry.grid(row=0, column=0, sticky="ew", padx=(0,10), pady=(0,10))

        row1 = ctk.CTkFrame(root)
        row1.grid(row=1, column=0, sticky="w", padx=(0,10), pady=(0,10))
        add_btn = ctk.CTkButton(row1, text="Add", width=120)
        del_btn = ctk.CTkButton(row1, text="Delete Selected", width=160)
        add_btn.pack(side="left", padx=(0,10))
        del_btn.pack(side="left")

        list_frame = ctk.CTkFrame(root)
        list_frame.grid(row=2, column=0, sticky="nsew", padx=(0,10))
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        lb = tk.Listbox(list_frame, height=14)
        lb.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=lb.yview)
        sb.grid(row=0, column=1, sticky="ns")
        lb.configure(yscrollcommand=sb.set)

        close_btn = ctk.CTkButton(root, text="Close", width=120, command=top.destroy)
        close_btn.grid(row=3, column=1, sticky="e", pady=(10,0))

        def refresh_list():
            lb.delete(0, "end")
            for s in list_staffs():
                lb.insert("end", s)

        def do_add():
            if add_staff(entry.get()):
                entry.delete(0, "end")
                refresh_list()

        def do_del():
            sel = lb.curselection()
            if not sel: return
            name = lb.get(sel[0])
            delete_staff(name)
            refresh_list()

        add_btn.configure(command=do_add)
        del_btn.configure(command=do_del)
        refresh_list()

    # ---- Row actions ----
    def _selected_voucher(self):
        sel = self.tree.focus()
        if not sel:
            return None
        values = self.tree.item(sel)["values"]
        if not values:
            return None
        return values  # [vid, created_at, ... , pdf]

    def mark_selected_completed(self):
        values = self._selected_voucher()
        if not values:
            messagebox.showerror("Error", "Select a record first."); return
        voucher_id = values[0]
        mark_completed(voucher_id)
        messagebox.showinfo("Updated", f"Voucher {voucher_id} marked as Completed.")
        self.perform_search()

    def mark_selected_deleted(self):
        values = self._selected_voucher()
        if not values:
            messagebox.showerror("Error", "Select a record first."); return
        voucher_id = str(values[0])
        if not messagebox.askyesno("Confirm Delete (Soft)",
                                   f"Mark voucher {voucher_id} as Deleted?\n"
                                   f"This will NOT renumber other vouchers."):
            return
        try:
            mark_deleted(voucher_id)
            messagebox.showinfo("Deleted", f"Voucher {voucher_id} marked as Deleted.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed: {e}")
        self.perform_search()

    def open_pdf(self):
        values = self._selected_voucher()
        if not values:
            messagebox.showerror("Error", "Select a record first."); return
        pdf_path = values[-1]
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Error", "PDF not found for this voucher."); return
        try:
            webbrowser.open_new(os.path.abspath(pdf_path))
        except Exception:
            if os.name == "nt":
                os.startfile(pdf_path)  # type: ignore
            else:
                os.system(f"open '{pdf_path}'" if sys.platform == "darwin" else f"xdg-open '{pdf_path}'")

    def open_pdf_folder(self):
        folder = PDF_DIR
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore
            elif sys.platform == "darwin":
                os.system(f"open '{folder}'")
            else:
                os.system(f"xdg-open '{folder}'")
        except Exception as e:
            messagebox.showerror("Error", f"Unable to open folder:\n{e}")

    def edit_selected(self):
        values = self._selected_voucher()
        if not values:
            messagebox.showerror("Error", "Select a record first."); return

        voucher_id = values[0]

        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("""
            SELECT voucher_id, created_at, customer_name, contact_number, units,
                   particulars, problem, staff_name, status, recipient, remark, pdf_path
            FROM vouchers WHERE voucher_id = ?
        """, (voucher_id,))
        row = cur.fetchone()
        conn.close()
        if not row: return

        _, created_at, customer_name, contact_number, units, particulars, problem, staff_name, status, recipient, remark, old_pdf_path = row

        top = ctk.CTkToplevel(self)
        top.title(f"Edit Voucher {voucher_id}")
        top.geometry("900x740")
        top.grab_set()

        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=12, pady=12)
        frm.grid_columnconfigure(1, weight=1)

        WIDE = 540
        r = 0
        ctk.CTkLabel(frm, text="Customer Name").grid(row=r, column=0, sticky="w")
        e_name = ctk.CTkEntry(frm, width=WIDE); e_name.insert(0, customer_name)
        e_name.grid(row=r, column=1, sticky="ew", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Contact Number").grid(row=r, column=0, sticky="w")
        e_contact = ctk.CTkEntry(frm, width=WIDE); e_contact.insert(0, contact_number)
        e_contact.grid(row=r, column=1, sticky="ew", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="No. of Units").grid(row=r, column=0, sticky="w")
        e_units = ctk.CTkEntry(frm, width=120); e_units.insert(0, str(units))
        e_units.grid(row=r, column=1, sticky="w", padx=10, pady=6); r+=1

        def mk_text(parent, width, height):
            wrap = ctk.CTkFrame(parent)
            wrap.grid_columnconfigure(0, weight=1)
            txt = tk.Text(wrap, width=width, height=height)
            sb  = ttk.Scrollbar(wrap, orient="vertical", command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            txt.grid(row=0, column=0, sticky="nsew")
            sb.grid(row=0, column=1, sticky="ns")
            return wrap, txt

        ctk.CTkLabel(frm, text="Particulars").grid(row=r, column=0, sticky="w")
        part_wrap, t_part = mk_text(frm, width=66, height=6); t_part.insert("1.0", particulars or "")
        part_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Problem").grid(row=r, column=0, sticky="w")
        prob_wrap, t_prob = mk_text(frm, width=66, height=5); t_prob.insert("1.0", problem or "")
        prob_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Recipient").grid(row=r, column=0, sticky="w")
        staff_values = list_staffs() or [""]
        e_recipient = ctk.CTkComboBox(frm, values=staff_values, width=WIDE)
        e_recipient.set(recipient or staff_name or "")
        e_recipient.grid(row=r, column=1, sticky="w", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Remark").grid(row=r, column=0, sticky="w")
        rem_wrap, t_remark = mk_text(frm, width=66, height=4); t_remark.insert("1.0", remark or "")
        rem_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6); r+=1

        def save_edit():
            name = e_name.get().strip()
            contact = e_contact.get().strip()
            try:
                units_val = int((e_units.get() or "1").strip())
                if units_val <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid", "Units must be a positive integer."); return
            particulars_val = t_part.get("1.0","end").strip()
            problem_val     = t_prob.get("1.0","end").strip()
            recipient_val   = e_recipient.get().strip()
            remark_val      = t_remark.get("1.0","end").strip()
            staff_val       = recipient_val

            conn = sqlite3.connect(DB_FILE)
            cur = conn.cursor()
            cur.execute("""UPDATE vouchers
                              SET customer_name=?, contact_number=?, units=?, particulars=?,
                                  problem=?, staff_name=?, recipient=?, remark=?
                            WHERE voucher_id=?""",
                        (name, contact, units_val, particulars_val,
                         problem_val, staff_val, recipient_val, remark_val, voucher_id))
            conn.commit()
            # regenerate PDF (remove old first)
            try:
                if old_pdf_path and os.path.exists(old_pdf_path):
                    os.remove(old_pdf_path)
            except Exception:
                pass

            pdf_path = generate_pdf(
                str(voucher_id), name, contact, units_val,
                particulars_val, problem_val, staff_val, status, created_at, recipient_val
            )
            cur.execute("UPDATE vouchers SET pdf_path=? WHERE voucher_id=?", (pdf_path, str(voucher_id)))
            conn.commit(); conn.close()

            messagebox.showinfo("Updated", f"Voucher {voucher_id} updated.")
            top.destroy(); self.perform_search()

        btns = ctk.CTkFrame(top); btns.pack(fill="x", padx=12, pady=12)
        ctk.CTkButton(btns, text="Save Changes", command=save_edit, width=160).pack(side="right", padx=6)
        ctk.CTkButton(btns, text="Cancel", command=top.destroy, width=100).pack(side="right", padx=6)

    # ---- Modify Base VID UI ----
    def modify_base_vid_ui(self):
        top = ctk.CTkToplevel(self)
        top.title("Modify Base Voucher ID")
        top.geometry("420x220")
        top.grab_set()

        frm = ctk.CTkFrame(top); frm.pack(fill="both", expand=True, padx=16, pady=16)

        current_base = _read_base_vid()
        ctk.CTkLabel(frm, text=f"Current Base VID: {current_base}", anchor="w").pack(fill="x", pady=(0,8))
        entry = ctk.CTkEntry(frm, placeholder_text="Enter new base voucher ID (integer)")
        entry.pack(fill="x"); entry.insert(0, str(current_base))
        info = ctk.CTkLabel(frm, text="Existing vouchers will shift by the difference. PDFs will be regenerated.", justify="left", wraplength=360)
        info.pack(fill="x", pady=8)

        def apply():
            new_base_str = entry.get().strip()
            if not new_base_str.isdigit():
                messagebox.showerror("Invalid", "Please enter a positive integer."); return
            new_base = int(new_base_str)
            try:
                delta = modify_base_vid(new_base)
                messagebox.showinfo("Done", f"Base VID set to {new_base}. Shift: {delta:+d}.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to modify base VID:\n{e}")
            top.destroy(); self.perform_search()

        btns = ctk.CTkFrame(frm); btns.pack(fill="x", pady=(8,0))
        ctk.CTkButton(btns, text="Apply", command=apply, width=120).pack(side="right", padx=4)
        ctk.CTkButton(btns, text="Cancel", command=top.destroy, width=100).pack(side="right", padx=4)

    # ---- Export helpers ----
    def _rows_for_export(self):
        filters = self._get_filters()
        rows = search_vouchers(filters)
        # Convert to list of dicts for exporting
        export_rows = []
        for (vid, created_at, customer, contact, units, recipient, status, remark, pdf) in rows:
            export_rows.append({
                "VoucherID": vid,
                "CreatedAt": created_at,
                "Customer": customer,
                "Contact": contact,
                "Units": units,
                "Recipient": recipient or "",
                "Status": status,
                "Remark": remark or "",
                "PDFPath": pdf or "",
            })
        return export_rows

    def export_csv(self):
        rows = self._rows_for_export()
        if not rows:
            messagebox.showinfo("Export", "No rows to export."); return
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV", "*.csv")],
                                            title="Save CSV")
        if not path:
            return
        try:
            # CSV without pandas to keep deps minimal
            import csv
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
                writer.writeheader()
                writer.writerows(rows)
            messagebox.showinfo("Export", f"CSV exported:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export CSV:\n{e}")

    def export_excel(self):
        if not _HAS_PANDAS:
            messagebox.showerror("Excel Export", "Excel export requires 'pandas' and 'openpyxl'.\nInstall: pip install pandas openpyxl")
            return
        rows = self._rows_for_export()
        if not rows:
            messagebox.showinfo("Export", "No rows to export."); return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel Workbook", "*.xlsx")],
                                            title="Save Excel")
        if not path:
            return
        try:
            df = pd.DataFrame(rows)
            # Use engine=openpyxl explicitly to ensure correct format/extension
            df.to_excel(path, index=False, engine="openpyxl")
            messagebox.showinfo("Export", f"Excel exported:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export Excel:\n{e}")

# ------------------ Run ------------------
if __name__ == "__main__":
    init_db()
    app = VoucherApp()
    app.mainloop()
