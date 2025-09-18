#!/usr/bin/env python3
# Service Voucher Management System (SVMS) — Monolith
# Core deps: customtkinter, reportlab, qrcode, pillow, bcrypt
# Optional for Excel export: pandas, openpyxl
# Run: python main.py

import os, sys, io, json, zipfile, shutil, sqlite3, webbrowser, re
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import simpledialog
import customtkinter as ctk
import qrcode
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from PIL import Image, ImageOps
import bcrypt

# ---- Optional: Excel export ----
try:
    import pandas as pd  # type: ignore
    _HAS_PANDAS = True
except Exception:
    _HAS_PANDAS = False

# ---------- Wrapped text helpers for PDF ----------
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
    _, h_used = para.wrap(w, h)
    para.drawOn(c, x, y + h - h_used)
    return h_used

def draw_wrapped_top(c, text, x, top_y, w, fontsize=10, bold=False, leading=None):
    style = _styleN.clone('wrapTop')
    style.fontName = "Helvetica-Bold" if bold else "Helvetica"
    style.fontSize = fontsize
    style.leading = leading if leading else fontsize + 2
    para = Paragraph((text or "-").replace("\n", "<br/>"), style)
    _, h_used = para.wrap(w, 1000*mm)
    para.drawOn(c, x, top_y - h_used)
    return h_used

# ------------------ Paths/Config ------------------
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

DB_FILE   = os.path.join(APP_DIR, "vouchers.db")
PDF_DIR   = os.path.join(APP_DIR, "pdfs")
IMG_DIR   = os.path.join(APP_DIR, "images")
IMG_STAFF = os.path.join(IMG_DIR, "staff")
IMG_BILLS = os.path.join(IMG_DIR, "bills")
os.makedirs(PDF_DIR, exist_ok=True)
for d in (IMG_DIR, IMG_STAFF, IMG_BILLS): os.makedirs(d, exist_ok=True)

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
    staff_name TEXT,            -- kept for legacy
    status TEXT DEFAULT 'Pending',
    recipient TEXT,
    solution TEXT,              -- was 'remark'
    pdf_path TEXT,
    technician_id TEXT,
    technician_name TEXT
);
CREATE TABLE IF NOT EXISTS staffs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    position TEXT,
    staff_id_opt TEXT,
    name TEXT,
    phone TEXT,
    photo_path TEXT,
    created_at TEXT,
    updated_at TEXT
);
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT
);
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE,
    role TEXT CHECK(role IN ('admin','supervisor','user')),
    password_hash BLOB,
    is_active INTEGER DEFAULT 1,
    must_change_pwd INTEGER DEFAULT 0,
    created_at TEXT,
    updated_at TEXT
);
CREATE TABLE IF NOT EXISTS commissions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_id INTEGER REFERENCES staffs(id) ON DELETE CASCADE,
    bill_type TEXT CHECK(bill_type IN ('CS','IV')),
    bill_no TEXT,
    total_amount REAL,
    commission_amount REAL,
    bill_image_path TEXT,
    created_at TEXT,
    updated_at TEXT
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

def _hash_pwd(pwd:str) -> bytes:
    return bcrypt.hashpw(pwd.encode("utf-8"), bcrypt.gensalt())

def _verify_pwd(pwd:str, hp:bytes) -> bool:
    try:
        return bcrypt.checkpw(pwd.encode("utf-8"), hp)
    except Exception:
        return False

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.executescript(BASE_SCHEMA)

    # migrate legacy 'remark' -> 'solution'
    if not _column_exists(cur, "vouchers", "solution"):
        cur.execute("ALTER TABLE vouchers ADD COLUMN solution TEXT")
    # move data if old column existed
    try:
        cur.execute("PRAGMA table_info(vouchers)")
        cols = [c[1] for c in cur.fetchall()]
        if "remark" in cols:
            cur.execute("UPDATE vouchers SET solution = COALESCE(solution, remark)")
    except Exception:
        pass

    # ensure new tech columns exist
    for col, typ in [("technician_id","TEXT"), ("technician_name","TEXT")]:
        if not _column_exists(cur, "vouchers", col):
            cur.execute(f"ALTER TABLE vouchers ADD COLUMN {col} {typ}")

    try:
        cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_vouchers_vid ON vouchers(voucher_id)")
        cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username ON users(username)")
        cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_staffs_name ON staffs(name)")
    except Exception:
        pass

    # settings defaults
    _get_setting(cur, "base_vid", DEFAULT_BASE_VID)

    # bootstrap admin if none
    cur.execute("SELECT COUNT(*) FROM users WHERE role='admin'")
    if cur.fetchone()[0] == 0:
        cur.execute(
            "INSERT INTO users (username, role, password_hash, created_at, updated_at) VALUES (?,?,?,?,?)",
            ("tonycom", "admin", _hash_pwd("1234567890!@#$%^&*()"),
             datetime.now().isoformat(sep=" ", timespec="seconds"),
             datetime.now().isoformat(sep=" ", timespec="seconds"))
        )

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

# ---- Recipient ops (simple list) ----
def list_staffs_names():
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
    cur.execute("SELECT name FROM staffs ORDER BY name COLLATE NOCASE ASC")
    rows = [r[0] for r in cur.fetchall()]
    conn.close(); return rows

def add_staff_simple(name: str):
    name = (name or "").strip()
    if not name: return False
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
    try:
        cur.execute("""INSERT OR IGNORE INTO staffs
                       (position, staff_id_opt, name, phone, photo_path, created_at, updated_at)
                       VALUES ('Technician', '', ?, '', '', ?, ?)""",
                    (name, datetime.now().isoformat(sep=" ", timespec="seconds"),
                     datetime.now().isoformat(sep=" ", timespec="seconds")))
        conn.commit()
    finally:
        conn.close()
    return True

def delete_staff_simple(name: str):
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
    cur.execute("DELETE FROM staffs WHERE name = ?", (name,))
    conn.commit(); conn.close()

# ------------------ PDF (helpers) ------------------
def _draw_header(c, left, right, top_y, voucher_id):
    y = top_y
    LOGO_W = 28*mm; LOGO_H = 18*mm
    if LOGO_PATH and os.path.exists(LOGO_PATH):
        try:
            c.drawImage(LOGO_PATH, right - LOGO_W, y - LOGO_H, LOGO_W, LOGO_H, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass
    c.setFont("Helvetica-Bold", 14); c.drawString(left, y, SHOP_NAME)
    c.setFont("Helvetica", 9.2)
    c.drawString(left, y - 5.0*mm, SHOP_ADDR)
    c.drawString(left, y - 9.0*mm, SHOP_TEL)
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString((left+right)/2, y - 16.0*mm, "SERVICE VOUCHER")
    c.drawRightString(right, y - 16.0*mm, f"No : {voucher_id}")
    return y - 16.0*mm

def _draw_datetime_row(c, left, right, base_y, created_at):
    try:
        dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
        date_str = dt.strftime("%d-%m-%Y"); time_str = dt.strftime("%H:%M:%S")
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

    c.rect(left, bottom_table, right-left, (row1_h + row2_h), stroke=1, fill=0)
    c.line(name_col_x, top_table, name_col_x, bottom_table)
    c.line(qty_col_x, top_table, qty_col_x, mid_y)
    c.line(left, mid_y, right, mid_y)

    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left + pad, top_table - pad - 8, "CUSTOMER NAME")
    draw_wrapped(c, customer_name, left + pad, mid_y + pad,
                 w=left_col_w - 2*pad, h=row1_h - 2*pad - 10, fontsize=10)

    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(name_col_x + pad, top_table - pad - 8, "PARTICULARS")
    draw_wrapped(c, particulars or "-", name_col_x + pad, mid_y + pad,
                 w=middle_col_w - 2*pad, h=row1_h - 2*pad - 10, fontsize=10)

    c.setFont("Helvetica-Bold", 10.4)
    c.drawCentredString(qty_col_x + qty_col_w/2, top_table - pad - 8, "QTY")
    c.setFont("Helvetica", 11)
    c.drawCentredString(qty_col_x + qty_col_w/2, mid_y + (row1_h/2) - 3, str(max(1, int(units or 1))))

    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left + pad, mid_y - pad - 8, "TEL")
    draw_wrapped(c, contact_number, left + pad, bottom_table + pad,
                 w=left_col_w - 2*pad, h=row2_h - 2*pad - 10, fontsize=10)

    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(name_col_x + pad, mid_y - pad - 8, "PROBLEM")
    draw_wrapped(c, problem or "-", name_col_x + pad, bottom_table + pad,
                 w=(middle_col_w + qty_col_w) - 2*pad, h=row2_h - 2*pad - 10, fontsize=10)

    return bottom_table, left_col_w

def _draw_policies_and_signatures(c, left, right, bottom_table, left_col_w, recipient, voucher_id, customer_name, contact_number, date_str):
    ack_text = ("WE HEREBY CONFIRMED THAT THE MACHINE WAS SERVICE AND "
                "REPAIRED SATISFACTORILY")
    ack_left   = left + left_col_w + 10*mm
    ack_right  = right - 6*mm
    ack_top_y  = bottom_table - 5*mm
    draw_wrapped_top(c, ack_text, ack_left, ack_top_y, max(20*mm, ack_right-ack_left), fontsize=9, bold=True, leading=11)

    y_rec = bottom_table - 9*mm
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left, y_rec, "RECIPIENT :")
    label_w = c.stringWidth("RECIPIENT :", "Helvetica-Bold", 10.4)
    line_x0 = left + label_w + 6
    line_x1 = left + left_col_w - 2*mm
    line_y  = y_rec - 3*mm
    c.line(line_x0, line_y, line_x1, line_y)
    if recipient:
        c.setFont("Helvetica", 9); c.drawString(line_x0 + 1*mm, line_y + 2.2*mm, recipient)

    policies_top = y_rec - 7*mm
    policies_w   = left_col_w - 1.5*mm

    p1 = "Kindly collect your goods within <font color='red' size='9'>60 days</font> from date of sending for repair."
    used_h = draw_wrapped_top(c, p1, left, policies_top, policies_w, fontsize=6.5, leading=10)
    y_cursor = policies_top - used_h - 2
    used_h = draw_wrapped_top(c, "A) We do not hold ourselves responsible for any loss or damage.", left, y_cursor, policies_w, fontsize=6.5, leading=10)
    y_cursor -= used_h - 1
    used_h = draw_wrapped_top(c, "B) We reserve our right to sell off the goods to cover our cost and loss.", left, y_cursor, policies_w, fontsize=6.5, leading=10)
    y_cursor -= used_h + 2
    p4 = ("MINIMUM <font color='red' size='9'><b>RM60.00</b></font> WILL BE CHARGED ON TROUBLESHOOTING, "
          "INSPECTION AND SERVICE ON ALL KIND OF HARDWARE AND SOFTWARE.")
    used_h = draw_wrapped_top(c, p4, left, y_cursor, policies_w, fontsize=8, leading=10)
    y_cursor -= used_h - 1
    used_h = draw_wrapped_top(c, "PLEASE BRING ALONG THIS SERVICE VOUCHER TO COLLECT YOUR GOODS", left, y_cursor, policies_w, fontsize=8, leading=10)
    y_cursor -= used_h - 1
    used_h = draw_wrapped_top(c, "NO ATTENTION GIVEN WITHOUT SERVICE VOUCHER", left, y_cursor, policies_w, fontsize=8, leading=10)
    policies_bottom = y_cursor - used_h

    qr_size = 20*mm
    try:
        qr_data = f"Voucher:{voucher_id}|Name:{customer_name}|Tel:{contact_number}|Date:{date_str}"
        qr_img  = qrcode.make(qr_data)
        qr_x = right - qr_size
        qr_y = max(policies_bottom + 3*mm, 10*mm + qr_size)
        c.drawImage(ImageReader(qr_img), qr_x, qr_y - qr_size, qr_size, qr_size)
    except Exception:
        pass

    SIG_LINE_W = 45*mm; SIG_GAP = 6*mm
    y_sig = max(policies_bottom + 4*mm, (A4[1]/2) - 20*mm)
    sig_left_start = right - (2*SIG_LINE_W + SIG_GAP)
    c.line(sig_left_start, y_sig, sig_left_start + SIG_LINE_W, y_sig)
    c.setFont("Helvetica", 8.8); c.drawString(sig_left_start, y_sig - 3.6*mm, "CUSTOMER SIGNATURE")
    right_line_x0 = sig_left_start + SIG_LINE_W + SIG_GAP
    c.line(right_line_x0, y_sig, right_line_x0 + SIG_LINE_W, y_sig)
    c.drawString(right_line_x0, y_sig - 3.6*mm, "DATE COLLECTED")

def generate_pdf(voucher_id, customer_name, contact_number, units,
                 particulars, problem, staff_name, status, created_at, recipient):
    filename = os.path.join(PDF_DIR, f"voucher_{voucher_id}.pdf")
    c = rl_canvas.Canvas(filename, pagesize=A4)
    width, height = A4
    left, right, top_y = 12*mm, width - 12*mm, height - 15*mm
    title_baseline = _draw_header(c, left, right, top_y, voucher_id)
    _draw_datetime_row(c, left, right, title_baseline, created_at)
    top_table = title_baseline - 12*mm
    bottom_table, left_col_w = _draw_main_table(c, left, right, top_table, customer_name, particulars, units, contact_number, problem)
    try:
        dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S"); date_str = dt.strftime("%d-%m-%Y")
    except Exception: date_str = created_at[:10]
    _draw_policies_and_signatures(c, left, right, bottom_table, left_col_w, recipient, voucher_id, customer_name, contact_number, date_str)
    c.showPage(); c.save()
    return filename

# ------------------ DB ops (vouchers) ------------------
def add_voucher(customer_name, contact_number, units, particulars, problem, staff_name, recipient="", solution="", technician_id="", technician_name=""):
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()

    voucher_id = next_voucher_id()
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    status = "Pending"

    pdf_path = generate_pdf(voucher_id, customer_name, contact_number, units, particulars, problem, staff_name, status, created_at, recipient)

    cur.execute("""
        INSERT INTO vouchers (voucher_id, created_at, customer_name, contact_number, units,
                              particulars, problem, staff_name, status, recipient, solution, pdf_path,
                              technician_id, technician_name)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (voucher_id, created_at, customer_name, contact_number, units, particulars, problem, staff_name,
          status, recipient, solution, pdf_path, technician_id, technician_name))

    conn.commit(); conn.close()
    return voucher_id, pdf_path

def update_voucher_fields(voucher_id, **fields):
    if not fields: return
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
    cols, params = [], []
    for k, v in fields.items():
        cols.append(f"{k}=?"); params.append(v)
    params.append(voucher_id)
    cur.execute(f"UPDATE vouchers SET {', '.join(cols)} WHERE voucher_id = ?", params)
    conn.commit(); conn.close()

def bulk_update_status(voucher_ids, new_status):
    if not voucher_ids: return
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
    cur.executemany("UPDATE vouchers SET status=? WHERE voucher_id=?", [(new_status, vid) for vid in voucher_ids])
    conn.commit(); conn.close()

def _build_search_sql(filters):
    sql = ("SELECT voucher_id, created_at, customer_name, contact_number, units, "
           "recipient, technician_id, technician_name, status, solution, pdf_path "
           "FROM vouchers WHERE 1=1")
    params = []
    if filters.get("voucher_id"):
        sql += " AND voucher_id LIKE ?"; params.append(f"%{filters['voucher_id']}%")
    if filters.get("name"):
        sql += " AND customer_name LIKE ? COLLATE NOCASE"; params.append(f"%{filters['name']}%")
    if filters.get("contact"):
        sql += " AND contact_number LIKE ?"; params.append(f"%{filters['contact']}%")
    df = filters.get("date_from"); dt = filters.get("date_to")
    if df: sql += " AND created_at >= ?"; params.append(df + " 00:00:00")
    if dt: sql += " AND created_at <= ?"; params.append(dt + " 23:59:59")
    if filters.get("status") and filters["status"] != "All":
        sql += " AND status = ?"; params.append(filters["status"])
    sql += " ORDER BY CAST(voucher_id AS INTEGER) DESC"
    return sql, params

def search_vouchers(filters):
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
    sql, params = _build_search_sql(filters)
    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()
    return rows

def modify_base_vid(new_base: int):
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
    cur.execute("SELECT MIN(CAST(voucher_id AS INTEGER)) FROM vouchers")
    row = cur.fetchone()
    if not row or row[0] is None:
        _set_setting(cur, "base_vid", new_base); conn.commit(); conn.close(); return 0
    current_min = int(row[0]); delta = int(new_base) - current_min
    if delta == 0:
        _set_setting(cur, "base_vid", new_base); conn.commit(); conn.close(); return 0

    order = "DESC" if delta > 0 else "ASC"
    cur.execute(f"""
        SELECT voucher_id, created_at, customer_name, contact_number, units,
               particulars, problem, staff_name, status, recipient, solution, pdf_path,
               technician_id, technician_name
        FROM vouchers
        ORDER BY CAST(voucher_id AS INTEGER) {order}
    """)
    rows = cur.fetchall()

    for (vid, created_at, customer_name, contact_number, units, particulars, problem, staff_name, status, recipient, solution, old_pdf, tech_id, tech_name) in rows:
        old_id = int(vid); new_id = old_id + delta
        try:
            if old_pdf and os.path.exists(old_pdf): os.remove(old_pdf)
        except Exception: pass
        cur.execute("UPDATE vouchers SET voucher_id=? WHERE voucher_id=?", (str(new_id), str(old_id)))
        new_pdf = generate_pdf(str(new_id), customer_name, contact_number, int(units or 1),
                               particulars, problem, staff_name, status, created_at, recipient)
        cur.execute("UPDATE vouchers SET pdf_path=? WHERE voucher_id=?", (new_pdf, str(new_id)))

    _set_setting(cur, "base_vid", new_base)
    conn.commit(); conn.close()
    return delta

# ------------------ Users (auth & admin) ------------------
def get_user_by_username(u):
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
    cur.execute("SELECT id, username, role, password_hash, is_active, must_change_pwd FROM users WHERE username=?", (u,))
    row = cur.fetchone(); conn.close(); return row

def list_users():
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
    cur.execute("SELECT id, username, role, is_active, must_change_pwd FROM users ORDER BY role, username")
    rows = cur.fetchall(); conn.close(); return rows

def create_user(username, role, password):
    if role not in ("supervisor","user","admin"): raise ValueError("Invalid role")
    # enforce single admin
    if role == "admin":
        conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM users WHERE role='admin'"); n = cur.fetchone()[0]
        conn.close()
        if n >= 1: raise ValueError("Only one admin allowed")
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
    cur.execute("""INSERT INTO users (username, role, password_hash, created_at, updated_at)
                   VALUES (?,?,?,?,?)""",
                (username, role, _hash_pwd(password),
                 datetime.now().isoformat(sep=" ", timespec="seconds"),
                 datetime.now().isoformat(sep=" ", timespec="seconds")))
    conn.commit(); conn.close()

def update_user(user_id, **fields):
    if not fields: return
    if "role" in fields and fields["role"] == "admin":
        conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM users WHERE role='admin' AND id<>?", (user_id,)); n = cur.fetchone()[0]
        conn.close()
        if n >= 1: raise ValueError("Only one admin allowed")
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
    cols, params = [], []
    for k, v in fields.items():
        cols.append(f"{k}=?"); params.append(v)
    params.append(user_id)
    cur.execute(f"UPDATE users SET {', '.join(cols)}, updated_at=? WHERE id=?", params[:-1] + [datetime.now().isoformat(sep=" ", timespec="seconds"), user_id])
    conn.commit(); conn.close()

def reset_password(user_id, new_pwd):
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
    cur.execute("UPDATE users SET password_hash=?, must_change_pwd=1, updated_at=? WHERE id=?",
                (_hash_pwd(new_pwd), datetime.now().isoformat(sep=" ", timespec="seconds"), user_id))
    conn.commit(); conn.close()

def delete_user(user_id):
    conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
    cur.execute("SELECT role FROM users WHERE id=?", (user_id,))
    row = cur.fetchone()
    if row and row[0] == "admin":
        conn.close(); raise ValueError("Cannot delete admin")
    cur.execute("DELETE FROM users WHERE id=?", (user_id,))
    conn.commit(); conn.close()

# ------------------ Staff utilities ------------------
def _process_square_image(path_in, path_out, max_px=400):
    img = Image.open(path_in).convert("RGB")
    img = ImageOps.exif_transpose(img)
    size = min(img.width, img.height)
    left = (img.width - size)//2; top = (img.height - size)//2
    img = img.crop((left, top, left+size, top+size))
    if size > max_px:
        img = img.resize((max_px, max_px), Image.LANCZOS)
    img.save(path_out, format="JPEG", quality=92)

# ------------------ Commission utilities ------------------
BILL_RE_CS = re.compile(r"^CS-(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])/\d{4}$")
BILL_RE_IV = re.compile(r"^IV-(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])/\d{4}$")

# ------------------ UI ------------------
FONT_FAMILY = "Segoe UI"
UI_FONT_SIZE = 14  # (req. 14)

class LoginDialog(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Login")
        self.geometry("380x220")
        self.resizable(False, False)
        self.grab_set()
        frm = ctk.CTkFrame(self); frm.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(frm, text="Username").grid(row=0, column=0, sticky="w", pady=(0,6))
        self.e_user = ctk.CTkEntry(frm, width=220); self.e_user.grid(row=0, column=1, sticky="w", pady=(0,6))
        ctk.CTkLabel(frm, text="Password").grid(row=1, column=0, sticky="w", pady=(0,6))
        self.e_pwd = ctk.CTkEntry(frm, width=220, show="•"); self.e_pwd.grid(row=1, column=1, sticky="w", pady=(0,6))
        self.var_show = tk.BooleanVar(value=False)
        chk = ctk.CTkCheckBox(frm, text="Show", variable=self.var_show, command=self._toggle); chk.grid(row=1, column=2, padx=(6,0))
        btns = ctk.CTkFrame(frm); btns.grid(row=2, column=0, columnspan=3, sticky="e", pady=(12,0))
        self.btn_login = ctk.CTkButton(btns, text="Login", command=self._do_login, width=120); self.btn_login.pack(side="right")
        self.result = None

    def _toggle(self):
        self.e_pwd.configure(show="" if self.var_show.get() else "•")

    def _do_login(self):
        u = (self.e_user.get() or "").strip()
        p = (self.e_pwd.get() or "").strip()
        if not u or not p:
            messagebox.showerror("Login", "Please enter username and password.")
            return
        row = get_user_by_username(u)
        if not row:
            messagebox.showerror("Login", "Invalid credentials.")
            return
        uid, username, role, pwdhash, is_active, must_change = row
        if not is_active:
            messagebox.showerror("Login", "Account is disabled.")
            return
        if not _verify_pwd(p, pwdhash):
            messagebox.showerror("Login", "Invalid credentials.")
            return
        self.result = {"id": uid, "username": username, "role": role}
        self.destroy()


# ---------- Helpers for CTk buttons (white, black outline) ----------
def white_btn(parent, **kwargs):
    kwargs.setdefault("fg_color", "white")
    kwargs.setdefault("text_color", "black")
    kwargs.setdefault("border_color", "black")
    kwargs.setdefault("border_width", 1)
    kwargs.setdefault("hover_color", "#F0F0F0")
    return ctk.CTkButton(parent, **kwargs)


# ---------- Voucher App ----------
STATUS_VALUES = ["All", "Pending", "Completed", "Deleted", "1st call", "2nd reminder", "3rd reminder"]

class VoucherApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Service Voucher Management System")
        self.geometry("1280x780")
        self.minsize(1024, 640)
        ctk.set_appearance_mode("light")

        # Global font size to 14
        try:
            import tkinter.font as tkfont
            for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkTooltipFont"):
                f = tkfont.nametofont(name)
                f.configure(size=UI_FONT_SIZE)
        except Exception:
            pass

        # Login
        self.current_user = None
        self._do_login_flow()

        # Menu bar (top)
        self._build_menubar()

        # Root layout: top filter strip is horizontally scrollable to keep controls reachable
        root = ctk.CTkFrame(self)
        root.pack(fill="both", expand=True)

        self._build_filters(root)
        self._build_table(root)
        self._build_bottom_bar(root)

        self.after(80, self._go_fullscreen)
        self.perform_search()

    # ---------- Login & role ----------
    def _do_login_flow(self):
        dlg = LoginDialog(self)
        self.wait_window(dlg)
        if not dlg.result:
            self.destroy()
            sys.exit(0)
        self.current_user = dlg.result
        self.title(f"SVMS — logged in as {self.current_user['username']} ({self.current_user['role']})")

    def _role_is(self, *roles):
        return self.current_user and self.current_user["role"] in roles

    # ---------- UI building ----------
    def _build_menubar(self):
        self.menu = tk.Menu(self)
        self.config(menu=self.menu)

        self.menu_svms = tk.Menu(self.menu, tearoff=0)
        self.menu_svms.add_command(label="Open PDF (Selected)", command=self.open_pdf)
        self.menu_svms.add_command(label="Open PDF Folder", command=self.open_pdf_folder)
        self.menu_svms.add_command(label="Regenerate PDF (Selected)", command=self.regen_pdf_selected)
        self.menu_svms.add_separator()
        self.menu_svms.add_command(label="Export to Excel", command=self.export_excel)
        self.menu_svms.add_separator()
        self.menu_svms.add_command(label="Backup Data (.zip)", command=self.backup_all)
        self.menu_svms.add_command(label="Restore Data (.zip)", command=self.restore_all)
        self.menu_svms.add_separator()
        self.menu_svms.add_command(label="Exit", command=self.destroy)

        self.menu.add_cascade(label="SVMS Menu", menu=self.menu_svms)

        self.menu_user = tk.Menu(self.menu, tearoff=0)
        self.menu_user.add_command(label="Manage Users", command=self.manage_users, state=tk.NORMAL if self._role_is("admin") else tk.DISABLED)
        self.menu.add_cascade(label="User Profile", menu=self.menu_user)

        self.menu_staff = tk.Menu(self.menu, tearoff=0)
        self.menu_staff.add_command(label="Staff Profile", command=self.staff_profile)
        self.menu.add_cascade(label="Staff Profile", menu=self.menu_staff)

        self.menu_comm = tk.Menu(self.menu, tearoff=0)
        self.menu_comm.add_command(label="Add Commission", command=self.add_commission)
        self.menu_comm.add_command(label="View/Edit Commissions", command=self.view_commissions)
        self.menu.add_cascade(label="Sales Commission", menu=self.menu_comm)

    def _build_filters(self, parent):
        # Wrap in a canvas for horizontal scrollability
        wrap = ctk.CTkFrame(parent)
        wrap.pack(fill="x", padx=8, pady=(8, 6))

        self.filter_canvas = tk.Canvas(wrap, height=52, borderwidth=0, highlightthickness=0)
        hscroll = ttk.Scrollbar(wrap, orient="horizontal", command=self.filter_canvas.xview)
        self.filter_canvas.configure(xscrollcommand=hscroll.set)
        self.filter_canvas.pack(fill="x", side="top")
        hscroll.pack(fill="x", side="bottom")

        self.filter_inner = ctk.CTkFrame(self.filter_canvas)
        self.filter_canvas.create_window((0, 0), window=self.filter_inner, anchor="nw")

        # Inputs
        today_ui = _to_ui_date(datetime.now())

        self.f_voucher = ctk.CTkEntry(self.filter_inner, width=140, placeholder_text="VoucherID");    self.f_voucher.grid(row=0, column=0, padx=5, pady=4)
        self.f_name    = ctk.CTkEntry(self.filter_inner, width=230, placeholder_text="Customer Name"); self.f_name.grid(row=0, column=1, padx=5, pady=4)
        self.f_contact = ctk.CTkEntry(self.filter_inner, width=190, placeholder_text="Contact Number"); self.f_contact.grid(row=0, column=2, padx=5, pady=4)
        self.f_from    = ctk.CTkEntry(self.filter_inner, width=180, placeholder_text="Date From (DD-MM-YYYY)"); self.f_from.grid(row=0, column=3, padx=5, pady=4)
        self.f_to      = ctk.CTkEntry(self.filter_inner, width=180, placeholder_text="Date To (DD-MM-YYYY)");   self.f_to.grid(row=0, column=4, padx=5, pady=4)
        self.f_to.insert(0, today_ui)

        self.f_status  = ctk.CTkOptionMenu(self.filter_inner, values=STATUS_VALUES, width=140); self.f_status.grid(row=0, column=5, padx=5, pady=4)
        self.f_status.set("All")

        self.btn_search = white_btn(self.filter_inner, text="Search", command=self.perform_search, width=110); self.btn_search.grid(row=0, column=6, padx=5, pady=4)
        self.btn_reset  = white_btn(self.filter_inner, text="Reset",  command=self.reset_filters, width=100);  self.btn_reset.grid(row=0, column=7, padx=5, pady=4)

        self.filter_inner.update_idletasks()
        self.filter_canvas.configure(scrollregion=self.filter_canvas.bbox("all"))

        def _resize_filter(_evt=None):
            self.filter_canvas.configure(scrollregion=self.filter_canvas.bbox("all"))
        self.filter_inner.bind("<Configure>", _resize_filter)

    def _build_table(self, parent):
        table_frame = ctk.CTkFrame(parent)
        table_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_frame,
            columns=("VoucherID","Date","Customer","Contact","Units","Recipient","TechID","TechName","Status","Solution","PDF"),
            show="headings", selectmode="extended")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")

        # Headings / widths
        for col, text, w in [
            ("VoucherID","VoucherID",100), ("Date","Date",170), ("Customer","Customer",200),
            ("Contact","Contact",140), ("Units","Units",60), ("Recipient","Recipient",160),
            ("TechID","Technician ID",120), ("TechName","Technician Name",180),
            ("Status","Status",120), ("Solution","Solution",260)
        ]:
            self.tree.heading(col, text=text)
            self.tree.column(col, anchor="w", width=w, stretch=True)
        self.tree.heading("PDF", text="PDF"); self.tree.column("PDF", width=0, stretch=False)

        # Row colors / tags
        # base statuses
        self.tree.tag_configure("Pending", background="#FFF4B3")      # yellow
        self.tree.tag_configure("Completed", background="#CDEEC8")    # green
        self.tree.tag_configure("Deleted", background="#F8D7DA")      # red
        self.tree.tag_configure("1st call", background="#CCE5FF")     # light blue
        self.tree.tag_configure("2nd reminder", background="#D9CCE5") # lavender
        self.tree.tag_configure("3rd reminder", background="#FFD9B3") # peach
        # outstanding tiers ( >60 days )
        self.tree.tag_configure("out_7",  background="#FFF0F0")       # subtle pink
        self.tree.tag_configure("out_14", background="#FFD6D6")       # stronger pink
        self.tree.tag_configure("out_30", background="#FFB3B3")       # strong red-pink

        self.tree.bind("<Double-1>", lambda e: self.open_pdf())

        # Context menu
        self._make_context_menu(self.tree)

    def _build_bottom_bar(self, parent):
        bar = ctk.CTkFrame(parent)
        bar.pack(fill="x", padx=8, pady=(0, 10))

        # left side action buttons
        b1 = white_btn(bar, text="Add Voucher", command=self.add_voucher_ui, width=140)
        b1.pack(side="left", padx=5, pady=8)

        b_pdf = white_btn(bar, text="Open PDF", command=self.open_pdf, width=120)
        b_pdf.pack(side="left", padx=5, pady=8)

        b_folder = white_btn(bar, text="Open PDF Folder", command=self.open_pdf_folder, width=150)
        b_folder.pack(side="left", padx=5, pady=8)

        b_manage_rec = white_btn(bar, text="Manage Recipients", command=self.manage_staffs_ui, width=170)
        b_manage_rec.pack(side="left", padx=5, pady=8)
        if not self._role_is("admin","supervisor"):
            b_manage_rec.configure(state=tk.DISABLED)

        b_base = white_btn(bar, text="Modify Base VID", command=self.modify_base_vid_ui, width=160)
        b_base.pack(side="left", padx=5, pady=8)
        if not self._role_is("admin"):
            b_base.configure(state=tk.DISABLED)

    # ---------- Context Menu ----------
    def _make_context_menu(self, tree):
        self.ctx = tk.Menu(tree, tearoff=0)
        self.ctx.add_command(label="Edit", command=self.edit_selected)
        mark = tk.Menu(self.ctx, tearoff=0)
        for s in ["Pending","Completed","Deleted","1st call","2nd reminder","3rd reminder"]:
            mark.add_command(label=s, command=lambda st=s: self._bulk_mark(st))
        self.ctx.add_cascade(label="Mark", menu=mark)
        self.ctx.add_separator()
        self.ctx.add_command(label="Open PDF", command=self.open_pdf)
        self.ctx.add_command(label="Regenerate PDF", command=self.regen_pdf_selected)
        self.ctx.add_separator()
        self.ctx.add_command(label="Unmark Completed → Pending", command=lambda: self._bulk_mark("Pending"))
        # Bind right-click
        def _popup(event):
            try:
                row = self.tree.identify_row(event.y)
                if row:
                    if row not in self.tree.selection():
                        self.tree.selection_set(row)
                self.ctx.post(event.x_root, event.y_root)
            finally:
                self.ctx.grab_release()
        self.tree.bind("<Button-3>", _popup)

    # ---------- Filters ----------
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
        self.f_to.insert(0, _to_ui_date(datetime.now()))
        self.f_status.set("All")
        self.perform_search()

    # ---------- Data ops ----------
    def perform_search(self):
        rows = search_vouchers(self._get_filters())
        self.tree.delete(*self.tree.get_children())
        now = datetime.now()
        for (vid, created_at, customer, contact, units, recipient, tech_id, tech_name, status, solution, pdf) in rows:
            tags = [status]
            # outstanding tiers: if not completed/deleted and older than 60 days
            try:
                dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
            except Exception:
                dt = now
            days = (now - dt).days
            if status not in ("Completed","Deleted") and days > 60:
                # 60–67, 68–74, >=75
                if days <= 67:
                    tags.append("out_7")
                elif days <= 74:
                    tags.append("out_14")
                else:
                    tags.append("out_30")
            self.tree.insert("", "end", values=(
                vid, _to_ui_datetime_str(created_at), customer, contact, units,
                recipient or "", tech_id or "", tech_name or "", status, solution or "", pdf or ""
            ), tags=tuple(tags))
        self.tree.update_idletasks()

    # ---------- Selection helpers ----------
    def _selected_ids(self):
        sels = self.tree.selection()
        vids = []
        for iid in sels:
            vals = self.tree.item(iid)["values"]
            if vals:
                vids.append(str(vals[0]))
        return vids

    # ---------- Actions ----------
    def _bulk_mark(self, new_status):
        vids = self._selected_ids()
        if not vids:
            messagebox.showerror("Mark", "Select record(s) first.")
            return
        if new_status == "Deleted" and not self._role_is("admin","supervisor"):
            messagebox.showerror("Mark", "You do not have permission to mark as Deleted.")
            return
        bulk_update_status(vids, new_status)
        self.perform_search()

    def open_pdf(self):
        sels = self.tree.selection()
        if not sels:
            messagebox.showerror("Open PDF", "Select a record first.")
            return
        vals = self.tree.item(sels[0])["values"]
        if not vals: return
        pdf_path = vals[-1]
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Open PDF", "PDF not found for this voucher.")
            return
        try:
            webbrowser.open_new(os.path.abspath(pdf_path))
        except Exception:
            if os.name == "nt":
                os.startfile(pdf_path)  # type: ignore
            else:
                os.system(f"open '{pdf_path}'" if sys.platform == "darwin" else f"xdg-open '{pdf_path}'")

    def regen_pdf_selected(self):
        sels = self.tree.selection()
        if not sels:
            messagebox.showerror("Regenerate", "Select a record first.")
            return
        vid = str(self.tree.item(sels[0])["values"][0])

        conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
        cur.execute("""SELECT voucher_id, created_at, customer_name, contact_number, units,
                              particulars, problem, staff_name, status, recipient, solution, pdf_path
                       FROM vouchers WHERE voucher_id=?""", (vid,))
        row = cur.fetchone(); conn.close()
        if not row:
            messagebox.showerror("Regenerate", "Voucher not found.")
            return
        (voucher_id, created_at, customer_name, contact_number, units, particulars,
         problem, staff_name, status, recipient, solution, old_pdf) = row
        try:
            if old_pdf and os.path.exists(old_pdf):
                os.remove(old_pdf)
        except Exception:
            pass
        new_pdf = generate_pdf(voucher_id, customer_name, contact_number, units, particulars, problem, staff_name, status, created_at, recipient)
        update_voucher_fields(voucher_id, pdf_path=new_pdf)
        self.perform_search()
        messagebox.showinfo("Regenerate", f"PDF regenerated for voucher {voucher_id}.")

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

    # ---------- Add / Edit voucher ----------
    def add_voucher_ui(self):
        top = ctk.CTkToplevel(self)
        top.title("Create Voucher")
        top.geometry("980x780")
        top.grab_set()
        frm = ctk.CTkFrame(top); frm.pack(fill="both", expand=True, padx=12, pady=12)
        frm.grid_columnconfigure(1, weight=1)

        WIDE = 560
        r = 0
        ctk.CTkLabel(frm, text="Customer Name").grid(row=r, column=0, sticky="w")
        e_name = ctk.CTkEntry(frm, width=WIDE); e_name.grid(row=r, column=1, sticky="ew", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Contact Number").grid(row=r, column=0, sticky="w")
        e_contact = ctk.CTkEntry(frm, width=WIDE); e_contact.grid(row=r, column=1, sticky="ew", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="No. of Units").grid(row=r, column=0, sticky="w")
        e_units = ctk.CTkEntry(frm, width=120); e_units.insert(0,"1")
        e_units.grid(row=r, column=1, sticky="w", padx=10, pady=6); r+=1

        def mk_text(parent, height):
            wrap = ctk.CTkFrame(parent)
            wrap.grid_columnconfigure(0, weight=1)
            txt = tk.Text(wrap, height=height, font=(FONT_FAMILY, UI_FONT_SIZE))
            sb  = ttk.Scrollbar(wrap, orient="vertical", command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            txt.grid(row=0, column=0, sticky="nsew")
            sb.grid(row=0, column=1, sticky="ns")
            return wrap, txt

        ctk.CTkLabel(frm, text="Particulars").grid(row=r, column=0, sticky="w")
        part_wrap, t_part = mk_text(frm, height=6); part_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Problem").grid(row=r, column=0, sticky="w")
        prob_wrap, t_prob = mk_text(frm, height=5); prob_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Recipient").grid(row=r, column=0, sticky="w")
        staff_values = list_staffs_names() or [""]
        e_recipient = ctk.CTkComboBox(frm, values=staff_values, width=WIDE)
        if staff_values and staff_values[0] != "": e_recipient.set(staff_values[0])
        e_recipient.grid(row=r, column=1, sticky="w", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Technician ID").grid(row=r, column=0, sticky="w")
        e_tid = ctk.CTkEntry(frm, width=240); e_tid.grid(row=r, column=1, sticky="w", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Technician Name").grid(row=r, column=0, sticky="w")
        e_tname = ctk.CTkEntry(frm, width=WIDE); e_tname.grid(row=r, column=1, sticky="ew", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Solution").grid(row=r, column=0, sticky="w")
        sol_wrap, t_sol = mk_text(frm, height=4); sol_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6); r+=1

        btns = ctk.CTkFrame(top); btns.pack(fill="x", padx=12, pady=(0,12))
        def save():
            name = e_name.get().strip()
            contact = e_contact.get().strip()
            try:
                units = int((e_units.get() or "1").strip())
                if units <= 0: raise ValueError
            except ValueError:
                messagebox.showerror("Invalid", "Units must be a positive integer."); return
            particulars = t_part.get("1.0","end").strip()
            problem     = t_prob.get("1.0","end").strip()
            recipient   = e_recipient.get().strip()
            staff_name  = recipient
            solution    = t_sol.get("1.0","end").strip()
            technician_id   = e_tid.get().strip()
            technician_name = e_tname.get().strip()

            if not name or not contact:
                messagebox.showerror("Missing", "Customer name and contact are required."); return

            voucher_id, pdf_path = add_voucher(name, contact, units, particulars, problem, staff_name,
                                               recipient, solution, technician_id, technician_name)
            # Open PDF but keep user on SV page (dialog will close only after OK)
            try: webbrowser.open_new(os.path.abspath(pdf_path))
            except Exception: pass
            messagebox.showinfo("Saved", f"Voucher {voucher_id} created.\nPDF opened.")
            top.destroy(); self.perform_search()

        white_btn(btns, text="Save & Open PDF", command=save, width=200).pack(side="right")

    def edit_selected(self):
        sels = self.tree.selection()
        if len(sels) != 1:
            messagebox.showerror("Edit", "Select exactly one record to edit.")
            return
        voucher_id = self.tree.item(sels[0])["values"][0]

        conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
        cur.execute("""SELECT voucher_id, created_at, customer_name, contact_number, units,
                              particulars, problem, staff_name, status, recipient, solution,
                              technician_id, technician_name
                       FROM vouchers WHERE voucher_id=?""", (voucher_id,))
        row = cur.fetchone(); conn.close()
        if not row:
            messagebox.showerror("Edit", "Voucher not found."); return

        (_, created_at, customer_name, contact_number, units, particulars, problem, staff_name,
         status, recipient, solution, tech_id, tech_name) = row

        top = ctk.CTkToplevel(self)
        top.title(f"Edit Voucher {voucher_id}")
        top.geometry("980x780")
        top.grab_set()
        frm = ctk.CTkFrame(top); frm.pack(fill="both", expand=True, padx=12, pady=12)
        frm.grid_columnconfigure(1, weight=1)

        WIDE = 560
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

        def mk_text(parent, height, seed=""):
            wrap = ctk.CTkFrame(parent)
            wrap.grid_columnconfigure(0, weight=1)
            txt = tk.Text(wrap, height=height, font=(FONT_FAMILY, UI_FONT_SIZE))
            txt.insert("1.0", seed)
            sb  = ttk.Scrollbar(wrap, orient="vertical", command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            txt.grid(row=0, column=0, sticky="nsew")
            sb.grid(row=0, column=1, sticky="ns")
            return wrap, txt

        ctk.CTkLabel(frm, text="Particulars").grid(row=r, column=0, sticky="w")
        part_wrap, t_part = mk_text(frm, 6, particulars or ""); part_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Problem").grid(row=r, column=0, sticky="w")
        prob_wrap, t_prob = mk_text(frm, 5, problem or ""); prob_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Recipient").grid(row=r, column=0, sticky="w")
        staff_values = list_staffs_names() or [""]
        e_recipient = ctk.CTkComboBox(frm, values=staff_values, width=WIDE)
        e_recipient.set(recipient or staff_name or "")
        e_recipient.grid(row=r, column=1, sticky="w", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Technician ID").grid(row=r, column=0, sticky="w")
        e_tid = ctk.CTkEntry(frm, width=240); e_tid.insert(0, tech_id or "")
        e_tid.grid(row=r, column=1, sticky="w", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Technician Name").grid(row=r, column=0, sticky="w")
        e_tname = ctk.CTkEntry(frm, width=WIDE); e_tname.insert(0, tech_name or "")
        e_tname.grid(row=r, column=1, sticky="ew", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Solution").grid(row=r, column=0, sticky="w")
        sol_wrap, t_sol = mk_text(frm, 4, solution or ""); sol_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6); r+=1

        btns = ctk.CTkFrame(top); btns.pack(fill="x", padx=12, pady=12)
        def save_edit():
            name = e_name.get().strip()
            contact = e_contact.get().strip()
            try:
                units_val = int((e_units.get() or "1").strip())
                if units_val <= 0: raise ValueError
            except ValueError:
                messagebox.showerror("Invalid", "Units must be a positive integer."); return
            particulars_val = t_part.get("1.0","end").strip()
            problem_val     = t_prob.get("1.0","end").strip()
            recipient_val   = e_recipient.get().strip()
            solution_val    = t_sol.get("1.0","end").strip()
            staff_val       = recipient_val
            tech_id_val     = e_tid.get().strip()
            tech_name_val   = e_tname.get().strip()

            update_voucher_fields(voucher_id,
                                  customer_name=name, contact_number=contact, units=units_val,
                                  particulars=particulars_val, problem=problem_val,
                                  staff_name=staff_val, recipient=recipient_val,
                                  solution=solution_val, technician_id=tech_id_val,
                                  technician_name=tech_name_val)
            # regen PDF for consistency
            conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
            cur.execute("SELECT created_at, status, pdf_path FROM vouchers WHERE voucher_id=?", (voucher_id,))
            created_at, status, old_pdf_path = cur.fetchone()
            try:
                if old_pdf_path and os.path.exists(old_pdf_path): os.remove(old_pdf_path)
            except Exception:
                pass
            pdf_path = generate_pdf(str(voucher_id), name, contact, units_val,
                                    particulars_val, problem_val, staff_val, status, created_at, recipient_val)
            cur.execute("UPDATE vouchers SET pdf_path=? WHERE voucher_id=?", (pdf_path, str(voucher_id)))
            conn.commit(); conn.close()

            messagebox.showinfo("Updated", f"Voucher {voucher_id} updated.")
            top.destroy(); self.perform_search()

        white_btn(btns, text="Save Changes", command=save_edit, width=160).pack(side="right", padx=6)
        white_btn(btns, text="Cancel", command=top.destroy, width=100).pack(side="right", padx=6)

    # ---------- Manage Recipient (simple) ----------
    def manage_staffs_ui(self):
        if not self._role_is("admin","supervisor"):
            messagebox.showerror("Permission", "Not allowed."); return

        top = ctk.CTkToplevel(self)
        top.title("Manage Recipients")
        top.geometry("680x480")
        top.grab_set()

        root = ctk.CTkFrame(top); root.pack(fill="both", expand=True, padx=14, pady=14)
        root.grid_rowconfigure(2, weight=1)
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=0)

        entry = ctk.CTkEntry(root, placeholder_text="New recipient name")
        entry.grid(row=0, column=0, sticky="ew", padx=(0,10), pady=(0,10))

        row1 = ctk.CTkFrame(root); row1.grid(row=1, column=0, sticky="w", padx=(0,10), pady=(0,10))
        add_btn = white_btn(row1, text="Add", width=120)
        del_btn = white_btn(row1, text="Delete Selected", width=160)
        add_btn.pack(side="left", padx=(0,10)); del_btn.pack(side="left")

        list_frame = ctk.CTkFrame(root); list_frame.grid(row=2, column=0, sticky="nsew", padx=(0,10))
        list_frame.grid_rowconfigure(0, weight=1); list_frame.grid_columnconfigure(0, weight=1)
        lb = tk.Listbox(list_frame, height=14, font=(FONT_FAMILY, UI_FONT_SIZE))
        lb.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=lb.yview)
        sb.grid(row=0, column=1, sticky="ns"); lb.configure(yscrollcommand=sb.set)

        close_btn = white_btn(root, text="Close", width=120, command=top.destroy)
        close_btn.grid(row=3, column=1, sticky="e", pady=(10,0))

        def refresh_list():
            lb.delete(0, "end")
            for s in list_staffs_names():
                lb.insert("end", s)

        def do_add():
            if add_staff_simple(entry.get()):
                entry.delete(0, "end")
                refresh_list()

        def do_del():
            sel = lb.curselection()
            if not sel: return
            name = lb.get(sel[0])
            delete_staff_simple(name)
            refresh_list()

        add_btn.configure(command=do_add)
        del_btn.configure(command=do_del)
        refresh_list()

    # ---------- Modify Base VID ----------
    def modify_base_vid_ui(self):
        if not self._role_is("admin"):
            messagebox.showerror("Permission", "Admin only."); return

        top = ctk.CTkToplevel(self)
        top.title("Modify Base Voucher ID")
        top.geometry("420x220")
        top.grab_set()

        frm = ctk.CTkFrame(top); frm.pack(fill="both", expand=True, padx=16, pady=16)
        current_base = _read_base_vid()
        ctk.CTkLabel(frm, text=f"Current Base VID: {current_base}", anchor="w").pack(fill="x", pady=(0,8))
        entry = ctk.CTkEntry(frm, placeholder_text="Enter new base voucher ID (integer)")
        entry.pack(fill="x"); entry.insert(0, str(current_base))
        info = ctk.CTkLabel(frm, text="Existing vouchers will shift by the difference. PDFs will be regenerated.",
                            justify="left", wraplength=360)
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
        white_btn(btns, text="Apply", command=apply, width=120).pack(side="right", padx=4)
        white_btn(btns, text="Cancel", command=top.destroy, width=100).pack(side="right", padx=4)

    # ---------- Export Excel ----------
    def export_excel(self):
        if not _HAS_PANDAS:
            messagebox.showerror("Excel Export", "Excel export requires 'pandas' and 'openpyxl'.\nInstall: pip install pandas openpyxl")
            return
        rows = search_vouchers(self._get_filters())
        if not rows:
            messagebox.showinfo("Export", "No rows to export."); return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel Workbook", "*.xlsx")],
                                            title="Save Excel")
        if not path: return
        try:
            df = pd.DataFrame([{
                "VoucherID": r[0], "CreatedAt": r[1], "Customer": r[2], "Contact": r[3], "Units": r[4],
                "Recipient": r[5], "TechnicianID": r[6], "TechnicianName": r[7],
                "Status": r[8], "Solution": r[9], "PDFPath": r[10]
            } for r in rows])
            df.to_excel(path, index=False, engine="openpyxl")
            messagebox.showinfo("Export", f"Excel exported:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export Excel:\n{e}")

    # ---------- Backup / Restore ----------
    def backup_all(self):
        path = filedialog.asksaveasfilename(defaultextension=".zip",
                                            filetypes=[("Zip Archive", "*.zip")],
                                            title="Create Backup (.zip)")
        if not path: return
        try:
            with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
                # DB
                if os.path.exists(DB_FILE):
                    z.write(DB_FILE, arcname="vouchers.db")
                # PDFs
                for root, _, files in os.walk(PDF_DIR):
                    for f in files:
                        fp = os.path.join(root, f)
                        arc = os.path.relpath(fp, APP_DIR)
                        z.write(fp, arcname=arc)
                # Images
                for root, _, files in os.walk(IMG_DIR):
                    for f in files:
                        fp = os.path.join(root, f)
                        arc = os.path.relpath(fp, APP_DIR)
                        z.write(fp, arcname=arc)
            messagebox.showinfo("Backup", f"Backup created:\n{path}")
        except Exception as e:
            messagebox.showerror("Backup Error", f"Failed to create backup:\n{e}")

    def restore_all(self):
        path = filedialog.askopenfilename(filetypes=[("Zip Archive", "*.zip")], title="Select Backup Zip")
        if not path: return
        if not messagebox.askyesno("Restore", "Restoring will overwrite current data. Continue?"):
            return
        try:
            with zipfile.ZipFile(path, "r") as z:
                # Extract to a temp and then copy to target to avoid partial restore
                temp_dir = os.path.join(APP_DIR, "_restore_tmp")
                if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
                z.extractall(temp_dir)
                # Move DB
                src_db = os.path.join(temp_dir, "vouchers.db")
                if os.path.exists(src_db):
                    shutil.copy2(src_db, DB_FILE)
                # Copy PDFs and images, preserving structure
                for sub in ("pdfs", "images"):
                    src = os.path.join(temp_dir, sub)
                    if os.path.exists(src):
                        dst = os.path.join(APP_DIR, sub)
                        if os.path.exists(dst):
                            shutil.rmtree(dst)
                        shutil.copytree(src, dst)
                shutil.rmtree(temp_dir, ignore_errors=True)
            messagebox.showinfo("Restore", "Restore completed. Please restart the application.")
        except Exception as e:
            messagebox.showerror("Restore Error", f"Failed to restore backup:\n{e}")

    # ---------- Users (Admin-only UI) ----------
    def manage_users(self):
        if not self._role_is("admin"):
            messagebox.showerror("Permission", "Admin only."); return

        top = ctk.CTkToplevel(self)
        top.title("User Accounts")
        top.geometry("720x520"); top.grab_set()

        frm = ctk.CTkFrame(top); frm.pack(fill="both", expand=True, padx=10, pady=10)
        frm.grid_columnconfigure(0, weight=1)
        frm.grid_rowconfigure(1, weight=1)

        # create panel
        panel = ctk.CTkFrame(frm); panel.grid(row=0, column=0, sticky="ew", pady=(0,8))
        ctk.CTkLabel(panel, text="Username").grid(row=0, column=0, padx=6)
        e_user = ctk.CTkEntry(panel, width=180); e_user.grid(row=0, column=1, padx=6)
        ctk.CTkLabel(panel, text="Role").grid(row=0, column=2, padx=6)
        role_box = ctk.CTkComboBox(panel, values=["supervisor","user"], width=160); role_box.set("user")
        role_box.grid(row=0, column=3, padx=6)
        ctk.CTkLabel(panel, text="Password").grid(row=0, column=4, padx=6)
        e_pwd = ctk.CTkEntry(panel, width=180, show="•"); e_pwd.grid(row=0, column=5, padx=6)
        white_btn(panel, text="Add User", width=120,
                  command=lambda: self._add_user_and_refresh(e_user, role_box, e_pwd, tree)).grid(row=0, column=6, padx=6)

        # table
        cols = ("id","username","role","active","must_change")
        tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="browse")
        for c, t, w in [("id","ID",60), ("username","Username",200), ("role","Role",120),
                        ("active","Active",80), ("must_change","Must Change",120)]:
            tree.heading(c, text=t); tree.column(c, width=w, anchor="w")
        tree.grid(row=1, column=0, sticky="nsew")
        sb = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
        sb.grid(row=1, column=1, sticky="ns"); tree.configure(yscrollcommand=sb.set)

        # buttons
        actions = ctk.CTkFrame(frm); actions.grid(row=2, column=0, sticky="e", pady=(8,0))
        white_btn(actions, text="Reset Password", width=140,
                  command=lambda: self._reset_user_password(tree)).pack(side="left", padx=5)
        white_btn(actions, text="Toggle Active", width=140,
                  command=lambda: self._toggle_user_active(tree)).pack(side="left", padx=5)
        white_btn(actions, text="Delete User", width=130,
                  command=lambda: self._delete_user(tree)).pack(side="left", padx=5)

        # load
        self._refresh_users(tree)

    def _refresh_users(self, tree):
        tree.delete(*tree.get_children())
        for (uid, username, role, is_active, must_change) in list_users():
            tree.insert("", "end", values=(uid, username, role, "Yes" if is_active else "No", "Yes" if must_change else "No"))

    def _add_user_and_refresh(self, e_user, role_box, e_pwd, tree):
        u = (e_user.get() or "").strip()
        r = role_box.get()
        p = (e_pwd.get() or "").strip()
        if not u or not p:
            messagebox.showerror("User", "Username and password required."); return
        try:
            create_user(u, r, p)
            self._refresh_users(tree)
            e_user.delete(0, "end"); e_pwd.delete(0, "end"); role_box.set("user")
        except Exception as e:
            messagebox.showerror("User", f"Failed: {e}")

    def _reset_user_password(self, tree):
        sel = tree.selection()
        if not sel: return
        vals = tree.item(sel[0])["values"]
        uid, username = vals[0], vals[1]
        new = tk.simpledialog.askstring("Reset Password", f"Enter new password for {username}:", show="•")
        if not new: return
        try:
            reset_password(uid, new)
            self._refresh_users(tree)
            messagebox.showinfo("User", "Password reset; user must change on next login.")
        except Exception as e:
            messagebox.showerror("User", f"Failed: {e}")

    def _toggle_user_active(self, tree):
        sel = tree.selection()
        if not sel: return
        vals = tree.item(sel[0])["values"]
        uid, active = vals[0], vals[3] == "Yes"
        try:
            update_user(uid, is_active=0 if active else 1)
            self._refresh_users(tree)
        except Exception as e:
            messagebox.showerror("User", f"Failed: {e}")

    def _delete_user(self, tree):
        sel = tree.selection()
        if not sel: return
        vals = tree.item(sel[0])["values"]
        uid, username, role = vals[0], vals[1], vals[2]
        if role == "admin":
            messagebox.showerror("User", "Cannot delete admin."); return
        if not messagebox.askyesno("Delete User", f"Delete user '{username}'?"): return
        try:
            delete_user(uid)
            self._refresh_users(tree)
        except Exception as e:
            messagebox.showerror("User", f"Failed: {e}")

    # ---------- Staff Profile UI (basic CRUD) ----------
    def staff_profile(self):
        top = ctk.CTkToplevel(self)
        top.title("Staff Profile")
        top.geometry("900x600")
        top.grab_set()

        frm = ctk.CTkFrame(top); frm.pack(fill="both", expand=True, padx=10, pady=10)
        frm.grid_columnconfigure(1, weight=1)
        frm.grid_rowconfigure(4, weight=1)

        # form
        ctk.CTkLabel(frm, text="Position").grid(row=0, column=0, sticky="w")
        roles = ["Salesman","Technician","Boss"]
        cb_pos = ctk.CTkComboBox(frm, values=roles, width=200); cb_pos.set("Technician")
        cb_pos.grid(row=0, column=1, sticky="w", padx=8, pady=4)

        ctk.CTkLabel(frm, text="Staff/Technician ID (optional)").grid(row=0, column=2, sticky="w")
        e_sid = ctk.CTkEntry(frm, width=200); e_sid.grid(row=0, column=3, sticky="w", padx=8, pady=4)

        ctk.CTkLabel(frm, text="Name").grid(row=1, column=0, sticky="w")
        e_name = ctk.CTkEntry(frm, width=260); e_name.grid(row=1, column=1, sticky="w", padx=8, pady=4)

        ctk.CTkLabel(frm, text="Contact Number").grid(row=1, column=2, sticky="w")
        e_phone = ctk.CTkEntry(frm, width=200); e_phone.grid(row=1, column=3, sticky="w", padx=8, pady=4)

        photo_lbl = ctk.CTkLabel(frm, text="No photo selected")
        photo_lbl.grid(row=2, column=0, columnspan=4, sticky="w", padx=8, pady=(6,0))

        def choose_photo():
            p = filedialog.askopenfilename(filetypes=[("Image", "*.jpg;*.jpeg;*.png")])
            if p:
                photo_lbl.configure(text=p)
        white_btn(frm, text="Choose Photo", command=choose_photo, width=140).grid(row=2, column=3, sticky="e", pady=(6,0))

        def _insert_staff():
            name = e_name.get().strip()
            if not name:
                messagebox.showerror("Staff", "Name required."); return
            pos = cb_pos.get()
            sid = e_sid.get().strip()
            phone = e_phone.get().strip()
            photo_src = photo_lbl.cget("text") if photo_lbl.cget("text") != "No photo selected" else ""
            photo_path = ""
            if photo_src:
                os.makedirs(IMG_STAFF, exist_ok=True)
                fn = f"{name.replace(' ','_')}_{int(datetime.now().timestamp())}.jpg"
                photo_path = os.path.join(IMG_STAFF, fn)
                try:
                    _process_square_image(photo_src, photo_path, max_px=400)
                except Exception as e:
                    messagebox.showerror("Photo", f"Failed to process photo: {e}")
                    photo_path = ""
            conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
            cur.execute("""INSERT INTO staffs (position, staff_id_opt, name, phone, photo_path, created_at, updated_at)
                           VALUES (?,?,?,?,?,?,?)""",
                        (pos, sid, name, phone, photo_path,
                         datetime.now().isoformat(sep=" ", timespec="seconds"),
                         datetime.now().isoformat(sep=" ", timespec="seconds")))
            conn.commit(); conn.close()
            refresh()
            e_name.delete(0,"end"); e_phone.delete(0,"end"); e_sid.delete(0,"end"); photo_lbl.configure(text="No photo selected")

        white_btn(frm, text="Add Staff", command=_insert_staff, width=130).grid(row=3, column=3, sticky="e", pady=6)

        # table
        cols = ("id","position","staff_id","name","phone","photo")
        tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="browse")
        for c, t, w in [("id","ID",60), ("position","Position",120), ("staff_id","StaffID",120),
                        ("name","Name",200), ("phone","Phone",160), ("photo","Photo",260)]:
            tree.heading(c, text=t); tree.column(c, width=w, anchor="w")
        tree.grid(row=4, column=0, columnspan=4, sticky="nsew")
        sb = ttk.Scrollbar(frm, orient="vertical", command=tree.yview); sb.grid(row=4, column=4, sticky="ns")
        tree.configure(yscrollcommand=sb.set)

        def refresh():
            tree.delete(*tree.get_children())
            conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
            cur.execute("SELECT id, position, staff_id_opt, name, phone, photo_path FROM staffs ORDER BY name")
            for r in cur.fetchall():
                tree.insert("", "end", values=r)
            conn.close()
        refresh()

        def delete_sel():
            sel = tree.selection()
            if not sel: return
            sid = tree.item(sel[0])["values"][0]
            if not messagebox.askyesno("Delete", "Delete selected staff?"): return
            conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
            cur.execute("DELETE FROM staffs WHERE id=?", (sid,))
            conn.commit(); conn.close()
            refresh()
        white_btn(frm, text="Delete Selected", command=delete_sel, width=150).grid(row=5, column=3, sticky="e", pady=8)

    # ---------- Commission UI ----------
    def add_commission(self):
        top = ctk.CTkToplevel(self); top.title("Add Commission"); top.geometry("760x540"); top.grab_set()
        frm = ctk.CTkFrame(top); frm.pack(fill="both", expand=True, padx=10, pady=10)
        frm.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(frm, text="Staff").grid(row=0, column=0, sticky="w")
        conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
        cur.execute("SELECT id, name FROM staffs ORDER BY name"); staff_rows = cur.fetchall(); conn.close()
        staff_names = [f"{i}:{n}" for (i,n) in staff_rows] or [""]
        cb_staff = ctk.CTkComboBox(frm, values=staff_names, width=320); 
        if staff_names and staff_names[0] != "": cb_staff.set(staff_names[0])
        cb_staff.grid(row=0, column=1, sticky="w", padx=8, pady=6)

        ctk.CTkLabel(frm, text="Bill Type").grid(row=1, column=0, sticky="w")
        cb_type = ctk.CTkComboBox(frm, values=["Cash Bill","Invoice Bill"], width=200); cb_type.set("Cash Bill")
        cb_type.grid(row=1, column=1, sticky="w", padx=8, pady=6)

        ctk.CTkLabel(frm, text="Bill No.").grid(row=2, column=0, sticky="w")
        e_bill = ctk.CTkEntry(frm, width=260); e_bill.grid(row=2, column=1, sticky="w", padx=8, pady=6)

        ctk.CTkLabel(frm, text="Total Amount").grid(row=3, column=0, sticky="w")
        e_total = ctk.CTkEntry(frm, width=200); e_total.grid(row=3, column=1, sticky="w", padx=8, pady=6)

        ctk.CTkLabel(frm, text="Commission Amount").grid(row=4, column=0, sticky="w")
        e_comm = ctk.CTkEntry(frm, width=200); e_comm.grid(row=4, column=1, sticky="w", padx=8, pady=6)

        photo_lbl = ctk.CTkLabel(frm, text="No bill image selected"); photo_lbl.grid(row=5, column=0, columnspan=2, sticky="w", padx=8, pady=(6,2))
        white_btn(frm, text="Choose Bill Image", width=160,
                  command=lambda: photo_lbl.configure(text=filedialog.askopenfilename(filetypes=[("Image","*.jpg;*.jpeg;*.png")]) or "No bill image selected")
                 ).grid(row=5, column=1, sticky="e")

        def validate_bill():
            s = (e_bill.get() or "").strip().upper()
            btype = cb_type.get()
            if btype == "Cash Bill":
                ok = bool(BILL_RE_CS.match(s))
            else:
                ok = bool(BILL_RE_IV.match(s))
            if not ok:
                messagebox.showerror("Bill No.", "Invalid bill number format.\nCash: CS-MMDD/XXXX\nInvoice: IV-MMDD/XXXX")
            return ok

        def save():
            if not staff_rows:
                messagebox.showerror("Commission", "No staff found. Add staff first."); return
            if not validate_bill(): return
            try:
                total = float(e_total.get())
                commission = float(e_comm.get())
            except Exception:
                messagebox.showerror("Commission", "Amounts must be numeric."); return
            staff_id = int((cb_staff.get().split(":",1)[0] or "0"))
            bill_type = "CS" if cb_type.get()=="Cash Bill" else "IV"
            bill_no = (e_bill.get() or "").strip().upper()
            img_src = photo_lbl.cget("text")
            img_path = ""
            if img_src and img_src != "No bill image selected":
                os.makedirs(IMG_BILLS, exist_ok=True)
                fn = f"{bill_type}_{bill_no.replace('/','-')}_{int(datetime.now().timestamp())}.jpg"
                img_path = os.path.join(IMG_BILLS, fn)
                try:
                    _process_square_image(img_src, img_path, max_px=400)
                except Exception as e:
                    messagebox.showerror("Bill Image", f"Failed to process image: {e}")
                    img_path = ""
            conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
            cur.execute("""INSERT INTO commissions (staff_id, bill_type, bill_no, total_amount, commission_amount, bill_image_path, created_at, updated_at)
                           VALUES (?,?,?,?,?,?,?,?)""",
                        (staff_id, bill_type, bill_no, total, commission, img_path,
                         datetime.now().isoformat(sep=" ", timespec="seconds"),
                         datetime.now().isoformat(sep=" ", timespec="seconds")))
            conn.commit(); conn.close()
            messagebox.showinfo("Commission", "Saved.")
            top.destroy()

        white_btn(frm, text="Save", command=save, width=120).grid(row=6, column=1, sticky="e", pady=10)

    def view_commissions(self):
        top = ctk.CTkToplevel(self); top.title("Commissions"); top.geometry("1000x620"); top.grab_set()
        frm = ctk.CTkFrame(top); frm.pack(fill="both", expand=True, padx=10, pady=10)
        frm.grid_rowconfigure(1, weight=1); frm.grid_columnconfigure(0, weight=1)

        # filter / actions
        bar = ctk.CTkFrame(frm); bar.grid(row=0, column=0, sticky="ew", pady=(0,8))
        e_search = ctk.CTkEntry(bar, width=280, placeholder_text="Search bill no. or staff name")
        e_search.pack(side="left", padx=5)
        white_btn(bar, text="Search", width=110, command=lambda: refresh()).pack(side="left", padx=5)
        white_btn(bar, text="Delete Selected", width=150, command=lambda: delete_sel()).pack(side="right", padx=5)

        # table with horizontal scroll
        wrap = ctk.CTkFrame(frm); wrap.grid(row=1, column=0, sticky="nsew")
        wrap.grid_rowconfigure(0, weight=1); wrap.grid_columnconfigure(0, weight=1)
        tree = ttk.Treeview(wrap, columns=("id","staff","type","billno","total","commission","img","created"),
                            show="headings", selectmode="browse")
        for c, t, w in [
            ("id","ID",60), ("staff","Staff",220), ("type","Type",80),
            ("billno","Bill No.",160), ("total","Total",110), ("commission","Commission",120),
            ("img","Image",240), ("created","Created",160)
        ]:
            tree.heading(c, text=t); tree.column(c, width=w, anchor="w", stretch=True)
        tree.grid(row=0, column=0, sticky="nsew")
        sbv = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview); sbv.grid(row=0, column=1, sticky="ns")
        sbh = ttk.Scrollbar(wrap, orient="horizontal", command=tree.xview); sbh.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=sbv.set, xscrollcommand=sbh.set)

        def refresh():
            q = (e_search.get() or "").strip().lower()
            tree.delete(*tree.get_children())
            conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
            cur.execute("""SELECT c.id, s.name, c.bill_type, c.bill_no, c.total_amount, c.commission_amount, c.bill_image_path, c.created_at
                           FROM commissions c JOIN staffs s ON c.staff_id=s.id
                           ORDER BY c.id DESC""")
            for r in cur.fetchall():
                if q:
                    if q not in str(r[1]).lower() and q not in str(r[3]).lower():
                        continue
                tree.insert("", "end", values=r)
            conn.close()

        def open_edit():
            sel = tree.selection()
            if not sel:
                return
            rid = tree.item(sel[0])["values"][0]

            # Load record
            conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
            cur.execute("""SELECT staff_id, bill_type, bill_no, total_amount, commission_amount, bill_image_path
                           FROM commissions WHERE id=?""", (rid,))
            row = cur.fetchone(); conn.close()
            if not row:
                return
            staff_id, bill_type, bill_no, total_amount, commission_amount, bill_image_path = row

            # Dialog
            dlg = ctk.CTkToplevel(self)
            dlg.title(f"Edit Commission {rid}")
            dlg.geometry("760x560")
            dlg.grab_set()

            f = ctk.CTkFrame(dlg); f.pack(fill="both", expand=True, padx=10, pady=10)
            f.grid_columnconfigure(1, weight=1)

            # Staff list
            conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
            cur.execute("SELECT id, name FROM staffs ORDER BY name")
            staff_rows = cur.fetchall(); conn.close()
            staff_names = [f"{i}:{n}" for (i, n) in staff_rows] or [""]

            ctk.CTkLabel(f, text="Staff").grid(row=0, column=0, sticky="w")
            cb_staff = ctk.CTkComboBox(f, values=staff_names, width=320)
            name_map = {i: n for (i, n) in staff_rows}
            pre_staff = f"{staff_id}:{name_map.get(staff_id, 'Unknown')}" if staff_rows else ""
            # Ensure the prefill is one of the options (CTkComboBox can show non-listed text, but nicer to include)
            if pre_staff and pre_staff not in staff_names:
                cb_staff.configure(values=staff_names + [pre_staff])
            if pre_staff:
                cb_staff.set(pre_staff)
            cb_staff.grid(row=0, column=1, sticky="w", padx=8, pady=6)

            # Bill type
            ctk.CTkLabel(f, text="Bill Type").grid(row=1, column=0, sticky="w")
            cb_type = ctk.CTkComboBox(f, values=["Cash Bill", "Invoice Bill"], width=200)
            cb_type.set("Cash Bill" if bill_type == "CS" else "Invoice Bill")
            cb_type.grid(row=1, column=1, sticky="w", padx=8, pady=6)

            # Bill no
            ctk.CTkLabel(f, text="Bill No.").grid(row=2, column=0, sticky="w")
            e_bill = ctk.CTkEntry(f, width=260); e_bill.insert(0, bill_no)
            e_bill.grid(row=2, column=1, sticky="w", padx=8, pady=6)

            # Amounts
            ctk.CTkLabel(f, text="Total Amount").grid(row=3, column=0, sticky="w")
            e_total = ctk.CTkEntry(f, width=200); e_total.insert(0, str(total_amount))
            e_total.grid(row=3, column=1, sticky="w", padx=8, pady=6)

            ctk.CTkLabel(f, text="Commission Amount").grid(row=4, column=0, sticky="w")
            e_comm = ctk.CTkEntry(f, width=200); e_comm.insert(0, str(commission_amount))
            e_comm.grid(row=4, column=1, sticky="w", padx=8, pady=6)

            # Bill image
            bill_lbl = ctk.CTkLabel(f, text=bill_image_path or "No bill image selected")
            bill_lbl.grid(row=5, column=0, columnspan=2, sticky="w", padx=8, pady=(6, 2))
            white_btn(
                f, text="Choose Bill Image", width=160,
                command=lambda: bill_lbl.configure(
                    text=filedialog.askopenfilename(filetypes=[("Image", "*.jpg;*.jpeg;*.png")]) or bill_lbl.cget("text")
                )
            ).grid(row=5, column=1, sticky="e")


            def validate_bill():
                s = (e_bill.get() or "").strip().upper()
                btype = cb_type.get()
                return bool(BILL_RE_CS.match(s)) if btype=="Cash Bill" else bool(BILL_RE_IV.match(s))

            def save_edit():
                if not validate_bill():
                    messagebox.showerror("Bill No.", "Invalid bill number format."); return
                try:
                    total = float(e_total.get()); commission = float(e_comm.get())
                except Exception:
                    messagebox.showerror("Commission", "Amounts must be numeric."); return
                staff_id_new = int((cb_staff.get().split(":",1)[0] or "0"))
                bill_type_new = "CS" if cb_type.get()=="Cash Bill" else "IV"
                bill_no_new = (e_bill.get() or "").strip().upper()
                img_src = bill_lbl.cget("text")
                img_path_new = bill_image_path
                if img_src and img_src not in ("", bill_image_path) and os.path.exists(img_src):
                    os.makedirs(IMG_BILLS, exist_ok=True)
                    fn = f"{bill_type_new}_{bill_no_new.replace('/','-')}_{int(datetime.now().timestamp())}.jpg"
                    img_path_new = os.path.join(IMG_BILLS, fn)
                    try:
                        _process_square_image(img_src, img_path_new, max_px=400)
                    except Exception as e:
                        messagebox.showerror("Bill Image", f"Failed to process image: {e}")
                        img_path_new = bill_image_path

                conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
                cur.execute("""UPDATE commissions SET staff_id=?, bill_type=?, bill_no=?, total_amount=?, commission_amount=?, bill_image_path=?, updated_at=?
                               WHERE id=?""",
                            (staff_id_new, bill_type_new, bill_no_new, total, commission, img_path_new,
                             datetime.now().isoformat(sep=" ", timespec="seconds"), rid))
                conn.commit(); conn.close()
                messagebox.showinfo("Commission", "Updated.")
                dlg.destroy(); refresh()

            white_btn(f, text="Save Changes", width=150, command=save_edit).grid(row=6, column=1, sticky="e", pady=10)

        def delete_sel():
            sel = tree.selection()
            if not sel: return
            rid = tree.item(sel[0])["values"][0]
            if not messagebox.askyesno("Delete", "Delete this commission record? This cannot be undone."): return
            conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
            cur.execute("DELETE FROM commissions WHERE id=?", (rid,))
            conn.commit(); conn.close()
            refresh()

        tree.bind("<Double-1>", lambda e: open_edit())
        refresh()

    # ---------- Utility ----------
    def _go_fullscreen(self):
        try:
            if sys.platform.startswith("win") or sys.platform.startswith("linux"):
                self.state("zoomed")
            else:
                self.attributes("-fullscreen", True)
        except Exception:
            self.attributes("-fullscreen", True)


# ------------------ Run ------------------
if __name__ == "__main__":
    init_db()
    app = VoucherApp()
    app.mainloop()

