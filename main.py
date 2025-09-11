#!/usr/bin/env python3
# Service Voucher Management System (Monolith)
# Deps: customtkinter, reportlab, qrcode, pillow
# Run:  python main.py

import os
import sys
import sqlite3
import webbrowser
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk
import qrcode
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader

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

# ------------------ Config ------------------
DB_FILE   = "vouchers.db"
PDF_DIR   = "pdfs"
os.makedirs(PDF_DIR, exist_ok=True)

SHOP_NAME = "TONY.COM"
SHOP_ADDR = "TB4318, Lot 5, Block 31, Fajar Complex  91000 Tawau Sabah, Malaysia"
SHOP_TEL  = "Tel : 089-763778, H/P: 0168260533"
LOGO_PATH = ""  # optional: path to logo image (png/jpg). leave blank to skip

BASE_VOUCHER_NO = 41000

# ------------------ Date helpers (DD-MM-YYYY for UI/PDF) ------------------
def _to_ui_date(dt: datetime) -> str:
    return dt.strftime("%d-%m-%Y")

def _to_ui_datetime_str(iso_str: str) -> str:
    try:
        dt = datetime.strptime(iso_str, "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%d-%m-%Y %H:%M:%S")
    except Exception:
        return iso_str

def _from_ui_date_to_sqldate(s: str) -> str:
    """UI date 'DD-MM-YYYY' -> 'YYYY-MM-DD' for SQL range filtering."""
    s = (s or "").strip()
    if not s:
        return ""
    try:
        d = datetime.strptime(s, "%d-%m-%Y")
        return d.strftime("%Y-%m-%d")
    except Exception:
        return ""

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
    pdf_path TEXT
);
CREATE TABLE IF NOT EXISTS staffs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE
);
"""

def _column_exists(cur, table, column):
    cur.execute(f"PRAGMA table_info({table})")
    return any(row[1] == column for row in cur.fetchall())

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.executescript(BASE_SCHEMA)
    for tbl, wanted in {
        "vouchers": [
            ("voucher_id","TEXT"), ("created_at","TEXT"), ("customer_name","TEXT"),
            ("contact_number","TEXT"), ("units","INTEGER"),
            ("particulars","TEXT"), ("problem","TEXT"), ("staff_name","TEXT"),
            ("status","TEXT"), ("recipient","TEXT"), ("pdf_path","TEXT")
        ],
        "staffs": [("name","TEXT")]
    }.items():
        for col, typ in wanted:
            if not _column_exists(cur, tbl, col):
                cur.execute(f"ALTER TABLE {tbl} ADD COLUMN {col} {typ}")
    try:
        cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_vouchers_vid ON vouchers(voucher_id)")
        cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_staffs_name ON staffs(name)")
    except Exception:
        pass
    conn.commit()
    conn.close()

def next_voucher_id():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT MAX(CAST(voucher_id AS INTEGER)) FROM vouchers")
    row = cur.fetchone()
    conn.close()
    if not row or row[0] is None:
        return str(BASE_VOUCHER_NO)
    return str(int(row[0]) + 1)

# ---- Staff/Recipient ops ----
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

# ------------------ PDF ------------------
def _draw_voucher(c, width, height, voucher_id, customer_name, contact_number,
                  units, particulars, problem, staff_name, created_at, recipient):
    """
    A4 voucher fits in half a page.
    - Narrower left column.
    - QTY split: top-only; bottom merged with PROBLEM.
    - Ack sentence removed.
    - Recipient line 3mm below label (moved up by 2mm from before).
    """

    # Margins / frame
    left   = 12*mm
    right  = width - 12*mm
    top_y  = height - 15*mm
    y      = top_y

    # Column widths (narrower left)
    qty_col_w    = 20*mm
    left_col_w   = 74*mm
    middle_col_w = (right - left) - left_col_w - qty_col_w

    name_col_x = left + left_col_w
    qty_col_x  = right - qty_col_w

    # Row heights
    row1_h = 20*mm
    row2_h = 20*mm
    total_h = row1_h + row2_h

    # --- Header ---
    c.setFont("Helvetica-Bold", 14); c.drawString(left, y, SHOP_NAME)
    c.setFont("Helvetica", 9.2)
    c.drawString(left, y - 5.0*mm, SHOP_ADDR)
    c.drawString(left, y - 9.0*mm, SHOP_TEL)

    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString((left+right)/2, y - 16.0*mm, "SERVICE VOUCHER")
    c.drawRightString(right, y - 16.0*mm, f"No : {voucher_id}")

    # Parse created_at -> DD-MM-YYYY and HH:MM:SS
    try:
        dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
        date_str = dt.strftime("%d-%m-%Y")
        time_str = dt.strftime("%H:%M:%S")
    except Exception:
        date_str = created_at[:10]
        time_str = created_at[11:19]

    c.setFont("Helvetica", 10)
    c.drawString(left, y - 24.0*mm, "Date :")
    c.drawString(left + 18*mm, y - 24.0*mm, date_str)
    c.drawRightString(right - 27*mm, y - 24.0*mm, "Time In :")
    c.drawRightString(right, y - 24.0*mm, time_str)

    # --- Table frame ---
    top_table = y - 28*mm
    bottom_table = top_table - total_h
    mid_y = top_table - row1_h
    pad = 3*mm

    # Outer border
    c.rect(left, bottom_table, right-left, total_h, stroke=1, fill=0)
    # Vertical at middle (full)
    c.line(name_col_x, top_table, name_col_x, bottom_table)
    # Vertical for QTY ONLY on the TOP row (split)
    c.line(qty_col_x, top_table, qty_col_x, mid_y)
    # Horizontal split between rows (full width)
    c.line(left, mid_y, right, mid_y)

    # Row 1 Left: CUSTOMER NAME
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left + pad, top_table - pad - 8, "CUSTOMER NAME")
    draw_wrapped(c, customer_name,
                 left + pad, mid_y + pad,
                 w=left_col_w - 2*pad, h=row1_h - 2*pad - 10, fontsize=10)

    # Row 1 Middle: PARTICULARS
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(name_col_x + pad, top_table - pad - 8, "PARTICULARS")
    draw_wrapped(c, particulars or "-",
                 name_col_x + pad, mid_y + pad,
                 w=middle_col_w - 2*pad, h=row1_h - 2*pad - 10, fontsize=10)

    # Row 1 Right: QTY top cell
    c.setFont("Helvetica-Bold", 10.4)
    c.drawCentredString(qty_col_x + qty_col_w/2, top_table - pad - 8, "QTY")
    c.setFont("Helvetica", 11)
    # Center of the TOP row = mid_y + row1_h/2
    c.drawCentredString(qty_col_x + qty_col_w/2, mid_y + (row1_h/2) - 3, str(units or 1))

    # Row 2 Left: TEL
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left + pad, mid_y - pad - 8, "TEL")
    draw_wrapped(c, contact_number,
                 left + pad, bottom_table + pad,
                 w=left_col_w - 2*pad, h=row2_h - 2*pad - 10, fontsize=10)

    # Row 2 Middle+Right merged: PROBLEM (spans middle_col_w + qty_col_w)
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(name_col_x + pad, mid_y - pad - 8, "PROBLEM")
    draw_wrapped(c, problem or "-",
                 name_col_x + pad, bottom_table + pad,
                 w=(middle_col_w + qty_col_w) - 2*pad, h=row2_h - 2*pad - 10, fontsize=10)

    # ------- Recipient label + line (line 3mm below the label) -------
    y_rec = bottom_table - 9*mm
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left, y_rec, "RECIPIENT :")
    label_w = c.stringWidth("RECIPIENT :", "Helvetica-Bold", 10.4)
    line_x0 = left + label_w + 6
    line_x1 = left + left_col_w - 2*mm
    line_y  = y_rec - 3*mm  # moved UP by 2mm from previous 5mm
    c.line(line_x0, line_y, line_x1, line_y)
    if recipient:
        c.setFont("Helvetica", 9)
        c.drawString(line_x0 + 1*mm, line_y + 2.2*mm, recipient)

    # ------- Company policies (left half only) -------
    policies_top = y_rec - 7*mm
    policies_w   = left_col_w - 1.5*mm

    p1 = "Kindly collect your goods within <font color='red' size='9'>60 days</font> from date of sending for repair."
    used_h = draw_wrapped_top(c, p1, left, policies_top, policies_w, fontsize=7.5, leading=10)
    y_cursor = policies_top - used_h - 2

    used_h = draw_wrapped_top(c, "A) We do not hold ourselves responsible for any loss or damage.",
                              left, y_cursor, policies_w, fontsize=7.5, leading=10)
    y_cursor -= used_h - 1

    used_h = draw_wrapped_top(c, "B) We reserve our right to sell off the goods to cover our cost and loss.",
                              left, y_cursor, policies_w, fontsize=7.5, leading=10)
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

    # QR code (aligned roughly with policies block)
    qr_size = 20*mm
    try:
        qr_data = f"Voucher:{voucher_id}|Name:{customer_name}|Tel:{contact_number}|Date:{date_str}"
        qr_img  = qrcode.make(qr_data)
        qr_y = policies_top - (qr_size * 0.2)
        c.drawImage(ImageReader(qr_img), right - qr_size, qr_y - qr_size, qr_size, qr_size)
    except Exception:
        pass

    # ------- Signatures (side-by-side, a bit lower, stay in top half) -------
    sig_gap_above = 6*mm
    candidate_y_sig = policies_bottom - sig_gap_above
    half_limit = height/2 + 5*mm
    y_sig = max(candidate_y_sig, half_limit)

    SIG_LINE_W = 45*mm
    SIG_GAP    = 6*mm
    sig_left_start = right - (2*SIG_LINE_W + SIG_GAP)

    c.line(sig_left_start, y_sig, sig_left_start + SIG_LINE_W, y_sig)
    c.setFont("Helvetica", 8.8)
    c.drawString(sig_left_start, y_sig - 3.6*mm, "CUSTOMER SIGNATURE")

    right_line_x0 = sig_left_start + SIG_LINE_W + SIG_GAP
    c.line(right_line_x0, y_sig, right_line_x0 + SIG_LINE_W, y_sig)
    c.drawString(right_line_x0, y_sig - 3.6*mm, "DATE COLLECTED")

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
def add_voucher(customer_name, contact_number, units, particulars, problem, staff_name, recipient=""):
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
                              particulars, problem, staff_name, status, recipient, pdf_path)
        VALUES (?,?,?,?,?,?,?,?,?,?,?)
    """, (voucher_id, created_at, customer_name, contact_number, units,
          particulars, problem, staff_name, status, recipient, pdf_path))

    conn.commit(); conn.close()
    return voucher_id, pdf_path

def mark_collected(voucher_id):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("UPDATE vouchers SET status='Collected' WHERE voucher_id = ?", (voucher_id,))
    conn.commit()
    conn.close()

def delete_and_compact(voucher_id_str: str):
    """
    Delete voucher with ID = voucher_id_str and then renumber
    all vouchers with ID > deleted so numbers remain contiguous.
    PDFs are regenerated and file paths updated.
    """
    target = int(voucher_id_str)
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    # 1) Delete the row and remove its PDF
    cur.execute("SELECT pdf_path FROM vouchers WHERE voucher_id = ?", (voucher_id_str,))
    row = cur.fetchone()
    if not row:
        conn.close()
        raise ValueError("Voucher not found.")

    old_pdf = row[0]
    cur.execute("DELETE FROM vouchers WHERE voucher_id = ?", (voucher_id_str,))
    if old_pdf and os.path.exists(old_pdf):
        try: os.remove(old_pdf)
        except Exception: pass

    # 2) Fetch all vouchers with id > target, in ascending order
    cur.execute("""
        SELECT voucher_id, created_at, customer_name, contact_number, units,
               particulars, problem, staff_name, status, recipient
        FROM vouchers
        WHERE CAST(voucher_id AS INTEGER) > ?
        ORDER BY CAST(voucher_id AS INTEGER) ASC
    """, (target,))
    rows = cur.fetchall()

    # 3) For each, decrement by 1 and regenerate PDF
    for (vid, created_at, customer_name, contact_number, units,
         particulars, problem, staff_name, status, recipient) in rows:
        old_id = int(vid)
        new_id = old_id - 1

        # Remove old PDF file if exists
        cur.execute("SELECT pdf_path FROM vouchers WHERE voucher_id=?", (vid,))
        r = cur.fetchone()
        if r and r[0] and os.path.exists(r[0]):
            try: os.remove(r[0])
            except Exception: pass

        # Update voucher_id first
        cur.execute("UPDATE vouchers SET voucher_id=? WHERE voucher_id=?", (str(new_id), vid))

        # Regenerate PDF with new voucher number
        new_pdf = generate_pdf(
            str(new_id), customer_name, contact_number, units,
            particulars, problem, staff_name, status, created_at, recipient
        )
        cur.execute("UPDATE vouchers SET pdf_path=? WHERE voucher_id=?", (new_pdf, str(new_id)))

    conn.commit()
    conn.close()

def search_vouchers(filters):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    sql = ("SELECT voucher_id, created_at, customer_name, contact_number, units, status, recipient, pdf_path "
           "FROM vouchers WHERE 1=1")
    params = []
    if filters.get("voucher_id"):
        sql += " AND voucher_id LIKE ?"; params.append(f"%{filters['voucher_id']}%")
    if filters.get("name"):
        sql += " AND customer_name LIKE ?"; params.append(f"%{filters['name']}%")
    if filters.get("contact"):
        sql += " AND contact_number LIKE ?"; params.append(f"%{filters['contact']}%")
    if filters.get("date_from"):
        sql += " AND created_at >= ?"; params.append(filters["date_from"].strip() + " 00:00:00")
    if filters.get("date_to"):
        sql += " AND created_at <= ?"; params.append(filters["date_to"].strip() + " 23:59:59")
    if filters.get("status") and filters["status"] != "All":
        sql += " AND status = ?"; params.append(filters["status"])
    sql += " ORDER BY CAST(voucher_id AS INTEGER) DESC"
    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()
    return rows

# ------------------ UI ------------------
class VoucherApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Service Voucher Management System")
        self.geometry("1100x650")
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.minsize(900, 520)

        # ===== Layout =====
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # ---- Filters ----
        filt = ctk.CTkFrame(self)
        filt.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        for i in range(10):
            filt.grid_columnconfigure(i, weight=0)
        filt.grid_columnconfigure(9, weight=1)

        today_ui = _to_ui_date(datetime.now())
        self.f_voucher = ctk.CTkEntry(filt, width=130, placeholder_text="VoucherID");    self.f_voucher.grid(row=0, column=0, padx=5)
        self.f_name    = ctk.CTkEntry(filt, width=220, placeholder_text="Customer Name"); self.f_name.grid(row=0, column=1, padx=5)
        self.f_contact = ctk.CTkEntry(filt, width=180, placeholder_text="Contact Number"); self.f_contact.grid(row=0, column=2, padx=5)
        self.f_from    = ctk.CTkEntry(filt, width=160, placeholder_text="Date From (DD-MM-YYYY)"); self.f_from.grid(row=0, column=3, padx=5)
        self.f_to      = ctk.CTkEntry(filt, width=160, placeholder_text="Date To (DD-MM-YYYY)");   self.f_to.grid(row=0, column=4, padx=5)
        self.f_from.insert(0, today_ui); self.f_to.insert(0, today_ui)

        self.f_status  = ctk.CTkOptionMenu(filt, values=["All","Pending","Collected"], width=120); self.f_status.grid(row=0, column=5, padx=(8,5))
        self.f_status.set("All")
        self.btn_search = ctk.CTkButton(filt, text="Search", command=self.perform_search, width=100); self.btn_search.grid(row=0, column=6, padx=5)
        self.btn_reset  = ctk.CTkButton(filt, text="Reset",  command=self.reset_filters, width=80);   self.btn_reset.grid(row=0, column=7, padx=5)

        # ---- Table (Recipient before Status; Status last) ----
        table_frame = ctk.CTkFrame(self)
        table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 8))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_frame,
            columns=("VoucherID","Date","Customer","Contact","Units","Recipient","Status","PDF"),
            show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew")

        table_vbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        table_hbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=table_vbar.set, xscrollcommand=table_hbar.set)
        table_vbar.grid(row=0, column=1, sticky="ns")
        table_hbar.grid(row=1, column=0, sticky="ew")

        self.col_weights = {"VoucherID":1, "Date":2, "Customer":2, "Contact":2, "Units":1, "Recipient":2, "Status":1}
        for col in ("VoucherID","Date","Customer","Contact","Units","Recipient","Status"):
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

        self.tree.bind("<Double-1>", lambda e: self.open_pdf())

        # ---- Actions ----
        bar = ctk.CTkFrame(self)
        bar.grid(row=2, column=0, sticky="ew", padx=10, pady=(0,10))
        for i in range(10): bar.grid_columnconfigure(i, weight=0)

        ctk.CTkButton(bar, text="Add Voucher",       command=self.add_voucher_ui,    width=120).grid(row=0, column=0, padx=6, pady=8)
        ctk.CTkButton(bar, text="Edit Selected",     command=self.edit_selected,     width=120).grid(row=0, column=1, padx=6, pady=8)
        ctk.CTkButton(bar, text="Mark as Collected", command=self.mark_selected,     width=140).grid(row=0, column=2, padx=6, pady=8)
        ctk.CTkButton(bar, text="Open PDF",          command=self.open_pdf,          width=110).grid(row=0, column=3, padx=6, pady=8)
        ctk.CTkButton(bar, text="Manage Recipients", command=self.manage_staffs_ui,  width=160).grid(row=0, column=4, padx=6, pady=8)
        ctk.CTkButton(bar, text="Delete Selected",   command=self.delete_selected,   width=130, fg_color="#d9534f", hover_color="#c9302c").grid(row=0, column=5, padx=6, pady=8)

        self.perform_search()

    # ---- Filters helpers ----
    def _get_filters(self):
        return {
            "voucher_id": self.f_voucher.get().strip(),
            "name": self.f_name.get().strip(),
            "contact": self.f_contact.get().strip(),
            # Convert UI DD-MM-YYYY -> SQL YYYY-MM-DD
            "date_from": _from_ui_date_to_sqldate(self.f_from.get().strip()),
            "date_to": _from_ui_date_to_sqldate(self.f_to.get().strip()),
            "status": self.f_status.get(),
        }

    def reset_filters(self):
        for e in (self.f_voucher, self.f_name, self.f_contact, self.f_from, self.f_to):
            e.delete(0, "end")
        today_ui = _to_ui_date(datetime.now())
        self.f_from.insert(0, today_ui)
        self.f_to.insert(0, today_ui)
        self.f_status.set("All")
        self.perform_search()

    def perform_search(self):
        rows = search_vouchers(self._get_filters())
        self.tree.delete(*self.tree.get_children())
        for (vid, created_at, customer, contact, units, status, recipient, pdf) in rows:
            # match new column order: ..., Recipient, Status
            self.tree.insert("", "end", values=(
                vid,
                _to_ui_datetime_str(created_at),
                customer, contact, units,
                recipient, status, pdf
            ))

    # ---- Create Voucher ----
    def add_voucher_ui(self):
        top = ctk.CTkToplevel(self)
        top.title("Create Voucher")
        top.geometry("860x620")
        top.grab_set()

        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        WIDE = 520

        r = 0
        ctk.CTkLabel(frm, text="Customer Name", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        e_name = ctk.CTkEntry(frm, width=WIDE); e_name.grid(row=r, column=1, sticky="w", padx=10, pady=(0,8)); r+=1

        ctk.CTkLabel(frm, text="Contact Number", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        e_contact = ctk.CTkEntry(frm, width=WIDE); e_contact.grid(row=r, column=1, sticky="w", padx=10, pady=(0,8)); r+=1

        ctk.CTkLabel(frm, text="No. of Units", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        e_units = ctk.CTkEntry(frm, width=120); e_units.insert(0,"1")
        e_units.grid(row=r, column=1, sticky="w", padx=10, pady=(0,8)); r+=1

        ctk.CTkLabel(frm, text="Particulars", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        t_part = tk.Text(frm, width=66, height=5); t_part.grid(row=r, column=1, sticky="w", padx=10, pady=(0,8)); r+=1

        ctk.CTkLabel(frm, text="Problem", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        t_prob = tk.Text(frm, width=66, height=4); t_prob.grid(row=r, column=1, sticky="w", padx=10, pady=(0,8)); r+=1

        # Staff Name becomes Recipient (same combobox, same source list)
        ctk.CTkLabel(frm, text="Recipient", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        staff_values = list_staffs() or [""]
        e_recipient = ctk.CTkComboBox(frm, values=staff_values, width=WIDE)
        if staff_values and staff_values[0] != "":
            e_recipient.set(staff_values[0])
        e_recipient.grid(row=r, column=1, sticky="w", padx=10, pady=(0,8)); r+=1

        btns = ctk.CTkFrame(top); btns.pack(fill="x", padx=12, pady=(0,12))

        def save():
            name = e_name.get().strip()
            contact = e_contact.get().strip()
            try:
                units = int((e_units.get() or "1").strip())
            except ValueError:
                messagebox.showerror("Invalid", "Units must be a number."); return
            particulars = t_part.get("1.0","end").strip()
            problem     = t_prob.get("1.0","end").strip()
            recipient   = e_recipient.get().strip()
            staff_name  = recipient  # inherit properties: keep both columns synced

            if not name or not contact:
                messagebox.showerror("Missing", "Customer name and contact are required."); return
            voucher_id, pdf_path = add_voucher(name, contact, units, particulars, problem, staff_name, recipient)
            messagebox.showinfo("Saved", f"Voucher {voucher_id} created.")
            try:
                webbrowser.open_new(os.path.abspath(pdf_path))
            except Exception:
                pass
            top.destroy()
            self.perform_search()

        ctk.CTkButton(btns, text="Save & Open PDF", command=save, width=180).pack(side="right")

    # ---- Manage Recipients (formerly Staffs) ----
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
    def mark_selected(self):
        sel = self.tree.focus()
        if not sel:
            messagebox.showerror("Error", "Select a record first."); return
        voucher_id = self.tree.item(sel)["values"][0]
        mark_collected(voucher_id)
        messagebox.showinfo("Updated", f"Voucher {voucher_id} marked as Collected.")
        self.perform_search()

    def delete_selected(self):
        sel = self.tree.focus()
        if not sel:
            messagebox.showerror("Error", "Select a record first."); return
        voucher_id = str(self.tree.item(sel)["values"][0])

        if not messagebox.askyesno("Confirm Delete",
                                   f"Delete voucher {voucher_id}?\n"
                                   f"All later vouchers will shift up by 1 and PDFs will be regenerated."):
            return
        try:
            delete_and_compact(voucher_id)
            messagebox.showinfo("Deleted", f"Voucher {voucher_id} deleted and IDs compacted.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete: {e}")
        self.perform_search()

    def open_pdf(self):
        sel = self.tree.focus()
        if not sel:
            messagebox.showerror("Error", "Select a record first."); return
        pdf_path = self.tree.item(sel)["values"][-1]
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Error", "PDF not found for this voucher."); return
        try:
            webbrowser.open_new(os.path.abspath(pdf_path))
        except Exception:
            if os.name == "nt":
                os.startfile(pdf_path)  # type: ignore
            else:
                os.system(f"open '{pdf_path}'" if sys.platform == "darwin" else f"xdg-open '{pdf_path}'")

    def edit_selected(self):
        sel = self.tree.focus()
        if not sel:
            messagebox.showerror("Error", "Select a record first."); return

        values = self.tree.item(sel)["values"]
        if not values:
            return

        voucher_id = values[0]

        # fetch full record from DB
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("""
            SELECT voucher_id, created_at, customer_name, contact_number, units,
                   particulars, problem, staff_name, status, recipient
            FROM vouchers WHERE voucher_id = ?
        """, (voucher_id,))
        row = cur.fetchone()
        conn.close()
        if not row: return

        _, created_at, customer_name, contact_number, units, particulars, problem, staff_name, status, recipient = row

        # open edit window
        top = ctk.CTkToplevel(self)
        top.title(f"Edit Voucher {voucher_id}")
        top.geometry("860x620")
        top.grab_set()

        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        WIDE = 520
        r = 0
        ctk.CTkLabel(frm, text="Customer Name").grid(row=r, column=0, sticky="w")
        e_name = ctk.CTkEntry(frm, width=WIDE)
        e_name.insert(0, customer_name)
        e_name.grid(row=r, column=1, sticky="w", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Contact Number").grid(row=r, column=0, sticky="w")
        e_contact = ctk.CTkEntry(frm, width=WIDE)
        e_contact.insert(0, contact_number)
        e_contact.grid(row=r, column=1, sticky="w", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="No. of Units").grid(row=r, column=0, sticky="w")
        e_units = ctk.CTkEntry(frm, width=120)
        e_units.insert(0, str(units))
        e_units.grid(row=r, column=1, sticky="w", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Particulars").grid(row=r, column=0, sticky="w")
        t_part = tk.Text(frm, width=66, height=5)
        t_part.insert("1.0", particulars or "")
        t_part.grid(row=r, column=1, sticky="w", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Problem").grid(row=r, column=0, sticky="w")
        t_prob = tk.Text(frm, width=66, height=4)
        t_prob.insert("1.0", problem or "")
        t_prob.grid(row=r, column=1, sticky="w", padx=10, pady=6); r+=1

        # Recipient combobox (replaces Staff)
        ctk.CTkLabel(frm, text="Recipient").grid(row=r, column=0, sticky="w")
        staff_values = list_staffs() or [""]
        e_recipient = ctk.CTkComboBox(frm, values=staff_values, width=WIDE)
        e_recipient.set(recipient or staff_name or "")
        e_recipient.grid(row=r, column=1, sticky="w", padx=10, pady=6); r+=1

        def save_edit():
            name = e_name.get().strip()
            contact = e_contact.get().strip()
            try:
                units_val = int((e_units.get() or "1").strip())
            except ValueError:
                messagebox.showerror("Invalid", "Units must be a number."); return
            particulars_val = t_part.get("1.0","end").strip()
            problem_val     = t_prob.get("1.0","end").strip()
            recipient_val   = e_recipient.get().strip()
            staff_val       = recipient_val  # keep in sync

            # update DB
            conn = sqlite3.connect(DB_FILE)
            cur = conn.cursor()
            cur.execute("""UPDATE vouchers
                              SET customer_name=?, contact_number=?, units=?, particulars=?,
                                  problem=?, staff_name=?, recipient=?
                            WHERE voucher_id=?""",
                        (name, contact, units_val, particulars_val,
                         problem_val, staff_val, recipient_val, voucher_id))
            conn.commit()

            # regenerate PDF to reflect edits
            pdf_path = generate_pdf(
                str(voucher_id), name, contact, units_val,
                particulars_val, problem_val, staff_val, status, created_at, recipient_val
            )
            cur.execute("UPDATE vouchers SET pdf_path=? WHERE voucher_id=?", (pdf_path, str(voucher_id)))
            conn.commit()
            conn.close()

            messagebox.showinfo("Updated", f"Voucher {voucher_id} updated.")
            top.destroy()
            self.perform_search()

        btns = ctk.CTkFrame(top); btns.pack(fill="x", padx=12, pady=12)
        ctk.CTkButton(btns, text="Save Changes", command=save_edit, width=160).pack(side="right", padx=6)
        ctk.CTkButton(btns, text="Cancel", command=top.destroy, width=100).pack(side="right", padx=6)

    def open_pdf(self):
        sel = self.tree.focus()
        if not sel:
            messagebox.showerror("Error", "Select a record first."); return
        pdf_path = self.tree.item(sel)["values"][-1]
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Error", "PDF not found for this voucher."); return
        try:
            webbrowser.open_new(os.path.abspath(pdf_path))
        except Exception:
            if os.name == "nt":
                os.startfile(pdf_path)  # type: ignore
            else:
                os.system(f"open '{pdf_path}'" if sys.platform == "darwin" else f"xdg-open '{pdf_path}'")

# ------------------ Run ------------------
if __name__ == "__main__":
    init_db()
    app = VoucherApp()
    app.mainloop()
