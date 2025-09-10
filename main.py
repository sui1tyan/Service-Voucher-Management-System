#!/usr/bin/env python3
# Service Voucher Management System (Monolith)
# Deps: customtkinter, reportlab, qrcode, pillow
# Run:  python main.py

import os
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

# ---------- Wrapped text helper (keeps text neatly inside cells) ----------
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet
_styles = getSampleStyleSheet()
_styleN = _styles["Normal"]

def draw_wrapped(c, text, x, y, w, h, fontsize=10, bold=False, leading=None):
    """
    Draw text wrapped inside box (x, y, w, h).
    NOTE: y is the *bottom* of the box, but we draw from the TOP down
    so labels/borders never collide with content.
    """
    style = _styleN.clone('wrap')
    style.fontName = "Helvetica-Bold" if bold else "Helvetica"
    style.fontSize = fontsize
    style.leading = leading if leading else fontsize + 2
    para = Paragraph((text or "-").replace("\n", "<br/>"), style)
    w_used, h_used = para.wrap(w, h)
    para.drawOn(c, x, y + h - h_used)

# ------------------ Config ------------------
DB_FILE   = "vouchers.db"
PDF_DIR   = "pdfs"
os.makedirs(PDF_DIR, exist_ok=True)

SHOP_NAME = "TONY.COM"
SHOP_ADDR = "TB4318, Lot 5, Block 31, Fajar Complex  91000 Tawau Sabah, Malaysia"
SHOP_TEL  = "Tel : 089-763778, H/P: 0168260533"
LOGO_PATH = ""  # optional: path to logo image (png/jpg). leave blank to skip

BASE_VOUCHER_NO = 41000

# ------------------ DB ------------------
BASE_SCHEMA = """
CREATE TABLE IF NOT EXISTS vouchers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    voucher_id TEXT UNIQUE,
    created_at TEXT,
    customer_name TEXT,
    contact_number TEXT,
    units INTEGER DEFAULT 1,
    remark TEXT,
    particulars TEXT,
    problem TEXT,
    staff_name TEXT,
    status TEXT DEFAULT 'Pending',
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
    # Back-compat safety (ensure all columns exist)
    for tbl, wanted in {
        "vouchers": [
            ("voucher_id","TEXT"), ("created_at","TEXT"), ("customer_name","TEXT"),
            ("contact_number","TEXT"), ("units","INTEGER"), ("remark","TEXT"),
            ("particulars","TEXT"), ("problem","TEXT"), ("staff_name","TEXT"),
            ("status","TEXT"), ("pdf_path","TEXT")
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

# ---- Staff ops ----
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
                  units, remark, particulars, problem, staff_name, created_at):
    """
    A4 single voucher matching your printed template:
    - 2-row table (row1: CUSTOMER NAME + PARTICULARS, row2: TEL + PROBLEM)
    - QTY is one tall cell spanning both rows
    - Labels & values share the same cell
    - 'Recipient' line + acknowledgement sentence to the right
    - 'Remark' outside table
    - Signature + Date Collected on the same line
    - Bottom disclaimers including highlighted RM60.00
    """

    # Page frame
    left   = 12*mm
    right  = width - 12*mm
    top_y  = height - 15*mm
    y      = top_y

    # Column widths (tuned for your original print)
    qty_col_w    = 20*mm
    left_col_w   = 84*mm
    middle_col_w = (right - left) - left_col_w - qty_col_w

    name_col_x = left + left_col_w
    qty_col_x  = right - qty_col_w

    # Row heights (tighter to fit half A4)
    row1_h = 20*mm
    row2_h = 20*mm
    total_h = row1_h + row2_h

    # --- Header ---
    if LOGO_PATH and os.path.exists(LOGO_PATH):
        try:
            c.drawImage(LOGO_PATH, left, y-12*mm, width=18*mm, height=18*mm,
                        preserveAspectRatio=True, mask='auto')
            text_x = left + 20*mm
        except Exception:
            text_x = left
    else:
        text_x = left

    c.setFont("Helvetica-Bold", 14)
    c.drawString(text_x, y, SHOP_NAME)
    c.setFont("Helvetica", 9.2)
    c.drawString(text_x, y - 5.0*mm, SHOP_ADDR)
    c.drawString(text_x, y - 9.0*mm, SHOP_TEL)

    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString((left+right)/2, y - 16.0*mm, "SERVICE VOUCHER")
    c.drawRightString(right, y - 16.0*mm, f"No : {voucher_id}")

    c.setFont("Helvetica", 10)
    c.drawString(left, y - 24.0*mm, "Date :")
    c.drawString(left + 18*mm, y - 24.0*mm, created_at[:10])
    c.drawRightString(right - 27*mm, y - 24.0*mm, "Time In :")
    c.drawRightString(right, y - 24.0*mm, created_at[11:19])

    # --- Table frame ---
    top_table = y - 28*mm
    bottom_table = top_table - total_h

    # outer border
    c.rect(left, bottom_table, right-left, total_h, stroke=1, fill=0)
    # verticals
    c.line(name_col_x, top_table, name_col_x, bottom_table)
    c.line(qty_col_x,  top_table, qty_col_x,  bottom_table)
    # single horizontal split for left+middle only
    mid_y = top_table - row1_h
    c.line(left, mid_y, qty_col_x, mid_y)

    pad = 3*mm

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

    # QTY (spans both rows)
    c.setFont("Helvetica-Bold", 10.4)
    c.drawCentredString(qty_col_x + qty_col_w/2, top_table - pad - 8, "QTY")
    c.setFont("Helvetica", 11)
    c.drawCentredString(qty_col_x + qty_col_w/2, bottom_table + total_h/2 - 3, str(units or 1))

    # Row 2 Left: TEL
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left + pad, mid_y - pad - 8, "TEL")
    draw_wrapped(c, contact_number,
                 left + pad, bottom_table + pad,
                 w=left_col_w - 2*pad, h=row2_h - 2*pad - 10, fontsize=10)

    # Row 2 Middle: PROBLEM
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(name_col_x + pad, mid_y - pad - 8, "PROBLEM")
    draw_wrapped(c, problem or "-",
                 name_col_x + pad, bottom_table + pad,
                 w=middle_col_w - 2*pad, h=row2_h - 2*pad - 10, fontsize=10)

    # ---- Recipient line + ack sentence ----
    y_rec = bottom_table - 10.5*mm
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left, y_rec, "RECIPIENT :")
    c.line(left + 24*mm, y_rec, left + 110*mm, y_rec)  # signature line

    draw_wrapped(
        c,
        "CUSTOMER ACKNOWLEDGE WE HEREBY CONFIRMED THAT THE MACHINE WAS SERVICE AND REPAIRED SATISFACTORILY",
        left + 118*mm, y_rec - 14,
        right - (left + 118*mm), 18*mm, fontsize=8.9, leading=10
    )

    # ---- Remark outside table ----
    y_remark = y_rec - 11*mm
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left, y_remark, "REMARK:")
    draw_wrapped(c, remark or "-", left + 22*mm, y_remark - 10,
                 w=(right - (left + 22*mm)), h=18*mm, fontsize=10)

    # ---- Signatures (same line) ----
    y_sig = y_remark - 20*mm
    c.line(left, y_sig, left + 60*mm, y_sig)
    c.setFont("Helvetica", 9)
    c.drawString(left, y_sig - 4*mm, "CUSTOMER SIGNATURE")

    c.line(right - 60*mm, y_sig, right, y_sig)
    c.drawString(right - 60*mm, y_sig - 4*mm, "DATE COLLECTED")

    # ---- Disclaimers + QR (tight, tidy, fits half A4) ----
    qr_size = 20*mm
    disc_left = left
    y_disc = y_sig - 10*mm

    c.setFont("Helvetica", 8.5)
    c.drawString(disc_left, y_disc, "Kindly collect your goods within 60 days from date of sending for repair.")
    c.drawString(disc_left, y_disc - 12, "A) We do not hold ourselves responsible for any loss or damage.")
    c.drawString(disc_left, y_disc - 24, "B) We reserve our right to sell off the goods to cover our cost and loss.")

    # Highlighted RM60.00 in red (split into two rows, smaller font)
    text1 = "MINIMUM "
    text2 = "RM60.00"
    text3a = " WILL BE CHARGED ON TROUBLESHOOTING, INSPECTION AND SERVICE"
    text3b = "ON ALL KIND OF HARDWARE AND SOFTWARE."

    # First row
    c.setFont("Helvetica", 7.5)
    c.drawString(disc_left, y_disc - 40, text1)
    c.setFillColorRGB(1, 0, 0)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(disc_left + c.stringWidth(text1, "Helvetica", 7.5), y_disc - 40, text2)
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica", 7.5)
    c.drawString(disc_left + c.stringWidth(text1, "Helvetica", 7.5) + c.stringWidth(text2, "Helvetica-Bold", 9),
                 y_disc - 40, text3a)

    # Second row
    c.setFont("Helvetica", 7.5)
    c.drawString(disc_left, y_disc - 52, text3b)

    # Follow-up disclaimers
    c.drawString(disc_left, y_disc - 66, "PLEASE BRING ALONG THIS SERVICE VOUCHER TO COLLECT YOUR GOODS")
    c.drawString(disc_left, y_disc - 78, "NO ATTENTION GIVEN WITHOUT SERVICE VOUCHER")


    # QR code (in-memory, no temp files)
    try:
        qr_data = f"Voucher:{voucher_id}|Name:{customer_name}|Tel:{contact_number}|Date:{created_at[:10]}"
        qr_img  = qrcode.make(qr_data)
        c.drawImage(ImageReader(qr_img), right - qr_size, y_disc - 50, qr_size, qr_size)
    except Exception:
        pass


def generate_pdf(voucher_id, customer_name, contact_number, units, remark,
                 particulars, problem, staff_name, status, created_at):
    filename = os.path.join(PDF_DIR, f"voucher_{voucher_id}.pdf")
    c = rl_canvas.Canvas(filename, pagesize=A4)
    width, height = A4
    _draw_voucher(c, width, height, voucher_id, customer_name, contact_number,
                  units, remark, particulars, problem, staff_name, created_at)
    c.showPage()
    c.save()
    return filename

# ------------------ DB ops ------------------
def add_voucher(customer_name, contact_number, units, remark, particulars, problem, staff_name):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    voucher_id = next_voucher_id()
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    status = "Pending"

    pdf_path = generate_pdf(
        voucher_id, customer_name, contact_number, units, remark,
        particulars, problem, staff_name, status, created_at
    )

    cur.execute("""
        INSERT INTO vouchers (voucher_id, created_at, customer_name, contact_number, units,
                              remark, particulars, problem, staff_name, status, pdf_path)
        VALUES (?,?,?,?,?,?,?,?,?,?,?)
    """, (voucher_id, created_at, customer_name, contact_number, units,
          remark, particulars, problem, staff_name, status, pdf_path))
    conn.commit()
    conn.close()
    return voucher_id, pdf_path

def mark_collected(voucher_id):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("UPDATE vouchers SET status='Collected' WHERE voucher_id = ?", (voucher_id,))
    conn.commit()
    conn.close()

def search_vouchers(filters):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    sql = ("SELECT voucher_id, created_at, customer_name, contact_number, units, status, remark, pdf_path "
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
    sql += " ORDER BY created_at DESC"
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
        self.minsize(800, 500)

        # ----- Scrollable main container -----
        outer = ctk.CTkFrame(self)
        outer.pack(fill="both", expand=True)

        # Canvas + scrollbars (h: manual, v: mouse wheel anywhere)
        self.uicanvas = tk.Canvas(outer, highlightthickness=0)
        self.uicanvas.pack(side="left", fill="both", expand=True)

        vbar = ttk.Scrollbar(outer, orient="vertical", command=self.uicanvas.yview)
        vbar.pack(side="right", fill="y")
        hbar = ttk.Scrollbar(self, orient="horizontal", command=self.uicanvas.xview)
        hbar.pack(side="bottom", fill="x")

        self.uicanvas.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)

        self.scrollable_frame = ctk.CTkFrame(self.uicanvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.uicanvas.configure(scrollregion=self.uicanvas.bbox("all"))
        )
        self.uicanvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # Global vertical scrolling (no need to hover the bar)
        def _on_mousewheel(event):
            # Windows/macOS: event.delta is +/- 120-ish; Linux uses Button-4/5 below
            delta = int(-1 * (event.delta / 120)) if event.delta else 0
            self.uicanvas.yview_scroll(delta, "units")
        self.bind_all("<MouseWheel>", _on_mousewheel)            # Windows/macOS
        self.bind_all("<Button-4>", lambda e: self.uicanvas.yview_scroll(-1, "units"))  # Linux up
        self.bind_all("<Button-5>", lambda e: self.uicanvas.yview_scroll(1, "units"))   # Linux down

        # ---- Filters ----
        filt = ctk.CTkFrame(self.scrollable_frame)
        filt.pack(fill="x", padx=10, pady=(10, 4))
        today = datetime.now().strftime("%Y-%m-%d")

        self.f_voucher = ctk.CTkEntry(filt, width=130, placeholder_text="VoucherID")
        self.f_name    = ctk.CTkEntry(filt, width=220, placeholder_text="Customer Name")
        self.f_contact = ctk.CTkEntry(filt, width=180, placeholder_text="Contact Number")
        self.f_from    = ctk.CTkEntry(filt, width=160, placeholder_text="Date From (YYYY-MM-DD)")
        self.f_to      = ctk.CTkEntry(filt, width=160, placeholder_text="Date To (YYYY-MM-DD)")
        self.f_from.insert(0, today)
        self.f_to.insert(0, today)

        self.f_status  = ctk.CTkOptionMenu(filt, values=["All","Pending","Collected"], width=120)
        self.f_status.set("All")
        self.btn_search = ctk.CTkButton(filt, text="Search", command=self.perform_search, width=100)
        self.btn_reset  = ctk.CTkButton(filt, text="Reset",  command=self.reset_filters, width=80)

        for w in (self.f_voucher, self.f_name, self.f_contact, self.f_from, self.f_to, self.f_status,
                  self.btn_search, self.btn_reset):
            w.pack(side="left", padx=5, pady=8)

        # ---- Table ----
        self.tree = ttk.Treeview(self.scrollable_frame,
            columns=("VoucherID","Date","Customer","Contact","Units","Status","Remark","PDF"),
            show="headings", height=18)

        widths = {"VoucherID":120, "Date":160, "Customer":220, "Contact":160,
                  "Units":70, "Status":100, "Remark":300, "PDF":0}
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=widths[col], anchor="w", stretch=(col != "PDF"))

        self.tree.pack(expand=True, fill="both", padx=10, pady=(4, 8))
        self.tree.bind("<Double-1>", lambda e: self.open_pdf())


        # ---- Actions ----
        bar = ctk.CTkFrame(self.scrollable_frame)
        bar.pack(fill="x", padx=10, pady=(0,10))
        ctk.CTkButton(bar, text="Add Voucher",       command=self.add_voucher_ui).pack(side="left", padx=8, pady=8)
        ctk.CTkButton(bar, text="Edit Selected",     command=self.edit_selected).pack(side="left", padx=8, pady=8)
        ctk.CTkButton(bar, text="Mark as Collected", command=self.mark_selected).pack(side="left", padx=8, pady=8)
        ctk.CTkButton(bar, text="Open PDF",          command=self.open_pdf).pack(side="left", padx=8, pady=8)
        ctk.CTkButton(bar, text="Manage Staffs",     command=self.manage_staffs_ui).pack(side="left", padx=8, pady=8)

        self.perform_search()

    # ---- Filters helpers ----
    def _get_filters(self):
        return {
            "voucher_id": self.f_voucher.get().strip(),
            "name": self.f_name.get().strip(),
            "contact": self.f_contact.get().strip(),
            "date_from": self.f_from.get().strip(),
            "date_to": self.f_to.get().strip(),
            "status": self.f_status.get(),
        }

    def reset_filters(self):
        for e in (self.f_voucher, self.f_name, self.f_contact, self.f_from, self.f_to):
            e.delete(0, "end")
        today = datetime.now().strftime("%Y-%m-%d")
        self.f_from.insert(0, today)
        self.f_to.insert(0, today)
        self.f_status.set("All")
        self.perform_search()

    def perform_search(self):
        rows = search_vouchers(self._get_filters())
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("", "end", values=r)

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

        ctk.CTkLabel(frm, text="Staff Name", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        staff_values = list_staffs() or [""]
        e_staff = ctk.CTkComboBox(frm, values=staff_values, width=WIDE)
        if staff_values and staff_values[0] != "":
            e_staff.set(staff_values[0])
        e_staff.grid(row=r, column=1, sticky="w", padx=10, pady=(0,8)); r+=1

        ctk.CTkLabel(frm, text="Remark", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        t_remark = tk.Text(frm, width=66, height=4); t_remark.grid(row=r, column=1, sticky="w", padx=10, pady=(0,8))

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
            remark      = t_remark.get("1.0","end").strip()
            staff_name  = e_staff.get().strip()
            if not name or not contact:
                messagebox.showerror("Missing", "Customer name and contact are required."); return
            voucher_id, pdf_path = add_voucher(name, contact, units, remark, particulars, problem, staff_name)
            messagebox.showinfo("Saved", f"Voucher {voucher_id} created.")
            try:
                webbrowser.open_new(os.path.abspath(pdf_path))
            except Exception:
                pass
            top.destroy()
            self.perform_search()

        ctk.CTkButton(btns, text="Save & Open PDF", command=save, width=180).pack(side="right")

    # ---- Manage Staffs (tidy) ----
    def manage_staffs_ui(self):
        top = ctk.CTkToplevel(self)
        top.title("Manage Staffs")
        top.geometry("680x480")
        top.grab_set()

        root = ctk.CTkFrame(top)
        root.pack(fill="both", expand=True, padx=14, pady=14)
        root.grid_rowconfigure(2, weight=1)
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=0)

        entry = ctk.CTkEntry(root, placeholder_text="New staff name")
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
                   particulars, problem, remark, staff_name, status
            FROM vouchers WHERE voucher_id = ?
        """, (voucher_id,))
        row = cur.fetchone()
        conn.close()
        if not row: return

        _, created_at, customer_name, contact_number, units, particulars, problem, remark, staff_name, status = row

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

        ctk.CTkLabel(frm, text="Staff").grid(row=r, column=0, sticky="w")
        staff_values = list_staffs() or [""]
        e_staff = ctk.CTkComboBox(frm, values=staff_values, width=WIDE)
        e_staff.set(staff_name or "")
        e_staff.grid(row=r, column=1, sticky="w", padx=10, pady=6); r+=1

        ctk.CTkLabel(frm, text="Remark").grid(row=r, column=0, sticky="w")
        t_remark = tk.Text(frm, width=66, height=4)
        t_remark.insert("1.0", remark or "")
        t_remark.grid(row=r, column=1, sticky="w", padx=10, pady=6)

        def save_edit():
            name = e_name.get().strip()
            contact = e_contact.get().strip()
            try:
                units_val = int((e_units.get() or "1").strip())
            except ValueError:
                messagebox.showerror("Invalid", "Units must be a number."); return
            particulars_val = t_part.get("1.0","end").strip()
            problem_val     = t_prob.get("1.0","end").strip()
            remark_val      = t_remark.get("1.0","end").strip()
            staff_val       = e_staff.get().strip()

            # update DB
            conn = sqlite3.connect(DB_FILE)
            cur = conn.cursor()
            cur.execute("""UPDATE vouchers
                              SET customer_name=?, contact_number=?, units=?, particulars=?,
                                  problem=?, remark=?, staff_name=?
                            WHERE voucher_id=?""",
                        (name, contact, units_val, particulars_val,
                         problem_val, remark_val, staff_val, voucher_id))
            conn.commit()

            # regenerate PDF to reflect edits
            pdf_path = generate_pdf(
                voucher_id, name, contact, units_val, remark_val,
                particulars_val, problem_val, staff_val, status, created_at
            )
            cur.execute("UPDATE vouchers SET pdf_path=? WHERE voucher_id=?", (pdf_path, voucher_id))
            conn.commit()
            conn.close()

            messagebox.showinfo("Updated", f"Voucher {voucher_id} updated.")
            top.destroy()
            self.perform_search()

        btns = ctk.CTkFrame(top); btns.pack(fill="x", padx=12, pady=12)
        ctk.CTkButton(btns, text="Save Changes", command=save_edit, width=160).pack(side="right", padx=6)
        ctk.CTkButton(btns, text="Cancel", command=top.destroy, width=100).pack(side="right", padx=6)


# ------------------ Run ------------------
if __name__ == "__main__":
    init_db()
    app = VoucherApp()
    app.mainloop()
