#!/usr/bin/env python3
# Service Voucher Management System (Monolith)
# Deps: customtkinter, reportlab, qrcode, pillow, pandas, openpyxl
# Run:  python main.py

import os
import sqlite3
import webbrowser
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import pandas as pd
import qrcode
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

# ------------------ Config ------------------
DB_FILE   = "vouchers.db"
PDF_DIR   = "pdfs"
os.makedirs(PDF_DIR, exist_ok=True)

# Header text (edit to your shop details)
SHOP_NAME = "TONY.COM"
SHOP_ADDR = "TB4318, Lot 5, Block 31, Fajar Complex  91000 Tawau Sabah, Malaysia"
SHOP_TEL  = "Tel : 089-763778, H/P: 0168260533"
LOGO_PATH = ""  # optional: path to logo image (png/jpg). leave blank to skip

# When DB is empty, first voucher number will start here (e.g., 41000)
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
"""

def _column_exists(cur, table, column):
    cur.execute(f"PRAGMA table_info({table})")
    return any(row[1] == column for row in cur.fetchall())

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.executescript(BASE_SCHEMA)
    # Back-compat safety (in case table exists without some cols)
    wanted = [
        ("voucher_id","TEXT"), ("created_at","TEXT"), ("customer_name","TEXT"),
        ("contact_number","TEXT"), ("units","INTEGER"), ("remark","TEXT"),
        ("particulars","TEXT"), ("problem","TEXT"), ("staff_name","TEXT"),
        ("status","TEXT"), ("pdf_path","TEXT")
    ]
    for col, typ in wanted:
        if not _column_exists(cur, "vouchers", col):
            cur.execute(f"ALTER TABLE vouchers ADD COLUMN {col} {typ}")
    try:
        cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_vouchers_vid ON vouchers(voucher_id)")
    except Exception:
        pass
    conn.commit()
    conn.close()

def next_voucher_id():
    """Return next numeric voucher id (41000, 41001, ...)."""
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT MAX(CAST(voucher_id AS INTEGER)) FROM vouchers")
    row = cur.fetchone()
    conn.close()
    if not row or row[0] is None:
        return str(BASE_VOUCHER_NO)
    return str(int(row[0]) + 1)

# ------------------ PDF ------------------
def _draw_voucher_half(c, width, top_y, voucher_id, customer_name, contact_number,
                       units, remark, particulars, problem, staff_name, created_at):
    """
    Draws one voucher (half page). top_y is the top baseline to start from.
    Returns a y-end reference (not used downstream but kept for clarity).
    """
    left   = 12*mm
    right  = width - 12*mm
    y      = top_y
    lh     = 4.2*mm   # label line height
    row_h  = 9.5*mm   # table row height

    # --- Header (compact, spaced) ---
    # Optional logo
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
    c.drawString(text_x,       y - 5.0*mm, SHOP_ADDR)
    c.drawString(text_x,       y - 9.0*mm, SHOP_TEL)

    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString((left+right)/2, y - 16.0*mm, "SERVICE VOUCHER")
    c.drawRightString(right,            y - 16.0*mm, f"No : {voucher_id}")

    # Date / Time
    c.setFont("Helvetica", 10)
    c.drawString(left,               y - 24.0*mm, "Date :")
    c.drawString(left + 18*mm,       y - 24.0*mm, created_at[:10])
    c.drawRightString(right - 27*mm, y - 24.0*mm, "Time In :")
    c.drawRightString(right,         y - 24.0*mm, created_at[11:19])

    # --- Main grid: Customer / Particulars / Qty ---
    top_table = y - 28.0*mm
    # outer rectangle (3 rows)
    c.rect(left, top_table - row_h*3, right-left, row_h*3, stroke=1, fill=0)

    # vertical splits
    name_col_x = left + 58*mm
    qty_col_x  = right - 20*mm
    c.line(name_col_x, top_table,           name_col_x, top_table - row_h*3)
    c.line(qty_col_x,  top_table,           qty_col_x,  top_table - row_h*3)

    # horizontal inner lines
    c.line(left,        top_table - row_h,  right, top_table - row_h)
    c.line(left,        top_table - row_h*2,right, top_table - row_h*2)

    # headers (no collisions)
    c.setFont("Helvetica-Bold", 10.2)
    c.drawString(left+2*mm,                 top_table - row_h + lh, "CUSTOMER NAME")
    c.drawCentredString((name_col_x+qty_col_x)/2, top_table - row_h + lh, "PARTICULARS")
    c.drawCentredString((qty_col_x+right)/2,      top_table - row_h + lh, "QTY")
    c.drawString(left+2*mm,                 top_table - row_h*2 + lh, "TEL:")

    # values
    c.setFont("Helvetica", 10)
    c.drawString(left+34*mm,                top_table - row_h + lh,   customer_name)
    c.drawString(left+12*mm,                top_table - row_h*2 + lh, contact_number)
    c.drawString(name_col_x + 3*mm,         top_table - row_h*2 + lh, (particulars or "-"))
    c.drawCentredString((qty_col_x+right)/2, top_table - row_h*2 + lh, str(units or 1))

    # problem row
    y2 = top_table - row_h*3 - 6
    c.setFont("Helvetica-Bold", 10.2); c.drawString(left, y2, "PROBLEM:")
    c.setFont("Helvetica", 10);        c.drawString(left+24*mm, y2, (problem or "-"))

    # recipient / staff
    y3 = y2 - 8.5*mm
    c.setFont("Helvetica-Bold", 10.2); c.drawString(left, y3, "RECIPIENT :")
    c.setFont("Helvetica-Bold", 10.2); c.drawString(left+70*mm, y3, "NAME OF STAFF :")
    c.setFont("Helvetica", 10);        c.drawString(left+110*mm, y3, (staff_name or "-"))

    # remark
    y4 = y3 - 8.5*mm
    c.setFont("Helvetica-Bold", 10.2); c.drawString(left, y4, "REMARK:")
    c.setFont("Helvetica", 10);        c.drawString(left+20*mm, y4, (remark or "-"))

    # signatures (closer to the above block)
    y5 = y4 - 11*mm
    c.line(left,          y5, left+60*mm, y5)
    c.setFont("Helvetica", 9); c.drawString(left, y5 - 4*mm, "CUSTOMER SIGNATURE")
    c.line(right-60*mm,   y5, right,      y5)
    c.drawString(right-60*mm, y5 - 4*mm, "DATE COLLECTED")

    # warnings (just below signatures)
    y6 = y5 - 7.5*mm
    c.setFont("Helvetica", 8.5)
    c.drawString(left, y6,            "* Kindly collect your goods within 60 days of sending for repair.")
    c.drawString(left, y6 - 4*mm,     "A) We do not hold ourselves responsible for any loss or damage.")
    c.drawString(left, y6 - 8*mm,     "B) We reserve our right to sell off the goods to cover our cost and loss.")
    c.drawString(left, y6 - 12*mm,    "* MINIMUM RM45.00 WILL BE CHARGED ON TROUBLESHOOTING / INSPECTION / SERVICE.")

    # QR (kept away from lines)
    try:
        qr_size = 18*mm
        qr_data = f"Voucher:{voucher_id}|Name:{customer_name}|Tel:{contact_number}|Date:{created_at[:10]}"
        qr_img  = qrcode.make(qr_data)
        qr_path = os.path.join(PDF_DIR, f"qr_{voucher_id}.png")
        qr_img.save(qr_path)
        c.drawImage(qr_path, right - qr_size, top_table + 2*mm, qr_size, qr_size)
        os.remove(qr_path)
    except Exception:
        pass

    return y6 - 14*mm

def generate_pdf(voucher_id, customer_name, contact_number, units, remark,
                 particulars, problem, staff_name, status, created_at):
    """A4 page with TWO vouchers (top & bottom) and a dashed cut line."""
    filename = os.path.join(PDF_DIR, f"voucher_{voucher_id}.pdf")
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    top_margin = 10*mm
    mid_y      = height / 2.0

    # Top half
    _draw_voucher_half(
        c, width, height - top_margin,
        voucher_id, customer_name, contact_number,
        units, remark, particulars, problem, staff_name, created_at
    )

    # dashed cut line
    c.setDash(6, 3)
    c.line(12*mm, mid_y, width - 12*mm, mid_y)
    c.setDash()

    # Bottom half
    _draw_voucher_half(
        c, width, mid_y - 6*mm,
        voucher_id, customer_name, contact_number,
        units, remark, particulars, problem, staff_name, created_at
    )

    c.showPage()
    c.save()
    return filename

# ------------------ DB operations ------------------
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
    """Return list of tuples for the table."""
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

def export_filtered(filters, filename):
    rows = search_vouchers(filters)
    df = pd.DataFrame(rows, columns=["VoucherID","Date","Customer","Contact","Units","Status","Remark","PDF"])
    if filename.endswith(".csv"):
        df.to_csv(filename, index=False)
    else:
        df.to_excel(filename, index=False)

# ------------------ UI ------------------
class VoucherApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Service Voucher Management System")
        self.geometry("1100x650")
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        # ---- Filters row ----
        filt = ctk.CTkFrame(self)
        filt.pack(fill="x", padx=10, pady=(10, 4))

        self.f_voucher = ctk.CTkEntry(filt, width=130, placeholder_text="VoucherID")
        self.f_name    = ctk.CTkEntry(filt, width=200, placeholder_text="Customer Name")
        self.f_contact = ctk.CTkEntry(filt, width=160, placeholder_text="Contact Number")
        self.f_from    = ctk.CTkEntry(filt, width=150, placeholder_text="Date From (YYYY-MM-DD)")
        self.f_to      = ctk.CTkEntry(filt, width=150, placeholder_text="Date To (YYYY-MM-DD)")
        self.f_status  = ctk.CTkOptionMenu(filt, values=["All","Pending","Collected"], width=120)
        self.f_status.set("All")
        self.btn_search = ctk.CTkButton(filt, text="Search", command=self.perform_search, width=100)
        self.btn_reset  = ctk.CTkButton(filt, text="Reset",  command=self.reset_filters, width=80)

        for w in (self.f_voucher, self.f_name, self.f_contact, self.f_from, self.f_to, self.f_status,
                  self.btn_search, self.btn_reset):
            w.pack(side="left", padx=5, pady=8)

        # ---- Table ----
        self.tree = ttk.Treeview(self,
            columns=("VoucherID","Date","Customer","Contact","Units","Status","Remark","PDF"),
            show="headings", height=18)
        widths = {"VoucherID":120, "Date":160, "Customer":200, "Contact":140, "Units":70, "Status":100, "Remark":300, "PDF":0}
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=widths[col], anchor="w", stretch=(col != "PDF"))
        self.tree.pack(expand=True, fill="both", padx=10, pady=(4, 8))
        self.tree.bind("<Double-1>", lambda e: self.open_pdf())

        # ---- Action bar ----
        bar = ctk.CTkFrame(self)
        bar.pack(fill="x", padx=10, pady=(0,10))
        ctk.CTkButton(bar, text="Add Voucher",        command=self.add_voucher_ui).pack(side="left", padx=8, pady=8)
        ctk.CTkButton(bar, text="Mark as Collected",  command=self.mark_selected).pack(side="left", padx=8, pady=8)
        ctk.CTkButton(bar, text="Open PDF",           command=self.open_pdf).pack(side="left", padx=8, pady=8)
        ctk.CTkButton(bar, text="Export (CSV/XLSX)",  command=self.export_data).pack(side="left", padx=8, pady=8)

        self.perform_search()

    # ---- Filters ----
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
        top.geometry("760x560")     # bigger window
        top.grab_set()

        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        r = 0
        ctk.CTkLabel(frm, text="Customer Name", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        e_name = ctk.CTkEntry(frm, width=280); e_name.grid(row=r, column=1, sticky="w", padx=10, pady=(0,6)); r+=1

        ctk.CTkLabel(frm, text="Contact Number", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        e_contact = ctk.CTkEntry(frm, width=280); e_contact.grid(row=r, column=1, sticky="w", padx=10, pady=(0,6)); r+=1

        ctk.CTkLabel(frm, text="No. of Units", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        e_units = ctk.CTkEntry(frm, width=120); e_units.insert(0,"1")
        e_units.grid(row=r, column=1, sticky="w", padx=10, pady=(0,6)); r+=1

        ctk.CTkLabel(frm, text="Particulars", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        t_part = tk.Text(frm, width=46, height=4); t_part.grid(row=r, column=1, sticky="w", padx=10, pady=(0,6)); r+=1

        ctk.CTkLabel(frm, text="Problem", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        t_prob = tk.Text(frm, width=46, height=4); t_prob.grid(row=r, column=1, sticky="w", padx=10, pady=(0,6)); r+=1

        ctk.CTkLabel(frm, text="Staff Name", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        e_staff = ctk.CTkEntry(frm, width=280); e_staff.grid(row=r, column=1, sticky="w", padx=10, pady=(0,6)); r+=1

        ctk.CTkLabel(frm, text="Remark", anchor="w").grid(row=r, column=0, sticky="w", pady=(0,2))
        t_remark = tk.Text(frm, width=46, height=4); t_remark.grid(row=r, column=1, sticky="w", padx=10, pady=(0,6))

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
            if not name or not contact:
                messagebox.showerror("Missing", "Customer name and contact are required."); return
            voucher_id, pdf_path = add_voucher(name, contact, units, remark, particulars, problem, e_staff.get().strip())
            messagebox.showinfo("Saved", f"Voucher {voucher_id} created.")
            try:
                webbrowser.open_new(os.path.abspath(pdf_path))
            except Exception:
                pass
            top.destroy()
            self.perform_search()

        ctk.CTkButton(btns, text="Save & Open PDF", command=save, width=160).pack(side="right")

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
                os.startfile(pdf_path)
            else:
                os.system(f"open '{pdf_path}'")

    def export_data(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files","*.csv"), ("Excel Files","*.xlsx")]
        )
        if not filename: return
        export_filtered(self._get_filters(), filename)
        messagebox.showinfo("Exported", f"Exported to {filename}")

# ------------------ Run ------------------
if __name__ == "__main__":
    init_db()
    app = VoucherApp()
    app.mainloop()
