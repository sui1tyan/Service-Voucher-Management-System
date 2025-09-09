#!/usr/bin/env python3
# Service Voucher Management System (Monolith, enhanced)
# Deps: customtkinter, reportlab, qrcode, pillow, pandas, openpyxl

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
SHOP_NAME = "TONY.COM"
SHOP_ADDR = "TB4318, Lot 5, Block 31, Fajar Complex\n91000 Tawau Sabah, Malaysia"
SHOP_TEL  = "Tel : 089-763778, H/P: 0168260533"
LOGO_PATH = ""  # put a path to a PNG/JPG logo if you have one (optional)

os.makedirs(PDF_DIR, exist_ok=True)

# ------------------ DB ------------------
BASE_SCHEMA = """
CREATE TABLE IF NOT EXISTS vouchers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    voucher_id TEXT,
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

    # Backward-compat: add missing columns if user already has an old DB
    wanted = [
        ("voucher_id", "TEXT"),
        ("created_at", "TEXT"),
        ("customer_name", "TEXT"),
        ("contact_number", "TEXT"),
        ("units", "INTEGER DEFAULT 1"),
        ("remark", "TEXT"),
        ("particulars", "TEXT"),
        ("problem", "TEXT"),
        ("staff_name", "TEXT"),
        ("status", "TEXT DEFAULT 'Pending'"),
        ("pdf_path", "TEXT"),
    ]
    for col, coltype in wanted:
        if not _column_exists(cur, "vouchers", col):
            cur.execute(f"ALTER TABLE vouchers ADD COLUMN {col} {coltype}")
    conn.commit()
    conn.close()

# ------------------ PDF ------------------
def _draw_header(c, width, height, voucher_id):
    top_y = height - 30*mm
    left_x = 15*mm
    right_x = width - 15*mm

    # Logo (optional)
    if LOGO_PATH and os.path.exists(LOGO_PATH):
        try:
            c.drawImage(LOGO_PATH, left_x, top_y-10*mm, width=20*mm, height=20*mm, preserveAspectRatio=True, mask='auto')
            text_x = left_x + 25*mm
        except Exception:
            text_x = left_x
    else:
        text_x = left_x

    # Shop info
    c.setFont("Helvetica-Bold", 18)
    c.drawString(text_x, top_y, SHOP_NAME)
    c.setFont("Helvetica", 10)
    c.drawString(text_x, top_y - 6*mm, SHOP_ADDR)
    c.drawString(text_x, top_y - 11*mm, SHOP_TEL)

    # Title
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width/2, top_y - 20*mm, "SERVICE VOUCHER")

    # Voucher No: big on right
    c.setFont("Helvetica-Bold", 16)
    c.drawRightString(right_x, top_y - 20*mm, f"No : {voucher_id}")

def _draw_grid(c, width, height, fields):
    """
    fields expects dict with:
    date, time_in, customer_name, contact_number, particulars, units, problem, staff_name, remark
    """
    left = 15*mm
    right = width - 15*mm
    top = height - 60*mm
    row_h = 10*mm

    # Outer rectangle
    c.rect(left, top - 4*row_h - 20*mm, right-left, 4*row_h + 20*mm, stroke=1, fill=0)

    # First row: Date / Time In
    c.setFont("Helvetica-Bold", 11)
    c.drawString(left+2*mm, top, "Date :")
    c.setFont("Helvetica", 11)
    c.drawString(left+22*mm, top, fields["date"])

    c.setFont("Helvetica-Bold", 11)
    c.drawString((right-35*mm), top, "Time In :")
    c.setFont("Helvetica", 11)
    c.drawString((right-20*mm), top, fields["time_in"])

    # Big table: Customer/Particulars/Qty
    # Draw vertical splits
    y = top - row_h
    c.line(left, y, right, y)
    c.line(left + 55*mm, y, left + 55*mm, y - 3*row_h)   # Customer Name column
    c.line(right - 25*mm, y, right - 25*mm, y - 3*row_h) # QTY column
    # Horizontal internal lines
    for i in range(1, 3):
        c.line(left + 55*mm, y - i*row_h, right, y - i*row_h)

    # Left label
    c.setFont("Helvetica-Bold", 11)
    c.drawString(left+2*mm, y - 0.7*row_h, "CUSTOMER NAME")
    c.setFont("Helvetica", 11)
    c.drawString(left+2*mm, y - 1.7*row_h, "TEL:")

    # Right labels
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString((left+55*mm + right-25*mm)/2, y - 0.7*row_h, "PARTICULARS")
    c.drawCentredString(right - 12.5*mm, y - 0.7*row_h, "QTY")

    # Fill values
    c.setFont("Helvetica", 11)
    c.drawString(left+58*mm, y - 1.7*row_h, fields["particulars"])
    c.drawCentredString(right - 12.5*mm, y - 1.7*row_h, str(fields["units"]))

    c.drawString(left+25*mm, y - 0.7*row_h, fields["customer_name"])
    c.drawString(left+12*mm, y - 1.7*row_h, fields["contact_number"])

    # Second block: PROBLEM
    y2 = y - 3*row_h
    c.line(left, y2, right, y2)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(left+2*mm, y2 - 0.7*row_h, "PROBLEM:")
    c.setFont("Helvetica", 11)
    c.drawString(left+25*mm, y2 - 0.7*row_h, fields["problem"])

    # Recipient / Staff
    y3 = y2 - row_h
    c.line(left, y3, right, y3)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(left+2*mm, y3 - 0.7*row_h, "RECIPIENT :")
    c.setFont("Helvetica-Bold", 11)
    c.drawString(left+65*mm, y3 - 0.7*row_h, "NAME OF STAFF :")
    c.setFont("Helvetica", 11)
    c.drawString(left+95*mm, y3 - 0.7*row_h, fields["staff_name"])

    # Footer notes and signatures
    bottom = 25*mm
    c.setFont("Helvetica", 9)
    c.drawString(left, bottom + 12*mm,
                 "* Kindly collect your goods within 60 days from date of sending for repair.")
    c.drawString(left, bottom + 8*mm,
                 "A) We do not hold ourselves responsible for any loss or damage.")
    c.drawString(left, bottom + 4*mm,
                 "B) We reserve our right to sell off the goods to cover our cost and loss.")

    c.drawString(left, bottom,
                 "* MINIMUM RM45.00 WILL BE CHARGED ON TROUBLESHOOTING, INSPECTION AND SERVICE "
                 "ON ALL KIND OF HARDWARE AND SOFTWARE.")

    # Signatures
    c.line(left, 15*mm, left+60*mm, 15*mm)
    c.drawString(left, 10*mm, "CUSTOMER SIGNATURE")

    c.line(right-60*mm, 15*mm, right, 15*mm)
    c.drawString(right-60*mm, 10*mm, "DATE COLLECTED")

    # Remark box
    c.setFont("Helvetica-Bold", 11)
    c.drawString(left, 32*mm, "REMARK:")
    c.setFont("Helvetica", 11)
    c.drawString(left+25*mm, 32*mm, fields["remark"])

def generate_pdf(voucher_id, customer_name, contact_number, units, remark,
                 particulars, problem, staff_name, status, created_at):
    filename = os.path.join(PDF_DIR, f"voucher_{voucher_id}.pdf")
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    _draw_header(c, width, height, voucher_id)

    fields = {
        "date": created_at[:10],
        "time_in": created_at[11:19],
        "customer_name": customer_name,
        "contact_number": contact_number,
        "particulars": particulars or "-",
        "units": units or 1,
        "problem": problem or "-",
        "staff_name": staff_name or "-",
        "remark": remark or "-",
    }
    _draw_grid(c, width, height, fields)

    # QR (encodes quick info)
    qr_data = f"Voucher:{voucher_id}\nName:{customer_name}\nContact:{contact_number}\nDate:{fields['date']}"
    qr_img = qrcode.make(qr_data)
    qr_path = os.path.join(PDF_DIR, f"qr_{voucher_id}.png")
    qr_img.save(qr_path)
    c.drawImage(qr_path, width - 35*mm, height - 80*mm, width=20*mm, height=20*mm)

    c.showPage()
    c.save()
    try:
        os.remove(qr_path)
    except Exception:
        pass
    return filename

# ------------------ DB operations ------------------
def add_voucher(customer_name, contact_number, units, remark, particulars, problem, staff_name):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    voucher_id = f"{int(datetime.now().timestamp())}"[1:]  # numeric-ish, looks like 41000+ style
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
    """
    filters: dict keys voucher_id, name, contact, date_from, date_to, status
    Returns rows with columns used in the table.
    """
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    base = ("SELECT voucher_id, created_at, customer_name, contact_number, units, status, remark, pdf_path "
            "FROM vouchers WHERE 1=1")
    params = []

    if filters.get("voucher_id"):
        base += " AND voucher_id LIKE ?"
        params.append(f"%{filters['voucher_id']}%")
    if filters.get("name"):
        base += " AND customer_name LIKE ?"
        params.append(f"%{filters['name']}%")
    if filters.get("contact"):
        base += " AND contact_number LIKE ?"
        params.append(f"%{filters['contact']}%")
    if filters.get("date_from"):
        base += " AND created_at >= ?"
        params.append(filters["date_from"].strip() + " 00:00:00")
    if filters.get("date_to"):
        base += " AND created_at <= ?"
        params.append(filters["date_to"].strip() + " 23:59:59")
    if filters.get("status") and filters["status"] != "All":
        base += " AND status = ?"
        params.append(filters["status"])

    base += " ORDER BY created_at DESC"
    cur.execute(base, params)
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
        filt_frame = ctk.CTkFrame(self)
        filt_frame.pack(fill="x", padx=10, pady=(10, 4))

        self.f_voucher = ctk.CTkEntry(filt_frame, width=130, placeholder_text="VoucherID")
        self.f_name    = ctk.CTkEntry(filt_frame, width=200, placeholder_text="Customer Name")
        self.f_contact = ctk.CTkEntry(filt_frame, width=160, placeholder_text="Contact Number")
        self.f_from    = ctk.CTkEntry(filt_frame, width=120, placeholder_text="Date From (YYYY-MM-DD)")
        self.f_to      = ctk.CTkEntry(filt_frame, width=120, placeholder_text="Date To (YYYY-MM-DD)")
        self.f_status  = ctk.CTkOptionMenu(filt_frame, values=["All","Pending","Collected"], width=120)
        self.f_status.set("All")

        self.btn_search = ctk.CTkButton(filt_frame, text="Search", command=self.perform_search, width=100)
        self.btn_reset  = ctk.CTkButton(filt_frame, text="Reset",  command=self.reset_filters, width=80)

        for w in (self.f_voucher, self.f_name, self.f_contact, self.f_from, self.f_to, self.f_status,
                  self.btn_search, self.btn_reset):
            w.pack(side="left", padx=5, pady=8)

        # ---- Table ----
        self.tree = ttk.Treeview(self,
            columns=("VoucherID","Date","Customer","Contact","Units","Status","Remark","PDF"),
            show="headings", height=18)
        for col, w in zip(self.tree["columns"],
                          (120,160,200,140,70,100,300,0)):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="w", stretch=(col!="PDF"))
        # hide PDF path column visually
        self.tree.column("PDF", width=0, stretch=False)
        self.tree.pack(expand=True, fill="both", padx=10, pady=(4, 8))

        # ---- Actions bar ----
        bar = ctk.CTkFrame(self)
        bar.pack(fill="x", padx=10, pady=(0,10))

        self.btn_add   = ctk.CTkButton(bar, text="Add Voucher",        command=self.add_voucher_ui)
        self.btn_mark  = ctk.CTkButton(bar, text="Mark as Collected",  command=self.mark_selected)
        self.btn_open  = ctk.CTkButton(bar, text="Open PDF",           command=self.open_pdf)
        self.btn_export= ctk.CTkButton(bar, text="Export (CSV/XLSX)",  command=self.export_data)

        for b in (self.btn_add, self.btn_mark, self.btn_open, self.btn_export):
            b.pack(side="left", padx=8, pady=8)

        self.perform_search()

    # -------- Filters / Search ----------
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
        self.f_voucher.delete(0, "end")
        self.f_name.delete(0, "end")
        self.f_contact.delete(0, "end")
        self.f_from.delete(0, "end")
        self.f_to.delete(0, "end")
        self.f_status.set("All")
        self.perform_search()

    def perform_search(self):
        rows = search_vouchers(self._get_filters())
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("", "end", values=r)

    # -------- Create Voucher ----------
    def add_voucher_ui(self):
        top = ctk.CTkToplevel(self)
        top.title("Create Voucher")
        top.geometry("760x560")  # bigger windowed mode
        top.grab_set()

        body = ctk.CTkFrame(top)
        body.pack(fill="both", expand=True, padx=12, pady=12)

        # Left column
        lbl_style = {"anchor":"w"}
        row = 0
        ctk.CTkLabel(body, text="Customer Name", **lbl_style).grid(row=row, column=0, sticky="w", pady=(0,2))
        e_name = ctk.CTkEntry(body, width=280); e_name.grid(row=row, column=1, sticky="w", padx=10, pady=(0,6))
        row += 1

        ctk.CTkLabel(body, text="Contact Number", **lbl_style).grid(row=row, column=0, sticky="w", pady=(0,2))
        e_contact = ctk.CTkEntry(body, width=280); e_contact.grid(row=row, column=1, sticky="w", padx=10, pady=(0,6))
        row += 1

        ctk.CTkLabel(body, text="No. of Units", **lbl_style).grid(row=row, column=0, sticky="w", pady=(0,2))
        e_units = ctk.CTkEntry(body, width=120); e_units.insert(0,"1")
        e_units.grid(row=row, column=1, sticky="w", padx=10, pady=(0,6))
        row += 1

        ctk.CTkLabel(body, text="Particulars", **lbl_style).grid(row=row, column=0, sticky="w", pady=(0,2))
        t_part = tk.Text(body, width=46, height=4); t_part.grid(row=row, column=1, sticky="w", padx=10, pady=(0,6))
        row += 1

        ctk.CTkLabel(body, text="Problem", **lbl_style).grid(row=row, column=0, sticky="w", pady=(0,2))
        t_prob = tk.Text(body, width=46, height=4); t_prob.grid(row=row, column=1, sticky="w", padx=10, pady=(0,6))
        row += 1

        ctk.CTkLabel(body, text="Staff Name", **lbl_style).grid(row=row, column=0, sticky="w", pady=(0,2))
        e_staff = ctk.CTkEntry(body, width=280); e_staff.grid(row=row, column=1, sticky="w", padx=10, pady=(0,6))
        row += 1

        ctk.CTkLabel(body, text="Remark", **lbl_style).grid(row=row, column=0, sticky="w", pady=(0,2))
        t_remark = tk.Text(body, width=46, height=4); t_remark.grid(row=row, column=1, sticky="w", padx=10, pady=(0,6))

        # Buttons
        btns = ctk.CTkFrame(top)
        btns.pack(fill="x", padx=12, pady=(0,12))
        def save():
            name = e_name.get().strip()
            contact = e_contact.get().strip()
            try:
                units = int(e_units.get().strip() or "1")
            except ValueError:
                messagebox.showerror("Invalid", "Units must be a number")
                return
            particulars = t_part.get("1.0","end").strip()
            problem     = t_prob.get("1.0","end").strip()
            remark      = t_remark.get("1.0","end").strip()
            if not name or not contact:
                messagebox.showerror("Missing", "Customer name and contact are required.")
                return
            voucher_id, pdf_path = add_voucher(name, contact, units, remark, particulars, problem, e_staff.get().strip())
            messagebox.showinfo("Saved", f"Voucher {voucher_id} created.")
            # Open PDF in default browser
            try:
                webbrowser.open_new(os.path.abspath(pdf_path))
            except Exception:
                pass
            top.destroy()
            self.perform_search()

        ctk.CTkButton(btns, text="Save & Open PDF", command=save, width=160).pack(side="right")

    # -------- Row actions ----------
    def mark_selected(self):
        sel = self.tree.focus()
        if not sel:
            messagebox.showerror("Error", "Select a record first.")
            return
        voucher_id = self.tree.item(sel)["values"][0]
        mark_collected(voucher_id)
        messagebox.showinfo("Updated", f"Voucher {voucher_id} marked as Collected.")
        self.perform_search()

    def open_pdf(self):
        sel = self.tree.focus()
        if not sel:
            messagebox.showerror("Error", "Select a record first.")
            return
        pdf_path = self.tree.item(sel)["values"][-1]
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Error", "PDF not found for this voucher.")
            return
        try:
            webbrowser.open_new(os.path.abspath(pdf_path))
        except Exception:
            # Fallback to OS open
            if os.name == "nt":
                os.startfile(pdf_path)
            else:
                os.system(f"open '{pdf_path}'")

    def export_data(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files","*.csv"), ("Excel Files","*.xlsx")]
        )
        if not filename:
            return
        export_filtered(self._get_filters(), filename)
        messagebox.showinfo("Exported", f"Exported to {filename}")

# ------------------ Run ------------------
if __name__ == "__main__":
    init_db()
    app = VoucherApp()
    app.mainloop()
