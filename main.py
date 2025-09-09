#!/usr/bin/env python3
# Service Voucher Management System (Monolith Version)
# Dependencies: customtkinter, reportlab, qrcode, pillow, pandas, openpyxl (optional for Excel)

import os
import sqlite3
import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox, filedialog, ttk
from datetime import datetime
import qrcode
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import pandas as pd

DB_FILE = "vouchers.db"
PDF_DIR = "pdfs"
os.makedirs(PDF_DIR, exist_ok=True)


# ---------- Database Initialization ----------
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS vouchers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        customer_name TEXT,
        contact_number TEXT,
        voucher_id TEXT,
        status TEXT DEFAULT 'Pending',
        created_at TEXT,
        pdf_path TEXT
    )
    """)
    conn.commit()
    conn.close()


# ---------- PDF Generation ----------
def generate_pdf(voucher_id, customer_name, contact_number, status):
    filename = os.path.join(PDF_DIR, f"voucher_{voucher_id}.pdf")

    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    # Shop header
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, height - 50, "My Repair Shop")
    c.setFont("Helvetica", 10)
    c.drawString(50, height - 65, "123 Main Street, City, Country")
    c.drawString(50, height - 80, "Contact: +6012-3456789")

    # Voucher ID in bold
    c.setFont("Helvetica-Bold", 24)
    c.drawString(200, height - 140, f"Voucher ID: {voucher_id}")

    # QR Code
    qr_data = f"Voucher: {voucher_id}\nName: {customer_name}\nContact: {contact_number}"
    qr_img = qrcode.make(qr_data)
    qr_path = os.path.join(PDF_DIR, f"qr_{voucher_id}.png")
    qr_img.save(qr_path)
    c.drawImage(qr_path, 50, height - 250, 100, 100)

    # Customer details
    c.setFont("Helvetica", 12)
    c.drawString(50, height - 300, f"Customer Name: {customer_name}")
    c.drawString(50, height - 320, f"Contact Number: {contact_number}")
    c.drawString(50, height - 340, f"Status: {status}")
    c.drawString(50, height - 360, f"Date Issued: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Signature lines
    c.line(50, 100, 250, 100)
    c.drawString(50, 85, "Customer Signature")

    c.line(300, 100, 500, 100)
    c.drawString(300, 85, "Staff Signature")

    c.save()
    return filename


# ---------- DB Operations ----------
def add_voucher(customer_name, contact_number):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    voucher_id = f"V{int(datetime.now().timestamp())}"  # unique ID
    status = "Pending"
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    pdf_path = generate_pdf(voucher_id, customer_name, contact_number, status)

    cur.execute("""
    INSERT INTO vouchers (customer_name, contact_number, voucher_id, status, created_at, pdf_path)
    VALUES (?, ?, ?, ?, ?, ?)
    """, (customer_name, contact_number, voucher_id, status, created_at, pdf_path))

    conn.commit()
    conn.close()
    return voucher_id


def search_vouchers(keyword):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("""
    SELECT * FROM vouchers
    WHERE customer_name LIKE ? OR contact_number LIKE ? OR voucher_id LIKE ?
    """, (f"%{keyword}%", f"%{keyword}%", f"%{keyword}%"))
    rows = cur.fetchall()
    conn.close()
    return rows


def mark_collected(voucher_id):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("UPDATE vouchers SET status = 'Collected' WHERE voucher_id = ?", (voucher_id,))
    conn.commit()
    conn.close()


def export_vouchers(filename):
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM vouchers", conn)
    conn.close()
    if filename.endswith(".csv"):
        df.to_csv(filename, index=False)
    elif filename.endswith(".xlsx"):
        df.to_excel(filename, index=False)


# ---------- UI ----------
class VoucherApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Service Voucher Management System")
        self.geometry("900x600")
        ctk.set_appearance_mode("light")

        # Search bar
        self.search_entry = ctk.CTkEntry(self, placeholder_text="Search by Name, Contact, or Voucher ID")
        self.search_entry.pack(pady=10)
        self.search_button = ctk.CTkButton(self, text="Search", command=self.perform_search)
        self.search_button.pack(pady=5)

        # Table
        self.tree = ttk.Treeview(self, columns=("ID", "Name", "Contact", "VoucherID", "Status", "Date"), show="headings")
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
        self.tree.pack(expand=True, fill="both", padx=10, pady=10)

        # Buttons
        self.add_button = ctk.CTkButton(self, text="Add Voucher", command=self.add_voucher_ui)
        self.add_button.pack(side="left", padx=10, pady=10)

        self.mark_button = ctk.CTkButton(self, text="Mark as Collected", command=self.mark_selected)
        self.mark_button.pack(side="left", padx=10, pady=10)

        self.open_pdf_button = ctk.CTkButton(self, text="Open PDF", command=self.open_pdf)
        self.open_pdf_button.pack(side="left", padx=10, pady=10)

        self.export_button = ctk.CTkButton(self, text="Export Data", command=self.export_data)
        self.export_button.pack(side="left", padx=10, pady=10)

    def perform_search(self):
        keyword = self.search_entry.get()
        rows = search_vouchers(keyword)
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("", "end", values=r)

    def add_voucher_ui(self):
        top = tk.Toplevel(self)
        top.title("Add Voucher")

        tk.Label(top, text="Customer Name").pack(pady=5)
        name_entry = tk.Entry(top)
        name_entry.pack(pady=5)

        tk.Label(top, text="Contact Number").pack(pady=5)
        contact_entry = tk.Entry(top)
        contact_entry.pack(pady=5)

        def save():
            customer_name = name_entry.get()
            contact_number = contact_entry.get()
            if not customer_name or not contact_number:
                messagebox.showerror("Error", "All fields are required")
                return
            voucher_id = add_voucher(customer_name, contact_number)
            messagebox.showinfo("Success", f"Voucher {voucher_id} created")
            top.destroy()

        tk.Button(top, text="Save", command=save).pack(pady=10)

    def mark_selected(self):
        selected = self.tree.focus()
        if not selected:
            messagebox.showerror("Error", "No voucher selected")
            return
        voucher_id = self.tree.item(selected)["values"][3]
        mark_collected(voucher_id)
        messagebox.showinfo("Updated", f"Voucher {voucher_id} marked as Collected")
        self.perform_search()

    def open_pdf(self):
        selected = self.tree.focus()
        if not selected:
            messagebox.showerror("Error", "No voucher selected")
            return
        pdf_path = self.tree.item(selected)["values"][6]
        if os.path.exists(pdf_path):
            os.startfile(pdf_path) if os.name == "nt" else os.system(f"open '{pdf_path}'")
        else:
            messagebox.showerror("Error", "PDF not found")

    def export_data(self):
        filename = filedialog.asksaveasfilename(defaultextension=".csv",
                                                filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx")])
        if filename:
            export_vouchers(filename)
            messagebox.showinfo("Exported", f"Data exported to {filename}")


# ---------- Run ----------
if __name__ == "__main__":
    init_db()
    app = VoucherApp()
    app.mainloop()
