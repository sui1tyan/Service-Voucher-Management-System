import tkinter as tk
import customtkinter as ctk
from tkinter import ttk, messagebox, simpledialog
import threading
import os
import sys
import webbrowser
from datetime import datetime

from config import FONT_FAMILY, UI_FONT_SIZE, PDF_DIR, logger
from database import get_conn, search_vouchers, list_staffs_names, get_next_voucher_id
from auth import verify_pwd, validate_password_policy, hash_pwd
from pdf_utils import generate_pdf

def white_btn(parent, **kwargs):
    kwargs.setdefault("fg_color", "white")
    kwargs.setdefault("text_color", "black")
    return ctk.CTkButton(parent, **kwargs)

class LoginDialog(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Login")
        self.geometry("400x250")
        self.resizable(False, False)
        self.grab_set()
        
        frm = ctk.CTkFrame(self)
        frm.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(frm, text="Username").pack(anchor="w")
        self.e_user = ctk.CTkEntry(frm)
        self.e_user.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(frm, text="Password").pack(anchor="w")
        self.e_pwd = ctk.CTkEntry(frm, show="*")
        self.e_pwd.pack(fill="x", pady=(0, 20))
        
        ctk.CTkButton(frm, text="Login", command=self._login).pack(fill="x")
        self.result = None

    def _login(self):
        u = self.e_user.get()
        p = self.e_pwd.get()
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT id, username, role, password_hash FROM users WHERE username=?", (u,))
        row = cur.fetchone()
        conn.close()
        
        if row and verify_pwd(p, row["password_hash"]):
            self.result = {"id": row["id"], "username": row["username"], "role": row["role"]}
            self.destroy()
        else:
            messagebox.showerror("Error", "Invalid credentials")

class VoucherApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("SVMS")
        self.geometry("1100x700")
        
        self.user = None
        self._login()
        
        # Layout
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        self._build_filters()
        self._build_table()
        self._build_bottom()
        
        self.perform_search()

    def _login(self):
        dlg = LoginDialog(self)
        self.wait_window(dlg)
        if not dlg.result:
            self.destroy()
            sys.exit()
        self.user = dlg.result
        self.title(f"SVMS - Logged in as {self.user['username']}")

    def _build_filters(self):
        frm = ctk.CTkFrame(self)
        frm.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        self.e_vid = ctk.CTkEntry(frm, placeholder_text="Voucher ID")
        self.e_vid.pack(side="left", padx=5)
        
        self.e_name = ctk.CTkEntry(frm, placeholder_text="Customer Name")
        self.e_name.pack(side="left", padx=5)
        
        self.btn_search = ctk.CTkButton(frm, text="Search", command=self.perform_search)
        self.btn_search.pack(side="left", padx=10)
        
        ctk.CTkButton(frm, text="Reset", command=self.reset, fg_color="gray").pack(side="left")

    def _build_table(self):
        frm = ctk.CTkFrame(self)
        frm.grid(row=1, column=0, sticky="nsew", padx=10)
        
        cols = ("ID", "Date", "Customer", "Contact", "Status")
        self.tree = ttk.Treeview(frm, columns=cols, show="headings")
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=150)
            
        self.tree.pack(fill="both", expand=True, side="left")
        
        sb = ttk.Scrollbar(frm, command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)
        
        self.tree.bind("<Double-1>", lambda e: self.open_pdf())

    def _build_bottom(self):
        frm = ctk.CTkFrame(self)
        frm.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        
        ctk.CTkButton(frm, text="Add Voucher", command=self.add_voucher_ui).pack(side="left", padx=5)
        ctk.CTkButton(frm, text="Open PDF", command=self.open_pdf).pack(side="left", padx=5)

    def perform_search(self):
        filters = {
            "voucher_id": self.e_vid.get(),
            "customer_name": self.e_name.get(),
            "status": "All"
        }
        
        def _bg():
            try:
                rows = search_vouchers(filters)
                self.after(0, lambda: self._update_tree(rows))
            except Exception as e:
                logger.error(f"Search failed: {e}")
        
        threading.Thread(target=_bg, daemon=True).start()

    def _update_tree(self, rows):
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            # row: voucher_id, created, name, contact, ... status ... pdf
            self.tree.insert("", "end", values=(r[0], r[1][:10], r[2], r[3], r[8], r[10]))

    def reset(self):
        self.e_vid.delete(0, "end")
        self.e_name.delete(0, "end")
        self.perform_search()

    def open_pdf(self):
        sel = self.tree.selection()
        if not sel: return
        # The stored rows in tree don't contain PDF path in displayed cols, 
        # but I packed it in previous logic? Let's fix search_vouchers return.
        # It returns ... status, solution, pdf_path. 
        # tree values are tuples. 
        # We need to fetch the full row or store path hidden.
        # Simplified: fetch path from DB by ID.
        vid = self.tree.item(sel[0])["values"][0]
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT pdf_path FROM vouchers WHERE voucher_id=?", (vid,))
        row = cur.fetchone()
        conn.close()
        
        if row and row[0] and os.path.exists(row[0]):
            webbrowser.open(row[0])
        else:
            messagebox.showerror("Error", "PDF not found")

    def add_voucher_ui(self):
        top = ctk.CTkToplevel(self)
        top.title("New Voucher")
        top.geometry("500x550")
        
        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=20, pady=20)
        
        entries = {}
        for label in ["Customer Name", "Contact", "Particulars", "Problem"]:
            ctk.CTkLabel(frm, text=label).pack(anchor="w")
            e = ctk.CTkEntry(frm)
            e.pack(fill="x", pady=(0, 10))
            entries[label] = e
            
        ctk.CTkLabel(frm, text="Recipient").pack(anchor="w")
        cb = ctk.CTkComboBox(frm, values=list_staffs_names())
        cb.pack(fill="x", pady=(0, 20))
        
        def save():
            vid = get_next_voucher_id()
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            pdf = generate_pdf(vid, entries["Customer Name"].get(), entries["Contact"].get(), 1,
                               entries["Particulars"].get(), entries["Problem"].get(), 
                               cb.get(), "Pending", ts, cb.get())
            
            conn = get_conn()
            conn.execute("""
                INSERT INTO vouchers (voucher_id, created_at, customer_name, contact_number, 
                particulars, problem, staff_name, recipient, pdf_path, status)
                VALUES (?,?,?,?,?,?,?,?,?,?)
            """, (vid, ts, entries["Customer Name"].get(), entries["Contact"].get(), 
                  entries["Particulars"].get(), entries["Problem"].get(), cb.get(), cb.get(), pdf, "Pending"))
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Success", f"Voucher {vid} created.")
            top.destroy()
            self.perform_search()
            
        ctk.CTkButton(frm, text="Save", command=save).pack(fill="x")
