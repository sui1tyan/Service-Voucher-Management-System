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

# --- Helpers ---
def white_btn(parent, **kwargs):
    kwargs.setdefault("fg_color", "white")
    kwargs.setdefault("text_color", "black")
    kwargs.setdefault("border_color", "black")
    kwargs.setdefault("border_width", 1)
    kwargs.setdefault("hover_color", "#F0F0F0")
    return ctk.CTkButton(parent, **kwargs)

def _to_ui_date(dt_obj):
    return dt_obj.strftime("%d-%m-%Y")

# --- Login Dialog ---
class LoginDialog(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Login")
        self.geometry("460x240")
        self.resizable(False, False)
        self.grab_set()
        
        frm = ctk.CTkFrame(self)
        frm.pack(fill="both", expand=True, padx=16, pady=16)
        
        ctk.CTkLabel(frm, text="Username:").grid(row=0, column=0, sticky="w", pady=6)
        self.e_user = ctk.CTkEntry(frm, width=220)
        self.e_user.grid(row=0, column=1, sticky="w", pady=6)
        
        ctk.CTkLabel(frm, text="Password:").grid(row=1, column=0, sticky="w", pady=6)
        self.e_pwd = ctk.CTkEntry(frm, width=220, show="•")
        self.e_pwd.grid(row=1, column=1, sticky="w", pady=6)
        
        self.btn_login = ctk.CTkButton(frm, text="Login", command=self._do_login, width=120)
        self.btn_login.grid(row=2, column=0, columnspan=2, pady=20)
        
        self.result = None
        self.bind("<Return>", lambda e: self._do_login())

    def _do_login(self):
        u = self.e_user.get().strip()
        p = self.e_pwd.get().strip()
        
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT id, username, role, password_hash, is_active, must_change_pwd FROM users WHERE username=?", (u,))
        row = cur.fetchone()
        conn.close()
        
        if row and row['is_active'] and verify_pwd(p, row['password_hash']):
            if row['must_change_pwd']:
                messagebox.showinfo("Change Password", "You must change your password.")
                new_pwd = simpledialog.askstring("New Password", "Enter new password:", show="*", parent=self)
                if not new_pwd or validate_password_policy(new_pwd):
                    messagebox.showerror("Error", "Invalid password or cancelled.")
                    return
                # Update Pwd
                conn = get_conn()
                conn.execute("UPDATE users SET password_hash=?, must_change_pwd=0 WHERE id=?", (hash_pwd(new_pwd), row['id']))
                conn.commit()
                conn.close()
            
            self.result = {"id": row['id'], "username": row['username'], "role": row['role']}
            self.destroy()
        else:
            messagebox.showerror("Login", "Invalid credentials")

# --- Main App ---
class VoucherApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Service Voucher Management System")
        self.geometry("1280x780")
        ctk.set_appearance_mode("light")
        
        self.current_user = None
        self._do_login_flow()
        
        # Main Layout
        root = ctk.CTkFrame(self)
        root.pack(fill="both", expand=True)
        
        self._build_filters(root)
        self._build_table(root)
        self._build_bottom_bar(root)
        
        self.after(100, self.perform_search)

    def _do_login_flow(self):
        dlg = LoginDialog(self)
        self.wait_window(dlg)
        if not dlg.result:
            self.destroy()
            sys.exit(0)
        self.current_user = dlg.result
        self.title(f"SVMS — {self.current_user['username']} ({self.current_user['role']})")

    def _build_filters(self, parent):
        wrap = ctk.CTkFrame(parent)
        wrap.pack(fill="x", padx=8, pady=8)
        
        self.f_voucher = ctk.CTkEntry(wrap, width=140, placeholder_text="VoucherID")
        self.f_voucher.pack(side="left", padx=5)
        
        self.f_name = ctk.CTkEntry(wrap, width=200, placeholder_text="Customer Name")
        self.f_name.pack(side="left", padx=5)
        
        self.f_status = ctk.CTkOptionMenu(wrap, values=["All", "Pending", "Completed"], width=140)
        self.f_status.set("All")
        self.f_status.pack(side="left", padx=5)
        
        self.btn_search = white_btn(wrap, text="Search", command=self.perform_search, width=100)
        self.btn_search.pack(side="left", padx=5)
        
        white_btn(wrap, text="Reset", command=self.reset_filters, width=100).pack(side="left", padx=5)

    def _build_table(self, parent):
        frame = ctk.CTkFrame(parent)
        frame.pack(fill="both", expand=True, padx=8, pady=4)
        
        cols = ("VoucherID", "Date", "Customer", "Contact", "Status", "PDF")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        
        widths = [100, 150, 250, 150, 120, 0]
        for c, w in zip(cols, widths):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w)
            
        self.tree.pack(fill="both", expand=True, side="left")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.bind("<Double-1>", lambda e: self.open_pdf())

    def _build_bottom_bar(self, parent):
        bar = ctk.CTkFrame(parent)
        bar.pack(fill="x", padx=8, pady=8)
        white_btn(bar, text="Add Voucher", command=self.add_voucher_ui, width=140).pack(side="left", padx=5)
        white_btn(bar, text="Open PDF", command=self.open_pdf, width=140).pack(side="left", padx=5)

    # --- Logic with Threading Patch ---
    def perform_search(self):
        filters = {
            "voucher_id": self.f_voucher.get().strip(),
            "customer_name": self.f_name.get().strip(),
            "status": self.f_status.get()
        }
        
        self.btn_search.configure(state="disabled")
        
        def _bg():
            try:
                rows = search_vouchers(filters)
                self.after(0, lambda: self._update_table(rows))
            except Exception as e:
                logger.exception("Search failed")
            finally:
                self.after(0, lambda: self.btn_search.configure(state="normal"))
        
        threading.Thread(target=_bg, daemon=True).start()

    def _update_table(self, rows):
        self.tree.delete(*self.tree.get_children())
        for row in rows:
            # Row mapping: voucher_id, created, name, contact, units, recipient, tech_id, tech_name, status, sol, pdf
            self.tree.insert("", "end", values=(row[0], row[1], row[2], row[3], row[8], row[10]))

    def reset_filters(self):
        self.f_voucher.delete(0, "end")
        self.f_name.delete(0, "end")
        self.f_status.set("All")
        self.perform_search()

    def open_pdf(self):
        sel = self.tree.selection()
        if not sel: return
        pdf_path = self.tree.item(sel[0])["values"][-1]
        if pdf_path and os.path.exists(pdf_path):
            webbrowser.open(pdf_path)
        else:
            messagebox.showerror("Error", "PDF not found.")

    # --- Add Voucher UI (Refactored to reduce repetition) ---
    def add_voucher_ui(self):
        top = ctk.CTkToplevel(self)
        top.title("Create Voucher")
        top.geometry("600x600")
        top.grab_set()
        
        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Helper for fields
        def add_field(label, row):
            ctk.CTkLabel(frm, text=label).grid(row=row, column=0, sticky="w", pady=5)
            e = ctk.CTkEntry(frm, width=300)
            e.grid(row=row, column=1, sticky="w", pady=5, padx=10)
            return e
            
        e_name = add_field("Customer Name", 0)
        e_contact = add_field("Contact", 1)
        e_part = add_field("Particulars", 2)
        e_prob = add_field("Problem", 3)
        
        ctk.CTkLabel(frm, text="Recipient").grid(row=4, column=0, sticky="w")
        staffs = list_staffs_names()
        cb_staff = ctk.CTkComboBox(frm, values=staffs, width=300)
        cb_staff.grid(row=4, column=1, sticky="w", pady=5, padx=10)
        
        def save():
            vid = get_next_voucher_id()
            pdf_path = generate_pdf(
                vid, e_name.get(), e_contact.get(), 1, 
                e_part.get(), e_prob.get(), cb_staff.get(), 
                "Pending", datetime.now().strftime("%Y-%m-%d %H:%M:%S"), cb_staff.get()
            )
            
            conn = get_conn()
            conn.execute(
                "INSERT INTO vouchers (voucher_id, created_at, customer_name, contact_number, particulars, problem, staff_name, recipient, pdf_path, status) VALUES (?,?,?,?,?,?,?,?,?,?)",
                (vid, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), e_name.get(), e_contact.get(), e_part.get(), e_prob.get(), cb_staff.get(), cb_staff.get(), pdf_path, "Pending")
            )
            conn.commit()
            conn.close()
            messagebox.showinfo("Success", f"Voucher {vid} created.")
            top.destroy()
            self.perform_search()
            
        white_btn(frm, text="Save", command=save, width=120).grid(row=6, column=1, pady=20)
