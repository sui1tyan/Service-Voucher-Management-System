"""CustomTkinter desktop interface for SVMS."""

from __future__ import annotations

import os
import subprocess
import sys
import time
import webbrowser
from pathlib import Path
from typing import Any, Callable, Mapping

import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog, ttk

from auth import validate_password_policy
from backup_utils import create_backup, restore_backup
from config import (
    APP_NAME,
    BILL_TYPES,
    FONT_FAMILY,
    PDF_DIR,
    ROLE_PERMISSIONS,
    ROLES,
    VOUCHER_STATUSES,
    logger,
)
from database import (
    authenticate_user,
    change_own_password,
    create_commission,
    create_initial_admin,
    create_staff,
    create_user,
    create_voucher,
    delete_commission,
    delete_staff,
    delete_user,
    delete_voucher,
    force_change_password,
    get_commission,
    get_staff,
    get_voucher,
    has_admin_user,
    list_audit_entries,
    list_commissions,
    list_staffs,
    list_staffs_names,
    list_users,
    reset_user_password,
    search_vouchers,
    set_base_voucher_id,
    update_commission,
    update_staff,
    update_user,
    update_voucher,
    update_voucher_pdf_path,
)
from export_utils import export_commissions_csv, export_vouchers_csv
from pdf_utils import generate_voucher_pdf
from validators import (
    ValidationError,
    validate_commission,
    validate_staff,
    validate_voucher,
)


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


def has_permission(user: Mapping[str, Any], permission: str) -> bool:
    return permission in ROLE_PERMISSIONS.get(str(user.get("role") or ""), set())


def open_local_file(path: str | Path) -> None:
    target = Path(path).resolve()
    if not target.is_file():
        raise FileNotFoundError(str(target))
    if sys.platform.startswith("win"):
        os.startfile(str(target))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(target)], check=False)
    else:
        webbrowser.open(target.as_uri())


def _selected_id(tree: ttk.Treeview) -> int | None:
    selection = tree.selection()
    if not selection:
        return None
    values = tree.item(selection[0], "values")
    return int(values[0]) if values else None


class Modal(ctk.CTkToplevel):
    """Base class for modal dialogs with a result value."""

    def __init__(self, master: Any, title: str, geometry: str) -> None:
        super().__init__(master)
        self.title(title)
        self.geometry(geometry)
        self.minsize(420, 240)
        self.transient(master)
        self.grab_set()
        self.result: Any = None
        self.after(80, self.focus_force)


class FirstRunAdminDialog(Modal):
    def __init__(self, master: Any) -> None:
        super().__init__(master, "Create First Administrator", "520x460")
        frame = ctk.CTkFrame(self)
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(
            frame,
            text="First-run administrator setup",
            font=(FONT_FAMILY, 20, "bold"),
        ).pack(pady=(8, 4))
        ctk.CTkLabel(
            frame,
            text="No default password is used. Create the first administrator account.",
            wraplength=430,
        ).pack(pady=(0, 16))

        self.username = self._field(frame, "Username")
        self.full_name = self._field(frame, "Full name")
        self.password = self._field(frame, "Password", show="*")
        self.confirm = self._field(frame, "Confirm password", show="*")
        self.show_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            frame,
            text="Show passwords",
            variable=self.show_var,
            command=self._toggle_passwords,
        ).pack(anchor="w", padx=8, pady=8)
        ctk.CTkButton(frame, text="Create Administrator", command=self._save).pack(
            fill="x", padx=8, pady=12
        )

    @staticmethod
    def _field(parent: Any, label: str, **kwargs: Any) -> ctk.CTkEntry:
        ctk.CTkLabel(parent, text=label).pack(anchor="w", padx=8)
        entry = ctk.CTkEntry(parent, **kwargs)
        entry.pack(fill="x", padx=8, pady=(2, 10))
        return entry

    def _toggle_passwords(self) -> None:
        show = "" if self.show_var.get() else "*"
        self.password.configure(show=show)
        self.confirm.configure(show=show)

    def _save(self) -> None:
        password = self.password.get()
        if password != self.confirm.get():
            messagebox.showerror("Setup", "Passwords do not match.", parent=self)
            return
        try:
            self.result = create_initial_admin(
                self.username.get().strip(),
                password,
                self.full_name.get().strip(),
            )
        except (ValueError, ValidationError) as exc:
            messagebox.showerror("Setup", str(exc), parent=self)
            return
        self.destroy()


class LoginDialog(Modal):
    MAX_ATTEMPTS = 5
    LOCK_SECONDS = 30

    def __init__(self, master: Any) -> None:
        super().__init__(master, "SVMS Login", "440x340")
        self.resizable(False, False)
        self.attempts = 0
        self.locked_until = 0.0

        frame = ctk.CTkFrame(self)
        frame.pack(fill="both", expand=True, padx=24, pady=24)
        ctk.CTkLabel(
            frame, text=APP_NAME, font=(FONT_FAMILY, 20, "bold")
        ).pack(pady=(10, 20))
        ctk.CTkLabel(frame, text="Username").pack(anchor="w", padx=8)
        self.username = ctk.CTkEntry(frame)
        self.username.pack(fill="x", padx=8, pady=(2, 12))
        ctk.CTkLabel(frame, text="Password").pack(anchor="w", padx=8)
        self.password = ctk.CTkEntry(frame, show="*")
        self.password.pack(fill="x", padx=8, pady=(2, 18))
        self.login_button = ctk.CTkButton(frame, text="Login", command=self._login)
        self.login_button.pack(fill="x", padx=8)
        self.status_label = ctk.CTkLabel(frame, text="", text_color="#B00020")
        self.status_label.pack(pady=8)
        self.bind("<Return>", lambda _event: self._login())
        self.after(100, self.username.focus_set)

    def _login(self) -> None:
        remaining = self.locked_until - time.monotonic()
        if remaining > 0:
            self.status_label.configure(
                text=f"Too many attempts. Try again in {int(remaining) + 1} seconds."
            )
            return
        user = authenticate_user(self.username.get(), self.password.get())
        if user:
            self.result = user
            self.destroy()
            return
        self.attempts += 1
        self.password.delete(0, "end")
        if self.attempts >= self.MAX_ATTEMPTS:
            self.attempts = 0
            self.locked_until = time.monotonic() + self.LOCK_SECONDS
            self.status_label.configure(text="Too many attempts. Login temporarily locked.")
        else:
            self.status_label.configure(text="Invalid username or password.")


class PasswordDialog(Modal):
    def __init__(
        self,
        master: Any,
        user: Mapping[str, Any],
        *,
        forced: bool,
    ) -> None:
        super().__init__(
            master,
            "Required Password Change" if forced else "Change Password",
            "480x430" if not forced else "480x360",
        )
        self.user = user
        self.forced = forced
        frame = ctk.CTkFrame(self)
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(
            frame,
            text="Choose a secure password",
            font=(FONT_FAMILY, 18, "bold"),
        ).pack(pady=(8, 4))
        ctk.CTkLabel(
            frame,
            text="At least 12 characters with uppercase, lowercase, digit, and symbol.",
            wraplength=400,
        ).pack(pady=(0, 12))

        self.current: ctk.CTkEntry | None = None
        if not forced:
            self.current = self._field(frame, "Current password")
        self.new = self._field(frame, "New password")
        self.confirm = self._field(frame, "Confirm new password")
        ctk.CTkButton(frame, text="Save Password", command=self._save).pack(
            fill="x", padx=8, pady=12
        )

    @staticmethod
    def _field(parent: Any, label: str) -> ctk.CTkEntry:
        ctk.CTkLabel(parent, text=label).pack(anchor="w", padx=8)
        entry = ctk.CTkEntry(parent, show="*")
        entry.pack(fill="x", padx=8, pady=(2, 9))
        return entry

    def _save(self) -> None:
        new_password = self.new.get()
        if new_password != self.confirm.get():
            messagebox.showerror("Password", "Passwords do not match.", parent=self)
            return
        if error := validate_password_policy(new_password):
            messagebox.showerror("Password", error, parent=self)
            return
        try:
            if self.forced:
                force_change_password(
                    int(self.user["id"]), new_password, str(self.user["username"])
                )
            else:
                assert self.current is not None
                change_own_password(
                    int(self.user["id"]),
                    self.current.get(),
                    new_password,
                    str(self.user["username"]),
                )
        except ValueError as exc:
            messagebox.showerror("Password", str(exc), parent=self)
            return
        self.result = True
        self.destroy()


class VoucherDialog(Modal):
    ENTRY_FIELDS = (
        ("customer_name", "Customer name"),
        ("contact_number", "Contact number"),
        ("units", "Units"),
        ("technician_id", "Technician ID"),
        ("technician_name", "Technician name"),
        ("ref_bill", "Reference bill"),
        ("ref_bill_date", "Bill date (YYYY-MM-DD)"),
        ("amount_rm", "Amount (RM)"),
        ("tech_commission", "Technician commission (RM)"),
    )

    def __init__(
        self,
        master: Any,
        *,
        voucher: Mapping[str, Any] | None = None,
    ) -> None:
        super().__init__(
            master, "Edit Voucher" if voucher else "New Voucher", "760x780"
        )
        self.voucher = dict(voucher or {})
        self.entries: dict[str, ctk.CTkEntry] = {}

        outer = ctk.CTkFrame(self)
        outer.pack(fill="both", expand=True, padx=12, pady=12)
        form = ctk.CTkScrollableFrame(outer)
        form.pack(fill="both", expand=True, padx=8, pady=8)
        form.grid_columnconfigure(1, weight=1)

        row = 0
        for field, label in self.ENTRY_FIELDS:
            ctk.CTkLabel(form, text=label).grid(
                row=row, column=0, sticky="w", padx=8, pady=6
            )
            entry = ctk.CTkEntry(form)
            entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
            entry.insert(0, str(self.voucher.get(field) or ("1" if field == "units" else "")))
            self.entries[field] = entry
            row += 1

        ctk.CTkLabel(form, text="Recipient").grid(
            row=row, column=0, sticky="w", padx=8, pady=6
        )
        recipients = list_staffs_names()
        current_recipient = str(self.voucher.get("recipient") or "")
        if current_recipient and current_recipient not in recipients:
            recipients.append(current_recipient)
        self.recipient = ctk.CTkComboBox(form, values=recipients or [""])
        self.recipient.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        self.recipient.set(current_recipient)
        row += 1

        ctk.CTkLabel(form, text="Status").grid(
            row=row, column=0, sticky="w", padx=8, pady=6
        )
        self.status = ctk.CTkComboBox(form, values=list(VOUCHER_STATUSES))
        self.status.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        self.status.set(str(self.voucher.get("status") or "Pending"))
        row += 1

        self.textboxes: dict[str, ctk.CTkTextbox] = {}
        for field, label in (
            ("particulars", "Particulars"),
            ("problem", "Reported problem"),
            ("solution", "Solution / work performed"),
        ):
            ctk.CTkLabel(form, text=label).grid(
                row=row, column=0, sticky="nw", padx=8, pady=6
            )
            textbox = ctk.CTkTextbox(form, height=90)
            textbox.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
            textbox.insert("1.0", str(self.voucher.get(field) or ""))
            self.textboxes[field] = textbox
            row += 1

        buttons = ctk.CTkFrame(outer, fg_color="transparent")
        buttons.pack(fill="x", padx=8, pady=8)
        ctk.CTkButton(buttons, text="Save", command=self._save).pack(
            side="right", padx=4
        )
        ctk.CTkButton(
            buttons, text="Cancel", command=self.destroy, fg_color="#666666"
        ).pack(side="right", padx=4)

    def _save(self) -> None:
        payload: dict[str, Any] = {
            field: entry.get() for field, entry in self.entries.items()
        }
        payload.update(
            {
                field: textbox.get("1.0", "end").strip()
                for field, textbox in self.textboxes.items()
            }
        )
        payload["recipient"] = self.recipient.get()
        payload["staff_name"] = self.recipient.get()
        payload["status"] = self.status.get()
        try:
            self.result = validate_voucher(payload)
        except ValidationError as exc:
            messagebox.showerror("Voucher", str(exc), parent=self)
            return
        self.destroy()


class StatusDialog(Modal):
    def __init__(self, master: Any, voucher: Mapping[str, Any]) -> None:
        super().__init__(master, "Update Voucher Status", "560x520")
        self.voucher = dict(voucher)
        frame = ctk.CTkFrame(self)
        frame.pack(fill="both", expand=True, padx=18, pady=18)
        ctk.CTkLabel(
            frame,
            text=f"Voucher {voucher.get('voucher_id')}",
            font=(FONT_FAMILY, 18, "bold"),
        ).pack(pady=(4, 12))
        ctk.CTkLabel(frame, text="Status").pack(anchor="w")
        self.status = ctk.CTkComboBox(frame, values=list(VOUCHER_STATUSES))
        self.status.pack(fill="x", pady=(2, 10))
        self.status.set(str(voucher.get("status") or "Pending"))
        ctk.CTkLabel(frame, text="Technician ID").pack(anchor="w")
        self.technician_id = ctk.CTkEntry(frame)
        self.technician_id.pack(fill="x", pady=(2, 10))
        self.technician_id.insert(0, str(voucher.get("technician_id") or ""))
        ctk.CTkLabel(frame, text="Technician name").pack(anchor="w")
        self.technician_name = ctk.CTkEntry(frame)
        self.technician_name.pack(fill="x", pady=(2, 10))
        self.technician_name.insert(0, str(voucher.get("technician_name") or ""))
        ctk.CTkLabel(frame, text="Solution / work performed").pack(anchor="w")
        self.solution = ctk.CTkTextbox(frame, height=130)
        self.solution.pack(fill="both", expand=True, pady=(2, 12))
        self.solution.insert("1.0", str(voucher.get("solution") or ""))
        ctk.CTkButton(frame, text="Save Status", command=self._save).pack(fill="x")

    def _save(self) -> None:
        updated = dict(self.voucher)
        updated["status"] = self.status.get()
        updated["technician_id"] = self.technician_id.get()
        updated["technician_name"] = self.technician_name.get()
        updated["solution"] = self.solution.get("1.0", "end").strip()
        try:
            self.result = validate_voucher(updated)
        except ValidationError as exc:
            messagebox.showerror("Status", str(exc), parent=self)
            return
        self.destroy()


class StaffEditorDialog(Modal):
    FIELDS = (
        ("staff_id_opt", "Staff ID"),
        ("name", "Name"),
        ("position", "Position"),
        ("phone", "Phone"),
        ("email", "Email"),
        ("note", "Note"),
    )

    def __init__(
        self, master: Any, staff: Mapping[str, Any] | None = None
    ) -> None:
        super().__init__(master, "Edit Staff" if staff else "Add Staff", "520x580")
        self.staff = dict(staff or {})
        self.entries: dict[str, ctk.CTkEntry] = {}
        frame = ctk.CTkFrame(self)
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        for field, label in self.FIELDS:
            ctk.CTkLabel(frame, text=label).pack(anchor="w", padx=6)
            entry = ctk.CTkEntry(frame)
            entry.pack(fill="x", padx=6, pady=(2, 8))
            entry.insert(0, str(self.staff.get(field) or ""))
            self.entries[field] = entry
        ctk.CTkButton(frame, text="Save", command=self._save).pack(
            fill="x", padx=6, pady=12
        )

    def _save(self) -> None:
        payload = {field: entry.get() for field, entry in self.entries.items()}
        payload["photo_path"] = str(self.staff.get("photo_path") or "")
        try:
            self.result = validate_staff(payload)
        except ValidationError as exc:
            messagebox.showerror("Staff", str(exc), parent=self)
            return
        self.destroy()


class StaffManagementDialog(Modal):
    def __init__(
        self, master: Any, actor: str, on_change: Callable[[], None] | None = None
    ) -> None:
        super().__init__(master, "Manage Staff and Recipients", "900x560")
        self.actor = actor
        self.on_change = on_change
        self.tree = ttk.Treeview(
            self,
            columns=("ID", "Staff ID", "Name", "Position", "Phone", "Email"),
            show="headings",
        )
        for column, width in (
            ("ID", 60),
            ("Staff ID", 100),
            ("Name", 180),
            ("Position", 150),
            ("Phone", 130),
            ("Email", 200),
        ):
            self.tree.heading(column, text=column)
            self.tree.column(column, width=width, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=12, pady=(12, 6))
        controls = ctk.CTkFrame(self)
        controls.pack(fill="x", padx=12, pady=(6, 12))
        for label, command in (
            ("Add", self._add),
            ("Edit", self._edit),
            ("Delete", self._delete),
            ("Refresh", self.refresh),
        ):
            ctk.CTkButton(controls, text=label, command=command, width=110).pack(
                side="left", padx=4, pady=8
            )
        self.refresh()

    def refresh(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for row in list_staffs():
            self.tree.insert(
                "",
                "end",
                values=(
                    row["id"],
                    row.get("staff_id_opt") or "",
                    row["name"],
                    row.get("position") or "",
                    row.get("phone") or "",
                    row.get("email") or "",
                ),
            )
        if self.on_change:
            self.on_change()

    def _add(self) -> None:
        dialog = StaffEditorDialog(self)
        self.wait_window(dialog)
        if dialog.result:
            try:
                create_staff(dialog.result, self.actor)
            except ValueError as exc:
                messagebox.showerror("Staff", str(exc), parent=self)
            self.refresh()

    def _edit(self) -> None:
        staff_id = _selected_id(self.tree)
        if staff_id is None:
            messagebox.showinfo("Staff", "Select a staff member.", parent=self)
            return
        staff = get_staff(staff_id)
        dialog = StaffEditorDialog(self, staff)
        self.wait_window(dialog)
        if dialog.result:
            try:
                update_staff(staff_id, dialog.result, self.actor)
            except ValueError as exc:
                messagebox.showerror("Staff", str(exc), parent=self)
            self.refresh()

    def _delete(self) -> None:
        staff_id = _selected_id(self.tree)
        if staff_id is None:
            messagebox.showinfo("Staff", "Select a staff member.", parent=self)
            return
        if not messagebox.askyesno(
            "Delete Staff", "Delete the selected staff member?", parent=self
        ):
            return
        try:
            delete_staff(staff_id, self.actor)
        except ValueError as exc:
            messagebox.showerror("Staff", str(exc), parent=self)
        self.refresh()


class UserEditorDialog(Modal):
    def __init__(
        self, master: Any, user: Mapping[str, Any] | None = None
    ) -> None:
        super().__init__(master, "Edit User" if user else "Add User", "540x650")
        self.user = dict(user or {})
        self.entries: dict[str, ctk.CTkEntry] = {}
        frame = ctk.CTkScrollableFrame(self)
        frame.pack(fill="both", expand=True, padx=18, pady=18)
        for field, label in (
            ("username", "Username"),
            ("full_name", "Full name"),
            ("phone", "Phone"),
            ("email", "Email"),
            ("note", "Note"),
        ):
            ctk.CTkLabel(frame, text=label).pack(anchor="w", padx=6)
            entry = ctk.CTkEntry(frame)
            entry.pack(fill="x", padx=6, pady=(2, 8))
            entry.insert(0, str(self.user.get(field) or ""))
            self.entries[field] = entry
        ctk.CTkLabel(frame, text="Role").pack(anchor="w", padx=6)
        self.role = ctk.CTkComboBox(frame, values=list(ROLES))
        self.role.pack(fill="x", padx=6, pady=(2, 8))
        self.role.set(str(self.user.get("role") or "user"))

        self.password: ctk.CTkEntry | None = None
        self.confirm: ctk.CTkEntry | None = None
        if not user:
            ctk.CTkLabel(frame, text="Temporary password").pack(anchor="w", padx=6)
            self.password = ctk.CTkEntry(frame, show="*")
            self.password.pack(fill="x", padx=6, pady=(2, 8))
            ctk.CTkLabel(frame, text="Confirm password").pack(anchor="w", padx=6)
            self.confirm = ctk.CTkEntry(frame, show="*")
            self.confirm.pack(fill="x", padx=6, pady=(2, 8))

        self.active = ctk.BooleanVar(value=bool(self.user.get("is_active", True)))
        ctk.CTkCheckBox(frame, text="Active account", variable=self.active).pack(
            anchor="w", padx=6, pady=8
        )
        ctk.CTkButton(frame, text="Save", command=self._save).pack(
            fill="x", padx=6, pady=12
        )

    def _save(self) -> None:
        payload = {field: entry.get().strip() for field, entry in self.entries.items()}
        payload["role"] = self.role.get()
        payload["is_active"] = self.active.get()
        if self.password is not None and self.confirm is not None:
            if self.password.get() != self.confirm.get():
                messagebox.showerror("User", "Passwords do not match.", parent=self)
                return
            payload["password"] = self.password.get()
        self.result = payload
        self.destroy()


class UserManagementDialog(Modal):
    def __init__(self, master: Any, actor: Mapping[str, Any]) -> None:
        super().__init__(master, "Manage Users", "980x580")
        self.actor = actor
        self.tree = ttk.Treeview(
            self,
            columns=("ID", "Username", "Name", "Role", "Active", "Must Change", "Last Login"),
            show="headings",
        )
        for column, width in (
            ("ID", 55),
            ("Username", 140),
            ("Name", 180),
            ("Role", 130),
            ("Active", 75),
            ("Must Change", 100),
            ("Last Login", 170),
        ):
            self.tree.heading(column, text=column)
            self.tree.column(column, width=width, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=12, pady=(12, 6))
        controls = ctk.CTkFrame(self)
        controls.pack(fill="x", padx=12, pady=(6, 12))
        for label, command in (
            ("Add", self._add),
            ("Edit", self._edit),
            ("Reset Password", self._reset_password),
            ("Delete", self._delete),
            ("Refresh", self.refresh),
        ):
            ctk.CTkButton(controls, text=label, command=command, width=130).pack(
                side="left", padx=4, pady=8
            )
        self.refresh()

    def refresh(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for row in list_users():
            self.tree.insert(
                "",
                "end",
                values=(
                    row["id"],
                    row["username"],
                    row.get("full_name") or "",
                    row["role"],
                    "Yes" if row["is_active"] else "No",
                    "Yes" if row["must_change_pwd"] else "No",
                    row.get("last_login") or "",
                ),
            )

    def _add(self) -> None:
        dialog = UserEditorDialog(self)
        self.wait_window(dialog)
        if not dialog.result:
            return
        payload = dict(dialog.result)
        password = payload.pop("password", "")
        try:
            create_user(payload, password, str(self.actor["username"]))
        except ValueError as exc:
            messagebox.showerror("User", str(exc), parent=self)
        self.refresh()

    def _edit(self) -> None:
        user_id = _selected_id(self.tree)
        if user_id is None:
            messagebox.showinfo("User", "Select a user.", parent=self)
            return
        user = next((row for row in list_users() if row["id"] == user_id), None)
        dialog = UserEditorDialog(self, user)
        self.wait_window(dialog)
        if dialog.result:
            try:
                update_user(user_id, dialog.result, str(self.actor["username"]))
            except ValueError as exc:
                messagebox.showerror("User", str(exc), parent=self)
            self.refresh()

    def _reset_password(self) -> None:
        user_id = _selected_id(self.tree)
        if user_id is None:
            messagebox.showinfo("User", "Select a user.", parent=self)
            return
        first = simpledialog.askstring(
            "Reset Password", "Enter a temporary password:", show="*", parent=self
        )
        if not first:
            return
        second = simpledialog.askstring(
            "Reset Password", "Confirm temporary password:", show="*", parent=self
        )
        if first != second:
            messagebox.showerror("User", "Passwords do not match.", parent=self)
            return
        try:
            reset_user_password(user_id, first, str(self.actor["username"]))
        except ValueError as exc:
            messagebox.showerror("User", str(exc), parent=self)
        else:
            messagebox.showinfo(
                "User",
                "Password reset. The user must change it at next login.",
                parent=self,
            )
        self.refresh()

    def _delete(self) -> None:
        user_id = _selected_id(self.tree)
        if user_id is None:
            messagebox.showinfo("User", "Select a user.", parent=self)
            return
        if not messagebox.askyesno(
            "Delete User", "Permanently delete the selected user?", parent=self
        ):
            return
        try:
            delete_user(
                user_id,
                int(self.actor["id"]),
                str(self.actor["username"]),
            )
        except ValueError as exc:
            messagebox.showerror("User", str(exc), parent=self)
        self.refresh()


class CommissionEditorDialog(Modal):
    def __init__(
        self,
        master: Any,
        commission: Mapping[str, Any] | None = None,
    ) -> None:
        super().__init__(
            master, "Edit Commission" if commission else "Add Commission", "560x650"
        )
        self.commission = dict(commission or {})
        self.staff_rows = list_staffs()
        self.staff_by_label = {
            f"{row['name']} (ID {row['id']})": int(row["id"]) for row in self.staff_rows
        }
        frame = ctk.CTkScrollableFrame(self)
        frame.pack(fill="both", expand=True, padx=18, pady=18)

        ctk.CTkLabel(frame, text="Staff").pack(anchor="w", padx=6)
        labels = list(self.staff_by_label) or ["No staff available"]
        self.staff = ctk.CTkComboBox(frame, values=labels)
        self.staff.pack(fill="x", padx=6, pady=(2, 8))
        current_staff_id = self.commission.get("staff_id")
        for label, staff_id in self.staff_by_label.items():
            if staff_id == current_staff_id:
                self.staff.set(label)
                break

        ctk.CTkLabel(frame, text="Bill type").pack(anchor="w", padx=6)
        self.bill_type = ctk.CTkComboBox(frame, values=list(BILL_TYPES))
        self.bill_type.pack(fill="x", padx=6, pady=(2, 8))
        self.bill_type.set(str(self.commission.get("bill_type") or "CS"))

        self.entries: dict[str, ctk.CTkEntry] = {}
        for field, label in (
            ("bill_no", "Bill number"),
            ("voucher_id", "Voucher ID (optional)"),
            ("total_amount", "Total amount (RM)"),
            ("commission_amount", "Commission amount (RM)"),
            ("bill_image_path", "Bill image path (optional)"),
            ("note", "Note"),
        ):
            ctk.CTkLabel(frame, text=label).pack(anchor="w", padx=6)
            entry = ctk.CTkEntry(frame)
            entry.pack(fill="x", padx=6, pady=(2, 8))
            entry.insert(0, str(self.commission.get(field) or ""))
            self.entries[field] = entry
        ctk.CTkButton(frame, text="Save", command=self._save).pack(
            fill="x", padx=6, pady=12
        )

    def _save(self) -> None:
        staff_id = self.staff_by_label.get(self.staff.get())
        payload: dict[str, Any] = {
            field: entry.get() for field, entry in self.entries.items()
        }
        payload["staff_id"] = staff_id
        payload["bill_type"] = self.bill_type.get()
        try:
            self.result = validate_commission(payload)
        except ValidationError as exc:
            messagebox.showerror("Commission", str(exc), parent=self)
            return
        self.destroy()


class CommissionManagementDialog(Modal):
    def __init__(self, master: Any, actor: str) -> None:
        super().__init__(master, "Commission Management", "1080x620")
        self.actor = actor
        self.rows: list[dict[str, Any]] = []
        self.tree = ttk.Treeview(
            self,
            columns=(
                "ID",
                "Created",
                "Staff",
                "Type",
                "Bill",
                "Voucher",
                "Total",
                "Commission",
            ),
            show="headings",
        )
        for column, width in (
            ("ID", 55),
            ("Created", 145),
            ("Staff", 170),
            ("Type", 60),
            ("Bill", 110),
            ("Voucher", 90),
            ("Total", 100),
            ("Commission", 110),
        ):
            self.tree.heading(column, text=column)
            self.tree.column(column, width=width, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=12, pady=(12, 6))
        self.total_label = ctk.CTkLabel(self, text="")
        self.total_label.pack(anchor="e", padx=16)
        controls = ctk.CTkFrame(self)
        controls.pack(fill="x", padx=12, pady=(6, 12))
        for label, command in (
            ("Add", self._add),
            ("Edit", self._edit),
            ("Delete", self._delete),
            ("Export CSV", self._export),
            ("Refresh", self.refresh),
        ):
            ctk.CTkButton(controls, text=label, command=command, width=120).pack(
                side="left", padx=4, pady=8
            )
        self.refresh()

    def refresh(self) -> None:
        self.rows = list_commissions()
        self.tree.delete(*self.tree.get_children())
        for row in self.rows:
            self.tree.insert(
                "",
                "end",
                values=(
                    row["id"],
                    row.get("created_at") or "",
                    row.get("staff_name") or "",
                    row.get("bill_type") or "",
                    row.get("bill_no") or "",
                    row.get("voucher_id") or "",
                    f"{float(row.get('total_amount') or 0):.2f}",
                    f"{float(row.get('commission_amount') or 0):.2f}",
                ),
            )
        total = sum(float(row.get("commission_amount") or 0) for row in self.rows)
        self.total_label.configure(text=f"Total commission: RM {total:,.2f}")

    def _add(self) -> None:
        if not list_staffs():
            messagebox.showinfo(
                "Commission", "Add a staff member first.", parent=self
            )
            return
        dialog = CommissionEditorDialog(self)
        self.wait_window(dialog)
        if dialog.result:
            try:
                create_commission(dialog.result, self.actor)
            except ValueError as exc:
                messagebox.showerror("Commission", str(exc), parent=self)
            self.refresh()

    def _edit(self) -> None:
        commission_id = _selected_id(self.tree)
        if commission_id is None:
            messagebox.showinfo("Commission", "Select a record.", parent=self)
            return
        dialog = CommissionEditorDialog(self, get_commission(commission_id))
        self.wait_window(dialog)
        if dialog.result:
            try:
                update_commission(commission_id, dialog.result, self.actor)
            except ValueError as exc:
                messagebox.showerror("Commission", str(exc), parent=self)
            self.refresh()

    def _delete(self) -> None:
        commission_id = _selected_id(self.tree)
        if commission_id is None:
            messagebox.showinfo("Commission", "Select a record.", parent=self)
            return
        if not messagebox.askyesno(
            "Commission", "Delete the selected commission record?", parent=self
        ):
            return
        try:
            delete_commission(commission_id, self.actor)
        except ValueError as exc:
            messagebox.showerror("Commission", str(exc), parent=self)
        self.refresh()

    def _export(self) -> None:
        destination = filedialog.asksaveasfilename(
            parent=self,
            title="Export commissions",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
        )
        if destination:
            path = export_commissions_csv(destination, self.rows)
            messagebox.showinfo("Export", f"Saved to:\n{path}", parent=self)


class AuditLogDialog(Modal):
    def __init__(self, master: Any) -> None:
        super().__init__(master, "Audit Log", "1080x620")
        tree = ttk.Treeview(
            self,
            columns=("Time", "User", "Action", "Entity", "ID", "Details"),
            show="headings",
        )
        for column, width in (
            ("Time", 155),
            ("User", 120),
            ("Action", 170),
            ("Entity", 90),
            ("ID", 90),
            ("Details", 380),
        ):
            tree.heading(column, text=column)
            tree.column(column, width=width, anchor="w")
        tree.pack(fill="both", expand=True, padx=12, pady=12)
        for row in list_audit_entries():
            tree.insert(
                "",
                "end",
                values=(
                    row.get("occurred_at") or "",
                    row.get("username") or "",
                    row.get("action") or "",
                    row.get("entity_type") or "",
                    row.get("entity_id") or "",
                    row.get("details") or "",
                ),
            )


class VoucherApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.ready = False
        self.title(APP_NAME)
        self.geometry("1420x820")
        self.minsize(1120, 680)
        self.user: dict[str, Any] | None = None
        self.current_rows: list[dict[str, Any]] = []
        self.rows_by_id: dict[str, dict[str, Any]] = {}

        if not has_admin_user() and not self._first_run_setup():
            self.destroy()
            return
        if not self._login_flow():
            self.destroy()
            return

        self._configure_tree_style()
        self._build_header()
        self._build_filters()
        self._build_table()
        self._build_actions()
        self.perform_search()
        self.ready = True

    def _first_run_setup(self) -> bool:
        self.withdraw()
        dialog = FirstRunAdminDialog(self)
        self.wait_window(dialog)
        self.deiconify()
        return bool(dialog.result)

    def _login_flow(self) -> bool:
        self.withdraw()
        dialog = LoginDialog(self)
        self.wait_window(dialog)
        if not dialog.result:
            return False
        self.user = dict(dialog.result)
        if self.user.get("must_change_pwd"):
            password_dialog = PasswordDialog(self, self.user, forced=True)
            self.wait_window(password_dialog)
            if not password_dialog.result:
                return False
            self.user["must_change_pwd"] = 0
        self.deiconify()
        self.title(
            f"{APP_NAME} — {self.user['username']} ({self.user['role']})"
        )
        return True

    def _configure_tree_style(self) -> None:
        style = ttk.Style(self)
        style.configure("Treeview", rowheight=28, font=(FONT_FAMILY, 10))
        style.configure("Treeview.Heading", font=(FONT_FAMILY, 10, "bold"))

    def _build_header(self) -> None:
        assert self.user is not None
        header = ctk.CTkFrame(self)
        header.pack(fill="x", padx=10, pady=(10, 4))
        ctk.CTkLabel(
            header, text=APP_NAME, font=(FONT_FAMILY, 22, "bold")
        ).pack(side="left", padx=12, pady=10)
        ctk.CTkLabel(
            header,
            text=f"Signed in: {self.user['username']} · {self.user['role']}",
        ).pack(side="right", padx=10)
        ctk.CTkButton(
            header, text="Change Password", width=135, command=self._change_password
        ).pack(side="right", padx=4)

    def _build_filters(self) -> None:
        filters = ctk.CTkFrame(self)
        filters.pack(fill="x", padx=10, pady=4)
        self.filter_entries: dict[str, ctk.CTkEntry] = {}
        for field, placeholder, width in (
            ("voucher_id", "Voucher ID", 120),
            ("customer_name", "Customer name", 180),
            ("contact_number", "Contact", 130),
            ("recipient", "Recipient", 150),
            ("date_from", "Date from YYYY-MM-DD", 170),
            ("date_to", "Date to YYYY-MM-DD", 170),
        ):
            entry = ctk.CTkEntry(filters, placeholder_text=placeholder, width=width)
            entry.pack(side="left", padx=4, pady=10)
            self.filter_entries[field] = entry
        self.status_filter = ctk.CTkComboBox(
            filters, values=["All", *VOUCHER_STATUSES], width=130
        )
        self.status_filter.set("All")
        self.status_filter.pack(side="left", padx=4)
        ctk.CTkButton(filters, text="Search", width=90, command=self.perform_search).pack(
            side="left", padx=4
        )
        ctk.CTkButton(
            filters, text="Reset", width=80, fg_color="#666666", command=self.reset_filters
        ).pack(side="left", padx=4)

    def _build_table(self) -> None:
        frame = ctk.CTkFrame(self)
        frame.pack(fill="both", expand=True, padx=10, pady=4)
        columns = (
            "ID",
            "Date",
            "Customer",
            "Contact",
            "Units",
            "Recipient",
            "Technician",
            "Status",
            "Amount",
        )
        self.tree = ttk.Treeview(frame, columns=columns, show="headings")
        widths = (90, 105, 180, 130, 60, 140, 150, 100, 90)
        for column, width in zip(columns, widths):
            self.tree.heading(column, text=column)
            self.tree.column(column, width=width, anchor="w")
        y_scroll = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", lambda _event: self._double_click())

    def _build_actions(self) -> None:
        assert self.user is not None
        actions = ctk.CTkFrame(self)
        actions.pack(fill="x", padx=10, pady=(4, 10))

        definitions = (
            ("Add Voucher", self.add_voucher, "voucher.create"),
            ("Edit", self.edit_voucher, "voucher.edit"),
            ("Update Status", self.update_status, "voucher.status"),
            ("Delete", self.delete_selected, "voucher.delete"),
            ("Open PDF", self.open_pdf, None),
            ("Regenerate PDF", self.regenerate_pdf, "voucher.edit"),
            ("Export CSV", self.export_vouchers, "voucher.export"),
            ("Staff", self.manage_staff, "staff.manage"),
            ("Users", self.manage_users, "user.manage"),
            ("Commissions", self.manage_commissions, "commission.manage"),
            ("Backup", self.backup_data, "backup.manage"),
            ("Restore", self.restore_data, "backup.manage"),
            ("Audit Log", self.open_audit_log, "user.manage"),
            ("Base ID", self.modify_base_id, "settings.manage"),
        )
        for index, (label, command, permission) in enumerate(definitions):
            button = ctk.CTkButton(actions, text=label, command=command, width=115)
            button.grid(row=index // 7, column=index % 7, padx=4, pady=5, sticky="ew")
            if permission and not has_permission(self.user, permission):
                button.configure(state="disabled")
        for column in range(7):
            actions.grid_columnconfigure(column, weight=1)

    def _filters(self) -> dict[str, str]:
        result = {field: entry.get().strip() for field, entry in self.filter_entries.items()}
        result["status"] = self.status_filter.get()
        return result

    def perform_search(self) -> None:
        try:
            self.current_rows = search_vouchers(self._filters())
        except Exception as exc:
            logger.exception("Voucher search failed")
            messagebox.showerror("Search", str(exc), parent=self)
            return
        self.rows_by_id = {str(row["voucher_id"]): row for row in self.current_rows}
        self.tree.delete(*self.tree.get_children())
        for row in self.current_rows:
            voucher_id = str(row["voucher_id"])
            self.tree.insert(
                "",
                "end",
                iid=voucher_id,
                values=(
                    voucher_id,
                    str(row.get("created_at") or "")[:10],
                    row.get("customer_name") or "",
                    row.get("contact_number") or "",
                    row.get("units") or 1,
                    row.get("recipient") or "",
                    row.get("technician_name") or "",
                    row.get("status") or "",
                    f"{float(row.get('amount_rm') or 0):.2f}",
                ),
            )

    def reset_filters(self) -> None:
        for entry in self.filter_entries.values():
            entry.delete(0, "end")
        self.status_filter.set("All")
        self.perform_search()

    def _selected_voucher(self) -> dict[str, Any] | None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Voucher", "Select a voucher first.", parent=self)
            return None
        voucher_id = str(selection[0])
        return get_voucher(voucher_id)

    def _double_click(self) -> None:
        assert self.user is not None
        if has_permission(self.user, "voucher.edit"):
            self.edit_voucher()
        else:
            self.open_pdf()

    def _save_pdf_for(self, voucher: Mapping[str, Any]) -> str:
        assert self.user is not None
        path = generate_voucher_pdf(voucher)
        update_voucher_pdf_path(
            str(voucher["voucher_id"]), path, str(self.user["username"])
        )
        return path

    def add_voucher(self) -> None:
        assert self.user is not None
        dialog = VoucherDialog(self)
        self.wait_window(dialog)
        if not dialog.result:
            return
        try:
            voucher = create_voucher(dialog.result, str(self.user["username"]))
            self._save_pdf_for(voucher)
        except (ValueError, ValidationError, OSError) as exc:
            logger.exception("Voucher creation failed")
            messagebox.showerror("Voucher", str(exc), parent=self)
        else:
            messagebox.showinfo(
                "Voucher", f"Voucher {voucher['voucher_id']} created.", parent=self
            )
        self.perform_search()

    def edit_voucher(self) -> None:
        assert self.user is not None
        voucher = self._selected_voucher()
        if not voucher:
            return
        dialog = VoucherDialog(self, voucher=voucher)
        self.wait_window(dialog)
        if not dialog.result:
            return
        try:
            updated = update_voucher(
                str(voucher["voucher_id"]),
                dialog.result,
                str(self.user["username"]),
            )
            self._save_pdf_for(updated)
        except (ValueError, ValidationError, OSError) as exc:
            logger.exception("Voucher update failed")
            messagebox.showerror("Voucher", str(exc), parent=self)
        self.perform_search()

    def update_status(self) -> None:
        assert self.user is not None
        voucher = self._selected_voucher()
        if not voucher:
            return
        dialog = StatusDialog(self, voucher)
        self.wait_window(dialog)
        if not dialog.result:
            return
        try:
            updated = update_voucher(
                str(voucher["voucher_id"]),
                dialog.result,
                str(self.user["username"]),
            )
            self._save_pdf_for(updated)
        except (ValueError, ValidationError, OSError) as exc:
            logger.exception("Status update failed")
            messagebox.showerror("Voucher", str(exc), parent=self)
        self.perform_search()

    def delete_selected(self) -> None:
        assert self.user is not None
        voucher = self._selected_voucher()
        if not voucher:
            return
        if not messagebox.askyesno(
            "Delete Voucher",
            f"Permanently delete voucher {voucher['voucher_id']}?",
            parent=self,
        ):
            return
        try:
            pdf_path = delete_voucher(
                str(voucher["voucher_id"]), str(self.user["username"])
            )
            if pdf_path:
                target = Path(pdf_path).resolve()
                pdf_root = Path(PDF_DIR).resolve()
                if target.is_file() and target.is_relative_to(pdf_root):
                    target.unlink()
        except (ValueError, OSError) as exc:
            messagebox.showerror("Voucher", str(exc), parent=self)
        self.perform_search()

    def open_pdf(self) -> None:
        voucher = self._selected_voucher()
        if not voucher:
            return
        path = str(voucher.get("pdf_path") or "")
        try:
            if not path or not Path(path).is_file():
                path = self._save_pdf_for(voucher)
            open_local_file(path)
        except Exception as exc:
            logger.exception("Opening PDF failed")
            messagebox.showerror("PDF", str(exc), parent=self)

    def regenerate_pdf(self) -> None:
        voucher = self._selected_voucher()
        if not voucher:
            return
        try:
            path = self._save_pdf_for(voucher)
        except Exception as exc:
            logger.exception("PDF regeneration failed")
            messagebox.showerror("PDF", str(exc), parent=self)
        else:
            messagebox.showinfo("PDF", f"PDF regenerated:\n{path}", parent=self)
        self.perform_search()

    def export_vouchers(self) -> None:
        destination = filedialog.asksaveasfilename(
            parent=self,
            title="Export displayed vouchers",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
        )
        if destination:
            path = export_vouchers_csv(destination, self.current_rows)
            messagebox.showinfo("Export", f"Saved to:\n{path}", parent=self)

    def manage_staff(self) -> None:
        assert self.user is not None
        dialog = StaffManagementDialog(
            self, str(self.user["username"]), on_change=self.perform_search
        )
        self.wait_window(dialog)

    def manage_users(self) -> None:
        assert self.user is not None
        dialog = UserManagementDialog(self, self.user)
        self.wait_window(dialog)

    def manage_commissions(self) -> None:
        assert self.user is not None
        dialog = CommissionManagementDialog(self, str(self.user["username"]))
        self.wait_window(dialog)

    def backup_data(self) -> None:
        destination = filedialog.asksaveasfilename(
            parent=self,
            title="Create SVMS backup",
            defaultextension=".zip",
            filetypes=[("SVMS backup", "*.zip")],
        )
        if not destination:
            return
        try:
            path = create_backup(destination)
        except Exception as exc:
            logger.exception("Backup failed")
            messagebox.showerror("Backup", str(exc), parent=self)
        else:
            messagebox.showinfo("Backup", f"Backup created:\n{path}", parent=self)

    def restore_data(self) -> None:
        source = filedialog.askopenfilename(
            parent=self,
            title="Restore SVMS backup",
            filetypes=[("SVMS backup", "*.zip")],
        )
        if not source:
            return
        if not messagebox.askyesno(
            "Restore Backup",
            "Restore this backup? A pre-restore backup will be created automatically, "
            "and the application will close afterwards.",
            parent=self,
        ):
            return
        try:
            pre_restore = restore_backup(source)
        except Exception as exc:
            logger.exception("Restore failed")
            messagebox.showerror("Restore", str(exc), parent=self)
            return
        messagebox.showinfo(
            "Restore Complete",
            f"Backup restored. Previous data was saved to:\n{pre_restore}\n\n"
            "Please reopen the application.",
            parent=self,
        )
        self.destroy()

    def open_audit_log(self) -> None:
        dialog = AuditLogDialog(self)
        self.wait_window(dialog)

    def modify_base_id(self) -> None:
        assert self.user is not None
        value = simpledialog.askinteger(
            "Base Voucher ID",
            "Set the minimum number used for future vouchers. Existing IDs are unchanged:",
            parent=self,
            minvalue=1,
            maxvalue=999_999_999,
        )
        if value is None:
            return
        try:
            set_base_voucher_id(value, str(self.user["username"]))
        except ValueError as exc:
            messagebox.showerror("Base Voucher ID", str(exc), parent=self)
        else:
            messagebox.showinfo(
                "Base Voucher ID",
                "Future voucher numbering has been updated.",
                parent=self,
            )

    def _change_password(self) -> None:
        assert self.user is not None
        dialog = PasswordDialog(self, self.user, forced=False)
        self.wait_window(dialog)
        if dialog.result:
            messagebox.showinfo("Password", "Password changed.", parent=self)
