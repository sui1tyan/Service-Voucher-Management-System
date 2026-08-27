from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

import database


def test_first_run_admin_and_authentication(isolated_db: Path) -> None:
    assert not database.has_admin_user()
    admin = database.create_initial_admin(
        "owner", "SecureAdmin!2026", "System Owner"
    )
    assert admin["role"] == "admin"
    assert database.has_admin_user()

    authenticated = database.authenticate_user("OWNER", "SecureAdmin!2026")
    assert authenticated is not None
    assert authenticated["username"] == "owner"
    assert "password_hash" not in authenticated
    assert database.authenticate_user("owner", "wrong") is None

    with pytest.raises(ValueError, match="already exists"):
        database.create_initial_admin("second", "AnotherSecure!2026")


def test_user_management_preserves_an_active_admin(admin: dict) -> None:
    user = database.create_user(
        {
            "username": "assistant",
            "role": "sales assistant",
            "full_name": "Sales Assistant",
        },
        "TemporaryPass!2026",
        "owner",
    )
    assert user["must_change_pwd"] == 1
    database.reset_user_password(
        user["id"], "SecondTemporary!2026", "owner", force_change=True
    )
    login = database.authenticate_user("assistant", "SecondTemporary!2026")
    assert login is not None and login["must_change_pwd"] == 1

    with pytest.raises(ValueError, match="At least one active administrator"):
        database.update_user(
            admin["id"],
            {
                "username": "owner",
                "role": "user",
                "is_active": True,
            },
            "owner",
        )
    with pytest.raises(ValueError, match="own account"):
        database.delete_user(admin["id"], admin["id"], "owner")


def test_voucher_staff_and_commission_workflow(
    admin: dict, voucher_payload: dict
) -> None:
    staff = database.create_staff(
        {
            "staff_id_opt": "S001",
            "name": "Alice",
            "position": "Technician",
            "phone": "016-1234567",
            "email": "alice@example.test",
        },
        "owner",
    )
    assert database.list_staffs_names() == ["Alice"]

    voucher = database.create_voucher(voucher_payload, "owner")
    assert voucher["voucher_id"] == "41000"
    assert voucher["contact_number"] == "0123456789"
    assert database.get_next_voucher_id() == "41001"

    matches = database.search_vouchers(
        {"customer_name": "test", "status": "Pending"}
    )
    assert [row["voucher_id"] for row in matches] == ["41000"]

    updated_payload = dict(voucher_payload)
    updated_payload.update(
        {
            "status": "Completed",
            "solution": "Replaced the damaged power adapter.",
            "technician_id": "S001",
            "technician_name": "Alice",
            "amount_rm": "120.50",
            "tech_commission": "12.50",
        }
    )
    updated = database.update_voucher("41000", updated_payload, "owner")
    assert updated["status"] == "Completed"
    assert updated["amount_rm"] == 120.5

    commission = database.create_commission(
        {
            "staff_id": staff["id"],
            "bill_type": "CS",
            "bill_no": "CS-1001",
            "voucher_id": "41000",
            "total_amount": 120.5,
            "commission_amount": 12.5,
            "note": "Repair commission",
        },
        "owner",
    )
    assert commission["staff_name"] == "Alice"
    assert database.list_commissions()[0]["voucher_id"] == "41000"
    with pytest.raises(ValueError, match="commission records"):
        database.delete_staff(staff["id"], "owner")

    database.delete_commission(commission["id"], "owner")
    database.delete_staff(staff["id"], "owner")
    database.delete_voucher("41000", "owner")
    assert database.get_voucher("41000") is None
    assert len(database.list_audit_entries()) >= 7


def test_base_voucher_id_does_not_renumber_existing_records(
    admin: dict, voucher_payload: dict
) -> None:
    first = database.create_voucher(voucher_payload, "owner")
    database.set_base_voucher_id(50_000, "owner")
    second = database.create_voucher(voucher_payload, "owner")
    assert first["voucher_id"] == "41000"
    assert second["voucher_id"] == "50000"


def test_legacy_schema_is_migrated(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    db_path = tmp_path / "legacy.db"
    with sqlite3.connect(db_path) as conn:
        conn.executescript(
            """
            CREATE TABLE users (
                id INTEGER PRIMARY KEY,
                username TEXT UNIQUE,
                role TEXT,
                password_hash BLOB,
                is_active INTEGER DEFAULT 1,
                must_change_pwd INTEGER DEFAULT 0,
                created_at TEXT,
                updated_at TEXT
            );
            CREATE TABLE vouchers (
                id INTEGER PRIMARY KEY,
                voucher_id TEXT UNIQUE,
                created_at TEXT,
                customer_name TEXT,
                contact_number TEXT,
                units INTEGER,
                particulars TEXT,
                problem TEXT,
                staff_name TEXT,
                status TEXT,
                recipient TEXT,
                solution TEXT,
                pdf_path TEXT,
                technician_id TEXT,
                technician_name TEXT,
                ref_bill TEXT,
                ref_bill_date TEXT,
                amount_rm REAL,
                tech_commission REAL
            );
            """
        )
    monkeypatch.setattr(database, "DB_FILE", db_path)
    database.init_db()
    with sqlite3.connect(db_path) as conn:
        user_columns = {
            row[1] for row in conn.execute("PRAGMA table_info(users)").fetchall()
        }
        voucher_columns = {
            row[1] for row in conn.execute("PRAGMA table_info(vouchers)").fetchall()
        }
        tables = {
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        }
    assert {"full_name", "last_login", "email"}.issubset(user_columns)
    assert {"updated_at", "created_by", "updated_by"}.issubset(voucher_columns)
    assert {"staffs", "commissions", "audit_log", "settings"}.issubset(tables)
