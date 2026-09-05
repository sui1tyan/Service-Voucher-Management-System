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


def test_voucher_pagination_and_count_use_identical_filters(
    admin: dict, voucher_payload: dict
) -> None:
    for index in range(130):
        payload = dict(voucher_payload)
        is_match = index < 125
        payload.update(
            {
                "customer_name": f"{'Match' if is_match else 'Other'} Customer {index}",
                "technician_name": "Alice" if is_match else "Bob",
                "ref_bill": "BATCH-A" if is_match else "BATCH-B",
            }
        )
        database.create_voucher(payload, "owner")

    filters = {
        "customer_name": "match customer",
        "technician_name": "alice",
        "ref_bill": "batch-a",
        "status": "Pending",
    }
    first_page = database.search_vouchers(filters, limit=100, offset=0)
    second_page = database.search_vouchers(filters, limit=100, offset=100)

    assert database.count_vouchers(filters) == 125
    assert len(first_page) == 100
    assert len(second_page) == 25
    assert {row["voucher_id"] for row in first_page}.isdisjoint(
        row["voucher_id"] for row in second_page
    )


def test_base_voucher_id_does_not_renumber_existing_records(
    admin: dict, voucher_payload: dict
) -> None:
    first = database.create_voucher(voucher_payload, "owner")
    database.set_base_voucher_id(50_000, "owner")
    second = database.create_voucher(voucher_payload, "owner")
    assert first["voucher_id"] == "41000"
    assert second["voucher_id"] == "50000"


def test_voucher_id_editing_uses_a_monotonic_next_id(
    admin: dict, voucher_payload: dict
) -> None:
    first = database.create_voucher(voucher_payload, "owner")
    edited_payload = dict(voucher_payload)
    edited_payload["voucher_id"] = "42000"
    edited = database.update_voucher(first["voucher_id"], edited_payload, "owner")

    assert edited["voucher_id"] == "42000"
    assert database.get_voucher("41000") is None
    assert database.get_next_voucher_id() == "42001"

    database.delete_voucher("42000", "owner")
    assert database.get_next_voucher_id() == "42001"
    replacement = database.create_voucher(voucher_payload, "owner")
    assert replacement["voucher_id"] == "42001"

    lower_manual = dict(voucher_payload)
    lower_manual["voucher_id"] = "41500"
    database.create_voucher(lower_manual, "owner")
    assert database.get_next_voucher_id() == "42002"

    duplicate = dict(voucher_payload)
    duplicate["voucher_id"] = "42001"
    with pytest.raises(ValueError, match="already exists"):
        database.create_voucher(duplicate, "owner")


def test_voucher_ids_are_sorted_numerically(
    admin: dict, voucher_payload: dict
) -> None:
    for voucher_id in ("9", "100", "20"):
        payload = dict(voucher_payload)
        payload["voucher_id"] = voucher_id
        database.create_voucher(payload, "owner")

    descending = database.search_vouchers(sort_direction="desc")
    ascending = database.search_vouchers(sort_direction="asc")
    assert [row["voucher_id"] for row in descending] == ["100", "20", "9"]
    assert [row["voucher_id"] for row in ascending] == ["9", "20", "100"]


def test_voucher_changes_automatically_synchronize_commission(
    admin: dict, voucher_payload: dict
) -> None:
    staff = database.create_staff(
        {
            "staff_id_opt": "TECH-7",
            "name": "Alex Technician",
            "position": "Technician",
        },
        "owner",
    )
    payload = dict(voucher_payload)
    payload.update(
        {
            "voucher_id": "43000",
            "technician_id": "TECH-7",
            "technician_name": "Alex Technician",
            "ref_bill": "INV-9001",
            "amount_rm": 250,
            "tech_commission": 25,
        }
    )
    database.create_voucher(payload, "owner")

    commissions = database.list_commissions()
    assert len(commissions) == 1
    assert commissions[0]["voucher_id"] == "43000"
    assert commissions[0]["bill_type"] == "INV"
    assert commissions[0]["total_amount"] == 250
    assert commissions[0]["commission_amount"] == 25
    assert commissions[0]["managed_by_voucher"] == 1

    payload.update(
        {
            "voucher_id": "43005",
            "ref_bill": "CS-9002",
            "amount_rm": 300,
            "tech_commission": 30,
        }
    )
    database.update_voucher("43000", payload, "owner")
    synchronized = database.list_commissions()
    assert len(synchronized) == 1
    assert synchronized[0]["voucher_id"] == "43005"
    assert synchronized[0]["bill_type"] == "CS"
    assert synchronized[0]["bill_no"] == "CS-9002"
    assert synchronized[0]["commission_amount"] == 30

    payload["ref_bill"] = ""
    database.update_voucher("43005", payload, "owner")
    assert database.list_commissions() == []

    manual = database.create_commission(
        {
            "staff_id": staff["id"],
            "bill_type": "CS",
            "bill_no": "MANUAL-1",
            "voucher_id": "43005",
            "total_amount": 100,
            "commission_amount": 10,
        },
        "owner",
    )
    payload["voucher_id"] = "43006"
    database.update_voucher("43005", payload, "owner")
    renamed_link = database.get_commission(manual["id"])
    assert renamed_link is not None
    assert renamed_link["voucher_id"] == "43006"

    database.delete_voucher("43006", "owner")
    preserved = database.get_commission(manual["id"])
    assert preserved is not None
    assert preserved["voucher_id"] is None
    assert preserved["managed_by_voucher"] == 0


def test_commission_and_audit_pagination(
    admin: dict, voucher_payload: dict
) -> None:
    staff = database.create_staff(
        {"staff_id_opt": "S-9", "name": "Paged Staff", "position": "Technician"},
        "owner",
    )
    for index in range(125):
        database.create_commission(
            {
                "staff_id": staff["id"],
                "bill_type": "CS",
                "bill_no": f"BATCH-A-{index:03d}",
                "total_amount": 10,
                "commission_amount": 1,
            },
            "owner",
        )
    for index in range(5):
        database.create_commission(
            {
                "staff_id": staff["id"],
                "bill_type": "INV",
                "bill_no": f"OTHER-{index:03d}",
                "total_amount": 20,
                "commission_amount": 2,
            },
            "owner",
        )

    filters = {"query": "batch-a", "bill_type": "CS"}
    first_page = database.search_commissions(filters, limit=100, offset=0)
    second_page = database.search_commissions(filters, limit=100, offset=100)
    assert database.count_commissions(filters) == 125
    assert database.sum_commissions(filters) == 125
    assert len(first_page) == 100
    assert len(second_page) == 25
    assert len(list(database.iter_commissions(filters))) == 125

    audit_total = database.count_audit_entries()
    first_audit_page = database.list_audit_entries(100, offset=0)
    second_audit_page = database.list_audit_entries(100, offset=100)
    assert audit_total >= 132
    assert len(first_audit_page) == 100
    assert len(second_audit_page) == audit_total - 100
    assert {row["id"] for row in first_audit_page}.isdisjoint(
        row["id"] for row in second_audit_page
    )


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
                technician_name TEXT
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
        commission_columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(commissions)").fetchall()
        }
        next_voucher_id = conn.execute(
            "SELECT value FROM settings WHERE key='next_vid'"
        ).fetchone()[0]
    assert {"full_name", "last_login", "email"}.issubset(user_columns)
    assert {
        "ref_bill",
        "ref_bill_date",
        "amount_rm",
        "tech_commission",
        "updated_at",
        "created_by",
        "updated_by",
    }.issubset(voucher_columns)
    assert {"staffs", "commissions", "audit_log", "settings"}.issubset(tables)
    assert "managed_by_voucher" in commission_columns
    assert next_voucher_id == "41000"
