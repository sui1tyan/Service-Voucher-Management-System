from __future__ import annotations

from pathlib import Path

import pytest

import database


@pytest.fixture
def isolated_db(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    db_path = tmp_path / "test_vouchers.db"
    monkeypatch.setattr(database, "DB_FILE", db_path)
    database.init_db()
    return db_path


@pytest.fixture
def admin(isolated_db: Path) -> dict:
    return database.create_initial_admin(
        "owner",
        "SecureAdmin!2026",
        "System Owner",
    )


@pytest.fixture
def voucher_payload() -> dict:
    return {
        "customer_name": "Test Customer",
        "contact_number": "012-3456789",
        "units": 1,
        "particulars": "Notebook computer and charger",
        "problem": "Does not power on",
        "staff_name": "Alice",
        "recipient": "Alice",
        "status": "Pending",
        "solution": "",
        "technician_id": "",
        "technician_name": "",
        "ref_bill": "",
        "ref_bill_date": "",
        "amount_rm": 0,
        "tech_commission": 0,
    }
