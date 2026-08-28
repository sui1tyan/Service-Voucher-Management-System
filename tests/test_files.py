from __future__ import annotations

import csv
from pathlib import Path

import pytest

import backup_utils
import database
import pdf_utils
from export_utils import export_vouchers_csv


def test_pdf_generation(
    admin: dict,
    voucher_payload: dict,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    voucher = database.create_voucher(voucher_payload, "owner")
    output_dir = tmp_path / "pdfs"
    monkeypatch.setattr(pdf_utils, "PDF_DIR", output_dir)
    path = Path(pdf_utils.generate_voucher_pdf(voucher))
    assert path.is_file()
    assert path.name == "voucher_41000.pdf"
    assert path.stat().st_size > 1_000


def test_csv_export(
    admin: dict, voucher_payload: dict, tmp_path: Path
) -> None:
    for index in range(101):
        payload = dict(voucher_payload)
        payload["customer_name"] = f"Export Customer {index}"
        database.create_voucher(payload, "owner")

    filters = {"customer_name": "export customer"}
    assert len(database.search_vouchers(filters, limit=100)) == 100
    path = Path(
        export_vouchers_csv(tmp_path / "vouchers", database.iter_vouchers(filters))
    )
    assert path.suffix == ".csv"
    with path.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    assert len(rows) == 101
    assert {row["voucher_id"] for row in rows} == {
        str(voucher_id) for voucher_id in range(41_000, 41_101)
    }
    assert all(row["customer_name"].startswith("Export Customer") for row in rows)


def test_backup_validation_and_restore(
    admin: dict,
    voucher_payload: dict,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    pdf_dir = tmp_path / "pdfs"
    staff_dir = tmp_path / "staffs"
    backup_dir = tmp_path / "backups"
    pdf_dir.mkdir()
    staff_dir.mkdir()
    monkeypatch.setattr(backup_utils, "PDF_DIR", pdf_dir)
    monkeypatch.setattr(backup_utils, "STAFFS_ROOT", staff_dir)
    monkeypatch.setattr(backup_utils, "BACKUP_DIR", backup_dir)

    database.create_voucher(voucher_payload, "owner")
    (pdf_dir / "voucher_41000.pdf").write_bytes(b"%PDF-test")
    backup = Path(backup_utils.create_backup(backup_dir / "snapshot.zip"))
    assert backup.is_file()
    backup_utils.validate_backup(backup)

    database.delete_voucher("41000", "owner")
    assert database.get_voucher("41000") is None
    pre_restore = Path(backup_utils.restore_backup(backup))
    assert pre_restore.is_file()
    assert database.get_voucher("41000") is not None
