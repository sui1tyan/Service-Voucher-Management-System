from __future__ import annotations

import csv
import sqlite3
import zipfile
from pathlib import Path

import pytest

import backup_utils
import database
import pdf_utils
from export_utils import export_commissions_csv, export_vouchers_csv


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


def test_commission_csv_export_includes_every_filtered_page(
    admin: dict, tmp_path: Path
) -> None:
    staff = database.create_staff(
        {"staff_id_opt": "CSV-1", "name": "CSV Staff", "position": "Technician"},
        "owner",
    )
    for index in range(105):
        database.create_commission(
            {
                "staff_id": staff["id"],
                "bill_type": "CS",
                "bill_no": f"EXPORT-{index:03d}",
                "total_amount": 50,
                "commission_amount": 5,
            },
            "owner",
        )
    database.create_commission(
        {
            "staff_id": staff["id"],
            "bill_type": "INV",
            "bill_no": "DO-NOT-EXPORT",
            "total_amount": 50,
            "commission_amount": 5,
        },
        "owner",
    )

    filters = {"query": "export-", "bill_type": "CS"}
    assert len(database.search_commissions(filters, limit=100)) == 100
    path = Path(
        export_commissions_csv(
            tmp_path / "commissions",
            database.iter_commissions(filters),
        )
    )
    with path.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    assert len(rows) == 105
    assert all(row["bill_no"].startswith("EXPORT-") for row in rows)


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


def test_restore_accepts_nested_metadata_free_legacy_backup(
    admin: dict,
    voucher_payload: dict,
    isolated_db: Path,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    backup_dir = tmp_path / "backups"
    restored_images = tmp_path / "restored-images"
    monkeypatch.setattr(backup_utils, "BACKUP_DIR", backup_dir)
    monkeypatch.setattr(backup_utils, "PDF_DIR", tmp_path / "pdfs")
    monkeypatch.setattr(backup_utils, "STAFFS_ROOT", tmp_path / "staffs")
    monkeypatch.setattr(backup_utils, "LEGACY_IMAGES_DIR", restored_images)

    database.create_voucher(voucher_payload, "owner")
    legacy_database = tmp_path / "legacy-vouchers.db"
    source = database.get_conn()
    try:
        with sqlite3.connect(legacy_database) as destination:
            source.backup(destination)
    finally:
        source.close()

    legacy_backup = tmp_path / "legacy-layout.zip"
    with zipfile.ZipFile(legacy_backup, "w") as archive:
        archive.write(legacy_database, "ServiceVoucherApp/vouchers.db")
        archive.writestr("ServiceVoucherApp/images/legacy-receipt.txt", "legacy")

    backup_utils.validate_backup(legacy_backup)
    database.delete_voucher("41000", "owner")
    backup_utils.restore_backup(legacy_backup)

    assert database.get_voucher("41000") is not None
    assert (restored_images / "legacy-receipt.txt").read_text() == "legacy"
    with sqlite3.connect(isolated_db) as conn:
        assert conn.execute("PRAGMA user_version").fetchone()[0] == database.SCHEMA_VERSION
