"""Safe backup and restore helpers for local SVMS data."""

from __future__ import annotations

import json
import shutil
import sqlite3
import tempfile
import zipfile
from contextlib import closing
from datetime import datetime
from pathlib import Path, PurePosixPath

from config import BACKUP_DIR, PDF_DIR, STAFFS_ROOT, logger
from database import db_session


REQUIRED_TABLES = {"vouchers", "staffs", "settings", "users", "commissions"}


def _timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _resolved_destination(destination: str | Path | None) -> Path:
    if destination is None:
        path = BACKUP_DIR / f"svms_backup_{_timestamp()}.zip"
    else:
        path = Path(destination).expanduser()
        if path.is_dir():
            path = path / f"svms_backup_{_timestamp()}.zip"
        elif path.suffix.lower() != ".zip":
            path = path.with_suffix(".zip")
    path.parent.mkdir(parents=True, exist_ok=True)
    return path.resolve()


def create_backup(destination: str | Path | None = None) -> str:
    """Create a consistent SQLite snapshot plus PDF and staff attachments."""

    target = _resolved_destination(destination)
    with tempfile.TemporaryDirectory(prefix="svms-backup-") as temp_name:
        temp_dir = Path(temp_name)
        snapshot = temp_dir / "vouchers.db"
        with db_session() as source, closing(sqlite3.connect(snapshot)) as dest:
            source.backup(dest)

        metadata = {
            "format": 1,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "application": "Service Voucher Management System",
        }
        (temp_dir / "metadata.json").write_text(
            json.dumps(metadata, indent=2), encoding="utf-8"
        )

        temp_zip = target.with_suffix(target.suffix + ".part")
        with zipfile.ZipFile(temp_zip, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.write(snapshot, "vouchers.db")
            archive.write(temp_dir / "metadata.json", "metadata.json")
            for source_dir, archive_root in (
                (Path(PDF_DIR), "pdfs"),
                (Path(STAFFS_ROOT), "staffs"),
            ):
                if not source_dir.exists():
                    continue
                for item in source_dir.rglob("*"):
                    if item.is_file() and not item.name.endswith(".part"):
                        archive.write(item, str(PurePosixPath(archive_root, *item.relative_to(source_dir).parts)))
        temp_zip.replace(target)
    logger.info("Backup created at %s", target)
    return str(target)


def _safe_members(archive: zipfile.ZipFile) -> list[zipfile.ZipInfo]:
    members = archive.infolist()
    for member in members:
        path = PurePosixPath(member.filename)
        if path.is_absolute() or ".." in path.parts:
            raise ValueError("Backup contains an unsafe path.")
    return members


def _check_database(db_path: Path) -> None:
    try:
        with sqlite3.connect(db_path) as conn:
            integrity = conn.execute("PRAGMA integrity_check").fetchone()[0]
            if integrity != "ok":
                raise ValueError(f"Backup database failed integrity check: {integrity}")
            tables = {
                row[0]
                for row in conn.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                ).fetchall()
            }
    except sqlite3.DatabaseError as exc:
        raise ValueError("Backup does not contain a valid SQLite database.") from exc
    missing = REQUIRED_TABLES - tables
    if missing:
        raise ValueError(f"Backup database is missing: {', '.join(sorted(missing))}.")


def validate_backup(backup_path: str | Path) -> None:
    """Validate archive paths and the embedded SQLite database."""

    path = Path(backup_path)
    if not path.is_file():
        raise ValueError("Backup file was not found.")
    try:
        with zipfile.ZipFile(path) as archive:
            members = _safe_members(archive)
            if "vouchers.db" not in {member.filename for member in members}:
                raise ValueError("Backup does not contain vouchers.db.")
            with tempfile.TemporaryDirectory(prefix="svms-validate-") as temp_name:
                archive.extract("vouchers.db", temp_name)
                _check_database(Path(temp_name) / "vouchers.db")
    except zipfile.BadZipFile as exc:
        raise ValueError("Selected file is not a valid SVMS backup.") from exc


def restore_backup(backup_path: str | Path) -> str:
    """Restore a validated backup after automatically preserving current data."""

    path = Path(backup_path).resolve()
    validate_backup(path)
    pre_restore = create_backup(BACKUP_DIR / f"pre_restore_{_timestamp()}.zip")

    with tempfile.TemporaryDirectory(prefix="svms-restore-") as temp_name:
        temp_dir = Path(temp_name)
        with zipfile.ZipFile(path) as archive:
            _safe_members(archive)
            archive.extractall(temp_dir)

        restored_db = temp_dir / "vouchers.db"
        _check_database(restored_db)
        with closing(sqlite3.connect(restored_db)) as source, db_session() as destination:
            source.backup(destination)

        for archive_root, destination_root in (
            ("pdfs", Path(PDF_DIR)),
            ("staffs", Path(STAFFS_ROOT)),
        ):
            source_dir = temp_dir / archive_root
            if source_dir.exists():
                destination_root.mkdir(parents=True, exist_ok=True)
                shutil.copytree(source_dir, destination_root, dirs_exist_ok=True)

    logger.warning("Backup restored from %s; pre-restore backup: %s", path, pre_restore)
    return pre_restore
