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
from database import SCHEMA_VERSION, db_session, init_db

BACKUP_FORMAT = 2
LEGACY_REQUIRED_TABLES = {"vouchers"}
LEGACY_IMAGES_DIR = Path(PDF_DIR).parent / "images"


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
            "format": BACKUP_FORMAT,
            "schema_version": SCHEMA_VERSION,
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
                (Path(LEGACY_IMAGES_DIR), "images"),
            ):
                if not source_dir.exists():
                    continue
                for item in source_dir.rglob("*"):
                    if item.is_file() and not item.name.endswith(".part"):
                        archive.write(
                            item,
                            str(
                                PurePosixPath(
                                    archive_root,
                                    *item.relative_to(source_dir).parts,
                                )
                            ),
                        )
        temp_zip.replace(target)
    logger.info("Backup created at %s", target)
    return str(target)


def _safe_members(archive: zipfile.ZipFile) -> list[zipfile.ZipInfo]:
    members = archive.infolist()
    for member in members:
        path = PurePosixPath(member.filename.replace("\\", "/"))
        if (
            not member.filename
            or "\x00" in member.filename
            or path.is_absolute()
            or ".." in path.parts
            or (path.parts and path.parts[0].endswith(":"))
        ):
            raise ValueError("Backup contains an unsafe path.")
    return members


def _normalized_member_path(member: zipfile.ZipInfo) -> PurePosixPath:
    return PurePosixPath(member.filename.replace("\\", "/"))


def _find_database_member(members: list[zipfile.ZipInfo]) -> zipfile.ZipInfo:
    candidates = [
        member
        for member in members
        if not member.is_dir()
        and _normalized_member_path(member).name.casefold() == "vouchers.db"
    ]
    root_matches = [
        member
        for member in candidates
        if _normalized_member_path(member).parts == ("vouchers.db",)
    ]
    if len(root_matches) == 1:
        return root_matches[0]
    if not candidates:
        raise ValueError("Backup does not contain vouchers.db.")

    minimum_depth = min(len(_normalized_member_path(item).parts) for item in candidates)
    shallowest = [
        item
        for item in candidates
        if len(_normalized_member_path(item).parts) == minimum_depth
    ]
    if len(shallowest) != 1:
        raise ValueError("Backup contains more than one possible vouchers.db file.")
    return shallowest[0]


def _validate_metadata(
    archive: zipfile.ZipFile,
    members: list[zipfile.ZipInfo],
    database_member: zipfile.ZipInfo,
) -> None:
    database_parent = _normalized_member_path(database_member).parent
    expected = database_parent / "metadata.json"
    metadata_members = [
        member
        for member in members
        if not member.is_dir() and _normalized_member_path(member) == expected
    ]
    if not metadata_members:
        return  # Legacy backups did not contain metadata.
    try:
        metadata = json.loads(archive.read(metadata_members[0]).decode("utf-8-sig"))
        backup_format = int(metadata.get("format", 1))
    except (OSError, UnicodeError, ValueError, TypeError, json.JSONDecodeError) as exc:
        raise ValueError("Backup metadata is not valid.") from exc
    if not 1 <= backup_format <= BACKUP_FORMAT:
        raise ValueError(
            f"Backup format {backup_format} is not supported by this application."
        )


def _extract_member(
    archive: zipfile.ZipFile,
    member: zipfile.ZipInfo,
    destination: Path,
) -> Path:
    member_path = _normalized_member_path(member)
    target = destination.joinpath(*member_path.parts)
    if member.is_dir():
        target.mkdir(parents=True, exist_ok=True)
        return target
    target.parent.mkdir(parents=True, exist_ok=True)
    with archive.open(member) as source, target.open("wb") as output:
        shutil.copyfileobj(source, output)
    return target


def _extract_all_safe(
    archive: zipfile.ZipFile,
    members: list[zipfile.ZipInfo],
    destination: Path,
) -> None:
    for member in members:
        _extract_member(archive, member, destination)


def _check_database(db_path: Path) -> None:
    try:
        with closing(sqlite3.connect(db_path)) as conn:
            integrity = conn.execute("PRAGMA integrity_check").fetchone()[0]
            if integrity != "ok":
                raise ValueError(f"Backup database failed integrity check: {integrity}")
            tables = {
                row[0]
                for row in conn.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                ).fetchall()
            }
            voucher_columns = {
                row[1] for row in conn.execute("PRAGMA table_info(vouchers)").fetchall()
            }
    except sqlite3.DatabaseError as exc:
        raise ValueError("Backup does not contain a valid SQLite database.") from exc
    missing = LEGACY_REQUIRED_TABLES - tables
    if missing:
        raise ValueError(f"Backup database is missing: {', '.join(sorted(missing))}.")
    if "voucher_id" not in voucher_columns:
        raise ValueError("Backup vouchers table does not contain voucher_id.")


def validate_backup(backup_path: str | Path) -> None:
    """Validate archive paths and the embedded SQLite database."""

    path = Path(backup_path)
    if not path.is_file():
        raise ValueError("Backup file was not found.")
    try:
        with zipfile.ZipFile(path) as archive:
            members = _safe_members(archive)
            database_member = _find_database_member(members)
            _validate_metadata(archive, members, database_member)
            with tempfile.TemporaryDirectory(prefix="svms-validate-") as temp_name:
                database_path = _extract_member(
                    archive, database_member, Path(temp_name)
                )
                _check_database(database_path)
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
            members = _safe_members(archive)
            database_member = _find_database_member(members)
            _extract_all_safe(archive, members, temp_dir)

        restored_db = temp_dir.joinpath(*_normalized_member_path(database_member).parts)
        _check_database(restored_db)
        init_db(restored_db)
        with closing(sqlite3.connect(restored_db)) as source, db_session() as destination:
            source.backup(destination)

        backup_root = restored_db.parent
        for archive_root, destination_root in (
            ("pdfs", Path(PDF_DIR)),
            ("staffs", Path(STAFFS_ROOT)),
            ("images", Path(LEGACY_IMAGES_DIR)),
        ):
            source_dir = backup_root / archive_root
            if source_dir.exists():
                destination_root.mkdir(parents=True, exist_ok=True)
                shutil.copytree(source_dir, destination_root, dirs_exist_ok=True)

    logger.warning("Backup restored from %s; pre-restore backup: %s", path, pre_restore)
    return pre_restore
