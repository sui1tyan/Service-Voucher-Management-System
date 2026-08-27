"""SQLite persistence and migrations for SVMS."""

from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Iterator, Mapping

from auth import (
    hash_pwd,
    validate_password_policy,
    validate_username,
    verify_pwd,
)
from config import DB_FILE, DEFAULT_BASE_VID, ROLES, logger
from validators import validate_commission, validate_staff, validate_voucher


SCHEMA_VERSION = 2

VOUCHER_FIELDS = (
    "customer_name",
    "contact_number",
    "units",
    "particulars",
    "problem",
    "staff_name",
    "status",
    "recipient",
    "solution",
    "technician_id",
    "technician_name",
    "ref_bill",
    "ref_bill_date",
    "amount_rm",
    "tech_commission",
)


def now_text() -> str:
    return datetime.now().isoformat(sep=" ", timespec="seconds")


def get_conn(db_file: str | Path | None = None) -> sqlite3.Connection:
    """Open a short-lived SQLite connection with safe concurrency defaults."""

    target = Path(db_file or DB_FILE)
    target.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(target, timeout=15)
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous = NORMAL")
    conn.execute("PRAGMA busy_timeout = 15000")
    conn.row_factory = sqlite3.Row
    return conn


@contextmanager
def db_session(
    db_file: str | Path | None = None,
) -> Iterator[sqlite3.Connection]:
    """Yield a connection, commit or roll back, and always close it."""

    conn = get_conn(db_file)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _column_names(conn: sqlite3.Connection, table: str) -> set[str]:
    return {row["name"] for row in conn.execute(f"PRAGMA table_info({table})")}


def _ensure_column(
    conn: sqlite3.Connection, table: str, column: str, definition: str
) -> None:
    if column not in _column_names(conn, table):
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def init_db() -> None:
    """Create a new database or safely migrate an older SVMS database."""

    try:
        with db_session() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS vouchers (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    voucher_id TEXT NOT NULL UNIQUE,
                    created_at TEXT NOT NULL,
                    customer_name TEXT NOT NULL,
                    contact_number TEXT NOT NULL,
                    units INTEGER NOT NULL DEFAULT 1,
                    particulars TEXT NOT NULL,
                    problem TEXT NOT NULL,
                    staff_name TEXT,
                    status TEXT NOT NULL DEFAULT 'Pending',
                    recipient TEXT,
                    solution TEXT,
                    pdf_path TEXT,
                    technician_id TEXT,
                    technician_name TEXT,
                    ref_bill TEXT,
                    ref_bill_date TEXT,
                    amount_rm REAL NOT NULL DEFAULT 0,
                    tech_commission REAL NOT NULL DEFAULT 0,
                    updated_at TEXT,
                    created_by TEXT,
                    updated_by TEXT
                );

                CREATE TABLE IF NOT EXISTS staffs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    position TEXT,
                    staff_id_opt TEXT,
                    name TEXT NOT NULL UNIQUE,
                    phone TEXT,
                    photo_path TEXT,
                    email TEXT,
                    note TEXT,
                    created_at TEXT,
                    updated_at TEXT
                );

                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT NOT NULL UNIQUE,
                    role TEXT NOT NULL CHECK(role IN ('admin','sales assistant','user','technician')),
                    password_hash BLOB NOT NULL,
                    is_active INTEGER NOT NULL DEFAULT 1,
                    must_change_pwd INTEGER NOT NULL DEFAULT 0,
                    full_name TEXT,
                    phone TEXT,
                    email TEXT,
                    note TEXT,
                    last_login TEXT,
                    created_at TEXT,
                    updated_at TEXT
                );

                CREATE TABLE IF NOT EXISTS commissions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    staff_id INTEGER NOT NULL REFERENCES staffs(id) ON DELETE RESTRICT,
                    bill_type TEXT NOT NULL CHECK(bill_type IN ('CS','INV')),
                    bill_no TEXT NOT NULL,
                    total_amount REAL NOT NULL DEFAULT 0,
                    commission_amount REAL NOT NULL DEFAULT 0,
                    bill_image_path TEXT,
                    created_at TEXT,
                    updated_at TEXT,
                    voucher_id TEXT,
                    note TEXT
                );

                CREATE TABLE IF NOT EXISTS audit_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    occurred_at TEXT NOT NULL,
                    username TEXT,
                    action TEXT NOT NULL,
                    entity_type TEXT,
                    entity_id TEXT,
                    details TEXT
                );
                """
            )

            migrations = {
                "vouchers": {
                    "updated_at": "TEXT",
                    "created_by": "TEXT",
                    "updated_by": "TEXT",
                },
                "staffs": {"email": "TEXT", "note": "TEXT"},
                "users": {
                    "full_name": "TEXT",
                    "phone": "TEXT",
                    "email": "TEXT",
                    "note": "TEXT",
                    "last_login": "TEXT",
                },
                "commissions": {"voucher_id": "TEXT", "note": "TEXT"},
            }
            for table, columns in migrations.items():
                for column, definition in columns.items():
                    _ensure_column(conn, table, column, definition)

            conn.executescript(
                """
                CREATE INDEX IF NOT EXISTS idx_vouchers_created_at
                    ON vouchers(created_at);
                CREATE INDEX IF NOT EXISTS idx_vouchers_customer
                    ON vouchers(customer_name COLLATE NOCASE);
                CREATE INDEX IF NOT EXISTS idx_vouchers_status
                    ON vouchers(status);
                CREATE INDEX IF NOT EXISTS idx_vouchers_recipient
                    ON vouchers(recipient COLLATE NOCASE);
                CREATE INDEX IF NOT EXISTS idx_commissions_staff
                    ON commissions(staff_id);
                CREATE INDEX IF NOT EXISTS idx_commissions_voucher
                    ON commissions(voucher_id);
                CREATE INDEX IF NOT EXISTS idx_audit_occurred_at
                    ON audit_log(occurred_at);
                """
            )
            conn.execute(
                "INSERT OR IGNORE INTO settings(key, value) VALUES('base_vid', ?)",
                (str(DEFAULT_BASE_VID),),
            )
            conn.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")
        logger.info("Database initialized (schema version %s).", SCHEMA_VERSION)
    except Exception:
        logger.exception("Database initialization failed")
        raise


def _dict(row: sqlite3.Row | None) -> dict[str, Any] | None:
    return dict(row) if row is not None else None


def _audit(
    conn: sqlite3.Connection,
    username: str,
    action: str,
    entity_type: str = "",
    entity_id: str = "",
    details: Mapping[str, Any] | str | None = None,
) -> None:
    if isinstance(details, Mapping):
        detail_text = json.dumps(details, ensure_ascii=False, sort_keys=True)
    else:
        detail_text = str(details or "")
    conn.execute(
        """
        INSERT INTO audit_log
            (occurred_at, username, action, entity_type, entity_id, details)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (now_text(), username, action, entity_type, str(entity_id), detail_text),
    )


# ---------------------------------------------------------------------------
# Users and authentication
# ---------------------------------------------------------------------------


def has_admin_user() -> bool:
    with db_session() as conn:
        row = conn.execute(
            "SELECT 1 FROM users WHERE role='admin' AND is_active=1 LIMIT 1"
        ).fetchone()
    return row is not None


def create_initial_admin(
    username: str, password: str, full_name: str = ""
) -> dict[str, Any]:
    username = username.strip()
    if error := validate_username(username):
        raise ValueError(error)
    if error := validate_password_policy(password):
        raise ValueError(error)

    with db_session() as conn:
        conn.execute("BEGIN IMMEDIATE")
        if conn.execute(
            "SELECT 1 FROM users WHERE role='admin' AND is_active=1 LIMIT 1"
        ).fetchone():
            raise ValueError("An administrator account already exists.")
        if conn.execute(
            "SELECT 1 FROM users WHERE LOWER(username)=LOWER(?)", (username,)
        ).fetchone():
            raise ValueError("That username already exists.")
        timestamp = now_text()
        cursor = conn.execute(
            """
            INSERT INTO users
                (username, role, password_hash, is_active, must_change_pwd,
                 full_name, created_at, updated_at)
            VALUES (?, 'admin', ?, 1, 0, ?, ?, ?)
            """,
            (username, hash_pwd(password), full_name.strip(), timestamp, timestamp),
        )
        _audit(conn, username, "initial_admin_created", "user", str(cursor.lastrowid))
    return get_user(cursor.lastrowid)  # type: ignore[return-value]


def authenticate_user(username: str, password: str) -> dict[str, Any] | None:
    with db_session() as conn:
        row = conn.execute(
            "SELECT * FROM users WHERE username=? COLLATE NOCASE AND is_active=1",
            (username.strip(),),
        ).fetchone()
        if row is None or not verify_pwd(password, row["password_hash"]):
            return None
        login_time = now_text()
        conn.execute("UPDATE users SET last_login=? WHERE id=?", (login_time, row["id"]))
        _audit(conn, row["username"], "login", "user", str(row["id"]))
        result = dict(row)
        result["last_login"] = login_time
        result.pop("password_hash", None)
        return result


def get_user(user_id: int) -> dict[str, Any] | None:
    with db_session() as conn:
        row = conn.execute(
            """
            SELECT id, username, role, is_active, must_change_pwd, full_name,
                   phone, email, note, last_login, created_at, updated_at
            FROM users WHERE id=?
            """,
            (user_id,),
        ).fetchone()
    return _dict(row)


def list_users() -> list[dict[str, Any]]:
    with db_session() as conn:
        rows = conn.execute(
            """
            SELECT id, username, role, is_active, must_change_pwd, full_name,
                   phone, email, note, last_login, created_at, updated_at
            FROM users ORDER BY username COLLATE NOCASE
            """
        ).fetchall()
    return [dict(row) for row in rows]


def create_user(
    data: Mapping[str, Any], password: str, actor: str
) -> dict[str, Any]:
    username = str(data.get("username") or "").strip()
    role = str(data.get("role") or "").strip().lower()
    if error := validate_username(username):
        raise ValueError(error)
    if error := validate_password_policy(password):
        raise ValueError(error)
    if role not in ROLES:
        raise ValueError("Select a valid role.")

    timestamp = now_text()
    try:
        with db_session() as conn:
            if conn.execute(
                "SELECT 1 FROM users WHERE LOWER(username)=LOWER(?)", (username,)
            ).fetchone():
                raise ValueError("That username already exists.")
            cursor = conn.execute(
                """
                INSERT INTO users
                    (username, role, password_hash, is_active, must_change_pwd,
                     full_name, phone, email, note, created_at, updated_at)
                VALUES (?, ?, ?, 1, 1, ?, ?, ?, ?, ?, ?)
                """,
                (
                    username,
                    role,
                    hash_pwd(password),
                    str(data.get("full_name") or "").strip(),
                    str(data.get("phone") or "").strip(),
                    str(data.get("email") or "").strip(),
                    str(data.get("note") or "").strip(),
                    timestamp,
                    timestamp,
                ),
            )
            _audit(conn, actor, "user_created", "user", str(cursor.lastrowid), {"role": role})
    except sqlite3.IntegrityError as exc:
        raise ValueError("That username already exists.") from exc
    return get_user(cursor.lastrowid)  # type: ignore[return-value]


def update_user(user_id: int, data: Mapping[str, Any], actor: str) -> dict[str, Any]:
    username = str(data.get("username") or "").strip()
    role = str(data.get("role") or "").strip().lower()
    if error := validate_username(username):
        raise ValueError(error)
    if role not in ROLES:
        raise ValueError("Select a valid role.")
    is_active = 1 if bool(data.get("is_active", True)) else 0

    with db_session() as conn:
        current = conn.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
        if current is None:
            raise ValueError("User not found.")
        duplicate = conn.execute(
            "SELECT 1 FROM users WHERE LOWER(username)=LOWER(?) AND id<>?",
            (username, user_id),
        ).fetchone()
        if duplicate:
            raise ValueError("That username already exists.")
        if current["role"] == "admin" and (role != "admin" or not is_active):
            other_admins = conn.execute(
                """
                SELECT COUNT(*) FROM users
                WHERE role='admin' AND is_active=1 AND id<>?
                """,
                (user_id,),
            ).fetchone()[0]
            if other_admins == 0:
                raise ValueError("At least one active administrator is required.")
        try:
            conn.execute(
                """
                UPDATE users
                SET username=?, role=?, is_active=?, full_name=?, phone=?,
                    email=?, note=?, updated_at=?
                WHERE id=?
                """,
                (
                    username,
                    role,
                    is_active,
                    str(data.get("full_name") or "").strip(),
                    str(data.get("phone") or "").strip(),
                    str(data.get("email") or "").strip(),
                    str(data.get("note") or "").strip(),
                    now_text(),
                    user_id,
                ),
            )
        except sqlite3.IntegrityError as exc:
            raise ValueError("That username already exists.") from exc
        _audit(conn, actor, "user_updated", "user", str(user_id), {"role": role})
    return get_user(user_id)  # type: ignore[return-value]


def reset_user_password(
    user_id: int, new_password: str, actor: str, *, force_change: bool = True
) -> None:
    if error := validate_password_policy(new_password):
        raise ValueError(error)
    with db_session() as conn:
        if not conn.execute("SELECT 1 FROM users WHERE id=?", (user_id,)).fetchone():
            raise ValueError("User not found.")
        conn.execute(
            """
            UPDATE users
            SET password_hash=?, must_change_pwd=?, updated_at=?
            WHERE id=?
            """,
            (hash_pwd(new_password), int(force_change), now_text(), user_id),
        )
        _audit(conn, actor, "password_reset", "user", str(user_id))


def change_own_password(
    user_id: int, current_password: str, new_password: str, actor: str
) -> None:
    if error := validate_password_policy(new_password):
        raise ValueError(error)
    with db_session() as conn:
        row = conn.execute(
            "SELECT password_hash FROM users WHERE id=?", (user_id,)
        ).fetchone()
        if row is None or not verify_pwd(current_password, row["password_hash"]):
            raise ValueError("Current password is incorrect.")
        conn.execute(
            """
            UPDATE users
            SET password_hash=?, must_change_pwd=0, updated_at=?
            WHERE id=?
            """,
            (hash_pwd(new_password), now_text(), user_id),
        )
        _audit(conn, actor, "password_changed", "user", str(user_id))


def force_change_password(user_id: int, new_password: str, actor: str) -> None:
    if error := validate_password_policy(new_password):
        raise ValueError(error)
    with db_session() as conn:
        conn.execute(
            """
            UPDATE users
            SET password_hash=?, must_change_pwd=0, updated_at=?
            WHERE id=?
            """,
            (hash_pwd(new_password), now_text(), user_id),
        )
        _audit(conn, actor, "forced_password_changed", "user", str(user_id))


def delete_user(user_id: int, actor_user_id: int, actor: str) -> None:
    if user_id == actor_user_id:
        raise ValueError("You cannot delete your own account.")
    with db_session() as conn:
        row = conn.execute("SELECT role FROM users WHERE id=?", (user_id,)).fetchone()
        if row is None:
            raise ValueError("User not found.")
        if row["role"] == "admin":
            count = conn.execute(
                "SELECT COUNT(*) FROM users WHERE role='admin' AND is_active=1"
            ).fetchone()[0]
            if count <= 1:
                raise ValueError("The last active administrator cannot be deleted.")
        conn.execute("DELETE FROM users WHERE id=?", (user_id,))
        _audit(conn, actor, "user_deleted", "user", str(user_id))


# ---------------------------------------------------------------------------
# Vouchers
# ---------------------------------------------------------------------------


def get_setting(key: str, default: str = "") -> str:
    with db_session() as conn:
        row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return str(row["value"]) if row else default


def set_base_voucher_id(value: int, actor: str) -> None:
    if not 1 <= int(value) <= 999_999_999:
        raise ValueError("Base voucher ID must be between 1 and 999,999,999.")
    with db_session() as conn:
        conn.execute(
            """
            INSERT INTO settings(key, value) VALUES('base_vid', ?)
            ON CONFLICT(key) DO UPDATE SET value=excluded.value
            """,
            (str(int(value)),),
        )
        _audit(conn, actor, "base_voucher_id_updated", "settings", "base_vid", {"value": value})


def _next_voucher_id(conn: sqlite3.Connection) -> str:
    row = conn.execute(
        """
        SELECT MAX(CAST(voucher_id AS INTEGER))
        FROM vouchers
        WHERE voucher_id <> '' AND voucher_id NOT GLOB '*[^0-9]*'
        """
    ).fetchone()
    base = int(
        conn.execute(
            "SELECT value FROM settings WHERE key='base_vid'"
        ).fetchone()["value"]
    )
    return str(max(base, (int(row[0]) + 1) if row and row[0] is not None else base))


def get_next_voucher_id() -> str:
    with db_session() as conn:
        return _next_voucher_id(conn)


def create_voucher(data: Mapping[str, Any], actor: str) -> dict[str, Any]:
    values = validate_voucher(data)
    timestamp = now_text()
    with db_session() as conn:
        conn.execute("BEGIN IMMEDIATE")
        voucher_id = _next_voucher_id(conn)
        columns = ", ".join(VOUCHER_FIELDS)
        placeholders = ", ".join("?" for _ in VOUCHER_FIELDS)
        conn.execute(
            f"""
            INSERT INTO vouchers
                (voucher_id, created_at, {columns}, updated_at, created_by, updated_by)
            VALUES (?, ?, {placeholders}, ?, ?, ?)
            """,
            (
                voucher_id,
                timestamp,
                *(values[field] for field in VOUCHER_FIELDS),
                timestamp,
                actor,
                actor,
            ),
        )
        _audit(conn, actor, "voucher_created", "voucher", voucher_id)
    return get_voucher(voucher_id)  # type: ignore[return-value]


def get_voucher(voucher_id: str) -> dict[str, Any] | None:
    with db_session() as conn:
        row = conn.execute(
            "SELECT * FROM vouchers WHERE voucher_id=?", (str(voucher_id),)
        ).fetchone()
    return _dict(row)


def search_vouchers(
    filters: Mapping[str, Any] | None = None,
    *,
    limit: int = 100,
    offset: int = 0,
) -> list[dict[str, Any]]:
    
    filters = filters or {}
    sql = "SELECT * FROM vouchers WHERE 1=1"
    params: list[Any] = []

    filter_map = (
        ("voucher_id", "voucher_id"),
        ("customer_name", "customer_name"),
        ("contact_number", "contact_number"),
        ("recipient", "recipient"),
        ("technician_name", "technician_name"),
        ("ref_bill", "ref_bill"),
    )
    for key, column in filter_map:
        value = str(filters.get(key) or "").strip()
        if value:
            sql += f" AND LOWER(COALESCE({column}, '')) LIKE ?"
            params.append(f"%{value.lower()}%")

    status = str(filters.get("status") or "").strip()
    if status and status != "All":
        sql += " AND status=?"
        params.append(status)

    date_from = str(filters.get("date_from") or "").strip()
    date_to = str(filters.get("date_to") or "").strip()
    if date_from:
        sql += " AND DATE(created_at) >= DATE(?)"
        params.append(date_from)
    if date_to:
        sql += " AND DATE(created_at) <= DATE(?)"
        params.append(date_to)

    sql += " ORDER BY created_at DESC, id DESC LIMIT ? OFFSET ?"
    params.extend((limit, offset))
    with db_session() as conn:
        rows = conn.execute(sql, params).fetchall()
    return [dict(row) for row in rows]

def count_vouchers(
    filters: Mapping[str, Any] | None = None
) -> int:
    filters = filters or {}
    sql = "SELECT COUNT(*) FROM vouchers WHERE 1=1"
    params: list[Any] = []

    voucher_id = str(filters.get("voucher_id") or "").strip()
    if voucher_id:
        sql += " AND voucher_id LIKE ?"
        params.append(f"%{voucher_id}%")

    customer_name = str(filters.get("customer_name") or "").strip()
    if customer_name:
        sql += " AND customer_name LIKE ? COLLATE NOCASE"
        params.append(f"%{customer_name}%")

    contact_number = str(filters.get("contact_number") or "").strip()
    if contact_number:
        sql += " AND contact_number LIKE ?"
        params.append(f"%{contact_number}%")

    recipient = str(filters.get("recipient") or "").strip()
    if recipient:
        sql += " AND recipient LIKE ? COLLATE NOCASE"
        params.append(f"%{recipient}%")

    status = str(filters.get("status") or "").strip()
    if status and status != "All":
        sql += " AND status = ?"
        params.append(status)

    date_from = str(filters.get("date_from") or "").strip()
    if date_from:
        sql += " AND DATE(created_at) >= DATE(?)"
        params.append(date_from)

    date_to = str(filters.get("date_to") or "").strip()
    if date_to:
        sql += " AND DATE(created_at) <= DATE(?)"
        params.append(date_to)

    with db_session() as conn:
        row = conn.execute(sql, params).fetchone()

    return int(row[0]) if row else 0

def update_voucher(
    voucher_id: str, data: Mapping[str, Any], actor: str
) -> dict[str, Any]:
    values = validate_voucher(data)
    assignments = ", ".join(f"{field}=?" for field in VOUCHER_FIELDS)
    with db_session() as conn:
        cursor = conn.execute(
            f"""
            UPDATE vouchers SET {assignments}, updated_at=?, updated_by=?
            WHERE voucher_id=?
            """,
            (
                *(values[field] for field in VOUCHER_FIELDS),
                now_text(),
                actor,
                str(voucher_id),
            ),
        )
        if cursor.rowcount != 1:
            raise ValueError("Voucher not found.")
        _audit(conn, actor, "voucher_updated", "voucher", str(voucher_id))
    return get_voucher(str(voucher_id))  # type: ignore[return-value]


def update_voucher_pdf_path(voucher_id: str, pdf_path: str, actor: str) -> None:
    with db_session() as conn:
        cursor = conn.execute(
            """
            UPDATE vouchers SET pdf_path=?, updated_at=?, updated_by=?
            WHERE voucher_id=?
            """,
            (str(pdf_path), now_text(), actor, str(voucher_id)),
        )
        if cursor.rowcount != 1:
            raise ValueError("Voucher not found.")


def delete_voucher(voucher_id: str, actor: str) -> str:
    with db_session() as conn:
        row = conn.execute(
            "SELECT pdf_path FROM vouchers WHERE voucher_id=?", (str(voucher_id),)
        ).fetchone()
        if row is None:
            raise ValueError("Voucher not found.")
        conn.execute("DELETE FROM vouchers WHERE voucher_id=?", (str(voucher_id),))
        _audit(conn, actor, "voucher_deleted", "voucher", str(voucher_id))
    return str(row["pdf_path"] or "")


# ---------------------------------------------------------------------------
# Staff
# ---------------------------------------------------------------------------


def list_staffs() -> list[dict[str, Any]]:
    with db_session() as conn:
        rows = conn.execute(
            "SELECT * FROM staffs ORDER BY name COLLATE NOCASE"
        ).fetchall()
    return [dict(row) for row in rows]


def list_staffs_names() -> list[str]:
    return [str(row["name"]) for row in list_staffs()]


def get_staff(staff_id: int) -> dict[str, Any] | None:
    with db_session() as conn:
        row = conn.execute("SELECT * FROM staffs WHERE id=?", (staff_id,)).fetchone()
    return _dict(row)


def create_staff(data: Mapping[str, Any], actor: str) -> dict[str, Any]:
    values = validate_staff(data)
    timestamp = now_text()
    try:
        with db_session() as conn:
            cursor = conn.execute(
                """
                INSERT INTO staffs
                    (position, staff_id_opt, name, phone, photo_path, email,
                     note, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    values["position"],
                    values["staff_id_opt"],
                    values["name"],
                    values["phone"],
                    values["photo_path"],
                    values["email"],
                    values["note"],
                    timestamp,
                    timestamp,
                ),
            )
            _audit(conn, actor, "staff_created", "staff", str(cursor.lastrowid))
    except sqlite3.IntegrityError as exc:
        raise ValueError("A staff member with that name already exists.") from exc
    return get_staff(cursor.lastrowid)  # type: ignore[return-value]


def update_staff(
    staff_id: int, data: Mapping[str, Any], actor: str
) -> dict[str, Any]:
    values = validate_staff(data)
    try:
        with db_session() as conn:
            cursor = conn.execute(
                """
                UPDATE staffs
                SET position=?, staff_id_opt=?, name=?, phone=?, photo_path=?,
                    email=?, note=?, updated_at=?
                WHERE id=?
                """,
                (
                    values["position"],
                    values["staff_id_opt"],
                    values["name"],
                    values["phone"],
                    values["photo_path"],
                    values["email"],
                    values["note"],
                    now_text(),
                    staff_id,
                ),
            )
            if cursor.rowcount != 1:
                raise ValueError("Staff member not found.")
            _audit(conn, actor, "staff_updated", "staff", str(staff_id))
    except sqlite3.IntegrityError as exc:
        raise ValueError("A staff member with that name already exists.") from exc
    return get_staff(staff_id)  # type: ignore[return-value]


def delete_staff(staff_id: int, actor: str) -> None:
    try:
        with db_session() as conn:
            commission_count = conn.execute(
                "SELECT COUNT(*) FROM commissions WHERE staff_id=?", (staff_id,)
            ).fetchone()[0]
            if commission_count:
                raise ValueError(
                    "This staff member has commission records and cannot be deleted."
                )
            cursor = conn.execute("DELETE FROM staffs WHERE id=?", (staff_id,))
            if cursor.rowcount != 1:
                raise ValueError("Staff member not found.")
            _audit(conn, actor, "staff_deleted", "staff", str(staff_id))
    except sqlite3.IntegrityError as exc:
        raise ValueError(
            "This staff member has commission records and cannot be deleted."
        ) from exc


# ---------------------------------------------------------------------------
# Commissions
# ---------------------------------------------------------------------------


def list_commissions() -> list[dict[str, Any]]:
    with db_session() as conn:
        rows = conn.execute(
            """
            SELECT c.*, COALESCE(s.name, 'Unassigned') AS staff_name
            FROM commissions c
            LEFT JOIN staffs s ON s.id=c.staff_id
            ORDER BY datetime(c.created_at) DESC, c.id DESC
            """
        ).fetchall()
    return [dict(row) for row in rows]


def get_commission(commission_id: int) -> dict[str, Any] | None:
    with db_session() as conn:
        row = conn.execute(
            """
            SELECT c.*, COALESCE(s.name, 'Unassigned') AS staff_name
            FROM commissions c LEFT JOIN staffs s ON s.id=c.staff_id
            WHERE c.id=?
            """,
            (commission_id,),
        ).fetchone()
    return _dict(row)


def create_commission(
    data: Mapping[str, Any], actor: str
) -> dict[str, Any]:
    values = validate_commission(data)
    timestamp = now_text()
    try:
        with db_session() as conn:
            cursor = conn.execute(
                """
                INSERT INTO commissions
                    (staff_id, bill_type, bill_no, total_amount,
                     commission_amount, bill_image_path, created_at,
                     updated_at, voucher_id, note)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    values["staff_id"],
                    values["bill_type"],
                    values["bill_no"],
                    values["total_amount"],
                    values["commission_amount"],
                    values["bill_image_path"],
                    timestamp,
                    timestamp,
                    values["voucher_id"],
                    values["note"],
                ),
            )
            _audit(
                conn,
                actor,
                "commission_created",
                "commission",
                str(cursor.lastrowid),
            )
    except sqlite3.IntegrityError as exc:
        raise ValueError("Select an existing staff member.") from exc
    return get_commission(cursor.lastrowid)  # type: ignore[return-value]


def update_commission(
    commission_id: int, data: Mapping[str, Any], actor: str
) -> dict[str, Any]:
    values = validate_commission(data)
    with db_session() as conn:
        cursor = conn.execute(
            """
            UPDATE commissions
            SET staff_id=?, bill_type=?, bill_no=?, total_amount=?,
                commission_amount=?, bill_image_path=?, voucher_id=?, note=?,
                updated_at=?
            WHERE id=?
            """,
            (
                values["staff_id"],
                values["bill_type"],
                values["bill_no"],
                values["total_amount"],
                values["commission_amount"],
                values["bill_image_path"],
                values["voucher_id"],
                values["note"],
                now_text(),
                commission_id,
            ),
        )
        if cursor.rowcount != 1:
            raise ValueError("Commission record not found.")
        _audit(conn, actor, "commission_updated", "commission", str(commission_id))
    return get_commission(commission_id)  # type: ignore[return-value]


def delete_commission(commission_id: int, actor: str) -> None:
    with db_session() as conn:
        cursor = conn.execute("DELETE FROM commissions WHERE id=?", (commission_id,))
        if cursor.rowcount != 1:
            raise ValueError("Commission record not found.")
        _audit(conn, actor, "commission_deleted", "commission", str(commission_id))


def list_audit_entries(limit: int = 500) -> list[dict[str, Any]]:
    with db_session() as conn:
        rows = conn.execute(
            "SELECT * FROM audit_log ORDER BY id DESC LIMIT ?",
            (max(1, min(int(limit), 5_000)),),
        ).fetchall()
    return [dict(row) for row in rows]
