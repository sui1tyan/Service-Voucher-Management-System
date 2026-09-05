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
from validators import (
    MAX_VOUCHER_ID,
    normalize_voucher_id,
    validate_commission,
    validate_staff,
    validate_voucher,
)

SCHEMA_VERSION = 3

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


def init_db(db_file: str | Path | None = None) -> None:
    """Create a new database or safely migrate an older SVMS database."""

    try:
        with db_session(db_file) as conn:
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
                    note TEXT,
                    managed_by_voucher INTEGER NOT NULL DEFAULT 0
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
                    "created_at": "TEXT",
                    "customer_name": "TEXT",
                    "contact_number": "TEXT",
                    "units": "INTEGER NOT NULL DEFAULT 1",
                    "particulars": "TEXT",
                    "problem": "TEXT",
                    "staff_name": "TEXT",
                    "status": "TEXT NOT NULL DEFAULT 'Pending'",
                    "recipient": "TEXT",
                    "solution": "TEXT",
                    "pdf_path": "TEXT",
                    "technician_id": "TEXT",
                    "technician_name": "TEXT",
                    "ref_bill": "TEXT",
                    "ref_bill_date": "TEXT",
                    "amount_rm": "REAL NOT NULL DEFAULT 0",
                    "tech_commission": "REAL NOT NULL DEFAULT 0",
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
                "commissions": {
                    "voucher_id": "TEXT",
                    "note": "TEXT",
                    "managed_by_voucher": "INTEGER NOT NULL DEFAULT 0",
                },
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
                CREATE INDEX IF NOT EXISTS idx_commissions_created_at
                    ON commissions(created_at);
                CREATE INDEX IF NOT EXISTS idx_commissions_bill
                    ON commissions(bill_type, bill_no COLLATE NOCASE);
                CREATE INDEX IF NOT EXISTS idx_audit_occurred_at
                    ON audit_log(occurred_at);
                """
            )
            conn.execute(
                "INSERT OR IGNORE INTO settings(key, value) VALUES('base_vid', ?)",
                (str(DEFAULT_BASE_VID),),
            )
            _ensure_next_voucher_id_setting(conn)
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


def _integer_setting(conn: sqlite3.Connection, key: str, default: int) -> int:
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    try:
        return int(row["value"]) if row else int(default)
    except (TypeError, ValueError):
        return int(default)


def _maximum_numeric_voucher_id(conn: sqlite3.Connection) -> int | None:
    row = conn.execute(
        """
        SELECT MAX(CAST(voucher_id AS INTEGER))
        FROM vouchers
        WHERE voucher_id <> '' AND voucher_id NOT GLOB '*[^0-9]*'
        """
    ).fetchone()
    return int(row[0]) if row and row[0] is not None else None


def _ensure_next_voucher_id_setting(conn: sqlite3.Connection) -> int:
    """Migrate and maintain the monotonic next voucher ID setting."""

    base = _integer_setting(conn, "base_vid", DEFAULT_BASE_VID)
    if not 1 <= base <= MAX_VOUCHER_ID:
        base = DEFAULT_BASE_VID
        conn.execute(
            """
            INSERT INTO settings(key, value) VALUES('base_vid', ?)
            ON CONFLICT(key) DO UPDATE SET value=excluded.value
            """,
            (str(base),),
        )

    maximum = _maximum_numeric_voucher_id(conn)
    existing_next = _integer_setting(conn, "next_vid", base)
    candidate = max(base, existing_next, (maximum + 1) if maximum is not None else base)
    conn.execute(
        """
        INSERT INTO settings(key, value) VALUES('next_vid', ?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value
        """,
        (str(candidate),),
    )
    return candidate


def _advance_next_voucher_id_setting(
    conn: sqlite3.Connection, minimum_next: int
) -> None:
    current = _ensure_next_voucher_id_setting(conn)
    if int(minimum_next) > current:
        conn.execute(
            "UPDATE settings SET value=? WHERE key='next_vid'",
            (str(int(minimum_next)),),
        )


def get_setting(key: str, default: str = "") -> str:
    with db_session() as conn:
        row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return str(row["value"]) if row else default


def set_base_voucher_id(value: int, actor: str) -> None:
    if not 1 <= int(value) <= MAX_VOUCHER_ID:
        raise ValueError("Base voucher ID must be between 1 and 999,999,999.")
    with db_session() as conn:
        conn.execute(
            """
            INSERT INTO settings(key, value) VALUES('base_vid', ?)
            ON CONFLICT(key) DO UPDATE SET value=excluded.value
            """,
            (str(int(value)),),
        )
        _advance_next_voucher_id_setting(conn, int(value))
        _audit(conn, actor, "base_voucher_id_updated", "settings", "base_vid", {"value": value})


def _next_voucher_id(conn: sqlite3.Connection) -> str:
    candidate = _ensure_next_voucher_id_setting(conn)
    if candidate > MAX_VOUCHER_ID:
        raise ValueError("No more numeric voucher IDs are available.")
    return str(candidate)


def get_next_voucher_id() -> str:
    with db_session() as conn:
        return _next_voucher_id(conn)


def _resolve_voucher_staff_id(
    conn: sqlite3.Connection, voucher: Mapping[str, Any]
) -> int | None:
    voucher_data = dict(voucher)
    technician_id = str(voucher_data.get("technician_id") or "").strip()
    if technician_id:
        row = conn.execute(
            """
            SELECT id FROM staffs
            WHERE LOWER(TRIM(COALESCE(staff_id_opt, ''))) = LOWER(?)
            ORDER BY id LIMIT 1
            """,
            (technician_id,),
        ).fetchone()
        if row:
            return int(row["id"])

    technician_name = str(voucher_data.get("technician_name") or "").strip()
    if technician_name:
        row = conn.execute(
            """
            SELECT id FROM staffs
            WHERE LOWER(TRIM(name)) = LOWER(?)
            ORDER BY id LIMIT 1
            """,
            (technician_name,),
        ).fetchone()
        if row:
            return int(row["id"])
    return None


def _bill_type_from_reference(reference: str) -> str:
    return "INV" if reference.strip().upper().startswith("INV") else "CS"


def _sync_commission_from_voucher(
    conn: sqlite3.Connection, voucher_id: str, actor: str
) -> None:
    """Make the voucher the source of truth for its linked commission."""

    voucher = conn.execute(
        "SELECT * FROM vouchers WHERE voucher_id=?", (str(voucher_id),)
    ).fetchone()
    if voucher is None:
        return

    linked = conn.execute(
        """
        SELECT * FROM commissions
        WHERE voucher_id=?
        ORDER BY managed_by_voucher DESC, id
        """,
        (str(voucher_id),),
    ).fetchall()
    staff_id = _resolve_voucher_staff_id(conn, voucher)
    bill_no = str(voucher["ref_bill"] or "").strip()

    if staff_id is None or not bill_no:
        for commission in linked:
            commission_id = int(commission["id"])
            if int(commission["managed_by_voucher"] or 0):
                conn.execute("DELETE FROM commissions WHERE id=?", (commission_id,))
                _audit(
                    conn,
                    actor,
                    "commission_removed_after_voucher_sync",
                    "commission",
                    str(commission_id),
                    {"voucher_id": str(voucher_id)},
                )
        return

    bill_type = _bill_type_from_reference(bill_no)
    total_amount = float(voucher["amount_rm"] or 0)
    commission_amount = float(voucher["tech_commission"] or 0)
    timestamp = now_text()

    if not linked:
        matching = conn.execute(
            """
            SELECT * FROM commissions
            WHERE (voucher_id IS NULL OR TRIM(voucher_id)='')
              AND staff_id=? AND bill_type=?
              AND LOWER(TRIM(bill_no))=LOWER(?)
            ORDER BY id
            """,
            (staff_id, bill_type, bill_no),
        ).fetchall()
        if len(matching) == 1:
            linked = matching

    if linked:
        for commission in linked:
            commission_id = int(commission["id"])
            conn.execute(
                """
                UPDATE commissions
                SET staff_id=?, bill_type=?, bill_no=?, total_amount=?,
                    commission_amount=?, voucher_id=?, updated_at=?
                WHERE id=?
                """,
                (
                    staff_id,
                    bill_type,
                    bill_no,
                    total_amount,
                    commission_amount,
                    str(voucher_id),
                    timestamp,
                    commission_id,
                ),
            )
            _audit(
                conn,
                actor,
                "commission_synced_from_voucher",
                "commission",
                str(commission_id),
                {"voucher_id": str(voucher_id)},
            )
        return

    cursor = conn.execute(
        """
        INSERT INTO commissions
            (staff_id, bill_type, bill_no, total_amount, commission_amount,
             bill_image_path, created_at, updated_at, voucher_id, note,
             managed_by_voucher)
        VALUES (?, ?, ?, ?, ?, '', ?, ?, ?, '', 1)
        """,
        (
            staff_id,
            bill_type,
            bill_no,
            total_amount,
            commission_amount,
            timestamp,
            timestamp,
            str(voucher_id),
        ),
    )
    _audit(
        conn,
        actor,
        "commission_created_from_voucher",
        "commission",
        str(cursor.lastrowid),
        {"voucher_id": str(voucher_id)},
    )


def _detach_commissions_for_voucher(
    conn: sqlite3.Connection, voucher_id: str, actor: str
) -> None:
    linked = conn.execute(
        "SELECT id, managed_by_voucher FROM commissions WHERE voucher_id=?",
        (str(voucher_id),),
    ).fetchall()
    for commission in linked:
        commission_id = int(commission["id"])
        if int(commission["managed_by_voucher"] or 0):
            conn.execute("DELETE FROM commissions WHERE id=?", (commission_id,))
            action = "commission_removed_with_voucher"
        else:
            conn.execute(
                "UPDATE commissions SET voucher_id=NULL, updated_at=? WHERE id=?",
                (now_text(), commission_id),
            )
            action = "commission_unlinked_from_deleted_voucher"
        _audit(
            conn,
            actor,
            action,
            "commission",
            str(commission_id),
            {"voucher_id": str(voucher_id)},
        )


def create_voucher(data: Mapping[str, Any], actor: str) -> dict[str, Any]:
    values = validate_voucher(data)
    requested_id = normalize_voucher_id(data.get("voucher_id"), optional=True)
    timestamp = now_text()
    try:
        with db_session() as conn:
            conn.execute("BEGIN IMMEDIATE")
            voucher_id = requested_id or _next_voucher_id(conn)
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
            _advance_next_voucher_id_setting(conn, int(voucher_id) + 1)
            _audit(conn, actor, "voucher_created", "voucher", voucher_id)
            _sync_commission_from_voucher(conn, voucher_id, actor)
    except sqlite3.IntegrityError as exc:
        duplicate_id = requested_id or "selected"
        raise ValueError(f"Voucher ID {duplicate_id} already exists.") from exc
    return get_voucher(voucher_id)  # type: ignore[return-value]


def get_voucher(voucher_id: str) -> dict[str, Any] | None:
    with db_session() as conn:
        row = conn.execute(
            "SELECT * FROM vouchers WHERE voucher_id=?", (str(voucher_id),)
        ).fetchone()
    return _dict(row)


def _voucher_filter_clause(
    filters: Mapping[str, Any] | None = None,
) -> tuple[str, list[Any]]:
    """Build the shared WHERE clause used by list, count, and export queries."""

    filters = filters or {}
    clauses: list[str] = []
    params: list[Any] = []

    text_filters = (
        ("voucher_id", "voucher_id"),
        ("customer_name", "customer_name"),
        ("contact_number", "contact_number"),
        ("recipient", "recipient"),
        ("technician_name", "technician_name"),
        ("ref_bill", "ref_bill"),
    )
    for key, column in text_filters:
        value = str(filters.get(key) or "").strip()
        if value:
            clauses.append(f"LOWER(COALESCE({column}, '')) LIKE ?")
            params.append(f"%{value.lower()}%")

    status = str(filters.get("status") or "").strip()
    if status and status.casefold() != "all":
        clauses.append("status = ?")
        params.append(status)

    date_from = str(filters.get("date_from") or "").strip()
    if date_from:
        clauses.append("DATE(created_at) >= DATE(?)")
        params.append(date_from)

    date_to = str(filters.get("date_to") or "").strip()
    if date_to:
        clauses.append("DATE(created_at) <= DATE(?)")
        params.append(date_to)

    if not clauses:
        return "", params
    return f" WHERE {' AND '.join(clauses)}", params


def _voucher_order_clause(sort_direction: str) -> str:
    direction = str(sort_direction).lower()
    if direction not in {"asc", "desc"}:
        raise ValueError("Voucher sort direction must be 'asc' or 'desc'.")
    sql_direction = direction.upper()
    numeric = "voucher_id <> '' AND voucher_id NOT GLOB '*[^0-9]*'"
    return (
        f" ORDER BY CASE WHEN {numeric} THEN 0 ELSE 1 END ASC,"
        f" CASE WHEN {numeric} THEN CAST(voucher_id AS INTEGER) END {sql_direction},"
        f" voucher_id COLLATE NOCASE {sql_direction}, id {sql_direction}"
    )


def search_vouchers(
    filters: Mapping[str, Any] | None = None,
    *,
    limit: int = 100,
    offset: int = 0,
    sort_direction: str = "desc",
) -> list[dict[str, Any]]:
    page_limit = max(1, min(int(limit), 5_000))
    page_offset = max(0, int(offset))
    where_clause, params = _voucher_filter_clause(filters)
    sql = (
        f"SELECT * FROM vouchers{where_clause}"
        f"{_voucher_order_clause(sort_direction)} LIMIT ? OFFSET ?"
    )
    params.extend((page_limit, page_offset))
    with db_session() as conn:
        rows = conn.execute(sql, params).fetchall()
    return [dict(row) for row in rows]


def count_vouchers(filters: Mapping[str, Any] | None = None) -> int:
    where_clause, params = _voucher_filter_clause(filters)
    with db_session() as conn:
        row = conn.execute(
            f"SELECT COUNT(*) FROM vouchers{where_clause}", params
        ).fetchone()

    return int(row[0]) if row else 0


def iter_vouchers(
    filters: Mapping[str, Any] | None = None,
    *,
    batch_size: int = 1_000,
    sort_direction: str = "desc",
) -> Iterator[dict[str, Any]]:
    """Yield every matching voucher in stable order without page-size limits."""

    fetch_size = max(1, min(int(batch_size), 5_000))
    where_clause, params = _voucher_filter_clause(filters)
    sql = (
        f"SELECT * FROM vouchers{where_clause}"
        f"{_voucher_order_clause(sort_direction)}"
    )

    with db_session() as conn:
        cursor = conn.execute(sql, params)
        while rows := cursor.fetchmany(fetch_size):
            for row in rows:
                yield dict(row)


def update_voucher(
    voucher_id: str, data: Mapping[str, Any], actor: str
) -> dict[str, Any]:
    values = validate_voucher(data)
    old_id = str(voucher_id)
    supplied_id = str(data.get("voucher_id", old_id)).strip()
    new_id = old_id if supplied_id == old_id else normalize_voucher_id(supplied_id)
    assignments = [f"{field}=?" for field in VOUCHER_FIELDS]
    parameters: list[Any] = [values[field] for field in VOUCHER_FIELDS]
    if new_id != old_id:
        assignments.extend(("voucher_id=?", "pdf_path=''"))
        parameters.append(new_id)
    assignments.extend(("updated_at=?", "updated_by=?"))
    parameters.extend((now_text(), actor, old_id))

    try:
        with db_session() as conn:
            conn.execute("BEGIN IMMEDIATE")
            if conn.execute(
                "SELECT 1 FROM vouchers WHERE voucher_id=?", (old_id,)
            ).fetchone() is None:
                raise ValueError("Voucher not found.")
            cursor = conn.execute(
                f"UPDATE vouchers SET {', '.join(assignments)} WHERE voucher_id=?",
                parameters,
            )
            if cursor.rowcount != 1:
                raise ValueError("Voucher not found.")
            if new_id != old_id:
                conn.execute(
                    "UPDATE commissions SET voucher_id=?, updated_at=? WHERE voucher_id=?",
                    (new_id, now_text(), old_id),
                )
                _advance_next_voucher_id_setting(conn, int(new_id) + 1)
            _audit(
                conn,
                actor,
                "voucher_updated",
                "voucher",
                new_id,
                {"previous_voucher_id": old_id} if new_id != old_id else None,
            )
            _sync_commission_from_voucher(conn, new_id, actor)
    except sqlite3.IntegrityError as exc:
        raise ValueError(f"Voucher ID {new_id} already exists.") from exc
    return get_voucher(new_id)  # type: ignore[return-value]


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
        _detach_commissions_for_voucher(conn, str(voucher_id), actor)
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


def _commission_filter_clause(
    filters: Mapping[str, Any] | None = None,
) -> tuple[str, list[Any]]:
    filters = filters or {}
    clauses: list[str] = []
    params: list[Any] = []

    query = str(filters.get("query") or "").strip().lower()
    if query:
        wildcard = f"%{query}%"
        clauses.append(
            "(LOWER(COALESCE(s.name, '')) LIKE ? "
            "OR LOWER(COALESCE(c.bill_no, '')) LIKE ? "
            "OR LOWER(COALESCE(c.voucher_id, '')) LIKE ?)"
        )
        params.extend((wildcard, wildcard, wildcard))

    for key, expression in (
        ("staff_name", "s.name"),
        ("bill_no", "c.bill_no"),
        ("voucher_id", "c.voucher_id"),
    ):
        value = str(filters.get(key) or "").strip().lower()
        if value:
            clauses.append(f"LOWER(COALESCE({expression}, '')) LIKE ?")
            params.append(f"%{value}%")

    bill_type = str(filters.get("bill_type") or "").strip().upper()
    if bill_type and bill_type != "ALL":
        clauses.append("c.bill_type=?")
        params.append(bill_type)

    staff_id = filters.get("staff_id")
    if staff_id not in (None, ""):
        clauses.append("c.staff_id=?")
        params.append(int(staff_id))

    date_from = str(filters.get("date_from") or "").strip()
    if date_from:
        clauses.append("DATE(c.created_at) >= DATE(?)")
        params.append(date_from)

    date_to = str(filters.get("date_to") or "").strip()
    if date_to:
        clauses.append("DATE(c.created_at) <= DATE(?)")
        params.append(date_to)

    if not clauses:
        return "", params
    return f" WHERE {' AND '.join(clauses)}", params


def _commission_order_clause(sort_direction: str) -> str:
    direction = str(sort_direction).lower()
    if direction not in {"asc", "desc"}:
        raise ValueError("Commission sort direction must be 'asc' or 'desc'.")
    return f" ORDER BY c.id {direction.upper()}"


def search_commissions(
    filters: Mapping[str, Any] | None = None,
    *,
    limit: int = 100,
    offset: int = 0,
    sort_direction: str = "desc",
) -> list[dict[str, Any]]:
    page_limit = max(1, min(int(limit), 5_000))
    page_offset = max(0, int(offset))
    where_clause, params = _commission_filter_clause(filters)
    sql = (
        "SELECT c.*, COALESCE(s.name, 'Unassigned') AS staff_name "
        "FROM commissions c LEFT JOIN staffs s ON s.id=c.staff_id"
        f"{where_clause}{_commission_order_clause(sort_direction)} LIMIT ? OFFSET ?"
    )
    params.extend((page_limit, page_offset))
    with db_session() as conn:
        rows = conn.execute(sql, params).fetchall()
    return [dict(row) for row in rows]


def count_commissions(filters: Mapping[str, Any] | None = None) -> int:
    where_clause, params = _commission_filter_clause(filters)
    with db_session() as conn:
        row = conn.execute(
            "SELECT COUNT(*) FROM commissions c "
            f"LEFT JOIN staffs s ON s.id=c.staff_id{where_clause}",
            params,
        ).fetchone()
    return int(row[0]) if row else 0


def sum_commissions(filters: Mapping[str, Any] | None = None) -> float:
    where_clause, params = _commission_filter_clause(filters)
    with db_session() as conn:
        row = conn.execute(
            "SELECT COALESCE(SUM(c.commission_amount), 0) FROM commissions c "
            f"LEFT JOIN staffs s ON s.id=c.staff_id{where_clause}",
            params,
        ).fetchone()
    return float(row[0] or 0) if row else 0.0


def iter_commissions(
    filters: Mapping[str, Any] | None = None,
    *,
    batch_size: int = 1_000,
    sort_direction: str = "desc",
) -> Iterator[dict[str, Any]]:
    fetch_size = max(1, min(int(batch_size), 5_000))
    where_clause, params = _commission_filter_clause(filters)
    sql = (
        "SELECT c.*, COALESCE(s.name, 'Unassigned') AS staff_name "
        "FROM commissions c LEFT JOIN staffs s ON s.id=c.staff_id"
        f"{where_clause}{_commission_order_clause(sort_direction)}"
    )
    with db_session() as conn:
        cursor = conn.execute(sql, params)
        while rows := cursor.fetchmany(fetch_size):
            for row in rows:
                yield dict(row)


def list_commissions() -> list[dict[str, Any]]:
    """Compatibility helper returning every commission record."""

    return list(iter_commissions())


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


def count_audit_entries() -> int:
    with db_session() as conn:
        row = conn.execute("SELECT COUNT(*) FROM audit_log").fetchone()
    return int(row[0]) if row else 0


def list_audit_entries(
    limit: int = 500, *, offset: int = 0
) -> list[dict[str, Any]]:
    with db_session() as conn:
        rows = conn.execute(
            "SELECT * FROM audit_log ORDER BY id DESC LIMIT ? OFFSET ?",
            (
                max(1, min(int(limit), 5_000)),
                max(0, int(offset)),
            ),
        ).fetchall()
    return [dict(row) for row in rows]
