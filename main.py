import os, sys, io, json, zipfile, shutil, sqlite3, webbrowser, re
import tkinter as tk
import customtkinter as ctk
import time
import io
import bcrypt
import logging
from datetime import datetime, timedelta
from tkinter import ttk, messagebox, filedialog
from tkinter import simpledialog
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas 
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from PIL import Image, ImageOps, ImageTk
from logging.handlers import RotatingFileHandler

# ------------------ Paths/Config (EARLY so logging can use APP_DIR) ------------------
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

LOG_DIR = os.path.join(APP_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)
logger = logging.getLogger("svms")
logger.setLevel(logging.INFO)
_handler = RotatingFileHandler(
    os.path.join(LOG_DIR, "app.log"), maxBytes=512_000, backupCount=3, encoding="utf-8"
)
_formatter = logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
_handler.setFormatter(_formatter)
if not logger.handlers:
    logger.addHandler(_handler)

def get_conn():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous = NORMAL")
    return conn

def retry_db_operation(callable_fn, retries: int = 5, base_delay: float = 0.08):
    """
    Run a callable that performs DB I/O and retry on transient SQLITE 'database is locked' errors.
    callable_fn: a zero-arg callable that performs the DB action (e.g., lambda: cur.execute(...))
    Returns whatever callable_fn returns.
    Raises the last exception if retries exhausted.
    """
    last_exc = None
    for attempt in range(retries):
        try:
            return callable_fn()
        except sqlite3.OperationalError as e:
            last_exc = e
            msg = str(e).lower()
            if "database is locked" in msg or "database is busy" in msg:
                # small backoff, increase delay slightly on each attempt
                if attempt < retries - 1:
                    time.sleep(base_delay * (1 + attempt * 0.7))
                    continue
            # Other OperationalError: re-raise
            raise
    # exhausted
    if last_exc:
        raise last_exc

def restart_app():
    """Restart the current app (PyInstaller EXE or python script) with same args."""
    if getattr(sys, "frozen", False):  # packaged exe
        os.execl(sys.executable, sys.executable, *sys.argv[1:])
    else:  # run as script
        os.execl(sys.executable, sys.executable, os.path.abspath(__file__), *sys.argv[1:])

PHONE_MIN_DIGITS = 5
PHONE_MAX_DIGITS = 15
PHONE_ALLOWED_RE = re.compile(r"^\+?[0-9]+$")  # only digits and optional leading +

def normalize_phone(raw_phone, default_cc=None):
    """
    Clean and normalize a phone number string.

    Returns:
      - A normalized string in E.164-like form: '+<countrycode><number>' if possible,
        or digits-only string if no country code and default_cc not provided.
      - Returns None if input is invalid / cannot be normalized.

    Rules:
      - Accepts numbers with spaces, dashes, parentheses; removes those characters.
      - Allows a leading '+' for country code.
      - If the number begins with '00' it treats it as international prefix and converts to '+'.
      - If number starts with local trunk '0' and default_cc is provided, the 0 is removed and default_cc is prepended.
      - Ensures total digits (excluding '+') between PHONE_MIN_DIGITS and PHONE_MAX_DIGITS.
    """
    if not raw_phone:
        return None
    # Trim whitespace
    s = str(raw_phone).strip()
    # Replace common separators: spaces, dashes, parentheses, dots
    s = re.sub(r"[ \-\.\(\)]", "", s)
    # Convert leading international 00 to +
    if s.startswith("00"):
        s = "+" + s[2:]
    # Only allow digits and optional leading +
    if not PHONE_ALLOWED_RE.match(s):
        return None

    # Extract digits-only part
    digits = s[1:] if s.startswith("+") else s

    # If leading '0' trunk and default_cc provided, convert
    if not s.startswith("+") and default_cc and digits.startswith("0"):
        # remove leading 0 and prepend default country code
        digits = digits.lstrip("0")
        digits = default_cc + digits
        s = "+" + digits
    elif not s.startswith("+") and default_cc and not digits.startswith("0"):
        # No leading + and no trunk 0, but default_cc set: treat as local and prepend default cc
        # (only if it's sensible length)
        if PHONE_MIN_DIGITS <= len(digits) <= PHONE_MAX_DIGITS:
            s = "+" + default_cc + digits
            digits = default_cc + digits
    elif not s.startswith("+"):
        # no default_cc: leave as digits-only (no +)
        pass

    # Validate length
    if not (PHONE_MIN_DIGITS <= len(digits) <= PHONE_MAX_DIGITS):
        return None

    # Return normalized: prefer + form if present, else digits-only
    return s if s.startswith("+") else digits
    
def is_valid_phone(raw_phone, default_cc=None):
    """
    Returns True if phone can be normalized and meets length/character constraints.
    """
    try:
        norm = normalize_phone(raw_phone, default_cc=default_cc)
        return norm is not None
    except Exception:
        logger.exception("Error validating phone")
        return False

def validate_password_policy(pw: str) -> str | None:
    """
    Validate password according to the app policy.
    Returns:
      - None if password is acceptable,
      - String error message if not acceptable.
    Policy (same rules used in UI forced-change):
      - at least 10 chars
      - at least one uppercase
      - at least one lowercase
      - at least one digit
      - at least one symbol (non-alphanumeric)
    """
    if pw is None:
        return "Password cannot be empty."
    s = str(pw)
    if len(s) < 10:
        return "Password must be at least 10 characters."
    if not re.search(r"[A-Z]", s):
        return "Include at least one uppercase letter."
    if not re.search(r"[a-z]", s):
        return "Include at least one lowercase letter."
    if not re.search(r"\d", s):
        return "Include at least one digit."
    if not re.search(r"[^\w\s]", s):
        return "Include at least one symbol."
    return None
    
# --- User / role / column whitelists ---
ALLOWED_ROLES = {"admin", "sales assistant", "technician", "user"}  # update this set to match your app
USER_UPDATABLE_COLUMNS = {
    "username", "password_hash", "role", "must_change_pwd",
    "is_active", "full_name", "phone", "email", "note"
}
# you'll likely have slightly different column names — adjust USER_UPDATABLE_COLUMNS accordingly.

def _begin_immediate_transaction(conn):
    """
    For SQLite: acquire a reserved lock to prevent concurrent writers.
    Use BEFORE checks that must be atomic (like checking admin count then inserting).
    """
    try:
        cur = conn.cursor()
        cur.execute("BEGIN IMMEDIATE")
        return cur
    except Exception:
        # If BEGIN IMMEDIATE fails, fall back to normal transaction
        return conn.cursor()
        
# ---------- Wrapped text helpers for PDF ----------
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet

_styles = getSampleStyleSheet()
_styleN = _styles["Normal"]


def draw_wrapped(c, text, x, y, w, h, fontsize=10, bold=False, leading=None):
    style = _styleN.clone('wrap')
    style.fontName = "Helvetica-Bold" if bold else "Helvetica"
    style.fontSize = fontsize
    style.leading = leading if leading else fontsize + 2
    para = Paragraph((text or "-").replace("\n", "<br/>"), style)
    _, h_used = para.wrap(w, h)
    para.drawOn(c, x, y + h - h_used)
    return h_used


def draw_wrapped_top(c, text, x, top_y, w, fontsize=10, bold=False, leading=None):
    style = _styleN.clone('wrapTop')
    style.fontName = "Helvetica-Bold" if bold else "Helvetica"
    style.fontSize = fontsize
    style.leading = leading if leading else fontsize + 2
    para = Paragraph((text or "-").replace("\n", "<br/>"), style)
    _, h_used = para.wrap(w, 1000 * mm)
    para.drawOn(c, x, top_y - h_used)
    return h_used


# ------------------ Paths/Config ------------------
DB_FILE = os.path.join(APP_DIR, "vouchers.db")
PDF_DIR = os.path.join(APP_DIR, "pdfs")
IMAGES_DIR = os.path.join(APP_DIR, "images")  # legacy images root (kept for backup/restore)
STAFFS_ROOT = os.path.join(APP_DIR, "staffs")  # NEW: per-staff folders
os.makedirs(PDF_DIR, exist_ok=True)
os.makedirs(STAFFS_ROOT, exist_ok=True)

SHOP_NAME = "TONY.COM"
SHOP_ADDR = "TB4318, Lot 5, Block 31, Fajar Complex  91000 Tawau Sabah, Malaysia"
SHOP_TEL = "Tel : 089-763778, H/P: 0168260533"
LOGO_PATH = ""  # optional: path to logo image (png/jpg). leave blank to skip

DEFAULT_BASE_VID = 41000

# Default admin account (created if no admin exists).
# NOTE: Using 'admin123' is weak — the code will force password-change if policy
# rejects it. Consider using a stronger default or forcing immediate change.
DEFAULT_ADMIN_USERNAME = "tonycom"
DEFAULT_ADMIN_PASSWORD = "admin123"


def safe_folder_name(name: str) -> str:
    base = re.sub(r"[^A-Za-z0-9._ -]", "_", name).strip()
    return base or "unknown"


def staff_dirs_for(name: str):
    sname = safe_folder_name(name)
    base = os.path.join(STAFFS_ROOT, sname)
    com = os.path.join(base, "commissions")
    os.makedirs(com, exist_ok=True)  # ensures both base and commissions exist
    return base, com


# ------------------ Date helpers (DD-MM-YYYY for UI/PDF) ------------------
def _to_ui_date(dt: datetime) -> str:
    return dt.strftime("%d-%m-%Y")


def _to_ui_datetime_str(iso_str: str) -> str:
    try:
        dt = datetime.strptime(iso_str, "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%d-%m-%Y %H:%M:%S")
    except Exception:
        return iso_str


def _from_ui_date_to_sqldate(s: str):
    s = (s or "").strip()
    if not s:
        return None
    try:
        d = datetime.strptime(s, "%d-%m-%Y")
        return d.strftime("%Y-%m-%d")
    except Exception:
        return None


def open_path(path: str):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore
        elif sys.platform == "darwin":
            os.system(f"open '{path}'")
        else:
            os.system(f"xdg-open '{path}'")
    except Exception as e:
        messagebox.showerror("Open Error", f"Unable to open:\n{e}")


# ------------------ DB ------------------
BASE_SCHEMA = """
CREATE TABLE IF NOT EXISTS vouchers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    voucher_id TEXT UNIQUE,
    created_at TEXT,
    customer_name TEXT,
    contact_number TEXT,
    units INTEGER DEFAULT 1,
    particulars TEXT,
    problem TEXT,
    staff_name TEXT,
    status TEXT DEFAULT 'Pending',
    recipient TEXT,
    solution TEXT,
    pdf_path TEXT,
    technician_id TEXT,
    technician_name TEXT
);
CREATE TABLE IF NOT EXISTS staffs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    position TEXT,
    staff_id_opt TEXT,
    name TEXT UNIQUE,
    phone TEXT,
    photo_path TEXT,
    created_at TEXT,
    updated_at TEXT
);
CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE,
    role TEXT CHECK(role IN ('admin','sales assistant','user','technician')),
    password_hash BLOB,
    is_active INTEGER DEFAULT 1,
    must_change_pwd INTEGER DEFAULT 0,
    created_at TEXT,
    updated_at TEXT
);

CREATE TABLE IF NOT EXISTS commissions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_id INTEGER REFERENCES staffs(id) ON DELETE CASCADE,
    bill_type TEXT CHECK(bill_type IN ('CS','INV')),
    bill_no TEXT,
    total_amount REAL,
    commission_amount REAL,
    bill_image_path TEXT,
    created_at TEXT,
    updated_at TEXT
);
"""

def ensure_commissions_schema(conn):
    """
    Ensure the `commissions` table has the expected additional columns (voucher_id TEXT, note TEXT).
    Returns True if schema is present or successfully updated, False on failure.
    """
    try:
        cur = conn.cursor()

        # Best-effort backup helper if present
        try:
            if "_ensure_db_backup" in globals():
                try:
                    _ensure_db_backup(cur)
                except Exception:
                    logger.exception("Backup helper raised while ensuring commissions schema")
        except Exception:
            logger.exception("Failed trying to create DB backup before commissions schema migration")

        # If commissions table missing, warn and return False (unexpected)
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='commissions'")
        if cur.fetchone() is None:
            logger.warning("commissions table not found in DB; ensure migrations created it.")
            return False

        # Check columns present
        cur.execute("PRAGMA table_info('commissions')")
        cols = [r[1] for r in cur.fetchall()]

        # required columns we want to ensure (column_name -> SQL type)
        required = {"voucher_id": "TEXT", "note": "TEXT"}
        missing = [c for c in required.keys() if c not in cols]

        if not missing:
            logger.info("commissions schema OK (voucher_id and note present)")
            return True

        # Add missing columns safely
        try:
            for col in missing:
                typ = required[col]
                if "add_column_safe" in globals():
                    add_column_safe(cur, "commissions", col, typ)
                else:
                    if not re.match(r"^[A-Za-z0-9_]+$", col):
                        raise ValueError(f"Invalid column name '{col}'")
                    cur.execute(f"ALTER TABLE commissions ADD COLUMN {col} {typ}")
            conn.commit()
            logger.info("Added missing columns to commissions: %s", ", ".join(missing))
            return True
        except Exception:
            logger.exception("Failed to add missing columns to commissions: %s", missing)
            conn.rollback()
            return False

    except Exception:
        logger.exception("Unexpected failure in ensure_commissions_schema")
        try:
            conn.rollback()
        except Exception:
            pass
        return False
        
def _column_exists(cur, table, column):
    cur.execute(f"PRAGMA table_info({table})")
    return any(row[1] == column for row in cur.fetchall())


def _get_setting(cur, key, default=None):
    cur.execute("SELECT value FROM settings WHERE key=?", (key,))
    row = cur.fetchone()
    if row and row[0] is not None:
        return row[0]
    if default is not None:
        cur.execute("INSERT OR IGNORE INTO settings(key,value) VALUES (?,?)", (key, str(default)))
        return str(default)
    return None

def _safe_voucher_int_expr(col_name):
    """Return a SQL expression that safely casts a voucher id (text) to integer,
    returning NULL for non-numeric values to avoid CAST errors when voucher_id
    contains non-numeric characters.
    Usage in SQL: ORDER BY _safe_voucher_int_expr('voucher_id') DESC
    """
    return "CASE WHEN {col} GLOB '[0-9]*' THEN CAST({col} AS INTEGER) ELSE NULL END".replace("{col}", col_name)

def add_column_safe(cur, table, col, typ):
    """Safely add a column to a table after validating the column name and type against whitelists.
    This prevents accidental SQL injection via f-strings for identifiers.
    """
    # NOTE: 'note' added to commissions allowed set
    allowed_cols = {
        "vouchers": {"technician_id","technician_name","ref_bill","ref_bill_date",
                     "amount_rm","tech_commission","reminder_pickup_1","reminder_pickup_2","reminder_pickup_3"},
        "commissions": {"voucher_id", "note"}
    }

    allowed_types = {"TEXT","REAL","INTEGER","BLOB"}
    col_clean = col.strip()
    typ_clean = typ.strip().upper()
    if table not in allowed_cols:
        raise ValueError(f"add_column_safe: table not allowed: {table}")
    if col_clean not in allowed_cols[table]:
        raise ValueError(f"add_column_safe: unexpected column name: {col_clean}")
    if typ_clean not in allowed_types:
        raise ValueError(f"add_column_safe: unexpected column type: {typ_clean}")
    sql = f"ALTER TABLE {table} ADD COLUMN {col_clean} {typ_clean}"
    cur.execute(sql)

def _set_setting(cur, key, value):
    cur.execute("INSERT OR REPLACE INTO settings(key,value) VALUES (?,?)", (key, str(value)))


def _hash_pwd(pwd: str) -> bytes:
    return bcrypt.hashpw(pwd.encode("utf-8"), bcrypt.gensalt())


def _verify_pwd(pwd: str, hp: bytes) -> bool:
    try:
        return bcrypt.checkpw(pwd.encode("utf-8"), hp)
    except Exception:
        return False

STATUS_CANONICAL = {
    "open": "open",
    "pending": "pending",
    "in progress": "in progress",
    "in_progress": "in progress",
    "completed": "completed",
    "done": "completed",
    "closed": "closed",
    "cancelled": "cancelled",
    "canceled": "cancelled",
    "on hold": "on hold",
    "returned": "returned"
}

def normalize_status_for_search(raw_status, fuzzy_status=False):
    """
    Normalize a user-provided status string for safe searching.
    - raw_status: the string entered by the user (may be None/empty)
    - fuzzy_status: when True allow a safe LIKE ('%...%') fallback on the full token (NOT prefix).
    Returns:
      (sql_clause_fragment, params)
      - sql_clause_fragment: e.g. "AND LOWER(status) = ?"
      - params: list of params to bind
    If raw_status is falsy, returns ("", [])
    """
    if not raw_status:
        return "", []

    s = raw_status.strip().lower()
    # Map exact synonyms to canonical
    if s in STATUS_CANONICAL:
        canon = STATUS_CANONICAL[s]
        return "AND LOWER(status) = ?", [canon.lower()]

    # If user supplied a comma-separated list, support explicit list (clean & whitelist)
    if "," in s:
        items = [it.strip().lower() for it in s.split(",") if it.strip()]
        mapped = []
        for it in items:
            if it in STATUS_CANONICAL:
                mapped.append(STATUS_CANONICAL[it].lower())
            else:
                # unknown token -> keep as-is only if fuzzy permitted
                if fuzzy_status:
                    mapped.append(it)
        if mapped:
            placeholders = ",".join("?" for _ in mapped)
            return f"AND LOWER(status) IN ({placeholders})", mapped

    # As a last resort, if fuzzy_status True, do a safe LIKE on the full token
    if fuzzy_status:
        return "AND LOWER(status) LIKE ?", [f"%{s}%"]

    # Unknown/ambiguous status and no fuzzy allowed: return a clause that matches nothing
    # (this is safer than guessing)
    return "AND 0", []

def init_db(db_path=DB_FILE):
    """
    Initialize DB and run safe migrations.
    - Creates a timestamped DB backup before destructive changes (if helper exists).
    - Ensures users table migration (adds 'sales assistant' role).
    - Adds missing voucher columns safely.
    - Ensures commissions schema (voucher_id column) via ensure_commissions_schema.
    - Creates useful indices and enforces uniqueness after cleaning legacy duplicates.
    - Creates a default admin account if none exists (may force password change).
    Returns sqlite3.Connection on success.
    Raises on failure.
    """
    conn = None
    try:
        # Open connection early so we can run scripts and PRAGMAs safely
        conn = sqlite3.connect(db_path, timeout=30)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()

        # Apply base schema (idempotent). Run inside try/except so we continue if it fails.
        try:
            cur.executescript(BASE_SCHEMA)
            conn.commit()
        except Exception:
            logger.exception("Failed to apply BASE_SCHEMA (continuing)")

        # Helper to safely cast voucher_id to integer for ordering/de-dup purposes
        def _safe_vid_expr(col_name="voucher_id"):
            return f"CASE WHEN {col_name} GLOB '[0-9]*' THEN CAST({col_name} AS INTEGER) ELSE NULL END"

        # --- MIGRATION: ensure users.role allows 'sales assistant' ---
        try:
            cur.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='users'")
            row = cur.fetchone()
            ddl = (row[0] or "") if row else ""
            if "sales assistant" not in ddl:
                # attempt backup if available
                try:
                    if "_ensure_db_backup" in globals():
                        try:
                            _ensure_db_backup(cur)
                        except Exception:
                            logger.exception("Failed to create DB backup before users table migration")
                except Exception:
                    logger.exception("Unexpected error invoking backup helper before users migration")

                # Rename and recreate users table with sales assistant role allowed
                cur.execute("ALTER TABLE users RENAME TO users_old")
                cur.execute("""
                    CREATE TABLE users (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        username TEXT UNIQUE,
                        role TEXT CHECK(role IN ('admin','sales assistant','user','technician')),
                        password_hash BLOB,
                        is_active INTEGER DEFAULT 1,
                        must_change_pwd INTEGER DEFAULT 0,
                        created_at TEXT,
                        updated_at TEXT
                    )
                """)
                cur.execute("""
                    INSERT INTO users(id, username, role, password_hash, is_active, must_change_pwd, created_at, updated_at)
                    SELECT id,
                           username,
                           CASE WHEN role='supervisor' THEN 'sales assistant' ELSE role END,
                           password_hash, is_active, must_change_pwd, created_at, updated_at
                    FROM users_old
                """)
                cur.execute("DROP TABLE users_old")
                conn.commit()
                logger.info("users table migration (sales assistant role) completed.")
        except Exception:
            logger.exception("users table migration failed; check DB state")

        # Trim status whitespace (safe no-op if column missing)
        try:
            cur.execute("UPDATE vouchers SET status = TRIM(status)")
            conn.commit()
        except Exception:
            logger.exception("Failed to TRIM(vouchers.status) — continuing")

        # --- Add missing voucher columns safely (whitelisted) ---
        voucher_columns = [
            ("technician_id", "TEXT"),
            ("technician_name", "TEXT"),
            ("ref_bill", "TEXT"),
            ("ref_bill_date", "TEXT"),
            ("amount_rm", "REAL"),
            ("tech_commission", "REAL"),
            ("reminder_pickup_1", "TEXT"),
            ("reminder_pickup_2", "TEXT"),
            ("reminder_pickup_3", "TEXT"),
        ]
        for col, typ in voucher_columns:
            try:
                cur.execute(f"PRAGMA table_info(vouchers)")
                existing_cols = [r[1] for r in cur.fetchall()]
                if col not in existing_cols:
                    if "add_column_safe" in globals():
                        add_column_safe(cur, "vouchers", col, typ)
                    else:
                        if not re.match(r"^[A-Za-z0-9_]+$", col):
                            raise ValueError(f"Unsafe column name: {col}")
                        cur.execute(f"ALTER TABLE vouchers ADD COLUMN {col} {typ}")
                    logger.info("Added column %s to vouchers", col)
            except Exception:
                logger.exception("Failed adding column %s to vouchers (continuing)", col)
        try:
            conn.commit()
        except Exception:
            logger.exception("commit failed after adding voucher columns")

        # --- Ensure commissions schema (voucher_id column) ---
        try:
            if "ensure_commissions_schema" in globals():
                ok = ensure_commissions_schema(conn)
                if not ok:
                    logger.warning("ensure_commissions_schema reported failure. Check DB and logs.")
        except Exception:
            logger.exception("ensure_commissions_schema call failed during init_db()")

        # --- Indexes and uniqueness constraints; create non-unique indexes ---
        try:
            cur.execute("CREATE INDEX IF NOT EXISTS idx_vouchers_created ON vouchers(created_at)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_vouchers_status ON vouchers(status)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_vouchers_customer ON vouchers(customer_name)")
            conn.commit()
        except Exception:
            logger.exception("Failed creating non-unique indexes; continuing")

        # --- Clean legacy duplicates for vouchers.ref_bill (destructive) ---
        try:
            try:
                if "_ensure_db_backup" in globals():
                    _ensure_db_backup(cur)
            except Exception:
                logger.exception("Failed to create DB backup before cleaning vouchers.ref_bill duplicates")

            cur.execute("BEGIN")
            safe_vid = _safe_vid_expr("voucher_id")
            cur.execute(f"""
                DELETE FROM vouchers
                WHERE ref_bill IS NOT NULL AND ref_bill <> ''
                  AND {safe_vid} NOT IN (
                    SELECT MAX({safe_vid}) FROM vouchers
                    WHERE ref_bill IS NOT NULL AND ref_bill <> ''
                    GROUP BY LOWER(ref_bill)
                  )
            """)
            conn.commit()
            logger.info("Legacy voucher duplicates cleaned (ref_bill).")
        except Exception:
            logger.exception("Failed to clean voucher duplicates; performing ROLLBACK")
            try:
                conn.rollback()
            except Exception:
                pass

        # Create unique index for ref_bill
        try:
            cur.execute("""
                CREATE UNIQUE INDEX IF NOT EXISTS idx_vouchers_ref_bill_unique_ci
                ON vouchers(LOWER(ref_bill))
                WHERE ref_bill IS NOT NULL AND ref_bill <> ''
            """)
            conn.commit()
        except Exception:
            logger.exception("Failed creating unique index for vouchers.ref_bill; check duplicates/logs")

        # --- Clean legacy duplicates for commissions ---
        try:
            try:
                if "_ensure_db_backup" in globals():
                    _ensure_db_backup(cur)
            except Exception:
                logger.exception("Failed to create DB backup before cleaning commissions duplicates")

            cur.execute("BEGIN")
            cur.execute("""
                DELETE FROM commissions
                WHERE id NOT IN (
                    SELECT MAX(id) FROM commissions
                    WHERE bill_no IS NOT NULL AND bill_no <> ''
                    GROUP BY bill_type, LOWER(bill_no)
                )
            """)
            conn.commit()
            logger.info("Legacy commissions duplicates cleaned.")
        except Exception:
            logger.exception("Failed to clean commissions duplicates; rolling back")
            try:
                conn.rollback()
            except Exception:
                pass

        # Create unique index for commissions (bill_type + LOWER(bill_no))
        try:
            cur.execute("""
                CREATE UNIQUE INDEX IF NOT EXISTS idx_comm_unique_global
                ON commissions(bill_type, LOWER(bill_no))
                WHERE bill_no IS NOT NULL AND bill_no <> ''
            """)
            conn.commit()
        except Exception:
            logger.exception("Failed creating unique index for commissions; check duplicates/logs")

        # --- Ensure base_vid setting exists ---
        try:
            _get_setting(cur, "base_vid", DEFAULT_BASE_VID)
            conn.commit()
        except Exception:
            logger.exception("Failed to ensure base_vid setting")

        # --- Ensure at least one admin user exists; create default if none ---
        try:
            cur.execute("SELECT COUNT(*) FROM users WHERE role='admin'")
            admin_count = cur.fetchone()[0]
            if admin_count == 0:
                # Use module-level defaults if present
                default_user = globals().get("DEFAULT_ADMIN_USERNAME", "tonycom")
                default_pwd = globals().get("DEFAULT_ADMIN_PASSWORD", "admin123")

                must_change = 0
                try:
                    policy_err = validate_password_policy(default_pwd)
                    if policy_err:
                        must_change = 1
                        logger.warning(
                            "Default admin password does not satisfy password policy: %s. "
                            "Creating account with must_change_pwd=1.",
                            policy_err
                        )
                except Exception:
                    must_change = 1
                    logger.exception("validate_password_policy raised exception; setting must_change_pwd=1")

                ts = datetime.now().isoformat(sep=' ', timespec='seconds')
                try:
                    hashed = _hash_pwd(default_pwd)
                    cur.execute(
                        "INSERT OR IGNORE INTO users (username, role, password_hash, must_change_pwd, created_at, updated_at) VALUES (?,?,?,?,?,?)",
                        (default_user, "admin", hashed, must_change, ts, ts)
                    )
                    conn.commit()
                    if must_change:
                        logger.warning("Default admin '%s' created; user must change password at first login.", default_user)
                    else:
                        logger.info("Default admin '%s' created with default password.", default_user)
                except sqlite3.IntegrityError:
                    # Already exists - ignore
                    logger.info("Admin user already exists (race or unique constraint).")
                except Exception:
                    logger.exception("Failed to create default admin user.")
        except Exception:
            logger.exception("Failed to ensure default admin user")

        # All migrations completed successfully; return the live connection
        return conn

    except Exception:
        logger.exception("init_db failed")
        if conn:
            try:
                conn.close()
            except Exception:
                pass
        raise

def bind_commission_to_voucher(commission_id, voucher_id):
    """
    Bind a commission row to a voucher by setting commissions.voucher_id = ? for the given commission_id.

    Returns True on success, False on failure.
    This variant opens its own DB connection (so callers just pass two args).
    """
    try:
        # Use context manager so connection is always closed and transactions roll back on exception
        with get_conn() as conn:
            cur = conn.cursor()

            # Check schema: make sure voucher_id column exists
            cur.execute("PRAGMA table_info('commissions')")
            cols = [r[1] for r in cur.fetchall()]
            if "voucher_id" not in cols:
                logger.error(
                    "bind_commission_to_voucher: 'voucher_id' column missing from commissions table. "
                    "Please run init_db() or ensure migrations have been applied."
                )
                return False

            # Best-effort backup helper (if present)
            try:
                if "_ensure_db_backup" in globals():
                    try:
                        _ensure_db_backup(cur)
                    except Exception:
                        logger.exception("Backup helper failed inside bind_commission_to_voucher")
            except Exception:
                logger.exception("Unexpected error invoking backup helper inside bind_commission_to_voucher")

            # Defensive checks: commission exists?
            cur.execute("SELECT id, voucher_id FROM commissions WHERE id = ?", (commission_id,))
            row = cur.fetchone()
            if not row:
                logger.warning("bind_commission_to_voucher: no commission row found for id=%s", commission_id)
                return False

            existing_bound_vid = row[1]
            # If already bound to same voucher, treat as success (idempotent)
            if existing_bound_vid and str(existing_bound_vid) == str(voucher_id):
                logger.info("bind_commission_to_voucher: commission %s already bound to voucher %s (idempotent)", commission_id, voucher_id)
                return True

            # If voucher_id already bound to another commission, fail (avoid duplicates)
            cur.execute("SELECT id FROM commissions WHERE voucher_id = ? AND id <> ?", (voucher_id, commission_id))
            other = cur.fetchone()
            if other:
                logger.warning("bind_commission_to_voucher: voucher %s already bound to commission id=%s", voucher_id, other[0])
                return False

            # Perform the binding. Use retry_db_operation if available to handle transient locks.
            def _do_update():
                cur.execute("UPDATE commissions SET voucher_id = ?, updated_at = ? WHERE id = ?",
                            (voucher_id, datetime.now().isoformat(sep=' ', timespec='seconds'), commission_id))
                return cur.rowcount

            try:
                if "retry_db_operation" in globals():
                    rows = retry_db_operation(_do_update)
                else:
                    rows = _do_update()
            except Exception:
                logger.exception("Failed to update commission voucher_id; operation aborted")
                # Let context manager rollback and close
                return False

            if not rows:
                logger.warning("bind_commission_to_voucher: update affected 0 rows for commission id=%s", commission_id)
                return False

            logger.info("Bound commission id=%s to voucher_id=%s", commission_id, voucher_id)
            # commit happens automatically when leaving the 'with' block without exception
            return True

    except Exception:
        # No assumption that conn exists here — safe to just log and return False
        logger.exception("Unexpected error in bind_commission_to_voucher")
        return False

def _read_base_vid():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT value FROM settings WHERE key='base_vid'")
    row = cur.fetchone();
    conn.close()
    return int(row[0]) if row and row[0] else DEFAULT_BASE_VID


def next_voucher_id():
    conn = get_conn();
    cur = conn.cursor()
    cur.execute("SELECT MAX(CAST(voucher_id AS INTEGER)) FROM vouchers")
    row = cur.fetchone()
    if not row or row[0] is None:
        base = _get_setting(cur, "base_vid", DEFAULT_BASE_VID)
        conn.commit();
        conn.close()
        return str(base)
    nxt = str(int(row[0]) + 1)
    conn.close()
    return nxt


# ---- Recipient ops ----
def list_staffs_names():
    conn = get_conn();
    cur = conn.cursor()
    cur.execute("SELECT name FROM staffs ORDER BY name COLLATE NOCASE ASC")
    rows = [r[0] for r in cur.fetchall()]
    conn.close();
    return rows


def add_staff_simple(name: str):
    name = (name or "").strip()
    if not name: return False
    conn = get_conn();
    cur = conn.cursor()
    try:
        cur.execute("""INSERT OR IGNORE INTO staffs
                       (position, staff_id_opt, name, phone, photo_path, created_at, updated_at)
                       VALUES ('Technician', '', ?, '', '', ?, ?)""",
                    (name, datetime.now().isoformat(sep=' ', timespec='seconds'),
                     datetime.now().isoformat(sep=' ', timespec='seconds')))
        conn.commit()
        # Create folders
        staff_dirs_for(name)
    finally:
        conn.close()
    return True


def delete_staff_simple(name: str):
    conn = get_conn();
    cur = conn.cursor()
    cur.execute("DELETE FROM staffs WHERE name = ?", (name,))
    conn.commit();
    conn.close()


# ------------------ PDF helpers ------------------
def _draw_header(c, left, right, top_y, voucher_id):
    y = top_y
    LOGO_W = 28 * mm;
    LOGO_H = 18 * mm
    if LOGO_PATH and os.path.exists(LOGO_PATH):
        try:
            c.drawImage(LOGO_PATH, right - LOGO_W, y - LOGO_H, LOGO_W, LOGO_H, preserveAspectRatio=True, mask='auto')
        except Exception as e:
                logger.exception("Caught exception", exc_info=e)
                pass
    c.setFont("Helvetica-Bold", 14);
    c.drawString(left, y, SHOP_NAME)
    c.setFont("Helvetica", 9.2)
    c.drawString(left, y - 5.0 * mm, SHOP_ADDR)
    c.drawString(left, y - 9.0 * mm, SHOP_TEL)
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString((left + right) / 2, y - 16.0 * mm, "SERVICE VOUCHER")
    c.drawRightString(right, y - 16.0 * mm, f"No : {voucher_id}")
    return y - 16.0 * mm


def _draw_datetime_row(c, left, right, base_y, created_at):
    try:
        dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
        date_str = dt.strftime("%d-%m-%Y");
        time_str = dt.strftime("%H:%M:%S")
    except Exception:
        date_str = created_at[:10];
        time_str = created_at[11:19]
    c.setFont("Helvetica", 10)
    c.drawString(left, base_y - 8.0 * mm, "Date :")
    c.drawString(left + 18 * mm, base_y - 8.0 * mm, date_str)
    c.drawRightString(right - 27 * mm, base_y - 8.0 * mm, "Time In :")
    c.drawRightString(right, base_y - 8.0 * mm, time_str)


def _draw_main_table(c, left, right, top_table, customer_name, particulars, units, contact_number, problem):
    qty_col_w = 20 * mm
    left_col_w = 74 * mm
    middle_col_w = (right - left) - left_col_w - qty_col_w
    name_col_x = left + left_col_w
    qty_col_x = right - qty_col_w
    row1_h = 20 * mm
    row2_h = 20 * mm
    bottom_table = top_table - (row1_h + row2_h)
    mid_y = top_table - row1_h
    pad = 3 * mm

    c.rect(left, bottom_table, right - left, (row1_h + row2_h), stroke=1, fill=0)
    c.line(name_col_x, top_table, name_col_x, bottom_table)
    c.line(qty_col_x, top_table, qty_col_x, mid_y)
    c.line(left, mid_y, right, mid_y)

    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left + pad, top_table - pad - 8, "CUSTOMER NAME")
    draw_wrapped(c, customer_name, left + pad, mid_y + pad,
                 w=left_col_w - 2 * pad, h=row1_h - 2 * pad - 10, fontsize=10)

    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(name_col_x + pad, top_table - pad - 8, "PARTICULARS")
    draw_wrapped(c, particulars or "-", name_col_x + pad, mid_y + pad,
                 w=middle_col_w - 2 * pad, h=row1_h - 2 * pad - 10, fontsize=10)

    c.setFont("Helvetica-Bold", 10.4)
    c.drawCentredString(qty_col_x + qty_col_w / 2, top_table - pad - 8, "QTY")
    c.setFont("Helvetica", 11)
    c.drawCentredString(qty_col_x + qty_col_w / 2, mid_y + (row1_h / 2) - 3, str(max(1, int(units or 1))))

    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left + pad, mid_y - pad - 8, "TEL")
    draw_wrapped(c, contact_number, left + pad, bottom_table + pad,
                 w=left_col_w - 2 * pad, h=row2_h - 2 * pad - 10, fontsize=10)

    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(name_col_x + pad, mid_y - pad - 8, "PROBLEM")
    draw_wrapped(c, problem or "-", name_col_x + pad, bottom_table + pad,
                 w=(middle_col_w + qty_col_w) - 2 * pad, h=row2_h - 2 * pad - 10, fontsize=10)

    return bottom_table, left_col_w


def _draw_policies_and_signatures(c, left, right, bottom_table, left_col_w, recipient, voucher_id, customer_name,
                                  contact_number, date_str):
    ack_text = ("WE HEREBY CONFIRM THAT THE MACHINE WAS SERVICED AND REPAIRED SATISFACTORILY")
    ack_left = left + left_col_w + 10 * mm
    ack_right = right - 6 * mm
    ack_top_y = bottom_table - 5 * mm
    draw_wrapped_top(c, ack_text, ack_left, ack_top_y, max(20 * mm, ack_right - ack_left), fontsize=9, bold=True,
                     leading=11)

    y_rec = bottom_table - 9 * mm
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left, y_rec, "RECIPIENT :")
    label_w = c.stringWidth("RECIPIENT :", "Helvetica-Bold", 10.4)
    line_x0 = left + label_w + 6
    line_x1 = left + left_col_w - 2 * mm
    line_y = y_rec - 3 * mm
    c.line(line_x0, line_y, line_x1, line_y)
    if recipient:
        c.setFont("Helvetica", 9);
        c.drawString(line_x0 + 1 * mm, line_y + 2.2 * mm, recipient)

    policies_top = y_rec - 7 * mm
    policies_w = left_col_w - 1.5 * mm

    p1 = "Kindly collect your goods within <font color='red' size='9'>60 days</font> from date of sending for repair."
    used_h = draw_wrapped_top(c, p1, left, policies_top, policies_w, fontsize=6.5, leading=10)
    y_cursor = policies_top - used_h - 2
    used_h = draw_wrapped_top(c, "A) We do not hold ourselves responsible for any loss or damage.", left, y_cursor,
                              policies_w, fontsize=6.5, leading=10)
    y_cursor -= used_h - 1
    used_h = draw_wrapped_top(c, "B) We reserve our right to sell off the goods to cover our cost and loss.", left,
                              y_cursor, policies_w, fontsize=6.5, leading=10)
    y_cursor -= used_h + 2
    p4 = (
        "MINIMUM <font color='red' size='9'><b>RM60.00</b></font> WILL BE CHARGED ON TROUBLESHOOTING, INSPECTION AND SERVICE ON ALL KIND OF HARDWARE AND SOFTWARE.")
    used_h = draw_wrapped_top(c, p4, left, y_cursor, policies_w, fontsize=8, leading=10)
    y_cursor -= used_h - 1
    used_h = draw_wrapped_top(c, "PLEASE BRING ALONG THIS SERVICE VOUCHER TO COLLECT YOUR GOODS", left, y_cursor,
                              policies_w, fontsize=8, leading=10)
    y_cursor -= used_h - 1
    used_h = draw_wrapped_top(c, "NO ATTENTION GIVEN WITHOUT SERVICE VOUCHER", left, y_cursor, policies_w, fontsize=8,
                              leading=10)
    policies_bottom = y_cursor - used_h

    SIG_LINE_W = 45 * mm;
    SIG_GAP = 6 * mm
    y_sig = max(policies_bottom + 4 * mm, (A4[1] / 2) - 20 * mm)
    sig_left_start = right - (2 * SIG_LINE_W + SIG_GAP)
    c.line(sig_left_start, y_sig, sig_left_start + SIG_LINE_W, y_sig)
    c.setFont("Helvetica", 8.8);
    c.drawString(sig_left_start, y_sig - 3.6 * mm, "CUSTOMER SIGNATURE")
    right_line_x0 = sig_left_start + SIG_LINE_W + SIG_GAP
    c.line(right_line_x0, y_sig, right_line_x0 + SIG_LINE_W, y_sig)
    c.drawString(right_line_x0, y_sig - 3.6 * mm, "DATE COLLECTED")


def generate_pdf(voucher_id, customer_name, contact_number, units,
                 particulars, problem, staff_name, status, created_at, recipient):
    """
    Safe PDF generator:
    - Writes to a temp file (same directory) and atomically replaces the final file.
    - Returns the final (absolute) filename on success.
    """
    os.makedirs(PDF_DIR, exist_ok=True)
    final_pdf = os.path.join(PDF_DIR, f"voucher_{voucher_id}.pdf")
    tmp_pdf = final_pdf + ".part"
    try:
        # Write into the temporary file first
        c = rl_canvas.Canvas(tmp_pdf, pagesize=A4)
        width, height = A4
        left, right, top_y = 12 * mm, width - 12 * mm, height - 15 * mm
        title_baseline = _draw_header(c, left, right, top_y, voucher_id)
        _draw_datetime_row(c, left, right, title_baseline, created_at)
        top_table = title_baseline - 12 * mm
        bottom_table, left_col_w = _draw_main_table(c, left, right, top_table, customer_name, particulars, units,
                                                    contact_number, problem)
        try:
            dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
            date_str = dt.strftime("%d-%m-%Y")
        except Exception:
            date_str = created_at[:10]
        _draw_policies_and_signatures(c, left, right, bottom_table, left_col_w, recipient, voucher_id, customer_name,
                                      contact_number, date_str)
        c.showPage()
        c.save()
        # Atomically move into final place
        try:
            os.replace(tmp_pdf, final_pdf)
        except Exception:
            # If os.replace is not available on platform, fall back to rename
            try:
                os.rename(tmp_pdf, final_pdf)
            except Exception:
                # If final replace fails, keep the temp (so no data loss) and raise
                logger.exception("Failed to move temporary PDF into place")
                raise
        return final_pdf
    except Exception:
        # Clean up temp file if it exists (best-effort)
        try:
            if os.path.exists(tmp_pdf):
                os.remove(tmp_pdf)
        except Exception:
            pass
        raise

# ------------------ DB ops (vouchers) ------------------
def add_voucher(
    customer_name,
    contact_number,
    units,
    particulars,
    problem,
    staff_name,
    recipient="",
    solution="",
    technician_id="",
    technician_name="",
    status=None,
    ref_bill=None,
    ref_bill_date=None,
    amount_rm=None,
    tech_commission=None,
):
    """
    Create voucher row and generate PDF. Robust INSERT: detect which columns
    exist in the vouchers table and only insert those columns to avoid schema
    mismatch errors on older DBs.
    """
    with get_conn() as conn:
        cur = conn.cursor()
        voucher_id = next_voucher_id()
        created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status = status or "Pending"

        # Generate to a temp name first (keeps previous safety logic)
        final_pdf = os.path.join(PDF_DIR, f"voucher_{voucher_id}.pdf")
        temp_pdf = final_pdf + ".part"
        try:
            c = rl_canvas.Canvas(temp_pdf, pagesize=A4)
            c.showPage(); c.save()
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
            pdf_path = generate_pdf(
                voucher_id, customer_name, contact_number, units,
                particulars, problem, staff_name, status, created_at, recipient
            )
            if pdf_path != final_pdf and os.path.exists(pdf_path):
                final_pdf = pdf_path
        except Exception:
            # re-raise so callers get an exception
            raise

        # Determine which columns actually exist in the vouchers table
        cur.execute("PRAGMA table_info(vouchers)")
        existing_cols = [r[1] for r in cur.fetchall()]

        # Map of candidate columns -> values
        col_map = [
            ("voucher_id", voucher_id),
            ("created_at", created_at),
            ("customer_name", customer_name),
            ("contact_number", contact_number),
            ("units", units),
            ("particulars", particulars),
            ("problem", problem),
            ("staff_name", staff_name),
            ("status", status),
            ("recipient", recipient),
            ("solution", solution),
            ("pdf_path", final_pdf),
            ("technician_id", technician_id),
            ("technician_name", technician_name),
            ("ref_bill", ref_bill if ref_bill is not None else ""),
            ("ref_bill_date", ref_bill_date if ref_bill_date is not None else None),
            ("amount_rm", float(amount_rm) if amount_rm is not None else None),
            ("tech_commission", float(tech_commission) if tech_commission is not None else None),
        ]

        cols_to_insert = []
        params = []
        for col, val in col_map:
            if col in existing_cols:
                cols_to_insert.append(col)
                params.append(val)

        if not cols_to_insert:
            raise RuntimeError("No writable columns found in vouchers table.")

        placeholders = ",".join("?" for _ in params)
        sql = f"INSERT INTO vouchers ({', '.join(cols_to_insert)}) VALUES ({placeholders})"
        try:
            cur.execute(sql, params)
        except Exception:
            # on failure, attempt to remove pdf and re-raise
            try:
                if os.path.exists(final_pdf):
                    os.remove(final_pdf)
            except Exception as e:
                logger.exception("Caught exception", exc_info=e)
                pass
            raise

        conn.commit()
        return voucher_id, final_pdf

def update_voucher_fields(voucher_id, **fields):
    if not fields:
        return
    if "status" in fields:
        _validate_status_or_raise(fields["status"])
    with get_conn() as conn:
        cur = conn.cursor()
        cols = []
        params = []
        for k, v in fields.items():
            cols.append(f"{k}=?")
            params.append(v)
        params.append(voucher_id)
        cur.execute(f"UPDATE vouchers SET {', '.join(cols)} WHERE voucher_id = ?", params)
        conn.commit()

def bulk_update_status(voucher_ids, new_status):
    _validate_status_or_raise(new_status)
    if not voucher_ids: return
    conn = get_conn();
    cur = conn.cursor()
    cur.executemany("UPDATE vouchers SET status=? WHERE voucher_id=?", [(new_status, vid) for vid in voucher_ids])
    conn.commit();
    conn.close()

def _parse_ui_date_to_iso(d: str):
    """
    Convert a date string from UI into YYYY-MM-DD for SQLite DATE() comparisons.
    Accepts:
      - DD-MM-YYYY  -> converts to YYYY-MM-DD
      - D-M-YYYY    -> converts (single-digit day/month)
      - YYYY-MM-DD  -> returned as-is
      - empty/invalid -> returns empty string
    """
    if not d:
        return ""
    s = d.strip()
    # If already ISO-ish (YYYY-...), return as-is
    m_iso = re.match(r"^\d{4}-\d{2}-\d{2}$", s)
    if m_iso:
        return s
    # Try DD-MM-YYYY or D-M-YYYY
    m = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$", s)
    if m:
        dd, mm, yyyy = m.group(1).zfill(2), m.group(2).zfill(2), m.group(3)
        return f"{yyyy}-{mm}-{dd}"
    # Unknown format -> return empty string to avoid broken DATE(...) comparisons
    return ""

def _build_search_sql(filters: dict):
    """
    Build SQL and params from the given filters dictionary.

    Defensive: normalize status, parse UI dates (DD-MM-YYYY -> YYYY-MM-DD),
    and handle status_list, limit/offset and numeric ordering on voucher_id.
    """
    # defensive default
    filters = filters or {}

    sql = (
        "SELECT voucher_id, created_at, customer_name, contact_number, units, "
        "recipient, technician_id, technician_name, status, solution, pdf_path "
        "FROM vouchers WHERE 1=1"
    )
    params = []

    # voucher id (partial)
    vid = (filters.get("voucher_id") or "").strip()
    if vid:
        sql += " AND voucher_id LIKE ?"
        params.append(f"%{vid}%")

    # customer name (case-insensitive)
    cname = (filters.get("customer_name") or "").strip()
    if cname:
        sql += " AND LOWER(customer_name) LIKE ?"
        params.append(f"%{cname.lower()}%")

    # contact (partial)
    contact = (filters.get("contact_number") or "").strip()
    if contact:
        sql += " AND contact_number LIKE ?"
        params.append(f"%{contact}%")

    # status_list (whitelisted already) - takes precedence if provided and non-empty
    status_list = filters.get("status_list")
    if status_list:
        status_list_normalized = [s.strip().lower() for s in status_list if (s or "").strip()]
        if status_list_normalized:
            placeholders = ",".join("?" for _ in status_list_normalized)
            sql += f" AND LOWER(TRIM(status)) IN ({placeholders})"
            params.extend(status_list_normalized)
    else:
        # status: treat "all" (case-insensitive) as no-filter
        status_raw = (filters.get("status") or "").strip()
        if status_raw and status_raw.lower() != "all":
            sql += " AND LOWER(TRIM(status)) = ?"
            params.append(status_raw.lower())

    # date filters: convert UI-formatted dates (DD-MM-YYYY) into ISO (YYYY-MM-DD)
    date_from_raw = (filters.get("date_from") or "").strip()
    date_to_raw = (filters.get("date_to") or "").strip()

    date_from = _parse_ui_date_to_iso(date_from_raw)
    date_to = _parse_ui_date_to_iso(date_to_raw)

    if date_from:
        sql += " AND DATE(created_at) >= DATE(?)"
        params.append(date_from)
    if date_to:
        sql += " AND DATE(created_at) <= DATE(?)"
        params.append(date_to)

    # ordering: prefer numeric ordering on voucher_id (if helper available),
    # otherwise use created_at DESC then voucher_id DESC
    try:
        numeric_vid_expr = _safe_voucher_int_expr("voucher_id")  # your helper
        sql += f" ORDER BY {numeric_vid_expr} DESC, voucher_id DESC"
    except Exception:
        sql += " ORDER BY created_at DESC"

    # limit / offset
    limit = filters.get("limit")
    offset = filters.get("offset")
    if isinstance(limit, int) and limit > 0:
        sql += " LIMIT ?"
        params.append(limit)
        if isinstance(offset, int) and offset >= 0:
            sql += " OFFSET ?"
            params.append(offset)
    elif isinstance(offset, int) and offset > 0:
        sql += " LIMIT -1 OFFSET ?"
        params.append(offset)

    return sql, params

def search_vouchers(filters: dict):
    conn = get_conn()
    cur = conn.cursor()
    sql, params = _build_search_sql(filters)

    # Debug: log SQL and params so we can inspect why rows might be missing
    try:
        logger.debug("search_vouchers SQL: %s", sql)
        logger.debug("search_vouchers params: %r", params)
    except Exception:
        pass

    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()
    return rows

# ------------------ Users (auth & admin) ------------------
def get_user_by_username(u):
    conn = get_conn();
    cur = conn.cursor()
    cur.execute("SELECT id, username, role, password_hash, is_active, must_change_pwd FROM users WHERE username=?",
                (u,))
    row = cur.fetchone();
    conn.close();
    return row


def list_users():
    conn = get_conn();
    cur = conn.cursor()
    cur.execute("SELECT id, username, role, is_active, must_change_pwd FROM users ORDER BY role, username")
    rows = cur.fetchall();
    conn.close();
    return rows

def create_user(username, password, role="user", must_change_pwd=0, extra_fields=None):
    """
    Create a user. Accepts plaintext password (it will be hashed).
    Returns new user id.
    """
    extra_fields = extra_fields or {}
    username = (username or "").strip()
    if not username:
        raise ValueError("username required")
    # align allowed roles with DB CHECK
    if role not in {"admin", "sales assistant", "user", "technician"}:
        raise ValueError(f"role must be one of {{'admin','sales assistant','user','technician'}}")

    # Validate extra fields keys
    for k in extra_fields.keys():
        if k not in USER_UPDATABLE_COLUMNS:
            raise ValueError(f"unsupported field for create_user: {k}")

    conn = get_conn()
    try:
        cur = _begin_immediate_transaction(conn)

        # If creating admin, ensure no other admin exists
        if role == "admin":
            cur.execute("SELECT COUNT(1) FROM users WHERE role = ?", ("admin",))
            count_admins = cur.fetchone()[0]
            if count_admins > 0:
                raise ValueError("An admin user already exists; cannot create a second admin.")

        cols = ["username", "password_hash", "role", "must_change_pwd"]
        placeholders = ["?"] * len(cols)
        hashed = _hash_pwd(password)
        values = [username, hashed, role, int(bool(must_change_pwd))]

        for k in extra_fields:
            cols.append(k)
            placeholders.append("?")
            values.append(extra_fields[k])

        sql = f"INSERT INTO users ({', '.join(cols)}) VALUES ({', '.join(placeholders)})"
        cur.execute(sql, values)
        uid = cur.lastrowid
        conn.commit()
        return uid
    except Exception:
        logger.exception("Failed to create user")
        conn.rollback()
        raise
    finally:
        conn.close()


def update_user(user_id, **fields):
    """
    Update a user by id. Accepts keyword args for columns (whitelisted).
    Returns number of rows updated.
    """
    if not fields:
        raise ValueError("no fields to update")
    # Validate keys
    for k in fields.keys():
        if k not in USER_UPDATABLE_COLUMNS:
            raise ValueError(f"unsupported update column: {k}")
    conn = get_conn()
    try:
        cur = _begin_immediate_transaction(conn)

        # If role is changing to admin, ensure no other admin exists (excluding this user)
        if "role" in fields and fields["role"] == "admin":
            cur.execute("SELECT COUNT(1) FROM users WHERE role = ? AND id != ?", ("admin", user_id))
            count_admins = cur.fetchone()[0]
            if count_admins > 0:
                raise ValueError("Another admin already exists; cannot promote this user to admin.")

        set_clauses = []
        params = []
        for k, v in fields.items():
            set_clauses.append(f"{k} = ?")
            params.append(v)
        params.append(user_id)
        sql = f"UPDATE users SET {', '.join(set_clauses)} WHERE id = ?"
        cur.execute(sql, params)
        updated = cur.rowcount
        conn.commit()
        return updated
    except Exception:
        logger.exception("Failed to update user")
        conn.rollback()
        raise
    finally:
        conn.close()

def reset_password(user_id, new_pwd):
    conn = get_conn();
    cur = conn.cursor()
    cur.execute(
        "UPDATE users SET password_hash=?, must_change_pwd=0, updated_at=? WHERE id=?",
        (_hash_pwd(new_pwd), datetime.now().isoformat(sep=" ", timespec="seconds"), user_id)
    )
    conn.commit();
    conn.close()


def delete_user(user_id):
    conn = get_conn();
    cur = conn.cursor()
    cur.execute("SELECT role FROM users WHERE id=?", (user_id,))
    row = cur.fetchone()
    if row and row[0] == "admin":
        conn.close();
        raise ValueError("Cannot delete admin")
    cur.execute("DELETE FROM users WHERE id=?", (user_id,))
    conn.commit();
    conn.close()


# ------------------ Staff utilities ------------------
def _process_square_image(path_in, path_out, max_px=400):
    img = Image.open(path_in).convert("RGB")
    img = ImageOps.exif_transpose(img)
    size = min(img.width, img.height)
    left = (img.width - size) // 2;
    top = (img.height - size) // 2
    img = img.crop((left, top, left + size, top + size))
    if size > max_px:
        img = img.resize((max_px, max_px), Image.LANCZOS)
    img.save(path_out, format="JPEG", quality=92)


# ------------------ Commission utilities ------------------
BILL_RE_CS = re.compile(r"^CS-(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])/\d{4}$")
BILL_RE_INV = re.compile(r"^INV-(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])/\d{4}$")

def _resolve_staff_db_id(tech_id_opt: str, tech_name: str) -> int | None:
    """Find staffs.id by staff_id_opt (preferred) or by name."""
    conn = get_conn()
    cur = conn.cursor()
    dbid = None
    if tech_id_opt:
        cur.execute("SELECT id FROM staffs WHERE staff_id_opt=?", (tech_id_opt,))
        row = cur.fetchone()
        if row: dbid = row[0]
    if dbid is None and tech_name:
        cur.execute("SELECT id FROM staffs WHERE name=?", (tech_name,))
        row = cur.fetchone()
        if row: dbid = row[0]
    conn.close()
    return dbid

def _parse_bill(ref_bill: str) -> tuple[str | None, str | None]:
    """Return (bill_type, bill_no) if valid else (None, None)."""
    s = (ref_bill or "").strip().upper()
    if BILL_RE_CS.match(s):
        return "CS", s
    if BILL_RE_INV.match(s):
        return "INV", s
    return None, None

# ------------------ UI ------------------
FONT_FAMILY = "Segoe UI"
UI_FONT_SIZE = 14  # requested 14


def freeze_tree_columns(tree: ttk.Treeview):
    # Make widths constant and disable drag-resize
    style = ttk.Style(tree)
    style.layout("NoResize.Treeview.Heading", [
        ("Treeheading.cell", {"sticky": "nswe"}),
        ("Treeheading.border", {"sticky": "nswe", "children": [
            ("Treeheading.padding", {"sticky": "nswe", "children": [
                ("Treeheading.image", {"side": "right", "sticky": ""}),
                ("Treeheading.text", {"sticky": "we"}),
            ]})
        ]})
    ])
    tree.configure(style="NoResize.Treeview")

    def block_sep(event):
        # If the pointer is over a heading separator, block it.
        region = tree.identify_region(event.x, event.y)
        if region == "separator":
            return "break"

    tree.bind("<Button-1>", block_sep, add="+")
    # also prevent double-click autosize
    tree.bind("<Double-1>", lambda e: "break", add="+")


class LoginDialog(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Login")
        self.geometry("460x240")
        self.resizable(False, False)
        self.grab_set()

        frm = ctk.CTkFrame(self);
        frm.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(frm, text="Username:").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.e_user = ctk.CTkEntry(frm, width=220);
        self.e_user.grid(row=0, column=1, sticky="w", pady=(0, 6))
        ctk.CTkLabel(frm, text="Password:").grid(row=1, column=0, sticky="w", pady=(0, 6))
        self.e_pwd = ctk.CTkEntry(frm, width=220, show="•");
        self.e_pwd.grid(row=1, column=1, sticky="w", pady=(0, 6))
        self.var_show = tk.BooleanVar(value=False)
        chk = ctk.CTkCheckBox(frm, text="Show", variable=self.var_show, command=self._toggle);
        chk.grid(row=1, column=2, padx=(6, 0))
        btns = ctk.CTkFrame(frm);
        btns.grid(row=2, column=0, columnspan=3, sticky="e", pady=(12, 0))
        self.btn_login = ctk.CTkButton(btns, text="Login", command=self._do_login, width=120);
        self.btn_login.pack(side="right")
        self.result = None

        # optional niceties
        self.e_user.focus_set()
        self.bind("<Return>", lambda _e: self._do_login())

    def _toggle(self):
        self.e_pwd.configure(show="" if self.var_show.get() else "•")

    def _do_login(self):
        u = (self.e_user.get() or "").strip()
        p = (self.e_pwd.get() or "").strip()
        if not u or not p:
            messagebox.showerror("Login", "Please enter username and password.")
            return

        row = get_user_by_username(u)
        if not row:
            messagebox.showerror("Login", "Invalid credentials.")
            return

        uid, username, role, pwdhash, is_active, must_change = row
        if not is_active:
            messagebox.showerror("Login", "Account is disabled.")
            return
        if not _verify_pwd(p, pwdhash):
            messagebox.showerror("Login", "Invalid credentials.")
            return

        # Force password change flow
        if must_change:
            # Show a helpful note listing the password requirements
            req_text = (
                "Your account requires a password change.\n\n"
                "Password requirements:\n"
                "- At least 10 characters\n"
                "- At least one uppercase letter\n"
                "- At least one lowercase letter\n"
                "- At least one digit\n"
                "- At least one symbol (e.g. !@#$%^&*)\n\n"
                "You can cancel to abort login."
            )
            messagebox.showinfo("Change Password — Requirements", req_text, parent=self)

            # Loop to ask for new password until it validates or user cancels
            while True:
                new1 = simpledialog.askstring("Change Password", "Enter new password:", show="•", parent=self)
                if new1 is None:
                    # user cancelled
                    return
                new2 = simpledialog.askstring("Change Password", "Confirm new password:", show="•", parent=self)
                if new2 is None:
                    # user cancelled confirmation
                    return
                if new1 != new2:
                    messagebox.showerror("Password", "Passwords do not match.", parent=self)
                    continue

                # Use global validate_password_policy()
                err = validate_password_policy(new1)
                if err:
                    messagebox.showerror("Password", err, parent=self)
                    # loop again so user can correct
                    continue

                # All good — apply change and continue with restart flow
                try:
                    reset_password(uid, new1)  # clears must_change flag
                except Exception as e:
                    logger.exception("Failed to reset password during forced change", exc_info=e)
                    messagebox.showerror("Password", f"Failed to set new password: {e}", parent=self)
                    return

                messagebox.showinfo("Password Changed",
                                    "Your password has been changed.\nThe application will restart now.",
                                    parent=self)
                # Safe window teardown + restart flow
                try:
                    try:
                        self.grab_release()
                    except Exception:
                        logger.debug("grab_release failed (ignorable)")
                except Exception:
                    pass

                # Close main window safely (if it still exists)
                try:
                    if getattr(self, "master", None) is not None:
                        try:
                            if hasattr(self.master, "winfo_exists") and self.master.winfo_exists():
                                self.master.destroy()
                        except Exception:
                            # master already gone or cannot be destroyed; ignore
                            logger.debug("master destroy failed or already destroyed")
                except Exception:
                    logger.exception("Unexpected when destroying master during password change")

                # destroy the dialog itself (safe)
                try:
                    if hasattr(self, "winfo_exists") and self.winfo_exists():
                        self.destroy()
                except tk.TclError:
                    # application may be shutting down; ignore
                    logger.debug("dialog destroy raised TclError (ignorable)")
                except Exception:
                    logger.exception("Unexpected when destroying login dialog")

                # restart app (this will re-launch)
                restart_app()
                return

        # Success path: set result and close dialog safely
        self.result = {"id": uid, "username": username, "role": role}
        try:
            if hasattr(self, "winfo_exists") and self.winfo_exists():
                self.destroy()
        except Exception:
            logger.exception("Failed destroying login dialog on success (ignorable)")

def white_btn(parent, **kwargs):
    kwargs.setdefault("fg_color", "white")
    kwargs.setdefault("text_color", "black")
    kwargs.setdefault("border_color", "black")
    kwargs.setdefault("border_width", 1)
    kwargs.setdefault("hover_color", "#F0F0F0")
    return ctk.CTkButton(parent, **kwargs)

def make_xy_scroller(parent):
    """Return (outer_frame, inner_frame) where inner_frame is scrollable on both axes."""
    outer = ctk.CTkFrame(parent)
    # Canvas for scrolling
    canvas = tk.Canvas(outer, highlightthickness=0)
    xbar = ttk.Scrollbar(outer, orient="horizontal", command=canvas.xview)
    ybar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    canvas.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)

    # Layout
    canvas.grid(row=0, column=0, sticky="nsew")
    ybar.grid(row=0, column=1, sticky="ns")
    xbar.grid(row=1, column=0, sticky="ew")
    outer.grid_rowconfigure(0, weight=1)
    outer.grid_columnconfigure(0, weight=1)

    # The actual content frame goes inside the canvas
    inner = ctk.CTkFrame(canvas)
    win = canvas.create_window((0, 0), window=inner, anchor="nw")

    # Auto-resize scrollregion
    def _update_scrollregion(_=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    inner.bind("<Configure>", _update_scrollregion)

    # Resize canvas when outer changes
    def _resize_canvas(event):
        canvas.configure(width=event.width - ybar.winfo_width(),
                         height=event.height - xbar.winfo_height())
        # Keep inner pinned to the left while allowing horizontal scroll
        canvas.itemconfigure(win, anchor="nw")

    outer.bind("<Configure>", _resize_canvas)

    # Wheel scrolling
    def _on_mousewheel(e):
        if e.state & 0x0001:  # Shift = horizontal
            canvas.xview_scroll(-1 if e.delta > 0 else 1, "units")
        else:
            canvas.yview_scroll(-1 if e.delta > 0 else 1, "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel, add="+")  # Windows

    return outer, inner

STATUS_VALUES = ["All", "Pending", "Completed", "Deleted", "1st call", "2nd reminder", "3rd reminder"]

STATUS_ALLOWED = {"Pending", "Completed", "Deleted", "1st call", "2nd reminder", "3rd reminder"}

def _validate_status_or_raise(s: str):
    if s not in STATUS_ALLOWED:
        raise ValueError(f"Invalid status: {s}")

class VoucherApp(ctk.CTk):
    def _copy_selected_row(self, max_chars: int = 5000):
        """
        Copy currently selected row to clipboard. Defensive guards:
          - Ensure max_chars is an int (if UI accidentally passed a widget, we coerce/ignore).
          - Ensure self.tree exists and is a Treeview.
          - Handles very large rows by saving full content to a timestamped file and copying a truncated preview.
        """
        try:
            # Defensive: if caller mistakenly passed a Treeview as the first arg (legacy bug),
            # detect and ignore it by coercing to default max_chars.
            if not isinstance(max_chars, int):
                max_chars = 5000

            # Defensive: ensure tree exists and has selection method
            if not hasattr(self, "tree") or not hasattr(self.tree, "selection"):
                logger.warning("_copy_selected_row called but self.tree not available")
                try:
                    messagebox.showinfo("Copy", "No row selected.")
                except Exception:
                    pass
                return

            # Get selection
            try:
                sel = self.tree.selection()
            except Exception:
                sel = ()
            if not sel:
                try:
                    messagebox.showinfo("Copy", "No row selected.")
                except Exception:
                    logger.info("No row selected to copy.")
                return

            # Use the first selected item
            iid = sel[0]
            try:
                item = self.tree.item(iid)
                values = item.get("values", []) if isinstance(item, dict) else []
            except Exception:
                values = []

            # Clean values into a readable string (column headers not needed)
            # ensure all values coerced to plain strings
            row_text = "\t".join([str(v) for v in (values or [])])

            # If small enough, copy directly
            if len(row_text) <= max_chars:
                try:
                    self.clipboard_clear()
                    self.clipboard_append(row_text)
                    # On some platforms, update() helps flush the clipboard
                    try:
                        self.update()
                    except Exception:
                        pass
                    try:
                        messagebox.showinfo("Copy", "Row copied to clipboard.")
                    except Exception:
                        logger.info("Row copied to clipboard.")
                    return
                except Exception:
                    logger.exception("Failed to copy to clipboard; will try file fallback")

            # Large content path: truncate for clipboard, save full to temp file
            preview = row_text[: max(0, max_chars - 200)]  # leave room for notice
            notice = "\n\n[TRUNCATED] Full content saved to file."
            clipboard_text = preview + notice

            # Save full content to a temp file with timestamp
            try:
                ts = datetime.now().strftime("%Y%m%d%H%M%S")
                safe_dir = os.path.join(os.path.dirname(DB_FILE) if "DB_FILE" in globals() else os.getcwd(), "exports")
                os.makedirs(safe_dir, exist_ok=True)
                filename = os.path.join(safe_dir, f"voucher_row_full_{ts}.txt")
                with open(filename, "w", encoding="utf-8") as fh:
                    fh.write(row_text)

                # Write truncated preview to clipboard
                try:
                    self.clipboard_clear()
                    self.clipboard_append(clipboard_text)
                    try:
                        self.update()
                    except Exception:
                        pass
                except Exception:
                    logger.exception("Failed to place truncated text on clipboard")

                # Inform the user where full content is saved
                try:
                    messagebox.showinfo(
                        "Copy (truncated)",
                        f"Row content was too large — a preview was copied to the clipboard.\n\n"
                        f"The full content was saved to:\n{filename}"
                    )
                except Exception:
                    logger.info("Large row saved to %s", filename)

            except Exception:
                logger.exception("Failed to save large clipboard content to file", exc_info=True)
                try:
                    messagebox.showerror("Copy failed", "Failed to copy or save row content. See logs for details.")
                except Exception:
                    pass

        except Exception:
            # Top-level catch so user doesn't see a crash; log stack trace
            logger.exception("Unexpected error in _copy_selected_row", exc_info=True)
            try:    
                messagebox.showerror("Copy", "Unexpected error when copying row. Check logs.")
            except Exception:
                pass

    def __init__(self):
        super().__init__()
        self.title("Service Voucher Management System")
        self.geometry("1280x780")
        self.minsize(1024, 640)
        ctk.set_appearance_mode("light")

        # bump base Tk fonts to UI_FONT_SIZE
        try:
            import tkinter.font as tkfont
            for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkTooltipFont"):
                try:
                    tkfont.nametofont(name).configure(size=UI_FONT_SIZE)
                except Exception as e:
                    logger.exception("Caught exception", exc_info=e)
                    pass
        except Exception as e:
            logger.exception("Caught exception", exc_info=e)
            pass

        self.current_user = None
        self._do_login_flow()
        self._build_menubar()

        root = ctk.CTkFrame(self)
        root.pack(fill="both", expand=True)

        self._build_filters(root)
        self._build_table(root)
        self._build_bottom_bar(root)

        self.after(80, lambda: self._go_fullscreen())
        self.perform_search()

    def _do_login_flow(self):
        dlg = LoginDialog(self)
        self.wait_window(dlg)
        if not dlg.result:
            self.destroy()
            sys.exit(0)
        self.current_user = dlg.result
        self.title(f"SVMS — logged in as {self.current_user['username']} ({self.current_user['role']})")

    def _role_is(self, *roles):
        return bool(self.current_user and self.current_user["role"] in roles)

    def _build_menubar(self):
        self.menu = tk.Menu(self)
        self.config(menu=self.menu)

        self.menu_svms = tk.Menu(self.menu, tearoff=0)
        self.menu_svms.add_command(label="Open PDF (Selected)", command=self.open_pdf)
        self.menu_svms.add_command(label="Open PDF Folder", command=self.open_pdf_folder)
        self.menu_svms.add_command(label="Regenerate PDF (Selected)", command=self.regen_pdf_selected)
        self.menu_svms.add_separator()
        self.menu_svms.add_command(label="Backup Data (.zip)", command=self.backup_all)
        self.menu_svms.add_command(label="Restore Data (.zip)", command=self.restore_all)
        self.menu_svms.add_separator()
        self.menu_svms.add_command(label="Exit", command=self.destroy)
        self.menu.add_cascade(label="SVMS Menu", menu=self.menu_svms)

        self.menu_user = tk.Menu(self.menu, tearoff=0)
        self.menu_user.add_command(
            label="Manage Users",
            command=self.manage_users,
            state=tk.NORMAL if self._role_is("admin") else tk.DISABLED,
        )
        self.menu.add_cascade(label="User Profile", menu=self.menu_user)

        self.menu_staff = tk.Menu(self.menu, tearoff=0)
        self.menu_staff.add_command(label="Staff Profile", command=self.staff_profile)
        self.menu.add_cascade(label="Staff Profile", menu=self.menu_staff)

        # Sales Commission menu (now includes Service & Sales Report)
        self.menu_comm = tk.Menu(self.menu, tearoff=0)
        self.menu_comm.add_command(label="Add Commission", command=self.add_commission)
        self.menu_comm.add_command(label="View/Edit Commissions", command=self.view_commissions)

        # NEW: Service & Sales Report entry (defensive wiring)
        try:
            if hasattr(self, "service_sales_report_ui") and callable(getattr(self, "service_sales_report_ui")):
                self.menu_comm.add_separator()
                self.menu_comm.add_command(label="Service & Sales Report", command=self.service_sales_report_ui)
            else:
                # If the method isn't present, add a disabled menu item so users see the option but it does nothing.
                self.menu_comm.add_separator()
                self.menu_comm.add_command(label="Service & Sales Report (Not available)", state=tk.DISABLED)
        except Exception:
            # Fail-safe: if anything goes wrong adding the menu item, still show the other items.
            logger.exception("Failed adding Service & Sales Report menu item", exc_info=True)

        self.menu.add_cascade(label="Sales Commission", menu=self.menu_comm)

    def _build_filters(self, parent):
        wrap = ctk.CTkFrame(parent)
        wrap.pack(fill="x", padx=8, pady=(8, 6))

        self.filter_canvas = tk.Canvas(wrap, height=52, borderwidth=0, highlightthickness=0)
        hscroll = ttk.Scrollbar(wrap, orient="horizontal", command=self.filter_canvas.xview)
        self.filter_canvas.configure(xscrollcommand=hscroll.set)
        self.filter_canvas.pack(fill="x", side="top")
        hscroll.pack(fill="x", side="bottom")

        self.filter_inner = ctk.CTkFrame(self.filter_canvas)
        self.filter_canvas.create_window((0, 0), window=self.filter_inner, anchor="nw")

        today_ui = _to_ui_date(datetime.now())

        self.f_voucher = ctk.CTkEntry(self.filter_inner, width=140, placeholder_text="VoucherID")
        self.f_voucher.grid(row=0, column=0, padx=5, pady=4)

        self.f_name = ctk.CTkEntry(self.filter_inner, width=230, placeholder_text="Customer Name")
        self.f_name.grid(row=0, column=1, padx=5, pady=4)

        self.f_contact = ctk.CTkEntry(self.filter_inner, width=190, placeholder_text="Contact Number")
        self.f_contact.grid(row=0, column=2, padx=5, pady=4)

        self.f_from = ctk.CTkEntry(self.filter_inner, width=180, placeholder_text="Date From (DD-MM-YYYY)")
        self.f_from.grid(row=0, column=3, padx=5, pady=4)

        self.f_to = ctk.CTkEntry(self.filter_inner, width=180, placeholder_text="Date To (DD-MM-YYYY)")
        self.f_to.grid(row=0, column=4, padx=5, pady=4)
        self.f_to.insert(0, today_ui)

        self.f_status = ctk.CTkOptionMenu(self.filter_inner, values=STATUS_VALUES, width=140)
        self.f_status.grid(row=0, column=5, padx=5, pady=4)
        self.f_status.set("All")

        self.btn_search = white_btn(self.filter_inner, text="Search", command=self.perform_search, width=110)
        self.btn_search.grid(row=0, column=6, padx=5, pady=4)

        self.btn_reset = white_btn(self.filter_inner, text="Reset", command=self.reset_filters, width=100)
        self.btn_reset.grid(row=0, column=7, padx=5, pady=4)

        self.filter_inner.update_idletasks()
        self.filter_canvas.configure(scrollregion=self.filter_canvas.bbox("all"))
        self.filter_inner.bind(
            "<Configure>", lambda _e=None: self.filter_canvas.configure(scrollregion=self.filter_canvas.bbox("all"))
        )

    def _build_table(self, parent):
        table_frame = ctk.CTkFrame(parent)
        table_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            table_frame,
            columns=(
                "VoucherID",
                "Date",
                "Customer",
                "Contact",
                "Units",
                "Recipient",
                "TechID",
                "TechName",
                "Status",
                "Solution",
                "PDF",
            ),
            show="headings",
            selectmode="extended",
        )
        self.tree.grid(row=0, column=0, sticky="nsew")

        vbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")

        for col, text, w in [
            ("VoucherID", "VoucherID", 100),
            ("Date", "Date", 220),
            ("Customer", "Customer", 220),
            ("Contact", "Contact", 160),
            ("Units", "Units", 70),
            ("Recipient", "Recipient", 180),
            ("TechID", "Technician ID", 130),
            ("TechName", "Technician Name", 200),
            ("Status", "Status", 130),
            ("Solution", "Solution", 360),
        ]:
            self.tree.heading(col, text=text)
            self.tree.column(col, anchor="w", width=w, stretch=False)
        self.tree.heading("PDF", text="PDF")
        self.tree.column("PDF", width=0, stretch=False)

        self.tree.tag_configure("Pending", background="#FFF4B3")
        self.tree.tag_configure("Completed", background="#CDEEC8")
        self.tree.tag_configure("Deleted", background="#F8D7DA")
        self.tree.tag_configure("1st call", background="#CCE5FF")
        self.tree.tag_configure("2nd reminder", background="#D9CCE5")
        self.tree.tag_configure("3rd reminder", background="#FFD9B3")
        self.tree.tag_configure("out_7", background="#FFF0F0")
        self.tree.tag_configure("out_14", background="#FFD6D6")
        self.tree.tag_configure("out_30", background="#FFB3B3")

        self.tree.bind("<Double-1>", lambda e: self.open_pdf())
        self._make_context_menu(self.tree)
        # fix: refer to self.tree instead of an undefined 'tree'
        self.ctx.add_command(
            label="Copy row (tab-separated)", command=lambda: self._copy_selected_row(self.tree)
        )
        freeze_tree_columns(self.tree)

    def _build_bottom_bar(self, parent):
        bar = ctk.CTkFrame(parent)
        bar.pack(fill="x", padx=8, pady=(0, 10))
        white_btn(bar, text="Add Voucher", command=self.add_voucher_ui, width=140).pack(side="left", padx=5, pady=8)
        white_btn(bar, text="Open PDF", command=self.open_pdf, width=120).pack(side="left", padx=5, pady=8)
        white_btn(bar, text="Open PDF Folder", command=self.open_pdf_folder, width=150).pack(side="left", padx=5,
                                                                                             pady=8)
        b_manage_rec = white_btn(bar, text="Manage Recipients", command=self.manage_staffs_ui, width=170)
        b_manage_rec.pack(side="left", padx=5, pady=8)
        if not self._role_is("admin", "supervisor"):
            b_manage_rec.configure(state=tk.DISABLED)
        b_base = white_btn(bar, text="Modify Base VID", command=self.modify_base_vid_ui, width=160)
        b_base.pack(side="left", padx=5, pady=8)
        if not self._role_is("admin"):
            b_base.configure(state=tk.DISABLED)

    def _make_context_menu(self, tree):
        """
        Safe context menu builder for a treeview.
        - Ensures self.ctx is always set before return.
        - Registers 'Copy row' correctly (no accidental argument passing).
        - Cross-platform: binds Button-3, Button-2 and Control-Button-1 (mac).
        """
        try:
            self.ctx = tk.Menu(tree, tearoff=0)
            # Basic actions
            self.ctx.add_command(label="Copy row (tab-separated)", command=self._copy_selected_row)
            self.ctx.add_command(label="Edit", command=self.edit_selected)
            # Mark submenu
            mark = tk.Menu(self.ctx, tearoff=0)
            for s in ["Pending", "Completed", "Deleted", "1st call", "2nd reminder", "3rd reminder"]:
                # use default arg capture to avoid late-binding trap
                mark.add_command(label=s, command=(lambda st=s: self._bulk_mark(st)))
            self.ctx.add_cascade(label="Mark", menu=mark)
            self.ctx.add_separator()
            self.ctx.add_command(label="Open PDF", command=self.open_pdf)
            self.ctx.add_command(label="Regenerate PDF", command=self.regen_pdf_selected)
            self.ctx.add_separator()
            self.ctx.add_command(label="Unmark Completed → Pending", command=(lambda: self._bulk_mark("Pending")))
        except Exception:
            # Make sure self.ctx exists even if menu building fails
            try:
                self.ctx = tk.Menu(tree, tearoff=0)
            except Exception:
                self.ctx = None
            logger.exception("Failed building context menu", exc_info=True)

        def _popup(event):
            try:
                # Identify row and select if not already selected
                try:
                    row = tree.identify_row(event.y)
                    if row and row not in tree.selection():
                        tree.selection_set(row)
                except Exception:
                    pass
                # Show menu if available
                if getattr(self, "ctx", None):
                    try:
                        self.ctx.post(event.x_root, event.y_root)
                    except Exception:
                        try:
                            # sometimes post can fail if widget destroyed; ignore
                            pass
                        except Exception:
                            pass
            finally:
                try:
                    if getattr(self, "ctx", None):
                        self.ctx.grab_release()
                except Exception:
                    pass

        # Bind right-clicks in a cross-platform way:
        try:
            tree.bind("<Button-3>", _popup, add="+")
        except Exception:
            pass
        # Some X11 setups use Button-2 for popup
        try:
            tree.bind("<Button-2>", _popup, add="+")
        except Exception:
            pass
        # On macOS, Control-Button-1 is often used
        try:
            tree.bind("<Control-Button-1>", _popup, add="+")
        except Exception:
            pass

    def _get_filters(self):
        """
        Return a filters dict compatible with _build_search_sql:
          - voucher_id
          - customer_name
          - contact_number
          - date_from (UI DD-MM-YYYY)  <-- _build_search_sql will call _parse_ui_date_to_iso
          - date_to
          - status
        """
        return {
            "voucher_id": (self.f_voucher.get() or "").strip(),
            "customer_name": (self.f_name.get() or "").strip(),
            # pass raw contact text (search will use LIKE). If you want normalized matching,
            # use normalize_phone(...) but be aware normalized stored values may differ.
            "contact_number": (self.f_contact.get() or "").strip(),
            "date_from": (self.f_from.get() or "").strip(),
            "date_to": (self.f_to.get() or "").strip(),
            "status": (self.f_status.get() or "").strip(),
        }

    def reset_filters(self):
        for e in (self.f_voucher, self.f_name, self.f_contact, self.f_from, self.f_to):
            e.delete(0, "end")
        self.f_to.insert(0, _to_ui_date(datetime.now()))
        self.f_status.set("All")
        self.perform_search()

    def perform_search(self):
        rows = search_vouchers(self._get_filters())
        self.tree.delete(*self.tree.get_children())
        now = datetime.now()

        for row in rows:
            # row expected: (voucher_id, created_at, customer_name, contact_number, units,
            #                recipient, technician_id, technician_name, status, solution, pdf_path)
            # be defensive about length and None values
            (vid, created_at, customer, contact, units, recipient,
             tech_id, tech_name, status, solution, pdf) = tuple(list(row) + [""] * (11 - len(row)))

            status = status or ""
            customer = customer or ""
            contact = contact or ""
            units = units or 1
            recipient = recipient or ""
            tech_id = tech_id or ""
            tech_name = tech_name or ""
            solution = solution or ""
            pdf = pdf or ""

            tags = [status] if status else []

            try:
                dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
            except Exception:
                dt = now
            days = (now - dt).days
            if status not in ("Completed", "Deleted") and days > 60:
                if days <= 67:
                    tags.append("out_7")
                elif days <= 74:
                    tags.append("out_14")
                else:
                    tags.append("out_30")

            self.tree.insert(
                "",
                "end",
                values=(
                    vid,
                    _to_ui_datetime_str(created_at),
                    customer,
                    contact,
                    units,
                    recipient,
                    tech_id,
                    tech_name,
                    status,
                    solution,
                    pdf,
                ),
                tags=tuple(tags),
            )
        self.tree.update_idletasks()

    def _selected_ids(self):
        sels = self.tree.selection()
        vids = []
        for iid in sels:
            vals = self.tree.item(iid)["values"]
            if vals:
                vids.append(str(vals[0]))
        return vids

    def _bulk_mark(self, new_status):
        vids = self._selected_ids()
        if not vids:
            messagebox.showerror("Mark", "Select record(s) first.")
            return
        if new_status == "Deleted" and not self._role_is("admin", "supervisor"):
            messagebox.showerror("Mark", "You do not have permission to mark as Deleted.")
            return
        bulk_update_status(vids, new_status)
        self.perform_search()

    def open_pdf(self):
        sels = self.tree.selection()
        if not sels:
            messagebox.showerror("Open PDF", "Select a record first.")
            return
        vals = self.tree.item(sels[0])["values"]
        if not vals:
            return
        pdf_path = vals[-1]
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Open PDF", "PDF not found for this voucher.")
            return
        try:
            webbrowser.open_new(os.path.abspath(pdf_path))
        except Exception:
            if os.name == "nt":
                os.startfile(pdf_path)  # type: ignore
            else:
                os.system(f"open '{pdf_path}'" if sys.platform == "darwin" else f"xdg-open '{pdf_path}'")

    def regen_pdf_selected(self):
        sels = self.tree.selection()
        if not sels:
            messagebox.showerror("Regenerate", "Select a record first.")
            return
        vid = str(self.tree.item(sels[0])["values"][0])

        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            """SELECT voucher_id, created_at, customer_name, contact_number, units,
                              particulars, problem, staff_name, status, recipient, solution, pdf_path
                       FROM vouchers WHERE voucher_id=?""",
            (vid,),
        )
        row = cur.fetchone()
        conn.close()
        if not row:
            messagebox.showerror("Regenerate", "Voucher not found.")
            return
        (
            voucher_id,
            created_at,
            customer_name,
            contact_number,
            units,
            particulars,
            problem,
            staff_name,
            status,
            recipient,
            solution,
            old_pdf,
        ) = row
        try:
            if old_pdf and os.path.exists(old_pdf):
                os.remove(old_pdf)
        except Exception as e:
            logger.exception("Caught exception", exc_info=e)
            pass
        new_pdf = generate_pdf(
            voucher_id,
            customer_name,
            contact_number,
            units,
            particulars,
            problem,
            staff_name,
            status,
            created_at,
            recipient,
        )
        update_voucher_fields(voucher_id, pdf_path=new_pdf)
        self.perform_search()
        messagebox.showinfo("Regenerate", f"PDF regenerated for voucher {voucher_id}.")

    def open_pdf_folder(self):
        open_path(PDF_DIR)

    # ---------- Add / Edit voucher ----------
    def add_voucher_ui(self):
        """Create Voucher window — updated layout and commission-binding flow."""
        # single-instance guard
        self._open_windows = getattr(self, "_open_windows", {})
        if self._open_windows.get("add_voucher_ui"):
            try:
                win = self._open_windows["add_voucher_ui"]
                if win.winfo_exists():
                    win.deiconify()
                    win.lift()
                    win.focus_force()
                    return
            except Exception:
                self._open_windows.pop("add_voucher_ui", None)

        top = ctk.CTkToplevel(self)
        top.title("Create Voucher")
        top.geometry("980x780")
        top.resizable(False, False)
        top.grab_set()

        # register window
        self._open_windows["add_voucher_ui"] = top

        def _on_close():
            try:
                del self._open_windows["add_voucher_ui"]
            except Exception as e:
                logger.exception("Caught exception", exc_info=e)
                pass
            try:
                top.destroy()
            except Exception as e:
                logger.exception("Caught exception", exc_info=e)
                pass

        try:
            top.protocol("WM_DELETE_WINDOW", _on_close)
        except Exception as e:
            logger.exception("Caught exception", exc_info=e)
            pass

        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=12, pady=12)
        frm.grid_columnconfigure(1, weight=1)

        # Make fields shorter per your request
        WIDE = 420   # shortened from previous 560
        ENTRY_SHORT = 240
        SMALL = 120
        r = 0

        def mk_text(parent, height, seed=""):
            wrap = ctk.CTkFrame(parent)
            wrap.grid_columnconfigure(0, weight=1)
            txt = tk.Text(wrap, height=height, font=(FONT_FAMILY, UI_FONT_SIZE), wrap='word')
            txt.insert("1.0", seed)
            sb = ttk.Scrollbar(wrap, orient="vertical", command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            txt.grid(row=0, column=0, sticky="nsew")
            sb.grid(row=0, column=1, sticky="ns")
            return wrap, txt

        # --- Customer (shorter)
        ctk.CTkLabel(frm, text="Customer Name").grid(row=r, column=0, sticky="w")
        e_name = ctk.CTkEntry(frm, width=WIDE)
        e_name.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="Contact Number").grid(row=r, column=0, sticky="w")
        e_contact = ctk.CTkEntry(frm, width=WIDE)
        e_contact.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="No. of Units").grid(row=r, column=0, sticky="w")
        e_units = ctk.CTkEntry(frm, width=SMALL)
        e_units.insert(0, "1")
        e_units.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        # Particulars (shorter)
        ctk.CTkLabel(frm, text="Particulars").grid(row=r, column=0, sticky="nw")
        part_wrap, t_part = mk_text(frm, height=3)
        part_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6)
        r += 1

        # Problem (shorter)
        ctk.CTkLabel(frm, text="Problem").grid(row=r, column=0, sticky="nw")
        prob_wrap, t_prob = mk_text(frm, height=3)
        prob_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6)
        r += 1

        # Recipient (shorter)
        ctk.CTkLabel(frm, text="Recipient").grid(row=r, column=0, sticky="w")
        try:
            staff_values = list_staffs_names()
        except Exception:
            staff_values = []
        staff_values = (["— Select —"] + staff_values) if staff_values else ["— Select —"]
        e_recipient = ctk.CTkComboBox(frm, values=staff_values, width=ENTRY_SHORT)
        e_recipient.set(staff_values[0])
        e_recipient.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        # ---------------- Technician ID and Technician Name separated vertically ----------------
        tech_id_values = ["— Select —"]
        tech_name_values = ["— Select —"]
        id_to_name = {}
        name_to_id = {}
        try:
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("SELECT staff_id_opt, name FROM staffs ORDER BY name COLLATE NOCASE")
            tech_rows = cur.fetchall()
            conn.close()
            for sid_opt, nm in tech_rows:
                if sid_opt and sid_opt.strip():
                    tech_id_values.append(sid_opt)
                    id_to_name[sid_opt] = nm
                tech_name_values.append(nm)
                name_to_id[nm] = sid_opt or ""
        except Exception as e:
            logger.exception("Caught exception", exc_info=e)
            pass

        # Technician ID (row)
        ctk.CTkLabel(frm, text="Technician ID").grid(row=r, column=0, sticky="w")
        cb_tech_id = ctk.CTkComboBox(frm, values=tech_id_values, width=200)
        cb_tech_id.set(tech_id_values[0])
        cb_tech_id.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        # Technician Name (placed under ID)
        ctk.CTkLabel(frm, text="Technician Name").grid(row=r, column=0, sticky="w")
        cb_tech_name = ctk.CTkComboBox(frm, values=tech_name_values, width=300)
        cb_tech_name.set(tech_name_values[0])
        cb_tech_name.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        def _on_tech_id_change(new=None):
            sel = cb_tech_id.get().strip()
            if not sel or sel == "— Select —":
                cb_tech_name.set("— Select —")
                return
            nm = id_to_name.get(sel, "")
            if nm:
                cb_tech_name.set(nm)

        def _on_tech_name_change(new=None):
            sel = cb_tech_name.get().strip()
            if not sel or sel == "— Select —":
                cb_tech_id.set("— Select —")
                return
            sid = name_to_id.get(sel, "")
            if sid:
                cb_tech_id.set(sid)
            else:
                cb_tech_id.set("— Select —")

        try:
            cb_tech_id.configure(command=_on_tech_id_change)
        except Exception:
            cb_tech_id.bind("<<ComboboxSelected>>", lambda e: _on_tech_id_change())
        try:
            cb_tech_name.configure(command=_on_tech_name_change)
        except Exception:
            cb_tech_name.bind("<<ComboboxSelected>>", lambda e: _on_tech_name_change())

        # ---------------- Bill Type & Commission Picker (replace bill no textbox) ----------------
        ctk.CTkLabel(frm, text="Bill Type").grid(row=r, column=0, sticky="w")
        bill_type_cb = ctk.CTkComboBox(frm, values=["CS", "INV", "None"], width=120)
        bill_type_cb.set("CS")
        bill_type_cb.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        # Commission picker: no free-text for bill no anymore. User must choose a commission record (or None).
        ctk.CTkLabel(frm, text="Commission (choose)").grid(row=r, column=0, sticky="w")
        # label to display chosen commission summary
        lbl_chosen_comm = ctk.CTkLabel(frm, text="No commission chosen", width=ENTRY_SHORT, anchor="w")
        lbl_chosen_comm.grid(row=r, column=1, sticky="w", padx=10, pady=6)

        def _open_comm_picker():
            """Open a small dialog listing unbound commissions for user to choose."""
            pick = ctk.CTkToplevel(top)
            pick.title("Choose Commission")
            # Bigger, fixed-size dialog so all columns are visible.
            pick.geometry("1400x520")
            pick.resizable(False, False)
            pick.grab_set()
            
            wrap = ctk.CTkFrame(pick)
            wrap.pack(fill="both", expand=True, padx=8, pady=8)
            # search/filter
            fl = ctk.CTkFrame(wrap)
            fl.pack(fill="x", padx=6, pady=(0,6))
            e_q = ctk.CTkEntry(fl, placeholder_text="Search staff or bill no", width=400)
            e_q.pack(side="left", padx=(0,8))
            b_search = white_btn(fl, text="Filter", width=100)
            b_search.pack(side="left")

            # tree
            cols = ("ID","StaffID","StaffName","BillType","BillNo","Total","Commission","Created")
            tree = ttk.Treeview(wrap, columns=cols, show="headings", selectmode="browse")
            for ccol, title, wcol in [
                ("ID","ID",60), ("StaffID","StaffID",120), ("StaffName","Staff",220),
                ("BillType","Type",90), ("BillNo","Bill No",260), ("Total","Total",120),
                ("Commission","Commission",200), ("Created","Created",140)
            ]:
                tree.heading(ccol, text=title)
                tree.column(ccol, width=wcol, anchor="w", stretch=False)
            tree.pack(fill="both", expand=True, padx=6, pady=(0,6))

            # populate
            def _load(q=""):
                qlow = (q or "").strip().lower()
                conn = get_conn()
                cur = conn.cursor()
                cur.execute("""
                    SELECT c.id, s.staff_id_opt, s.name, c.bill_type, c.bill_no, c.total_amount, c.commission_amount, c.created_at
                    FROM commissions c
                    JOIN staffs s ON c.staff_id = s.id
                    WHERE (c.voucher_id IS NULL OR c.voucher_id = '')
                    ORDER BY c.id DESC
                """)
                rows = cur.fetchall()
                conn.close()
                tree.delete(*tree.get_children())
                for (cid, staff_opt, staff_name, bt, bno, tot, comm, created_at) in rows:
                    if qlow and qlow not in str(staff_name).lower() and qlow not in str(bno).lower():
                        continue
                    tree.insert("", "end", iid=str(cid), values=(cid, staff_opt or "", staff_name, bt, bno, tot or "", comm or "", created_at or ""))
            _load()

            def _pick_selected():
                sel = tree.selection()
                if not sel:
                    messagebox.showerror("Choose", "Select a commission first.", parent=pick)
                    return
                cid = int(sel[0])
                vals = tree.item(sel[0])["values"]
                # store into outer scope
                chosen = {
                    "id": cid,
                    "bill_type": vals[3],
                    "bill_no": vals[4],
                    "total": vals[5],
                    "commission": vals[6],
                }
                # set UI label in add-voucher form
                lbl_chosen_comm.configure(text=f"{chosen['bill_type']} {chosen['bill_no']} — Total:{chosen['total']} Commission:{chosen['commission']}")
                # save chosen into outer variable
                nonlocal_chosen.clear()
                nonlocal_chosen.update(chosen)
                pick.destroy()

            def _do_filter():
                _load(e_q.get())

            b_search.configure(command=_do_filter)

            # double click = pick
            tree.bind("<Double-1>", lambda e: _pick_selected())

            btns = ctk.CTkFrame(pick)
            btns.pack(fill="x", padx=8, pady=(4,8))
            white_btn(btns, text="Select", command=_pick_selected, width=120).pack(side="right", padx=(6,0))
            white_btn(btns, text="Close", command=pick.destroy, width=120).pack(side="right", padx=(0,6))

        # Actual storage for chosen commission (dict) — use a mutable so nested func can modify
        nonlocal_chosen = {}
        # Choose button
        white_btn(frm, text="Choose Commission...", command=_open_comm_picker, width=160).grid(row=r, column=1, sticky="e", padx=10, pady=6)
        r += 1

        # Bill Date (optional) — keep for record but user won't manually enter bill no now
        ctk.CTkLabel(frm, text="Bill Date (DD-MM-YYYY)").grid(row=r, column=0, sticky="w")
        e_ref_bill_date = ctk.CTkEntry(frm, width=180)
        e_ref_bill_date.insert(0, _to_ui_date(datetime.now()))
        e_ref_bill_date.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        # Behavior: when bill type is "None", disable bill no entry
        def _on_bill_type_change(*_a):
            """
            Safely handle changes in bill type dropdown.
            Enables/disables reference bill fields, and auto-prefixes new bill numbers.
            Guards against cases where e_ref_bill or e_ref_bill_date may not exist.
            """
            try:
                # Safely try to access widgets if they exist in this scope
                try:
                    ref_entry = e_ref_bill
                except NameError:
                    ref_entry = None
                try:
                    ref_date_entry = e_ref_bill_date
                except NameError:
                    ref_date_entry = None

                # Get selected bill type
                bt = (bill_type_cb.get() or "").strip().upper()

                # If widgets not present, log once and return
                if ref_entry is None or ref_date_entry is None:
                    logger.debug("Skipped _on_bill_type_change: ref_entry or ref_date_entry missing")
                    return

                if bt == "NONE":
                    # Disable both widgets
                    try:
                        ref_entry.delete(0, "end")
                        ref_entry.configure(state="disabled")
                        ref_date_entry.configure(state="disabled")
                    except Exception as e:
                        logger.exception("Caught exception disabling ref bill fields", exc_info=e)
                        pass
                else:
                    # Enable both and auto-prefix if needed
                    try:
                        ref_entry.configure(state="normal")
                        ref_date_entry.configure(state="normal")
                        today = datetime.now()
                        mm = f"{today.month:02d}"
                        dd = f"{today.day:02d}"
                        prefix = f"{bt}-{mm}{dd}/"
                        curv = (ref_entry.get() or "").strip().upper()
                        # Auto-prefix if blank or old prefix format
                        if not curv or re.match(r"^(CS|INV)-\d{4}/", curv, re.IGNORECASE):
                            ref_entry.delete(0, "end")
                            ref_entry.insert(0, prefix)
                    except Exception as e:
                        logger.exception("Caught exception enabling/prefixing ref bill fields", exc_info=e)
                        pass

            except Exception as outer_e:
                logger.exception("Caught outer exception in _on_bill_type_change", exc_info=outer_e)
                pass

        # Bind the combobox to the change handler
        try:
            bill_type_cb.configure(command=_on_bill_type_change)
        except Exception:
            bill_type_cb.bind("<<ComboboxSelected>>", lambda e: _on_bill_type_change())

        # Run once at start to initialize field states
        _on_bill_type_change()

        def save():
            name = e_name.get().strip()
            contact = normalize_phone(e_contact.get())
            try:
                units = int((e_units.get() or "1").strip())
                if units <= 0:
                    raise ValueError
            except Exception:
                messagebox.showerror("Invalid", "Units must be a positive integer.")
                return
            particulars = t_part.get("1.0", "end").strip()
            problem = t_prob.get("1.0", "end").strip()
            recipient = e_recipient.get().strip()
            if recipient in ("— Select —", ""):
                messagebox.showerror("Missing", "Please choose a Recipient.")
                return

            # For Add UI we don't have a solution text area; default to empty
            solution = ""

            # Technician resolution
            tech_id_sel = cb_tech_id.get().strip()
            tech_name_sel = cb_tech_name.get().strip()
            if tech_id_sel in ("— Select —", "") and tech_name_sel in ("— Select —", ""):
                messagebox.showerror("Missing", "Please select a Technician in charge.")
                return
            technician_id = tech_id_sel if tech_id_sel not in ("— Select —", "") else ""
            technician_name = tech_name_sel if tech_name_sel not in ("— Select —", "") else ""

            if not name or not contact:
                messagebox.showerror("Missing", "Customer name and contact are required.")
                return

            # use chosen commission if any (nonlocal_chosen), otherwise create without binding
            chosen_comm = nonlocal_chosen.get("id")

            # For Add UI we set status to Pending
            status_val = "Pending"

            # Determine ref_bill and amounts if we selected a commission
            ref_bill_val = nonlocal_chosen.get("bill_no") if nonlocal_chosen else None
            ref_bill_date_sql = _from_ui_date_to_sqldate(e_ref_bill_date.get().strip()) if e_ref_bill_date else None
            amount_val = None
            tech_comm_val = None
            if nonlocal_chosen:
                try:
                    amount_val = float(nonlocal_chosen.get("total")) if nonlocal_chosen.get("total") not in (None, "") else None
                except Exception:
                    amount_val = None
                try:
                    tech_comm_val = float(nonlocal_chosen.get("commission")) if nonlocal_chosen.get("commission") not in (None, "") else None
                except Exception:
                    tech_comm_val = None

            try:
                voucher_id, _pdf = add_voucher(
                    customer_name=name,
                    contact_number=contact,
                    units=units,
                    particulars=particulars,
                    problem=problem,
                    staff_name=recipient,
                    recipient=recipient,
                    solution=solution,
                    technician_id=technician_id,
                    technician_name=technician_name,
                    status=status_val,
                    ref_bill=ref_bill_val,
                    ref_bill_date=ref_bill_date_sql,
                    amount_rm=amount_val,
                    tech_commission=tech_comm_val,
                )
            except Exception as ex:
                messagebox.showerror("Save Failed", f"Failed to create voucher:\n{ex}")
                return

            # Bind commission if a chosen commission existed
            if chosen_comm:
                try:
                    bind_commission_to_voucher(chosen_comm, voucher_id)
                    # optionally write commission details into voucher (safe-update)
                    try:
                        conn = get_conn()
                        cur = conn.cursor()
                        cur.execute("PRAGMA table_info(vouchers)")
                        vcols = [r[1] for r in cur.fetchall()]
                        updates = []
                        params = []
                        if "ref_bill" in vcols and ref_bill_val:
                            updates.append("ref_bill=?"); params.append(ref_bill_val)
                        if "amount_rm" in vcols and amount_val is not None:
                            updates.append("amount_rm=?"); params.append(amount_val)
                        if "tech_commission" in vcols and tech_comm_val is not None:
                            updates.append("tech_commission=?"); params.append(tech_comm_val)
                        if updates:
                            params.append(voucher_id)
                            cur.execute(f"UPDATE vouchers SET {', '.join(updates)} WHERE voucher_id=?", params)
                            conn.commit()
                    except Exception:
                        logger.exception("Failed writing commission details into voucher after binding", exc_info=True)
                    finally:
                        try:
                            conn.close()
                        except Exception:
                            pass
                except Exception as e:
                    logger.exception("Failed binding commission after voucher creation", exc_info=e)
                    messagebox.showwarning("Bound Failed", f"Voucher created ({voucher_id}) but failed to bind commission: {e}")

            messagebox.showinfo("Saved", f"Voucher {voucher_id} created.")
            try:
                _on_close()
            except Exception:
                pass

            # --- Ensure UI shows the newly created voucher:
            # Reset filter inputs (clears date / name / contact filters) then refresh table.
            try:
                # Reset filters so newly created voucher is not hidden by an active filter
                try:
                    self.reset_filters()
                except Exception:
                    # Fallback: perform a direct refresh if reset_filters not available
                    try:
                        self.perform_search()
                    except Exception:
                        pass
            except Exception:
                logger.exception("Failed to reset filters / refresh after voucher create", exc_info=True)

        # Add Save / Cancel buttons (were missing previously)
        btns = ctk.CTkFrame(top)
        btns.pack(fill="x", padx=12, pady=(6, 12))
        
        white_btn(btns, text="Save", command=save, width=140).pack(side="right")
        white_btn(btns, text="Cancel", command=lambda: _on_close(), width=100).pack(side="right", padx=(0,8))

    def edit_selected(self):
        sels = self.tree.selection()
        if len(sels) != 1:
            messagebox.showerror("Edit", "Select exactly one record to edit.")
            return
        voucher_id = str(self.tree.item(sels[0])["values"][0])

        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            """SELECT voucher_id, created_at, customer_name, contact_number, units,
              particulars, problem, staff_name, status, recipient, solution,
              technician_id, technician_name,
              ref_bill, ref_bill_date, amount_rm, tech_commission,
              reminder_pickup_1, reminder_pickup_2, reminder_pickup_3
       FROM vouchers WHERE voucher_id=?""",
            (voucher_id,),
        )
        row = cur.fetchone()
        conn.close()
        if not row:
            messagebox.showerror("Edit", "Voucher not found.")
            return

        (
            _,
            created_at,
            customer_name,
            contact_number,
            units,
            particulars,
            problem,
            staff_name,
            status,
            recipient,
            solution,
            tech_id,
            tech_name,
            ref_bill,
            ref_bill_date,
            amount_rm,
            tech_commission,
            reminder1,
            reminder2,
            reminder3,
        ) = row

        top = ctk.CTkToplevel(self)
        top.title(f"Edit Voucher {voucher_id}")
        top.geometry("980x780")
        top.grab_set()
        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=12, pady=12)
        frm.grid_columnconfigure(1, weight=1)

        WIDE = 560
        r = 0

        ctk.CTkLabel(frm, text="Customer Name").grid(row=r, column=0, sticky="w")
        e_name = ctk.CTkEntry(frm, width=WIDE)
        e_name.insert(0, customer_name or "")
        e_name.grid(row=r, column=1, sticky="ew", padx=10, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="Contact Number").grid(row=r, column=0, sticky="w")
        e_contact = ctk.CTkEntry(frm, width=WIDE)
        e_contact.insert(0, contact_number or "")
        e_contact.grid(row=r, column=1, sticky="ew", padx=10, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="No. of Units").grid(row=r, column=0, sticky="w")
        e_units = ctk.CTkEntry(frm, width=120)
        e_units.insert(0, str(units or 1))
        e_units.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        def mk_text(parent, height, seed=""):
            wrap = ctk.CTkFrame(parent)
            wrap.grid_columnconfigure(0, weight=1)
            txt = tk.Text(wrap, height=height, font=(FONT_FAMILY, UI_FONT_SIZE), wrap='word')
            txt.insert("1.0", seed or "")
            sb = ttk.Scrollbar(wrap, orient="vertical", command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            txt.grid(row=0, column=0, sticky="nsew")
            sb.grid(row=0, column=1, sticky="ns")
            return wrap, txt

        ctk.CTkLabel(frm, text="Particulars").grid(row=r, column=0, sticky="nw")
        part_wrap, t_part = mk_text(frm, height=3, seed=(particulars or ""))
        part_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="Problem").grid(row=r, column=0, sticky="nw")
        prob_wrap, t_prob = mk_text(frm, height=3, seed=(problem or ""))
        prob_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="Recipient").grid(row=r, column=0, sticky="w")
        try:
            staff_values = list_staffs_names()
        except Exception:
            staff_values = []
        staff_values = (["— Select —"] + staff_values) if staff_values else ["— Select —"]
        e_recipient = ctk.CTkComboBox(frm, values=staff_values, width=240)
        e_recipient.set(recipient or (staff_values[0] if staff_values else "— Select —"))
        e_recipient.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        # Technician ID and Name (separate, linked)
        tech_id_values = ["— Select —"]
        tech_name_values = ["— Select —"]
        id_to_name = {}
        name_to_id = {}
        try:
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("SELECT staff_id_opt, name FROM staffs ORDER BY name COLLATE NOCASE")
            tech_rows = cur.fetchall()
            conn.close()
            for sid_opt, nm in tech_rows:
                if sid_opt and sid_opt.strip():
                    tech_id_values.append(sid_opt)
                    id_to_name[sid_opt] = nm
                tech_name_values.append(nm)
                name_to_id[nm] = sid_opt or ""
        except Exception as e:
            logger.exception("Caught exception", exc_info=e)
            pass

        ctk.CTkLabel(frm, text="Technician ID").grid(row=r, column=0, sticky="w")
        cb_tech_id = ctk.CTkComboBox(frm, values=tech_id_values, width=200)
        cb_tech_id.set(tech_id or tech_id_values[0])
        cb_tech_id.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="Technician Name").grid(row=r, column=0, sticky="w")
        cb_tech_name = ctk.CTkComboBox(frm, values=tech_name_values, width=300)
        cb_tech_name.set(tech_name or tech_name_values[0])
        cb_tech_name.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        def _on_tech_id_change(new=None):
            sel = cb_tech_id.get().strip()
            if not sel or sel == "— Select —":
                cb_tech_name.set("— Select —")
                return
            nm = id_to_name.get(sel, "")
            if nm:
                cb_tech_name.set(nm)

        def _on_tech_name_change(new=None):
            sel = cb_tech_name.get().strip()
            if not sel or sel == "— Select —":
                cb_tech_id.set("— Select —")
                return
            sid = name_to_id.get(sel, "")
            if sid:
                cb_tech_id.set(sid)
            else:
                cb_tech_id.set("— Select —")

        try:
            cb_tech_id.configure(command=_on_tech_id_change)
        except Exception:
            cb_tech_id.bind("<<ComboboxSelected>>", lambda e: _on_tech_id_change())
        try:
            cb_tech_name.configure(command=_on_tech_name_change)
        except Exception:
            cb_tech_name.bind("<<ComboboxSelected>>", lambda e: _on_tech_name_change())

        # ---------------- Bill Type & Commission picker for Edit ----------------
        ctk.CTkLabel(frm, text="Bill Type").grid(row=r, column=0, sticky="w")
        # display same options as add page
        bill_type_cb = ctk.CTkComboBox(frm, values=["CS", "INV", "None"], width=120)
        # if ref_bill present, derive type from prefix, otherwise keep existing selection
        if ref_bill and isinstance(ref_bill, str) and ref_bill.upper().startswith("INV-"):
            bill_type_cb.set("INV")
        elif ref_bill and isinstance(ref_bill, str) and ref_bill.upper().startswith("CS-"):
            bill_type_cb.set("CS")
        else:
            bill_type_cb.set("CS" if (status or "").lower() != "none" else "None")
        bill_type_cb.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="Commission (choose)").grid(row=r, column=0, sticky="w")
        lbl_chosen_comm = ctk.CTkLabel(frm, text="No commission chosen", width=240, anchor="w")
        lbl_chosen_comm.grid(row=r, column=1, sticky="w", padx=10, pady=6)

        # Try to detect existing bound commission for this voucher (if any)
        existing_comm = None
        try:
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("SELECT id, bill_type, bill_no, total_amount, commission_amount FROM commissions WHERE voucher_id=?", (voucher_id,))
            existing_comm = cur.fetchone()
            # if not found, attempt to match by ref_bill
            if not existing_comm and ref_bill:
                cur.execute("SELECT id, bill_type, bill_no, total_amount, commission_amount FROM commissions WHERE LOWER(bill_no)=LOWER(?)", (ref_bill,))
                existing_comm = cur.fetchone()
            conn.close()
        except Exception:
            existing_comm = None

        nonlocal_chosen = {}
        if existing_comm:
            # existing_comm = (id, bill_type, bill_no, total, commission)
            nonlocal_chosen.update({
                "id": existing_comm[0],
                "bill_type": existing_comm[1],
                "bill_no": existing_comm[2],
                "total": existing_comm[3],
                "commission": existing_comm[4],
            })
            lbl_chosen_comm.configure(text=f"{nonlocal_chosen['bill_type']} {nonlocal_chosen['bill_no']} — Total:{nonlocal_chosen['total']} Commission:{nonlocal_chosen['commission']}")

        def _open_comm_picker():
            """Open commission picker dialog (similar to Add)."""
            pick = ctk.CTkToplevel(top)
            pick.title("Choose Commission")
            # Bigger, fixed-size dialog so all columns are visible.
            pick.geometry("1400x520")
            pick.resizable(False, False)
            pick.grab_set()

            wrap = ctk.CTkFrame(pick)
            wrap.pack(fill="both", expand=True, padx=8, pady=8)
            fl = ctk.CTkFrame(wrap)
            fl.pack(fill="x", padx=6, pady=(0,6))
            e_q = ctk.CTkEntry(fl, placeholder_text="Search staff or bill no", width=400)
            e_q.pack(side="left", padx=(0,8))
            b_search = white_btn(fl, text="Filter", width=100)
            b_search.pack(side="left")

            cols = ("ID","StaffID","StaffName","BillType","BillNo","Total","Commission","Created")
            tree = ttk.Treeview(wrap, columns=cols, show="headings", selectmode="browse")
            for ccol, title, wcol in [
                ("ID","ID",60), ("StaffID","StaffID",120), ("StaffName","Staff",220),
                ("BillType","Type",90), ("BillNo","Bill No",260), ("Total","Total",120),
                ("Commission","Commission",200), ("Created","Created",140)
            ]:
                tree.heading(ccol, text=title)
                tree.column(ccol, width=wcol, anchor="w", stretch=False)
            tree.pack(fill="both", expand=True, padx=6, pady=(0,6))

            def _load(q=""):
                qlow = (q or "").strip().lower()
                conn = get_conn()
                cur = conn.cursor()
                cur.execute("""
                    SELECT c.id, s.staff_id_opt, s.name, c.bill_type, c.bill_no, c.total_amount, c.commission_amount, c.created_at
                    FROM commissions c
                    JOIN staffs s ON c.staff_id = s.id
                    WHERE (c.voucher_id IS NULL OR c.voucher_id = '')
                    ORDER BY c.id DESC
                """)
                rows = cur.fetchall()
                conn.close()
                tree.delete(*tree.get_children())
                for (cid, staff_opt, staff_name, bt, bno, tot, comm, created_at) in rows:
                    if qlow and qlow not in str(staff_name).lower() and qlow not in str(bno).lower():
                        continue
                    tree.insert("", "end", iid=str(cid), values=(cid, staff_opt or "", staff_name, bt, bno, tot or "", comm or "", created_at or ""))

            _load()

            def _pick_selected():
                sel = tree.selection()
                if not sel:
                    messagebox.showerror("Choose", "Select a commission first.", parent=pick)
                    return
                cid = int(sel[0])
                vals = tree.item(sel[0])["values"]
                chosen = {
                    "id": cid,
                    "bill_type": vals[3],
                    "bill_no": vals[4],
                    "total": vals[5],
                    "commission": vals[6],
                }
                lbl_chosen_comm.configure(text=f"{chosen['bill_type']} {chosen['bill_no']} — Total:{chosen['total']} Commission:{chosen['commission']}")
                nonlocal_chosen.clear()
                nonlocal_chosen.update(chosen)
                pick.destroy()

            def _do_filter():
                _load(e_q.get())

            b_search.configure(command=_do_filter)
            tree.bind("<Double-1>", lambda e: _pick_selected())

            btns = ctk.CTkFrame(pick)
            btns.pack(fill="x", padx=8, pady=(4,8))
            white_btn(btns, text="Select", command=_pick_selected, width=120).pack(side="right", padx=(6,0))
            white_btn(btns, text="Close", command=pick.destroy, width=120).pack(side="right", padx=(0,6))

        white_btn(frm, text="Choose Commission...", command=_open_comm_picker, width=160).grid(row=r, column=1, sticky="e", padx=10, pady=6)
        r += 1

        # Bill Date (optional)
        ctk.CTkLabel(frm, text="Bill Date (DD-MM-YYYY)").grid(row=r, column=0, sticky="w")
        e_ref_bill_date = ctk.CTkEntry(frm, width=180)
        if ref_bill_date:
            try:
                if isinstance(ref_bill_date, str):
                    dt = datetime.strptime(ref_bill_date, "%Y-%m-%d")
                    e_ref_bill_date.insert(0, _to_ui_date(dt))
                else:
                    e_ref_bill_date.insert(0, str(ref_bill_date))
            except Exception:
                # fallback: show raw (or blank)
                e_ref_bill_date.insert(0, str(ref_bill_date or ""))
        else:
            e_ref_bill_date.insert(0, _to_ui_date(datetime.now()))
        e_ref_bill_date.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        # Amount and Commission Amount (prefill from voucher if present)
        ctk.CTkLabel(frm, text="Amount (Total)").grid(row=r, column=0, sticky="w")
        e_amount = ctk.CTkEntry(frm, width=180)
        if amount_rm is not None:
            e_amount.insert(0, str(amount_rm))
        e_amount.grid(row=r, column=1, sticky="w", padx=(10,0), pady=6)

        ctk.CTkLabel(frm, text="Commission Amount").grid(row=r, column=1, sticky="w", padx=(220,0))
        e_tech_comm = ctk.CTkEntry(frm, width=140)
        if tech_commission is not None:
            e_tech_comm.insert(0, str(tech_commission))
        e_tech_comm.grid(row=r, column=1, sticky="w", padx=(340,10), pady=6)
        r += 1

        # Status
        ctk.CTkLabel(frm, text="Status").grid(row=r, column=0, sticky="w")
        status_choices = ["Pending", "Completed", "Deleted", "1st call", "2nd reminder", "3rd reminder"]
        cb_status = ctk.CTkComboBox(frm, values=status_choices, width=200)
        cb_status.set(status or "Pending")
        cb_status.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        # Solution
        ctk.CTkLabel(frm, text="Solution").grid(row=r, column=0, sticky="nw")
        sol_wrap, t_sol = mk_text(frm, height=3, seed=(solution or ""))
        sol_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6)
        r += 1

        # Buttons
        btns = ctk.CTkFrame(top)
        btns.pack(fill="x", padx=12, pady=(6, 12))

        def save_edit():
            name = e_name.get().strip()
            contact = normalize_phone(e_contact.get())
            try:
                units_val = int((e_units.get() or "1").strip())
                if units_val <= 0:
                    raise ValueError
            except Exception:
                messagebox.showerror("Invalid", "Units must be a positive integer.")
                return
            particulars_val = t_part.get("1.0", "end").strip()
            problem_val = t_prob.get("1.0", "end").strip()
            recipient_val = e_recipient.get().strip()
            if recipient_val in ("— Select —", ""):
                messagebox.showerror("Missing", "Please choose a Recipient.")
                return
            solution_val = t_sol.get("1.0", "end").strip()

            # Technician resolution
            tech_id_sel = cb_tech_id.get().strip()
            tech_name_sel = cb_tech_name.get().strip()
            if tech_id_sel in ("— Select —", "") and tech_name_sel in ("— Select —", ""):
                messagebox.showerror("Missing", "Please select a Technician in charge.")
                return
            technician_id_val = tech_id_sel if tech_id_sel not in ("— Select —", "") else ""
            technician_name_val = tech_name_sel if tech_name_sel not in ("— Select —", "") else ""

            if not name or not contact:
                messagebox.showerror("Missing", "Customer name and contact are required.")
                return

            # amounts
            amount_val = None
            tech_comm_val = None
            if (e_amount.get() or "").strip():
                try:
                    amount_val = float(e_amount.get().strip())
                except Exception:
                    messagebox.showerror("Invalid", "Amount must be a number.")
                    return
            if (e_tech_comm.get() or "").strip():
                try:
                    tech_comm_val = float(e_tech_comm.get().strip())
                except Exception:
                    messagebox.showerror("Invalid", "Commission amount must be a number.")
                    return

            # Determine ref_bill & ref_bill_date from chosen commission if any
            ref_bill_val = None
            if nonlocal_chosen.get("id"):
                ref_bill_val = nonlocal_chosen.get("bill_no")
            else:
                # user did not pick commission: keep existing ref_bill if any, but do not allow creating a new ref_bill by typing
                ref_bill_val = ref_bill or None

            ref_bill_date_sql = _from_ui_date_to_sqldate(e_ref_bill_date.get().strip())

            # Prepare fields to update
            fields = {
                "customer_name": name,
                "contact_number": contact,
                "units": units_val,
                "particulars": particulars_val,
                "problem": problem_val,
                "staff_name": recipient_val,
                "status": cb_status.get().strip(),
                "recipient": recipient_val,
                "solution": solution_val,
                "technician_id": technician_id_val,
                "technician_name": technician_name_val,
            }
            # optional columns may or may not exist; update_voucher_fields will only write valid columns
            if ref_bill_val is not None:
                fields["ref_bill"] = ref_bill_val
            if ref_bill_date_sql:
                fields["ref_bill_date"] = ref_bill_date_sql
            if amount_val is not None:
                fields["amount_rm"] = amount_val
            if tech_comm_val is not None:
                fields["tech_commission"] = tech_comm_val

            # Attempt to update voucher
            try:
                update_voucher_fields(voucher_id, **fields)
            except Exception as ex:
                messagebox.showerror("Save Failed", f"Failed to update voucher:\n{ex}")
                return

            # If nonlocal_chosen present, attempt to bind; otherwise leave existing binding intact
            if nonlocal_chosen.get("id"):
                chosen_id = nonlocal_chosen["id"]
                try:
                    # Unbind any previous commission that referenced this voucher (so uniqueness holds)
                    try:
                        conn = get_conn()
                        cur = conn.cursor()
                        cur.execute("PRAGMA table_info(commissions)")
                        if any(r[1] == "voucher_id" for r in cur.fetchall()):
                            cur.execute("UPDATE commissions SET voucher_id=NULL, updated_at=? WHERE voucher_id=?", (datetime.now().isoformat(sep=' ', timespec='seconds'), voucher_id))
                            conn.commit()
                    except Exception as e:
                        logger.exception("Caught exception", exc_info=e)
                        pass
                    finally:
                        try:
                            conn.close()
                        except Exception as e:
                            logger.exception("Caught exception", exc_info=e)
                            pass

                    # Now bind chosen commission to this voucher
                    bind_commission_to_voucher(chosen_id, voucher_id)

                    # After binding, also ensure voucher contains commission data (bill_no/amounts)
                    try:
                        conn = get_conn()
                        cur = conn.cursor()
                        cur.execute("PRAGMA table_info(vouchers)")
                        vcols = [r[1] for r in cur.fetchall()]
                        updates = []
                        params = []
                        if "ref_bill" in vcols and nonlocal_chosen.get("bill_no"):
                            updates.append("ref_bill=?"); params.append(nonlocal_chosen.get("bill_no"))
                        if "amount_rm" in vcols and nonlocal_chosen.get("total") is not None:
                            updates.append("amount_rm=?"); params.append(nonlocal_chosen.get("total"))
                        if "tech_commission" in vcols and nonlocal_chosen.get("commission") is not None:
                            updates.append("tech_commission=?"); params.append(nonlocal_chosen.get("commission"))
                        if updates:
                            params.append(voucher_id)
                            cur.execute(f"UPDATE vouchers SET {', '.join(updates)} WHERE voucher_id=?", params)
                            conn.commit()
                    except Exception:
                        logger.exception("Failed writing commission details into voucher after binding", exc_info=True)
                    finally:
                        try:
                            conn.close()
                        except Exception as e:
                            logger.exception("Caught exception", exc_info=e)
                            pass

                except Exception as e:
                    logger.exception("Failed to bind commission during edit", exc_info=e)
                    messagebox.showwarning("Bind Failed", f"Voucher updated but failed to bind commission: {e}")

            messagebox.showinfo("Saved", f"Voucher {voucher_id} updated.")
            try:
                top.destroy()
            except Exception as e:
                logger.exception("Caught exception", exc_info=e)
                pass
            try:
                self.perform_search()
            except Exception as e:
                logger.exception("Caught exception", exc_info=e)
                pass

        white_btn(btns, text="Save", command=save_edit, width=140).pack(side="right")
        white_btn(btns, text="Cancel", command=lambda: top.destroy(), width=100).pack(side="right", padx=(0,8))

        try:
            top.update_idletasks()
        except Exception as e:
            logger.exception("Caught exception", exc_info=e)
            pass

    # ---------- Manage Recipient (simple) ----------
    def manage_staffs_ui(self):
        if not self._role_is("admin", "supervisor"):
            messagebox.showerror("Permission", "Not allowed.")
            return

        top = ctk.CTkToplevel(self)
        top.title("Manage Recipients")
        top.geometry("720x520")
        top.grab_set()

        root = ctk.CTkFrame(top)
        root.pack(fill="both", expand=True, padx=14, pady=14)
        root.grid_rowconfigure(2, weight=1)
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=0)

        entry = ctk.CTkEntry(root, placeholder_text="New recipient name")
        entry.grid(row=0, column=0, sticky="ew", padx=(0, 10), pady=(0, 10))

        row1 = ctk.CTkFrame(root)
        row1.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(0, 10))
        add_btn = white_btn(row1, text="Add", width=120)
        del_btn = white_btn(row1, text="Delete Selected", width=160)
        add_btn.pack(side="left", padx=(0, 10))
        del_btn.pack(side="left")

        list_frame = ctk.CTkFrame(root)
        list_frame.grid(row=2, column=0, sticky="nsew", padx=(0, 10))
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        lb = tk.Listbox(list_frame, height=14, font=(FONT_FAMILY, UI_FONT_SIZE))
        lb.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=lb.yview)
        sb.grid(row=0, column=1, sticky="ns")
        lb.configure(yscrollcommand=sb.set)

        close_btn = white_btn(root, text="Close", width=120, command=top.destroy)
        close_btn.grid(row=3, column=1, sticky="e", pady=(10, 0))

        def refresh_list():
            lb.delete(0, "end")
            for s in list_staffs_names():
                lb.insert("end", s)

        def do_add():
            if add_staff_simple(entry.get()):
                entry.delete(0, "end")
                refresh_list()

        def do_del():
            sel = lb.curselection()
            if not sel:
                return
            name = lb.get(sel[0])
            delete_staff_simple(name)
            refresh_list()

        add_btn.configure(command=do_add)
        del_btn.configure(command=do_del)
        refresh_list()

    # ---------- Modify Base VID ----------
    def modify_base_vid_ui(self):
        if not self._role_is("admin"):
            messagebox.showerror("Permission", "Admin only.")
            return

        top = ctk.CTkToplevel(self)
        top.title("Modify Base Voucher ID")
        top.geometry("420x220")
        top.grab_set()
        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=16, pady=16)
        current_base = _read_base_vid()
        ctk.CTkLabel(frm, text=f"Current Base VID: {current_base}", anchor="w").pack(fill="x", pady=(0, 8))
        entry = ctk.CTkEntry(frm, placeholder_text="Enter new base voucher ID (integer)")
        entry.pack(fill="x")
        entry.insert(0, str(current_base))
        info = ctk.CTkLabel(
            frm,
            text="Existing vouchers will shift by the difference. PDFs will be regenerated.",
            justify="left",
            wraplength=360,
        )
        info.pack(fill="x", pady=8)

        def apply():
            new_base_str = entry.get().strip()
            if not new_base_str.isdigit():
                messagebox.showerror("Invalid", "Please enter a positive integer.")
                return
            new_base = int(new_base_str)
            try:
                delta = modify_base_vid(new_base)
                messagebox.showinfo("Done", f"Base VID set to {new_base}. Shift: {delta:+d}.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to modify base VID:\n{e}")
            top.destroy()
            self.perform_search()

        btns = ctk.CTkFrame(frm)
        btns.pack(fill="x", pady=(8, 0))
        white_btn(btns, text="Apply", command=apply, width=120).pack(side="right", padx=4)
        white_btn(btns, text="Cancel", command=top.destroy, width=100).pack(side="right", padx=4)

    # ---------- Backup / Restore ----------
    def backup_all(self):
        # Admin-only guard: only allow users with 'admin' role to perform this action
        try:
            if not getattr(self, '_role_is', lambda r: False)('admin'):
                import tkinter.messagebox as messagebox
                messagebox.showerror('Permission', 'Admin only.')
                return
        except Exception as e:
            logger.exception("Caught exception", exc_info=e)
            pass

        # Admin-only guard (UI + safety)
        if not self._role_is("admin"):
            messagebox.showerror("Permission", "Admin only.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("Zip Archive", "*.zip")], title="Save backup as")
        if not path:
            return
        try:
            with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
                for root, _, files in os.walk(APP_DIR):
                    for f in files:
                        fp = os.path.join(root, f)
                        arc = os.path.relpath(fp, APP_DIR)
                        z.write(fp, arcname=arc)
                for root, _, files in os.walk(IMAGES_DIR):
                    for f in files:
                        fp = os.path.join(root, f)
                        arc = os.path.relpath(fp, APP_DIR)
                        z.write(fp, arcname=arc)
            messagebox.showinfo("Backup", f"Backup created:\n{path}")
        except Exception as e:
            messagebox.showerror("Backup Error", f"Failed to create backup:\n{e}")

    def restore_all(self):
        # Admin-only guard: only allow users with 'admin' role to perform this action
        try:
            if not getattr(self, '_role_is', lambda r: False)('admin'):
                import tkinter.messagebox as messagebox
                messagebox.showerror('Permission', 'Admin only.')
                return
        except Exception as e:
            logger.exception("Caught exception", exc_info=e)
            pass

        # Admin-only guard (UI + safety)
        if not self._role_is("admin"):
            messagebox.showerror("Permission", "Admin only.")
            return
        path = filedialog.askopenfilename(filetypes=[("Zip Archive", "*.zip")], title="Select Backup Zip")
        if not path:
            return
        if not messagebox.askyesno("Restore", "Restoring will overwrite current data. Continue?"):
            return
        try:
            with zipfile.ZipFile(path, "r") as z:
                temp_dir = os.path.join(APP_DIR, "_restore_tmp")
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                z.extractall(temp_dir)
                src_db = os.path.join(temp_dir, "vouchers.db")
                if os.path.exists(src_db):
                    shutil.copy2(src_db, DB_FILE)
                for sub in ("pdfs", "staffs", "images"):
                    src = os.path.join(temp_dir, sub)
                    if os.path.exists(src):
                        dst = os.path.join(APP_DIR, sub)
                        if os.path.exists(dst):
                            shutil.rmtree(dst)
                        shutil.copytree(src, dst)
                shutil.rmtree(temp_dir, ignore_errors=True)
            messagebox.showinfo("Restore", "Restore completed. The application will now restart.")
            restart_app()
        except Exception as e:
            messagebox.showerror("Restore Error", f"Failed to restore backup:\n{e}")
            
    # ---------- Users (Admin-only UI) ----------
    def manage_users(self):
        """User Accounts manager (admin-only).
        Replaced earlier version that referenced missing helpers.
        This self-contained implementation uses the DB helpers already in the file:
        list_users, create_user, update_user, delete_user, reset_password.
        """
        # Admin-only guard
        if not self._role_is("admin"):
            messagebox.showerror("Permission", "Admin only.")
            return

        # single-instance guard
        self._open_windows = getattr(self, "_open_windows", {})
        if self._open_windows.get("manage_users"):
            try:
                w = self._open_windows["manage_users"]
                if w.winfo_exists():
                    w.deiconify(); w.lift(); w.focus_force()
                    return
            except Exception:
                self._open_windows.pop("manage_users", None)

        top = ctk.CTkToplevel(self)
        top.title("User Accounts")
        TOP_W, TOP_H = 1200, 700
        top.geometry(f"{TOP_W}x{TOP_H}")
        top.resizable(False, False)
        top.grab_set()
        self._open_windows["manage_users"] = top
        top.protocol("WM_DELETE_WINDOW", lambda: (self._open_windows.pop("manage_users", None), top.destroy()))

        # --- Toolbar ---
        toolbar = ctk.CTkFrame(top)
        toolbar.pack(fill="x", padx=8, pady=(8, 6))

        # left side: search
        left_toolbar = ctk.CTkFrame(toolbar)
        left_toolbar.pack(side="left", fill="x", expand=True)

        search_var = tk.StringVar()
        ctk.CTkLabel(left_toolbar, text="Search:").pack(side="left", padx=(8,4))
        e_search = ctk.CTkEntry(left_toolbar, textvariable=search_var, width=320)
        e_search.pack(side="left", padx=(0,8))

        # right side: buttons container so buttons stay grouped to the right
        right_toolbar = ctk.CTkFrame(toolbar)
        right_toolbar.pack(side="right", padx=(0,8))

        # small helper to refresh tree (defined before buttons so we can reference it)
        def do_refresh():
            try:
                refresh_users()
            except Exception as e:
                # if refresh_users not available yet, ignore silently
                logger.exception("Caught exception", exc_info=e)
                pass

        try:
            e_search.bind("<Return>", lambda e: do_refresh())
        except Exception:
            pass

        # Button handlers implemented inline (use DB helper functions)
        def on_add():
            show_add_edit_dialog(mode="add")

        def on_edit():
            # tree will be defined later; lookup happens at call time
            try:
                sel = tree.selection()
            except Exception:
                messagebox.showinfo("Edit user", "Select a user row first.")
                return
            if not sel:
                messagebox.showinfo("Edit user", "Select a user row first.")
                return
            vals = tree.item(sel[0])["values"]
            uid = vals[0]
            show_add_edit_dialog(mode="edit", user_id=uid)

        def on_delete():
            try:
                sel = tree.selection()
            except Exception:
                messagebox.showinfo("Delete user", "Select a user row first.")
                return
            if not sel:
                messagebox.showinfo("Delete user", "Select a user row first.")
                return
            vals = tree.item(sel[0])["values"]
            uid = vals[0]
            uname = vals[1]
            if messagebox.askyesno("Delete", f"Delete user '{uname}'?"):
                try:
                    delete_user(uid)
                except Exception as e:
                    messagebox.showerror("Delete user", f"Delete failed:\n{e}")
            do_refresh()

        def on_reset():
            try:
                sel = tree.selection()
            except Exception:
                messagebox.showinfo("Reset Password", "Select a user row first.")
                return
            if not sel:
                messagebox.showinfo("Reset Password", "Select a user row first.")
                return
            vals = tree.item(sel[0])["values"]
            uid = vals[0]
            uname = vals[1]
            newpwd = simpledialog.askstring("Reset Password", f"Enter new password for {uname}:", parent=top, show="•")
            if not newpwd:
                return
            try:
                reset_password(uid, newpwd)
            except Exception as e:
                messagebox.showerror("Reset Password", f"Failed:\n{e}")
            else:
                messagebox.showinfo("Reset Password", "Password reset successfully.")

        def on_toggle_active():
            try:
                sel = tree.selection()
            except Exception:
                messagebox.showinfo("Toggle Active", "Select a user row first.")
                return
            if not sel:
                messagebox.showinfo("Toggle Active", "Select a user row first.")
                return
            vals = tree.item(sel[0])["values"]
            uid = vals[0]
            try:
                # read current is_active from DB (list_users returns that info)
                cur_active = None
                for u in list_users():
                    if u[0] == uid:
                        cur_active = u[3]
                        break
                if cur_active is None:
                    messagebox.showerror("Toggle Active", "User record not found.")
                    return
                update_user(uid, is_active=0 if cur_active else 1)
            except Exception as e:
                messagebox.showerror("Toggle Active", f"Failed:\n{e}")
            do_refresh()

        # toolbar buttons (grouped on right)
        white_btn(right_toolbar, text="Add", command=on_add, width=100).pack(side="left", padx=(0,6))
        white_btn(right_toolbar, text="Edit", command=on_edit, width=100).pack(side="left", padx=(0,6))
        white_btn(right_toolbar, text="Delete", command=on_delete, width=100).pack(side="left", padx=(0,6))
        white_btn(right_toolbar, text="Reset Password", command=on_reset, width=140).pack(side="left", padx=(0,6))
        white_btn(right_toolbar, text="Toggle Active", command=on_toggle_active, width=140).pack(side="left", padx=(0,6))
        white_btn(right_toolbar, text="Refresh", command=do_refresh, width=100).pack(side="left", padx=(8,0))

        # --- Main area: treeview + scrollbar ---
        mainf = ctk.CTkFrame(top)
        mainf.pack(fill="both", expand=True, padx=8, pady=(0,8))
        mainf.grid_rowconfigure(0, weight=1)
        mainf.grid_columnconfigure(0, weight=1)

        container = tk.Frame(mainf)
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        cols = ("id", "username", "role", "active", "must_change")
        display_names = {"id":"ID","username":"Username","role":"Role","active":"Active","must_change":"Must Change"}

        tree = ttk.Treeview(container, columns=cols, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        for c in cols:
            tree.heading(c, text=display_names.get(c, c.title()))

        def autosize_columns():
            try:
                container.update_idletasks()
                padding = 20
                total_w = container.winfo_width() - padding
                if total_w <= 0:
                    total_w = TOP_W - 40
                samples = {cid: [str(tree.heading(cid)['text'])] for cid in cols}
                for item in tree.get_children():
                    vals = tree.item(item)["values"]
                    for idx, cid in enumerate(cols):
                        try:
                            samples[cid].append(str(vals[idx] or ""))
                        except Exception as e:
                            logger.exception("Caught exception", exc_info=e)
                            pass
                min_widths = []
                for cid in cols:
                    longest = max((len(s) for s in samples[cid]), default=6)
                    min_widths.append(max(60, int(longest * 7 + 20)))
                sum_min = sum(min_widths)
                if sum_min <= total_w:
                    extra = total_w - sum_min
                    widths = [mw + int(extra * (mw / sum_min)) for mw in min_widths]
                else:
                    factor = total_w / sum_min
                    widths = [max(60, int(mw * factor)) for mw in min_widths]
                for cid, w in zip(cols, widths):
                    tree.column(cid, width=w, anchor="w", stretch=False)
            except Exception:
                for cid in cols:
                    tree.column(cid, width=150, anchor="w", stretch=True)

        def refresh_users():
            try:
                tree.delete(*tree.get_children())
                q = (search_var.get() or "").strip().lower()
                for (uid, username, role, is_active, must_change) in list_users():
                    display = (uid, username, role, "Yes" if is_active else "No", "Yes" if must_change else "No")
                    if not q or q in str(username).lower() or q in str(role).lower():
                        tree.insert("", "end", values=display)
            except Exception as e:
                logger.exception("Caught exception", exc_info=e)
                pass
            autosize_columns()

        # initial population and bind resizing
        refresh_users()
        try:
            top.bind("<Configure>", lambda e: autosize_columns())
        except Exception as e:
            logger.exception("Caught exception", exc_info=e)
            pass

        # double-click to edit
        try:
            tree.bind("<Double-1>", lambda e: on_edit())
        except Exception as e:
            logger.exception("Caught exception", exc_info=e)
            pass

        # --- Add / Edit dialog implementation (self-contained) ---
        def show_add_edit_dialog(mode="add", user_id=None):
            # mode: "add" or "edit"
            dlg = ctk.CTkToplevel(top)
            dlg.transient(top)
            dlg.grab_set()
            dlg.title("Add User" if mode=="add" else f"Edit User {user_id}")
            dlg.geometry("420x260")

            frm = ctk.CTkFrame(dlg)
            frm.pack(fill="both", expand=True, padx=12, pady=12)
            frm.grid_columnconfigure(1, weight=1)

            ctk.CTkLabel(frm, text="Username").grid(row=0, column=0, sticky="w", pady=(6,6))
            e_un = ctk.CTkEntry(frm, width=240)
            e_un.grid(row=0, column=1, sticky="ew", pady=(6,6))

            ctk.CTkLabel(frm, text="Role").grid(row=1, column=0, sticky="w", pady=(6,6))
            role_vals = ["admin", "sales assistant", "technician", "user"]
            cb_role = ctk.CTkComboBox(frm, values=role_vals, width=180)
            cb_role.set(role_vals[-1])
            cb_role.grid(row=1, column=1, sticky="w", pady=(6,6))

            ctk.CTkLabel(frm, text="Password").grid(row=2, column=0, sticky="w", pady=(6,6))
            e_pwd = ctk.CTkEntry(frm, width=240, show="•")
            e_pwd.grid(row=2, column=1, sticky="ew", pady=(6,6))

            chk_must_change = tk.IntVar(value=0)
            ctk.CTkCheckBox(frm, text="Must change password on next login", variable=chk_must_change).grid(row=3, column=1, sticky="w", pady=(6,6))

            # If edit: populate fields from DB
            if mode == "edit" and user_id is not None:
                for u in list_users():
                    if u[0] == user_id:
                        e_un.insert(0, u[1])
                        cb_role.set(u[2] or "user")
                        # do not pre-fill password
                        break

            btns = ctk.CTkFrame(frm)
            btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(12,0))

            def do_save():
                uname = e_un.get().strip()
                role = cb_role.get().strip()
                pwd = e_pwd.get().strip()
                must = 1 if chk_must_change.get() else 0
                if not uname:
                    messagebox.showerror("Missing", "Username required.", parent=dlg)
                    return
                if mode == "add":
                    if not pwd:
                        messagebox.showerror("Missing", "Password required for new user.", parent=dlg)
                        return
                    try:
                        # create_user(username, password, role=..., must_change_pwd=...)
                        create_user(uname, pwd, role=role, must_change_pwd=1 if must else 0)
                        messagebox.showinfo("Added", "User created.", parent=dlg)
                    except Exception as e:
                        messagebox.showerror("Add failed", f"{e}", parent=dlg)
                        return
                else:
                    # edit mode: update role and must_change (password change optional)
                    try:
                        update_user(user_id, role=role, must_change_pwd=1 if must else 0)
                        if pwd:
                            reset_password(user_id, pwd)
                        messagebox.showinfo("Updated", "User updated.", parent=dlg)
                    except Exception as e:
                        messagebox.showerror("Update failed", f"{e}", parent=dlg)
                        return
                dlg.destroy()
                refresh_users()

            white_btn(btns, text="Cancel", command=dlg.destroy, width=120).pack(side="right", padx=(6,0))
            white_btn(btns, text="Save", command=do_save, width=140).pack(side="right")

        # expose tree reference in case other code expects it
        try:
            self._manage_users_tree = tree
        except Exception as e:
            logger.exception("Caught exception", exc_info=e)
            pass

    # ---------- Staff Profile UI (preview + folder open) ----------
    def staff_profile(self, pick_callback=None):
        """
        Staff Profile window. Backwards-compatible: if pick_callback is provided (callable),
        double-clicking a staff will call `pick_callback(info_dict)` and close the window.
        info_dict contains: {"id": <db id or None>, "staff_id_opt": <staff id str>, "name": <staff name>, "position": <position>, "phone": <phone>}
        If pick_callback is None, behavior is unchanged (management / edit flow).
        """
        # single-instance guard
        self._open_windows = getattr(self, "_open_windows", {})
        if self._open_windows.get("staff_profile"):
            try:
                w = self._open_windows["staff_profile"]
                if w.winfo_exists():
                    w.deiconify(); w.lift(); w.focus_force()
                    return
            except Exception:
                self._open_windows.pop("staff_profile", None)

        top = ctk.CTkToplevel(self)
        top.title("Staff Profile")
        TOP_W, TOP_H = 900, 520
        top.geometry(f"{TOP_W}x{TOP_H}")
        top.resizable(False, False)
        top.grab_set()

        self._open_windows["staff_profile"] = top
        top.protocol("WM_DELETE_WINDOW", lambda: (self._open_windows.pop("staff_profile", None), top.destroy()))

        # Main layout
        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=10, pady=10)
        frm.grid_rowconfigure(1, weight=1)  # tree on row 1 will expand
        frm.grid_columnconfigure(0, weight=1)
    
        # Row 0: input panel
        panel = ctk.CTkFrame(frm)
        panel.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        panel.grid_columnconfigure(1, weight=1)

        # Shared width for Position and Name
        SHARED_WIDTH = 260

        ctk.CTkLabel(panel, text="Position").grid(row=0, column=0, padx=(6,8), sticky="w")
        cb_pos = ctk.CTkComboBox(panel, values=["Technician", "Sales", "Admin"], width=SHARED_WIDTH)
        cb_pos.set("Technician")
        cb_pos.grid(row=0, column=1, padx=(0,8), sticky="w")

        ctk.CTkLabel(panel, text="Staff ID").grid(row=0, column=2, padx=(6,8), sticky="w")
        e_staffid = ctk.CTkEntry(panel, width=160)
        e_staffid.grid(row=0, column=3, padx=(0,8), sticky="w")

        ctk.CTkLabel(panel, text="Name").grid(row=1, column=0, padx=(6,8), sticky="w", pady=(8,0))
        e_name = ctk.CTkEntry(panel, width=SHARED_WIDTH)
        e_name.grid(row=1, column=1, padx=(0,8), sticky="w", pady=(8,0))

        ctk.CTkLabel(panel, text="Phone Number").grid(row=1, column=2, padx=(6,8), sticky="w", pady=(8,0))
        e_phone = ctk.CTkEntry(panel, width=160)
        e_phone.grid(row=1, column=3, padx=(0,8), sticky="w", pady=(8,0))

        # Add Staff / Delete / Reset buttons (inline handlers)
        btn_panel = ctk.CTkFrame(panel)
        btn_panel.grid(row=0, column=4, rowspan=2, padx=(8,4), sticky="e")

        # ... (handlers on_add_staff / on_delete_staff / on_reset_fields are unchanged) ...
        # Reuse your existing handlers by copy/paste (safe) — for brevity keep same implementations:
        def on_add_staff():
            name = e_name.get().strip()
            staff_id_opt = e_staffid.get().strip()
            position = cb_pos.get().strip()
            phone = normalize_phone(e_phone.get())
            if not name:
                messagebox.showerror("Missing", "Please enter Name.")
                return
            now = datetime.now().isoformat(sep=" ", timespec="seconds")
            try:
                conn = get_conn()
                cur = conn.cursor()
                cur.execute("""
                    INSERT OR IGNORE INTO staffs (position, staff_id_opt, name, phone, photo_path, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (position, staff_id_opt, name, phone, "", now, now))
                cur.execute("""
                    UPDATE staffs
                       SET position=?, staff_id_opt=?, phone=?, updated_at=?
                     WHERE name=?
                """, (position, staff_id_opt, phone, now, name))
                conn.commit()
                conn.close()
            except Exception as e:
                messagebox.showerror("Add Staff", f"Failed to add staff:\n{e}")
                return
            e_staffid.delete(0, "end")
            e_name.delete(0, "end")
            e_phone.delete(0, "end")
            cb_pos.set("Technician")
            try:
                refresh()
            except Exception:
                pass

        def on_delete_staff():
            name = e_name.get().strip()
            staff_id_opt = e_staffid.get().strip()
            if not name and not staff_id_opt:
                messagebox.showerror("Delete", "Enter Staff ID or Name to delete.")
                return
            if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this staff record?"):
                return
            try:
                conn = get_conn()
                cur = conn.cursor()
                if staff_id_opt and name:
                    cur.execute("DELETE FROM staffs WHERE staff_id_opt=? AND name=?", (staff_id_opt, name))
                elif staff_id_opt:
                    cur.execute("DELETE FROM staffs WHERE staff_id_opt=?", (staff_id_opt,))
                else:
                    cur.execute("DELETE FROM staffs WHERE name=?", (name,))
                affected = cur.rowcount
                conn.commit()
                conn.close()
                if affected == 0:
                    messagebox.showinfo("Delete", "No matching staff found to delete.")
                else:
                    messagebox.showinfo("Delete", f"Deleted {affected} record(s).")
            except Exception as e:
                messagebox.showerror("Delete Staff", f"Failed to delete staff:\n{e}")
                return
            e_staffid.delete(0, "end")
            e_name.delete(0, "end")
            e_phone.delete(0, "end")
            cb_pos.set("Technician")
            try:
                refresh()
            except Exception:
                pass

        def on_reset_fields():
            e_staffid.delete(0, "end")
            e_name.delete(0, "end")
            e_phone.delete(0, "end")
            cb_pos.set("Technician")

        white_btn(btn_panel, text="Add Staff", width=120, command=on_add_staff).pack(side="top", pady=(0,6))
        white_btn(btn_panel, text="Delete", width=120, command=on_delete_staff).pack(side="top", pady=(0,6))
        white_btn(btn_panel, text="Reset", width=120, command=on_reset_fields).pack(side="top")

        # Row 1: Treeview container
        container = tk.Frame(frm)
        container.grid(row=1, column=0, sticky="nsew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        cols = ("position", "staff_id", "name", "phone")
        display_names = {"position":"Position", "staff_id":"Staff ID", "name":"Name", "phone":"Phone"}

        tree = ttk.Treeview(container, columns=cols, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        for c in cols:
            tree.heading(c, text=display_names.get(c, c.title()))

        def autosize_columns():
            try:
                container.update_idletasks()
                padding = 20
                total_w = container.winfo_width() - padding
                if total_w <= 0: total_w = TOP_W - 40
                samples = {cid: [str(tree.heading(cid)['text'])] for cid in cols}
                for iid in tree.get_children()[:40]:
                    vals = tree.item(iid)["values"]
                    for idx, cid in enumerate(cols):
                        try:
                            samples[cid].append(str(vals[idx] or ""))
                        except Exception:
                            pass
                min_widths = []
                for cid in cols:
                    longest = max((len(s) for s in samples[cid]), default=6)
                    min_w = max(80, int(longest * 7 + 20))
                    min_widths.append(min_w)
                sum_min = sum(min_widths)
                if sum_min <= total_w:
                    extra = total_w - sum_min
                    widths = [mw + int(extra * (mw / sum_min)) for mw in min_widths]
                else:
                    factor = total_w / sum_min
                    widths = [max(60, int(mw * factor)) for mw in min_widths]
                for cid, w in zip(cols, widths):
                    tree.column(cid, width=w, anchor="w", stretch=False)
            except Exception:
                base = int((TOP_W - 80) / len(cols))
                for cid in cols:
                    tree.column(cid, width=base, anchor="w", stretch=True)

        def refresh():
            tree.delete(*tree.get_children())
            try:
                conn = get_conn()
                cur = conn.cursor()
                cur.execute("SELECT position, staff_id_opt, name, phone FROM staffs ORDER BY name COLLATE NOCASE")
                rows = cur.fetchall()
                conn.close()
            except Exception:
                rows = [
                    ("Technician", "T001", "Alice Lee", "012-3456789"),
                    ("Sales", "S005", "Bob Tan", "019-888777"),
                ]
            for r in rows:
                pos = r[0] if len(r) > 0 else ""
                sid = r[1] if len(r) > 1 else ""
                nm = r[2] if len(r) > 2 else ""
                phone = r[3] if len(r) > 3 else ""
                tree.insert("", "end", values=(pos, sid, nm, phone))
            autosize_columns()

        # Double-click behavior:
        if callable(pick_callback):
            # Picker mode: choose a staff and close
            def _pick_for_report(e=None):
                sel = tree.selection()
                if not sel:
                    return
                vals = tree.item(sel[0])["values"]
                try:
                    picked = {
                        "id": None,
                        "staff_id_opt": vals[1] if len(vals) > 1 else "",
                        "name": vals[2] if len(vals) > 2 else "",
                        "position": vals[0] if len(vals) > 0 else "",
                        "phone": vals[3] if len(vals) > 3 else "",
                    }
                except Exception:
                    picked = {"id": None, "staff_id_opt": "", "name": "", "position": "", "phone": ""}
                try:
                    # resolve DB id if possible
                    try:
                        conn = get_conn()
                        cur = conn.cursor()
                        cur.execute("SELECT id FROM staffs WHERE name=? COLLATE NOCASE AND (staff_id_opt=? OR staff_id_opt='') LIMIT 1", (picked["name"], picked["staff_id_opt"]))
                        r = cur.fetchone()
                        if r:
                            picked["id"] = r[0]
                        conn.close()
                    except Exception:
                        pass
                except Exception:
                    pass
                try:
                    pick_callback(picked)
                except Exception:
                    logger.exception("pick_callback raised", exc_info=True)
                try:
                    top.destroy()
                except Exception:
                    pass
            tree.bind("<Double-1>", _pick_for_report)
        else:
            # Default management behavior (leave existing behavior as-is)
            try:
                tree.bind("<Double-1>", lambda e: self._edit_staff(tree, top))
            except Exception:
                # if _edit_staff doesn't exist or binding fails, ignore silently
                logger.debug("staff_profile: default double-click bind to _edit_staff failed or _edit_staff missing", exc_info=True)

        # initial load and sizing
        try:
            refresh()
            top.update_idletasks()
            autosize_columns()
        except Exception:
            pass

        # expose tree for outside helpers if needed
        try:
            self._staff_profile_tree = tree
        except Exception:
            pass

    # ----------------- Service / Sales Report UI -----------------
    def service_sales_report_ui(self):
        """
        Service & Sales Report UI.
        - Choose Staff (opens a small chooser dialog)
        - All Staff (aggregated) mode
        - Date from / to filters (UI: DD-MM-YYYY). Date To defaults to today.
        - Table columns: Bill Date | Bill No. | Bill Amount | Commission Amount
        - Total Commission displayed at bottom-right.
        - Show Details checkbox: double-click a row to open related vouchers.
        """
        # Single-instance guard
        self._open_windows = getattr(self, "_open_windows", {})
        if self._open_windows.get("service_sales_report"):
            try:
                w = self._open_windows["service_sales_report"]
                if w.winfo_exists():
                    w.deiconify(); w.lift(); w.focus_force()
                    return
            except Exception:
                self._open_windows.pop("service_sales_report", None)

        top = ctk.CTkToplevel(self)
        top.title("Service / Sales Report")
        top.geometry("1100x680")
        top.resizable(False, False)
        top.grab_set()
        self._open_windows["service_sales_report"] = top
        top.protocol("WM_DELETE_WINDOW", lambda: (self._open_windows.pop("service_sales_report", None), top.destroy()))

        # State holders for the chosen staff (DB id, staff_id_opt, name, role)
        selected = {"db_id": None, "staff_id_opt": None, "name": None, "role": None}
        # All-staff toggle
        all_staff_var = tk.BooleanVar(value=False)

        # ---------- Top filter frame ----------
        frm = ctk.CTkFrame(top)
        frm.pack(fill="x", padx=8, pady=(8, 6))

        # Row 0: Staff ID chooser and readonly fields for name + role
        lbl_sid = ctk.CTkLabel(frm, text="Staff ID:")
        lbl_sid.grid(row=0, column=0, padx=(6,6), pady=6, sticky="w")

        # Placeholder frame where either "Choose Staff" button or the readonly id label will appear
        sid_holder = ctk.CTkFrame(frm, corner_radius=4)
        sid_holder.grid(row=0, column=1, padx=(0,8), pady=6, sticky="w")

        # Function to render the Choose Staff button
        def render_choose_button():
            for w in sid_holder.winfo_children():
                w.destroy()
            btn = white_btn(sid_holder, text="Choose Staff", width=120)
            btn.pack(side="left")
            btn.configure(command=open_staff_chooser)

        # Readonly ID label (shown after selection or when All Staff)
        lbl_selected_id = ctk.CTkLabel(sid_holder, text="", anchor="w")

        # Staff Name label + non-editable entry (shortened)
        ctk.CTkLabel(frm, text="Staff Name:").grid(row=0, column=2, padx=(6,6), sticky="w")
        ent_name = ctk.CTkEntry(frm, width=300)
        ent_name.grid(row=0, column=3, padx=(0,8), sticky="w")
        ent_name.configure(state="disabled")

        # Role label + non-editable entry (aligned with date inputs)
        ctk.CTkLabel(frm, text="Role:").grid(row=0, column=4, padx=(6,6), sticky="w")
        ent_role = ctk.CTkEntry(frm, width=200)
        ent_role.grid(row=0, column=5, padx=(0,8), sticky="w")
        ent_role.configure(state="disabled")

        # All Staff checkbox placed near ID area
        all_chk = ctk.CTkCheckBox(frm, text="All Staff", variable=all_staff_var, width=120)

        def _on_all_toggle(*_a):
            if all_staff_var.get():
                # Clear selection and show aggregated text
                selected.update({"db_id": None, "staff_id_opt": None, "name": "All Staff (aggregated)", "role": ""})
                # hide choose button and show aggregated label
                for w in sid_holder.winfo_children():
                    w.destroy()
                lbl_selected_id.configure(text="All Staff (aggregated)")
                lbl_selected_id.pack(side="left", padx=(6,0))
                ent_name.configure(state="normal"); ent_name.delete(0, "end"); ent_name.insert(0, "All Staff (aggregated)"); ent_name.configure(state="disabled")
                ent_role.configure(state="normal"); ent_role.delete(0, "end"); ent_role.insert(0, ""); ent_role.configure(state="disabled")
            else:
                # revert to choose button and blank fields
                selected.update({"db_id": None, "staff_id_opt": None, "name": None, "role": None})
                for w in sid_holder.winfo_children():
                    w.destroy()
                render_choose_button()
                ent_name.configure(state="normal"); ent_name.delete(0, "end"); ent_name.configure(state="disabled")
                ent_role.configure(state="normal"); ent_role.delete(0, "end"); ent_role.configure(state="disabled")

        # Position the checkbox (right of the id holder)
        all_chk.grid(row=0, column=6, padx=(6,8), sticky="w")
        all_chk.configure(command=_on_all_toggle)

        # initialize choose button
        render_choose_button()

        # Row 1: Date filters
        ctk.CTkLabel(frm, text="Date From (DD-MM-YYYY):").grid(row=1, column=0, padx=(6,6), sticky="w")
        e_from = ctk.CTkEntry(frm, width=160)
        e_from.grid(row=1, column=1, padx=(0,8), sticky="w")

        ctk.CTkLabel(frm, text="Date To (DD-MM-YYYY):").grid(row=1, column=2, padx=(6,6), sticky="w")
        e_to = ctk.CTkEntry(frm, width=200)
        e_to.grid(row=1, column=3, padx=(0,8), sticky="w")
        e_to.delete(0, "end")
        e_to.insert(0, _to_ui_date(datetime.now()))

        # Buttons: Generate and Reset
        btn_generate = white_btn(frm, text="Generate", width=120)
        btn_generate.grid(row=1, column=5, padx=(6,6), sticky="e")

        btn_reset = white_btn(frm, text="Reset", width=120)
        btn_reset.grid(row=1, column=6, padx=(6,6), sticky="e")

        # Show Details checkbox (small)
        show_details_var = tk.BooleanVar(value=False)
        cb_show_details = ctk.CTkCheckBox(frm, text="Show Details", variable=show_details_var)
        cb_show_details.grid(row=0, column=7, padx=(12,8), sticky="e")

        # ---------- Table area ----------
        table_wrap = ctk.CTkFrame(top)
        table_wrap.pack(fill="both", expand=True, padx=8, pady=(6, 6))

        cols = ("bill_date", "bill_no", "bill_amount", "commission_amount")
        tree = ttk.Treeview(table_wrap, columns=cols, show="headings", selectmode="browse")
        headings = [
            ("bill_date", "Bill Date", 140),
            ("bill_no", "Bill No.", 220),     # slightly shorter
            ("bill_amount", "Bill Amount (RM)", 180),  # lengthened
            ("commission_amount", "Commission Amount (RM)", 200),  # lengthened
        ]
        for key, title, w in headings:
            tree.heading(key, text=title)
            tree.column(key, width=w, anchor="w", stretch=False)
        tree.pack(fill="both", expand=True, side="left", padx=(0,6), pady=6)

        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=tree.yview)
        vsb.pack(side="left", fill="y")
        tree.configure(yscrollcommand=vsb.set)

        # Bottom area: total commission display (right aligned)
        bottom = ctk.CTkFrame(top)
        bottom.pack(fill="x", padx=8, pady=(0, 12))
        total_lbl = ctk.CTkLabel(bottom, text="Total Commission (RM):", anchor="e")
        total_val = ctk.CTkEntry(bottom, width=200)
        total_val.configure(state="disabled")
        # place them to the right
        total_lbl.pack(side="right", padx=(6,4))
        total_val.pack(side="right", padx=(0,16))

        # ---------- Helpers for staff chooser and UI updates ----------
        def open_staff_chooser():
            """Open a small staff picker (double-click selects)."""
            pick = ctk.CTkToplevel(top)
            pick.title("Choose Staff")
            pick.geometry("720x420")
            pick.grab_set()
            wrap = ctk.CTkFrame(pick)
            wrap.pack(fill="both", expand=True, padx=8, pady=8)

            cols_local = ("staff_id_opt", "name", "position", "phone")
            t = ttk.Treeview(wrap, columns=cols_local, show="headings", selectmode="browse")
            widths = [140, 300, 140, 140]
            headers = [("staff_id_opt","Staff ID",widths[0]), ("name","Name",widths[1]), ("position","Position",widths[2]), ("phone","Phone",widths[3])]
            for k, title, w_ in headers:
                t.heading(k, text=title); t.column(k, width=w_, anchor="w", stretch=False)
            t.pack(fill="both", expand=True, side="left", padx=(0,6), pady=6)
            vs = ttk.Scrollbar(wrap, orient="vertical", command=t.yview); vs.pack(side="left", fill="y"); t.configure(yscrollcommand=vs.set)

            # load rows
            try:
                conn = get_conn(); cur = conn.cursor()
                cur.execute("SELECT staff_id_opt, name, position, phone, id FROM staffs ORDER BY name COLLATE NOCASE")
                rows = cur.fetchall()
                conn.close()
            except Exception:
                rows = []
            for r in rows:
                # r: (staff_id_opt, name, position, phone, id)
                sid = r[0] or ""
                nm = r[1] or ""
                pos = r[2] or ""
                phone = r[3] or ""
                dbid = r[4] if len(r) > 4 else None
                t.insert("", "end", iid=str(dbid), values=(sid, nm, pos, phone))

            def _pick(evt=None):
                sel = t.selection()
                if not sel:
                    messagebox.showinfo("Choose", "Select a staff first.", parent=pick)
                    return
                iid = sel[0]
                vals = t.item(iid)["values"]
                sid_opt = vals[0] or ""
                nm = vals[1] or ""
                pos = vals[2] or ""
                # set selected state
                selected["db_id"] = int(iid) if iid and iid.isdigit() else None
                selected["staff_id_opt"] = sid_opt
                selected["name"] = nm
                selected["role"] = pos
                # Update UI: hide choose button, show selected id label and readonly fields
                for w in sid_holder.winfo_children():
                    w.destroy()
                lbl_selected_id.configure(text=str(sid_opt or selected["db_id"] or ""))
                lbl_selected_id.pack(side="left", padx=(6,0))
                ent_name.configure(state="normal"); ent_name.delete(0, "end"); ent_name.insert(0, nm); ent_name.configure(state="disabled")
                ent_role.configure(state="normal"); ent_role.delete(0, "end"); ent_role.insert(0, pos); ent_role.configure(state="disabled")
                # ensure All Staff unchecked
                all_staff_var.set(False)
                try:
                    pick.destroy()
                except Exception:
                    pass

            t.bind("<Double-1>", _pick)
            # Buttons
            btnf = ctk.CTkFrame(pick); btnf.pack(fill="x", padx=8, pady=(6,8))
            white_btn(btnf, text="Select", width=120, command=_pick).pack(side="right", padx=(6,0))
            white_btn(btnf, text="Close", width=120, command=pick.destroy).pack(side="right", padx=(0,6))

        def reset_filters():
            # Reset visual state and filters
            all_staff_var.set(False)
            _on_all_toggle()
            e_from.delete(0, "end")
            e_to.delete(0, "end"); e_to.insert(0, _to_ui_date(datetime.now()))
            # clear table & totals
            tree.delete(*tree.get_children())
            total_val.configure(state="normal"); total_val.delete(0, "end"); total_val.configure(state="disabled")

        btn_reset.configure(command=reset_filters)

        # ---------- Data loading and mapping ----------
        def _parse_ui_date_safe(s):
            s = (s or "").strip()
            if not s:
                return ""
            iso = _parse_ui_date_to_iso(s)
            return iso

        def load_commissions():
            # Build query depending on selected staff / all-staff and dates
            date_from_iso = _parse_ui_date_safe(e_from.get().strip())
            date_to_iso = _parse_ui_date_safe(e_to.get().strip())

            params = []
            where_clauses = []
            # staff filtering
            if not all_staff_var.get():
                # require a selected staff
                if not selected.get("db_id") and not selected.get("staff_id_opt"):
                    messagebox.showerror("Missing", 'Choose a staff first or enable "All Staff".')
                    return False
                # prefer staff.db id if available
                if selected.get("db_id"):
                    where_clauses.append("c.staff_id = ?")
                    params.append(selected["db_id"])
                elif selected.get("staff_id_opt"):
                    where_clauses.append("s.staff_id_opt = ?")
                    params.append(selected["staff_id_opt"])
            # date filters (created_at of commissions)
            if date_from_iso:
                where_clauses.append("DATE(c.created_at) >= DATE(?)"); params.append(date_from_iso)
            if date_to_iso:
                where_clauses.append("DATE(c.created_at) <= DATE(?)"); params.append(date_to_iso)

            where_sql = ("WHERE " + " AND ".join(where_clauses)) if where_clauses else ""
            sql = f"""
                SELECT c.id, c.bill_type, c.bill_no, c.total_amount, c.commission_amount, c.created_at, c.voucher_id, s.staff_id_opt, s.name
                FROM commissions c
                LEFT JOIN staffs s ON c.staff_id = s.id
                {where_sql}
                ORDER BY c.created_at DESC
            """
            try:
                conn = get_conn(); cur = conn.cursor()
                cur.execute(sql, params)
                rows = cur.fetchall()
                conn.close()
            except Exception as e:
                logger.exception("Failed loading commissions for report", exc_info=e)
                rows = []

            # populate tree and compute total
            tree.delete(*tree.get_children())
            total = 0.0
            for row in rows:
                # row: (id, bill_type, bill_no, total_amount, commission_amount, created_at, voucher_id, staff_id_opt, name)
                cid = row[0]
                bt = (row[1] or "").upper()
                bno = row[2] or ""
                amt = row[3] if row[3] is not None else ""
                comm = row[4] if row[4] is not None else 0.0
                created_at = row[5] or ""
                # format date (dd-mm-yyyy)
                try:
                    dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
                    bill_date = dt.strftime("%d-%m-%Y")
                except Exception:
                    bill_date = (created_at or "")[:10]
                # display amounts with 2 decimal
                try:
                    amt_display = f"{float(amt):.2f}" if amt != "" else ""
                except Exception:
                    amt_display = str(amt)
                try:
                    comm_val = float(comm) if comm not in (None, "") else 0.0
                except Exception:
                    comm_val = 0.0
                comm_display = f"{comm_val:.2f}"
                total += comm_val

                tree.insert("", "end", iid=str(cid), values=(bill_date, bno, amt_display, comm_display))

            # update total display
            total_val.configure(state="normal"); total_val.delete(0, "end"); total_val.insert(0, f"{total:.2f}"); total_val.configure(state="disabled")
            return True

        def on_generate():
            ok = load_commissions()
            if ok:
                # scroll to top
                try:
                    if tree.get_children():
                        first = tree.get_children()[0]
                        tree.see(first)
                except Exception:
                    pass

        btn_generate.configure(command=on_generate)

        # ---------- Double-click details handler ----------
        def _show_voucher_details_for_commission(event=None):
            sel = tree.selection()
            if not sel:
                return
            cid = sel[0]
            # We need to look up the commission row to find voucher_id or bill_no
            try:
                conn = get_conn(); cur = conn.cursor()
                cur.execute("SELECT id, voucher_id, bill_no FROM commissions WHERE id=?", (cid,))
                crow = cur.fetchone()
                conn.close()
            except Exception:
                crow = None
            if not crow:
                messagebox.showinfo("Details", "No commission info found.")
                return
            _, voucher_id, bill_no = crow
            # Build voucher query
            q_params = []
            q_where = []
            if voucher_id:
                q_where.append("voucher_id = ?"); q_params.append(voucher_id)
            if bill_no:
                q_where.append("LOWER(ref_bill) = LOWER(?)"); q_params.append(bill_no)
            if not q_where:
                messagebox.showinfo("Details", "No linked vouchers found for this commission.")
                return
            qsql = f"SELECT voucher_id, created_at, customer_name, contact_number, amount_rm, tech_commission, status FROM vouchers WHERE {' OR '.join(q_where)} ORDER BY created_at DESC"
            try:
                conn = get_conn(); cur = conn.cursor()
                cur.execute(qsql, q_params)
                vrows = cur.fetchall()
                conn.close()
            except Exception as e:
                logger.exception("Failed loading vouchers for commission", exc_info=e)
                vrows = []

            if not vrows:
                messagebox.showinfo("Details", "No linked vouchers found for this commission.")
                return

            det = ctk.CTkToplevel(top)
            det.title(f"Voucher(s) for Commission {cid}")
            det.geometry("980x420")
            det.grab_set()

            wrap = ctk.CTkFrame(det); wrap.pack(fill="both", expand=True, padx=8, pady=8)
            cols_v = ("voucher_id","created_at","customer_name","contact_number","amount_rm","tech_commission","status")
            tv = ttk.Treeview(wrap, columns=cols_v, show="headings")
            widths_v = [100,140,220,120,110,120,100]
            for k,w_ in zip(cols_v, widths_v):
                tv.heading(k, text=k.replace("_"," ").title()); tv.column(k, width=w_, anchor="w", stretch=False)
            tv.pack(fill="both", expand=True, side="left", padx=(0,6))
            vs = ttk.Scrollbar(wrap, orient="vertical", command=tv.yview); vs.pack(side="left", fill="y"); tv.configure(yscrollcommand=vs.set)

            for vr in vrows:
                vid, created_at, cname, contact, amt, comm_amt, status = vr
                try:
                    dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
                    created_f = dt.strftime("%d-%m-%Y")
                except Exception:
                    created_f = (created_at or "")[:10]
                tv.insert("", "end", values=(vid, created_f, cname or "", contact or "", f"{amt or ''}", f"{comm_amt or ''}", status or ""))

            btns = ctk.CTkFrame(det); btns.pack(fill="x", padx=8, pady=(6,8))
            white_btn(btns, text="Close", width=120, command=det.destroy).pack(side="right", padx=(6,0))

        # Only show details dialog when Show Details is ticked
        def _maybe_show_details(event=None):
            if not show_details_var.get():
                return
            _show_voucher_details_for_commission(event)

        # bind double-click
        try:
            tree.bind("<Double-1>", _maybe_show_details)
        except Exception:
            pass

        # initial state: clear table
        reset_filters()

        try:
            top.update_idletasks()
        except Exception:
            pass


    # ---------- Commission UI (with preview + per-staff storage) ----------
    def add_commission(self):
        """Add Commission dialog (note field removed to match DB schema)."""
        top = ctk.CTkToplevel(self)
        top.title("Add Commission")
        # slightly larger so form fields & buttons fit comfortably
        top.geometry("820x520")
        top.grab_set()

        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=12, pady=12)
        for c in range(4):
            frm.grid_columnconfigure(c, weight=1)

        r = 0
        ctk.CTkLabel(frm, text="Staff").grid(row=r, column=0, sticky="w")
        staff_names = list_staffs_names()
        cb_staff = ctk.CTkComboBox(frm, values=["— Select —"] + staff_names, width=420)
        cb_staff.set("— Select —")
        cb_staff.grid(row=r, column=1, sticky="w", padx=8, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="Bill Type").grid(row=r, column=0, sticky="w")
        cb_type = ctk.CTkComboBox(frm, values=["Cash Bill", "Invoice"], width=200)
        cb_type.set("Cash Bill")
        cb_type.grid(row=r, column=1, sticky="w", padx=8, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="Bill No").grid(row=r, column=0, sticky="w")
        e_bill = ctk.CTkEntry(frm, width=420)
        e_bill.grid(row=r, column=1, sticky="w", padx=8, pady=6)
        r += 1

        # Total amount (under bill no)
        ctk.CTkLabel(frm, text="Total Amount (RM)").grid(row=r, column=0, sticky="w")
        e_total = ctk.CTkEntry(frm, width=200)
        e_total.grid(row=r, column=1, sticky="w", padx=8, pady=6)
        r += 1

        # -- NOTE field removed intentionally --

        # Auto-prefix for bill based on type
        def _on_comm_bill_type_change(*_a):
            bt = (cb_type.get() or "").strip()
            today = datetime.now()
            mm = f"{today.month:02d}"
            dd = f"{today.day:02d}"
            if bt == "Cash Bill":
                prefix = f"CS-{mm}{dd}/"
            else:
                prefix = f"INV-{mm}{dd}/"
            curv = (e_bill.get() or "").strip().upper()
            if not curv or re.match(r"^(CS|INV)-\d{4}/", curv, re.IGNORECASE):
                try:
                    e_bill.delete(0, "end")
                    e_bill.insert(0, prefix)
                except Exception:
                    pass

        try:
            cb_type.configure(command=_on_comm_bill_type_change)
        except Exception:
            cb_type.bind("<<ComboboxSelected>>", lambda e: _on_comm_bill_type_change())
        _on_comm_bill_type_change()

        # Buttons
        btns = ctk.CTkFrame(top)
        btns.pack(fill="x", padx=12, pady=(6, 12))

        def save_commission():
            # basic validations
            staff_name = cb_staff.get().strip()
            if staff_name in ("— Select —", ""):
                messagebox.showerror("Missing", "Please select a staff for this commission.", parent=top)
                return
            try:
                comm_amt = float((e_total.get() or "").strip()) if (e_total.get() or "").strip() else None
            except Exception:
                messagebox.showerror("Invalid", "Total Amount must be a number.", parent=top)
                return

            bill_type_text = cb_type.get()
            bill_type = "CS" if bill_type_text == "Cash Bill" else "INV"
            bill_no = (e_bill.get() or "").strip().upper()
            # validate bill format if not empty
            if bill_no:
                ok = False
                if bill_type == "CS" and BILL_RE_CS.match(bill_no):
                    ok = True
                if bill_type == "INV" and BILL_RE_INV.match(bill_no):
                    ok = True
                if not ok:
                    messagebox.showerror("Bill No.", "Invalid bill number format.\nCash: CS-MMDD/XXXX\nInvoice: INV-MMDD/XXXX", parent=top)
                    return

            # resolve staff id from name (if exists)
            staff_db_id = None
            try:
                conn = get_conn()
                cur = conn.cursor()
                cur.execute("SELECT id FROM staffs WHERE name=? COLLATE NOCASE", (staff_name,))
                srow = cur.fetchone()
                if srow:
                    staff_db_id = srow[0]
                conn.close()
            except Exception:
                staff_db_id = None

            # insert commission (NO 'note' column)
            try:
                conn = get_conn()
                cur = conn.cursor()
                now = datetime.now().isoformat(sep=' ', timespec='seconds')
                cur.execute("""
                    INSERT INTO commissions
                        (staff_id, bill_type, bill_no, total_amount, commission_amount, bill_image_path, created_at, updated_at)
                    VALUES (?,?,?,?,?,?,?,?)
                """, (staff_db_id, bill_type, bill_no or "", comm_amt if comm_amt is not None else None, 0.0, "", now, now))
                conn.commit()
                conn.close()
            except Exception as e:
                messagebox.showerror("Save Failed", f"Failed to create commission:\n{e}", parent=top)
                try:
                    conn.close()
                except Exception:
                    pass
                return

            messagebox.showinfo("Saved", "Commission record created.", parent=top)
            try:
                top.destroy()
            except Exception:
                pass
            # refresh commissions view if open
            try:
                self.perform_search()
            except Exception:
                pass

        white_btn(btns, text="Save", command=save_commission, width=140).pack(side="right")
        white_btn(btns, text="Cancel", command=lambda: top.destroy(), width=100).pack(side="right", padx=(0,8))

        try:
            top.update_idletasks()
        except Exception:
            pass

    def view_commissions(self):
        """View / Edit Commissions - cleaned, widened layout, single handlers."""
        top = ctk.CTkToplevel(self)
        top.title("Commissions")
        top.geometry("1280x720")
        top.resizable(False, False)
        top.grab_set()

        # --- Top toolbar ---
        bar = ctk.CTkFrame(top)
        bar.pack(fill="x", padx=8, pady=(8, 4))

        e_q = ctk.CTkEntry(bar, placeholder_text="Search staff, bill no or id", width=520)
        e_q.pack(side="left", padx=(6, 8))

        def do_search():
            refresh(e_q.get())

        # keep the main toolbar items compact, remove per-row action buttons
        white_btn(bar, text="Filter", width=100, command=do_search).pack(side="left", padx=(0, 6))
        white_btn(bar, text="Refresh", width=100, command=lambda: refresh("")).pack(side="left", padx=(0, 6))

        btn_frame = ctk.CTkFrame(bar)
        btn_frame.pack(side="right", padx=6)

        # Keep only Add Commission, make it wider so the full label is visible
        white_btn(btn_frame, text="Add Commission", width=200, command=lambda: (self.add_commission(), refresh(e_q.get()))).pack(side="right", padx=(6, 0))


        # --- Tree / list area ---
        wrap = ctk.CTkFrame(top)
        wrap.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        cols = ("id", "staff_id", "staff_name", "bill_type", "bill_no", "total_amount", "commission_amount",
                "voucher_id", "created_at")
        tree = ttk.Treeview(wrap, columns=cols, show="headings", selectmode="browse")
        headings = [
            ("id", "ID", 60),
            ("staff_id", "StaffID", 110),
            ("staff_name", "Staff", 260),
            ("bill_type", "Type", 90),
            ("bill_no", "Bill No", 340),
            ("total_amount", "Total", 140),
            ("commission_amount", "Commission", 160),
            ("voucher_id", "Voucher ID", 140),
            ("created_at", "Created", 180),
        ]
        for key, title, w in headings:
            tree.heading(key, text=title)
            tree.column(key, width=w, anchor="w", stretch=False)
        tree.pack(fill="both", expand=True, side="left", padx=(0, 6), pady=6)

        vsb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        vsb.pack(side="left", fill="y")
        tree.configure(yscrollcommand=vsb.set)

        # --- Loader ---
        def _load_rows(q=""):
            qlow = (q or "").strip().lower()
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("""
                SELECT c.id, c.staff_id, s.name, c.bill_type, c.bill_no, c.total_amount, c.commission_amount, c.voucher_id, c.created_at
                FROM commissions c
                LEFT JOIN staffs s ON c.staff_id = s.id
                ORDER BY c.created_at DESC
            """)
            rows = cur.fetchall()
            conn.close()
            tree.delete(*tree.get_children())
            for (cid, sid, sname, bt, bno, tot, comm, vid, created_at) in rows:
                display = (cid, sid or "", sname or "", bt or "", bno or "", tot or "", comm or "", vid or "",
                           created_at or "")
                if qlow and qlow not in " ".join(map(str, display)).lower():
                    continue
                tree.insert("", "end", iid=str(cid), values=display)

        _load_rows()

        def _do_filter():
            _load_rows(e_q.get())

        # Wire toolbar filter button defensively
        try:
            b_search = [w for w in bar.winfo_children() if isinstance(w, ctk.CTkButton) and w.cget("text") == "Filter"][0]
            b_search.configure(command=_do_filter)
        except Exception:
            pass

        # --- Actions ---
        def _delete_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("Delete", "Select a commission to delete.", parent=top)
                return
            cid = int(sel[0])
            if not messagebox.askyesno("Confirm", f"Delete commission {cid}?", parent=top):
                return
            try:
                conn = get_conn()
                cur = conn.cursor()
                cur.execute("DELETE FROM commissions WHERE id=?", (cid,))
                conn.commit()
                conn.close()
                refresh(e_q.get())
            except Exception as e:
                messagebox.showerror("Delete Failed", f"Failed to delete: {e}", parent=top)

        def bind_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("Bind", "Select a commission to bind.", parent=top)
                return
            try:
                cid = int(sel[0])
            except Exception:
                messagebox.showerror("Bind", "Invalid selection.", parent=top)
                return
            self.open_vouchers_for_binding(cid)
            refresh(e_q.get())

        # --- Edit dialog (opens for selected commission) ---
        def edit_comm():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("Edit", "Select a commission to edit.", parent=top)
                return
            cid = int(sel[0])

            # Load commission
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("""
                SELECT c.id, c.staff_id, s.name, c.bill_type, c.bill_no, c.total_amount, c.commission_amount, c.created_at
                FROM commissions c
                LEFT JOIN staffs s ON c.staff_id = s.id
                WHERE c.id = ?
            """, (cid,))
            row = cur.fetchone()
            conn.close()

            if not row:
                messagebox.showerror("Edit", "Commission record not found.", parent=top)
                return

            _id, staff_id, staff_name, bill_type, bill_no, total_amount, commission_amount, created_at = row

            etop = ctk.CTkToplevel(self)
            etop.title(f"Edit Commission {cid}")
            etop.geometry("980x560")
            etop.resizable(False, False)
            etop.grab_set()

            efrm = ctk.CTkFrame(etop)
            efrm.pack(fill="both", expand=True, padx=12, pady=12)
            for cc in range(4):
                efrm.grid_columnconfigure(cc, weight=1)

            # Staff (readonly)
            ctk.CTkLabel(efrm, text="Staff").grid(row=0, column=0, sticky="w")
            lbl_staff = ctk.CTkLabel(efrm, text=staff_name)
            lbl_staff.grid(row=0, column=1, sticky="w", padx=8, pady=6)

            # Bill Type (use cb_type_edit consistently)
            ctk.CTkLabel(efrm, text="Bill Type").grid(row=1, column=0, sticky="w")
            cb_type_edit = ctk.CTkComboBox(efrm, values=["Cash Bill", "Invoice"], width=240)
            cb_type_edit.set("Cash Bill" if (bill_type or "").upper() == "CS" else "Invoice")
            cb_type_edit.grid(row=1, column=1, sticky="w", padx=8, pady=6)

            # Bill No
            ctk.CTkLabel(efrm, text="Bill No").grid(row=2, column=0, sticky="w")
            e_bill = ctk.CTkEntry(efrm, width=520)
            e_bill.insert(0, bill_no or "")
            e_bill.grid(row=2, column=1, sticky="w", padx=8, pady=6)

            # Total / Commission
            ctk.CTkLabel(efrm, text="Total Amount (RM)").grid(row=3, column=0, sticky="w")
            e_total = ctk.CTkEntry(efrm, width=200)
            e_total.insert(0, str(total_amount or ""))
            e_total.grid(row=3, column=1, sticky="w", padx=8, pady=6)

            ctk.CTkLabel(efrm, text="Commission Amount").grid(row=4, column=0, sticky="w")
            e_comm = ctk.CTkEntry(efrm, width=200)
            e_comm.insert(0, str(commission_amount or ""))
            e_comm.grid(row=4, column=1, sticky="w", padx=8, pady=6)

            # Created (readonly)
            ctk.CTkLabel(efrm, text="Created").grid(row=5, column=0, sticky="w")
            lbl_created = ctk.CTkLabel(efrm, text=created_at or "")
            lbl_created.grid(row=5, column=1, sticky="w", padx=8, pady=6)

            # Buttons row
            btns = ctk.CTkFrame(etop)
            btns.pack(fill="x", padx=12, pady=(6, 12))

            def _validate_bill_format(b):
                s = (b or "").strip().upper()
                bt = cb_type_edit.get()
                if bt == "Cash Bill":
                    return bool(BILL_RE_CS.match(s))
                else:
                    return bool(BILL_RE_INV.match(s))

            def _save_edit():
                new_bill_no = (e_bill.get() or "").strip().upper()
                new_type = "CS" if cb_type_edit.get() == "Cash Bill" else "INV"
                try:
                    new_total = float(e_total.get()) if e_total.get().strip() != "" else None
                    new_comm = float(e_comm.get()) if e_comm.get().strip() != "" else None
                except Exception:
                    messagebox.showerror("Invalid", "Amounts must be numeric.", parent=etop)
                    return

                if not new_bill_no:
                    messagebox.showerror("Missing", "Bill number is required.", parent=etop)
                    return

                if not _validate_bill_format(new_bill_no):
                    messagebox.showerror("Bill No.",
                                         "Invalid bill number format.\nCash: CS-MMDD/XXXX\nInvoice: INV-MMDD/XXXX",
                                         parent=etop)
                    return

                # Duplicate check
                conn_chk = get_conn()
                cur_chk = conn_chk.cursor()
                cur_chk.execute("SELECT id FROM commissions WHERE LOWER(bill_no) = LOWER(?) AND id <> ?", (new_bill_no, cid))
                dup = cur_chk.fetchone()
                conn_chk.close()
                if dup:
                    messagebox.showerror("Duplicate Bill",
                                         f"This bill number ({new_bill_no}) already exists in the system.",
                                         parent=etop)
                    return

                # Commit update
                try:
                    conn2 = get_conn()
                    cur2 = conn2.cursor()
                    cur2.execute("""
                        UPDATE commissions
                        SET bill_type=?, bill_no=?, total_amount=?, commission_amount=?, updated_at=?
                        WHERE id=?
                    """, (
                        new_type, new_bill_no, new_total, new_comm,
                        datetime.now().isoformat(sep=" ", timespec="seconds"), cid
                    ))
                    conn2.commit()
                    conn2.close()
                except sqlite3.IntegrityError:
                    messagebox.showerror("Duplicate Bill", "This bill already exists (case-insensitive).", parent=etop)
                    return
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to update commission:\n{e}", parent=etop)
                    return

                messagebox.showinfo("Updated", "Commission updated.", parent=etop)
                etop.destroy()
                _load_rows(e_q.get())

            white_btn(btns, text="Cancel", width=120, command=etop.destroy).pack(side="right", padx=(6, 0))
            white_btn(btns, text="Save Changes", width=140, command=_save_edit).pack(side="right")

        # Single context menu (create AFTER edit_comm exists)
        try:
            comm_ctx = tk.Menu(tree, tearoff=0)
            comm_ctx.add_command(label="Edit", command=edit_comm)
            comm_ctx.add_command(label="Bind to Voucher", command=bind_selected)
            comm_ctx.add_command(label="Delete", command=_delete_selected)

            def _comm_popup(event):
                try:
                    row = tree.identify_row(event.y)
                    if row and row not in tree.selection():
                        tree.selection_set(row)
                    comm_ctx.post(event.x_root, event.y_root)
                finally:
                    try:
                        comm_ctx.grab_release()
                    except Exception:
                        pass

            tree.bind("<Button-3>", _comm_popup)
        except Exception:
            logger.exception("Caught exception setting up commission context menu", exc_info=True)

        # Admin double-click: edit_comm
        def _on_double(e=None):
            if self.current_user and self.current_user.get("role") == "admin":
                edit_comm()

        tree.bind("<Double-1>", _on_double)

        # helper to refresh
        def refresh(q=""):
            _load_rows(q)

        # expose refresh reference
        try:
            self._refresh_commissions = refresh
        except Exception:
            pass

        try:
            top.update_idletasks()
        except Exception:
            logger.exception("Caught exception while updating idletasks", exc_info=True)
            pass
            
    def open_vouchers_for_binding(self, commission_id: int):
        """
        Open a voucher-picker window so the user can visually select a voucher
        to bind the commission to. After selection, perform validation and bind.
        """
        # load commission info
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT id, bill_type, bill_no, total_amount, commission_amount, voucher_id FROM commissions WHERE id=?", (commission_id,))
        crow = cur.fetchone()
        conn.close()
        if not crow:
            messagebox.showerror("Bind", "Commission not found.")
            return
            
        cid, bill_type, bill_no, total_amount, commission_amount, existing_vid = crow
        # Allow re-binding a commission that was previously bound; we'll clear the old voucher's commission columns
        # when the user confirms binding to a new voucher (done later).
        # (No early return here.)

        pick = ctk.CTkToplevel(self)
        pick.title(f"Select Voucher to bind (Commission {commission_id})")
        # Larger, fixed dialog
        pick.geometry("1100x560")
        pick.resizable(False, False)
        pick.grab_set()

        hint = ctk.CTkLabel(pick, text=f"Choose a voucher to bind — Double-click a voucher or select and press Bind.", anchor="w")
        hint.pack(fill="x", padx=8, pady=(8,4))

        wrap = ctk.CTkFrame(pick)
        wrap.pack(fill="both", expand=True, padx=8, pady=8)

        # search
        fl = ctk.CTkFrame(wrap)
        fl.pack(fill="x", padx=6, pady=(0,6))
        e_q = ctk.CTkEntry(fl, placeholder_text="Search voucher id, customer or contact", width=420)
        e_q.pack(side="left", padx=(0,8))
        b_search = white_btn(fl, text="Filter", width=100)
        b_search.pack(side="left", padx=(4,6))

        cols = ("voucher_id","created_at","customer_name","contact_number","ref_bill","amount_rm","tech_commission","status")
        tree = ttk.Treeview(wrap, columns=cols, show="headings", selectmode="browse")
        headings = [
            ("voucher_id","Voucher ID",120),
            ("created_at","Created",140),
            ("customer_name","Customer",260),
            ("contact_number","Contact",140),
            ("ref_bill","Ref Bill",220),
            ("amount_rm","Amount",110),
            ("tech_commission","Commission",180),
            ("status","Status",120),
        ]
        for key, title, w in headings:
            tree.heading(key, text=title)
            tree.column(key, width=w, anchor="w", stretch=False)
        tree.pack(fill="both", expand=True, padx=6, pady=(0,6))

        vsb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        vsb.pack(side="left", fill="y")
        tree.configure(yscrollcommand=vsb.set)

        def _load_vouchers(q=""):
            qlow = (q or "").strip().lower()
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("""
                SELECT voucher_id, created_at, customer_name, contact_number, ref_bill, amount_rm, tech_commission, status
                FROM vouchers
                ORDER BY created_at DESC
            """)
            rows = cur.fetchall()
            conn.close()
            tree.delete(*tree.get_children())
            for (vid, created_at, cname, contact, refb, amt, comm, status) in rows:
                display = (vid, created_at or "", cname or "", contact or "", refb or "", amt or "", comm or "", status or "")
                if qlow:
                    if qlow not in " ".join(map(str, display)).lower():
                        continue
                tree.insert("", "end", iid=str(vid), values=display)

        _load_vouchers()

        def _do_filter():
            _load_vouchers(e_q.get())

        b_search.configure(command=_do_filter)

        def _bind_to_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("Bind", "Select a voucher to bind.", parent=pick)
                return
            vid = sel[0]
            
            # Validate voucher exists (should)
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("SELECT id FROM commissions WHERE voucher_id=?", (vid,))
            ex = cur.fetchone()
            conn.close()
            if ex and ex[0] != cid:
                messagebox.showerror("Bind", f"Voucher {vid} is already bound to commission id {ex[0]}.", parent=pick)
                return

            # Ensure voucher not already bound to another commission
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("SELECT id FROM commissions WHERE voucher_id=?", (vid,))
            ex = cur.fetchone()
            conn.close()
            if ex:
                messagebox.showerror("Bind", f"Voucher {vid} is already bound to commission id {ex[0]}.", parent=pick)
                return

            # If this commission was previously bound to another voucher, clear commission-related fields on the old voucher.
            try:
                if existing_vid:
                    conn_old = get_conn()
                    cur_old = conn_old.cursor()
                    cur_old.execute("PRAGMA table_info(vouchers)")
                    vcols_old = [r[1] for r in cur_old.fetchall()]
                    
                    # Build safe update to clear the common commission fields if they exist.
                    updates = []
                    params = []
                    if "tech_commission" in vcols_old:
                        updates.append("tech_commission = NULL")
                    if "amount_rm" in vcols_old:
                        updates.append("amount_rm = NULL")
                    if "technician_id" in vcols_old:
                        updates.append("technician_id = NULL")
                    if "technician_name" in vcols_old:
                        updates.append("technician_name = NULL")
                    if "ref_bill" in vcols_old:
                        updates.append("ref_bill = ''")
                    if updates:
                        sql_clear = f"UPDATE vouchers SET {', '.join(updates)} WHERE voucher_id = ?"
                        cur_old.execute(sql_clear, (existing_vid,))
                        conn_old.commit()
                    try:
                        conn_old.close()
                    except Exception:
                        pass
            except Exception:
                logger.exception("Failed clearing old voucher commission fields (continuing)")

            # Perform binding using helper (handles column creation if needed)
            try:
                bind_commission_to_voucher(cid, vid)
            except Exception as e:
                messagebox.showerror("Bind Failed", f"Failed to bind: {e}", parent=pick)
                return

            # Write commission details into voucher row (if columns exist)
            try:
                conn = get_conn()
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(vouchers)")
                vcols = [r[1] for r in cur.fetchall()]
                updates = []
                params = []
                if "technician_id" in vcols and commission_amount is not None:
                    updates.append("technician_id = ?"); params.append(None)
                if "ref_bill" in vcols and bill_no:
                    updates.append("ref_bill=?"); params.append(bill_no)
                if "amount_rm" in vcols and total_amount is not None:
                    updates.append("amount_rm=?"); params.append(total_amount)
                if "tech_commission" in vcols and commission_amount is not None:
                    updates.append("tech_commission=?"); params.append(commission_amount)
                if updates:
                    params.append(vid)
                    cur.execute(f"UPDATE vouchers SET {', '.join(updates)} WHERE voucher_id=?", params)
                    conn.commit()
            except Exception:
                logger.exception("Failed to update voucher with commission info after binding")
            finally:
                try:
                    conn.close()
                except Exception:
                    pass

            messagebox.showinfo("Bound", f"Commission {cid} bound to Voucher {vid}.", parent=pick)
            try:
                pick.destroy()
            except Exception:
                pass

            # Refresh main lists
            try:
                self.perform_search()
            except Exception:
                pass

        # Bind on double-click -> bind
        tree.bind("<Double-1>", lambda e: _bind_to_selected())

        # Buttons
        btns = ctk.CTkFrame(pick)
        btns.pack(fill="x", padx=8, pady=(6,8))
        white_btn(btns, text="Bind Selected", command=_bind_to_selected, width=140).pack(side="right", padx=(6,0))
        white_btn(btns, text="Close", command=lambda: pick.destroy(), width=120).pack(side="right", padx=(0,6))

        try:
            pick.update_idletasks()
        except Exception:
            logger.exception("Caught exception while updating voucher-picker idletasks", exc_info=True)
            pass

    def _go_fullscreen(self):
        """
        Safe fullscreen toggler called after startup.
        Leaves app as normal window on platforms where fullscreen is undesired.
        """
        try:
            # try a gentle fullscreen: maximize window without changing user resizing preferences
            try:
                # For Windows, use state('zoomed') as maximize
                if sys.platform.startswith("win"):
                    self.state("zoomed")
                else:
                    # On mac / linux try the fullscreen attribute
                    self.attributes("-fullscreen", True)
            except Exception:
                # last-resort: set geometry to screen size
                try:
                    w = self.winfo_screenwidth()
                    h = self.winfo_screenheight()
                    self.geometry(f"{w}x{h}")
                except Exception:
                    pass

            # Add an escape handler so user can exit fullscreen
            def _exit_fs(e=None):
                try:
                    if sys.platform.startswith("win"):
                        self.state("normal")
                    else:
                        self.attributes("-fullscreen", False)
                except Exception:
                    try:
                        self.geometry("1280x780")
                    except Exception:
                        pass

            # bind Escape to exit fullscreen
            try:
                self.bind("<Escape>", _exit_fs)
            except Exception:
                pass

        except Exception as e:
            # do not crash — log only
            logger.exception("Failed to enter fullscreen", exc_info=e)

# ------------------ Modify Base VID (unchanged core) ------------------
def modify_base_vid(new_base: int):
    conn = get_conn();
    cur = conn.cursor()
    try:
        conn.execute("BEGIN IMMEDIATE")
        cur.execute("SELECT MIN(CAST(voucher_id AS INTEGER)) FROM vouchers")
        row = cur.fetchone()
        if not row or row[0] is None:
            _set_setting(cur, "base_vid", new_base)
            conn.commit();
            return 0

        current_min = int(row[0]);
        delta = int(new_base) - current_min
        if delta == 0:
            _set_setting(cur, "base_vid", new_base)
            conn.commit();
            return 0

        order = "DESC" if delta > 0 else "ASC"
        cur.execute(f"""
            SELECT voucher_id, created_at, customer_name, contact_number, units,
                   particulars, problem, staff_name, status, recipient, solution, pdf_path,
                   technician_id, technician_name
            FROM vouchers
            ORDER BY CAST(voucher_id AS INTEGER) {order}
        """)
        rows = cur.fetchall()

        for (vid, created_at, customer_name, contact_number, units, particulars, problem,
             staff_name, status, recipient, solution, old_pdf, tech_id, tech_name) in rows:
            old_id = int(vid);
            new_id = old_id + delta
            try:
                if old_pdf and os.path.exists(old_pdf): os.remove(old_pdf)
            except Exception as e:
                logger.exception("Caught exception", exc_info=e)
                pass
            cur.execute("UPDATE vouchers SET voucher_id=? WHERE voucher_id=?", (str(new_id), str(old_id)))
            new_pdf = generate_pdf(str(new_id), customer_name, contact_number, int(units or 1),
                                   particulars, problem, staff_name, status, created_at, recipient)
            cur.execute("UPDATE vouchers SET pdf_path=? WHERE voucher_id=?", (new_pdf, str(new_id)))

        _set_setting(cur, "base_vid", new_base)
        conn.commit();
        return delta
    except Exception:
        conn.rollback();
        raise
    finally:
        conn.close()


# ------------------ Run ------------------
if __name__ == "__main__":
    init_db()
    app = VoucherApp()
    app.mainloop()
