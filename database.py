import sqlite3
import time
import re
import os
from datetime import datetime
from config import DB_FILE, DEFAULT_BASE_VID, DEFAULT_ADMIN_USERNAME, DEFAULT_ADMIN_PASSWORD, logger

def get_conn():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous = NORMAL")
    conn.row_factory = sqlite3.Row
    return conn

def retry_db_operation(callable_fn, retries=5, base_delay=0.08):
    last_exc = None
    for attempt in range(retries):
        try:
            return callable_fn()
        except sqlite3.OperationalError as e:
            last_exc = e
            msg = str(e).lower()
            if "database is locked" in msg or "database is busy" in msg:
                if attempt < retries - 1:
                    time.sleep(base_delay * (1 + attempt * 0.7))
                    continue
            raise
    if last_exc:
        raise last_exc

def init_db():
    conn = None
    try:
        conn = get_conn()
        cur = conn.cursor()

        # Base Tables
        cur.executescript("""
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
            updated_at TEXT,
            voucher_id TEXT,
            note TEXT
        );
        """)
        conn.commit()
        
        # --- Migrations & Updates ---
        # Ensure commission columns exist
        cur.execute("PRAGMA table_info('commissions')")
        c_cols = [r[1] for r in cur.fetchall()]
        if "voucher_id" not in c_cols:
            cur.execute("ALTER TABLE commissions ADD COLUMN voucher_id TEXT")
        if "note" not in c_cols:
            cur.execute("ALTER TABLE commissions ADD COLUMN note TEXT")

        # Ensure voucher columns exist
        voucher_extras = [
            ("technician_id", "TEXT"), ("technician_name", "TEXT"),
            ("ref_bill", "TEXT"), ("ref_bill_date", "TEXT"),
            ("amount_rm", "REAL"), ("tech_commission", "REAL"),
            ("reminder_pickup_1", "TEXT"), ("reminder_pickup_2", "TEXT"), ("reminder_pickup_3", "TEXT")
        ]
        cur.execute("PRAGMA table_info('vouchers')")
        v_cols = [r[1] for r in cur.fetchall()]
        for col, typ in voucher_extras:
            if col not in v_cols:
                try:
                    cur.execute(f"ALTER TABLE vouchers ADD COLUMN {col} {typ}")
                except Exception:
                    pass

        # Indexes
        cur.execute("CREATE INDEX IF NOT EXISTS idx_vouchers_created ON vouchers(created_at)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_vouchers_status ON vouchers(status)")
        
        # Default Admin
        from auth import hash_pwd
        cur.execute("SELECT COUNT(*) FROM users WHERE role='admin'")
        if cur.fetchone()[0] == 0:
            # Note: auth.py imports need careful handling, importing locally to avoid circle
            ts = datetime.now().isoformat(sep=' ', timespec='seconds')
            cur.execute(
                "INSERT OR IGNORE INTO users (username, role, password_hash, must_change_pwd, created_at, updated_at) VALUES (?,?,?,?,?,?)",
                ("tonycom", "admin", hash_pwd("admin123"), 1, ts, ts)
            )
            conn.commit()

        return conn
    except Exception as e:
        logger.exception("init_db failed")
        if conn: conn.close()
        raise

# --- Helper Queries ---

def search_vouchers(filters):
    conn = get_conn()
    cur = conn.cursor()
    sql = (
        "SELECT voucher_id, created_at, customer_name, contact_number, units, "
        "recipient, technician_id, technician_name, status, solution, pdf_path "
        "FROM vouchers WHERE 1=1"
    )
    params = []

    if filters.get("voucher_id"):
        sql += " AND voucher_id LIKE ?"
        params.append(f"%{filters['voucher_id']}%")
    if filters.get("customer_name"):
        sql += " AND LOWER(customer_name) LIKE ?"
        params.append(f"%{filters['customer_name'].lower()}%")
    if filters.get("contact_number"):
        sql += " AND contact_number LIKE ?"
        params.append(f"%{filters['contact_number']}%")
    
    status = filters.get("status")
    if status and status != "All":
        sql += " AND status = ?"
        params.append(status)

    # Simple date handling
    if filters.get("date_from"):
        # Assuming input is DD-MM-YYYY, convert to YYYY-MM-DD
        try:
            d = datetime.strptime(filters["date_from"], "%d-%m-%Y")
            sql += " AND DATE(created_at) >= DATE(?)"
            params.append(d.strftime("%Y-%m-%d"))
        except: pass
    if filters.get("date_to"):
        try:
            d = datetime.strptime(filters["date_to"], "%d-%m-%Y")
            sql += " AND DATE(created_at) <= DATE(?)"
            params.append(d.strftime("%Y-%m-%d"))
        except: pass

    sql += " ORDER BY created_at DESC"
    
    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()
    return rows

def get_next_voucher_id():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT MAX(CAST(voucher_id AS INTEGER)) FROM vouchers")
    row = cur.fetchone()
    if not row or row[0] is None:
        cur.execute("SELECT value FROM settings WHERE key='base_vid'")
        s_row = cur.fetchone()
        base = int(s_row[0]) if s_row else DEFAULT_BASE_VID
        conn.close()
        return str(base)
    nxt = str(int(row[0]) + 1)
    conn.close()
    return nxt

def list_staffs_names():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT name FROM staffs ORDER BY name COLLATE NOCASE ASC")
    rows = [r[0] for r in cur.fetchall()]
    conn.close()
    return rows
