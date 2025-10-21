import os, sys, io, json, zipfile, shutil, sqlite3, webbrowser, re
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import simpledialog
import customtkinter as ctk
import qrcode
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from PIL import Image, ImageOps, ImageTk
import bcrypt
import logging
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


def restart_app():
    """Restart the current app (PyInstaller EXE or python script) with same args."""
    if getattr(sys, "frozen", False):  # packaged exe
        os.execl(sys.executable, sys.executable, *sys.argv[1:])
    else:  # run as script
        os.execl(sys.executable, sys.executable, os.path.abspath(__file__), *sys.argv[1:])


PHONE_RE = re.compile(r"[0-9+()\- ]+")


def normalize_phone(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[^\d+()\- ]", "", s)
    return s[:32]


def validate_password_policy(pw: str) -> str | None:
    if not pw:
        return "Password cannot be empty."
    return None   # ✅ everything else is acceptable

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
    bill_type TEXT CHECK(bill_type IN ('CS','IV')),
    bill_no TEXT,
    total_amount REAL,
    commission_amount REAL,
    bill_image_path TEXT,
    created_at TEXT,
    updated_at TEXT
);
"""


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


def _set_setting(cur, key, value):
    cur.execute("INSERT OR REPLACE INTO settings(key,value) VALUES (?,?)", (key, str(value)))


def _hash_pwd(pwd: str) -> bytes:
    return bcrypt.hashpw(pwd.encode("utf-8"), bcrypt.gensalt())


def _verify_pwd(pwd: str, hp: bytes) -> bool:
    try:
        return bcrypt.checkpw(pwd.encode("utf-8"), hp)
    except Exception:
        return False


def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.executescript(BASE_SCHEMA)

    # --- MIGRATION: ensure users.role allows 'sales assistant' ---
    cur.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='users'")
    row = cur.fetchone()
    if row and "sales assistant" not in (row[0] or ""):
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


    cur.execute("UPDATE vouchers SET status = TRIM(status)")
    conn.commit()

    for col, typ in [
    ("technician_id", "TEXT"),
    ("technician_name", "TEXT"),
    ("ref_bill", "TEXT"),
    ("ref_bill_date", "TEXT"),
    ("amount_rm", "REAL"),
    ("tech_commission", "REAL"),
    ("reminder_pickup_1", "TEXT"),
    ("reminder_pickup_2", "TEXT"),
    ("reminder_pickup_3", "TEXT"),]:

        if not _column_exists(cur, "vouchers", col):
            cur.execute(f"ALTER TABLE vouchers ADD COLUMN {col} {typ}")

    try:
        # --- Indexes (and enforce uniqueness for commissions) ---
        cur.execute("CREATE INDEX IF NOT EXISTS idx_vouchers_created ON vouchers(created_at)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_vouchers_status ON vouchers(status)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_vouchers_customer ON vouchers(customer_name)")

        # --- VOUCHERS: make ref_bill globally unique (case-insensitive) ---
        # 1) Clean any legacy duplicates (keep the highest voucher_id)
        cur.execute("""
            DELETE FROM vouchers
            WHERE ref_bill IS NOT NULL AND ref_bill <> ''
              AND CAST(voucher_id AS INTEGER) NOT IN (
                SELECT MAX(CAST(voucher_id AS INTEGER)) FROM vouchers
                WHERE ref_bill IS NOT NULL AND ref_bill <> ''
                GROUP BY LOWER(ref_bill)
              )
        """)

        # 2) Enforce uniqueness (ignoring case)
        cur.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS idx_vouchers_ref_bill_unique_ci
            ON vouchers(LOWER(ref_bill))
            WHERE ref_bill IS NOT NULL AND ref_bill <> ''
        """)


        # Remove legacy duplicates so the UNIQUE index can be created safely
        # Remove legacy duplicates so the UNIQUE index can be created safely
        cur.execute("""
            DELETE FROM commissions
            WHERE id NOT IN (
                SELECT MAX(id) FROM commissions
                WHERE bill_no IS NOT NULL AND bill_no <> ''
                GROUP BY bill_type, LOWER(bill_no)
            )
        """)

        # Enforce global uniqueness (bill_type + bill_no), ignoring case
        cur.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS idx_comm_unique_global
            ON commissions(bill_type, LOWER(bill_no))
            WHERE bill_no IS NOT NULL AND bill_no <> ''
        """)


    except Exception as e:
        logger.exception("Failed setting up indexes/uniqueness for commissions", exc_info=e)

    _get_setting(cur, "base_vid", DEFAULT_BASE_VID)

    cur.execute("SELECT COUNT(*) FROM users WHERE role='admin'")
    if cur.fetchone()[0] == 0:
        cur.execute(
            "INSERT INTO users (username, role, password_hash, must_change_pwd, created_at, updated_at) VALUES (?,?,?,?,?,?)",
            ("tonycom", "admin", _hash_pwd("admin123"), 0,
             datetime.now().isoformat(sep=' ', timespec='seconds'),
             datetime.now().isoformat(sep=' ', timespec='seconds'))
        )

    conn.commit();
    conn.close()


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
        except Exception:
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

    qr_size = 20 * mm
    try:
        qr_data = f"Voucher:{voucher_id}|Name:{customer_name}|Tel:{contact_number}|Date:{date_str}"
        qr_img = qrcode.make(qr_data)
        qr_x = right - qr_size
        qr_y = max(policies_bottom + 3 * mm, 10 * mm + qr_size)
        c.drawImage(ImageReader(qr_img), qr_x, qr_y - qr_size, qr_size, qr_size)
    except Exception:
        pass

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
    filename = os.path.join(PDF_DIR, f"voucher_{voucher_id}.pdf")
    c = rl_canvas.Canvas(filename, pagesize=A4)
    width, height = A4
    left, right, top_y = 12 * mm, width - 12 * mm, height - 15 * mm
    title_baseline = _draw_header(c, left, right, top_y, voucher_id)
    _draw_datetime_row(c, left, right, title_baseline, created_at)
    top_table = title_baseline - 12 * mm
    bottom_table, left_col_w = _draw_main_table(c, left, right, top_table, customer_name, particulars, units,
                                                contact_number, problem)
    try:
        dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S");
        date_str = dt.strftime("%d-%m-%Y")
    except Exception:
        date_str = created_at[:10]
    _draw_policies_and_signatures(c, left, right, bottom_table, left_col_w, recipient, voucher_id, customer_name,
                                  contact_number, date_str)
    c.showPage();
    c.save()
    return filename


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
    status=None,   # <-- NEW
):
    with get_conn() as conn:
        cur = conn.cursor()
        voucher_id = next_voucher_id()
        created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status = status or "Pending"   # <-- NEW: default if not provided

        # Generate to a temp name first
        final_pdf = os.path.join(PDF_DIR, f"voucher_{voucher_id}.pdf")
        temp_pdf  = final_pdf + ".part"
        try:
            # generate to temp
            c = rl_canvas.Canvas(temp_pdf, pagesize=A4)
            c.showPage(); c.save()
            os.remove(temp_pdf)
            pdf_path = generate_pdf(
                voucher_id, customer_name, contact_number, units,
                particulars, problem, staff_name, status, created_at, recipient
            )
            if pdf_path != final_pdf and os.path.exists(pdf_path):
                final_pdf = pdf_path
        except Exception as e:
            raise

        # Insert DB row
        try:
            cur.execute("""
                INSERT INTO vouchers (voucher_id, created_at, customer_name, contact_number, units,
                                      particulars, problem, staff_name, status, recipient, solution, pdf_path,
                                      technician_id, technician_name)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (voucher_id, created_at, customer_name, contact_number, units, particulars, problem, staff_name,
                  status, recipient, solution, final_pdf, technician_id, technician_name))
        except Exception:
            try:
                if os.path.exists(final_pdf): os.remove(final_pdf)
            except Exception:
                pass
            raise
        return voucher_id, final_pdf



def update_voucher_fields(voucher_id, **fields):
    if not fields: return
    if "status" in fields: _validate_status_or_raise(fields["status"])
    with get_conn() as conn:
        cur = conn.cursor()
        cols, params = [], []
        for k, v in fields.items():
            cols.append(f"{k}=?");
            params.append(v)
        params.append(voucher_id)
        cur.execute(f"UPDATE vouchers SET {', '.join(cols)} WHERE voucher_id = ?", params)


def bulk_update_status(voucher_ids, new_status):
    _validate_status_or_raise(new_status)
    if not voucher_ids: return
    conn = get_conn();
    cur = conn.cursor()
    cur.executemany("UPDATE vouchers SET status=? WHERE voucher_id=?", [(new_status, vid) for vid in voucher_ids])
    conn.commit();
    conn.close()


def _build_search_sql(filters):
    sql = ("SELECT voucher_id, created_at, customer_name, contact_number, units, "
           "recipient, technician_id, technician_name, status, solution, pdf_path "
           "FROM vouchers WHERE 1=1")
    params = []

    if filters.get("voucher_id"):
        sql += " AND voucher_id LIKE ?"; params.append(f"%{filters['voucher_id']}%")
    if filters.get("name"):
        sql += " AND customer_name LIKE ? COLLATE NOCASE"; params.append(f"%{filters['name']}%")
    if filters.get("contact"):
        sql += " AND contact_number LIKE ?"; params.append(f"%{filters['contact']}%")

    df = filters.get("date_from"); dt = filters.get("date_to")
    if df: sql += " AND created_at >= ?"; params.append(df + " 00:00:00")
    if dt: sql += " AND created_at <= ?"; params.append(dt + " 23:59:59")

    # ---- robust status filter (ignores spaces/case, tolerates typos like "Pendina") ----
    st = (filters.get("status") or "").strip()
    if st and st.lower() != "all":
        # Prefer prefix match (so 'Pending', 'Pendina', 'pending ' all match).
        sql += " AND REPLACE(LOWER(IFNULL(status,'')),' ','') LIKE ?"
        params.append(re.sub(r"\s+", "", st).lower()[:5] + "%")   # 'Pending' -> 'pend%' , 'Completed' -> 'compl%'

    sql += " ORDER BY CAST(voucher_id AS INTEGER) DESC"
    return sql, params




def search_vouchers(filters):
    conn = get_conn();
    cur = conn.cursor()
    sql, params = _build_search_sql(filters)
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


def create_user(username, role, password):
    if role not in ("sales assistant", "user", "admin", "technician"):
        raise ValueError("Invalid role")

    if role == "admin":
        conn = get_conn();
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM users WHERE role='admin'");
        n = cur.fetchone()[0]
        conn.close()
        if n >= 1: raise ValueError("Only one admin allowed")
    conn = get_conn();
    cur = conn.cursor()
    err = validate_password_policy(password)
    if err: raise ValueError(err)
    cur.execute("""INSERT INTO users (username, role, password_hash, created_at, updated_at)
                   VALUES (?,?,?,?,?)""",
                (username, role, _hash_pwd(password),
                 datetime.now().isoformat(sep=" ", timespec="seconds"),
                 datetime.now().isoformat(sep=" ", timespec="seconds")))
    conn.commit();
    conn.close()


def update_user(user_id, **fields):
    if not fields: return
    if "role" in fields and fields["role"] == "admin":
        conn = get_conn();
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM users WHERE role='admin' AND id<>?", (user_id,));
        n = cur.fetchone()[0]
        conn.close()
        if n >= 1: raise ValueError("Only one admin allowed")
    conn = get_conn();
    cur = conn.cursor()
    cols, params = [], []
    for k, v in fields.items():
        cols.append(f"{k}=?");
        params.append(v)
    params.append(user_id)
    cur.execute(f"UPDATE users SET {', '.join(cols)}, updated_at=? WHERE id=?",
                params[:-1] + [datetime.now().isoformat(sep=" ", timespec="seconds"), user_id])
    conn.commit();
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
BILL_RE_IV = re.compile(r"^IV-(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])/\d{4}$")

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
    if BILL_RE_IV.match(s):
        return "IV", s
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
            def _validate_password(pw: str) -> str | None:
                if len(pw) < 10: return "Password must be at least 10 characters."
                if not re.search(r"[A-Z]", pw): return "Include at least one uppercase letter."
                if not re.search(r"[a-z]", pw): return "Include at least one lowercase letter."
                if not re.search(r"\d", pw): return "Include at least one digit."
                if not re.search(r"[^\w\s]", pw): return "Include at least one symbol."
                return None

            new1 = simpledialog.askstring("Change Password", "Enter new password:", show="•", parent=self)
            if not new1: return
            new2 = simpledialog.askstring("Change Password", "Confirm new password:", show="•", parent=self)
            if new1 != new2:
                messagebox.showerror("Password", "Passwords do not match.")
                return
            err = _validate_password(new1)
            if err:
                messagebox.showerror("Password", err)
                return

            reset_password(uid, new1)  # clears must_change flag
            messagebox.showinfo("Password Changed",
                                "Your password has been changed.\nThe application will restart now.")
            try:
                self.grab_release()
            except Exception:
                logger.exception("Failed to release grab during restart", exc_info=True)
            try:
                if self.master is not None:
                    self.master.destroy()
            except Exception:
                pass
            try:
                self.destroy()
            except Exception:
                pass
            restart_app()
            return

        # Success path: set result and close dialog
        self.result = {"id": uid, "username": username, "role": role}
        self.destroy()


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
    def _copy_selected_row(self, tree):
        sel = tree.selection()
        if not sel:
            return
        vals = tree.item(sel[0])["values"]
        try:
            self.clipboard_clear()
            self.clipboard_append("\t".join(str(v) for v in vals))
        except Exception as e:
            messagebox.showerror("Copy", f"Failed to copy:\n{e}")

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
                except Exception:
                    pass
        except Exception:
            pass

        self.current_user = None
        self._do_login_flow()
        self._build_menubar()

        root = ctk.CTkFrame(self)
        root.pack(fill="both", expand=True)

        self._build_filters(root)
        self._build_table(root)
        self._build_bottom_bar(root)

        self.after(80, self._go_fullscreen)
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

        self.menu_comm = tk.Menu(self.menu, tearoff=0)
        self.menu_comm.add_command(label="Add Commission", command=self.add_commission)
        self.menu_comm.add_command(label="View/Edit Commissions", command=self.view_commissions)
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
        self.ctx = tk.Menu(tree, tearoff=0)
        self.ctx.add_command(label="Edit", command=self.edit_selected)
        mark = tk.Menu(self.ctx, tearoff=0)
        for s in ["Pending", "Completed", "Deleted", "1st call", "2nd reminder", "3rd reminder"]:
            mark.add_command(label=s, command=lambda st=s: self._bulk_mark(st))
        self.ctx.add_cascade(label="Mark", menu=mark)
        self.ctx.add_separator()
        self.ctx.add_command(label="Open PDF", command=self.open_pdf)
        self.ctx.add_command(label="Regenerate PDF", command=self.regen_pdf_selected)
        self.ctx.add_separator()
        self.ctx.add_command(
            label="Unmark Completed → Pending", command=lambda: self._bulk_mark("Pending")
        )

        def _popup(event):
            try:
                row = tree.identify_row(event.y)
                if row and row not in tree.selection():
                    tree.selection_set(row)
                self.ctx.post(event.x_root, event.y_root)
            finally:
                self.ctx.grab_release()

        tree.bind("<Button-3>", _popup)

    def _get_filters(self):
        return {
            "voucher_id": self.f_voucher.get().strip(),
            "name": self.f_name.get().strip(),
            "contact": self.f_contact.get().strip(),
            "date_from": _from_ui_date_to_sqldate(self.f_from.get().strip()),
            "date_to": _from_ui_date_to_sqldate(self.f_to.get().strip()),
            "status": (self.f_status.get() or "").strip(),  # trim
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
        for (
                vid,
                created_at,
                customer,
                contact,
                units,
                recipient,
                tech_id,
                tech_name,
                status,
                solution,
                pdf,
        ) in rows:
            tags = [status]
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
                    recipient or "",
                    tech_id or "",
                    tech_name or "",
                    status,
                    solution or "",
                    pdf or "",
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
        except Exception:
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
        """Create Voucher window — single-instance, fixed size, no horizontal scrollbars,
        shortened entry widths, and textboxes wrapped to remove horizontal sliders.
        """
        # --- single-instance guard ---
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

        # --- create toplevel window ---
        top = ctk.CTkToplevel(self)
        top.title("Create Voucher")
        # fixed size (adjust numbers if you want a different constant)
        TOP_WIDTH, TOP_HEIGHT = 720, 620
        top.geometry(f"{TOP_WIDTH}x{TOP_HEIGHT}")
        top.resizable(False, False)
        top.grab_set()

        # register window for single-instance behavior
        self._open_windows["add_voucher_ui"] = top
        def _on_close():
            try:
                del self._open_windows["add_voucher_ui"]
            except Exception:
                pass
            try:
                top.destroy()
            except Exception:
                pass
        try:
            top.protocol("WM_DELETE_WINDOW", _on_close)
        except Exception:
            pass
    
        # --- main frame and grid layout ---
        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=12, pady=8)
        # columns: 0 -> labels (auto), 1 -> inputs (expand)
        frm.grid_columnconfigure(0, weight=0)
        frm.grid_columnconfigure(1, weight=1)
    
        # Common widths (status width used as baseline)
        STATUS_WIDTH = 240   # chosen baseline; adjust if you prefer different
        ENTRY_SHORT = STATUS_WIDTH  # customer/contact/recipient/technician
        SMALL_WIDTH = 120
    
        # --- helpers ---
        def mk_text(parent, height_lines=4):
            """Create a wrapped Text with vertical scrollbar only (no horizontal)."""
            wrap_frame = ctk.CTkFrame(parent)
            wrap_frame.grid_columnconfigure(0, weight=1)
            # use tk.Text for multi-line and control wrap
            txt = tk.Text(wrap_frame, height=height_lines, font=(FONT_FAMILY, UI_FONT_SIZE), wrap='word')
            vsb = ttk.Scrollbar(wrap_frame, orient="vertical", command=txt.yview)
            txt.configure(yscrollcommand=vsb.set)
            txt.grid(row=0, column=0, sticky="nsew")
            vsb.grid(row=0, column=1, sticky="ns")
            return wrap_frame, txt
    
        r = 0
    
        # Customer Name
        ctk.CTkLabel(frm, text="Customer Name").grid(row=r, column=0, sticky="w", padx=(4,8), pady=6)
        e_name = ctk.CTkEntry(frm, width=ENTRY_SHORT)
        e_name.grid(row=r, column=1, sticky="w", padx=(0,12), pady=6)
        r += 1
    
        # Contact Number
        ctk.CTkLabel(frm, text="Contact Number").grid(row=r, column=0, sticky="w", padx=(4,8), pady=6)
        e_contact = ctk.CTkEntry(frm, width=ENTRY_SHORT)
        e_contact.grid(row=r, column=1, sticky="w", padx=(0,12), pady=6)
        r += 1

        # No. of Units (kept small)
        ctk.CTkLabel(frm, text="No. of Units").grid(row=r, column=0, sticky="w", padx=(4,8), pady=6)
        e_units = ctk.CTkEntry(frm, width=SMALL_WIDTH)
        e_units.insert(0, "1")
        e_units.grid(row=r, column=1, sticky="w", padx=(0,12), pady=6)
        r += 1
    
        # Particulars (multi-line) — match length with recipient (uses available column width)
        ctk.CTkLabel(frm, text="Particulars").grid(row=r, column=0, sticky="nw", padx=(4,8), pady=6)
        part_wrap, t_part = mk_text(frm, height_lines=3)
        part_wrap.grid(row=r, column=1, sticky="nsew", padx=(0,12), pady=6)
        r += 1
    
        # Problem (multi-line)
        ctk.CTkLabel(frm, text="Problem").grid(row=r, column=0, sticky="nw", padx=(4,8), pady=6)
        prob_wrap, t_prob = mk_text(frm, height_lines=3)
        prob_wrap.grid(row=r, column=1, sticky="nsew", padx=(0,12), pady=6)
        r += 1
    
        # Recipient (combo)
        ctk.CTkLabel(frm, text="Recipient").grid(row=r, column=0, sticky="w", padx=(4,8), pady=6)
        # list_staffs_names() should return list of names; keep a safe default
        try:
            staff_values = list_staffs_names()
        except Exception:
            staff_values = []
        staff_values = (["— Select —"] + staff_values) if staff_values else ["— Select —"]
        e_recipient = ctk.CTkComboBox(frm, values=staff_values, width=ENTRY_SHORT)
        e_recipient.set(staff_values[0])
        e_recipient.grid(row=r, column=1, sticky="w", padx=(0,12), pady=6)
        r += 1
    
        # Technician (in charge) (combo)
        ctk.CTkLabel(frm, text="Technician (in charge)").grid(row=r, column=0, sticky="w", padx=(4,8), pady=6)
        # populate tech list from DB (keep robust)
        tech_values = ["— Select —"]
        tech_map = {}
        try:
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("SELECT staff_id_opt, name FROM staffs ORDER BY name COLLATE NOCASE")
            tech_rows = cur.fetchall()
            conn.close()
            for sid, nm in tech_rows:
                label = f"{sid}: {nm}" if sid else nm
                tech_values.append(label)
                tech_map[label] = (sid, nm)
        except Exception:
            tech_values = ["— Select —"]
        cb_tech = ctk.CTkComboBox(frm, values=tech_values, width=ENTRY_SHORT)
        cb_tech.set(tech_values[0])
        cb_tech.grid(row=r, column=1, sticky="w", padx=(0,12), pady=6)
        r += 1
    
        # Status (combo) — baseline width
        ctk.CTkLabel(frm, text="Status").grid(row=r, column=0, sticky="w", padx=(4,8), pady=6)
        status_choices = ["Pending", "Completed", "Deleted", "1st call", "2nd reminder", "3rd reminder"]
        cb_status = ctk.CTkComboBox(frm, values=status_choices, width=STATUS_WIDTH)
        cb_status.set("Pending")
        cb_status.grid(row=r, column=1, sticky="w", padx=(0,12), pady=6)
        r += 1
    
        # Solution (multi-line)
        ctk.CTkLabel(frm, text="Solution").grid(row=r, column=0, sticky="nw", padx=(4,8), pady=6)
        sol_wrap, t_sol = mk_text(frm, height_lines=3)
        sol_wrap.grid(row=r, column=1, sticky="nsew", padx=(0,12), pady=6)
        r += 1
    
        # Buttons frame (bottom) — ensure visible and aligned
        btns = ctk.CTkFrame(top)
        btns.pack(fill="x", padx=12, pady=(6, 12))
        # Save handler
        def save():
            name = e_name.get().strip()
            contact = e_contact.get().strip()
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
            solution = t_sol.get("1.0", "end").strip()
            sel = cb_tech.get().strip()
            if sel in ("— Select —", ""):
                messagebox.showerror("Missing", "Please select a Technician in charge.")
                return
            technician_id, technician_name = tech_map.get(sel, ("", ""))
            if not name or not contact:
                messagebox.showerror("Missing", "Customer name and contact are required.")
                return

            # call your existing add_voucher function (preserves current DB logic)
            voucher_id, _ = add_voucher(
                name, contact, units,
                particulars, problem,
                recipient, recipient,
                solution, technician_id, technician_name,
                status=cb_status.get().strip()
            )
            messagebox.showinfo("Saved", f"Voucher {voucher_id} created.")
            # close and refresh search/list in parent
            try:
                _on_close()
            except Exception:
                try: top.destroy()
                except Exception: pass
            try:
                self.perform_search()
            except Exception:
                pass
    
        # Right-aligned Save button
        white_btn(btns, text="Save", command=save, width=140).pack(side="right")
        # Optional Cancel button
        white_btn(btns, text="Cancel", command=_on_close, width=100).pack(side="right", padx=(0,8))
    
        # Final layout tweaks: ensure the frame expands properly and no widget is cropped
        try:
            # force initial geometry to ensure widgets compute correct sizes (helps some platforms)
            top.update_idletasks()
        except Exception:
            pass
    
    def edit_selected(self):
        sels = self.tree.selection()
        if len(sels) != 1:
            messagebox.showerror("Edit", "Select exactly one record to edit.")
            return
        voucher_id = self.tree.item(sels[0])["values"][0]

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
        e_name.insert(0, customer_name)
        e_name.grid(row=r, column=1, sticky="ew", padx=10, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="Contact Number").grid(row=r, column=0, sticky="w")
        e_contact = ctk.CTkEntry(frm, width=WIDE)
        e_contact.insert(0, contact_number)
        e_contact.grid(row=r, column=1, sticky="ew", padx=10, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="No. of Units").grid(row=r, column=0, sticky="w")
        e_units = ctk.CTkEntry(frm, width=120)
        e_units.insert(0, str(units))
        e_units.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        def mk_text(parent, height, seed=""):
            wrap = ctk.CTkFrame(parent)
            wrap.grid_columnconfigure(0, weight=1)
            txt = tk.Text(wrap, height=height, font=(FONT_FAMILY, UI_FONT_SIZE))
            txt.insert("1.0", seed)
            sb = ttk.Scrollbar(wrap, orient="vertical", command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            txt.grid(row=0, column=0, sticky="nsew")
            sb.grid(row=0, column=1, sticky="ns")
            return wrap, txt

        ctk.CTkLabel(frm, text="Particulars").grid(row=r, column=0, sticky="w")
        part_wrap, t_part = mk_text(frm, 3, particulars or "")
        part_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="Problem").grid(row=r, column=0, sticky="w")
        prob_wrap, t_prob = mk_text(frm, 4, problem or "")
        prob_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="Recipient").grid(row=r, column=0, sticky="w")
        staff_values = list_staffs_names() or [""]
        e_recipient = ctk.CTkComboBox(frm, values=staff_values, width=WIDE)
        e_recipient.set(recipient or staff_name or "")
        e_recipient.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        # --- Technician (in charge) ---
        ctk.CTkLabel(frm, text="Technician (in charge)").grid(row=r, column=0, sticky="w")

        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT staff_id_opt, name FROM staffs ORDER BY name COLLATE NOCASE")
        tech_rows = cur.fetchall()
        conn.close()

        tech_values = ["— Select —"] + [f"{sid}: {nm}" if sid else nm for sid, nm in tech_rows]
        tech_map = {f"{sid}: {nm}" if sid else nm: (sid, nm) for sid, nm in tech_rows}

        cb_tech = ctk.CTkComboBox(frm, values=tech_values, width=WIDE)

        preset_val = None
        if tech_id and tech_name:
            preset_val = f"{tech_id}: {tech_name}"
        elif tech_name:
            preset_val = tech_name

        if preset_val in tech_values:
            cb_tech.set(preset_val)
        else:
            cb_tech.set(tech_values[0])

        cb_tech.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        ctk.CTkLabel(frm, text="Solution").grid(row=r, column=0, sticky="w")
        sol_wrap, t_sol = mk_text(frm, 4, solution or "")
        sol_wrap.grid(row=r, column=1, sticky="nsew", padx=10, pady=6)
        r += 1

        # --- Referring to Cash/Inv bill + Date (side-by-side) ---
        ctk.CTkLabel(frm, text="Refering to Cash/Inv bill").grid(row=r, column=0, sticky="w")

        bill_frame = ctk.CTkFrame(frm)
        bill_frame.grid(row=r, column=1, sticky="ew", padx=10, pady=6)
        bill_frame.grid_columnconfigure(0, weight=1)
        bill_frame.grid_columnconfigure(1, weight=0)

        e_ref_bill = ctk.CTkEntry(bill_frame, width=320)
        e_ref_bill.insert(0, ref_bill or "")
        e_ref_bill.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        ctk.CTkLabel(bill_frame, text="Date:").grid(row=0, column=1, sticky="w", padx=(4, 6))
        e_ref_date = ctk.CTkEntry(bill_frame, width=150)
        e_ref_date.insert(0, ref_bill_date or "")
        e_ref_date.grid(row=0, column=2, sticky="w")
        r += 1

        # --- Amount and Commission (closer together, same row) ---
        ctk.CTkLabel(frm, text="Amount (RM)").grid(row=r, column=0, sticky="w")
        amount_comm_frame = ctk.CTkFrame(frm, fg_color="transparent")
        amount_comm_frame.grid(row=r, column=1, columnspan=3, sticky="w", padx=10, pady=6)

        e_amount = ctk.CTkEntry(amount_comm_frame, width=120)
        e_amount.pack(side="left", padx=(0, 10))  # small 10px gap between amount and commission

        ctk.CTkLabel(amount_comm_frame, text="Technician Commission").pack(side="left", padx=(0, 6))
        e_comm = ctk.CTkEntry(amount_comm_frame, width=120)
        e_comm.pack(side="left")
        r += 1

        # --- Reminder Customer pick up (three dates on one row) ---
        ctk.CTkLabel(frm, text="Reminder Customer pick up (Date)").grid(row=r, column=0, sticky="w")

        rem_frame = ctk.CTkFrame(frm)
        rem_frame.grid(row=r, column=1, sticky="ew", padx=10, pady=6)
        for i in range(6): rem_frame.grid_columnconfigure(i, weight=0)

        ctk.CTkLabel(rem_frame, text="1:").grid(row=0, column=0, sticky="w", padx=(0, 4))
        e_rem1 = ctk.CTkEntry(rem_frame, width=130);
        e_rem1.insert(0, reminder1 or "")
        e_rem1.grid(row=0, column=1, sticky="w", padx=(0, 12))

        ctk.CTkLabel(rem_frame, text="2:").grid(row=0, column=2, sticky="w", padx=(0, 4))
        e_rem2 = ctk.CTkEntry(rem_frame, width=130);
        e_rem2.insert(0, reminder2 or "")
        e_rem2.grid(row=0, column=3, sticky="w", padx=(0, 12))

        ctk.CTkLabel(rem_frame, text="3:").grid(row=0, column=4, sticky="w", padx=(0, 4))
        e_rem3 = ctk.CTkEntry(rem_frame, width=130);
        e_rem3.insert(0, reminder3 or "")
        e_rem3.grid(row=0, column=5, sticky="w")
        r += 1

        ctk.CTkLabel(frm, text="Status").grid(row=r, column=0, sticky="w")
        status_choices = ["Pending", "Completed", "Deleted", "1st call", "2nd reminder", "3rd reminder"]
        cb_status = ctk.CTkComboBox(frm, values=status_choices, width=240)
        cb_status.set(status or "Pending")  # you'll set again when loading the record
        cb_status.grid(row=r, column=1, sticky="w", padx=10, pady=6)
        r += 1

        # --- Buttons just under Status (right-aligned) ---
        btns = ctk.CTkFrame(frm, fg_color="transparent")
        btns.grid(row=r, column=0, columnspan=2, sticky="e", padx=12, pady=(4, 8))

        def save_edit():
            name = e_name.get().strip()
            contact = e_contact.get().strip()
            try:
                units_val = int((e_units.get() or "1").strip())
                if units_val <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid", "Units must be a positive integer.")
                return
            particulars_val = t_part.get("1.0", "end").strip()
            problem_val = t_prob.get("1.0", "end").strip()
            recipient_val = e_recipient.get().strip()
            solution_val = t_sol.get("1.0", "end").strip()
            staff_val = recipient_val
            sel = cb_tech.get().strip()
            if sel in ("— Select —", ""):
                messagebox.showerror("Missing", "Please select a Technician in charge.")
                return
            tech_id_val, tech_name_val = tech_map.get(sel, ("", ""))
            ref_bill_val = e_ref_bill.get().strip().upper()  # normalize to uppercase
            ref_date_val = e_ref_date.get().strip()
            amount_val = e_amount.get().strip()
            tech_comm_val = e_comm.get().strip()
            rem1_val = e_rem1.get().strip()
            rem2_val = e_rem2.get().strip()
            rem3_val = e_rem3.get().strip()

            # --- Duplicate check (case-insensitive) ---
            if ref_bill_val:
                conn_chk = get_conn()
                cur_chk = conn_chk.cursor()
                cur_chk.execute(
                    "SELECT voucher_id FROM vouchers WHERE LOWER(ref_bill) = LOWER(?) AND voucher_id <> ?",
                    (ref_bill_val, str(voucher_id)),
                )
                row_dup = cur_chk.fetchone()
                conn_chk.close()
                if row_dup:
                    messagebox.showerror(
                        "Duplicate Bill",
                        f"The bill '{ref_bill_val}' is already used by voucher {row_dup[0]}.\n"
                        "Please use a unique Cash/Inv bill number."
                    )
                    return

            def _num_or_none(s):
                try:
                    return float(s) if s != "" else None
                except:
                    return None

            # Try update, handle possible DB uniqueness failure gracefully
            try:
                update_voucher_fields(
                    voucher_id,
                    customer_name=name,
                    contact_number=contact,
                    units=units_val,
                    particulars=particulars_val,
                    problem=problem_val,
                    staff_name=staff_val,
                    recipient=recipient_val,
                    solution=solution_val,
                    technician_id=tech_id_val,
                    technician_name=tech_name_val,
                    status=cb_status.get().strip(),
                    ref_bill=ref_bill_val,
                    ref_bill_date=ref_date_val,
                    amount_rm=_num_or_none(amount_val),
                    tech_commission=_num_or_none(tech_comm_val),
                    reminder_pickup_1=rem1_val,
                    reminder_pickup_2=rem2_val,
                    reminder_pickup_3=rem3_val,
                )
            except sqlite3.IntegrityError:
                messagebox.showerror(
                    "Duplicate Bill",
                    "This Cash/Inv bill is already linked to another voucher (case-insensitive)."
                )
                return
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update voucher:\n{e}")
                return

            # --- Auto-sync commission into commissions table (View/Edit Commissions) ---
            def _num_or_zero(s):
                try:
                    return float(s) if s != "" else 0.0
                except:
                    return 0.0

            bill_type, bill_no = _parse_bill(ref_bill_val)
            total_amount = _num_or_zero(amount_val)
            commission_amount = _num_or_zero(tech_comm_val)

            # Only proceed if we have a valid bill AND a technician resolved AND at least one amount filled
            staff_db_id = _resolve_staff_db_id(tech_id_val, tech_name_val)

            if bill_type and bill_no and staff_db_id and (total_amount or commission_amount):
                try:
                    conn_comm = get_conn()
                    cur_comm = conn_comm.cursor()
                    # Upsert into unique (staff_id, bill_type, bill_no)
                    cur_comm.execute(
                        """
                        INSERT INTO commissions
                            (staff_id, bill_type, bill_no, total_amount, commission_amount, bill_image_path, created_at, updated_at)
                        VALUES (?,?,?,?,?,?,?,?)
                        ON CONFLICT(staff_id, bill_type, bill_no) DO UPDATE SET
                            total_amount=excluded.total_amount,
                            commission_amount=excluded.commission_amount,
                            updated_at=excluded.updated_at
                        """,
                        (
                            staff_db_id,
                            bill_type, bill_no,
                            total_amount, commission_amount,
                            "",  # no image from voucher edit
                            datetime.now().isoformat(sep=" ", timespec="seconds"),
                            datetime.now().isoformat(sep=" ", timespec="seconds"),
                        ),
                    )
                    conn_comm.commit()
                except Exception as e:
                    # non-fatal; user still updates voucher even if commission sync fails
                    logger.exception("Auto-sync commission failed for voucher %s: %s", voucher_id, e)
                finally:
                    try:
                        conn_comm.close()
                    except Exception:
                        pass

            # --- regenerate PDF and finish ---
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("SELECT created_at, status, pdf_path FROM vouchers WHERE voucher_id=?", (voucher_id,))
            created_at2, status2, old_pdf_path = cur.fetchone()
            try:
                if old_pdf_path and os.path.exists(old_pdf_path):
                    os.remove(old_pdf_path)
            except Exception:
                pass

            selected_status = cb_status.get().strip()
            pdf_path = generate_pdf(
                str(voucher_id),
                name,
                contact,
                units_val,
                particulars_val,
                problem_val,
                staff_val,
                selected_status,
                created_at2,
                recipient_val,
            )
            cur.execute("UPDATE vouchers SET pdf_path=? WHERE voucher_id=?", (pdf_path, str(voucher_id)))
            conn.commit()
            conn.close()
            messagebox.showinfo("Updated", f"Voucher {voucher_id} updated.")
            top.destroy()
            self.perform_search()

        white_btn(btns, text="Cancel", command=top.destroy, width=120).pack(side="right", padx=(6, 0))
        white_btn(btns, text="Save Changes", command=save_edit, width=140).pack(side="right")

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
        except Exception:
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
        except Exception:
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
            except Exception:
                # if refresh_users not available yet, ignore silently
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
                        except Exception:
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
            except Exception:
                pass
            autosize_columns()

        # initial population and bind resizing
        refresh_users()
        try:
            top.bind("<Configure>", lambda e: autosize_columns())
        except Exception:
            pass

        # double-click to edit
        try:
            tree.bind("<Double-1>", lambda e: on_edit())
        except Exception:
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
                        create_user(uname, role, pwd)
                        # if must_change, set flag
                        if must:
                            # find user id and set must_change flag
                            for u in list_users():
                                if u[1] == uname:
                                    update_user(u[0], must_change=1)
                                    break
                        messagebox.showinfo("Added", "User created.", parent=dlg)
                    except Exception as e:
                        messagebox.showerror("Add failed", f"{e}", parent=dlg)
                        return
                else:
                    # edit mode: update role and must_change (password change optional)
                    try:
                        update_user(user_id, role=role, must_change=must)
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
        except Exception:
            pass

    # ---------- Staff Profile UI (preview + folder open) ----------
    def staff_profile(self):
        """Staff Profile window:
           - Position and Name inputs share same width
           - Table columns (position, staff ID, name, phone) auto-fit to window
           - Fixed, non-resizable window; single-instance behavior
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

        # changed label text to "Phone Number"
        ctk.CTkLabel(panel, text="Phone Number").grid(row=1, column=2, padx=(6,8), sticky="w", pady=(8,0))
        e_phone = ctk.CTkEntry(panel, width=160)
        e_phone.grid(row=1, column=3, padx=(0,8), sticky="w", pady=(8,0))

        # Add Staff / Delete / Reset buttons (inline handlers)
        btn_panel = ctk.CTkFrame(panel)
        btn_panel.grid(row=0, column=4, rowspan=2, padx=(8,4), sticky="e")

        def on_add_staff():
            name = e_name.get().strip()
            staff_id_opt = e_staffid.get().strip()
            position = cb_pos.get().strip()
            phone = e_phone.get().strip()
            if not name:
                messagebox.showerror("Missing", "Please enter Name.")
                return
            now = datetime.now().isoformat(sep=" ", timespec="seconds")
            try:
                conn = get_conn()
                cur = conn.cursor()
                # Insert if not exists, otherwise update fields for that name
                cur.execute("""
                    INSERT OR IGNORE INTO staffs (position, staff_id_opt, name, phone, photo_path, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (position, staff_id_opt, name, phone, "", now, now))
                # Ensure the record has up-to-date fields (staff_id_opt/position/phone)
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
            # clear inputs and refresh table (refresh() defined later)
            e_staffid.delete(0, "end")
            e_name.delete(0, "end")
            e_phone.delete(0, "end")
            cb_pos.set("Technician")
            try:
                refresh()
            except Exception:
                pass

        def on_delete_staff():
            """Delete staff record. Match by staff_id if provided, else by name.
            If both provided, require both match to avoid accidental deletes."""
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
            # clear inputs and refresh
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

        # Columns: position, staff_id, name, phone
        cols = ("position", "staff_id", "name", "phone")
        display_names = {"position":"Position", "staff_id":"Staff ID", "name":"Name", "phone":"Phone"}

        tree = ttk.Treeview(container, columns=cols, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        # Headings
        for c in cols:
            tree.heading(c, text=display_names.get(c, c.title()))

        # Autosize columns to exactly fill available width
        def autosize_columns():
            try:
                container.update_idletasks()
                padding = 20  # adjust if your UI has extra paddings
                total_w = container.winfo_width() - padding
                if total_w <= 0: total_w = TOP_W - 40
                n = len(cols)

                # Gather sample strings for each column: header + a few rows
                samples = {cid: [str(tree.heading(cid)['text'])] for cid in cols}
                for iid in tree.get_children()[:40]:  # sample up to first 40 rows
                    vals = tree.item(iid)["values"]
                    for idx, cid in enumerate(cols):
                        try:
                            samples[cid].append(str(vals[idx] or ""))
                        except Exception:
                            pass

                # Estimate minimum widths (7 px per char) + margin
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

                # Apply widths
                for cid, w in zip(cols, widths):
                    tree.column(cid, width=w, anchor="w", stretch=False)
            except Exception:
                # fallback: reasonable defaults
                base = int((TOP_W - 80) / len(cols))
                for cid in cols:
                    tree.column(cid, width=base, anchor="w", stretch=True)

        # Populate data (use your DB helper if present)
        def refresh():
            tree.delete(*tree.get_children())
            try:
                conn = get_conn()
                cur = conn.cursor()
                cur.execute("SELECT position, staff_id_opt, name, phone FROM staffs ORDER BY name COLLATE NOCASE")
                rows = cur.fetchall()
                conn.close()
            except Exception:
                # fallback sample data
                rows = [
                    ("Technician", "T001", "Alice Lee", "012-3456789"),
                    ("Sales", "S005", "Bob Tan", "019-888777"),
                ]
            for r in rows:
                # Note: mapping depends on your schema; here we used staff_id_opt for staff ID field
                pos = r[0] if len(r) > 0 else ""
                sid = r[1] if len(r) > 1 else ""
                nm = r[2] if len(r) > 2 else ""
                phone = r[3] if len(r) > 3 else ""
                tree.insert("", "end", values=(pos, sid, nm, phone))
            autosize_columns()

        # binding: double-click to edit
        try:
            tree.bind("<Double-1>", lambda e: self._edit_staff(tree, top))
        except Exception:
            pass

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

    # ---------- Commission UI (with preview + per-staff storage) ----------
    def add_commission(self):
        top.title("Add Commission")
        # reduced height so form fields align horizontally with the Save button
        top.geometry("900x420")
        top.resizable(False, False)
        top.grab_set()
        
        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        for c in range(4):
            frm.grid_columnconfigure(c, weight=1)

        # ---------- Staff combo (display "staff_id_opt:Name", save real staffs.id) ----------
        ctk.CTkLabel(frm, text="Staff").grid(row=0, column=0, sticky="w")

        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT id, name, staff_id_opt FROM staffs ORDER BY name COLLATE NOCASE")
        staff_rows = cur.fetchall()
        conn.close()

        display_list = ["— Select —"]
        staff_map = {}  # display string -> staffs.id
        name_map = {}  # staffs.id -> name (for folder naming)
        for db_id, nm, opt in staff_rows:
            label = f"{opt}:{nm}" if (opt or "").strip() else nm
            display_list.append(label)
            staff_map[label] = db_id
            name_map[db_id] = nm

        cb_staff = ctk.CTkComboBox(frm, values=display_list, width=320)
        cb_staff.set(display_list[0])
        cb_staff.grid(row=0, column=1, sticky="w", padx=8, pady=6)

        # ---------- Other fields ----------
        ctk.CTkLabel(frm, text="Bill Type").grid(row=0, column=2, sticky="w")
        cb_type = ctk.CTkComboBox(frm, values=["Cash Bill", "Invoice Bill"], width=200)
        cb_type.set("Cash Bill")
        cb_type.grid(row=0, column=3, sticky="w", padx=8, pady=6)

        ctk.CTkLabel(frm, text="Bill No.").grid(row=1, column=0, sticky="w")
        e_bill = ctk.CTkEntry(frm)
        e_bill.grid(row=1, column=1, sticky="ew", padx=8, pady=6)

        ctk.CTkLabel(frm, text="Total Amount").grid(row=1, column=2, sticky="w")
        e_total = ctk.CTkEntry(frm)
        e_total.grid(row=1, column=3, sticky="ew", padx=8, pady=6)

        ctk.CTkLabel(frm, text="Commission Amount").grid(row=2, column=0, sticky="w")
        e_comm = ctk.CTkEntry(frm)
        e_comm.grid(row=2, column=1, sticky="ew", padx=8, pady=6)


        # ---------- Validation ----------
        def validate_bill():
            s = (e_bill.get() or "").strip().upper()
            btype = cb_type.get()
            ok = bool(BILL_RE_CS.match(s)) if btype == "Cash Bill" else bool(BILL_RE_IV.match(s))
            if not ok:
                messagebox.showerror(
                    "Bill No.", "Invalid bill number format.\nCash: CS-MMDD/XXXX\nInvoice: IV-MMDD/XXXX"
                )
            return ok

        # ---------- Save ----------
        def save():
            if not staff_rows:
                messagebox.showerror("Commission", "No staff found. Add staff first.")
                return

            selected_display = cb_staff.get().strip()
            if selected_display == "— Select —":
                messagebox.showerror("Commission", "Please select a staff.")
                return

            staff_db_id = staff_map.get(selected_display)
            if not staff_db_id:
                messagebox.showerror("Commission", "Invalid staff selection.")
                return

            if not validate_bill():
                return

            try:
                total = float(e_total.get())
                commission = float(e_comm.get())
            except Exception:
                messagebox.showerror("Commission", "Amounts must be numeric.")
                return

            sname = name_map.get(staff_db_id, "unknown")
            base_dir, com_dir = staff_dirs_for(sname)

            bill_type = "CS" if cb_type.get() == "Cash Bill" else "IV"
            bill_no = (e_bill.get() or "").strip().upper()

            # --- Global duplicate check (case-insensitive) ---
            conn_chk = get_conn()
            cur_chk = conn_chk.cursor()
            cur_chk.execute(
                "SELECT id FROM commissions WHERE LOWER(bill_no) = LOWER(?)",
                (bill_no,)
            )
            dup = cur_chk.fetchone()
            conn_chk.close()
            if dup:
                messagebox.showerror(
                    "Duplicate Bill",
                    f"This bill number ({bill_no}) already exists in the system. Cannot save duplicate."
                )
                return


            # ----- Since image feature removed, ensure img_path exists (empty) -----
            img_path = ""

            conn = get_conn()
            try:
                cur = conn.cursor()
                cur.execute("""
                    INSERT INTO commissions
                        (staff_id, bill_type, bill_no, total_amount, commission_amount, bill_image_path, created_at, updated_at)
                    VALUES (?,?,?,?,?,?,?,?)
                """, (
                    staff_db_id, bill_type, bill_no, total, commission, img_path,
                    datetime.now().isoformat(sep=" ", timespec="seconds"),
                    datetime.now().isoformat(sep=" ", timespec="seconds"),
                ))
                conn.commit()
            except sqlite3.IntegrityError:
                conn.rollback()
                try:
                    if img_path and os.path.exists(img_path):
                        os.remove(img_path)
                except Exception:
                    pass
                messagebox.showerror("Commission",
                                     f"This bill already exists for {sname}:\nType: {bill_type}, No: {bill_no}")
                return
            except Exception as e:
                conn.rollback()
                try:
                    if img_path and os.path.exists(img_path):
                        os.remove(img_path)
                except Exception:
                    pass
                messagebox.showerror("Commission", f"Failed to save commission:\n{e}")
                return
            finally:
                conn.close()

            messagebox.showinfo("Commission", "Saved.")
            top.destroy()

        # place Save immediately below input rows so dialog height can be compact
        white_btn(frm, text="Save", command=save, width=140).grid(row=3, column=3, sticky="e", pady=8)

    def view_commissions(self):
        top = ctk.CTkToplevel(self)
        top.title("Commissions")
        # increase height so the tree, horizontal scrollbar and buttons are visible
        top.geometry("950x720")
        top.resizable(False, False)
        top.grab_set()

        frm = ctk.CTkFrame(top)
        frm.pack(fill="both", expand=True, padx=10, pady=10)
        frm.grid_rowconfigure(1, weight=1)
        frm.grid_columnconfigure(0, weight=1)

        bar = ctk.CTkFrame(frm)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        # Search + Date range controls
        e_search = ctk.CTkEntry(bar, width=260, placeholder_text="Search bill no. or staff name")
        e_search.pack(side="left", padx=(4, 6))

        ctk.CTkLabel(bar, text="Date From").pack(side="left", padx=(6, 2))
        e_date_from = ctk.CTkEntry(bar, width=120, placeholder_text="DD-MM-YYYY")
        e_date_from.pack(side="left", padx=(0, 6))
        ctk.CTkLabel(bar, text="To").pack(side="left", padx=(6, 2))
        e_date_to = ctk.CTkEntry(bar, width=120, placeholder_text="DD-MM-YYYY")
        e_date_to.pack(side="left", padx=(0, 8))

        white_btn(bar, text="Search", width=110, command=lambda: refresh()).pack(side="left", padx=5)

        # Edit button: admin only
        btn_edit = white_btn(bar, text="Edit Selected", width=140, command=lambda: _edit_selected())
        if not (self.current_user and self.current_user.get("role") == "admin"):
            btn_edit.configure(state=tk.DISABLED)
        btn_edit.pack(side="right", padx=5)

        # Delete button: disable for technicians
        btn_delete = white_btn(bar, text="Delete Selected", width=150, command=lambda: delete_sel())
        if self.current_user and self.current_user.get("role") == "technician":
            btn_delete.configure(state=tk.DISABLED)
        btn_delete.pack(side="right", padx=5)

        wrap = ctk.CTkFrame(frm)
        wrap.grid(row=1, column=0, sticky="nsew")
        wrap.grid_rowconfigure(0, weight=1)
        wrap.grid_columnconfigure(0, weight=1)

        cols = ("staff_id", "staff", "type", "billno", "total", "commission", "created")
        tree = ttk.Treeview(wrap, columns=cols, show="headings", selectmode="browse")
        for c, t, w in [
            ("staff_id", "Staff ID", 120),
            ("staff", "Staff", 220),
            ("type", "Type", 80),
            ("billno", "Bill No.", 160),
            ("total", "Total", 110),
            ("commission", "Commission", 120),
            ("created", "Created", 160),
        ]:
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="w", stretch=False)

        tree.grid(row=0, column=0, sticky="nsew")
        sbv = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        sbv.grid(row=0, column=1, sticky="ns")
        sbh = ttk.Scrollbar(wrap, orient="horizontal", command=tree.xview)
        sbh.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=sbv.set, xscrollcommand=sbh.set)
        freeze_tree_columns(tree)

        def refresh():
            q = (e_search.get() or "").strip().lower()
            tree.delete(*tree.get_children())

            # Validate and convert date range
            df_ui = (e_date_from.get() or "").strip()
            dt_ui = (e_date_to.get() or "").strip()
            date_from_sql = _from_ui_date_to_sqldate(df_ui) if df_ui else None
            date_to_sql = _from_ui_date_to_sqldate(dt_ui) if dt_ui else None

            if df_ui and not date_from_sql:
                messagebox.showerror("Invalid Date", "Date From is invalid. Use DD-MM-YYYY.")
                return
            if dt_ui and not date_to_sql:
                messagebox.showerror("Invalid Date", "Date To is invalid. Use DD-MM-YYYY.")
                return

            conn = get_conn()
            cur = conn.cursor()

            # Build optional WHERE for technicians and date range
            where_clauses = []
            params = []

            if self.current_user and self.current_user.get("role") == "technician":
                # try to map user to a staff row
                uname = (self.current_user.get("username") or "").strip()
                cur.execute(
                    "SELECT id FROM staffs WHERE LOWER(name)=LOWER(?) OR LOWER(COALESCE(staff_id_opt,''))=LOWER(?)",
                    (uname, uname),
                )
                m = cur.fetchone()
                if m:
                    where_clauses.append("c.staff_id = ?")
                    params.append(m[0])
                # if not found, we won't filter (so they still see data)

            # Date filters (created_at is ISO 'YYYY-MM-DD HH:MM:SS')
            if date_from_sql:
                where_clauses.append("c.created_at >= ?")
                params.append(date_from_sql + " 00:00:00")
            if date_to_sql:
                where_clauses.append("c.created_at <= ?")
                params.append(date_to_sql + " 23:59:59")

            where_sql = ("WHERE " + " AND ".join(where_clauses)) if where_clauses else ""

            cur.execute(
                f"""
                SELECT
                    c.id,            -- use as iid
                    s.staff_id_opt,  -- visible staff ID
                    s.name,          -- staff name
                    c.bill_type, c.bill_no, c.total_amount, c.commission_amount, c.created_at
                FROM commissions c
                JOIN staffs s ON c.staff_id = s.id
                {where_sql}
                ORDER BY c.id DESC
                """,
                params
            )
            rows = cur.fetchall()
            conn.close()

            for (cid, staff_opt, staff_name, bill_type, bill_no, total, commission, created_at) in rows:
                if q and (q not in str(staff_name).lower() and q not in str(bill_no).lower()):
                    continue
                tree.insert(
                    "", "end", iid=str(cid),
                    values=(staff_opt or "", staff_name, bill_type, bill_no, total, commission, created_at)
                )

        def delete_sel():
            # Extra guard: technicians cannot delete even if button is re-enabled somehow
            if self.current_user and self.current_user.get("role") == "technician":
                messagebox.showerror("Permission", "Technician role cannot delete commissions.")
                return

            sel = tree.selection()
            if not sel:
                return
            cid = int(sel[0])
            if not messagebox.askyesno("Delete", "Delete this commission record? This cannot be undone."):
                return
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("DELETE FROM commissions WHERE id=?", (cid,))
            conn.commit()
            conn.close()
            refresh()

        # Local edit function so it can call refresh()
        def _edit_selected():
            # Permission guard
            if not (self.current_user and self.current_user.get("role") == "admin"):
                messagebox.showerror("Permission", "Only admin can edit commissions.")
                return

            sel = tree.selection()
            if not sel:
                messagebox.showinfo("Edit", "Select a commission to edit.")
                return
            cid = int(sel[0])

            # Load commission
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("""
                SELECT c.id, c.staff_id, s.name, c.bill_type, c.bill_no, c.total_amount, c.commission_amount, c.created_at
                FROM commissions c
                JOIN staffs s ON c.staff_id = s.id
                WHERE c.id = ?
            """, (cid,))
            row = cur.fetchone()
            conn.close()

            if not row:
                messagebox.showerror("Edit", "Commission record not found.")
                return

            _id, staff_id, staff_name, bill_type, bill_no, total_amount, commission_amount, created_at = row

            # Build edit dialog
            etop = ctk.CTkToplevel(self)
            etop.title(f"Edit Commission {cid}")
            # a bit taller so Save/Cancel row and fields are not cut off
            etop.geometry("900x520")
            etop.resizable(False, False)
            etop.grab_set()

            efrm = ctk.CTkFrame(etop)
            efrm.pack(fill="both", expand=True, padx=12, pady=12)
            for c in range(4):
                efrm.grid_columnconfigure(c, weight=1)

            # Staff (readonly)
            ctk.CTkLabel(efrm, text="Staff").grid(row=0, column=0, sticky="w")
            lbl_staff = ctk.CTkLabel(efrm, text=staff_name)
            lbl_staff.grid(row=0, column=1, sticky="w", padx=8, pady=6)

            # Bill Type
            ctk.CTkLabel(efrm, text="Bill Type").grid(row=0, column=2, sticky="w")
            cb_type = ctk.CTkComboBox(efrm, values=["Cash Bill", "Invoice Bill"], width=200)
            cb_type.set("Cash Bill" if bill_type == "CS" else "Invoice Bill")
            cb_type.grid(row=0, column=3, sticky="w", padx=8, pady=6)

            # Bill No
            ctk.CTkLabel(efrm, text="Bill No.").grid(row=1, column=0, sticky="w")
            e_bill = ctk.CTkEntry(efrm)
            e_bill.insert(0, bill_no or "")
            e_bill.grid(row=1, column=1, sticky="ew", padx=8, pady=6)

            # Total Amount
            ctk.CTkLabel(efrm, text="Total Amount").grid(row=1, column=2, sticky="w")
            e_total = ctk.CTkEntry(efrm)
            e_total.insert(0, str(total_amount or ""))
            e_total.grid(row=1, column=3, sticky="ew", padx=8, pady=6)

            # Commission Amount
            ctk.CTkLabel(efrm, text="Commission Amount").grid(row=2, column=0, sticky="w")
            e_comm = ctk.CTkEntry(efrm)
            e_comm.insert(0, str(commission_amount or ""))
            e_comm.grid(row=2, column=1, sticky="ew", padx=8, pady=6)

            # Created (readonly)
            ctk.CTkLabel(efrm, text="Created").grid(row=2, column=2, sticky="w")
            lbl_created = ctk.CTkLabel(efrm, text=created_at or "")
            lbl_created.grid(row=2, column=3, sticky="w", padx=8, pady=6)

            # Buttons
            btns = ctk.CTkFrame(efrm)
            btns.grid(row=5, column=0, columnspan=4, sticky="e", pady=(18, 0))

            def _validate_bill_format(b):
                s = (b or "").strip().upper()
                bt = cb_type.get()
                if bt == "Cash Bill":
                    return bool(BILL_RE_CS.match(s))
                else:
                    return bool(BILL_RE_IV.match(s))

            def _save_edit():
                new_bill_no = (e_bill.get() or "").strip().upper()
                new_type = "CS" if cb_type.get() == "Cash Bill" else "IV"
                try:
                    new_total = float(e_total.get()) if e_total.get().strip() != "" else None
                    new_comm = float(e_comm.get()) if e_comm.get().strip() != "" else None
                except Exception:
                    messagebox.showerror("Invalid", "Amounts must be numeric.")
                    return

                if not new_bill_no:
                    messagebox.showerror("Missing", "Bill number is required.")
                    return

                if not _validate_bill_format(new_bill_no):
                    messagebox.showerror("Bill No.",
                                         "Invalid bill number format.\nCash: CS-MMDD/XXXX\nInvoice: IV-MMDD/XXXX")
                    return

                # Global duplicate check (case-insensitive) - exclude current record id
                conn_chk = get_conn()
                cur_chk = conn_chk.cursor()
                cur_chk.execute(
                    "SELECT id FROM commissions WHERE LOWER(bill_no) = LOWER(?) AND id <> ?",
                    (new_bill_no, cid)
                )
                dup = cur_chk.fetchone()
                conn_chk.close()
                if dup:
                    messagebox.showerror("Duplicate Bill",
                                         f"This bill number ({new_bill_no}) already exists in the system.")
                    return

                # Update DB
                try:
                    conn = get_conn()
                    cur = conn.cursor()
                    cur.execute("""
                        UPDATE commissions
                        SET bill_type=?, bill_no=?, total_amount=?, commission_amount=?, updated_at=?
                        WHERE id=?
                    """, (
                        new_type, new_bill_no, new_total, new_comm,
                        datetime.now().isoformat(sep=" ", timespec="seconds"), cid
                    ))
                    conn.commit()
                    conn.close()
                except sqlite3.IntegrityError:
                    messagebox.showerror("Duplicate Bill", "This bill already exists (case-insensitive).")
                    return
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to update commission:\n{e}")
                    return

                messagebox.showinfo("Updated", "Commission updated.")
                etop.destroy()
                refresh()

            white_btn(btns, text="Cancel", width=120, command=etop.destroy).pack(side="right", padx=(6, 0))
            white_btn(btns, text="Save Changes", width=140, command=_save_edit).pack(side="right")

        # double-click to edit (admin only)
        def _on_double(e=None):
            if self.current_user and self.current_user.get("role") == "admin":
                _edit_selected()

        tree.bind("<Double-1>", _on_double)

        refresh()

    def _go_fullscreen(self):
        try:
            if sys.platform.startswith("win") or sys.platform.startswith("linux"):
                self.state("zoomed")
            else:
                self.attributes("-fullscreen", True)
        except Exception:
            self.attributes("-fullscreen", True)


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
            except Exception:
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
