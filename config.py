"""Application configuration, paths, and logging."""

from __future__ import annotations

import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path


APP_NAME = "Service Voucher Management System"
APP_SLUG = "svms"

if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent

RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))

# Store writable application data outside the installed application folder.
# Source/development builds continue to use the project directory.
if getattr(sys, "frozen", False):
    default_data_dir = (
        Path(os.environ.get("LOCALAPPDATA", Path.home()))
        / "TONY.COM"
        / "ServiceVoucherApp"
    )
else:
    default_data_dir = APP_DIR

DATA_DIR = Path(
    os.environ.get("SVMS_DATA_DIR", default_data_dir)
).expanduser().resolve()

LOG_DIR = DATA_DIR / "logs"
PDF_DIR = DATA_DIR / "pdfs"
STAFFS_ROOT = DATA_DIR / "staffs"
BACKUP_DIR = DATA_DIR / "backups"
DB_FILE = Path(os.environ.get("SVMS_DB_FILE", DATA_DIR / "vouchers.db")).expanduser().resolve()

for directory in (DATA_DIR, LOG_DIR, PDF_DIR, STAFFS_ROOT, BACKUP_DIR):
    directory.mkdir(parents=True, exist_ok=True)


def resource_path(filename: str) -> Path:
    """Return a source or PyInstaller bundled resource path."""

    return RESOURCE_DIR / filename


SHOP_NAME = os.environ.get("SVMS_SHOP_NAME", "TONY.COM")
SHOP_ADDR = os.environ.get(
    "SVMS_SHOP_ADDR",
    "TB4318, Lot 5, Block 31, Fajar Complex, 91000 Tawau, Sabah, Malaysia",
)
SHOP_TEL = os.environ.get("SVMS_SHOP_TEL", "Tel: 089-763778, H/P: 0168260533")
LOGO_PATH = resource_path("logo.jpg")

DEFAULT_BASE_VID = 41000
FONT_FAMILY = "Segoe UI"
UI_FONT_SIZE = 14

ROLES = ("admin", "sales assistant", "technician", "user")
VOUCHER_STATUSES = ("Pending", "In Progress", "Completed", "Cancelled")
BILL_TYPES = ("CS", "INV")

ROLE_PERMISSIONS = {
    "admin": {
        "voucher.create",
        "voucher.edit",
        "voucher.delete",
        "voucher.status",
        "voucher.export",
        "staff.manage",
        "user.manage",
        "commission.manage",
        "backup.manage",
        "settings.manage",
    },
    "sales assistant": {
        "voucher.create",
        "voucher.edit",
        "voucher.status",
        "voucher.export",
        "commission.manage",
    },
    "technician": {"voucher.status", "voucher.export"},
    "user": {"voucher.export"},
}


logger = logging.getLogger(APP_SLUG)
logger.setLevel(logging.INFO)
logger.propagate = False

if not logger.handlers:
    file_handler = RotatingFileHandler(
        LOG_DIR / "app.log",
        maxBytes=1_000_000,
        backupCount=5,
        encoding="utf-8",
    )
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
    )
    logger.addHandler(file_handler)

    if not getattr(sys, "frozen", False):
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        logger.addHandler(console_handler)
