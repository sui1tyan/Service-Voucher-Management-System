import os
import sys
import logging
from logging.handlers import RotatingFileHandler

# ------------------ Paths ------------------
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

LOG_DIR = os.path.join(APP_DIR, "logs")
DB_FILE = os.path.join(APP_DIR, "vouchers.db")
PDF_DIR = os.path.join(APP_DIR, "pdfs")
IMAGES_DIR = os.path.join(APP_DIR, "images")
STAFFS_ROOT = os.path.join(APP_DIR, "staffs")

os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)
os.makedirs(STAFFS_ROOT, exist_ok=True)

# ------------------ Logging ------------------
logger = logging.getLogger("svms")
logger.setLevel(logging.INFO)
_handler = RotatingFileHandler(
    os.path.join(LOG_DIR, "app.log"), maxBytes=512_000, backupCount=3, encoding="utf-8"
)
_formatter = logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
_handler.setFormatter(_formatter)
if not logger.handlers:
    logger.addHandler(_handler)

# ------------------ Constants ------------------
SHOP_NAME = "TONY.COM"
SHOP_ADDR = "TB4318, Lot 5, Block 31, Fajar Complex  91000 Tawau Sabah, Malaysia"
SHOP_TEL = "Tel : 089-763778, H/P: 0168260533"
LOGO_PATH = "" 
DEFAULT_BASE_VID = 41000
FONT_FAMILY = "Segoe UI"
UI_FONT_SIZE = 14
