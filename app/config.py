"""app/config.py — Application configuration.

All values are loaded from the .env file via python-dotenv.
Import this module to access settings throughout the app.
"""
import os
import re
import shutil
import datetime
import logging

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # fall back to real environment variables

# ---- Timezone helpers ----
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
    IST = ZoneInfo("Asia/Kolkata")
except Exception:
    IST = datetime.timezone(datetime.timedelta(hours=5, minutes=30))

UTC = datetime.timezone.utc

def utc_now() -> datetime.datetime:
    return datetime.datetime.now(tz=UTC)

def utc_now_str(sep: str = ' ', timespec: str = 'seconds') -> str:
    return utc_now().isoformat(sep=sep, timespec=timespec)

# ---- IST display helpers ----
def parse_dt_assume_utc(dt_str: str):
    """Parse ISO-like datetime string. If timezone is missing, assume UTC."""
    if not dt_str:
        return None
    try:
        dt = datetime.datetime.fromisoformat(str(dt_str))
    except Exception:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=UTC)
    return dt

def format_dt_ist(dt_str: str) -> str:
    """Return a friendly IST timestamp for display (YYYY-MM-DD HH:MM:SS IST)."""
    dt = parse_dt_assume_utc(dt_str)
    if not dt:
        return (dt_str or '')
    dt_ist = dt.astimezone(IST)
    return dt_ist.strftime('%Y-%m-%d %H:%M:%S IST')

import http.cookies
import uuid
import re
import logging
import traceback

from PIL import Image, ImageOps, ImageEnhance
# Excel export support
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
except Exception:
    Workbook = None
    Font = None
    Alignment = None
    Border = None
    Side = None
    PatternFill = None
try:
    import pytesseract
    from pytesseract import Output
except Exception:
    pytesseract = None
    Output = None

# =========================================================
# Configuration  (all values sourced from .env)
# =========================================================

# Server
DEBUG     = os.environ.get('DEBUG', 'true').strip().lower() in ('1', 'true', 'yes')
HOST      = os.environ.get('HOST', '127.0.0.1').strip()
PORT      = int(os.environ.get('PORT', '8000') or '8000')

# Storage
UPLOAD_DIR = os.environ.get('UPLOAD_DIR', 'uploads').strip()
DB_FILE    = os.environ.get('DB_FILE', 'users.db').strip()

# Feature flags
AUTO_LOCK_NTT_ON_OCR = os.environ.get('AUTO_LOCK_NTT_ON_OCR', 'true').strip().lower() in ('1', 'true', 'yes')

# Default admin seed (used only when DB has no users)
DEFAULT_ADMIN_USERNAME = os.environ.get('DEFAULT_ADMIN_USERNAME', 'admin').strip()
DEFAULT_ADMIN_EMAIL    = os.environ.get('DEFAULT_ADMIN_EMAIL', 'admin@example.com').strip()
DEFAULT_ADMIN_PASSWORD = os.environ.get('DEFAULT_ADMIN_PASSWORD', 'admin').strip()

# -------------------------
# Tesseract (Windows) runtime setup
# -------------------------
# Set TESSERACT_DIR in .env to the folder containing tesseract.exe.
# Leave blank on Linux/macOS if Tesseract is already on PATH.
TESSERACT_DIR = os.environ.get('TESSERACT_DIR', '').strip()

if os.name == 'nt' and TESSERACT_DIR:
    try:
        os.environ['PATH'] = TESSERACT_DIR + os.pathsep + os.environ.get('PATH', '')
        os.environ.setdefault('TESSDATA_PREFIX', os.path.join(TESSERACT_DIR, 'tessdata'))
    except Exception:
        pass



# --- Ensure pytesseract can find Tesseract after TESSERACT_DIR/PATH setup (Windows-safe) ---
def _ensure_tesseract_cmd_ready():
    if pytesseract is None:
        return
    # On Windows, prefer the configured directory
    try:
        if os.name == 'nt':
            candidate = os.path.join(TESSERACT_DIR, 'tesseract.exe')
            if os.path.exists(candidate):
                pytesseract.pytesseract.tesseract_cmd = candidate
                return
    except Exception:
        pass
    # Fallback to PATH
    try:
        found = shutil.which('tesseract')
        if found and os.path.exists(found):
            pytesseract.pytesseract.tesseract_cmd = found
            return
    except Exception:
        pass

_ensure_tesseract_cmd_ready()
os.makedirs(UPLOAD_DIR, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.FileHandler("server.log"),
        logging.StreamHandler()
    ]
)


# -------------------------
# Helpers
# -------------------------
def _safe_name_part(s: str, default="user"):
    s = (s or "").strip()
    if not s:
        return default
    s = re.sub(r"[^a-zA-Z0-9._-]+", "_", s)
    return s[:60] or default


def normalize_screen_access(val: str) -> str:
    val = (val or "BOTH").strip().upper()
    return val if val in ("PPM", "NTT", "EMAIL", "BOTH") else "BOTH"


# -------------------------
# DB connection helper (enables FK)
# -------------------------

# ---- SMTP configuration ----
SMTP_HOST = os.environ.get('SMTP_HOST', '').strip()
SMTP_PORT = int(os.environ.get('SMTP_PORT', '587') or '587')
SMTP_USER = os.environ.get('SMTP_USER', '').strip()
SMTP_PASS = os.environ.get('SMTP_PASS', '').strip()
SMTP_FROM = os.environ.get('SMTP_FROM', '').strip()
SMTP_SSL  = os.environ.get('SMTP_SSL', 'false').strip().lower() in ('1', 'true', 'yes')
