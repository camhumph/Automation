import os
import re
import csv
import io
import math
import json
import time
import hmac
import socket
import shutil
import secrets
import hashlib
import sqlite3
import asyncio
import logging
import threading
import urllib.request
import urllib.error
import mimetypes
import sys
import xml.etree.ElementTree as ET

from datetime import datetime, date
from typing import List

from fastapi import (
    FastAPI,
    HTTPException,
    WebSocket,
    WebSocketDisconnect,
    Body,
    Request,
    Response,
    UploadFile,
    File,
    Form,
)
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
import uvicorn


# Optional steel-sheet parsers.
# Install if needed:
#   pip install python-multipart openpyxl PyPDF2 pillow pytesseract
try:
    import openpyxl
except Exception:
    openpyxl = None

try:
    import PyPDF2
except Exception:
    PyPDF2 = None

try:
    from PIL import Image
    import pytesseract
except Exception:
    Image = None
    pytesseract = None


# ============================================================
# CONFIGURATION
# ============================================================
APP_NAME = "Elgin Custom Mold Services"
APP_SHORT_NAME = "Elgin CMS"
APP_SLUG = "elgin-custom-mold-services"
APP_PORT = int(os.environ.get("ELGIN_APP_PORT", os.environ.get("NEXUS_APP_PORT", "2926")))

MACHINE_CONFIGS = [
    {
        "name": "UMC-1000",
        "ip": "10.1.10.18",
        "port": 5051,
        "use_qcode": True,
        "use_mtconnect": True,
        "mtconnect_url": os.environ.get(
            "UMC1000_MTCONNECT_URL",
            "http://10.1.10.18:8082/current",
        ),
    },
    {
        "name": "VF-4 (Mill 2)",
        "ip": "10.1.10.19",
        "port": 5051,
        "use_qcode": True,
        "use_mtconnect": True,
        "mtconnect_url": os.environ.get(
            "VF4_MTCONNECT_URL",
            "http://10.1.10.19:8082/current",
        ),
    },
]

MACHINE_RENAME_MAP = {
    "VF-2 (Mill 1)": "UMC-1000",
}

DB_PATH = os.environ.get("ELGIN_DB_PATH", os.environ.get("NEXUS_DB_PATH", "shop_analytics_pro.db"))
ELGIN_SIGNATURE_API_KEY = os.environ.get("ELGIN_SIGNATURE_API_KEY", "cms-signature-upload")
ELGIN_MATCHING_SOFTWARE_DIR = os.environ.get(
    "ELGIN_MATCHING_SOFTWARE_DIR",
    r"\\Mycloudex2ultra\mexico\Cameron's stuff\Matching software",
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(APP_SLUG)


# ============================================================
# NETWORK-AWARE DATA ROOT
# On the company Netgear Wi-Fi  -> PUBLIC share (office can see it).
# Anywhere else                 -> PRIVATE local folder, so Elgin
# runs fine off the company network for testing right now.
# Force with ELGIN_NETWORK_MODE = auto | local | company
# ============================================================
ELGIN_COMPANY_SSID = os.environ.get("ELGIN_COMPANY_SSID", "NETGEAR")
ELGIN_NETWORK_MODE = os.environ.get("ELGIN_NETWORK_MODE", "auto").strip().lower()
ELGIN_PUBLIC_DATA_ROOT = os.environ.get(
    "ELGIN_PUBLIC_DATA_ROOT",
    r"\\Mycloudex2ultra\mexico\Cameron's stuff\Matching software",
)
ELGIN_PRIVATE_DATA_ROOT = os.environ.get(
    "ELGIN_PRIVATE_DATA_ROOT",
    r"C:\CMS_Local_Workspace\Matching",
)


def _ms_get_wifi_ssid() -> str:
    try:
        import subprocess
        out = subprocess.run(
            ["netsh", "wlan", "show", "interfaces"],
            capture_output=True, text=True, timeout=4,
        ).stdout or ""
    except Exception:
        return ""
    for line in out.splitlines():
        s = line.strip()
        # A line that starts with "SSID" (so "BSSID" is skipped).
        if s.upper().startswith("SSID") and ":" in s:
            return s.split(":", 1)[1].strip()
    return ""


def ms_on_company_network() -> bool:
    if ELGIN_NETWORK_MODE == "local":
        return False
    if ELGIN_NETWORK_MODE == "company":
        return True
    ssid = _ms_get_wifi_ssid()
    if ssid and ELGIN_COMPANY_SSID.upper() in ssid.upper():
        return True
    return False


def resolve_cms_data_root() -> str:
    return ELGIN_PUBLIC_DATA_ROOT if ms_on_company_network() else ELGIN_PRIVATE_DATA_ROOT


# Resolve the shared matching folder by network, unless the user pinned it via env.
ELGIN_CMS_ON_COMPANY = ms_on_company_network()
if "ELGIN_MATCHING_SOFTWARE_DIR" not in os.environ:
    ELGIN_MATCHING_SOFTWARE_DIR = resolve_cms_data_root()
try:
    os.makedirs(ELGIN_PRIVATE_DATA_ROOT, exist_ok=True)
except Exception:
    pass
logger.info(
    "CMS data root: %s  (network mode=%s, on_company=%s)",
    ELGIN_MATCHING_SOFTWARE_DIR, ELGIN_NETWORK_MODE, ELGIN_CMS_ON_COMPANY,
)


HTTPS_CERT_FILE = os.environ.get("ELGIN_SSL_CERT", os.environ.get("NEXUS_SSL_CERT", ""))
HTTPS_KEY_FILE = os.environ.get("ELGIN_SSL_KEY", os.environ.get("NEXUS_SSL_KEY", ""))

UPLOAD_DIR = os.environ.get("ELGIN_UPLOAD_DIR", os.environ.get("NEXUS_UPLOAD_DIR", "uploads"))
STEEL_UPLOAD_DIR = os.path.join(UPLOAD_DIR, "steel_sheets")
os.makedirs(STEEL_UPLOAD_DIR, exist_ok=True)

PROGRAM_SEARCH_DIRS = [
    p.strip()
    for p in os.environ.get(
        "ELGIN_PROGRAM_DIRS",
        os.environ.get(
            "NEXUS_PROGRAM_DIRS",
            r"C:\NC_PROGRAMS;C:\HAAS_PROGRAMS;C:\Users\Public\Documents\NC_PROGRAMS",
        ),
    ).split(";")
    if p.strip()
]

PROGRAM_EXTS = [
    ".nc",
    ".eia",
    ".min",
    ".txt",
    ".mpf",
    ".tap",
    ".ngc",
]

GCODE_ANALYSIS_MAX_FILE_BYTES = 5_000_000
GCODE_ANALYSIS_REFRESH_SEC = 6 * 60 * 60

# Used by the historical fallback estimate system.
# If Q-code and G-code do not produce a runtime estimate, the system looks for
# completed same-component jobs with similar physical job/catalog dimensions.
SIMILAR_JOB_TOLERANCE = float(os.environ.get("ELGIN_SIMILAR_JOB_TOLERANCE", "0.03"))

JOB_SIGNATURE_COMPONENTS = [
    "TCP",
    "BCP",
    "ID HOLDER",
    "OD HOLDER",
    "ID POT",
    "OD POT",
]

JOB_SIGNATURE_COMPONENT_ALIASES = {
    "TCP": "TCP",
    "TOPSMED": "TCP",
    "TOPSMEDPLATE": "TCP",
    "TOPCLAMPING": "TCP",
    "TOPCLAMPINGPLATE": "TCP",
    "BCP": "BCP",
    "BOTTOMSMED": "BCP",
    "BOTSMED": "BCP",
    "BOTTOMSMEDPLATE": "BCP",
    "BOTTOMCLAMPING": "BCP",
    "BOTTOMCLAMPINGPLATE": "BCP",
    "BOTCLAMPING": "BCP",
    "IDHOLDER": "ID HOLDER",
    "TOPHOLDER": "ID HOLDER",
    "IDTEHOLDER": "ID HOLDER",
    "ODHOLDER": "OD HOLDER",
    "BOTTOMHOLDER": "OD HOLDER",
    "BOTHOLDER": "OD HOLDER",
    "ODTEHOLDER": "OD HOLDER",
    "IDPOT": "ID POT",
    "IDPOTBLOCK": "ID POT",
    "TOPPOT": "ID POT",
    "TOPPOTBLOCK": "ID POT",
    "ODPOT": "OD POT",
    "ODPOTBLOCK": "OD POT",
    "BOTTOMPOT": "OD POT",
    "BOTTOMPOTBLOCK": "OD POT",
    "BOTPOT": "OD POT",
    "BOTPOTBLOCK": "OD POT",
}

JOB_SIGNATURE_DEFAULT_LIMIT = int(os.environ.get("ELGIN_JOB_SIGNATURE_MATCH_LIMIT", "10"))


# ============================================================
# ADMIN / SECURITY
# ============================================================
ADMIN_PASSWORD = os.environ.get(
    "ELGIN_ADMIN_PASSWORD",
    os.environ.get("NEXUS_ADMIN_PASSWORD", "cms123"),
)
ADMIN_SESSION_SECRET = os.environ.get(
    "ELGIN_ADMIN_SESSION_SECRET",
    os.environ.get("NEXUS_ADMIN_SESSION_SECRET", ""),
)

if not ADMIN_SESSION_SECRET:
    ADMIN_SESSION_SECRET = secrets.token_urlsafe(48)

ADMIN_COOKIE_NAME = "elgin_admin"
ADMIN_SESSION_TTL_SEC = 8 * 60 * 60

# These fields are hidden from non-admin responses and UI.
# Money/profit/cost/labor data remains admin-only.
FINANCIAL_FIELDS = [
    "revenue",
    "steel_cost",
    "grinding_cost",
    "pyropel_cost",
    "labor_rate",
    "labor_override_cost",
    "actual_sec",
    "labor_hours",
    "labor_cost",
    "total_cost",
    "profit",
    "margin_pct",
]

# Admin-only actions:
# - financial updates
# - delete job
# - rename/merge job
# - delete app-calendar entry
# - delete shop board row
# - tool replacement reset
# - tool inspection/calibration
# - confirm/ignore/clear replacement warnings
# - fake job cleanup
# - MTConnect diagnostic


# ============================================================
# MACHINE / TOOL CONSTANTS
# ============================================================
L_OFFSET_BASE = 2000
D_OFFSET_BASE = 2200

# Mark idle only after RPM has not been over zero/valid for 5 minutes.
IDLE_AFTER_NO_RPM_SEC = 5 * 60.0

RPM_HOLD_LAST_GOOD_SEC = 120.0
RPM_FALLBACK_FOR_WEAR = 1200.0

# Hold last valid tool to avoid temporary T00 / T0 reads during tool changes.
TOOL_HOLD_LAST_GOOD_SEC = 10 * 60.0

SPINDLE_LOAD_CUTTING_THRESHOLD = 3.0
FEED_CUTTING_THRESHOLD = 0.01

RPM_REGISTER_CANDIDATES = [
    3027,
    3028,
    3029,
    4119,
    4019,
]

DEFAULT_LABOR_RATE_PER_HOUR = 75.0
DEFAULT_WORK_MATERIAL = "4140"

DEFAULT_JOB_COMPONENTS = [
    "ID POT",
    "OD POT",
    "ID HOLDER",
    "OD HOLDER",
    "BCP",
    "TCP",
    "PYROPEL",
    "J BLOCK",
]

DEFAULT_COMPONENT_MACHINE = {
    "ID POT": "UMC-1000",
    "OD POT": "UMC-1000",
    "ID HOLDER": "UMC-1000",
    "OD HOLDER": "UMC-1000",
}

# Do NOT hard-code new jobs/components to 45 minutes.
# Runtime/progress uses:
#   1. Q-code live cycle time
#   2. G-code estimate
#   3. similar completed job actual time, matched by component and part signature
#   4. otherwise "no estimate"
DEFAULT_COMPONENT_EST_SEC = {
    "ID POT": 0,
    "OD POT": 0,
    "ID HOLDER": 0,
    "OD HOLDER": 0,
    "BCP": 0,
    "TCP": 0,
    "PYROPEL": 0,
    "J BLOCK": 0,
    "ENGRAVING": 0,
    "UNCLASSIFIED": 0,
}

DEFAULT_UNESTIMATED_COMPONENT_SEC = 0.0

# Programs/components that usually indicate a current direct job has reached
# a finishing/final operation. This is used later when old/template programs
# are mapped to current matched jobs.
FINAL_OPERATION_COMPONENTS = {
    "ENGRAVING",
}


# ============================================================
# TOOL HEALTH / REPLACEMENT DETECTION SETTINGS
# ============================================================
ENABLE_AUTO_TOOL_REPLACEMENT_CANDIDATES = os.environ.get(
    "ELGIN_AUTO_TOOL_REPLACEMENT_CANDIDATES",
    os.environ.get("NEXUS_AUTO_TOOL_REPLACEMENT_CANDIDATES", "1"),
).strip().lower() in ("1", "true", "yes", "on")

# Recommended:
#   setx NEXUS_USE_OFFSET_WEAR_IN_HEALTH "0"
#
# Haas offsets are often geometry/operator adjustments, not direct tool wear.
USE_OFFSET_WEAR_IN_HEALTH = os.environ.get(
    "ELGIN_USE_OFFSET_WEAR_IN_HEALTH",
    os.environ.get("NEXUS_USE_OFFSET_WEAR_IN_HEALTH", "0"),
).strip().lower() in ("1", "true", "yes", "on")

WIPS_LENGTH_CHANGE_THRESHOLD = 0.025
WIPS_DIAMETER_CHANGE_THRESHOLD = 0.010
WIPS_CANDIDATE_COOLDOWN_SEC = 900
WIPS_STABLE_CONFIRM_SAMPLES = 4
WIPS_MIN_IDLE_BEFORE_DETECT_SEC = 20.0
WIPS_MAX_REASONABLE_OFFSET_ABS = 5.0

TOOL_CATEGORY_BASE_LIFE_4140_MIN = {
    "Tap": 12,
    "Thread Mill": 28,
    "Drill": 35,
    "Spot Drill": 70,
    "Ball Endmill": 35,
    "Face Mill": 110,
    "Chamfer Mill": 55,
    "Inserted Cutter": 95,
    "Endmill": 45,
    "Tool": 40,
}


# ============================================================
# UMC-1000 SCHEDULER SETTINGS
# ============================================================
UMC_SCHEDULER_MACHINE = "UMC-1000"

UMC_SCHEDULER_COMPONENTS = [
    "ID POT",
    "OD POT",
]

UMC_DEFAULT_SETUP_CHANGE_SEC = 15 * 60
UMC_SIMILAR_JOB_SETUP_CHANGE_SEC = 5 * 60
UMC_SAME_COMPONENT_SETUP_CHANGE_SEC = 3 * 60

STEEL_READY_WORDS = [
    "READY",
    "DONE",
    "CUT",
    "RECEIVED",
    "RECIEVED",
    "YES",
    "Y",
    "IN HOUSE",
    "IN-HOUSE",
    "SAWED",
    "COMPLETE",
]

STEEL_WAIT_WORDS = [
    "WAIT",
    "WAITING",
    "ORDER",
    "ORDERED",
    "NO",
    "N",
    "HOLD",
    "NEED",
    "NEEDS",
    "BACKORDER",
    "BACK ORDER",
]


# ============================================================
# GENERAL HELPERS
# ============================================================
def normalize_job_num(v: str) -> str:
    s = str(v or "").upper().strip()
    m = re.search(r"[JO]?\s*(\d{3,8})", s)
    if m:
        return m.group(1)
    s = re.sub(r"[^A-Z0-9]", "", s)
    s = re.sub(r"^[JO]", "", s)
    return s


def job_num_is_all_zeros(v: str) -> bool:
    """
    True for bad/fake controller/default jobs:
      J0, J00, J000, J0000, J00000, etc.
    """
    j = normalize_job_num(v)
    return bool(j and re.fullmatch(r"0+", j))


def is_fake_job_num(v: str) -> bool:
    """
    Fake job numbers we should not auto-create.

    J1000 is not globally invalid because it could be a real shop job.
    UMC-1000 -> J1000 is blocked separately by program_looks_like_machine_name().
    """
    j = normalize_job_num(v)
    if not j:
        return True
    if job_num_is_all_zeros(j):
        return True
    return False


def validate_new_job_num_or_400(v: str) -> str:
    """
    Used when creating/renaming jobs.
    Prevents creating all-zero fake jobs like J0000.
    """
    j = normalize_job_num(v)
    if not j:
        raise HTTPException(400, "Job number required")
    if job_num_is_all_zeros(j):
        raise HTTPException(400, "All-zero job numbers like J0000 are invalid")
    return j


def normalize_component(v: str) -> str:
    return re.sub(r"\s+", " ", str(v or "").strip().upper())


def _today_str():
    return datetime.now().strftime("%Y-%m-%d")


def _date_obj(s):
    if not s:
        return None
    try:
        return datetime.strptime(str(s)[:10], "%Y-%m-%d").date()
    except Exception:
        return None


def _rowdicts(cur):
    return [dict(r) for r in cur.fetchall()]


def _money(v) -> float:
    try:
        return round(float(v or 0), 2)
    except Exception:
        return 0.0


def _clamp(x, lo, hi):
    return max(lo, min(hi, x))


def _b64safe(v: str) -> str:
    return v.replace("+", "-").replace("/", "_").replace("=", "")


def _compact_name(v: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(v or "").upper())


def program_looks_like_machine_name(program: str, machine_name: str) -> bool:
    """
    Prevent controller/machine names like UMC-1000 from becoming fake jobs like J1000.
    """
    p = _compact_name(program)
    m = _compact_name(machine_name)

    if not p or not m:
        return False

    if p == m:
        return True

    machine_like_prefixes = [
        "UMC",
        "VF",
        "VF2",
        "VF3",
        "VF4",
        "VF5",
        "MILL",
        "LATHE",
        "HAAS",
    ]

    if any(p.startswith(x) for x in machine_like_prefixes):
        return True

    return False


def safe_upload_filename(v: str) -> str:
    name = os.path.basename(str(v or "").strip())
    name = re.sub(r"[^A-Za-z0-9._\- ]", "_", name)
    name = re.sub(r"\s+", "_", name)
    return name or "steel_sheet.pdf"


def normalize_tool_num(raw) -> str:
    """
    Normalize tool values from Haas Q-code / MTConnect.

    Fixes issue where tool displayed as T00:
    - T00, 00, 000, 0, N/A, UNKNOWN -> "0"
    - T0182, 0182, TOOL 182 -> "182"
    """
    if raw is None:
        return "0"

    s = str(raw or "").strip().replace("\x00", "").upper()
    if not s:
        return "0"

    if s in ("N/A", "NA", "NONE", "NULL", "UNKNOWN", "UNAVAILABLE", "NOT_READY", "NOTREADY", "?"):
        return "0"

    nums = re.findall(r"\d+", s)
    if not nums:
        return "0"

    # Usually the tool number is the last numeric token returned.
    n_raw = nums[-1]
    try:
        n = int(n_raw)
    except Exception:
        return "0"

    if n <= 0:
        return "0"

    return str(n)


def apply_tool_hold_last_good(raw_tool, state: dict) -> str:
    """
    Avoid showing T00 during temporary controller reads/tool changes.

    If raw tool normalizes to 0, hold the last valid nonzero tool for a short time.
    """
    now = time.time()
    tool = normalize_tool_num(raw_tool)

    if tool != "0":
        state["_last_valid_tool"] = tool
        state["_last_valid_tool_at"] = now
        return tool

    last = normalize_tool_num(state.get("_last_valid_tool", "0"))
    last_at = float(state.get("_last_valid_tool_at", 0.0) or 0.0)

    if last != "0" and now - last_at <= TOOL_HOLD_LAST_GOOD_SEC:
        return last

    return "0"


# ============================================================
# NETWORK / LOCAL APP HOSTING HELPERS
# ============================================================
def get_local_network_addresses(port: int = APP_PORT) -> list:
    """
    Returns local URLs the host computer can use.

    This supports the setup where the main computer runs the app and other
    shop computers reach it over Ethernet/private LAN without needing Wi-Fi.
    """
    ips = set()
    ips.add("127.0.0.1")

    try:
        hostname = socket.gethostname()
        for info in socket.getaddrinfo(hostname, None, socket.AF_INET):
            ip = info[4][0]
            if ip and not ip.startswith("127."):
                ips.add(ip)
    except Exception:
        pass

    # Finds the default outbound adapter IP. No real traffic needs to complete.
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.settimeout(0.2)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            if ip:
                ips.add(ip)
    except Exception:
        pass

    out = []
    for ip in sorted(ips):
        out.append({
            "ip": ip,
            "http_url": f"http://{ip}:{port}",
            "https_url": f"https://{ip}:{port}",
            "is_localhost": ip.startswith("127."),
        })

    return out


def write_windows_launcher_file():
    """
    Creates a simple double-click launcher next to this Python file.

    This does not expose the app to the internet.
    It simply starts the local server on the host computer.
    """
    if os.name != "nt":
        return ""

    try:
        script_path = os.path.abspath(sys.argv[0])
        script_dir = os.path.dirname(script_path)
        launcher_path = os.path.join(script_dir, "start_elgin_custom_mold_services.bat")

        content = f"""@echo off
title {APP_NAME}
cd /d "{script_dir}"
echo Starting {APP_NAME}...
echo.
python "{script_path}"
echo.
pause
"""

        # Do not overwrite if user edited it.
        if not os.path.exists(launcher_path):
            with open(launcher_path, "w", encoding="utf-8") as f:
                f.write(content)

        return launcher_path

    except Exception:
        return ""


# ============================================================
# ADMIN TOKEN HELPERS
# ============================================================
def _admin_signature(exp: int, nonce: str) -> str:
    msg = f"{exp}:{nonce}".encode("utf-8")
    key = ADMIN_SESSION_SECRET.encode("utf-8")
    return _b64safe(hmac.new(key, msg, hashlib.sha256).hexdigest())


def make_admin_token() -> str:
    exp = int(time.time()) + ADMIN_SESSION_TTL_SEC
    nonce = secrets.token_urlsafe(18)
    sig = _admin_signature(exp, nonce)
    return f"{exp}.{nonce}.{sig}"


def verify_admin_token(token: str) -> bool:
    try:
        exp_s, nonce, sig = str(token or "").split(".", 2)
        exp = int(exp_s)
        if time.time() > exp:
            return False
        expected = _admin_signature(exp, nonce)
        return hmac.compare_digest(sig, expected)
    except Exception:
        return False


def is_admin_request(request: Request) -> bool:
    if not request:
        return False
    token = request.cookies.get(ADMIN_COOKIE_NAME, "")
    return verify_admin_token(token)


def require_admin_request(request: Request):
    if not is_admin_request(request):
        raise HTTPException(status_code=401, detail="Admin login required")


def is_signature_api_request(request: Request) -> bool:
    if not request:
        return False
    supplied = request.headers.get("X-Elgin-Api-Key", "")
    return bool(ELGIN_SIGNATURE_API_KEY) and hmac.compare_digest(supplied, ELGIN_SIGNATURE_API_KEY)


def require_signature_or_admin_request(request: Request):
    if is_admin_request(request) or is_signature_api_request(request):
        return
    raise HTTPException(status_code=401, detail="Admin login or signature API key required")


# ============================================================
# COMPONENT / PROGRAM PARSING
# ============================================================
COMPONENT_ALIASES = {
    "ID POT": ["ID POT", "I D POT", "IDPOT", "I.D. POT", "INNER POT", "INSIDE POT"],
    "OD POT": ["OD POT", "O D POT", "ODPOT", "O.D. POT", "OUTER POT", "OUTSIDE POT"],
    "ID HOLDER": ["ID HOLDER", "I D HOLDER", "IDHOLDER", "I.D. HOLDER", "INNER HOLDER"],
    "OD HOLDER": ["OD HOLDER", "O D HOLDER", "ODHOLDER", "O.D. HOLDER", "OUTER HOLDER"],
    "BCP": ["BCP"],
    "TCP": ["TCP"],
    "PYROPEL": ["PYROPEL", "PYRO", "PYRO PEL"],
    "J BLOCK": ["J BLOCK", "JBLOCK", "J-BLOCK"],
    "ENGRAVING": ["ENGRAVING", "ENGRAVE", "ENGRAV"],
}


def detect_component(name: str):
    raw = normalize_component(name)
    if not raw:
        return None

    spaced = re.sub(r"[^A-Z0-9]+", " ", raw).strip()
    compact = re.sub(r"[^A-Z0-9]", "", raw)

    if "ENGRAV" in compact:
        return "ENGRAVING"

    checks = []
    for comp, aliases in COMPONENT_ALIASES.items():
        for alias in aliases:
            a_spaced = re.sub(r"[^A-Z0-9]+", " ", alias.upper()).strip()
            a_compact = re.sub(r"[^A-Z0-9]", "", alias.upper())
            checks.append((comp, a_spaced, a_compact))

    checks.sort(key=lambda x: len(x[2]), reverse=True)
    spaced_padded = f" {spaced} "

    for comp, a_spaced, a_compact in checks:
        if a_spaced and f" {a_spaced} " in spaced_padded:
            return comp
        if a_compact and a_compact in compact:
            return comp

    return None


def canonical_component(name: str) -> str:
    detected = detect_component(name)
    if detected:
        return detected

    n = normalize_component(name)
    n = re.sub(r"\b\d+\s*X\b$", "", n).strip()
    n = re.sub(r"\bOP\s*\d+\b", "", n).strip()
    n = re.sub(r"\bREV\s*\d+\b", "", n).strip()
    n = re.sub(r"\.(NC|EIA|MIN|TXT|MPF)$", "", n, flags=re.I).strip()

    return n or "UNCLASSIFIED"


def _idle_program_parse_result(display: str = "IDLE"):
    return {
        "job_num": None,
        "sub_cat": None,
        "part_name": None,
        "display": display,
        "full_name": None,
    }


def parse_program_name(prog: str):
    """
    Parse a CNC program name into job/component info.

    Updated:
    - All-zero jobs like J0000 are ignored.
    - Machine names are blocked separately by program_looks_like_machine_name().
    """
    if not prog:
        return _idle_program_parse_result()

    raw = str(prog).strip().replace("\x00", "").strip()
    if raw.upper() in ("N/A", "", "0", "IDLE", "NONE", "NULL", "?"):
        return _idle_program_parse_result()

    p = os.path.basename(raw.replace("\\", "/"))
    p = re.sub(r"\.(NC|EIA|MIN|TXT|MPF)$", "", p, flags=re.I).strip()
    pu = p.upper()

    jm = re.search(r"(?:^|[^A-Z0-9])([JO])?(\d{3,8})(?=[^0-9]|$)", pu)

    job_num = None
    tail = ""

    if jm:
        job_num = normalize_job_num(jm.group(2))
        tail = pu[jm.end(2):]
    else:
        jm2 = re.search(r"([JO])(\d{3,8})", pu)
        if jm2:
            job_num = normalize_job_num(jm2.group(2))
            tail = pu[jm2.end(2):]

    if job_num and job_num_is_all_zeros(job_num):
        return _idle_program_parse_result("IDLE")

    detected = detect_component(tail) or detect_component(pu)

    if job_num:
        part_name = detected or canonical_component(tail) or None
        display_part = detected or part_name or ""
        return {
            "job_num": job_num,
            "program_job_num": job_num,
            "effective_job_num": job_num,
            "matched_from_program_job": "",
            "matched_reason": "",
            "sub_cat": display_part,
            "part_name": display_part,
            "display": f"J{job_num}" + (f" · {display_part}" if display_part else ""),
            "full_name": f"{job_num} {display_part}".strip(),
        }

    detected = detect_component(pu)
    part_name = detected or canonical_component(pu)

    return {
        "job_num": None,
        "program_job_num": None,
        "effective_job_num": None,
        "matched_from_program_job": "",
        "matched_reason": "",
        "sub_cat": part_name,
        "part_name": part_name,
        "display": part_name or pu,
        "full_name": part_name or pu,
    }


# ============================================================
# TOOL HELPERS
# ============================================================
def infer_tool_category(name: str) -> str:
    n = (name or "").upper()
    if "TAP" in n and "THREADMILL" not in n and "THREAD MILL" not in n:
        return "Tap"
    if "THREAD" in n:
        return "Thread Mill"
    if "DRILL" in n and "SPOT" not in n:
        return "Drill"
    if "SPOT" in n:
        return "Spot Drill"
    if "BALL" in n:
        return "Ball Endmill"
    if "FACE" in n or "SHELL" in n:
        return "Face Mill"
    if "CHAMF" in n:
        return "Chamfer Mill"
    if "DAPRA" in n or "INGERSOLL" in n or "DIJET" in n or "ZENIT" in n:
        return "Inserted Cutter"
    if "ENDMILL" in n or "END MILL" in n or re.search(r"\bEM\b", n):
        return "Endmill"
    return "Tool"


def infer_tool_diameter(tool_name: str) -> float:
    n = str(tool_name or "").upper()

    m = re.search(r"(?<!\d)(\d?\.\d{2,4})(?!\d)", n)
    if m:
        try:
            return float(m.group(1))
        except Exception:
            pass

    m = re.search(r"(?<!\d)(\d{1,2})\s*/\s*(\d{1,2})(?!\d)", n)
    if m:
        try:
            return float(m.group(1)) / float(m.group(2))
        except Exception:
            pass

    m = re.search(r"(?<!\d)(\d(?:\.\d+)?)\s*IN\b", n)
    if m:
        try:
            return float(m.group(1))
        except Exception:
            pass

    return 0.0


def infer_flutes(tool_name: str) -> int:
    n = str(tool_name or "").upper()
    m = re.search(r"(\d+)\s*FL", n)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            pass

    if "DRILL" in n:
        return 2
    if "TAP" in n:
        return 3
    return 4


def estimate_default_tool_life_4140_min(tool_name: str, category: str = "") -> float:
    category = category or infer_tool_category(tool_name)
    dia = infer_tool_diameter(tool_name)
    base = TOOL_CATEGORY_BASE_LIFE_4140_MIN.get(category, 40)

    if dia >= 2.0:
        base *= 1.8
    elif dia >= 1.0:
        base *= 1.35
    elif dia >= 0.5:
        base *= 1.05
    elif dia > 0 and dia < 0.1875:
        base *= 0.55

    n = str(tool_name or "").upper()

    if any(x in n for x in ("INGERSOLL", "DAPRA", "DIJET", "ZENIT", "RAPTOR")):
        base *= 1.25

    if "TAP" in n and "THREADMILL" not in n and "THREAD MILL" not in n:
        base *= 0.75

    if dia and dia <= 0.125:
        base *= 0.65

    return round(_clamp(base, 5, 240), 2)


def get_tool_name(tid: str) -> str:
    tid = normalize_tool_num(tid)
    if tid == "0":
        return "No Tool"

    for t, name in TOOL_CATALOG_SEED:
        if str(t) == tid:
            return name

    return f"Tool {tid}"


def tool_health_color(score: float) -> str:
    score = float(score or 0)
    if score >= 85:
        return "red"
    if score >= 65:
        return "amber"
    return "green"


# ============================================================
# TOOL CATALOG SEED
# ============================================================
TOOL_CATALOG_SEED = [
    ("191", "0.25 1/4-20 UNC Right-Handed Tap"),
    ("182", ".25 GORILLA LONG 5FL"),
    ("20", "0.5 RAPTOR"),
    ("37", "0.640 INGERSOLL PILOT"),
    ("184", "0.3125 WIDIN .02R"),
    ("8", "0.375 3/8 X 90 Spot Drill"),
    ("9", "0.25 Spot Drill"),
    ("25", ".5 DAPRA"),
    ("45", "1/8 NPT THREAD MILL / Chamfer Tool VERIFY"),
    ("152", "0.5 Drill / .500 OSG POT DRILL VERIFY"),
    ("175", "0.5 Flat Endmill"),
    ("159", ".9375 DRILL / 0.625 Drill VERIFY"),
    ("35", ".531 SPADE DRILL / 0.625 SPADE / 0.6875 Drill VERIFY"),
    ("180", "0.5 BALL MILL"),
    ("192", "7/16-20 THREADMILL"),
    ("3", "0.213 Drill"),
    ("28", "0.25 Spot Drill"),
    ("16", "0.75 FLAT ENDMILL WIDIN"),
    ("22", "0.75 ZENIT HF MILL / 2.25 Face Mill VERIFY"),
    ("176", "0.123 Flat Endmill / 0.125 Ball-Nosed Endmill / 0.06 Ball VERIFY"),
    ("173", "1IN LONG FIN EM"),
    ("24", "2IN ZENIT HF MILL"),
    ("160", "0.343 Drill"),
    ("27", "0.625 Spot Drill"),
    ("161", "0.3125 Drill"),
    ("11", "0.25 GORILLA HARDMILL / 0.25 Ball-Nosed Endmill VERIFY"),
    ("46", "1IN DIJET"),
    ("179", "0.36 Thread Mill"),
    ("30", "0.5 1/2-13 UNC Right-Handed Tap"),
    ("21", "1 Flat Endmill"),
    ("40", "0.6875 INGERSOLL PILOT"),
    ("38", "0.64 INGERSOLL LONGPOTS"),
    ("198", "0.3125 5FL GORILLA"),
    ("163", "0.250 Drill"),
    ("2", "2.25 Face Mill / Shell Mill"),
    ("44", "0.375 Drill"),
    ("221", ".625 ZENIT HF MILL"),
    ("212", "0.375 Flat Endmill"),
    ("219", "0.1875 Flat Endmill"),
    ("5", "GORILLA CHAMFER-0.375-CH"),
    ("200", "0.640 Drill"),
    ("42", "2 Bull-Nosed Endmill"),
    ("10", "0.25 FLAT ENDMILL"),
    ("153", "0.1875 Drill"),
    ("101", "0.14 Drill"),
    ("150", "0.343 Drill"),
    ("14", "0.5 FLAT ENDMILL WIDIN"),
    ("18", "0.5 Flat Endmill"),
    ("158", "0.375 LONG DRILL"),
    ("17", "0.75 FLAT FINISH ENDMILL"),
    ("32", "0.3125 5/16-18 UNC Right-Handed Tap"),
    ("151", "0.406 Drill"),
    ("190", "0.75 LONG FIN EM"),
    ("39", "0.64 INGERSOLL LONG"),
    ("81", "0.422 Drill"),
    ("43", "0.705 Dovetail Cutter"),
    ("15", "0.5 FLAT FINISH ENDMILL"),
    ("13", "0.375 FLAT FIN ENDMILL"),
    ("23", "1IN ZENIT HF MILL"),
    ("31", "0.25 1/4-20 UNC Right-Handed Tap"),
    ("47", "3IN FACEMILL FIN"),
    ("199", "0.5 CHAMF CARBIDE"),
    ("156", "0.281 Drill"),
    ("174", "0.499 Flat Endmill"),
    ("171", "0.375 3/8-16 UNC Right-Handed Tap"),
    ("164", "0.265 Drill"),
    ("195", ".500 SNAP RING CUTTER"),
    ("188", ".062 WIDIN GREASE CUTTER"),
    ("26", "1 Chamfer Mill"),
    ("183", ".25 WIDIN FINISH"),
    ("181", "0.236 WIDIN Endmill"),
    ("162", ".375 LONG ENDMILL"),
    ("120", ".5625 DRILL"),
    ("41", "0.6875 LONG SPADE"),
]


# ============================================================
# END PART 1
# PART 2 CONTINUES WITH:
# - database schema/init
# - request models
# - app/state setup
# ============================================================

# ============================================================
# DATABASE
# ============================================================
def _add_col(conn, table, coldef):
    try:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {coldef}")
    except Exception:
        pass


def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys=ON")

        # ------------------------
        # Inventory
        # ------------------------
        conn.execute("""CREATE TABLE IF NOT EXISTS inventory (
            sku TEXT PRIMARY KEY,
            description TEXT,
            category TEXT,
            quantity INTEGER DEFAULT 0,
            min_threshold INTEGER DEFAULT 5
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sku TEXT,
            direction TEXT,
            qty INTEGER,
            timestamp TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS inventory_aliases (
            alias TEXT PRIMARY KEY,
            sku TEXT NOT NULL,
            description TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        # ------------------------
        # Tools
        # ------------------------
        conn.execute("""CREATE TABLE IF NOT EXISTS tool_catalog (
            tid TEXT PRIMARY KEY,
            tool_name TEXT NOT NULL DEFAULT '',
            category TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            created_at TEXT DEFAULT '',
            updated_at TEXT DEFAULT ''
        )""")

        for col in [
            "diameter REAL DEFAULT 0",
            "flutes INTEGER DEFAULT 0",
            "tool_material TEXT DEFAULT ''",
            "coating TEXT DEFAULT ''",
            "work_material TEXT DEFAULT '4140'",
            "max_life_min REAL DEFAULT 0",
            "default_sfm REAL DEFAULT 0",
            "default_chipload REAL DEFAULT 0",
        ]:
            _add_col(conn, "tool_catalog", col)

        conn.execute("""CREATE TABLE IF NOT EXISTS tool_health (
            machine TEXT,
            tid TEXT,
            tool_name TEXT DEFAULT '',
            in_cut_sec REAL DEFAULT 0,
            max_life_min REAL DEFAULT 60,
            peak_load REAL DEFAULT 0,
            peak_rpm REAL DEFAULT 0,
            wear_score REAL DEFAULT 0,
            last_seen TEXT DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY(machine, tid)
        )""")

        for col in [
            "peak_rpm REAL DEFAULT 0",
            "wear_score REAL DEFAULT 0",
            "last_seen TEXT DEFAULT ''",
            "active_cycle_id INTEGER DEFAULT 0",
            "cycle_started_at TEXT DEFAULT ''",
            "last_l_wear REAL DEFAULT 0",
            "last_d_wear REAL DEFAULT 0",
            "last_probe_candidate_at REAL DEFAULT 0",
            "replacement_pending INTEGER DEFAULT 0",
        ]:
            _add_col(conn, "tool_health", col)

        conn.execute("""CREATE TABLE IF NOT EXISTS tool_life_cycles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine TEXT NOT NULL,
            tid TEXT NOT NULL,
            tool_name TEXT DEFAULT '',
            cycle_no INTEGER DEFAULT 1,
            started_at TEXT DEFAULT CURRENT_TIMESTAMP,
            ended_at TEXT DEFAULT '',
            total_cut_sec REAL DEFAULT 0,
            avg_load REAL DEFAULT 0,
            peak_load REAL DEFAULT 0,
            avg_rpm REAL DEFAULT 0,
            peak_rpm REAL DEFAULT 0,
            avg_feed REAL DEFAULT 0,
            sample_count INTEGER DEFAULT 0,
            final_wear_score REAL DEFAULT 0,
            replacement_reason TEXT DEFAULT '',
            replacement_notes TEXT DEFAULT '',
            replacement_source TEXT DEFAULT '',
            archived_by_event_id INTEGER DEFAULT 0,
            status TEXT DEFAULT 'ACTIVE'
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS tool_cut_samples (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine TEXT NOT NULL,
            tid TEXT NOT NULL,
            tool_cycle_id INTEGER DEFAULT 0,
            tool_name TEXT DEFAULT '',
            prog_raw TEXT DEFAULT '',
            job_num TEXT DEFAULT '',
            component_name TEXT DEFAULT '',
            spindle_load REAL DEFAULT 0,
            spindle_rpm REAL DEFAULT 0,
            feed_rate REAL DEFAULT 0,
            cut_delta_sec REAL DEFAULT 0,
            in_cut_total_sec REAL DEFAULT 0,
            wear_score REAL DEFAULT 0,
            l_wear REAL DEFAULT 0,
            d_wear REAL DEFAULT 0,
            actual_cutting INTEGER DEFAULT 0,
            timestamp TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS tool_replacement_candidates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine TEXT NOT NULL,
            tid TEXT NOT NULL,
            tool_name TEXT DEFAULT '',
            old_l_wear REAL DEFAULT 0,
            new_l_wear REAL DEFAULT 0,
            delta_l_wear REAL DEFAULT 0,
            old_d_wear REAL DEFAULT 0,
            new_d_wear REAL DEFAULT 0,
            delta_d_wear REAL DEFAULT 0,
            status TEXT DEFAULT 'PENDING',
            detected_reason TEXT DEFAULT '',
            machine_status TEXT DEFAULT '',
            prog_raw TEXT DEFAULT '',
            confirmed_at TEXT DEFAULT '',
            ignored_at TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            timestamp TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS tool_replacement_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            candidate_id INTEGER DEFAULT 0,
            machine TEXT NOT NULL,
            tid TEXT NOT NULL,
            tool_name TEXT DEFAULT '',
            reason TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            source TEXT DEFAULT '',
            old_cycle_id INTEGER DEFAULT 0,
            new_cycle_id INTEGER DEFAULT 0,
            timestamp TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS tool_inspections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine TEXT NOT NULL,
            tid TEXT NOT NULL,
            tool_name TEXT DEFAULT '',
            condition TEXT DEFAULT 'GOOD',
            inspected_by TEXT DEFAULT '',
            observed_wear_pct REAL DEFAULT 0,
            measured_diameter REAL DEFAULT 0,
            measured_length_offset REAL DEFAULT 0,
            edge_notes TEXT DEFAULT '',
            action_taken TEXT DEFAULT '',
            calibrate_health INTEGER DEFAULT 0,
            old_wear_score REAL DEFAULT 0,
            new_wear_score REAL DEFAULT 0,
            old_max_life_min REAL DEFAULT 0,
            new_max_life_min REAL DEFAULT 0,
            active_cycle_id INTEGER DEFAULT 0,
            notes TEXT DEFAULT '',
            timestamp TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        # ------------------------
        # Jobs / scheduling
        # ------------------------
        conn.execute("""CREATE TABLE IF NOT EXISTS job_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine TEXT NOT NULL,
            job_num TEXT,
            part_name TEXT,
            prog_raw TEXT,
            started_at TEXT DEFAULT CURRENT_TIMESTAMP,
            last_seen TEXT DEFAULT CURRENT_TIMESTAMP,
            cycle_time_sec REAL DEFAULT 0,
            estimated_total_sec REAL DEFAULT 0,
            avg_load REAL DEFAULT 0,
            max_load REAL DEFAULT 0,
            avg_rpm REAL DEFAULT 0,
            load_samples INTEGER DEFAULT 0,
            status TEXT DEFAULT 'RUNNING'
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS efficiency_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine TEXT NOT NULL,
            period TEXT NOT NULL,
            period_key TEXT NOT NULL,
            cutting_sec REAL DEFAULT 0,
            idle_sec REAL DEFAULT 0,
            alarm_sec REAL DEFAULT 0,
            parts_run INTEGER DEFAULT 0,
            avg_spindle_load REAL DEFAULT 0,
            load_samples INTEGER DEFAULT 0,
            UNIQUE(machine, period, period_key)
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS part_catalog (
            job_id TEXT PRIMARY KEY,
            weight REAL,
            com_x REAL,
            com_y REAL,
            com_z REAL,
            due_date TEXT DEFAULT '',
            steel_sheet_url TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS job_signature_components (
            job_num TEXT NOT NULL,
            component_role TEXT NOT NULL,
            quote_name TEXT DEFAULT '',
            cad_component TEXT DEFAULT '',
            clean_name TEXT DEFAULT '',
            length REAL DEFAULT 0,
            width REAL DEFAULT 0,
            thickness REAL DEFAULT 0,
            mass REAL DEFAULT 0,
            center_x REAL DEFAULT 0,
            center_y REAL DEFAULT 0,
            center_z REAL DEFAULT 0,
            has_center INTEGER DEFAULT 0,
            source_file TEXT DEFAULT '',
            imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (job_num, component_role)
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS match_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT DEFAULT CURRENT_TIMESTAMP,
            specs TEXT,
            matches TEXT
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS schedule (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_id TEXT,
            due_date TEXT,
            steel_sheet_url TEXT DEFAULT '',
            priority INTEGER DEFAULT 5,
            status TEXT DEFAULT 'QUEUED',
            notes TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS job_series (
            job_num TEXT PRIMARY KEY,
            due_date TEXT DEFAULT '',
            steel_sheet_url TEXT DEFAULT '',
            steel_sheet_file TEXT DEFAULT '',
            steel_sheet_original_name TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            status TEXT DEFAULT 'ACTIVE',
            matched_job_num TEXT DEFAULT '',
            matched_reason TEXT DEFAULT '',
            matched_at TEXT DEFAULT '',
            revenue REAL DEFAULT 0,
            steel_cost REAL DEFAULT 0,
            grinding_cost REAL DEFAULT 0,
            pyropel_cost REAL DEFAULT 0,
            components_cost REAL DEFAULT 0,
            labor_rate REAL DEFAULT 0,
            labor_override_cost REAL DEFAULT 0,
            other_cost REAL DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS job_components (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_num TEXT NOT NULL,
            component_name TEXT NOT NULL,
            machine TEXT DEFAULT '',
            status TEXT DEFAULT 'PENDING',
            sort_order INTEGER DEFAULT 0,
            estimated_sec REAL DEFAULT 0,
            actual_sec REAL DEFAULT 0,
            estimate_source TEXT DEFAULT '',
            estimate_basis TEXT DEFAULT '',
            estimate_updated_at TEXT DEFAULT '',
            started_at TEXT,
            completed_at TEXT,
            manual_complete INTEGER DEFAULT 0,
            last_seen_machine TEXT DEFAULT '',
            last_seen_program TEXT DEFAULT '',
            last_runtime_update TEXT,
            last_status_note TEXT DEFAULT '',
            UNIQUE(job_num, component_name)
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS job_status_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_num TEXT NOT NULL,
            component_name TEXT DEFAULT '',
            status TEXT DEFAULT '',
            machine TEXT DEFAULT '',
            prog_raw TEXT DEFAULT '',
            message TEXT DEFAULT '',
            timestamp TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS machine_schedule_plan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine TEXT NOT NULL,
            job_num TEXT NOT NULL,
            component_name TEXT NOT NULL,
            schedule_rank INTEGER DEFAULT 0,
            due_date TEXT DEFAULT '',
            steel_ready INTEGER DEFAULT 0,
            steel_status TEXT DEFAULT '',
            steel_item_count INTEGER DEFAULT 0,
            steel_materials TEXT DEFAULT '',
            estimated_sec REAL DEFAULT 0,
            setup_sec REAL DEFAULT 0,
            total_planned_sec REAL DEFAULT 0,
            planned_start TEXT DEFAULT '',
            planned_end TEXT DEFAULT '',
            priority_reason TEXT DEFAULT '',
            efficiency_reason TEXT DEFAULT '',
            status TEXT DEFAULT 'PLANNED',
            generated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        # ------------------------
        # Steel sheet parsing
        # ------------------------
        conn.execute("""CREATE TABLE IF NOT EXISTS steel_sheet_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_num TEXT NOT NULL,
            source_url TEXT DEFAULT '',
            source_type TEXT DEFAULT '',
            qty REAL DEFAULT 0,
            item_name TEXT DEFAULT '',
            dim_1 REAL DEFAULT 0,
            dim_2 REAL DEFAULT 0,
            dim_3 REAL DEFAULT 0,
            material TEXT DEFAULT '',
            heat_or_type TEXT DEFAULT '',
            raw_line TEXT DEFAULT '',
            confidence REAL DEFAULT 0,
            parsed_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS steel_sheet_parse_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_num TEXT DEFAULT '',
            source_url TEXT DEFAULT '',
            status TEXT DEFAULT '',
            message TEXT DEFAULT '',
            parsed_count INTEGER DEFAULT 0,
            raw_text_preview TEXT DEFAULT '',
            timestamp TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        # ------------------------
        # G-code analysis
        # ------------------------
        conn.execute("""CREATE TABLE IF NOT EXISTS gcode_tool_estimates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine TEXT DEFAULT '',
            program_key TEXT NOT NULL,
            program_path TEXT DEFAULT '',
            tid TEXT NOT NULL,

            estimated_cut_sec REAL DEFAULT 0,
            cut_distance_in REAL DEFAULT 0,
            rapid_distance_in REAL DEFAULT 0,
            plunge_distance_in REAL DEFAULT 0,

            avg_feed_ipm REAL DEFAULT 0,
            max_feed_ipm REAL DEFAULT 0,
            avg_rpm REAL DEFAULT 0,
            max_rpm REAL DEFAULT 0,

            max_depth_in REAL DEFAULT 0,
            avg_stepdown_in REAL DEFAULT 0,
            max_stepdown_in REAL DEFAULT 0,
            xy_cut_moves INTEGER DEFAULT 0,
            z_plunge_moves INTEGER DEFAULT 0,

            drill_cycles INTEGER DEFAULT 0,
            peck_cycles INTEGER DEFAULT 0,
            avg_peck_q_in REAL DEFAULT 0,

            coolant_on_seen INTEGER DEFAULT 0,
            severity_multiplier REAL DEFAULT 1.0,

            raw_json TEXT DEFAULT '',
            analyzed_at TEXT DEFAULT CURRENT_TIMESTAMP,

            UNIQUE(machine, program_key, tid)
        )""")

        # ------------------------
        # Material/tool rules
        # ------------------------
        conn.execute("""CREATE TABLE IF NOT EXISTS material_profiles (
            material TEXT PRIMARY KEY,
            hardness_note TEXT DEFAULT '',
            tool_life_factor REAL DEFAULT 1.0,
            notes TEXT DEFAULT '',
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS tool_life_rules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT NOT NULL,
            work_material TEXT DEFAULT '4140',
            diameter_min REAL DEFAULT 0,
            diameter_max REAL DEFAULT 999,
            max_life_min REAL DEFAULT 60,
            notes TEXT DEFAULT '',
            UNIQUE(category, work_material, diameter_min, diameter_max)
        )""")

        # ------------------------
        # Telemetry
        # ------------------------
        conn.execute("""CREATE TABLE IF NOT EXISTS telemetry_samples (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine TEXT NOT NULL,
            source TEXT DEFAULT '',
            status TEXT DEFAULT '',
            prog_raw TEXT DEFAULT '',
            tool TEXT DEFAULT '',
            spindle_rpm REAL DEFAULT 0,
            feed_rate REAL DEFAULT 0,
            spindle_load REAL DEFAULT 0,
            actual_cutting INTEGER DEFAULT 0,
            mtconnect_ok INTEGER DEFAULT 0,
            mtconnect_error TEXT DEFAULT '',
            timestamp TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        # ------------------------
        # Autoloader simulation
        # ------------------------
        conn.execute("""CREATE TABLE IF NOT EXISTS autoloader_shelves (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shelf_code TEXT UNIQUE NOT NULL,
            shelf_row INTEGER DEFAULT 0,
            shelf_col INTEGER DEFAULT 0,
            shelf_level INTEGER DEFAULT 0,
            status TEXT DEFAULT 'EMPTY',
            assigned_job_num TEXT DEFAULT '',
            component_name TEXT DEFAULT '',
            material TEXT DEFAULT '',
            pallet_id TEXT DEFAULT '',
            fixture_id TEXT DEFAULT '',
            part_present_confidence REAL DEFAULT 0,
            safe_to_run INTEGER DEFAULT 0,
            locked INTEGER DEFAULT 0,
            notes TEXT DEFAULT '',
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS autoloader_queue (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            requested_job_num TEXT DEFAULT '',
            component_name TEXT DEFAULT '',
            from_shelf TEXT DEFAULT '',
            to_machine TEXT DEFAULT '',
            requested_action TEXT DEFAULT 'LOAD',
            status TEXT DEFAULT 'SIMULATION_ONLY',
            safety_notes TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS autoloader_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_type TEXT DEFAULT '',
            shelf_code TEXT DEFAULT '',
            machine TEXT DEFAULT '',
            job_num TEXT DEFAULT '',
            component_name TEXT DEFAULT '',
            message TEXT DEFAULT '',
            timestamp TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        # ------------------------
        # Shop board
        # ------------------------
        conn.execute("""CREATE TABLE IF NOT EXISTS shop_board_rows (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_num TEXT DEFAULT '',
            description TEXT DEFAULT '',
            due_date TEXT DEFAULT '',
            priority INTEGER DEFAULT 5,
            row_status TEXT DEFAULT 'OPEN',
            row_color TEXT DEFAULT '',
            sort_order INTEGER DEFAULT 0,

            finals_mex TEXT DEFAULT '',
            finals_john TEXT DEFAULT '',

            steel TEXT DEFAULT '',
            ht200 TEXT DEFAULT '',
            xt_sent TEXT DEFAULT '',
            pyropel TEXT DEFAULT '',
            j_blk TEXT DEFAULT '',
            a2_1 TEXT DEFAULT '',
            req_1 TEXT DEFAULT '',
            d2 TEXT DEFAULT '',
            req_2 TEXT DEFAULT '',
            other TEXT DEFAULT '',

            pullcore_cam TEXT DEFAULT '',
            pullcore_key TEXT DEFAULT '',

            h13 TEXT DEFAULT '',
            req_3 TEXT DEFAULT '',
            a2_2 TEXT DEFAULT '',
            req_4 TEXT DEFAULT '',

            pdf_mc TEXT DEFAULT '',
            prog TEXT DEFAULT '',
            notes TEXT DEFAULT '',

            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        conn.execute("""CREATE TABLE IF NOT EXISTS shop_board_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            row_id INTEGER DEFAULT 0,
            job_num TEXT DEFAULT '',
            field_name TEXT DEFAULT '',
            old_value TEXT DEFAULT '',
            new_value TEXT DEFAULT '',
            message TEXT DEFAULT '',
            timestamp TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        # ------------------------
        # Legacy schema upgrades
        # ------------------------
        for col in [
            "due_date TEXT DEFAULT ''",
            "steel_sheet_url TEXT DEFAULT ''",
        ]:
            _add_col(conn, "part_catalog", col)

        for col in [
            "steel_sheet_url TEXT DEFAULT ''",
            "steel_sheet_file TEXT DEFAULT ''",
            "steel_sheet_original_name TEXT DEFAULT ''",
            "notes TEXT DEFAULT ''",
            "status TEXT DEFAULT 'ACTIVE'",
            "matched_job_num TEXT DEFAULT ''",
            "matched_reason TEXT DEFAULT ''",
            "matched_at TEXT DEFAULT ''",
            "revenue REAL DEFAULT 0",
            "steel_cost REAL DEFAULT 0",
            "grinding_cost REAL DEFAULT 0",
            "pyropel_cost REAL DEFAULT 0",
            "components_cost REAL DEFAULT 0",
            "labor_rate REAL DEFAULT 0",
            "labor_override_cost REAL DEFAULT 0",
            "other_cost REAL DEFAULT 0",
            "updated_at TEXT DEFAULT ''",
        ]:
            _add_col(conn, "job_series", col)

        for col in [
            "estimated_sec REAL DEFAULT 0",
            "actual_sec REAL DEFAULT 0",
            "estimate_source TEXT DEFAULT ''",
            "estimate_basis TEXT DEFAULT ''",
            "estimate_updated_at TEXT DEFAULT ''",
            "last_seen_machine TEXT DEFAULT ''",
            "last_seen_program TEXT DEFAULT ''",
            "last_runtime_update TEXT",
            "last_status_note TEXT DEFAULT ''",
        ]:
            _add_col(conn, "job_components", col)

        for col in [
            "steel_item_count INTEGER DEFAULT 0",
            "steel_materials TEXT DEFAULT ''",
        ]:
            _add_col(conn, "machine_schedule_plan", col)

        # ------------------------
        # Machine rename cleanup
        # ------------------------
        for old, new in MACHINE_RENAME_MAP.items():
            for table in [
                "tool_health",
                "job_runs",
                "efficiency_log",
                "tool_life_cycles",
                "tool_cut_samples",
                "machine_schedule_plan",
            ]:
                try:
                    conn.execute(f"UPDATE {table} SET machine=? WHERE machine=?", (new, old))
                except Exception:
                    pass

            try:
                conn.execute("UPDATE job_components SET machine=? WHERE machine=?", (new, old))
                conn.execute("UPDATE job_components SET last_seen_machine=? WHERE last_seen_machine=?", (new, old))
            except Exception:
                pass

        # ------------------------
        # Seed tool catalog
        # ------------------------
        conn.executemany("""
            INSERT OR IGNORE INTO tool_catalog (
                tid, tool_name, category, created_at, updated_at
            )
            VALUES (?, ?, ?, datetime('now'), datetime('now'))
        """, [(tid, name, infer_tool_category(name)) for tid, name in TOOL_CATALOG_SEED])

        conn.execute("""
            INSERT OR IGNORE INTO material_profiles (
                material, hardness_note, tool_life_factor, notes
            )
            VALUES (
                '4140',
                '4140 / 4140 prehard varies widely; tune per supplier/hardness',
                0.72,
                'Default conservative tool-life factor for alloy steel versus mild steel.'
            )
        """)

        default_rules = [
            ("Tap", "4140", 0, 999, 12, "Conservative tapping life in 4140."),
            ("Thread Mill", "4140", 0, 999, 28, "Thread mills generally safer than taps."),
            ("Drill", "4140", 0, 0.25, 22, "Small drills in 4140."),
            ("Drill", "4140", 0.25, 0.75, 35, "Medium drills in 4140."),
            ("Drill", "4140", 0.75, 999, 55, "Large drills/spades in 4140."),
            ("Spot Drill", "4140", 0, 999, 70, "Spot drills usually low engagement."),
            ("Endmill", "4140", 0, 0.1875, 18, "Tiny carbide endmills."),
            ("Endmill", "4140", 0.1875, 0.5, 40, "Small carbide endmills."),
            ("Endmill", "4140", 0.5, 999, 60, "Larger endmills."),
            ("Ball Endmill", "4140", 0, 999, 35, "Ball tools wear at contact point."),
            ("Face Mill", "4140", 0, 999, 110, "Inserted face milling."),
            ("Inserted Cutter", "4140", 0, 999, 95, "Inserted roughing/high-feed cutters."),
            ("Chamfer Mill", "4140", 0, 999, 55, "Chamfering/deburring tools."),
            ("Tool", "4140", 0, 999, 40, "Fallback rule."),
        ]

        conn.executemany("""
            INSERT OR IGNORE INTO tool_life_rules (
                category, work_material, diameter_min, diameter_max, max_life_min, notes
            )
            VALUES (?, ?, ?, ?, ?, ?)
        """, default_rules)

        for tid, name in TOOL_CATALOG_SEED:
            cat = infer_tool_category(name)
            dia = infer_tool_diameter(name)
            flutes = infer_flutes(name)
            life = estimate_default_tool_life_4140_min(name, cat)

            conn.execute("""
                UPDATE tool_catalog
                SET
                    category=CASE WHEN category IS NULL OR category='' THEN ? ELSE category END,
                    diameter=CASE WHEN COALESCE(diameter,0)=0 THEN ? ELSE diameter END,
                    flutes=CASE WHEN COALESCE(flutes,0)=0 THEN ? ELSE flutes END,
                    work_material=CASE WHEN work_material IS NULL OR work_material='' THEN '4140' ELSE work_material END,
                    max_life_min=CASE WHEN COALESCE(max_life_min,0)=0 THEN ? ELSE max_life_min END,
                    updated_at=datetime('now')
                WHERE tid=?
            """, (cat, dia, flutes, life, tid))

        # Seed 5-high x 4-wide virtual shelf positions A1-D5.
        shelf_rows = []
        for level in range(1, 6):
            for col_idx, col_letter in enumerate(["A", "B", "C", "D"], start=1):
                code = f"{col_letter}{level}"
                shelf_rows.append((code, level, col_idx, level))

        conn.executemany("""
            INSERT OR IGNORE INTO autoloader_shelves (
                shelf_code, shelf_row, shelf_col, shelf_level,
                status, notes, updated_at
            )
            VALUES (?, ?, ?, ?, 'EMPTY', 'Auto-created virtual shelf slot. Simulation only.', datetime('now'))
        """, shelf_rows)

        conn.commit()


# Initialize DB before web app starts.
init_db()


# ============================================================
# REQUEST MODELS
# ============================================================
class ScanRequest(BaseModel):
    sku: str
    qty: int = 1
    direction: str


class InventoryItemRequest(BaseModel):
    sku: str
    description: str
    category: str = "Tool"
    quantity: int = 0
    min_threshold: int = 5


class AutoScanRequest(BaseModel):
    barcode: str
    qty: int = 1
    direction: str = "IN"


class QRResolveRequest(BaseModel):
    barcode: str


class QRActionRequest(BaseModel):
    barcode: str
    direction: str
    qty: int = 1
    sku: str = ""
    description: str = ""
    category: str = "Tool"
    starting_quantity: int = 0
    min_threshold: int = 5


class QRAliasRequest(BaseModel):
    alias: str
    sku: str
    description: str = ""
    category: str = "Tool"
    quantity: int = 0
    min_threshold: int = 5


class AdminLoginRequest(BaseModel):
    password: str


class ToolReplaceRequest(BaseModel):
    machine: str
    tid: str
    reason: str = "Manual replacement"
    notes: str = ""


class ToolCandidateActionRequest(BaseModel):
    notes: str = ""
    reason: str = "Operator confirmed actual tool replacement after offset change"


class ToolInspectionRequest(BaseModel):
    machine: str
    tid: str
    condition: str = "GOOD"
    inspected_by: str = ""
    observed_wear_pct: float = 25.0
    measured_diameter: float = 0.0
    measured_length_offset: float = 0.0
    edge_notes: str = ""
    action_taken: str = "INSPECTED"
    calibrate_health: int = 1
    notes: str = ""


class JobRenameRequest(BaseModel):
    new_job_num: str
    merge_if_exists: bool = True


class JobDeleteRequest(BaseModel):
    confirm: str = ""


class ShelfUpdateRequest(BaseModel):
    shelf_row: int = 0
    shelf_col: int = 0
    shelf_level: int = 0
    status: str = "EMPTY"
    assigned_job_num: str = ""
    component_name: str = ""
    material: str = ""
    pallet_id: str = ""
    fixture_id: str = ""
    part_present_confidence: float = 0.0
    safe_to_run: int = 0
    locked: int = 0
    notes: str = ""


class AutoloaderPlanRequest(BaseModel):
    requested_job_num: str = ""
    component_name: str = ""
    from_shelf: str = ""
    to_machine: str = ""
    requested_action: str = "LOAD"


class ShopBoardRowRequest(BaseModel):
    job_num: str = ""
    description: str = ""
    due_date: str = ""
    priority: int = 5
    row_status: str = "OPEN"
    row_color: str = ""
    sort_order: int = 0

    finals_mex: str = ""
    finals_john: str = ""

    steel: str = ""
    ht200: str = ""
    xt_sent: str = ""
    pyropel: str = ""
    j_blk: str = ""
    a2_1: str = ""
    req_1: str = ""
    d2: str = ""
    req_2: str = ""
    other: str = ""

    pullcore_cam: str = ""
    pullcore_key: str = ""

    h13: str = ""
    req_3: str = ""
    a2_2: str = ""
    req_4: str = ""

    pdf_mc: str = ""
    prog: str = ""
    notes: str = ""


# ============================================================
# STATE / APP SETUP
# ============================================================
machine_states = {
    cfg["name"]: {
        "status": "OFFLINE",
        "prog": "N/A",
        "tool": "0",
        "load": 0.0,
        "uptime": "00:00:00",
        "cycle": "00:00:00",
        "cycle_sec": 0.0,
        "spindle_rpm": 0.0,
        "rpm_signal_ok": False,
        "rpm_source": "",
        "feed_rate": 0.0,
        "l_wear": 0.0,
        "d_wear": 0.0,
        "tool_name": "",
        "job_num": None,
        "part_name": None,
        "sub_cat": None,
        "prog_display": "IDLE",
        "actual_cutting": False,
        "raw_mode": "",
        "data_source": "",
        "mtconnect_ok": False,
        "mtconnect_url": cfg.get("mtconnect_url", ""),
        "mtconnect_error": "",
        "part_count": 0,

        # Program runtime/progress estimate.
        "program_est_sec": 0.0,
        "program_progress_pct": 0.0,
        "program_remaining_sec": 0.0,
        "program_est_finish": "",
        "program_progress_source": "",

        "_cut_sec": 0.0,
        "_idle_sec": 0.0,
        "_alarm_sec": 0.0,
        "_load_sum": 0.0,
        "_load_n": 0,
        "_last_prog": None,
        "_current_job_id": None,
        "_last_flush": time.time(),
        "_consec_offline": 0,

        # RPM idle logic.
        "_no_spindle_since": time.time(),
        "_last_nonzero_rpm": 0.0,
        "_last_nonzero_rpm_at": 0.0,

        # Tool-number hold logic, prevents temporary T00 display.
        "_last_valid_tool": "0",
        "_last_valid_tool_at": 0.0,

        "_last_component_key": None,
        "_last_cycle_sec_for_component": 0.0,
        "_offset_baseline_ready": set(),
        "_offset_candidate_streaks": {},
        "_last_progress_prog": "",
        "_program_started_at": 0.0,

        # Old/template program mapping cache.
        # Example: controller runs old J8206, but current shop job is J8342.
        # Key: "8206|ID POT" -> "8342"
        "_effective_job_for_template": {},
    }
    for cfg in MACHINE_CONFIGS
}

db_lock = threading.Lock()


class ConnectionManager:
    def __init__(self):
        self.active_connections: List[WebSocket] = []

    async def connect(self, ws):
        await ws.accept()
        self.active_connections.append(ws)

    def disconnect(self, ws):
        if ws in self.active_connections:
            self.active_connections.remove(ws)

    async def broadcast(self, msg):
        dead = []
        for c in self.active_connections:
            try:
                await c.send_text(msg)
            except Exception:
                dead.append(c)
        for d in dead:
            self.disconnect(d)


manager = ConnectionManager()
app = FastAPI(title=APP_NAME)


@app.on_event("startup")
async def startup_event():
    asyncio.create_task(broadcast_loop())


# ============================================================
# END PART 2
# PART 3 CONTINUES WITH:
# - Haas Q-code helpers
# - MTConnect helpers
# - tool health helpers
# - G-code analysis helpers
# - steel-sheet parsing helpers
# ============================================================
# ============================================================
# HAAS Q-CODE HELPERS
# ============================================================
def clean_haas_timer(raw: str, default: str = "00:00:00") -> str:
    if not raw:
        return default

    s = str(raw).strip().replace("\x00", "").strip()
    if not s:
        return default

    if s.startswith("?Q"):
        return default

    if s.upper() in ("?", ">", "N/A", "NULL", "NONE"):
        return default

    if re.match(r"^\d+:\d+(:\d+)?$", s):
        return s

    try:
        float(s)
        return s
    except Exception:
        return default


def parse_cycle_time(raw: str) -> float:
    if not raw or raw.strip() in ("N/A", ""):
        return 0.0

    raw = clean_haas_timer(raw)
    raw = raw.strip().replace(">", "").strip()

    m = re.match(r"^(\d+):(\d+):(\d+)$", raw)
    if m:
        return int(m.group(1)) * 3600 + int(m.group(2)) * 60 + int(m.group(3))

    m2 = re.match(r"^(\d+):(\d+)$", raw)
    if m2:
        return int(m2.group(1)) * 60 + int(m2.group(2))

    try:
        return float(raw)
    except Exception:
        return 0.0


def _is_active(mode: str) -> bool:
    m = (mode or "").strip().upper()
    return any(k in m for k in ("CYCLE", "RUNNING", "IN CYCLE", "FEED", "ACTIVE", "EXECUTING"))


def _is_alarm_mode(mode: str) -> bool:
    m = (mode or "").strip().upper()
    if not m:
        return False

    non_alarm_markers = [
        "NO ALARM",
        "NO ALARMS",
        "ALARM: 0",
        "ALARMS: 0",
        "0 ALARM",
        "0 ALARMS",
    ]

    if any(x in m for x in non_alarm_markers):
        return False

    if "E-STOP" in m or "ESTOP" in m or "EMERGENCY" in m:
        return True
    if "ALARM" in m:
        return True
    if "FAULT" in m:
        return True

    return False


def query_haas(ip: str, port: int, cmd: str) -> str:
    try:
        cmd_clean = cmd.strip()

        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(2.0)
            s.connect((ip, port))

            try:
                s.recv(4096)
            except Exception:
                pass

            s.sendall(f"{cmd_clean}\r\n".encode())
            time.sleep(0.12)

            raw = b""
            deadline = time.time() + 1.5

            while time.time() < deadline:
                try:
                    chunk = s.recv(4096)
                    if not chunk:
                        break
                    raw += chunk
                    if b">" in raw:
                        break
                except Exception:
                    break

            text = raw.decode(errors="replace")
            text = text.replace("\x00", "").strip()

            if not text:
                return ""

            lines = []
            for line in text.splitlines():
                line = line.strip().rstrip(">").strip()
                if line:
                    lines.append(line)

            if not lines:
                return ""

            cmd_upper = cmd_clean.upper()
            candidates = []

            for line in lines:
                u = line.upper().strip()

                if u == cmd_upper:
                    continue

                if "," in line:
                    parts = [p.strip() for p in line.split(",") if p.strip()]
                    if parts:
                        candidates.append(parts[-1])
                        continue

                if u.startswith(cmd_upper):
                    stripped = line[len(cmd_clean):].strip(" ,:\t")
                    if stripped:
                        candidates.append(stripped)
                    continue

                candidates.append(line)

            if not candidates:
                return ""

            answer = candidates[-1].strip().rstrip(">").strip()

            if answer.upper() == cmd_upper:
                return ""

            return answer

    except Exception:
        return ""


def safe_float(val: str) -> float:
    try:
        return float((val or "").strip().replace("\x00", ""))
    except Exception:
        return 0.0


def query_haas_float(ip: str, port: int, cmd: str) -> float:
    raw = query_haas(ip, port, cmd)
    if not raw:
        return 0.0

    s = str(raw).strip().replace("\x00", "").strip()
    if not s:
        return 0.0

    if s.upper().startswith("?Q") or s.upper().startswith("Q"):
        if "," not in s:
            return 0.0

    direct = safe_float(s)
    if direct != 0.0:
        return direct

    nums = re.findall(r"[-+]?\d+(?:\.\d+)?", s)
    if not nums:
        return 0.0

    try:
        return float(nums[-1])
    except Exception:
        return 0.0


def read_spindle_rpm(ip: str, port: int, state: dict):
    now = time.time()

    for reg in RPM_REGISTER_CANDIDATES:
        val = query_haas_float(ip, port, f"?Q600 {reg}")

        if 1 <= val <= 60000:
            if val > 0:
                state["_last_nonzero_rpm"] = val
                state["_last_nonzero_rpm_at"] = now
            return val, True, f"Q600 {reg}"

    last = float(state.get("_last_nonzero_rpm", 0.0) or 0.0)
    last_at = float(state.get("_last_nonzero_rpm_at", 0.0) or 0.0)

    if last > 0 and now - last_at <= RPM_HOLD_LAST_GOOD_SEC:
        return last, False, "HELD_LAST_GOOD"

    return 0.0, False, "NO_RPM_SIGNAL"


def read_qcode_snapshot(ip: str, port: int, state: dict) -> dict:
    mode_raw = query_haas(ip, port, "?Q104")
    q_ok = bool(mode_raw and mode_raw.strip() not in ("", "?", ">"))

    prog_raw = query_haas(ip, port, "?Q102") if q_ok else ""
    uptime_raw = clean_haas_timer(query_haas(ip, port, "?Q501")) if q_ok else "00:00:00"
    cycle_raw = clean_haas_timer(query_haas(ip, port, "?Q502")) if q_ok else "00:00:00"
    tool_raw = query_haas(ip, port, "?Q201") if q_ok else "0"

    spindle_load = query_haas_float(ip, port, "?Q600 1098") if q_ok else 0.0
    path_feed = query_haas_float(ip, port, "?Q600 4109") if q_ok else 0.0
    spindle_rpm, rpm_signal_ok, rpm_source = read_spindle_rpm(ip, port, state) if q_ok else (0.0, False, "")

    # Prevent temporary T00 / 00 / 0 reads from becoming the active displayed tool.
    tool = apply_tool_hold_last_good(tool_raw, state)

    return {
        "ok": q_ok,
        "source": "QCODE",
        "status": mode_raw.strip().upper() if q_ok else "OFFLINE",
        "raw_mode": mode_raw.strip().upper() if q_ok else "",
        "prog": prog_raw or "N/A",
        "uptime": uptime_raw,
        "cycle": cycle_raw,
        "cycle_sec": parse_cycle_time(cycle_raw),
        "tool": tool,
        "load": spindle_load,
        "feed_rate": path_feed,
        "spindle_rpm": spindle_rpm,
        "rpm_signal_ok": rpm_signal_ok,
        "rpm_source": rpm_source,
    }


# ============================================================
# MTCONNECT HELPERS
# ============================================================
MTCONNECT_TIMEOUT_SEC = 4.0

MT_AVAILABILITY_KEYS = ["availability"]

MT_PROGRAM_KEYS = [
    "program",
    "activeprogram",
    "mainprogram",
    "programname",
    "program_name",
    "path_program",
    "programcomment",
    "activeprogramname",
]

MT_EXECUTION_KEYS = ["execution", "exec"]
MT_MODE_KEYS = ["controllermode", "controller_mode", "mode"]
MT_TOOL_KEYS = ["toolnumber", "tool_number", "currenttool", "current_tool", "tool"]

MT_SPINDLE_RPM_KEYS = [
    "spindlespeed",
    "spindle_speed",
    "spindle",
    "srpm",
    "rotaryvelocity",
    "rotary_velocity",
]

MT_FEED_KEYS = [
    "pathfeedrate",
    "path_feedrate",
    "feedrate",
    "feed",
    "fact",
]

MT_LOAD_KEYS = [
    "load",
    "spindleload",
    "spindle_load",
]

MT_PART_COUNT_KEYS = [
    "partcount",
    "part_count",
]

MT_EMERGENCY_KEYS = [
    "emergencystop",
    "emergency_stop",
    "estop",
]

MT_TIMER_KEYS = [
    "accumulatedtime",
    "accumulated_time",
    "cycletime",
    "cycle_time",
]


def mtconnect_candidate_urls(cfg: dict) -> list:
    ip = cfg.get("ip", "")
    configured = cfg.get("mtconnect_url", "") or ""

    candidates = [
        configured,
        f"http://{ip}:8082/current",
        f"http://{ip}:8082/current?path=//Devices",
        f"http://{ip}:5000/current",
        f"http://{ip}:5000/current?path=//Devices",
        f"http://{ip}/current",
        f"http://{ip}/current?path=//Devices",
    ]

    out = []
    seen = set()

    for u in candidates:
        u = str(u or "").strip()
        if not u:
            continue
        if u in seen:
            continue
        seen.add(u)
        out.append(u)

    return out


def _strip_ns(tag: str) -> str:
    return str(tag or "").split("}", 1)[-1]


def _mt_norm_key(v: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(v or "").lower())


def _mt_clean_value(v: str) -> str:
    s = str(v or "").replace("\x00", "").strip()
    if not s:
        return ""
    if s.upper() in ("UNAVAILABLE", "NOT_READY", "NOTREADY", "UNKNOWN", "N/A", "NULL"):
        return ""
    return s


def _mt_float_value(v) -> float:
    s = _mt_clean_value(v)
    if not s:
        return 0.0
    nums = re.findall(r"[-+]?\d+(?:\.\d+)?", s)
    if not nums:
        return 0.0
    try:
        return float(nums[-1])
    except Exception:
        return 0.0


def fetch_mtconnect_current(url: str) -> dict:
    if not url:
        return {}

    try:
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": f"{APP_SLUG}/1.0",
                "Accept": "application/xml,text/xml,*/*",
            },
        )

        with urllib.request.urlopen(req, timeout=MTCONNECT_TIMEOUT_SEC) as resp:
            raw = resp.read(2_000_000)

        if not raw:
            return {}

        root = ET.fromstring(raw)
        out = {
            "_ok": True,
            "_url": url,
            "_sha1": hashlib.sha1(raw).hexdigest(),
            "_received_at": datetime.now().isoformat(timespec="seconds"),
        }

        condition_faults = []
        condition_warnings = []

        for elem in root.iter():
            tag = _strip_ns(elem.tag)
            tag_norm = _mt_norm_key(tag)
            text = _mt_clean_value(elem.text)

            if tag_norm in ("fault", "warning"):
                native_code = elem.attrib.get("nativeCode", "")
                qual = elem.attrib.get("qualifier", "")
                msg = text or elem.attrib.get("nativeSeverity", "") or tag
                packed = " ".join(x for x in [tag.upper(), native_code, qual, msg] if x)
                if tag_norm == "fault":
                    condition_faults.append(packed)
                else:
                    condition_warnings.append(packed)

            if not text:
                continue

            possible_keys = set()
            possible_keys.add(tag)

            for attr in ("name", "dataItemId", "id", "subType", "type"):
                if elem.attrib.get(attr):
                    possible_keys.add(elem.attrib.get(attr))

            for k in possible_keys:
                nk = _mt_norm_key(k)
                if not nk:
                    continue

                out[nk] = text
                out[f"raw_{nk}"] = {
                    "tag": tag,
                    "value": text,
                    "attrib": dict(elem.attrib),
                }

        out["_condition_faults"] = condition_faults
        out["_condition_warnings"] = condition_warnings
        return out

    except urllib.error.HTTPError as e:
        return {
            "_ok": False,
            "_url": url,
            "_error": f"HTTP {e.code}: {e.reason}",
            "_received_at": datetime.now().isoformat(timespec="seconds"),
        }

    except urllib.error.URLError as e:
        return {
            "_ok": False,
            "_url": url,
            "_error": f"URL error: {e.reason}",
            "_received_at": datetime.now().isoformat(timespec="seconds"),
        }

    except ET.ParseError as e:
        return {
            "_ok": False,
            "_url": url,
            "_error": f"XML parse error: {e}",
            "_received_at": datetime.now().isoformat(timespec="seconds"),
        }

    except Exception as e:
        return {
            "_ok": False,
            "_url": url,
            "_error": f"{type(e).__name__}: {e}",
            "_received_at": datetime.now().isoformat(timespec="seconds"),
        }


def mt_get(data: dict, keys: list, default: str = "") -> str:
    if not data:
        return default

    for k in keys:
        nk = _mt_norm_key(k)
        val = _mt_clean_value(data.get(nk, ""))
        if val:
            return val

    normalized_keys = [_mt_norm_key(k) for k in keys]
    for dk, dv in data.items():
        if dk.startswith("_") or dk.startswith("raw_"):
            continue
        ndk = _mt_norm_key(dk)
        for target in normalized_keys:
            if target and (target == ndk or target in ndk or ndk in target):
                val = _mt_clean_value(dv)
                if val:
                    return val

    return default


def mt_get_float(data: dict, keys: list, default: float = 0.0) -> float:
    return _mt_float_value(mt_get(data, keys, "")) or default


def mt_program_is_valid(v: str) -> bool:
    s = _mt_clean_value(v)
    if not s:
        return False
    u = s.upper()
    if u in ("0", "NONE", "IDLE", "UNAVAILABLE", "NOT_READY", "UNKNOWN"):
        return False
    return True


def choose_program_from_sources(mt_data: dict, q_program: str = "") -> str:
    mt_prog = mt_get(mt_data, MT_PROGRAM_KEYS, "")
    if mt_program_is_valid(mt_prog):
        return mt_prog

    for dk, dv in (mt_data or {}).items():
        if dk.startswith("_") or dk.startswith("raw_"):
            continue
        if "program" in dk:
            if mt_program_is_valid(str(dv)):
                return str(dv)

    if mt_program_is_valid(q_program):
        return q_program

    return "N/A"


def derive_status_from_mtconnect(mt_data: dict) -> str:
    if not mt_data or not mt_data.get("_ok"):
        return "OFFLINE"

    availability = mt_get(mt_data, MT_AVAILABILITY_KEYS, "")
    emergency = mt_get(mt_data, MT_EMERGENCY_KEYS, "")
    execution = mt_get(mt_data, MT_EXECUTION_KEYS, "")
    mode = mt_get(mt_data, MT_MODE_KEYS, "")

    faults = mt_data.get("_condition_faults") or []

    if availability and availability.upper() not in ("AVAILABLE", "READY"):
        return "OFFLINE"

    if emergency and emergency.upper() not in ("ARMED", "NORMAL", "READY"):
        return f"E-STOP {emergency}"

    if faults:
        return "ALARM " + faults[0][:80]

    if execution:
        return execution.upper()

    if mode:
        return mode.upper()

    return "AVAILABLE"


def read_mtconnect_snapshot(cfg: dict, state: dict, q_program_fallback: str = "") -> dict:
    if not cfg.get("use_mtconnect", True):
        return {
            "ok": False,
            "source": "MTCONNECT",
            "status": "OFFLINE",
            "mtconnect_error": "MTConnect disabled in machine config",
            "mtconnect_url": "",
        }

    urls = mtconnect_candidate_urls(cfg)

    if not urls:
        return {
            "ok": False,
            "source": "MTCONNECT",
            "status": "OFFLINE",
            "mtconnect_error": "No MTConnect URL configured",
            "mtconnect_url": "",
        }

    mt = {}
    good_url = ""
    last_error = ""

    for url in urls:
        test = fetch_mtconnect_current(url)
        logger.info("MTConnect test %s -> ok=%s error=%s", url, test.get("_ok"), test.get("_error", ""))

        if test.get("_ok"):
            mt = test
            good_url = url
            break

        last_error = test.get("_error", "") or "Unknown MTConnect error"

    if not mt.get("_ok"):
        return {
            "ok": False,
            "source": "MTCONNECT",
            "status": "OFFLINE",
            "mtconnect_error": last_error,
            "mtconnect_url": urls[0] if urls else "",
        }

    status = derive_status_from_mtconnect(mt)
    prog_raw = choose_program_from_sources(mt, q_program_fallback)

    tool_raw = mt_get(mt, MT_TOOL_KEYS, "")
    tool = apply_tool_hold_last_good(tool_raw, state)

    spindle_rpm = mt_get_float(mt, MT_SPINDLE_RPM_KEYS, 0.0)

    if spindle_rpm > 0:
        state["_last_nonzero_rpm"] = spindle_rpm
        state["_last_nonzero_rpm_at"] = time.time()

    load = mt_get_float(mt, MT_LOAD_KEYS, 0.0)
    feed = mt_get_float(mt, MT_FEED_KEYS, 0.0)
    timer = mt_get(mt, MT_TIMER_KEYS, "")
    part_count = int(mt_get_float(mt, MT_PART_COUNT_KEYS, 0.0))

    return {
        "ok": True,
        "source": "MTCONNECT",
        "status": status,
        "raw_mode": status,
        "prog": prog_raw or "N/A",
        "uptime": "00:00:00",
        "cycle": clean_haas_timer(timer) if timer else "00:00:00",
        "cycle_sec": parse_cycle_time(timer) if timer else 0.0,
        "tool": tool,
        "load": load,
        "feed_rate": feed,
        "spindle_rpm": spindle_rpm,
        "rpm_signal_ok": spindle_rpm > 0,
        "rpm_source": "MTCONNECT",
        "part_count": part_count,
        "mtconnect_raw": mt,
        "mtconnect_url": good_url,
        "mtconnect_error": "",
    }


def _valid_snapshot_value(v):
    if v is None:
        return False
    if isinstance(v, str):
        return bool(_mt_clean_value(v))
    if isinstance(v, (int, float)):
        return float(v) != 0.0
    return True


def merge_machine_snapshot(mt: dict, q: dict) -> dict:
    if mt.get("ok") and q.get("ok"):
        merged = dict(q)
        for k, v in mt.items():
            if k in ("prog",):
                if mt_program_is_valid(v):
                    merged[k] = v
            elif k in ("tool",):
                if normalize_tool_num(v) != "0":
                    merged[k] = v
            elif k in ("cycle", "cycle_sec"):
                if _valid_snapshot_value(v):
                    merged[k] = v
            elif k.startswith("mtconnect"):
                merged[k] = v
            elif k in (
                "status",
                "raw_mode",
                "load",
                "feed_rate",
                "spindle_rpm",
                "rpm_signal_ok",
                "rpm_source",
                "part_count",
            ):
                if _valid_snapshot_value(v):
                    merged[k] = v
        merged["source"] = "MTCONNECT+QCODE"
        return merged

    if mt.get("ok"):
        return mt

    if q.get("ok"):
        return q

    return {
        "ok": False,
        "source": "OFFLINE",
        "status": "OFFLINE",
        "raw_mode": "",
        "prog": "N/A",
        "uptime": "00:00:00",
        "cycle": "00:00:00",
        "cycle_sec": 0.0,
        "tool": "0",
        "load": 0.0,
        "feed_rate": 0.0,
        "spindle_rpm": 0.0,
        "rpm_signal_ok": False,
        "rpm_source": "",
        "part_count": 0,
    }


# ============================================================
# TOOL HEALTH HELPERS
# ============================================================
def tool_load_damage_multiplier(spindle_load: float) -> float:
    load = float(spindle_load or 0)
    if load <= 35:
        return 1.0
    if load <= 55:
        return 1.15
    if load <= 75:
        return 1.35
    if load <= 90:
        return 1.75
    return 2.35


def offset_wear_score(l_wear: float, d_wear: float, category: str) -> float:
    l = abs(float(l_wear or 0))
    d = abs(float(d_wear or 0))

    if category in ("Tap", "Thread Mill"):
        d_limit = 0.003
        l_limit = 0.010
    elif category in ("Drill", "Spot Drill"):
        d_limit = 0.006
        l_limit = 0.020
    else:
        d_limit = 0.010
        l_limit = 0.030

    score_d = (d / d_limit) * 100.0 if d_limit else 0
    score_l = (l / l_limit) * 100.0 if l_limit else 0
    return _clamp(max(score_d, score_l), 0, 100)


def get_tool_catalog_info(conn, tid: str, fallback_name: str = "") -> dict:
    tid = normalize_tool_num(tid)
    row = conn.execute("SELECT * FROM tool_catalog WHERE tid=?", (str(tid),)).fetchone()
    if row:
        return dict(row)

    name = fallback_name or get_tool_name(tid)
    cat = infer_tool_category(name)

    return {
        "tid": str(tid),
        "tool_name": name,
        "category": cat,
        "diameter": infer_tool_diameter(name),
        "flutes": infer_flutes(name),
        "max_life_min": estimate_default_tool_life_4140_min(name, cat),
        "work_material": DEFAULT_WORK_MATERIAL,
    }


def compute_tool_life_limit_min(conn, tid: str, tool_name: str) -> float:
    tid = normalize_tool_num(tid)
    info = get_tool_catalog_info(conn, tid, tool_name)
    configured = float(info.get("max_life_min") or 0)

    if configured > 0:
        return configured

    cat = info.get("category") or infer_tool_category(tool_name)
    return estimate_default_tool_life_4140_min(tool_name, cat)


def compute_wear_score(
    prior_cut_sec: float,
    cut_delta_sec: float,
    max_life_min: float,
    spindle_load: float,
    l_wear: float,
    d_wear: float,
    category: str,
    gcode_severity: float = 1.0,
) -> float:
    """
    Tool wear model.

    Default:
    - actual cutting time
    - spindle load factor
    - G-code severity factor
    - Haas offsets ignored unless ELGIN_USE_OFFSET_WEAR_IN_HEALTH/NEXUS_USE_OFFSET_WEAR_IN_HEALTH=1
    """
    load_mult = tool_load_damage_multiplier(spindle_load)
    gcode_mult = _clamp(float(gcode_severity or 1.0), 0.5, 3.0)

    damage_mult = load_mult * gcode_mult

    weighted_cut_sec = float(prior_cut_sec or 0) + float(cut_delta_sec or 0) * damage_mult
    time_score = (weighted_cut_sec / max(max_life_min * 60.0, 1.0)) * 100.0

    if USE_OFFSET_WEAR_IN_HEALTH:
        off_score = offset_wear_score(l_wear, d_wear, category)
        score = max(time_score, off_score)
    else:
        score = time_score

    return round(_clamp(score, 0, 100), 2)


def offset_read_looks_valid(l_wear: float, d_wear: float) -> bool:
    try:
        l = float(l_wear or 0)
        d = float(d_wear or 0)
    except Exception:
        return False

    if abs(l) < 0.000001 and abs(d) < 0.000001:
        return False

    if abs(l) > WIPS_MAX_REASONABLE_OFFSET_ABS:
        return False

    if abs(d) > WIPS_MAX_REASONABLE_OFFSET_ABS:
        return False

    return True


def machine_idle_long_enough_for_offset_detection(state: dict) -> bool:
    no_spindle_since = state.get("_no_spindle_since")

    if not no_spindle_since:
        return False

    try:
        return time.time() - float(no_spindle_since) >= WIPS_MIN_IDLE_BEFORE_DETECT_SEC
    except Exception:
        return False


# ============================================================
# G-CODE ANALYSIS HELPERS
# ============================================================
def normalize_program_key(prog_raw: str) -> str:
    p = str(prog_raw or "").strip().replace("\\", "/")
    p = os.path.basename(p)
    p = re.sub(r"\.(NC|EIA|MIN|TXT|MPF|TAP|NGC)$", "", p, flags=re.I)
    p = p.upper().strip()
    p = re.sub(r"[^A-Z0-9_\- ]", "", p)
    p = re.sub(r"\s+", " ", p)
    return p or "UNKNOWN"


def strip_gcode_comments(line: str) -> str:
    """
    Removes common G-code comments:
    - Parentheses comments: ( comment )
    - Semicolon comments: ; comment
    """
    s = str(line or "")

    # 1. Remove standard parentheses comments (e.g., (TOOL 1))
    s = re.sub(r"\([^)]*\)", "", s)

    # 2. Remove the specific MathML pattern
    # Fixed: Used single quotes for attributes and escaped the brackets \[ \]
    pattern = r"<math><mrow><msup><mo form='prefix' stretchy='false'>\[</mo><mo form='postfix' stretchy='false'>\)</mo></msup><mo form='postfix' stretchy='false'>\]</mo><mo>∗</></mrow></math>"
    s = re.sub(pattern, "", s)

    # 3. Remove semicolon comments and everything after
    s = s.split(";", 1)[0]

    return s.strip()


def parse_gcode_words(line: str) -> dict:
    out = {}
    for letter, value in re.findall(r"([A-Z])\s*([-+]?\d*\.?\d+)", line.upper()):
        try:
            out.setdefault(letter, []).append(float(value))
        except Exception:
            pass
    return out


def last_word(words: dict, letter: str, default=None):
    vals = words.get(letter, [])
    if not vals:
        return default
    return vals[-1]


def dist3(a, b):
    return (
        (float(a.get("X", 0)) - float(b.get("X", 0))) ** 2
        + (float(a.get("Y", 0)) - float(b.get("Y", 0))) ** 2
        + (float(a.get("Z", 0)) - float(b.get("Z", 0))) ** 2
    ) ** 0.5


def dist_xy(a, b):
    return (
        (float(a.get("X", 0)) - float(b.get("X", 0))) ** 2
        + (float(a.get("Y", 0)) - float(b.get("Y", 0))) ** 2
    ) ** 0.5


def estimate_arc_length_in(old_pos: dict, new_pos: dict, words: dict, clockwise: bool) -> float:
    x0 = float(old_pos.get("X", 0))
    y0 = float(old_pos.get("Y", 0))
    x1 = float(new_pos.get("X", x0))
    y1 = float(new_pos.get("Y", y0))

    i = last_word(words, "I", None)
    j = last_word(words, "J", None)

    chord = ((x1 - x0) ** 2 + (y1 - y0) ** 2) ** 0.5

    if i is None or j is None:
        r = last_word(words, "R", None)
        if r is None or abs(float(r)) < 0.000001:
            return chord

        r = abs(float(r))
        try:
            theta = 2.0 * math.asin(max(-1.0, min(1.0, chord / (2.0 * r))))
            return abs(r * theta)
        except Exception:
            return chord

    cx = x0 + float(i)
    cy = y0 + float(j)

    a0 = math.atan2(y0 - cy, x0 - cx)
    a1 = math.atan2(y1 - cy, x1 - cx)

    if clockwise:
        if a1 >= a0:
            a1 -= 2 * math.pi
    else:
        if a1 <= a0:
            a1 += 2 * math.pi

    theta = abs(a1 - a0)
    r = ((x0 - cx) ** 2 + (y0 - cy) ** 2) ** 0.5

    if r <= 0:
        return chord

    return abs(r * theta)


def tool_diameter_from_catalog_for_analysis(tid: str) -> float:
    tid = normalize_tool_num(tid)
    name = get_tool_name(str(tid))
    dia = infer_tool_diameter(name)
    return dia if dia > 0 else 0.5


def compute_gcode_severity(metrics: dict, tid: str) -> float:
    dia = tool_diameter_from_catalog_for_analysis(tid)
    dia = max(dia, 0.0625)

    est_cut_sec = float(metrics.get("estimated_cut_sec", 0) or 0)
    cut_distance = float(metrics.get("cut_distance_in", 0) or 0)
    max_depth = abs(float(metrics.get("max_depth_in", 0) or 0))
    avg_stepdown = abs(float(metrics.get("avg_stepdown_in", 0) or 0))
    max_stepdown = abs(float(metrics.get("max_stepdown_in", 0) or 0))
    drill_cycles = int(metrics.get("drill_cycles", 0) or 0)
    peck_cycles = int(metrics.get("peck_cycles", 0) or 0)
    coolant_on = int(metrics.get("coolant_on_seen", 0) or 0)

    severity = 1.0

    if max_depth > dia:
        severity += min(0.45, (max_depth / max(dia, 0.001) - 1.0) * 0.10)

    if avg_stepdown > 0:
        severity += min(0.35, avg_stepdown / max(dia, 0.001) * 0.30)

    if max_stepdown > dia * 0.75:
        severity += 0.20

    if drill_cycles > 0:
        severity += min(0.25, drill_cycles / 200.0)

    if peck_cycles > 0:
        severity += min(0.25, peck_cycles / 120.0)

    if cut_distance > 200:
        severity += min(0.35, cut_distance / 1500.0)

    if coolant_on:
        severity *= 0.92

    if dia <= 0.125:
        severity *= 1.25
    elif dia <= 0.1875:
        severity *= 1.15

    if est_cut_sec <= 0 and cut_distance <= 0 and drill_cycles <= 0:
        severity = 1.0

    return round(_clamp(severity, 0.65, 2.75), 3)


def analyze_gcode_text(text: str, program_key: str = "") -> dict:
    units_scale = 1.0
    abs_mode = True

    pos = {"X": 0.0, "Y": 0.0, "Z": 0.0}
    motion = 0
    feed = 0.0
    rpm = 0.0
    spindle_on = False

    pending_tool = ""
    current_tool = ""

    tools = {}

    def ensure_tool(tid):
        tid = normalize_tool_num(tid)
        if not tid or tid == "0":
            tid = "UNKNOWN"

        if tid not in tools:
            tools[tid] = {
                "tid": tid,
                "estimated_cut_sec": 0.0,
                "cut_distance_in": 0.0,
                "rapid_distance_in": 0.0,
                "plunge_distance_in": 0.0,
                "feed_sum": 0.0,
                "feed_n": 0,
                "max_feed_ipm": 0.0,
                "rpm_sum": 0.0,
                "rpm_n": 0,
                "max_rpm": 0.0,
                "max_depth_in": 0.0,
                "stepdown_sum": 0.0,
                "stepdown_n": 0,
                "max_stepdown_in": 0.0,
                "xy_cut_moves": 0,
                "z_plunge_moves": 0,
                "drill_cycles": 0,
                "peck_cycles": 0,
                "peck_q_sum": 0.0,
                "peck_q_n": 0,
                "coolant_on_seen": 0,
            }

        return tools[tid]

    def active_metrics():
        return ensure_tool(current_tool or pending_tool or "UNKNOWN")

    for raw_line in str(text or "").splitlines():
        line = strip_gcode_comments(raw_line).upper()
        if not line:
            continue

        words = parse_gcode_words(line)
        if not words:
            continue

        for g in words.get("G", []):
            gi = int(round(g))
            if gi == 20:
                units_scale = 1.0
            elif gi == 21:
                units_scale = 1.0 / 25.4
            elif gi == 90:
                abs_mode = True
            elif gi == 91:
                abs_mode = False

        t_val = last_word(words, "T", None)
        if t_val is not None:
            pending_tool = normalize_tool_num(str(int(round(t_val))))

        m_codes = [int(round(x)) for x in words.get("M", [])]

        if 6 in m_codes:
            if pending_tool:
                current_tool = pending_tool
                ensure_tool(current_tool)

        if 3 in m_codes or 4 in m_codes:
            spindle_on = True

        if 5 in m_codes:
            spindle_on = False

        if 8 in m_codes:
            active_metrics()["coolant_on_seen"] = 1

        s_val = last_word(words, "S", None)
        if s_val is not None:
            rpm = abs(float(s_val))

        f_val = last_word(words, "F", None)
        if f_val is not None:
            feed = abs(float(f_val)) * units_scale

        for g in words.get("G", []):
            gi = int(round(g))
            if gi in (0, 1, 2, 3):
                motion = gi
            elif gi in (81, 82, 83, 84, 85, 86, 87, 88, 89):
                motion = gi

        old_pos = dict(pos)
        new_pos = dict(pos)

        for axis in ("X", "Y", "Z"):
            val = last_word(words, axis, None)
            if val is not None:
                val = float(val) * units_scale
                if abs_mode:
                    new_pos[axis] = val
                else:
                    new_pos[axis] = float(new_pos.get(axis, 0)) + val

        metrics = active_metrics()

        if rpm > 0:
            metrics["rpm_sum"] += rpm
            metrics["rpm_n"] += 1
            metrics["max_rpm"] = max(metrics["max_rpm"], rpm)

        if feed > 0:
            metrics["feed_sum"] += feed
            metrics["feed_n"] += 1
            metrics["max_feed_ipm"] = max(metrics["max_feed_ipm"], feed)

        z = float(new_pos.get("Z", 0))
        if z < 0:
            metrics["max_depth_in"] = max(metrics["max_depth_in"], abs(z))

        if motion == 0:
            d = dist3(old_pos, new_pos)
            metrics["rapid_distance_in"] += d

        elif motion == 1:
            d = dist3(old_pos, new_pos)
            xy = dist_xy(old_pos, new_pos)
            dz = float(new_pos.get("Z", 0)) - float(old_pos.get("Z", 0))

            is_cut = spindle_on or float(new_pos.get("Z", 0)) < 0 or float(old_pos.get("Z", 0)) < 0

            if is_cut:
                metrics["cut_distance_in"] += d

                if feed > 0:
                    metrics["estimated_cut_sec"] += (d / feed) * 60.0

                if xy > 0.0001:
                    metrics["xy_cut_moves"] += 1

                if dz < -0.0001:
                    plunge = abs(dz)
                    metrics["plunge_distance_in"] += plunge
                    metrics["z_plunge_moves"] += 1
                    metrics["stepdown_sum"] += plunge
                    metrics["stepdown_n"] += 1
                    metrics["max_stepdown_in"] = max(metrics["max_stepdown_in"], plunge)

        elif motion in (2, 3):
            arc_len = estimate_arc_length_in(old_pos, new_pos, words, clockwise=(motion == 2))
            metrics["cut_distance_in"] += arc_len
            metrics["xy_cut_moves"] += 1

            if feed > 0:
                metrics["estimated_cut_sec"] += (arc_len / feed) * 60.0

        elif motion in (81, 82, 83, 84, 85, 86, 87, 88, 89):
            metrics["drill_cycles"] += 1

            z_target = last_word(words, "Z", None)
            r_plane = last_word(words, "R", 0.0)
            q_peck = last_word(words, "Q", None)

            if z_target is not None:
                z_target = float(z_target) * units_scale
                r_plane = float(r_plane or 0) * units_scale
                depth = abs(r_plane - z_target)

                metrics["plunge_distance_in"] += depth
                metrics["max_depth_in"] = max(metrics["max_depth_in"], abs(z_target))

                if feed > 0:
                    metrics["estimated_cut_sec"] += (depth / feed) * 60.0

            if motion == 83 or q_peck is not None:
                metrics["peck_cycles"] += 1

                if q_peck is not None:
                    q = abs(float(q_peck) * units_scale)
                    metrics["peck_q_sum"] += q
                    metrics["peck_q_n"] += 1

        pos = new_pos

    result = {
        "program_key": program_key,
        "tools": {},
    }

    for tid, m in tools.items():
        feed_n = max(int(m.get("feed_n") or 0), 1)
        rpm_n = max(int(m.get("rpm_n") or 0), 1)
        step_n = max(int(m.get("stepdown_n") or 0), 1)
        peck_n = max(int(m.get("peck_q_n") or 0), 1)

        finalized = dict(m)
        finalized["avg_feed_ipm"] = round(float(m.get("feed_sum") or 0) / feed_n, 3)
        finalized["avg_rpm"] = round(float(m.get("rpm_sum") or 0) / rpm_n, 1)
        finalized["avg_stepdown_in"] = round(float(m.get("stepdown_sum") or 0) / step_n, 4)
        finalized["avg_peck_q_in"] = round(float(m.get("peck_q_sum") or 0) / peck_n, 4)
        finalized["severity_multiplier"] = compute_gcode_severity(finalized, tid)

        for k in [
            "feed_sum",
            "feed_n",
            "rpm_sum",
            "rpm_n",
            "stepdown_sum",
            "stepdown_n",
            "peck_q_sum",
            "peck_q_n",
        ]:
            finalized.pop(k, None)

        result["tools"][tid] = finalized

    return result


def find_program_file(prog_raw: str) -> str:
    if not prog_raw:
        return ""

    raw = str(prog_raw or "").strip().replace("\\", "/")
    base = os.path.basename(raw)
    key = normalize_program_key(base)

    if os.path.exists(raw) and os.path.isfile(raw):
        return raw

    names_to_try = [base]
    root_no_ext = re.sub(r"\.(NC|EIA|MIN|TXT|MPF|TAP|NGC)$", "", base, flags=re.I)
    names_to_try.append(root_no_ext)

    for ext in PROGRAM_EXTS:
        names_to_try.append(root_no_ext + ext)
        names_to_try.append(root_no_ext.upper() + ext.upper())

    for folder in PROGRAM_SEARCH_DIRS:
        if not folder or not os.path.isdir(folder):
            continue

        for name in names_to_try:
            p = os.path.join(folder, name)
            if os.path.exists(p) and os.path.isfile(p):
                return p

    for folder in PROGRAM_SEARCH_DIRS:
        if not folder or not os.path.isdir(folder):
            continue

        try:
            for root, dirs, files in os.walk(folder):
                for fn in files:
                    fn_key = normalize_program_key(fn)
                    if fn_key == key:
                        return os.path.join(root, fn)

                if root.count(os.sep) - folder.count(os.sep) > 5:
                    dirs[:] = []
        except Exception:
            continue

    return ""


def gcode_analysis_is_fresh(conn, machine: str, program_key: str) -> bool:
    row = conn.execute("""
        SELECT analyzed_at
        FROM gcode_tool_estimates
        WHERE machine=? AND program_key=?
        ORDER BY analyzed_at DESC
        LIMIT 1
    """, (
        machine,
        program_key,
    )).fetchone()

    if not row:
        return False

    try:
        analyzed = datetime.strptime(str(row[0])[:19], "%Y-%m-%d %H:%M:%S")
        age = (datetime.now() - analyzed).total_seconds()
        return age <= GCODE_ANALYSIS_REFRESH_SEC
    except Exception:
        return False


def analyze_and_store_gcode_program(conn, machine: str, prog_raw: str) -> dict:
    program_key = normalize_program_key(prog_raw)

    if not program_key or program_key in ("N/A", "IDLE", "UNKNOWN"):
        return {}

    if gcode_analysis_is_fresh(conn, machine, program_key):
        return {
            "status": "fresh",
            "program_key": program_key,
        }

    path = find_program_file(prog_raw)

    if not path:
        return {
            "status": "missing",
            "program_key": program_key,
            "message": "Program file not found in ELGIN_PROGRAM_DIRS/NEXUS_PROGRAM_DIRS.",
        }

    try:
        if os.path.getsize(path) > GCODE_ANALYSIS_MAX_FILE_BYTES:
            return {
                "status": "too_large",
                "program_key": program_key,
                "program_path": path,
            }

        with open(path, "r", encoding="utf-8", errors="replace") as f:
            text = f.read()

        analysis = analyze_gcode_text(text, program_key=program_key)

        for tid, metrics in analysis.get("tools", {}).items():
            conn.execute("""
                INSERT INTO gcode_tool_estimates (
                    machine, program_key, program_path, tid,

                    estimated_cut_sec,
                    cut_distance_in,
                    rapid_distance_in,
                    plunge_distance_in,

                    avg_feed_ipm,
                    max_feed_ipm,
                    avg_rpm,
                    max_rpm,

                    max_depth_in,
                    avg_stepdown_in,
                    max_stepdown_in,
                    xy_cut_moves,
                    z_plunge_moves,

                    drill_cycles,
                    peck_cycles,
                    avg_peck_q_in,

                    coolant_on_seen,
                    severity_multiplier,

                    raw_json,
                    analyzed_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
                ON CONFLICT(machine, program_key, tid) DO UPDATE SET
                    program_path=excluded.program_path,

                    estimated_cut_sec=excluded.estimated_cut_sec,
                    cut_distance_in=excluded.cut_distance_in,
                    rapid_distance_in=excluded.rapid_distance_in,
                    plunge_distance_in=excluded.plunge_distance_in,

                    avg_feed_ipm=excluded.avg_feed_ipm,
                    max_feed_ipm=excluded.max_feed_ipm,
                    avg_rpm=excluded.avg_rpm,
                    max_rpm=excluded.max_rpm,

                    max_depth_in=excluded.max_depth_in,
                    avg_stepdown_in=excluded.avg_stepdown_in,
                    max_stepdown_in=excluded.max_stepdown_in,
                    xy_cut_moves=excluded.xy_cut_moves,
                    z_plunge_moves=excluded.z_plunge_moves,

                    drill_cycles=excluded.drill_cycles,
                    peck_cycles=excluded.peck_cycles,
                    avg_peck_q_in=excluded.avg_peck_q_in,

                    coolant_on_seen=excluded.coolant_on_seen,
                    severity_multiplier=excluded.severity_multiplier,

                    raw_json=excluded.raw_json,
                    analyzed_at=datetime('now')
            """, (
                machine,
                program_key,
                path,
                str(tid),

                float(metrics.get("estimated_cut_sec") or 0),
                float(metrics.get("cut_distance_in") or 0),
                float(metrics.get("rapid_distance_in") or 0),
                float(metrics.get("plunge_distance_in") or 0),

                float(metrics.get("avg_feed_ipm") or 0),
                float(metrics.get("max_feed_ipm") or 0),
                float(metrics.get("avg_rpm") or 0),
                float(metrics.get("max_rpm") or 0),

                float(metrics.get("max_depth_in") or 0),
                float(metrics.get("avg_stepdown_in") or 0),
                float(metrics.get("max_stepdown_in") or 0),
                int(metrics.get("xy_cut_moves") or 0),
                int(metrics.get("z_plunge_moves") or 0),

                int(metrics.get("drill_cycles") or 0),
                int(metrics.get("peck_cycles") or 0),
                float(metrics.get("avg_peck_q_in") or 0),

                int(metrics.get("coolant_on_seen") or 0),
                float(metrics.get("severity_multiplier") or 1.0),

                json.dumps(metrics),
            ))

        return {
            "status": "ok",
            "program_key": program_key,
            "program_path": path,
            "tools": list(analysis.get("tools", {}).keys()),
        }

    except Exception as e:
        return {
            "status": "error",
            "program_key": program_key,
            "program_path": path,
            "error": str(e),
        }


def get_gcode_tool_severity(conn, machine: str, prog_raw: str, tid: str) -> float:
    program_key = normalize_program_key(prog_raw)
    tid = normalize_tool_num(tid)

    if not program_key or not tid or tid == "0":
        return 1.0

    analyze_and_store_gcode_program(conn, machine, prog_raw)

    row = conn.execute("""
        SELECT severity_multiplier
        FROM gcode_tool_estimates
        WHERE machine=? AND program_key=? AND tid=?
        ORDER BY analyzed_at DESC
        LIMIT 1
    """, (
        machine,
        program_key,
        tid,
    )).fetchone()

    if not row:
        return 1.0

    try:
        return _clamp(float(row[0] or 1.0), 0.65, 2.75)
    except Exception:
        return 1.0


def get_gcode_program_runtime_estimate_sec(conn, machine: str, prog_raw: str) -> dict:
    program_key = normalize_program_key(prog_raw)

    if not program_key or program_key in ("N/A", "IDLE", "UNKNOWN"):
        return {
            "ok": False,
            "program_key": program_key,
            "estimated_sec": 0.0,
            "source": "NO_PROGRAM",
            "message": "No valid program.",
        }

    try:
        analyze_and_store_gcode_program(conn, machine, prog_raw)
    except Exception as e:
        return {
            "ok": False,
            "program_key": program_key,
            "estimated_sec": 0.0,
            "source": "GCODE_ANALYSIS_ERROR",
            "message": str(e),
        }

    rows = _rowdicts(conn.execute("""
        SELECT *
        FROM gcode_tool_estimates
        WHERE machine=? AND program_key=?
    """, (
        machine,
        program_key,
    )))

    if not rows:
        return {
            "ok": False,
            "program_key": program_key,
            "estimated_sec": 0.0,
            "source": "NO_GCODE_ESTIMATE",
            "message": "No G-code estimate found for program.",
        }

    cut_sec = sum(float(r.get("estimated_cut_sec") or 0) for r in rows)
    rapid_distance = sum(float(r.get("rapid_distance_in") or 0) for r in rows)

    assumed_rapid_ipm = 600.0
    rapid_sec = (rapid_distance / assumed_rapid_ipm) * 60.0 if rapid_distance > 0 else 0.0

    tool_count = len(set(str(r.get("tid") or "") for r in rows if str(r.get("tid") or "").strip()))
    toolchange_sec = max(0, tool_count - 1) * 20.0

    misc_overhead_sec = max(20.0, cut_sec * 0.04) if cut_sec > 0 else 0.0

    estimated_sec = cut_sec + rapid_sec + toolchange_sec + misc_overhead_sec

    return {
        "ok": estimated_sec > 0,
        "program_key": program_key,
        "estimated_sec": round(estimated_sec, 2),
        "cut_sec": round(cut_sec, 2),
        "rapid_sec": round(rapid_sec, 2),
        "toolchange_sec": round(toolchange_sec, 2),
        "misc_overhead_sec": round(misc_overhead_sec, 2),
        "tool_count": tool_count,
        "source": "GCODE_ESTIMATE",
        "message": f"Estimated from {tool_count} tool section(s).",
    }


def compute_live_program_progress(
    state: dict,
    prog_raw: str,
    estimated_sec: float,
    cycle_sec: float,
    actual_cutting: bool,
) -> dict:
    now = time.time()
    prog_key = normalize_program_key(prog_raw)

    if state.get("_last_progress_prog") != prog_key:
        state["_last_progress_prog"] = prog_key
        state["_program_started_at"] = now

    estimated_sec = float(estimated_sec or 0)
    cycle_sec = float(cycle_sec or 0)

    if estimated_sec <= 0:
        return {
            "program_est_sec": 0.0,
            "program_progress_pct": 0.0,
            "program_remaining_sec": 0.0,
            "program_est_finish": "",
            "program_progress_source": "NO_ESTIMATE",
        }

    elapsed = cycle_sec
    source = "ESTIMATE_PLUS_QCODE_CYCLE_TIMER"

    if elapsed <= 0 and state.get("_program_started_at"):
        elapsed = max(0.0, now - float(state.get("_program_started_at") or now))
        source = "ESTIMATE_PLUS_WALL_CLOCK"

    progress_pct = _clamp((elapsed / estimated_sec) * 100.0, 0, 100)
    remaining_sec = max(0.0, estimated_sec - elapsed)

    finish = ""
    if remaining_sec > 0:
        finish = datetime.fromtimestamp(now + remaining_sec).strftime("%H:%M:%S")

    return {
        "program_est_sec": round(estimated_sec, 2),
        "program_progress_pct": round(progress_pct, 1),
        "program_remaining_sec": round(remaining_sec, 2),
        "program_est_finish": finish,
        "program_progress_source": source,
    }


# ============================================================
# STEEL SHEET PARSING HELPERS
# ============================================================
STEEL_MATERIAL_PATTERNS = [
    r"#\s*\d+\s*4140",
    r"#\s*\d+\s*H13",
    r"#\s*\d+\s*A2",
    r"#\s*\d+\s*D2",
    r"\b4140\b",
    r"\bH13\b",
    r"\bA2\b",
    r"\bD2\b",
    r"\bO1\b",
    r"\bS7\b",
    r"\bALUMINUM\b",
    r"\bALUM\b",
    r"\b6061\b",
    r"\b7075\b",
]


def google_sheet_to_csv_url(url: str) -> str:
    u = str(url or "").strip()

    if "docs.google.com/spreadsheets" not in u:
        return u

    m = re.search(r"/spreadsheets/d/([^/]+)", u)
    if not m:
        return u

    sheet_id = m.group(1)

    gid = "0"
    gm = re.search(r"[?#&]gid=(\d+)", u)
    if gm:
        gid = gm.group(1)

    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"


def fetch_url_bytes(url: str) -> tuple:
    if not url:
        raise ValueError("No URL provided")

    fixed_url = google_sheet_to_csv_url(url)

    req = urllib.request.Request(
        fixed_url,
        headers={
            "User-Agent": f"{APP_SLUG}/1.0",
            "Accept": "*/*",
        },
    )

    with urllib.request.urlopen(req, timeout=12) as resp:
        data = resp.read(10_000_000)
        final_url = resp.geturl()
        content_type = resp.headers.get("Content-Type", "")

    return data, final_url, content_type


def normalize_steel_ocr_text(text: str) -> str:
    s = str(text or "")
    s = s.replace("\x00", " ")
    s = s.replace("×", " X ")
    s = s.replace("x", " X ")
    s = s.replace("|", " ")
    s = s.replace("\t", " ")
    s = re.sub(r"[ ]+", " ", s)
    s = re.sub(r"\n\s+", "\n", s)
    return s


def looks_like_material_token(v: str) -> bool:
    u = str(v or "").upper()
    return any(re.search(p, u) for p in STEEL_MATERIAL_PATTERNS)


def clean_material(v: str) -> str:
    s = str(v or "").upper().strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("# ", "#")
    return s


def extract_material_from_line(line: str) -> str:
    u = str(line or "").upper()

    for pat in STEEL_MATERIAL_PATTERNS:
        m = re.search(pat, u)
        if m:
            return clean_material(m.group(0))

    return ""


def parse_float_token(v: str) -> float:
    try:
        return float(str(v).strip())
    except Exception:
        return 0.0


def parse_steel_row_from_line(line: str) -> dict:
    raw = str(line or "").strip()
    if not raw:
        return {}

    line_clean = normalize_steel_ocr_text(raw)
    line_clean = re.sub(r"\s+", " ", line_clean).strip()

    material = extract_material_from_line(line_clean)
    nums = re.findall(r"[-+]?\d+(?:\.\d+)?", line_clean)

    if len(nums) < 4:
        return {}

    pattern = re.compile(
        r"^\s*(?P<qty>\d+(?:\.\d+)?)\s+"
        r"(?P<name>.+?)\s+"
        r"(?P<d1>\d+(?:\.\d+)?)\s*X\s*"
        r"(?P<d2>\d+(?:\.\d+)?)\s*X\s*"
        r"(?P<d3>\d+(?:\.\d+)?)"
        r"(?P<trail>.*)$",
        re.I,
    )

    m = pattern.search(line_clean)
    if m:
        trail = m.group("trail") or ""
        mat = material or extract_material_from_line(trail)
        confidence = 0.95 if mat or looks_like_material_token(trail) else 0.70

        return {
            "qty": parse_float_token(m.group("qty")),
            "item_name": re.sub(r"\s+", " ", m.group("name").strip()),
            "dim_1": parse_float_token(m.group("d1")),
            "dim_2": parse_float_token(m.group("d2")),
            "dim_3": parse_float_token(m.group("d3")),
            "material": mat,
            "raw_line": raw,
            "confidence": confidence,
        }

    qty = parse_float_token(nums[0])
    d1 = parse_float_token(nums[-3])
    d2 = parse_float_token(nums[-2])
    d3 = parse_float_token(nums[-1])

    name_part = line_clean
    name_part = re.sub(r"^\s*\d+(?:\.\d+)?\s+", "", name_part)
    name_part = re.sub(
        r"\d+(?:\.\d+)?\s*X\s*\d+(?:\.\d+)?\s*X\s*\d+(?:\.\d+)?.*$",
        "",
        name_part,
        flags=re.I,
    )
    name_part = re.sub(r"\s+", " ", name_part).strip()

    if not name_part:
        return {}

    return {
        "qty": qty,
        "item_name": name_part,
        "dim_1": d1,
        "dim_2": d2,
        "dim_3": d3,
        "material": material,
        "raw_line": raw,
        "confidence": 0.55 if material else 0.40,
    }


def parse_steel_items_from_text(text: str) -> list:
    text = normalize_steel_ocr_text(text)

    rows = []

    for line in text.splitlines():
        item = parse_steel_row_from_line(line)
        if item:
            rows.append(item)

    dedup = []
    seen = set()

    for r in rows:
        key = (
            round(float(r.get("qty") or 0), 4),
            str(r.get("item_name") or "").upper(),
            round(float(r.get("dim_1") or 0), 4),
            round(float(r.get("dim_2") or 0), 4),
            round(float(r.get("dim_3") or 0), 4),
            str(r.get("material") or "").upper(),
        )

        if key in seen:
            continue

        seen.add(key)
        dedup.append(r)

    return dedup


def extract_text_from_csv_bytes(data: bytes) -> str:
    s = data.decode("utf-8", errors="replace")
    out_lines = []

    reader = csv.reader(io.StringIO(s))
    for row in reader:
        vals = [str(v).strip() for v in row if str(v).strip()]
        if vals:
            out_lines.append(" ".join(vals))

    return "\n".join(out_lines)


def extract_text_from_xlsx_bytes(data: bytes) -> str:
    if openpyxl is None:
        raise RuntimeError("openpyxl is not installed. Run: pip install openpyxl")

    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    out_lines = []

    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            vals = []
            for v in row:
                if v is None:
                    continue
                vals.append(str(v).strip())
            if vals:
                out_lines.append(" ".join(vals))

    return "\n".join(out_lines)


def extract_text_from_pdf_bytes(data: bytes) -> str:
    if PyPDF2 is None:
        raise RuntimeError("PyPDF2 is not installed. Run: pip install PyPDF2")

    reader = PyPDF2.PdfReader(io.BytesIO(data))
    pages = []

    for page in reader.pages:
        try:
            pages.append(page.extract_text() or "")
        except Exception:
            pass

    return "\n".join(pages)


def extract_text_from_image_bytes(data: bytes) -> str:
    if Image is None or pytesseract is None:
        raise RuntimeError("Image OCR requires pillow and pytesseract. Run: pip install pillow pytesseract")

    img = Image.open(io.BytesIO(data))
    return pytesseract.image_to_string(img)


def guess_source_type(url: str, content_type: str) -> str:
    u = str(url or "").lower()
    ct = str(content_type or "").lower()

    if "csv" in ct or u.endswith(".csv") or "format=csv" in u:
        return "csv"
    if "spreadsheet" in ct or u.endswith(".xlsx") or u.endswith(".xlsm"):
        return "xlsx"
    if "pdf" in ct or u.endswith(".pdf"):
        return "pdf"
    if "image" in ct or any(u.endswith(ext) for ext in [".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"]):
        return "image"
    if "html" in ct:
        return "html"

    return "text"


def extract_text_from_steel_source(data: bytes, url: str, content_type: str) -> tuple:
    source_type = guess_source_type(url, content_type)

    if source_type == "csv":
        return extract_text_from_csv_bytes(data), "csv"

    if source_type == "xlsx":
        return extract_text_from_xlsx_bytes(data), "xlsx"

    if source_type == "pdf":
        return extract_text_from_pdf_bytes(data), "pdf"

    if source_type == "image":
        return extract_text_from_image_bytes(data), "image_ocr"

    text = data.decode("utf-8", errors="replace")

    if source_type == "html":
        text = re.sub(r"<br\s*/?>", "\n", text, flags=re.I)
        text = re.sub(r"</tr>", "\n", text, flags=re.I)
        text = re.sub(r"</td>", " ", text, flags=re.I)
        text = re.sub(r"<[^>]+>", " ", text)
        text = re.sub(r"&nbsp;", " ", text)
        text = re.sub(r"&amp;", "&", text)

    return text, source_type


def parse_and_store_steel_sheet(conn, job_num: str, source_url: str, force: bool = True) -> dict:
    """
    Legacy URL-based parser retained for compatibility.
    The UI no longer uses URL upload as the main workflow.
    New workflow is PDF upload from Job Matching.
    """
    job_num = validate_new_job_num_or_400(job_num)
    source_url = str(source_url or "").strip()

    if not source_url:
        raise HTTPException(400, "Steel sheet URL required")

    try:
        data, final_url, content_type = fetch_url_bytes(source_url)
        text, source_type = extract_text_from_steel_source(data, final_url, content_type)
        items = parse_steel_items_from_text(text)

        if force:
            conn.execute("DELETE FROM steel_sheet_items WHERE job_num=?", (job_num,))

        for item in items:
            mat = clean_material(item.get("material") or "")

            heat_or_type = ""
            m = re.search(r"#\s*(\d+)", mat)
            if m:
                heat_or_type = "#" + m.group(1)

            conn.execute("""
                INSERT INTO steel_sheet_items (
                    job_num, source_url, source_type,
                    qty, item_name,
                    dim_1, dim_2, dim_3,
                    material, heat_or_type,
                    raw_line, confidence,
                    parsed_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
            """, (
                job_num,
                final_url,
                source_type,
                float(item.get("qty") or 0),
                item.get("item_name") or "",
                float(item.get("dim_1") or 0),
                float(item.get("dim_2") or 0),
                float(item.get("dim_3") or 0),
                mat,
                heat_or_type,
                item.get("raw_line") or "",
                float(item.get("confidence") or 0),
            ))

        preview = normalize_steel_ocr_text(text)[:2000]
        status = "ok" if items else "no_items_found"
        message = f"Parsed {len(items)} steel item(s)." if items else "No steel rows found."

        conn.execute("""
            INSERT INTO steel_sheet_parse_log (
                job_num, source_url, status, message, parsed_count, raw_text_preview
            )
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            job_num,
            final_url,
            status,
            message,
            len(items),
            preview,
        ))

        return {
            "status": status,
            "job_num": job_num,
            "source_url": final_url,
            "source_type": source_type,
            "count": len(items),
            "items": items,
            "message": message,
            "preview": preview,
        }

    except Exception as e:
        conn.execute("""
            INSERT INTO steel_sheet_parse_log (
                job_num, source_url, status, message, parsed_count
            )
            VALUES (?, ?, 'error', ?, 0)
        """, (
            job_num,
            source_url,
            str(e),
        ))

        return {
            "status": "error",
            "job_num": job_num,
            "source_url": source_url,
            "count": 0,
            "items": [],
            "message": str(e),
        }


def parse_and_store_steel_pdf_bytes(
    conn,
    job_num: str,
    pdf_bytes: bytes,
    original_filename: str = "",
    stored_path: str = "",
    force: bool = True,
) -> dict:
    job_num = validate_new_job_num_or_400(job_num)

    if not pdf_bytes:
        raise HTTPException(400, "Empty PDF upload")

    try:
        text = extract_text_from_pdf_bytes(pdf_bytes)
        items = parse_steel_items_from_text(text)

        if force:
            conn.execute("DELETE FROM steel_sheet_items WHERE job_num=?", (job_num,))

        for item in items:
            mat = clean_material(item.get("material") or "")

            heat_or_type = ""
            m = re.search(r"#\s*(\d+)", mat)
            if m:
                heat_or_type = "#" + m.group(1)

            conn.execute("""
                INSERT INTO steel_sheet_items (
                    job_num, source_url, source_type,
                    qty, item_name,
                    dim_1, dim_2, dim_3,
                    material, heat_or_type,
                    raw_line, confidence,
                    parsed_at
                )
                VALUES (?, ?, 'uploaded_pdf', ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
            """, (
                job_num,
                stored_path or original_filename or "uploaded_pdf",
                float(item.get("qty") or 0),
                item.get("item_name") or "",
                float(item.get("dim_1") or 0),
                float(item.get("dim_2") or 0),
                float(item.get("dim_3") or 0),
                mat,
                heat_or_type,
                item.get("raw_line") or "",
                float(item.get("confidence") or 0),
            ))

        preview = normalize_steel_ocr_text(text)[:2000]
        status = "ok" if items else "no_items_found"
        message = f"Parsed {len(items)} steel item(s)." if items else "No steel rows found in uploaded PDF."

        conn.execute("""
            INSERT INTO steel_sheet_parse_log (
                job_num, source_url, status, message, parsed_count, raw_text_preview
            )
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            job_num,
            stored_path or original_filename or "uploaded_pdf",
            status,
            message,
            len(items),
            preview,
        ))

        return {
            "status": status,
            "job_num": job_num,
            "source_type": "uploaded_pdf",
            "stored_path": stored_path,
            "original_filename": original_filename,
            "count": len(items),
            "items": items,
            "message": message,
            "preview": preview,
        }

    except Exception as e:
        conn.execute("""
            INSERT INTO steel_sheet_parse_log (
                job_num, source_url, status, message, parsed_count
            )
            VALUES (?, ?, 'error', ?, 0)
        """, (
            job_num,
            stored_path or original_filename or "uploaded_pdf",
            str(e),
        ))

        return {
            "status": "error",
            "job_num": job_num,
            "source_type": "uploaded_pdf",
            "stored_path": stored_path,
            "original_filename": original_filename,
            "count": 0,
            "items": [],
            "message": str(e),
        }


def get_steel_sheet_items_for_job(conn, job_num: str) -> list:
    return _rowdicts(conn.execute("""
        SELECT *
        FROM steel_sheet_items
        WHERE job_num=?
        ORDER BY id
    """, (normalize_job_num(job_num),)))


def steel_sheet_summary_for_job(conn, job_num: str) -> dict:
    rows = get_steel_sheet_items_for_job(conn, job_num)

    if not rows:
        return {
            "count": 0,
            "materials": [],
            "has_4140": False,
            "largest_piece": None,
            "message": "No parsed steel-sheet items.",
        }

    materials = sorted(set((r.get("material") or "").strip() for r in rows if r.get("material")))

    largest = None
    largest_vol = -1

    for r in rows:
        vol = (
            float(r.get("dim_1") or 0)
            * float(r.get("dim_2") or 0)
            * float(r.get("dim_3") or 0)
            * float(r.get("qty") or 1)
        )
        if vol > largest_vol:
            largest_vol = vol
            largest = r

    return {
        "count": len(rows),
        "materials": materials,
        "has_4140": any("4140" in m for m in materials),
        "largest_piece": largest,
        "message": f"{len(rows)} steel item(s), material(s): {', '.join(materials) if materials else 'unknown'}",
    }


# ============================================================
# END PART 3
# PART 4 CONTINUES WITH:
# - job series / financial / cleanup helpers
# - smart historical estimate helpers
# - old/reused program matching helpers
# - UMC scheduler helpers
# - QR/barcode helpers
# - runtime update helpers
# - efficiency/tool-cycle/inspection helpers
# - machine polling loop
# ============================================================
# ============================================================
# JOB SERIES / FINANCIAL / CLEANUP HELPERS
# ============================================================
def ensure_job_series(
    conn,
    job_num: str,
    due_date: str = "",
    steel_sheet_url: str = "",
    notes: str = "",
    matched_job_num: str = "",
    matched_reason: str = "",
):
    """
    Create a job series and default components if it is a valid shop job.

    Important:
    - Does NOT auto-create all-zero fake jobs like J0000.
    - Does NOT assign fixed 45-minute default component estimates.
    - Component estimates start at 0 and are later updated from:
        1. Q-code cycle timer
        2. G-code estimate
        3. similar completed job actual runtime
        4. otherwise unknown
    """
    job_num = normalize_job_num(job_num)
    if not job_num:
        return

    if job_num_is_all_zeros(job_num):
        return

    due_date = due_date or ""
    steel_sheet_url = steel_sheet_url or ""
    notes = notes or ""
    matched_job_num = normalize_job_num(matched_job_num) if matched_job_num else ""
    matched_reason = matched_reason or ""

    if matched_job_num and job_num_is_all_zeros(matched_job_num):
        matched_job_num = ""
        matched_reason = ""

    conn.execute("""
        INSERT OR IGNORE INTO job_series (
            job_num, due_date, steel_sheet_url, notes, status,
            matched_job_num, matched_reason, matched_at,
            created_at, updated_at
        )
        VALUES (
            ?, ?, ?, ?, 'ACTIVE',
            ?, ?, CASE WHEN ?!='' THEN datetime('now') ELSE '' END,
            datetime('now'), datetime('now')
        )
    """, (
        job_num,
        due_date,
        steel_sheet_url,
        notes,
        matched_job_num,
        matched_reason,
        matched_job_num,
    ))

    conn.execute("""
        UPDATE job_series
        SET
            due_date=CASE WHEN ?!='' THEN ? ELSE due_date END,
            steel_sheet_url=CASE WHEN ?!='' THEN ? ELSE steel_sheet_url END,
            notes=CASE WHEN ?!='' THEN ? ELSE notes END,
            matched_job_num=CASE WHEN ?!='' THEN ? ELSE matched_job_num END,
            matched_reason=CASE WHEN ?!='' THEN ? ELSE matched_reason END,
            matched_at=CASE WHEN ?!='' THEN datetime('now') ELSE matched_at END,
            updated_at=datetime('now')
        WHERE job_num=?
    """, (
        due_date, due_date,
        steel_sheet_url, steel_sheet_url,
        notes, notes,
        matched_job_num, matched_job_num,
        matched_reason, matched_reason,
        matched_job_num,
        job_num,
    ))

    for idx, comp in enumerate(DEFAULT_JOB_COMPONENTS):
        mach = DEFAULT_COMPONENT_MACHINE.get(comp, "")
        est_sec = float(DEFAULT_COMPONENT_EST_SEC.get(comp, 0) or 0)

        conn.execute("""
            INSERT OR IGNORE INTO job_components (
                job_num, component_name, machine, status, sort_order, estimated_sec,
                estimate_source, estimate_basis, estimate_updated_at
            )
            VALUES (?, ?, ?, 'PENDING', ?, ?, '', '', '')
        """, (
            job_num,
            comp,
            mach,
            idx,
            est_sec,
        ))


def choose_match_with_due(matches):
    dated = []
    today = date.today()

    for m in matches:
        d = _date_obj(m.get("due_date"))
        if d:
            dated.append((
                abs((d - today).days),
                d,
                normalize_job_num(m.get("job_id", "")),
                m,
            ))

    if not dated:
        return None

    dated.sort(key=lambda x: (x[0], x[1], x[2]))
    return dated[0][3]


def try_inherit_due_from_catalog(conn, job_num: str, tolerance: float = 0.01):
    job_num = normalize_job_num(job_num)
    if not job_num or job_num_is_all_zeros(job_num):
        return

    js = conn.execute("SELECT * FROM job_series WHERE job_num=?", (job_num,)).fetchone()
    if js and dict(js).get("due_date"):
        return

    src = conn.execute("SELECT * FROM part_catalog WHERE job_id=?", (job_num,)).fetchone()
    if not src:
        return

    src = dict(src)
    w = float(src.get("weight") or 0)
    x = float(src.get("com_x") or 0)
    y = float(src.get("com_y") or 0)
    z = float(src.get("com_z") or 0)

    rows = [
        dict(r)
        for r in conn.execute(
            "SELECT * FROM part_catalog WHERE job_id!=?",
            (job_num,),
        ).fetchall()
    ]

    matches = [
        r for r in rows
        if not job_num_is_all_zeros(r.get("job_id", ""))
        and abs(float(r.get("weight") or 0) - w) <= tolerance
        and abs(float(r.get("com_x") or 0) - x) <= tolerance
        and abs(float(r.get("com_y") or 0) - y) <= tolerance
        and abs(float(r.get("com_z") or 0) - z) <= tolerance
    ]

    best = choose_match_with_due(matches)
    if not best:
        return

    inherited_from = normalize_job_num(best.get("job_id", ""))
    due_date = best.get("due_date") or ""
    steel = best.get("steel_sheet_url") or ""

    if not due_date:
        return

    ensure_job_series(
        conn,
        job_num,
        due_date,
        steel,
        "",
        inherited_from,
        "INHERITED_DUE_DATE_FROM_PHYSICAL_MATCH",
    )

    conn.execute("""
        UPDATE part_catalog
        SET
            due_date=CASE WHEN due_date IS NULL OR due_date='' THEN ? ELSE due_date END,
            steel_sheet_url=CASE WHEN steel_sheet_url IS NULL OR steel_sheet_url='' THEN ? ELSE steel_sheet_url END
        WHERE job_id=?
    """, (due_date, steel, job_num))


# ============================================================
# SMART JOB ESTIMATES + OLD PROGRAM / MATCHED JOB RESOLUTION
# ============================================================
def get_part_signature(conn, job_num: str) -> dict:
    j = normalize_job_num(job_num)
    if not j:
        return {}

    row = conn.execute("""
        SELECT *
        FROM part_catalog
        WHERE job_id=?
    """, (j,)).fetchone()

    if not row:
        return {}

    r = dict(row)
    return {
        "job_num": j,
        "weight": float(r.get("weight") or 0),
        "com_x": float(r.get("com_x") or 0),
        "com_y": float(r.get("com_y") or 0),
        "com_z": float(r.get("com_z") or 0),
        "due_date": r.get("due_date") or "",
    }


def part_signature_is_meaningful(sig: dict) -> bool:
    if not sig:
        return False

    vals = [
        float(sig.get("weight") or 0),
        float(sig.get("com_x") or 0),
        float(sig.get("com_y") or 0),
        float(sig.get("com_z") or 0),
    ]

    return any(abs(v) > 0.000001 for v in vals)


def part_signature_delta(a: dict, b: dict) -> float:
    return (
        abs(float(a.get("weight") or 0) - float(b.get("weight") or 0))
        + abs(float(a.get("com_x") or 0) - float(b.get("com_x") or 0))
        + abs(float(a.get("com_y") or 0) - float(b.get("com_y") or 0))
        + abs(float(a.get("com_z") or 0) - float(b.get("com_z") or 0))
    )


def parts_are_similar(a: dict, b: dict, tolerance: float = SIMILAR_JOB_TOLERANCE) -> bool:
    if not part_signature_is_meaningful(a) or not part_signature_is_meaningful(b):
        return False

    return (
        abs(float(a.get("weight") or 0) - float(b.get("weight") or 0)) <= tolerance
        and abs(float(a.get("com_x") or 0) - float(b.get("com_x") or 0)) <= tolerance
        and abs(float(a.get("com_y") or 0) - float(b.get("com_y") or 0)) <= tolerance
        and abs(float(a.get("com_z") or 0) - float(b.get("com_z") or 0)) <= tolerance
    )



def _signature_key(v: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", str(v or "").upper())


def normalize_signature_component(v: str) -> str:
    key = _signature_key(v)
    if key in JOB_SIGNATURE_COMPONENT_ALIASES:
        return JOB_SIGNATURE_COMPONENT_ALIASES[key]
    return ""


def _signature_float(row: dict, *names: str) -> float:
    for name in names:
        if name in row and row.get(name) not in (None, ""):
            try:
                return float(str(row.get(name)).strip())
            except Exception:
                pass
    return 0.0


def _signature_bool(row: dict, *names: str) -> bool:
    for name in names:
        v = row.get(name)
        if v is None:
            continue
        s = str(v).strip().upper()
        if s in {"TRUE", "YES", "Y", "1", "OK"}:
            return True
        if s in {"FALSE", "NO", "N", "0"}:
            return False
    return False


def parse_job_signature_csv_text(text: str, fallback_job_num: str = "") -> list[dict]:
    rows: list[dict] = []
    reader = csv.DictReader(io.StringIO(text or ""))

    for raw in reader:
        # Accept the VBA macro headers and a few hand-written variants.
        role = normalize_signature_component(
            raw.get("ComponentRole")
            or raw.get("component_role")
            or raw.get("Role")
            or raw.get("Component")
            or raw.get("QuoteName")
            or raw.get("quote_name")
        )
        if not role:
            continue

        job_num = normalize_job_num(
            raw.get("JobNumber")
            or raw.get("job_num")
            or raw.get("Job")
            or fallback_job_num
        )
        if not job_num:
            job_num = normalize_job_num(fallback_job_num)

        rows.append({
            "job_num": job_num,
            "component_role": role,
            "quote_name": str(raw.get("QuoteName") or raw.get("quote_name") or "").strip(),
            "cad_component": str(raw.get("CadComponent") or raw.get("cad_component") or "").strip(),
            "clean_name": str(raw.get("CleanName") or raw.get("clean_name") or "").strip(),
            "length": _signature_float(raw, "Length", "length", "L"),
            "width": _signature_float(raw, "Width", "width", "W"),
            "thickness": _signature_float(raw, "Thickness", "thickness", "T"),
            "mass": _signature_float(raw, "Mass", "mass", "Weight", "weight"),
            "center_x": _signature_float(raw, "CenterX", "center_x", "CogX", "cog_x", "ComX", "com_x"),
            "center_y": _signature_float(raw, "CenterY", "center_y", "CogY", "cog_y", "ComY", "com_y"),
            "center_z": _signature_float(raw, "CenterZ", "center_z", "CogZ", "cog_z", "ComZ", "com_z"),
            "has_center": 1 if _signature_bool(raw, "HasCenter", "has_center") else 0,
        })

    return rows


def upsert_job_signature_rows(conn, job_num: str, rows: list[dict], source_file: str = "") -> int:
    j = normalize_job_num(job_num)
    if not j:
        raise ValueError("job number required")

    cleaned = []
    seen = set()
    for row in rows:
        role = normalize_signature_component(row.get("component_role") or "")
        if not role or role in seen:
            continue
        seen.add(role)
        r = dict(row)
        r["job_num"] = j
        r["component_role"] = role
        cleaned.append(r)

    conn.execute("DELETE FROM job_signature_components WHERE job_num=?", (j,))

    for r in cleaned:
        conn.execute("""
            INSERT INTO job_signature_components (
                job_num, component_role, quote_name, cad_component, clean_name,
                length, width, thickness, mass,
                center_x, center_y, center_z, has_center,
                source_file, imported_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
        """, (
            j,
            r.get("component_role") or "",
            r.get("quote_name") or "",
            r.get("cad_component") or "",
            r.get("clean_name") or "",
            float(r.get("length") or 0),
            float(r.get("width") or 0),
            float(r.get("thickness") or 0),
            float(r.get("mass") or 0),
            float(r.get("center_x") or 0),
            float(r.get("center_y") or 0),
            float(r.get("center_z") or 0),
            1 if r.get("has_center") else 0,
            source_file or "",
        ))

    return len(cleaned)


def load_job_signature(conn, job_num: str) -> dict:
    j = normalize_job_num(job_num)
    if not j:
        return {}
    rows = _rowdicts(conn.execute("""
        SELECT *
        FROM job_signature_components
        WHERE job_num=?
    """, (j,)))
    return {normalize_signature_component(r.get("component_role") or ""): dict(r) for r in rows}


def _relative_diff(a: float, b: float, floor: float = 0.001) -> float:
    a = float(a or 0)
    b = float(b or 0)
    denom = max((abs(a) + abs(b)) / 2.0, floor)
    return abs(a - b) / denom


def compare_signature_component(current: dict, candidate: dict) -> dict:
    dims = ["length", "width", "thickness"]
    size_diff = sum(_relative_diff(current.get(k), candidate.get(k), 0.25) for k in dims) / 3.0
    mass_diff = _relative_diff(current.get("mass"), candidate.get("mass"), 0.1)

    center_distance = math.sqrt(
        (float(current.get("center_x") or 0) - float(candidate.get("center_x") or 0)) ** 2 +
        (float(current.get("center_y") or 0) - float(candidate.get("center_y") or 0)) ** 2 +
        (float(current.get("center_z") or 0) - float(candidate.get("center_z") or 0)) ** 2
    )
    cur_diag = math.sqrt(sum(float(current.get(k) or 0) ** 2 for k in dims))
    cand_diag = math.sqrt(sum(float(candidate.get(k) or 0) ** 2 for k in dims))
    center_diff = center_distance / max((cur_diag + cand_diag) / 2.0, 1.0)

    # Lower score is better. Size and mass dominate; center/COG separates same-size stacks.
    diff_score = (0.45 * size_diff) + (0.35 * mass_diff) + (0.20 * center_diff)
    similarity = max(0.0, min(100.0, 100.0 * (1.0 - diff_score)))

    return {
        "similarity": round(similarity, 2),
        "diff_score": round(diff_score, 6),
        "size_diff_pct": round(size_diff * 100.0, 3),
        "mass_diff_pct": round(mass_diff * 100.0, 3),
        "center_distance_in": round(center_distance, 4),
        "center_diff_pct": round(center_diff * 100.0, 3),
        "current": current,
        "candidate": candidate,
    }


def compare_job_signature_to_previous(conn, job_num: str, limit: int = JOB_SIGNATURE_DEFAULT_LIMIT) -> list[dict]:
    j = normalize_job_num(job_num)
    current = load_job_signature(conn, j)
    if not current:
        return []

    candidates = _rowdicts(conn.execute("""
        SELECT DISTINCT job_num
        FROM job_signature_components
        WHERE job_num<>?
        ORDER BY job_num DESC
    """, (j,)))

    results = []
    for c in candidates:
        cand_job = normalize_job_num(c.get("job_num") or "")
        if not cand_job:
            continue
        cand_sig = load_job_signature(conn, cand_job)
        component_results = []
        missing_current = []
        missing_candidate = []

        for role in JOB_SIGNATURE_COMPONENTS:
            cur = current.get(role)
            cand = cand_sig.get(role)
            if not cur:
                missing_current.append(role)
                continue
            if not cand:
                missing_candidate.append(role)
                continue
            comp = compare_signature_component(cur, cand)
            comp["component_role"] = role
            component_results.append(comp)

        if not component_results:
            continue

        avg_similarity = sum(x["similarity"] for x in component_results) / len(component_results)
        coverage = len(component_results) / len(JOB_SIGNATURE_COMPONENTS)
        final_score = avg_similarity * (0.75 + 0.25 * coverage)

        results.append({
            "job_num": cand_job,
            "score": round(final_score, 2),
            "avg_component_similarity": round(avg_similarity, 2),
            "matched_components": len(component_results),
            "coverage_pct": round(coverage * 100.0, 1),
            "missing_current": missing_current,
            "missing_candidate": missing_candidate,
            "components": sorted(component_results, key=lambda x: x["component_role"]),
        })

    results.sort(key=lambda x: x["score"], reverse=True)
    return results[:max(1, int(limit or JOB_SIGNATURE_DEFAULT_LIMIT))]


def sync_matching_software_signatures(conn, folder: str = ELGIN_MATCHING_SOFTWARE_DIR) -> dict:
    """
    Import every XT_Export_Job_Signature CSV found in the shared matching folder.
    Scans the matching root (flat files) and one level of job subfolders so
    macro publish paths both work without manual upload.
    """
    result = {"folder": folder, "scanned": 0, "imported_files": 0, "imported_components": 0, "errors": []}
    if not folder or not os.path.isdir(folder):
        return result

    seen_paths: set[str] = set()

    def _job_num_from_signature_filename(name: str) -> str:
        base = os.path.splitext(os.path.basename(name))[0]
        m = re.search(r"job_signature[_-]?(.*)$", base, re.I)
        if m and m.group(1):
            return normalize_job_num(m.group(1))
        return normalize_job_num(base.split("_")[0])

    def _consider(path: str) -> None:
        if path in seen_paths:
            return
        seen_paths.add(path)
        result["scanned"] += 1
        name = os.path.basename(path)
        try:
            with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
                data = f.read()
            fallback_job = _job_num_from_signature_filename(name)
            rows = parse_job_signature_csv_text(data, fallback_job_num=fallback_job)
            if not rows:
                return
            job_num = normalize_job_num(rows[0].get("job_num") or fallback_job)
            if not job_num:
                return
            imported = upsert_job_signature_rows(conn, job_num, rows, source_file=path)
            result["imported_files"] += 1
            result["imported_components"] += imported
        except Exception as exc:
            result["errors"].append(f"{name}: {exc}")

    try:
        for name in sorted(os.listdir(folder)):
            path = os.path.join(folder, name)
            if os.path.isfile(path) and name.lower().endswith(".csv") and "job_signature" in name.lower():
                _consider(path)
            elif os.path.isdir(path):
                try:
                    for fn in sorted(os.listdir(path)):
                        if not fn.lower().endswith(".csv") or "job_signature" not in fn.lower():
                            continue
                        sub_path = os.path.join(path, fn)
                        if os.path.isfile(sub_path):
                            _consider(sub_path)
                except Exception as exc:
                    result["errors"].append(f"{name}/: {exc}")
    except Exception as exc:
        result["errors"].append(str(exc))
    return result


def recent_signature_match_summaries(conn, limit: int = 8) -> list[dict]:
    jobs = _rowdicts(conn.execute("""
        SELECT job_num, MAX(imported_at) AS imported_at
        FROM job_signature_components
        GROUP BY job_num
        ORDER BY imported_at DESC
        LIMIT ?
    """, (max(1, int(limit or 8)),)))
    out = []
    for row in jobs:
        job_num = normalize_job_num(row.get("job_num") or "")
        if not job_num:
            continue
        matches = compare_job_signature_to_previous(conn, job_num, limit=3)
        out.append({
            "job_num": job_num,
            "imported_at": row.get("imported_at") or "",
            "matches": matches,
        })
    return out

def job_family_root(conn, job_num: str) -> str:
    """
    Returns the physical/matched family root.

    Example:
      J8342 matched old J8206 -> family root J8206.
    """
    j = normalize_job_num(job_num)
    if not j:
        return ""

    row = conn.execute("""
        SELECT matched_job_num
        FROM job_series
        WHERE job_num=?
    """, (j,)).fetchone()

    if row:
        matched = normalize_job_num(row[0] or "")
        if matched and not job_num_is_all_zeros(matched):
            return matched

    return j


def component_is_complete(conn, job_num: str, component_name: str) -> bool:
    j = normalize_job_num(job_num)
    c = canonical_component(component_name or "")

    if not j or not c:
        return False

    row = conn.execute("""
        SELECT status
        FROM job_components
        WHERE job_num=? AND component_name=?
    """, (
        j,
        c,
    )).fetchone()

    return bool(row and str(row[0] or "").upper() == "COMPLETE")


def job_is_complete(conn, job_num: str) -> bool:
    j = normalize_job_num(job_num)
    if not j:
        return False

    rows = _rowdicts(conn.execute("""
        SELECT status
        FROM job_components
        WHERE job_num=?
    """, (j,)))

    if not rows:
        return False

    return all(str(r.get("status") or "").upper() == "COMPLETE" for r in rows)


def job_due_is_old_or_missing(conn, job_num: str) -> bool:
    """
    True means this program/job number is likely old/template/reused.

    Used for cases like:
      controller program J8206 is running,
      but current real job is J8342/J8343/J8344 from matching.
    """
    j = normalize_job_num(job_num)

    if not j:
        return True

    row = conn.execute("""
        SELECT due_date, status
        FROM job_series
        WHERE job_num=?
    """, (j,)).fetchone()

    if not row:
        return True

    due = str(row[0] or "").strip()
    status = str(row[1] or "").upper().strip()

    if status == "COMPLETE":
        return True

    if not due:
        return True

    d = _date_obj(due)
    if not d:
        return True

    if d < date.today():
        return True

    return False


def same_component_required(component_name: str) -> bool:
    """
    ID and OD absolutely must not be mixed.
    This helper lets us keep that rule explicit.
    """
    c = canonical_component(component_name or "")
    if c in ("ID POT", "OD POT", "ID HOLDER", "OD HOLDER"):
        return True
    if c in DEFAULT_JOB_COMPONENTS:
        return True
    return True


def similar_completed_component_estimate_sec(
    conn,
    job_num: str,
    component_name: str,
    tolerance: float = SIMILAR_JOB_TOLERANCE,
    limit: int = 7,
) -> dict:
    """
    Historical estimate fallback.

    Critical behavior:
    - Matches the SAME component only.
      ID POT uses completed ID POT times.
      OD POT uses completed OD POT times.
    - Prefers same physical/matched family.
    - Then falls back to close part_catalog block/CoM/weight match.
    """
    target_job = normalize_job_num(job_num)
    comp = canonical_component(component_name or "")

    if not target_job or not comp or job_num_is_all_zeros(target_job):
        return {
            "ok": False,
            "estimated_sec": 0.0,
            "source": "NO_HISTORICAL_MATCH",
            "basis": "Bad target job/component.",
            "matches": [],
        }

    target_sig = get_part_signature(conn, target_job)
    target_root = job_family_root(conn, target_job)

    rows = _rowdicts(conn.execute("""
        SELECT
            jc.job_num,
            jc.component_name,
            jc.actual_sec,
            jc.completed_at,
            js.due_date,
            js.matched_job_num,
            pc.weight,
            pc.com_x,
            pc.com_y,
            pc.com_z
        FROM job_components jc
        LEFT JOIN job_series js ON js.job_num=jc.job_num
        LEFT JOIN part_catalog pc ON pc.job_id=jc.job_num
        WHERE jc.component_name=?
          AND COALESCE(jc.actual_sec, 0) > 60
          AND jc.job_num != ?
          AND COALESCE(jc.status, '')='COMPLETE'
    """, (
        comp,
        target_job,
    )))

    candidates = []

    for r in rows:
        cand_job = normalize_job_num(r.get("job_num", ""))
        if not cand_job or job_num_is_all_zeros(cand_job):
            continue

        matched_root = normalize_job_num(r.get("matched_job_num") or "")
        cand_root = matched_root or cand_job

        same_family = bool(target_root and cand_root and target_root == cand_root)

        cand_sig = {
            "job_num": cand_job,
            "weight": float(r.get("weight") or 0),
            "com_x": float(r.get("com_x") or 0),
            "com_y": float(r.get("com_y") or 0),
            "com_z": float(r.get("com_z") or 0),
        }

        physical_match = parts_are_similar(target_sig, cand_sig, tolerance=tolerance)

        if not same_family and not physical_match:
            continue

        delta = part_signature_delta(target_sig, cand_sig) if physical_match else 999.0

        # Lower score is better.
        score = delta
        if same_family:
            score -= 10.0

        candidates.append({
            "job_num": cand_job,
            "component_name": comp,
            "actual_sec": float(r.get("actual_sec") or 0),
            "completed_at": r.get("completed_at") or "",
            "due_date": r.get("due_date") or "",
            "same_family": same_family,
            "physical_match": physical_match,
            "signature_delta": round(delta, 5) if delta < 999 else None,
            "score": score,
        })

    if not candidates:
        return {
            "ok": False,
            "estimated_sec": 0.0,
            "source": "NO_HISTORICAL_MATCH",
            "basis": f"No completed same-component match found for {comp}.",
            "matches": [],
        }

    candidates.sort(key=lambda x: (
        x["score"],
        x.get("due_date") or "9999-99-99",
        x["job_num"],
    ))

    selected = candidates[:max(1, int(limit or 7))]
    vals = sorted(float(x["actual_sec"] or 0) for x in selected if float(x["actual_sec"] or 0) > 60)

    if not vals:
        return {
            "ok": False,
            "estimated_sec": 0.0,
            "source": "NO_HISTORICAL_MATCH",
            "basis": "Historical candidates had no useful actual time.",
            "matches": selected,
        }

    # Median is safer than average when one job had rework, resize, error recovery, etc.
    mid = len(vals) // 2
    if len(vals) % 2:
        estimate = vals[mid]
    else:
        estimate = (vals[mid - 1] + vals[mid]) / 2.0

    basis = (
        f"Historical estimate from {len(vals)} completed same-component "
        f"{comp} job(s): "
        + ", ".join("J" + x["job_num"] for x in selected[:5])
    )

    return {
        "ok": True,
        "estimated_sec": round(float(estimate), 2),
        "source": "SIMILAR_COMPLETED_JOB",
        "basis": basis,
        "matches": selected,
    }


def save_component_estimate(
    conn,
    job_num: str,
    component_name: str,
    estimated_sec: float,
    source: str,
    basis: str,
):
    """
    Saves a learned estimate without falling back to fixed 45 minutes.
    """
    j = normalize_job_num(job_num)
    c = canonical_component(component_name or "")

    if not j or not c or job_num_is_all_zeros(j):
        return 0.0

    estimated_sec = float(estimated_sec or 0)
    if estimated_sec <= 0:
        return 0.0

    row = conn.execute("""
        SELECT estimated_sec
        FROM job_components
        WHERE job_num=? AND component_name=?
    """, (
        j,
        c,
    )).fetchone()

    old_est = float(row[0] or 0) if row else 0.0

    if old_est <= 0:
        new_est = estimated_sec
    else:
        # Do not let a partial Q-code timer collapse a good estimate.
        if source == "QCODE_CYCLE_TIMER" and estimated_sec < old_est * 0.70:
            new_est = old_est
        else:
            new_est = (old_est * 0.65) + (estimated_sec * 0.35)

    note = f"Estimate updated from {source}: {fsec_py(estimated_sec)}"

    conn.execute("""
        UPDATE job_components
        SET
            estimated_sec=?,
            estimate_source=?,
            estimate_basis=?,
            estimate_updated_at=datetime('now'),
            last_status_note=CASE
                WHEN last_status_note IS NULL OR last_status_note=''
                THEN ?
                ELSE last_status_note || ' | ' || ?
            END
        WHERE job_num=? AND component_name=?
    """, (
        round(new_est, 2),
        source or "",
        basis or "",
        note,
        note,
        j,
        c,
    ))

    return round(new_est, 2)


def detected_component_for_program(parsed: dict, prog_raw: str) -> str:
    comp = (
        detect_component((parsed or {}).get("sub_cat") or "")
        or detect_component((parsed or {}).get("part_name") or "")
        or detect_component(prog_raw or "")
        or canonical_component((parsed or {}).get("sub_cat") or (parsed or {}).get("part_name") or prog_raw or "")
        or "UNCLASSIFIED"
    )

    return canonical_component(comp)


def find_active_matched_jobs_for_template(
    conn,
    template_job_num: str,
    component_name: str,
    tolerance: float = SIMILAR_JOB_TOLERANCE,
) -> list:
    """
    Finds active/current jobs that should receive runtime from an old reused program.

    Example:
      Program says J8206.
      J8206 is old/no due/overdue.
      Current matched jobs are J8342, J8343, J8344.
      Return whichever matched job still has this exact component open.
    """
    template = normalize_job_num(template_job_num)
    comp = canonical_component(component_name or "")

    if not template or job_num_is_all_zeros(template):
        return []

    direct = _rowdicts(conn.execute("""
        SELECT *
        FROM job_series
        WHERE matched_job_num=?
          AND job_num!=?
          AND COALESCE(status, 'ACTIVE')!='COMPLETE'
    """, (
        template,
        template,
    )))

    # Physical fallback if matched_job_num is not filled in.
    physical = []
    template_sig = get_part_signature(conn, template)

    if part_signature_is_meaningful(template_sig):
        rows = _rowdicts(conn.execute("""
            SELECT js.*
            FROM job_series js
            JOIN part_catalog pc ON pc.job_id=js.job_num
            WHERE js.job_num!=?
              AND COALESCE(js.status, 'ACTIVE')!='COMPLETE'
        """, (template,)))

        for r in rows:
            j = normalize_job_num(r.get("job_num", ""))
            if not j or job_num_is_all_zeros(j):
                continue

            sig = get_part_signature(conn, j)
            if parts_are_similar(template_sig, sig, tolerance=tolerance):
                physical.append(r)

    by_job = {}
    for r in direct + physical:
        j = normalize_job_num(r.get("job_num", ""))
        if j and not job_num_is_all_zeros(j):
            by_job[j] = r

    candidates = []

    for j, row in by_job.items():
        if job_is_complete(conn, j):
            continue

        # ID and OD must not be mixed.
        # Only choose jobs where this exact component is not complete.
        if comp and comp not in ("UNCLASSIFIED", "N/A", "IDLE"):
            if component_is_complete(conn, j, comp):
                continue

        due = row.get("due_date") or ""
        d = _date_obj(due)

        candidates.append({
            **row,
            "job_num": j,
            "_due_sort": d or date.max,
            "_has_due": bool(d),
        })

    candidates.sort(key=lambda r: (
        r["_due_sort"],
        r["job_num"],
    ))

    return candidates


def resolve_effective_job_from_matching(
    conn,
    state: dict,
    parsed: dict,
    prog_raw: str,
    machine: str,
) -> dict:
    """
    Converts a reused/old program job number into the current active shop job.

    Example:
      Machine program: J8206 ID POT
      J8206 is old/no due/overdue
      Current matched jobs: J8342, J8343, J8344
      Effective runtime should go to J8342, then after completion to J8343, etc.

    Important:
      We do NOT advance just because J8206 resets.
      We keep the selected effective job until its same component is completed,
      normally when the machine moves to a finishing/direct job such as engraving.
    """
    raw_job = normalize_job_num((parsed or {}).get("program_job_num") or (parsed or {}).get("job_num") or "")

    if not raw_job or job_num_is_all_zeros(raw_job):
        return parsed

    comp = detected_component_for_program(parsed, prog_raw)

    # If no candidate current matched jobs exist, use the program job directly.
    candidate_jobs = find_active_matched_jobs_for_template(conn, raw_job, comp)

    if not candidate_jobs:
        return parsed

    # If the program job itself is current, not old/missing/overdue, use it directly.
    if not job_due_is_old_or_missing(conn, raw_job):
        return parsed

    key = f"{raw_job}|{comp or 'UNCLASSIFIED'}"
    state.setdefault("_effective_job_for_template", {})

    current_effective = normalize_job_num(
        state["_effective_job_for_template"].get(key, "")
    )

    candidate_nums = [normalize_job_num(c.get("job_num", "")) for c in candidate_jobs]

    if (
        current_effective
        and current_effective in candidate_nums
        and not component_is_complete(conn, current_effective, comp)
    ):
        effective = current_effective
    else:
        effective = candidate_nums[0] if candidate_nums else ""

    if not effective:
        return parsed

    state["_effective_job_for_template"][key] = effective

    out = dict(parsed)
    out["program_job_num"] = raw_job
    out["effective_job_num"] = effective
    out["job_num"] = effective
    out["matched_from_program_job"] = raw_job
    out["matched_reason"] = (
        f"Program J{raw_job} treated as old/template; "
        f"runtime assigned to active matched job J{effective}."
    )

    display_part = out.get("part_name") or out.get("sub_cat") or comp or ""

    out["display"] = (
        f"J{effective}"
        + (f" · {display_part}" if display_part else "")
        + f"  ← program J{raw_job}"
    )
    out["full_name"] = f"{effective} {display_part}".strip()

    return out


def compute_job_financial(conn, job_num: str):
    """
    Admin-only calculations are performed here.

    Labor-rate behavior:
    - Stored labor_rate > 0 is honored.
    - Only if no labor_rate is saved does it fall back to DEFAULT_LABOR_RATE_PER_HOUR.
    """
    job_num = normalize_job_num(job_num)
    if not job_num:
        return {}

    row = conn.execute("SELECT * FROM job_series WHERE job_num=?", (job_num,)).fetchone()
    if not row:
        return {}

    js = dict(row)

    actual_sec = conn.execute("""
        SELECT COALESCE(SUM(actual_sec), 0)
        FROM job_components
        WHERE job_num=?
    """, (job_num,)).fetchone()[0] or 0

    revenue = _money(js.get("revenue"))
    steel = _money(js.get("steel_cost"))
    grinding = _money(js.get("grinding_cost"))
    pyropel = _money(js.get("pyropel_cost"))

    saved_labor_rate = _money(js.get("labor_rate"))
    effective_labor_rate = saved_labor_rate if saved_labor_rate > 0 else DEFAULT_LABOR_RATE_PER_HOUR

    labor_hours = float(actual_sec or 0) / 3600.0
    calculated_labor_cost = round(labor_hours * effective_labor_rate, 2)

    labor_override_cost = _money(js.get("labor_override_cost"))
    labor_cost = labor_override_cost if labor_override_cost > 0 else calculated_labor_cost

    total_cost = round(steel + grinding + pyropel + labor_cost, 2)
    profit = round(revenue - total_cost, 2)
    margin_pct = round((profit / revenue) * 100.0, 2) if revenue > 0 else 0.0

    return {
        "actual_sec": round(float(actual_sec or 0), 2),
        "labor_hours": round(labor_hours, 3),
        "labor_rate_saved": saved_labor_rate,
        "labor_rate_effective": effective_labor_rate,
        "labor_cost": labor_cost,
        "labor_override_cost": labor_override_cost,
        "total_cost": total_cost,
        "profit": profit,
        "margin_pct": margin_pct,
    }


def log_job_status(conn, job_num, component_name, status, machine, prog_raw, message):
    job_num = normalize_job_num(job_num)
    if not job_num:
        return

    conn.execute("""
        INSERT INTO job_status_log (
            job_num, component_name, status, machine, prog_raw, message
        )
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        job_num,
        component_name or "",
        status or "",
        machine or "",
        prog_raw or "",
        message or "",
    ))


def delete_job_everywhere(conn, job_num: str):
    """
    Deletes a job from every relevant table.

    This intentionally allows deleting fake all-zero jobs like J0000,
    because old bad rows may already exist in the database.
    """
    j = normalize_job_num(job_num)
    if not j:
        raise HTTPException(400, "Bad job number")

    deleted = {}

    tables_simple = [
        ("job_components", "job_num"),
        ("job_status_log", "job_num"),
        ("job_runs", "job_num"),
        ("job_series", "job_num"),
        ("schedule", "job_id"),
        ("part_catalog", "job_id"),
        ("shop_board_rows", "job_num"),
        ("machine_schedule_plan", "job_num"),
        ("autoloader_queue", "requested_job_num"),
        ("steel_sheet_items", "job_num"),
        ("steel_sheet_parse_log", "job_num"),
    ]

    for table, col in tables_simple:
        try:
            cur = conn.execute(f"DELETE FROM {table} WHERE {col}=?", (j,))
            deleted[table] = cur.rowcount
        except Exception:
            deleted[table] = 0

    try:
        cur = conn.execute("""
            UPDATE autoloader_shelves
            SET
                assigned_job_num='',
                component_name='',
                safe_to_run=0,
                notes=CASE
                    WHEN notes IS NULL OR notes=''
                    THEN ?
                    ELSE notes || ' | ' || ?
                END,
                updated_at=datetime('now')
            WHERE assigned_job_num=?
        """, (
            f"Cleared deleted job J{j}.",
            f"Cleared deleted job J{j}.",
            j,
        ))
        deleted["autoloader_shelves_cleared"] = cur.rowcount
    except Exception:
        deleted["autoloader_shelves_cleared"] = 0

    try:
        cur = conn.execute("""
            UPDATE tool_cut_samples
            SET job_num=''
            WHERE job_num=?
        """, (j,))
        deleted["tool_cut_samples_jobnum_cleared"] = cur.rowcount
    except Exception:
        deleted["tool_cut_samples_jobnum_cleared"] = 0

    return {
        "job_num": j,
        "deleted": deleted,
    }


def rename_or_merge_job_everywhere(conn, old_job_num: str, new_job_num: str, merge_if_exists: bool = True):
    """
    Rename or merge a job.

    - Old job may be fake like J0000 so it can be fixed.
    - New job must be valid and cannot be all-zero.
    """
    old = normalize_job_num(old_job_num)
    new = validate_new_job_num_or_400(new_job_num)

    if not old:
        raise HTTPException(400, "Old job number required")
    if old == new:
        return {"status": "noop", "old_job_num": old, "new_job_num": new}

    target_exists = conn.execute(
        "SELECT 1 FROM job_series WHERE job_num=?",
        (new,),
    ).fetchone()

    if target_exists and not merge_if_exists:
        raise HTTPException(409, f"J{new} already exists. Enable merge_if_exists to merge.")

    if target_exists:
        ensure_job_series(conn, new)

        old_components = _rowdicts(conn.execute("""
            SELECT *
            FROM job_components
            WHERE job_num=?
        """, (old,)))

        for c in old_components:
            comp = c.get("component_name") or "UNCLASSIFIED"

            target_comp = conn.execute("""
                SELECT *
                FROM job_components
                WHERE job_num=? AND component_name=?
            """, (
                new,
                comp,
            )).fetchone()

            if target_comp:
                conn.execute("""
                    UPDATE job_components
                    SET
                        actual_sec=COALESCE(actual_sec,0)+?,
                        estimated_sec=CASE
                            WHEN COALESCE(estimated_sec,0)>0 THEN estimated_sec
                            ELSE COALESCE(?,0)
                        END,
                        estimate_source=CASE
                            WHEN COALESCE(estimate_source,'')!='' THEN estimate_source
                            ELSE COALESCE(?, '')
                        END,
                        estimate_basis=CASE
                            WHEN COALESCE(estimate_basis,'')!='' THEN estimate_basis
                            ELSE COALESCE(?, '')
                        END,
                        status=CASE
                            WHEN status='COMPLETE' OR ?='COMPLETE' THEN 'COMPLETE'
                            WHEN status='RUNNING' OR ?='RUNNING' THEN 'RUNNING'
                            ELSE status
                        END,
                        last_status_note=CASE
                            WHEN last_status_note IS NULL OR last_status_note=''
                            THEN ?
                            ELSE last_status_note || ' | ' || ?
                        END
                    WHERE job_num=? AND component_name=?
                """, (
                    float(c.get("actual_sec") or 0),
                    float(c.get("estimated_sec") or 0),
                    c.get("estimate_source") or "",
                    c.get("estimate_basis") or "",
                    c.get("status") or "",
                    c.get("status") or "",
                    f"Merged from J{old}.",
                    f"Merged from J{old}.",
                    new,
                    comp,
                ))

                conn.execute("""
                    DELETE FROM job_components
                    WHERE job_num=? AND component_name=?
                """, (
                    old,
                    comp,
                ))

            else:
                conn.execute("""
                    UPDATE job_components
                    SET job_num=?
                    WHERE job_num=? AND component_name=?
                """, (
                    new,
                    old,
                    comp,
                ))

        conn.execute("DELETE FROM job_series WHERE job_num=?", (old,))

        target_catalog = conn.execute(
            "SELECT 1 FROM part_catalog WHERE job_id=?",
            (new,),
        ).fetchone()

        if target_catalog:
            conn.execute("DELETE FROM part_catalog WHERE job_id=?", (old,))
        else:
            conn.execute("UPDATE part_catalog SET job_id=? WHERE job_id=?", (new, old))

        target_board = conn.execute(
            "SELECT 1 FROM shop_board_rows WHERE job_num=?",
            (new,),
        ).fetchone()

        if target_board:
            conn.execute("DELETE FROM shop_board_rows WHERE job_num=?", (old,))
        else:
            conn.execute("""
                UPDATE shop_board_rows
                SET job_num=?, updated_at=datetime('now')
                WHERE job_num=?
            """, (new, old))

    else:
        rename_tables = [
            ("job_series", "job_num"),
            ("job_components", "job_num"),
            ("job_status_log", "job_num"),
            ("job_runs", "job_num"),
            ("schedule", "job_id"),
            ("part_catalog", "job_id"),
            ("shop_board_rows", "job_num"),
            ("steel_sheet_items", "job_num"),
            ("steel_sheet_parse_log", "job_num"),
        ]

        for table, col in rename_tables:
            try:
                conn.execute(f"UPDATE {table} SET {col}=? WHERE {col}=?", (new, old))
            except Exception:
                pass

        try:
            conn.execute("""
                UPDATE shop_board_rows
                SET updated_at=datetime('now')
                WHERE job_num=?
            """, (new,))
        except Exception:
            pass

        ensure_job_series(conn, new)

    more_updates = [
        ("machine_schedule_plan", "job_num"),
        ("autoloader_queue", "requested_job_num"),
        ("autoloader_shelves", "assigned_job_num"),
        ("tool_cut_samples", "job_num"),
    ]

    for table, col in more_updates:
        try:
            conn.execute(f"UPDATE {table} SET {col}=? WHERE {col}=?", (new, old))
        except Exception:
            pass

    try:
        conn.execute("""
            INSERT INTO job_status_log (
                job_num, status, message
            )
            VALUES (?, 'RENAMED', ?)
        """, (
            new,
            f"Admin renamed/merged J{old} into J{new}.",
        ))
    except Exception:
        pass

    return {
        "old_job_num": old,
        "new_job_num": new,
        "merged": bool(target_exists),
    }


# ============================================================
# JOB MATCHING + PDF UPLOAD HELPERS
# ============================================================
def save_uploaded_steel_pdf_for_job(
    conn,
    job_num: str,
    pdf_bytes: bytes,
    original_filename: str = "",
    due_date: str = "",
):
    """
    Main steel-sheet PDF workflow.

    - The UI uses this only from Job Matching.
    - URLs are no longer the main workflow.
    """
    j = validate_new_job_num_or_400(job_num)

    if not pdf_bytes:
        return {
            "status": "no_file",
            "job_num": j,
            "message": "No PDF uploaded.",
            "count": 0,
            "items": [],
        }

    original_name = safe_upload_filename(original_filename or "steel_sheet.pdf")

    if not original_name.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a PDF steel sheet")

    stored_name = f"J{j}_{int(time.time())}_{original_name}"
    stored_path = os.path.join(STEEL_UPLOAD_DIR, stored_name)

    with open(stored_path, "wb") as f:
        f.write(pdf_bytes)

    ensure_job_series(conn, j)

    if due_date:
        conn.execute("""
            UPDATE job_series
            SET due_date=?, updated_at=datetime('now')
            WHERE job_num=?
        """, (due_date, j))

        conn.execute("""
            UPDATE part_catalog
            SET due_date=?
            WHERE job_id=?
        """, (due_date, j))

        conn.execute("""
            UPDATE schedule
            SET due_date=?
            WHERE job_id=?
        """, (due_date, j))

        conn.execute("""
            UPDATE shop_board_rows
            SET due_date=?, updated_at=datetime('now')
            WHERE job_num=?
        """, (due_date, j))

    conn.execute("""
        UPDATE job_series
        SET
            steel_sheet_file=?,
            steel_sheet_original_name=?,
            steel_sheet_url='',
            updated_at=datetime('now')
        WHERE job_num=?
    """, (
        stored_path,
        original_name,
        j,
    ))

    result = parse_and_store_steel_pdf_bytes(
        conn=conn,
        job_num=j,
        pdf_bytes=pdf_bytes,
        original_filename=original_name,
        stored_path=stored_path,
        force=True,
    )

    return {
        "status": "ok",
        "job_num": j,
        "stored_path": stored_path,
        "original_filename": original_name,
        "parse": result,
    }


def catalog_lookup_core(
    conn,
    job_id: str,
    weight: float,
    com_x: float,
    com_y: float,
    com_z: float,
    tolerance: float = 0.01,
    due_date: str = "",
    notes: str = "",
):
    """
    Shared job-matching logic used by JSON lookup and multipart PDF lookup.

    This is also the foundation for old/reused CNC program resolution.
    If J8342 physically matches J8206, then J8342 can inherit J8206 as its
    matched_job_num/family root.
    """
    job_id = validate_new_job_num_or_400(job_id)

    tol = float(tolerance or 0.01)
    w = float(weight or 0)
    x = float(com_x or 0)
    y = float(com_y or 0)
    z = float(com_z or 0)
    due_date = due_date or ""

    inherited_from = ""
    inherited_reason = ""

    rows = _rowdicts(conn.execute("SELECT * FROM part_catalog"))

    matches = [
        r for r in rows
        if normalize_job_num(r.get("job_id", "")) != job_id
        and not job_num_is_all_zeros(r.get("job_id", ""))
        and abs(float(r.get("weight") or 0) - w) <= tol
        and abs(float(r.get("com_x") or 0) - x) <= tol
        and abs(float(r.get("com_y") or 0) - y) <= tol
        and abs(float(r.get("com_z") or 0) - z) <= tol
    ]

    if not due_date:
        best = choose_match_with_due(matches)
        if best:
            due_date = best.get("due_date") or ""
            inherited_from = normalize_job_num(best.get("job_id", ""))
            inherited_reason = "INHERITED_DUE_DATE_FROM_PHYSICAL_MATCH"

    # If due date is provided and there are physical matches, still keep a matched
    # root so reused old programs can map to the new current job.
    if not inherited_from and matches:
        # Prefer a match that is old/no-due/past, because this is often the template
        # program/job number currently reused by the CNC.
        sorted_matches = []
        for m in matches:
            mj = normalize_job_num(m.get("job_id", ""))
            mdue = m.get("due_date") or ""
            d = _date_obj(mdue)
            old_score = 0
            if not d:
                old_score -= 2
            elif d < date.today():
                old_score -= 1
            sorted_matches.append((old_score, d or date.min, mj, m))

        sorted_matches.sort(key=lambda x: (x[0], x[1], x[2]))
        inherited_from = normalize_job_num(sorted_matches[0][3].get("job_id", ""))
        inherited_reason = "PHYSICAL_MATCH_FOR_REUSED_PROGRAM"

    conn.execute("""
        INSERT OR REPLACE INTO part_catalog (
            job_id, weight, com_x, com_y, com_z, due_date, steel_sheet_url, notes
        )
        VALUES (?, ?, ?, ?, ?, ?, '', ?)
    """, (
        job_id,
        w,
        x,
        y,
        z,
        due_date,
        notes or "Reverse Lookup",
    ))

    ensure_job_series(
        conn,
        job_id,
        due_date,
        "",
        notes or "",
        inherited_from,
        inherited_reason,
    )

    conn.execute("""
        INSERT INTO shop_board_rows (
            job_num, description, due_date, steel, notes, row_status, updated_at
        )
        SELECT ?, ?, ?, ?, ?, 'OPEN', datetime('now')
        WHERE NOT EXISTS (
            SELECT 1 FROM shop_board_rows WHERE job_num=?
        )
    """, (
        job_id,
        notes or "Reverse lookup",
        due_date,
        "",
        "Created from reverse lookup.",
        job_id,
    ))

    specs = f"J{job_id} W:{w}, X:{x}, Y:{y}, Z:{z}, Due:{due_date}"
    match_str = ", ".join("J" + normalize_job_num(m["job_id"]) for m in matches) if matches else "NO MATCH"

    if inherited_from:
        match_str += f" | MATCHED FAMILY ROOT J{inherited_from}"

    conn.execute(
        "INSERT INTO match_history (specs, matches) VALUES (?, ?)",
        (specs, match_str),
    )

    return {
        "matches": matches,
        "specs": specs,
        "match_str": match_str,
        "job_id": job_id,
        "due_date": due_date,
        "inherited_from": inherited_from,
    }


# ============================================================
# UMC-1000 POT SCHEDULER HELPERS
# ============================================================
def _parse_dt_for_schedule(s: str):
    if not s:
        return None
    try:
        return datetime.strptime(str(s)[:10], "%Y-%m-%d")
    except Exception:
        return None


def steel_status_from_text(v: str) -> dict:
    raw = str(v or "").strip()
    u = raw.upper()

    if not u:
        return {
            "ready": False,
            "status": "",
            "reason": "No steel status entered.",
        }

    if any(w in u for w in STEEL_WAIT_WORDS):
        return {
            "ready": False,
            "status": raw,
            "reason": f"Steel appears not ready: {raw}",
        }

    if any(w in u for w in STEEL_READY_WORDS):
        return {
            "ready": True,
            "status": raw,
            "reason": f"Steel marked ready: {raw}",
        }

    return {
        "ready": False,
        "status": raw,
        "reason": f"Steel status unclear: {raw}",
    }


def get_shop_board_steel_status(conn, job_num: str) -> dict:
    job_num = normalize_job_num(job_num)

    row = conn.execute("""
        SELECT steel
        FROM shop_board_rows
        WHERE job_num=?
        ORDER BY id DESC
        LIMIT 1
    """, (job_num,)).fetchone()

    if row:
        return steel_status_from_text(row[0])

    return {
        "ready": False,
        "status": "",
        "reason": "No shop-board steel status found.",
    }


def infer_steel_ready(conn, job_series_row: dict) -> dict:
    """
    Uses:
    1. shop_board_rows.steel text first for actual physical readiness
    2. parsed steel-sheet rows as evidence that requirements exist
    3. uploaded steel sheet file as weak evidence only

    Important:
    Having a steel sheet does NOT mean steel is physically ready.
    """
    job_num = normalize_job_num(job_series_row.get("job_num", ""))

    board = get_shop_board_steel_status(conn, job_num)

    if board["ready"]:
        return board

    parsed_summary = steel_sheet_summary_for_job(conn, job_num)

    steel_file = str(job_series_row.get("steel_sheet_file") or "").strip()
    steel_url = str(job_series_row.get("steel_sheet_url") or "").strip()

    if parsed_summary.get("count", 0) > 0:
        if board.get("status"):
            return {
                "ready": False,
                "status": board.get("status", ""),
                "reason": board.get("reason", "") + " | Steel sheet parsed: " + parsed_summary.get("message", ""),
            }

        return {
            "ready": False,
            "status": "STEEL SHEET PARSED",
            "reason": "Steel requirements parsed, but shop board steel is not marked READY/CUT/RECEIVED.",
        }

    if steel_file or steel_url:
        if board.get("status"):
            return board

        return {
            "ready": False,
            "status": "STEEL SHEET ATTACHED",
            "reason": "Steel sheet exists, but steel is not marked ready.",
        }

    return board


def job_family_root_from_series(job: dict) -> str:
    job_num = normalize_job_num(job.get("job_num", ""))
    matched = normalize_job_num(job.get("matched_job_num", ""))
    return matched or job_num


def due_bucket(due_date: str) -> int:
    """
    Lower number means more urgent.
    """
    d = _date_obj(due_date)
    today = date.today()

    if not d:
        return 5

    days = (d - today).days

    if days < 0:
        return 0
    if days == 0:
        return 1
    if days <= 2:
        return 2
    if days <= 7:
        return 3
    return 4


def due_sort_value(due_date: str):
    d = _date_obj(due_date)
    if not d:
        return date.max
    return d


def scheduler_priority_reason(job: dict, component: dict, steel: dict) -> str:
    due = job.get("due_date") or ""
    bucket = due_bucket(due)

    if bucket == 0:
        due_reason = "OVERDUE"
    elif bucket == 1:
        due_reason = "DUE TODAY"
    elif bucket == 2:
        due_reason = "DUE WITHIN 2 DAYS"
    elif bucket == 3:
        due_reason = "DUE THIS WEEK"
    elif bucket == 4:
        due_reason = "DUE LATER"
    else:
        due_reason = "NO DUE DATE"

    steel_reason = "STEEL READY" if steel.get("ready") else "STEEL NOT CONFIRMED READY"

    return f"{due_reason}; {steel_reason}; {component.get('component_name', '')}"


def component_estimate_sec(component: dict, conn=None) -> float:
    """
    Runtime estimate priority:
    1. learned estimated_sec from Q-code/G-code/history
    2. actual_sec if meaningful
    3. optional historical similar-job fallback
    4. 0 = unknown estimate

    No fixed 45-minute fallback.
    """
    est = float(component.get("estimated_sec") or 0)
    actual = float(component.get("actual_sec") or 0)

    if est > 0:
        if actual > 60:
            return round((est * 0.70) + (actual * 0.30), 2)
        return round(est, 2)

    if actual > 60:
        return round(actual, 2)

    if conn is not None:
        hist = similar_completed_component_estimate_sec(
            conn,
            job_num=component.get("job_num", ""),
            component_name=component.get("component_name", ""),
        )
        if hist.get("ok") and float(hist.get("estimated_sec") or 0) > 0:
            return round(float(hist["estimated_sec"]), 2)

    return DEFAULT_UNESTIMATED_COMPONENT_SEC


def setup_time_between(prev_item: dict, item: dict) -> tuple:
    """
    Returns setup seconds and reason.
    """
    if not prev_item:
        return 0, "First scheduled item."

    prev_comp = normalize_component(prev_item.get("component_name", ""))
    comp = normalize_component(item.get("component_name", ""))

    prev_family = prev_item.get("family_root", "")
    family = item.get("family_root", "")

    if prev_comp == comp and prev_family == family:
        return UMC_SAME_COMPONENT_SETUP_CHANGE_SEC, "Same component and similar family; minimal setup transition."

    if prev_comp == comp:
        return UMC_SIMILAR_JOB_SETUP_CHANGE_SEC, "Same pot type; reduced setup transition."

    return UMC_DEFAULT_SETUP_CHANGE_SEC, f"Setup change from {prev_comp} to {comp}."


def build_umc1000_pot_schedule(conn, start_dt: datetime = None, save: bool = True) -> list:
    """
    Creates an efficient UMC-1000 pot schedule.

    Uses fluid estimated_sec values learned from G-code/Q-code/runtime/history.
    Does not assume every pot is 45 minutes.
    """
    if start_dt is None:
        start_dt = datetime.now()

    conn.row_factory = sqlite3.Row

    job_rows = _rowdicts(conn.execute("""
        SELECT *
        FROM job_series
        WHERE COALESCE(status, 'ACTIVE') != 'COMPLETE'
    """))

    jobs_by_num = {
        normalize_job_num(j.get("job_num", "")): j
        for j in job_rows
        if not job_num_is_all_zeros(j.get("job_num", ""))
    }

    comp_rows = _rowdicts(conn.execute("""
        SELECT *
        FROM job_components
        WHERE machine=?
          AND component_name IN (?, ?)
          AND COALESCE(status, 'PENDING') != 'COMPLETE'
        ORDER BY job_num, sort_order, component_name
    """, (
        UMC_SCHEDULER_MACHINE,
        "ID POT",
        "OD POT",
    )))

    candidates = []

    for c in comp_rows:
        job_num = normalize_job_num(c.get("job_num", ""))
        if not job_num or job_num_is_all_zeros(job_num):
            continue

        job = jobs_by_num.get(job_num)

        if not job:
            continue

        steel = infer_steel_ready(conn, job)
        steel_summary = steel_sheet_summary_for_job(conn, job_num)
        family = job_family_root_from_series(job)

        est = component_estimate_sec(c, conn=conn)

        estimate_source = c.get("estimate_source") or ""
        estimate_basis = c.get("estimate_basis") or ""

        if est > 0 and not estimate_source:
            hist = similar_completed_component_estimate_sec(
                conn,
                job_num=job_num,
                component_name=c.get("component_name", ""),
            )
            if hist.get("ok"):
                estimate_source = hist.get("source") or "SIMILAR_COMPLETED_JOB"
                estimate_basis = hist.get("basis") or ""

        item = {
            "machine": UMC_SCHEDULER_MACHINE,
            "job_num": job_num,
            "component_name": c.get("component_name", ""),
            "component_status": c.get("status", ""),
            "due_date": job.get("due_date") or "",
            "steel_ready": 1 if steel.get("ready") else 0,
            "steel_status": steel.get("status", ""),
            "steel_reason": steel.get("reason", ""),
            "steel_item_count": int(steel_summary.get("count") or 0),
            "steel_materials": ", ".join(steel_summary.get("materials") or []),
            "estimated_sec": est,
            "estimate_source": estimate_source,
            "estimate_basis": estimate_basis,
            "family_root": family,
            "matched_job_num": normalize_job_num(job.get("matched_job_num", "")),
            "notes": job.get("notes", ""),
            "priority_bucket": due_bucket(job.get("due_date") or ""),
            "priority_reason": scheduler_priority_reason(job, c, steel),
        }

        candidates.append(item)

    candidates.sort(key=lambda x: (
        x["priority_bucket"],
        due_sort_value(x["due_date"]),
        -int(x["steel_ready"]),
        normalize_component(x["component_name"]),
        x["family_root"],
        x["job_num"],
    ))

    optimized = []
    remaining = list(candidates)

    while remaining:
        if not optimized:
            optimized.append(remaining.pop(0))
            continue

        prev = optimized[-1]

        compatible = [
            (idx, item)
            for idx, item in enumerate(remaining)
            if item["priority_bucket"] == prev["priority_bucket"]
            and item["due_date"] == prev["due_date"]
        ]

        if not compatible:
            optimized.append(remaining.pop(0))
            continue

        def transition_score(pair):
            idx, item = pair
            setup_sec, _ = setup_time_between(prev, item)

            return (
                setup_sec,
                normalize_component(item["component_name"]) != normalize_component(prev["component_name"]),
                item["family_root"] != prev["family_root"],
                -int(item["steel_ready"]),
                item["job_num"],
            )

        best_idx, best_item = sorted(compatible, key=transition_score)[0]
        optimized.append(remaining.pop(best_idx))

    planned = []
    cursor = start_dt
    prev_item = None

    for rank, item in enumerate(optimized, start=1):
        setup_sec, setup_reason = setup_time_between(prev_item, item)
        planned_start = cursor

        # If no estimate exists, only schedule setup time and mark estimate as unknown.
        total_sec = float(item["estimated_sec"] or 0) + float(setup_sec or 0)
        planned_end_dt = datetime.fromtimestamp(planned_start.timestamp() + total_sec)

        efficiency_reason = setup_reason
        if item.get("estimate_source"):
            efficiency_reason += f" | Runtime estimate: {item.get('estimate_source')}"
        if item.get("estimate_basis"):
            efficiency_reason += f" | {item.get('estimate_basis')}"

        if not item["estimated_sec"]:
            efficiency_reason += " | Runtime estimate unknown."

        out = dict(item)
        out.update({
            "schedule_rank": rank,
            "setup_sec": setup_sec,
            "total_planned_sec": total_sec,
            "planned_start": planned_start.strftime("%Y-%m-%d %H:%M:%S"),
            "planned_end": planned_end_dt.strftime("%Y-%m-%d %H:%M:%S") if total_sec > 0 else "",
            "efficiency_reason": efficiency_reason,
        })

        planned.append(out)

        if total_sec > 0:
            cursor = planned_end_dt

        prev_item = item

    if save:
        conn.execute("""
            DELETE FROM machine_schedule_plan
            WHERE machine=?
              AND status='PLANNED'
        """, (UMC_SCHEDULER_MACHINE,))

        for item in planned:
            conn.execute("""
                INSERT INTO machine_schedule_plan (
                    machine, job_num, component_name,
                    schedule_rank, due_date,
                    steel_ready, steel_status,
                    steel_item_count, steel_materials,
                    estimated_sec, setup_sec, total_planned_sec,
                    planned_start, planned_end,
                    priority_reason, efficiency_reason,
                    status, generated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'PLANNED', datetime('now'))
            """, (
                item["machine"],
                item["job_num"],
                item["component_name"],
                item["schedule_rank"],
                item["due_date"],
                int(item["steel_ready"]),
                item["steel_status"],
                int(item.get("steel_item_count") or 0),
                item.get("steel_materials") or "",
                float(item["estimated_sec"] or 0),
                float(item["setup_sec"] or 0),
                float(item["total_planned_sec"] or 0),
                item["planned_start"],
                item["planned_end"],
                item["priority_reason"],
                item["efficiency_reason"],
            ))

    return planned


# ============================================================
# QR / BARCODE HELPERS
# ============================================================
def normalize_barcode_alias(v: str) -> str:
    return str(v or "").strip().replace("\x00", "").strip().upper()


def default_sku_from_barcode(raw: str) -> str:
    raw = str(raw or "").strip()
    cleaned = re.sub(r"[^A-Z0-9\-_\.]", "", raw.upper()).strip("._-")

    if not cleaned or len(cleaned) > 60:
        return "QR-" + hashlib.sha1(raw.encode("utf-8", errors="ignore")).hexdigest()[:12].upper()

    return cleaned[:60]


def extract_barcode_payload(raw: str) -> dict:
    """
    Supports QR contents:
      T182
      SKU:T182
      SKU=T182
      TOOL:T182
      MCN|SKU|T182|DESC|0.25 GORILLA LONG 5FL
      {"sku":"T182","description":"0.25 GORILLA LONG 5FL"}
    """
    raw_s = str(raw or "").strip().replace("\x00", "").strip()
    alias = normalize_barcode_alias(raw_s)

    out = {
        "alias": alias,
        "sku_hint": "",
        "description_hint": "",
    }

    if not raw_s:
        return out

    try:
        obj = json.loads(raw_s)
        if isinstance(obj, dict):
            out["sku_hint"] = str(obj.get("sku") or obj.get("SKU") or "").strip().upper()
            out["description_hint"] = str(obj.get("description") or obj.get("desc") or "").strip()
            return out
    except Exception:
        pass

    upper = raw_s.upper().strip()

    parts = [p.strip() for p in raw_s.split("|")]
    parts_u = [p.upper().strip() for p in parts]

    if "SKU" in parts_u:
        idx = parts_u.index("SKU")
        if idx + 1 < len(parts):
            out["sku_hint"] = default_sku_from_barcode(parts[idx + 1])

    if "DESC" in parts_u:
        idx = parts_u.index("DESC")
        if idx + 1 < len(parts):
            out["description_hint"] = parts[idx + 1].strip()

    if out["sku_hint"]:
        return out

    m = re.search(r"\bSKU\s*[:=]\s*([A-Z0-9._\-]+)", upper)
    if m:
        out["sku_hint"] = default_sku_from_barcode(m.group(1))
        return out

    m = re.search(r"\bTOOL\s*[:=]\s*([A-Z0-9._\-]+)", upper)
    if m:
        out["sku_hint"] = default_sku_from_barcode(m.group(1))
        return out

    out["sku_hint"] = default_sku_from_barcode(raw_s)
    return out


def resolve_qr_inventory(conn, barcode: str) -> dict:
    """
    QR workflow:
    - Scan first.
    - Resolve to known item if mapped.
    - If unknown, return needs_create with SKU/description hints.
    - User then chooses TAKE OUT or ADD/RESTOCK and enters quantity.
    """
    payload = extract_barcode_payload(barcode)
    alias = payload["alias"]
    sku_hint = payload["sku_hint"]
    desc_hint = payload["description_hint"]

    if not alias:
        raise HTTPException(400, "Barcode/QR required")

    alias_row = conn.execute("""
        SELECT sku, description
        FROM inventory_aliases
        WHERE alias=?
    """, (alias,)).fetchone()

    sku = ""
    mapped = False

    if alias_row:
        sku = alias_row["sku"]
        mapped = True
    else:
        sku = sku_hint or default_sku_from_barcode(alias)

    inv = conn.execute("""
        SELECT *
        FROM inventory
        WHERE sku=?
    """, (sku,)).fetchone()

    if inv:
        inv = dict(inv)
        return {
            "status": "found",
            "alias": alias,
            "sku": inv.get("sku", sku),
            "description": inv.get("description") or desc_hint or sku,
            "category": inv.get("category") or "Tool",
            "quantity": int(inv.get("quantity") or 0),
            "min_threshold": int(inv.get("min_threshold") or 0),
            "mapped": mapped,
            "needs_create": False,
            "sku_hint": sku,
            "description_hint": desc_hint,
        }

    return {
        "status": "needs_create",
        "alias": alias,
        "sku": sku,
        "description": desc_hint or "",
        "category": "Tool",
        "quantity": 0,
        "min_threshold": 5,
        "mapped": mapped,
        "needs_create": True,
        "sku_hint": sku,
        "description_hint": desc_hint,
    }


def apply_qr_inventory_action(conn, req: QRActionRequest) -> dict:
    """
    Scan-driven action:
    - Existing QR/item:
      direction OUT = take out/use
      direction IN = add/restock
    - Unknown QR:
      create item, map QR, then apply action.

    Quantity is supported for add/subtract.
    The QR identity itself is one identity per tool/SKU; this does not create
    multiple different QR codes for the same tool.
    """
    payload = extract_barcode_payload(req.barcode)
    alias = payload["alias"]

    if not alias:
        raise HTTPException(400, "Barcode/QR required")

    direction = normalize_component(req.direction or "")
    if direction not in ("IN", "OUT"):
        raise HTTPException(400, "direction must be IN or OUT")

    qty = max(1, int(req.qty or 1))

    resolved = resolve_qr_inventory(conn, req.barcode)

    sku = (req.sku or resolved.get("sku") or resolved.get("sku_hint") or "").strip().upper()
    desc = (req.description or resolved.get("description") or resolved.get("description_hint") or "").strip()
    cat = (req.category or resolved.get("category") or "Tool").strip()

    if not sku:
        sku = default_sku_from_barcode(alias)

    if not desc:
        desc = f"UNKNOWN TOOL - {sku}"

    created = False

    inv = conn.execute("""
        SELECT *
        FROM inventory
        WHERE sku=?
    """, (sku,)).fetchone()

    if not inv:
        starting_qty = max(0, int(req.starting_quantity or 0))

        conn.execute("""
            INSERT INTO inventory (
                sku, description, category, quantity, min_threshold
            )
            VALUES (?, ?, ?, ?, ?)
        """, (
            sku,
            desc,
            cat,
            starting_qty,
            max(0, int(req.min_threshold or 0)),
        ))

        created = True

    else:
        conn.execute("""
            UPDATE inventory
            SET
                description=CASE WHEN ?!='' THEN ? ELSE description END,
                category=CASE WHEN ?!='' THEN ? ELSE category END
            WHERE sku=?
        """, (
            desc, desc,
            cat, cat,
            sku,
        ))

    conn.execute("""
        INSERT INTO inventory_aliases (
            alias, sku, description, created_at, updated_at
        )
        VALUES (?, ?, ?, datetime('now'), datetime('now'))
        ON CONFLICT(alias) DO UPDATE SET
            sku=excluded.sku,
            description=excluded.description,
            updated_at=datetime('now')
    """, (
        alias,
        sku,
        desc,
    ))

    inv_qty_row = conn.execute("""
        SELECT quantity
        FROM inventory
        WHERE sku=?
    """, (sku,)).fetchone()

    old_qty = int(inv_qty_row[0] or 0) if inv_qty_row else 0

    if direction == "IN":
        new_qty = old_qty + qty
    else:
        new_qty = max(0, old_qty - qty)

    conn.execute("""
        UPDATE inventory
        SET quantity=?
        WHERE sku=?
    """, (
        new_qty,
        sku,
    ))

    conn.execute("""
        INSERT INTO transactions (sku, direction, qty)
        VALUES (?, ?, ?)
    """, (
        sku,
        direction,
        qty,
    ))

    return {
        "status": "ok",
        "alias": alias,
        "sku": sku,
        "description": desc,
        "category": cat,
        "direction": direction,
        "qty": qty,
        "old_quantity": old_qty,
        "new_quantity": new_qty,
        "created": created,
        "mapped": True,
    }


# ============================================================
# JOB COMPONENT RUNTIME UPDATE
# ============================================================
def fsec_py(s):
    s = int(round(float(s or 0)))
    h = s // 3600
    m = (s % 3600) // 60
    sec = s % 60
    if h > 0:
        return f"{h}h {m}m"
    return f"{m}m {sec}s"


def update_component_estimate_from_program(
    conn,
    job_num: str,
    component_name: str,
    machine: str,
    prog_raw: str,
    cycle_sec: float = 0.0,
):
    """
    Smart component estimate.

    Priority:
    1. Q-code cycle timer if meaningful.
    2. G-code program analysis.
    3. Similar completed job with same component and similar block/part signature.
    4. Unknown / 0.

    This replaces all fixed 45-minute assumptions.
    """
    job_num = normalize_job_num(job_num)
    comp = canonical_component(component_name or "")

    if not job_num or not comp or job_num_is_all_zeros(job_num):
        return 0.0

    row = conn.execute("""
        SELECT estimated_sec
        FROM job_components
        WHERE job_num=? AND component_name=?
    """, (
        job_num,
        comp,
    )).fetchone()

    old_est = float(row[0] or 0) if row else 0.0

    estimated_sec = 0.0
    source = ""
    basis = ""

    q_cycle = float(cycle_sec or 0)

    # Q-code timer is useful, but avoid shrinking a good estimate with a partial timer.
    if q_cycle > 60 and (old_est <= 0 or q_cycle >= old_est * 0.70):
        estimated_sec = q_cycle
        source = "QCODE_CYCLE_TIMER"
        basis = f"Controller cycle timer reported {fsec_py(q_cycle)}."

    if estimated_sec <= 0:
        try:
            runtime_est = get_gcode_program_runtime_estimate_sec(conn, machine, prog_raw)
            if runtime_est.get("ok") and float(runtime_est.get("estimated_sec") or 0) > 0:
                estimated_sec = float(runtime_est.get("estimated_sec") or 0)
                source = "GCODE_ESTIMATE"
                basis = runtime_est.get("message") or "Estimated from G-code analysis."
        except Exception as e:
            logger.debug("G-code estimate failed for %s: %s", prog_raw, e)

    if estimated_sec <= 0:
        hist = similar_completed_component_estimate_sec(
            conn,
            job_num=job_num,
            component_name=comp,
        )

        if hist.get("ok") and float(hist.get("estimated_sec") or 0) > 0:
            estimated_sec = float(hist.get("estimated_sec") or 0)
            source = hist.get("source") or "SIMILAR_COMPLETED_JOB"
            basis = hist.get("basis") or "Estimated from similar completed job."

    if estimated_sec <= 0:
        return 0.0

    return save_component_estimate(
        conn=conn,
        job_num=job_num,
        component_name=comp,
        estimated_sec=estimated_sec,
        source=source,
        basis=basis,
    )


def update_job_component_from_program(
    conn,
    state: dict,
    machine: str,
    prog_raw: str,
    parsed: dict,
    actual_cutting: bool,
    delta_t: float,
):
    job_num = normalize_job_num(parsed.get("job_num") or "")
    if not job_num or job_num_is_all_zeros(job_num):
        return

    ensure_job_series(conn, job_num)
    try_inherit_due_from_catalog(conn, job_num)

    comp = detected_component_for_program(parsed, prog_raw)
    comp = comp or "UNCLASSIFIED"

    if comp.upper() in ("N/A", "IDLE", "0", ""):
        comp = "UNCLASSIFIED"

    mach_for_component = DEFAULT_COMPONENT_MACHINE.get(comp, machine)
    est_sec = float(DEFAULT_COMPONENT_EST_SEC.get(comp, 0) or 0)
    sort_order = DEFAULT_JOB_COMPONENTS.index(comp) if comp in DEFAULT_JOB_COMPONENTS else 99

    conn.execute("""
        INSERT OR IGNORE INTO job_components (
            job_num, component_name, machine, status, sort_order, estimated_sec,
            estimate_source, estimate_basis, estimate_updated_at
        )
        VALUES (?, ?, ?, 'PENDING', ?, ?, '', '', '')
    """, (
        job_num,
        comp,
        mach_for_component,
        sort_order,
        est_sec,
    ))

    update_component_estimate_from_program(
        conn=conn,
        job_num=job_num,
        component_name=comp,
        machine=machine,
        prog_raw=prog_raw,
        cycle_sec=float(state.get("cycle_sec") or 0),
    )

    key = f"{job_num}|{comp}"
    prev_key = state.get("_last_component_key")

    if prev_key and prev_key != key and "|" in prev_key:
        prev_job, prev_comp = prev_key.split("|", 1)

        conn.execute("""
            UPDATE job_components
            SET
                status='COMPLETE',
                completed_at=COALESCE(completed_at, datetime('now')),
                last_status_note='Auto-completed because machine loaded a new program/component'
            WHERE job_num=?
              AND component_name=?
              AND status='RUNNING'
              AND last_seen_machine=?
              AND COALESCE(manual_complete, 0)=0
        """, (
            prev_job,
            prev_comp,
            machine,
        ))

        log_job_status(
            conn,
            prev_job,
            prev_comp,
            "COMPLETE",
            machine,
            prog_raw,
            "Auto-completed because machine loaded a new program/component",
        )

        state["_last_cycle_sec_for_component"] = float(state.get("cycle_sec", 0.0) or 0.0)

    same_component = prev_key == key

    cycle_now = float(state.get("cycle_sec", 0.0) or 0.0)
    cycle_prev = float(state.get("_last_cycle_sec_for_component", cycle_now) or 0.0)

    cycle_delta = 0.0
    if same_component:
        cycle_delta = cycle_now - cycle_prev
        if cycle_delta < 0 or cycle_delta > max(float(delta_t or 0) + 2.0, 10.0):
            cycle_delta = 0.0

    state["_last_cycle_sec_for_component"] = cycle_now
    state["_last_component_key"] = key

    actual_delta = cycle_delta if cycle_delta > 0 else (float(delta_t or 0.0) if actual_cutting else 0.0)

    run_label = "CUTTING" if actual_cutting else "LOADED / IDLE"

    program_job_note = ""
    if parsed.get("program_job_num") and parsed.get("effective_job_num"):
        pjob = normalize_job_num(parsed.get("program_job_num") or "")
        ejob = normalize_job_num(parsed.get("effective_job_num") or "")
        if pjob and ejob and pjob != ejob:
            program_job_note = (
                f"; controller_program_job=J{pjob}; "
                f"effective_shop_job=J{ejob}"
            )

    note = (
        f"{run_label} on {machine}; "
        f"program={prog_raw}; "
        f"source={state.get('data_source', '')}; "
        f"machine_uptime={state.get('uptime')}; "
        f"cycle={state.get('cycle')}"
        f"{program_job_note}"
    )

    conn.execute("""
        UPDATE job_components
        SET
            status=CASE
                WHEN status='COMPLETE' AND COALESCE(manual_complete, 0)=1 THEN status
                ELSE 'RUNNING'
            END,
            started_at=CASE
                WHEN started_at IS NULL THEN datetime('now')
                ELSE started_at
            END,
            last_seen_machine=?,
            last_seen_program=?,
            actual_sec=COALESCE(actual_sec, 0) + ?,
            last_runtime_update=datetime('now'),
            last_status_note=?
        WHERE job_num=?
          AND component_name=?
    """, (
        machine,
        prog_raw or "",
        actual_delta,
        note,
        job_num,
        comp,
    ))

    if prev_key != key:
        log_job_status(conn, job_num, comp, "RUNNING", machine, prog_raw, note)


# ============================================================
# EFFICIENCY UPDATE
# ============================================================
def upsert_efficiency(conn, machine, delta_cut, delta_idle, delta_alarm, load_sum, load_n):
    now = datetime.now()

    for period, key in [
        ("day", now.strftime("%Y-%m-%d")),
        ("month", now.strftime("%Y-%m")),
        ("year", now.strftime("%Y")),
    ]:
        avg = load_sum / max(load_n, 1)

        conn.execute("""
            INSERT INTO efficiency_log (
                machine, period, period_key, cutting_sec, idle_sec, alarm_sec, load_samples, avg_spindle_load
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(machine, period, period_key) DO UPDATE SET
                cutting_sec=cutting_sec+excluded.cutting_sec,
                idle_sec=idle_sec+excluded.idle_sec,
                alarm_sec=alarm_sec+excluded.alarm_sec,
                load_samples=load_samples+excluded.load_samples,
                avg_spindle_load=CASE
                    WHEN load_samples+excluded.load_samples>0
                    THEN ((avg_spindle_load*load_samples)+(excluded.avg_spindle_load*excluded.load_samples))/(load_samples+excluded.load_samples)
                    ELSE avg_spindle_load
                END
        """, (
            machine,
            period,
            key,
            delta_cut,
            delta_idle,
            delta_alarm,
            load_n,
            avg,
        ))


# ============================================================
# TOOL LIFE CYCLE / REPLACEMENT HELPERS
# ============================================================
def get_next_tool_cycle_no(conn, machine: str, tid: str) -> int:
    tid = normalize_tool_num(tid)
    row = conn.execute("""
        SELECT COALESCE(MAX(cycle_no), 0)
        FROM tool_life_cycles
        WHERE machine=? AND tid=?
    """, (machine, str(tid))).fetchone()
    return int((row[0] if row else 0) or 0) + 1


def ensure_active_tool_cycle(conn, machine: str, tid: str, tool_name: str) -> int:
    tid = normalize_tool_num(tid)
    if not tid or tid == "0":
        return 0

    th = conn.execute("""
        SELECT *
        FROM tool_health
        WHERE machine=? AND tid=?
    """, (machine, tid)).fetchone()

    if th:
        th = dict(th)
        active_id = int(th.get("active_cycle_id") or 0)
        if active_id > 0:
            cyc = conn.execute("""
                SELECT id
                FROM tool_life_cycles
                WHERE id=? AND status='ACTIVE'
            """, (active_id,)).fetchone()
            if cyc:
                return active_id

    cycle_no = get_next_tool_cycle_no(conn, machine, tid)

    cur = conn.execute("""
        INSERT INTO tool_life_cycles (
            machine, tid, tool_name, cycle_no,
            started_at, status
        )
        VALUES (?, ?, ?, ?, datetime('now'), 'ACTIVE')
    """, (
        machine,
        tid,
        tool_name or get_tool_name(tid),
        cycle_no,
    ))

    cycle_id = cur.lastrowid

    conn.execute("""
        INSERT INTO tool_health (
            machine, tid, tool_name, active_cycle_id, cycle_started_at, last_seen
        )
        VALUES (?, ?, ?, ?, datetime('now'), datetime('now'))
        ON CONFLICT(machine, tid) DO UPDATE SET
            active_cycle_id=excluded.active_cycle_id,
            cycle_started_at=datetime('now'),
            tool_name=excluded.tool_name,
            last_seen=datetime('now')
    """, (
        machine,
        tid,
        tool_name or get_tool_name(tid),
        cycle_id,
    ))

    return cycle_id


def update_tool_cycle_aggregate(
    conn,
    cycle_id: int,
    cut_delta_sec: float,
    spindle_load: float,
    spindle_rpm: float,
    feed_rate: float,
    wear_score: float,
):
    if not cycle_id:
        return

    row = conn.execute("""
        SELECT *
        FROM tool_life_cycles
        WHERE id=?
    """, (cycle_id,)).fetchone()

    if not row:
        return

    r = dict(row)
    n = int(r.get("sample_count") or 0)
    new_n = n + 1

    old_avg_load = float(r.get("avg_load") or 0)
    old_avg_rpm = float(r.get("avg_rpm") or 0)
    old_avg_feed = float(r.get("avg_feed") or 0)

    new_avg_load = ((old_avg_load * n) + float(spindle_load or 0)) / max(new_n, 1)
    new_avg_rpm = ((old_avg_rpm * n) + float(spindle_rpm or 0)) / max(new_n, 1)
    new_avg_feed = ((old_avg_feed * n) + float(feed_rate or 0)) / max(new_n, 1)

    conn.execute("""
        UPDATE tool_life_cycles
        SET
            total_cut_sec=COALESCE(total_cut_sec,0)+?,
            avg_load=?,
            peak_load=MAX(COALESCE(peak_load,0), ?),
            avg_rpm=?,
            peak_rpm=MAX(COALESCE(peak_rpm,0), ?),
            avg_feed=?,
            sample_count=?,
            final_wear_score=?
        WHERE id=?
    """, (
        float(cut_delta_sec or 0),
        new_avg_load,
        float(spindle_load or 0),
        new_avg_rpm,
        float(spindle_rpm or 0),
        new_avg_feed,
        new_n,
        float(wear_score or 0),
        cycle_id,
    ))


def record_tool_cut_sample(
    conn,
    machine: str,
    tid: str,
    cycle_id: int,
    tool_name: str,
    prog_raw: str,
    parsed: dict,
    component_name: str,
    spindle_load: float,
    spindle_rpm: float,
    feed_rate: float,
    cut_delta_sec: float,
    in_cut_total_sec: float,
    wear_score: float,
    l_wear: float,
    d_wear: float,
    actual_cutting: bool,
):
    tid = normalize_tool_num(tid)
    if not tid or tid == "0":
        return

    sample_job = normalize_job_num((parsed or {}).get("job_num") or "")
    if job_num_is_all_zeros(sample_job):
        sample_job = ""

    conn.execute("""
        INSERT INTO tool_cut_samples (
            machine, tid, tool_cycle_id, tool_name,
            prog_raw, job_num, component_name,
            spindle_load, spindle_rpm, feed_rate,
            cut_delta_sec, in_cut_total_sec,
            wear_score, l_wear, d_wear, actual_cutting
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        machine,
        tid,
        int(cycle_id or 0),
        tool_name or get_tool_name(tid),
        prog_raw or "",
        sample_job,
        component_name or "",
        float(spindle_load or 0),
        float(spindle_rpm or 0),
        float(feed_rate or 0),
        float(cut_delta_sec or 0),
        float(in_cut_total_sec or 0),
        float(wear_score or 0),
        float(l_wear or 0),
        float(d_wear or 0),
        1 if actual_cutting else 0,
    ))

    if actual_cutting:
        update_tool_cycle_aggregate(
            conn,
            cycle_id,
            cut_delta_sec,
            spindle_load,
            spindle_rpm,
            feed_rate,
            wear_score,
        )

    conn.execute("""
        DELETE FROM tool_cut_samples
        WHERE id NOT IN (
            SELECT id FROM tool_cut_samples
            ORDER BY id DESC
            LIMIT 120000
        )
    """)


def summarize_previous_tool_cycles(conn, machine: str, tid: str, limit: int = 5):
    tid = normalize_tool_num(tid)

    rows = _rowdicts(conn.execute("""
        SELECT *
        FROM tool_life_cycles
        WHERE machine=?
          AND tid=?
          AND status!='ACTIVE'
        ORDER BY ended_at DESC, id DESC
        LIMIT ?
    """, (
        machine,
        str(tid),
        int(limit),
    )))

    if not rows:
        return {
            "count": 0,
            "avg_cut_sec": 0,
            "avg_load": 0,
            "avg_rpm": 0,
            "avg_wear": 0,
            "message": "No previous completed cycles for comparison.",
        }

    avg_cut = sum(float(r.get("total_cut_sec") or 0) for r in rows) / len(rows)
    avg_load = sum(float(r.get("avg_load") or 0) for r in rows) / len(rows)
    avg_rpm = sum(float(r.get("avg_rpm") or 0) for r in rows) / len(rows)
    avg_wear = sum(float(r.get("final_wear_score") or 0) for r in rows) / len(rows)

    return {
        "count": len(rows),
        "avg_cut_sec": round(avg_cut, 2),
        "avg_load": round(avg_load, 2),
        "avg_rpm": round(avg_rpm, 2),
        "avg_wear": round(avg_wear, 2),
        "message": f"Compared against last {len(rows)} completed tool cycle(s).",
    }


def archive_and_start_new_tool_cycle(
    conn,
    machine: str,
    tid: str,
    reason: str = "Manual replacement",
    notes: str = "",
    source: str = "MANUAL",
    candidate_id: int = 0,
):
    tid = normalize_tool_num(tid)
    if not tid or tid == "0":
        raise HTTPException(400, "Bad tool number")

    tool_name = get_tool_name(tid)
    th = conn.execute("""
        SELECT *
        FROM tool_health
        WHERE machine=? AND tid=?
    """, (machine, tid)).fetchone()

    old_cycle_id = 0
    if th:
        th = dict(th)
        tool_name = th.get("tool_name") or tool_name
        old_cycle_id = int(th.get("active_cycle_id") or 0)

    if old_cycle_id:
        conn.execute("""
            UPDATE tool_life_cycles
            SET
                ended_at=datetime('now'),
                replacement_reason=?,
                replacement_notes=?,
                replacement_source=?,
                archived_by_event_id=?,
                status='REPLACED'
            WHERE id=?
              AND status='ACTIVE'
        """, (
            reason or "",
            notes or "",
            source or "",
            int(candidate_id or 0),
            old_cycle_id,
        ))

    new_cycle_no = get_next_tool_cycle_no(conn, machine, tid)
    cur = conn.execute("""
        INSERT INTO tool_life_cycles (
            machine, tid, tool_name, cycle_no,
            started_at, status
        )
        VALUES (?, ?, ?, ?, datetime('now'), 'ACTIVE')
    """, (
        machine,
        tid,
        tool_name,
        new_cycle_no,
    ))

    new_cycle_id = cur.lastrowid

    cur2 = conn.execute("""
        INSERT INTO tool_replacement_events (
            candidate_id, machine, tid, tool_name,
            reason, notes, source,
            old_cycle_id, new_cycle_id
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        int(candidate_id or 0),
        machine,
        tid,
        tool_name,
        reason or "",
        notes or "",
        source or "",
        old_cycle_id,
        new_cycle_id,
    ))

    event_id = cur2.lastrowid

    conn.execute("""
        INSERT INTO tool_health (
            machine, tid, tool_name,
            in_cut_sec, max_life_min,
            peak_load, peak_rpm, wear_score,
            active_cycle_id, cycle_started_at,
            replacement_pending, last_seen
        )
        VALUES (?, ?, ?, 0, ?, 0, 0, 0, ?, datetime('now'), 0, datetime('now'))
        ON CONFLICT(machine, tid) DO UPDATE SET
            tool_name=excluded.tool_name,
            in_cut_sec=0,
            peak_load=0,
            peak_rpm=0,
            wear_score=0,
            active_cycle_id=excluded.active_cycle_id,
            cycle_started_at=datetime('now'),
            replacement_pending=0,
            last_seen=datetime('now')
    """, (
        machine,
        tid,
        tool_name,
        compute_tool_life_limit_min(conn, tid, tool_name),
        new_cycle_id,
    ))

    if candidate_id:
        conn.execute("""
            UPDATE tool_replacement_candidates
            SET
                status='CONFIRMED',
                confirmed_at=datetime('now'),
                notes=CASE WHEN ?!='' THEN ? ELSE notes END
            WHERE id=?
        """, (
            notes or "",
            notes or "",
            int(candidate_id),
        ))

    return {
        "event_id": event_id,
        "old_cycle_id": old_cycle_id,
        "new_cycle_id": new_cycle_id,
    }


def ignore_tool_replacement_candidate(conn, candidate_id: int, notes: str = ""):
    row = conn.execute("""
        SELECT *
        FROM tool_replacement_candidates
        WHERE id=?
    """, (int(candidate_id),)).fetchone()

    if not row:
        raise HTTPException(404, "Candidate not found")

    r = dict(row)

    conn.execute("""
        UPDATE tool_replacement_candidates
        SET
            status='IGNORED',
            ignored_at=datetime('now'),
            notes=CASE WHEN ?!='' THEN ? ELSE notes END
        WHERE id=?
    """, (
        notes or "",
        notes or "",
        int(candidate_id),
    ))

    conn.execute("""
        UPDATE tool_health
        SET replacement_pending=0
        WHERE machine=? AND tid=?
    """, (
        r.get("machine"),
        r.get("tid"),
    ))

    return {"status": "ignored"}


# ============================================================
# TOOL INSPECTION / CALIBRATION HELPERS
# ============================================================
def inspection_condition_to_target_wear(condition: str, observed_wear_pct: float) -> float:
    condition = normalize_component(condition or "GOOD")
    observed = _clamp(float(observed_wear_pct or 0), 0, 100)

    if condition == "GOOD":
        return _clamp(observed if observed > 0 else 15.0, 0.0, 35.0)

    if condition == "OK":
        return _clamp(observed if observed > 0 else 35.0, 20.0, 55.0)

    if condition == "WATCH":
        return _clamp(observed if observed > 0 else 65.0, 45.0, 80.0)

    if condition == "WORN":
        return _clamp(observed if observed > 0 else 85.0, 70.0, 95.0)

    if condition == "CHIPPED":
        return _clamp(observed if observed > 0 else 95.0, 85.0, 100.0)

    if condition == "BROKEN":
        return 100.0

    if condition == "REPLACED":
        return 0.0

    return _clamp(observed, 0, 100)


def record_tool_inspection(conn, req: ToolInspectionRequest):
    machine = req.machine
    tid = normalize_tool_num(req.tid)

    if not machine or not tid or tid == "0":
        raise HTTPException(400, "machine and tid required")

    condition = normalize_component(req.condition or "GOOD")

    allowed = {
        "GOOD",
        "OK",
        "WATCH",
        "WORN",
        "CHIPPED",
        "BROKEN",
        "REPLACED",
    }

    if condition not in allowed:
        raise HTTPException(400, "Bad condition")

    th = conn.execute("""
        SELECT *
        FROM tool_health
        WHERE machine=? AND tid=?
    """, (
        machine,
        tid,
    )).fetchone()

    if not th:
        raise HTTPException(404, "Tool health not found")

    th = dict(th)

    tool_name = th.get("tool_name") or get_tool_name(tid)
    old_wear = float(th.get("wear_score") or 0)
    old_life = float(th.get("max_life_min") or compute_tool_life_limit_min(conn, tid, tool_name))
    in_cut_sec = float(th.get("in_cut_sec") or 0)
    active_cycle_id = int(th.get("active_cycle_id") or 0)

    observed = _clamp(float(req.observed_wear_pct or 0), 0, 100)

    new_wear = old_wear
    new_life = old_life

    calibrate = int(req.calibrate_health or 0)

    target_wear = inspection_condition_to_target_wear(condition, observed)

    if calibrate and condition in ("GOOD", "OK", "WATCH"):
        if in_cut_sec > 60 and target_wear > 0:
            estimated_life_from_inspection = (in_cut_sec / (target_wear / 100.0)) / 60.0

            new_life = _clamp(
                max(old_life, estimated_life_from_inspection),
                5.0,
                max(old_life * 4.0, 60.0),
            )

            new_wear = round(
                _clamp(
                    (in_cut_sec / max(new_life * 60.0, 1.0)) * 100.0,
                    0,
                    100,
                ),
                2,
            )

            if condition in ("GOOD", "OK"):
                new_wear = min(new_wear, target_wear)
        else:
            new_wear = target_wear

        conn.execute("""
            UPDATE tool_catalog
            SET
                max_life_min=?,
                updated_at=datetime('now')
            WHERE tid=?
        """, (
            float(new_life),
            tid,
        ))

        conn.execute("""
            UPDATE tool_health
            SET
                max_life_min=?,
                wear_score=?,
                replacement_pending=0,
                last_seen=datetime('now')
            WHERE machine=? AND tid=?
        """, (
            float(new_life),
            float(new_wear),
            machine,
            tid,
        ))

        conn.execute("""
            UPDATE tool_life_cycles
            SET final_wear_score=?
            WHERE id=?
        """, (
            float(new_wear),
            active_cycle_id,
        ))

    elif calibrate and condition in ("WORN", "CHIPPED", "BROKEN"):
        new_wear = max(old_wear, target_wear)

        conn.execute("""
            UPDATE tool_health
            SET
                wear_score=?,
                replacement_pending=CASE
                    WHEN ? IN ('CHIPPED','BROKEN') THEN 1
                    ELSE replacement_pending
                END,
                last_seen=datetime('now')
            WHERE machine=? AND tid=?
        """, (
            float(new_wear),
            condition,
            machine,
            tid,
        ))

        conn.execute("""
            UPDATE tool_life_cycles
            SET final_wear_score=?
            WHERE id=?
        """, (
            float(new_wear),
            active_cycle_id,
        ))

    elif calibrate and condition == "REPLACED":
        new_wear = 0.0

        conn.execute("""
            UPDATE tool_health
            SET
                wear_score=0,
                replacement_pending=0,
                last_seen=datetime('now')
            WHERE machine=? AND tid=?
        """, (
            machine,
            tid,
        ))

        conn.execute("""
            UPDATE tool_life_cycles
            SET final_wear_score=0
            WHERE id=?
        """, (
            active_cycle_id,
        ))

    cur = conn.execute("""
        INSERT INTO tool_inspections (
            machine, tid, tool_name,
            condition, inspected_by,
            observed_wear_pct,
            measured_diameter,
            measured_length_offset,
            edge_notes,
            action_taken,
            calibrate_health,
            old_wear_score,
            new_wear_score,
            old_max_life_min,
            new_max_life_min,
            active_cycle_id,
            notes
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        machine,
        tid,
        tool_name,
        condition,
        req.inspected_by or "",
        float(observed),
        float(req.measured_diameter or 0),
        float(req.measured_length_offset or 0),
        req.edge_notes or "",
        req.action_taken or "INSPECTED",
        1 if calibrate else 0,
        float(old_wear),
        float(new_wear),
        float(old_life),
        float(new_life),
        int(active_cycle_id or 0),
        req.notes or "",
    ))

    return {
        "inspection_id": cur.lastrowid,
        "machine": machine,
        "tid": tid,
        "condition": condition,
        "old_wear_score": round(old_wear, 2),
        "new_wear_score": round(new_wear, 2),
        "old_max_life_min": round(old_life, 2),
        "new_max_life_min": round(new_life, 2),
        "calibrated": bool(calibrate),
    }


# ============================================================
# SAFER OFFSET CHANGE DETECTION
# ============================================================
def detect_tool_replacement_candidate(
    conn,
    state: dict,
    machine: str,
    tid: str,
    tool_name: str,
    l_wear: float,
    d_wear: float,
    actual_cutting: bool,
    mode: str,
    prog_raw: str,
):
    """
    Conservative offset-change detection.

    This does NOT say a tool was definitely replaced.
    This does NOT reset tool life.
    This only creates a pending admin warning after a stable offset change.
    """
    if not ENABLE_AUTO_TOOL_REPLACEMENT_CANDIDATES:
        return None

    tid = normalize_tool_num(tid)
    if not tid or tid == "0":
        return None

    now = time.time()

    if actual_cutting:
        return None

    if not machine_idle_long_enough_for_offset_detection(state):
        return None

    mode_u = (mode or "").upper()
    if (
        "OFFLINE" in mode_u
        or "ALARM" in mode_u
        or "E-STOP" in mode_u
        or "ESTOP" in mode_u
        or "FAULT" in mode_u
    ):
        return None

    if not offset_read_looks_valid(l_wear, d_wear):
        return None

    new_l = float(l_wear or 0)
    new_d = float(d_wear or 0)
    baseline_key = f"{machine}|{tid}"

    if baseline_key not in state.get("_offset_baseline_ready", set()):
        state.setdefault("_offset_baseline_ready", set()).add(baseline_key)

        conn.execute("""
            INSERT INTO tool_health (
                machine, tid, tool_name,
                last_l_wear, last_d_wear,
                last_seen
            )
            VALUES (?, ?, ?, ?, ?, datetime('now'))
            ON CONFLICT(machine, tid) DO UPDATE SET
                tool_name=excluded.tool_name,
                last_l_wear=excluded.last_l_wear,
                last_d_wear=excluded.last_d_wear,
                last_seen=datetime('now')
        """, (
            machine,
            tid,
            tool_name or get_tool_name(tid),
            new_l,
            new_d,
        ))

        state.setdefault("_offset_candidate_streaks", {})
        state["_offset_candidate_streaks"].pop(baseline_key, None)
        return None

    th = conn.execute("""
        SELECT *
        FROM tool_health
        WHERE machine=? AND tid=?
    """, (machine, tid)).fetchone()

    if not th:
        conn.execute("""
            INSERT INTO tool_health (
                machine, tid, tool_name,
                last_l_wear, last_d_wear,
                last_seen
            )
            VALUES (?, ?, ?, ?, ?, datetime('now'))
        """, (
            machine,
            tid,
            tool_name or get_tool_name(tid),
            new_l,
            new_d,
        ))
        return None

    th = dict(th)

    old_l = float(th.get("last_l_wear") or 0)
    old_d = float(th.get("last_d_wear") or 0)

    if not offset_read_looks_valid(old_l, old_d):
        conn.execute("""
            UPDATE tool_health
            SET
                last_l_wear=?,
                last_d_wear=?,
                last_seen=datetime('now')
            WHERE machine=? AND tid=?
        """, (
            new_l,
            new_d,
            machine,
            tid,
        ))
        return None

    delta_l = new_l - old_l
    delta_d = new_d - old_d

    changed = (
        abs(delta_l) >= WIPS_LENGTH_CHANGE_THRESHOLD
        or abs(delta_d) >= WIPS_DIAMETER_CHANGE_THRESHOLD
    )

    state.setdefault("_offset_candidate_streaks", {})
    streaks = state["_offset_candidate_streaks"]

    if not changed:
        streaks.pop(baseline_key, None)

        conn.execute("""
            UPDATE tool_health
            SET
                last_l_wear=?,
                last_d_wear=?,
                last_seen=datetime('now')
            WHERE machine=? AND tid=?
        """, (
            new_l,
            new_d,
            machine,
            tid,
        ))

        return None

    streak = streaks.get(baseline_key)

    if not streak:
        streaks[baseline_key] = {
            "count": 1,
            "new_l": new_l,
            "new_d": new_d,
            "old_l": old_l,
            "old_d": old_d,
            "first_seen": now,
        }
        return None

    same_candidate = (
        abs(float(streak.get("new_l", 0)) - new_l) <= 0.001
        and abs(float(streak.get("new_d", 0)) - new_d) <= 0.001
    )

    if not same_candidate:
        streaks[baseline_key] = {
            "count": 1,
            "new_l": new_l,
            "new_d": new_d,
            "old_l": old_l,
            "old_d": old_d,
            "first_seen": now,
        }
        return None

    streak["count"] = int(streak.get("count") or 0) + 1

    if streak["count"] < WIPS_STABLE_CONFIRM_SAMPLES:
        return None

    last_candidate_at = float(th.get("last_probe_candidate_at") or 0)
    replacement_pending = int(th.get("replacement_pending") or 0)

    if replacement_pending:
        conn.execute("""
            UPDATE tool_health
            SET
                last_l_wear=?,
                last_d_wear=?,
                last_seen=datetime('now')
            WHERE machine=? AND tid=?
        """, (
            new_l,
            new_d,
            machine,
            tid,
        ))

        streaks.pop(baseline_key, None)
        return None

    if now - last_candidate_at < WIPS_CANDIDATE_COOLDOWN_SEC:
        conn.execute("""
            UPDATE tool_health
            SET
                last_l_wear=?,
                last_d_wear=?,
                last_seen=datetime('now')
            WHERE machine=? AND tid=?
        """, (
            new_l,
            new_d,
            machine,
            tid,
        ))

        streaks.pop(baseline_key, None)
        return None

    pending = conn.execute("""
        SELECT id
        FROM tool_replacement_candidates
        WHERE machine=?
          AND tid=?
          AND status='PENDING'
        ORDER BY id DESC
        LIMIT 1
    """, (
        machine,
        tid,
    )).fetchone()

    if pending:
        conn.execute("""
            UPDATE tool_health
            SET
                last_l_wear=?,
                last_d_wear=?,
                replacement_pending=1,
                last_seen=datetime('now')
            WHERE machine=? AND tid=?
        """, (
            new_l,
            new_d,
            machine,
            tid,
        ))

        streaks.pop(baseline_key, None)
        return dict(pending)

    reason = (
        "Stable tool offset change detected. "
        "This does NOT prove the tool was replaced. "
        "Confirm only if the operator actually replaced or reloaded the tool. "
        f"L delta={delta_l:.4f}, D delta={delta_d:.4f}."
    )

    cur = conn.execute("""
        INSERT INTO tool_replacement_candidates (
            machine, tid, tool_name,
            old_l_wear, new_l_wear, delta_l_wear,
            old_d_wear, new_d_wear, delta_d_wear,
            status, detected_reason, machine_status, prog_raw
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'PENDING', ?, ?, ?)
    """, (
        machine,
        tid,
        tool_name or get_tool_name(tid),
        old_l,
        new_l,
        delta_l,
        old_d,
        new_d,
        delta_d,
        reason,
        mode or "",
        prog_raw or "",
    ))

    candidate_id = cur.lastrowid

    conn.execute("""
        UPDATE tool_health
        SET
            last_l_wear=?,
            last_d_wear=?,
            replacement_pending=1,
            last_probe_candidate_at=?,
            last_seen=datetime('now')
        WHERE machine=? AND tid=?
    """, (
        new_l,
        new_d,
        now,
        machine,
        tid,
    ))

    streaks.pop(baseline_key, None)

    return {
        "id": candidate_id,
        "machine": machine,
        "tid": tid,
        "reason": reason,
    }


# ============================================================
# DASHBOARD DATA HELPERS
# ============================================================
def efficiency_rows_for_period(conn, period: str, key: str):
    return _rowdicts(conn.execute(
        "SELECT * FROM efficiency_log WHERE period=? AND period_key=?",
        (period, key),
    ))


def merge_efficiency_rows(rows, period):
    result = {r["machine"]: dict(r) for r in rows}

    for mn, st in machine_states.items():
        if mn not in result:
            result[mn] = {
                "machine": mn,
                "period": period,
                "cutting_sec": 0,
                "idle_sec": 0,
                "alarm_sec": 0,
                "avg_spindle_load": 0,
                "load_samples": 0,
                "parts_run": 0,
            }

        r = result[mn]
        r["cutting_sec"] = r.get("cutting_sec", 0) + st["_cut_sec"]
        r["idle_sec"] = r.get("idle_sec", 0) + st["_idle_sec"]
        r["alarm_sec"] = r.get("alarm_sec", 0) + st["_alarm_sec"]

    return list(result.values())


def safe_machine_public_state(tool_catalog):
    catalog_map = {str(r["tid"]): r["tool_name"] for r in tool_catalog}
    machines_out = {}

    for name, st in machine_states.items():
        mo = {k: v for k, v in st.items() if not k.startswith("_")}
        tid = normalize_tool_num(mo.get("tool", "0"))
        mo["tool"] = tid
        mo["tool_name"] = catalog_map.get(tid, get_tool_name(tid))
        machines_out[name] = mo

    return machines_out


# ============================================================
# MACHINE POLLING
# ============================================================
def poll_single_machine(m_name: str, ip: str, port: int):
    last_time = time.time()
    FLUSH_INTERVAL = 60.0
    OFFLINE_THRESHOLD = 3

    cfg = next((c for c in MACHINE_CONFIGS if c["name"] == m_name), {})
    use_qcode = cfg.get("use_qcode", True)
    use_mtconnect = cfg.get("use_mtconnect", True)

    while True:
        now = time.time()
        delta_t = min(now - last_time, 5.0)
        last_time = now
        state = machine_states[m_name]

        q_snapshot = {"ok": False}
        mt_snapshot = {"ok": False}

        if use_qcode:
            q_snapshot = read_qcode_snapshot(ip, port, state)

        if use_mtconnect:
            q_prog = q_snapshot.get("prog", "") if q_snapshot.get("ok") else ""
            mt_snapshot = read_mtconnect_snapshot(cfg, state, q_prog)

        snap = merge_machine_snapshot(mt_snapshot, q_snapshot)

        if not snap.get("ok"):
            state["_consec_offline"] = state.get("_consec_offline", 0) + 1
            if state["_consec_offline"] >= OFFLINE_THRESHOLD:
                state["status"] = "OFFLINE"
                state["data_source"] = "OFFLINE"
                state["mtconnect_ok"] = False
                state["mtconnect_error"] = mt_snapshot.get("mtconnect_error", "")
            time.sleep(2)
            continue

        state["_consec_offline"] = 0

        mode = str(snap.get("status") or "AVAILABLE").strip().upper()
        state["status"] = mode
        state["raw_mode"] = str(snap.get("raw_mode") or mode).strip().upper()
        state["data_source"] = snap.get("source", "")
        state["mtconnect_ok"] = bool(mt_snapshot.get("ok"))
        state["mtconnect_url"] = mt_snapshot.get("mtconnect_url") or cfg.get("mtconnect_url", "")
        state["mtconnect_error"] = "" if mt_snapshot.get("ok") else mt_snapshot.get("mtconnect_error", "")

        mode_cutting = _is_active(mode)
        is_alarm = _is_alarm_mode(mode)

        prog_raw = snap.get("prog") or "N/A"

        if program_looks_like_machine_name(prog_raw, m_name):
            prog_raw = "N/A"

        parsed = parse_program_name(prog_raw)

        if parsed.get("job_num") and job_num_is_all_zeros(parsed.get("job_num")):
            parsed = _idle_program_parse_result("IDLE")

        state["prog"] = prog_raw
        state["job_num"] = parsed["job_num"]
        state["part_name"] = parsed["part_name"]
        state["sub_cat"] = parsed["sub_cat"]
        state["prog_display"] = parsed["display"]

        state["uptime"] = snap.get("uptime") or "00:00:00"
        state["cycle"] = snap.get("cycle") or "00:00:00"
        state["cycle_sec"] = float(snap.get("cycle_sec") or 0.0)

        tool = apply_tool_hold_last_good(snap.get("tool") or "0", state)
        state["tool"] = tool

        spindle_load = float(snap.get("load") or 0.0)
        path_feed = float(snap.get("feed_rate") or 0.0)
        spindle_rpm = float(snap.get("spindle_rpm") or 0.0)
        rpm_signal_ok = bool(snap.get("rpm_signal_ok"))
        rpm_source = snap.get("rpm_source", "")

        part_count = int(float(snap.get("part_count") or 0))
        state["part_count"] = part_count

        if spindle_rpm > 0:
            state["_last_nonzero_rpm"] = spindle_rpm
            state["_last_nonzero_rpm_at"] = now
            state["_no_spindle_since"] = None
        else:
            if state.get("_no_spindle_since") is None:
                state["_no_spindle_since"] = now

        last_rpm_at = float(state.get("_last_nonzero_rpm_at", 0.0) or 0.0)
        rpm_zero_for_sec = now - last_rpm_at if last_rpm_at > 0 else now - float(state.get("_no_spindle_since") or now)
        rpm_idle_eligible = rpm_zero_for_sec >= IDLE_AFTER_NO_RPM_SEC

        load_says_cutting = spindle_load >= SPINDLE_LOAD_CUTTING_THRESHOLD
        feed_says_cutting = path_feed >= FEED_CUTTING_THRESHOLD

        actual_cutting = (
            spindle_rpm > 0
            or (mode_cutting and load_says_cutting)
            or (mode_cutting and feed_says_cutting and spindle_load > 0.5)
        )

        state["actual_cutting"] = actual_cutting
        state["rpm_signal_ok"] = bool(rpm_signal_ok)
        state["rpm_source"] = rpm_source

        spindle_rpm_for_runtime = spindle_rpm
        if actual_cutting and spindle_rpm_for_runtime <= 0:
            spindle_rpm_for_runtime = float(state.get("_last_nonzero_rpm", 0.0) or 0.0)
            if spindle_rpm_for_runtime <= 0:
                spindle_rpm_for_runtime = RPM_FALLBACK_FOR_WEAR

        l_wear = d_wear = 0.0
        tool_name = get_tool_name(tool)

        if use_qcode and tool.isdigit() and int(tool) > 0:
            t_id = int(tool)
            l_wear = query_haas_float(ip, port, f"?Q600 {L_OFFSET_BASE + t_id}")
            d_wear = query_haas_float(ip, port, f"?Q600 {D_OFFSET_BASE + t_id}")

        state.update({
            "load": spindle_load,
            "spindle_rpm": spindle_rpm,
            "feed_rate": path_feed,
            "l_wear": l_wear,
            "d_wear": d_wear,
            "tool_name": tool_name,
        })

        if actual_cutting:
            state["_cut_sec"] += delta_t
        elif is_alarm:
            state["_alarm_sec"] += delta_t
        else:
            if rpm_idle_eligible:
                state["_idle_sec"] += delta_t

        if spindle_load > 0:
            state["_load_sum"] += spindle_load
            state["_load_n"] += 1

        with db_lock, sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row

            # Resolve reused/old program job numbers to the current matched shop job.
            # Example: controller program J8206 may actually be current job J8342.
            if parsed.get("job_num"):
                parsed = resolve_effective_job_from_matching(
                    conn=conn,
                    state=state,
                    parsed=parsed,
                    prog_raw=prog_raw,
                    machine=m_name,
                )

                state["job_num"] = parsed.get("job_num")
                state["part_name"] = parsed.get("part_name")
                state["sub_cat"] = parsed.get("sub_cat")
                state["prog_display"] = parsed.get("display")

            conn.execute("""
                INSERT INTO telemetry_samples (
                    machine, source, status, prog_raw, tool,
                    spindle_rpm, feed_rate, spindle_load, actual_cutting,
                    mtconnect_ok, mtconnect_error
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                m_name,
                state.get("data_source", ""),
                state.get("status", ""),
                prog_raw,
                tool,
                spindle_rpm,
                path_feed,
                spindle_load,
                1 if actual_cutting else 0,
                1 if state.get("mtconnect_ok") else 0,
                state.get("mtconnect_error", ""),
            ))

            conn.execute("""
                DELETE FROM telemetry_samples
                WHERE id NOT IN (
                    SELECT id FROM telemetry_samples
                    ORDER BY id DESC
                    LIMIT 50000
                )
            """)

            if parsed["job_num"]:
                update_job_component_from_program(
                    conn=conn,
                    state=state,
                    machine=m_name,
                    prog_raw=prog_raw,
                    parsed=parsed,
                    actual_cutting=actual_cutting,
                    delta_t=delta_t,
                )

            if parsed["job_num"] and prog_raw and prog_raw.upper() not in ("N/A", "IDLE"):
                analyze_and_store_gcode_program(conn, m_name, prog_raw)

                runtime_est = get_gcode_program_runtime_estimate_sec(
                    conn=conn,
                    machine=m_name,
                    prog_raw=prog_raw,
                )

                if not runtime_est.get("ok"):
                    current_comp_for_est = detected_component_for_program(parsed, prog_raw)
                    hist = similar_completed_component_estimate_sec(
                        conn,
                        job_num=parsed["job_num"],
                        component_name=current_comp_for_est,
                    )
                    if hist.get("ok"):
                        runtime_est = {
                            "ok": True,
                            "estimated_sec": float(hist.get("estimated_sec") or 0),
                            "source": hist.get("source") or "SIMILAR_COMPLETED_JOB",
                            "message": hist.get("basis") or "",
                        }

                progress = compute_live_program_progress(
                    state=state,
                    prog_raw=prog_raw,
                    estimated_sec=runtime_est.get("estimated_sec", 0.0),
                    cycle_sec=float(state.get("cycle_sec") or 0.0),
                    actual_cutting=actual_cutting,
                )

                if runtime_est.get("source") and progress.get("program_progress_source") != "NO_ESTIMATE":
                    progress["program_progress_source"] = (
                        runtime_est.get("source") + "_PLUS_"
                        + progress.get("program_progress_source", "")
                    ).strip("_")

                state.update(progress)

            else:
                state.update({
                    "program_est_sec": 0.0,
                    "program_progress_pct": 0.0,
                    "program_remaining_sec": 0.0,
                    "program_est_finish": "",
                    "program_progress_source": "",
                })

            current_component = detected_component_for_program(parsed, prog_raw) or ""

            if tool.isdigit() and int(tool) > 0:
                cat_info = get_tool_catalog_info(conn, tool, tool_name)
                category = cat_info.get("category") or infer_tool_category(tool_name)
                max_life_min = compute_tool_life_limit_min(conn, tool, tool_name)
                cut_delta = delta_t if actual_cutting else 0.0

                cycle_id = ensure_active_tool_cycle(conn, m_name, tool, tool_name)

                detect_tool_replacement_candidate(
                    conn=conn,
                    state=state,
                    machine=m_name,
                    tid=tool,
                    tool_name=tool_name,
                    l_wear=l_wear,
                    d_wear=d_wear,
                    actual_cutting=actual_cutting,
                    mode=mode,
                    prog_raw=prog_raw,
                )

                row = conn.execute(
                    "SELECT * FROM tool_health WHERE machine=? AND tid=?",
                    (m_name, tool),
                ).fetchone()

                prior_cut = float(row["in_cut_sec"] or 0) if row else 0.0

                gcode_severity = get_gcode_tool_severity(
                    conn=conn,
                    machine=m_name,
                    prog_raw=prog_raw,
                    tid=tool,
                )

                wear_score = compute_wear_score(
                    prior_cut_sec=prior_cut,
                    cut_delta_sec=cut_delta,
                    max_life_min=max_life_min,
                    spindle_load=spindle_load,
                    l_wear=l_wear,
                    d_wear=d_wear,
                    category=category,
                    gcode_severity=gcode_severity,
                )

                projected_cut = prior_cut + cut_delta

                conn.execute("""
                    INSERT INTO tool_health (
                        machine, tid, tool_name, in_cut_sec, max_life_min,
                        peak_load, peak_rpm, wear_score,
                        active_cycle_id, cycle_started_at,
                        last_l_wear, last_d_wear,
                        last_seen
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'), ?, ?, datetime('now'))
                    ON CONFLICT(machine, tid) DO UPDATE SET
                        tool_name=excluded.tool_name,
                        in_cut_sec=tool_health.in_cut_sec+excluded.in_cut_sec,
                        max_life_min=excluded.max_life_min,
                        peak_load=MAX(tool_health.peak_load, excluded.peak_load),
                        peak_rpm=MAX(tool_health.peak_rpm, excluded.peak_rpm),
                        wear_score=excluded.wear_score,
                        active_cycle_id=CASE
                            WHEN COALESCE(tool_health.active_cycle_id,0)=0 THEN excluded.active_cycle_id
                            ELSE tool_health.active_cycle_id
                        END,
                        last_l_wear=CASE
                            WHEN ABS(excluded.last_l_wear) < 0.000001
                             AND ABS(excluded.last_d_wear) < 0.000001
                            THEN tool_health.last_l_wear
                            ELSE excluded.last_l_wear
                        END,
                        last_d_wear=CASE
                            WHEN ABS(excluded.last_l_wear) < 0.000001
                             AND ABS(excluded.last_d_wear) < 0.000001
                            THEN tool_health.last_d_wear
                            ELSE excluded.last_d_wear
                        END,
                        last_seen=datetime('now')
                """, (
                    m_name,
                    tool,
                    tool_name,
                    cut_delta,
                    max_life_min,
                    spindle_load,
                    spindle_rpm_for_runtime,
                    wear_score,
                    cycle_id,
                    l_wear,
                    d_wear,
                ))

                row2 = conn.execute(
                    "SELECT * FROM tool_health WHERE machine=? AND tid=?",
                    (m_name, tool),
                ).fetchone()

                row2d = dict(row2) if row2 else {}
                active_cycle_id = int(row2d.get("active_cycle_id") or cycle_id or 0)
                total_cut_now = float(row2d.get("in_cut_sec") or projected_cut or 0)

                record_tool_cut_sample(
                    conn=conn,
                    machine=m_name,
                    tid=tool,
                    cycle_id=active_cycle_id,
                    tool_name=tool_name,
                    prog_raw=prog_raw,
                    parsed=parsed,
                    component_name=current_component,
                    spindle_load=spindle_load,
                    spindle_rpm=spindle_rpm_for_runtime,
                    feed_rate=path_feed,
                    cut_delta_sec=cut_delta,
                    in_cut_total_sec=total_cut_now,
                    wear_score=wear_score,
                    l_wear=l_wear,
                    d_wear=d_wear,
                    actual_cutting=actual_cutting,
                )

            cur_prog = prog_raw if parsed["job_num"] else None
            if cur_prog and cur_prog != state.get("_last_prog"):
                state["_last_prog"] = cur_prog

                if state["_current_job_id"]:
                    conn.execute(
                        "UPDATE job_runs SET status='COMPLETE', last_seen=datetime('now') WHERE id=?",
                        (state["_current_job_id"],),
                    )

                cur = conn.execute(
                    "INSERT INTO job_runs (machine, job_num, part_name, prog_raw) VALUES (?, ?, ?, ?)",
                    (m_name, parsed["job_num"], parsed["part_name"], prog_raw),
                )
                state["_current_job_id"] = cur.lastrowid

            if state["_current_job_id"] and spindle_load > 0:
                conn.execute("""
                    UPDATE job_runs SET
                        last_seen=datetime('now'),
                        cycle_time_sec=CAST(strftime('%s','now')-strftime('%s',started_at) AS REAL),
                        avg_load=CASE WHEN load_samples>0 THEN (avg_load*load_samples+?)/(load_samples+1) ELSE ? END,
                        max_load=MAX(max_load, ?),
                        avg_rpm=CASE WHEN load_samples>0 THEN (avg_rpm*load_samples+?)/(load_samples+1) ELSE ? END,
                        load_samples=load_samples+1,
                        status='RUNNING'
                    WHERE id=?
                """, (
                    spindle_load,
                    spindle_load,
                    spindle_load,
                    spindle_rpm_for_runtime,
                    spindle_rpm_for_runtime,
                    state["_current_job_id"],
                ))

            if now - state["_last_flush"] >= FLUSH_INTERVAL:
                upsert_efficiency(
                    conn,
                    m_name,
                    state["_cut_sec"],
                    state["_idle_sec"],
                    state["_alarm_sec"],
                    state["_load_sum"],
                    state["_load_n"],
                )

                state["_cut_sec"] = 0.0
                state["_idle_sec"] = 0.0
                state["_alarm_sec"] = 0.0
                state["_load_sum"] = 0.0
                state["_load_n"] = 0
                state["_last_flush"] = now

            conn.commit()

        time.sleep(1)


async def broadcast_loop():
    while True:
        await manager.broadcast(json.dumps({"type": "telemetry"}))
        await asyncio.sleep(1.5)


# ============================================================
# END PART 4
# PART 5 CONTINUES WITH:
# - Admin API
# - Diagnostic/network API
# - Dashboard API
# - Inventory/QR API
# - Tool APIs
# - G-code API
# - Steel sheet API
# - Scheduler/job/catalog/shop-board/autoloader APIs
# ============================================================
# ============================================================
# ADMIN API
# ============================================================
@app.post("/api/admin/login")
async def admin_login(req: AdminLoginRequest, response: Response):
    if not hmac.compare_digest(req.password or "", ADMIN_PASSWORD):
        raise HTTPException(401, "Bad admin password")

    token = make_admin_token()
    response.set_cookie(
        ADMIN_COOKIE_NAME,
        token,
        httponly=True,
        secure=bool(HTTPS_CERT_FILE and HTTPS_KEY_FILE),
        samesite="lax",
        max_age=ADMIN_SESSION_TTL_SEC,
    )
    return {"status": "ok", "admin": True}


@app.post("/api/admin/logout")
async def admin_logout(response: Response):
    response.delete_cookie(ADMIN_COOKIE_NAME)
    return {"status": "ok", "admin": False}


@app.get("/api/admin/me")
async def admin_me(request: Request):
    return {"admin": is_admin_request(request)}


# ============================================================
# DIAGNOSTIC / NETWORK API
# ============================================================
@app.get("/api/network/info")
async def api_network_info():
    urls = get_local_network_addresses(APP_PORT)

    return {
        "status": "ok",
        "app_name": APP_NAME,
        "short_name": APP_SHORT_NAME,
        "port": APP_PORT,
        "https_enabled": bool(
            HTTPS_CERT_FILE
            and HTTPS_KEY_FILE
            and os.path.exists(HTTPS_CERT_FILE)
            and os.path.exists(HTTPS_KEY_FILE)
        ),
        "urls": urls,
        "notes": [
            "Run this app on the main shop computer.",
            "Other computers do not need Wi-Fi if they are connected by Ethernet/private LAN to the host computer.",
            "Open the host computer's Ethernet/private-network IP address in a browser.",
            f"Windows Firewall must allow inbound TCP port {APP_PORT}.",
            "Live iPhone camera scanning requires HTTPS or ngrok. HTTP still works for normal app pages and photo scan.",
        ],
    }


@app.get("/api/mtconnect/test")
async def test_mtconnect(request: Request):
    """
    Admin-only because this exposes machine IPs, MTConnect adapter URLs,
    raw MTConnect keys, and troubleshooting details.
    """
    require_admin_request(request)

    results = []

    for cfg in MACHINE_CONFIGS:
        machine_result = {
            "machine": cfg.get("name", ""),
            "ip": cfg.get("ip", ""),
            "configured_url": cfg.get("mtconnect_url", ""),
            "tests": [],
        }

        for url in mtconnect_candidate_urls(cfg):
            mt = fetch_mtconnect_current(url)

            item = {
                "url": url,
                "ok": bool(mt.get("_ok")),
                "error": mt.get("_error", ""),
                "status": derive_status_from_mtconnect(mt) if mt.get("_ok") else "OFFLINE",
                "program": choose_program_from_sources(mt, "") if mt.get("_ok") else "",
                "tool": mt_get(mt, MT_TOOL_KEYS, "") if mt.get("_ok") else "",
                "rpm": mt_get_float(mt, MT_SPINDLE_RPM_KEYS, 0.0) if mt.get("_ok") else 0,
                "feed": mt_get_float(mt, MT_FEED_KEYS, 0.0) if mt.get("_ok") else 0,
                "load": mt_get_float(mt, MT_LOAD_KEYS, 0.0) if mt.get("_ok") else 0,
                "sample_keys": [
                    k for k in mt.keys()
                    if not k.startswith("_") and not k.startswith("raw_")
                ][:80] if mt.get("_ok") else [],
            }

            machine_result["tests"].append(item)

        results.append(machine_result)

    return {
        "status": "ok",
        "results": results,
    }


# ============================================================
# DASHBOARD API
# ============================================================
@app.get("/api/dashboard")
async def get_dashboard(request: Request):
    admin_ok = is_admin_request(request)

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        tool_catalog = _rowdicts(conn.execute("""
            SELECT *
            FROM tool_catalog
            ORDER BY CAST(tid AS INTEGER)
        """))

        tools = _rowdicts(conn.execute("""
            SELECT
                th.*,
                COALESCE(NULLIF(th.tool_name,''), tc.tool_name, 'Tool ' || th.tid) AS display_tool_name,
                tc.category AS tool_category,
                tc.diameter AS diameter,
                tc.flutes AS flutes,
                tc.work_material AS work_material,
                tc.max_life_min AS catalog_max_life_min
            FROM tool_health th
            LEFT JOIN tool_catalog tc ON th.tid=tc.tid
            ORDER BY th.machine, CAST(th.tid AS INTEGER)
        """))

        for t in tools:
            tid = normalize_tool_num(t.get("tid", "0"))
            t["tid"] = tid
            t["tool_name"] = t.get("display_tool_name") or t.get("tool_name") or get_tool_name(tid)
            score = float(t.get("wear_score") or 0)
            t["health_color"] = tool_health_color(score)
            t["life_pct"] = round(_clamp(score, 0, 100), 2)
            t["life_remaining_pct"] = round(_clamp(100 - score, 0, 100), 2)

        pending_tool_candidates = []
        if admin_ok:
            pending_tool_candidates = _rowdicts(conn.execute("""
                SELECT *
                FROM tool_replacement_candidates
                WHERE status='PENDING'
                ORDER BY timestamp DESC
            """))

        # Inventory query includes checkout timing.
        # This lets the UI show how long it takes between tool checkouts,
        # last checkout, 7-day usage, 30-day usage, and runway.
        inventory = _rowdicts(conn.execute("""
            SELECT
                i.*,

                COALESCE((
                    SELECT SUM(t.qty)
                    FROM transactions t
                    WHERE t.sku=i.sku
                      AND t.direction='OUT'
                      AND t.timestamp>=datetime('now','-7 days')
                ), 0) AS used_7d,

                COALESCE((
                    SELECT SUM(t.qty)
                    FROM transactions t
                    WHERE t.sku=i.sku
                      AND t.direction='OUT'
                      AND t.timestamp>=datetime('now','-30 days')
                ), 0) AS used_30d,

                (
                    SELECT t.timestamp
                    FROM transactions t
                    WHERE t.sku=i.sku
                      AND t.direction='OUT'
                    ORDER BY t.timestamp DESC
                    LIMIT 1
                ) AS last_out_timestamp,

                (
                    SELECT t.timestamp
                    FROM transactions t
                    WHERE t.sku=i.sku
                      AND t.direction='OUT'
                    ORDER BY t.timestamp DESC
                    LIMIT 1 OFFSET 1
                ) AS prev_out_timestamp,

                ROUND((
                    SELECT AVG((strftime('%s', x.timestamp) - strftime('%s', x.prev_ts)) / 3600.0)
                    FROM (
                        SELECT
                            t.timestamp,
                            LAG(t.timestamp) OVER (ORDER BY t.timestamp) AS prev_ts
                        FROM transactions t
                        WHERE t.sku=i.sku
                          AND t.direction='OUT'
                    ) x
                    WHERE x.prev_ts IS NOT NULL
                ), 2) AS avg_hours_between_out,

                ROUND((
                    SELECT (strftime('%s','now') - strftime('%s', t.timestamp)) / 3600.0
                    FROM transactions t
                    WHERE t.sku=i.sku
                      AND t.direction='OUT'
                    ORDER BY t.timestamp DESC
                    LIMIT 1
                ), 2) AS hours_since_last_out

            FROM inventory i
            ORDER BY i.category, i.description
        """))

        total = len(inventory)
        low = len([i for i in inventory if 0 < i["quantity"] <= i["min_threshold"]])
        out_count = len([i for i in inventory if i["quantity"] <= 0])

        jobs = _rowdicts(conn.execute("""
            SELECT *
            FROM job_runs
            WHERE status='RUNNING'
               OR (status='COMPLETE' AND date(last_seen)=date('now'))
            ORDER BY last_seen DESC
        """))

        job_series = _rowdicts(conn.execute("""
            SELECT *
            FROM job_series
            WHERE job_num IS NOT NULL
              AND job_num != ''
            ORDER BY due_date='', due_date, job_num
        """))

        cleaned_job_series = []
        for js in job_series:
            job_num = normalize_job_num(js.get("job_num", ""))
            if not job_num or job_num_is_all_zeros(job_num):
                # Let fake jobs still be deleted through admin cleanup,
                # but avoid rendering them as normal active jobs for non-admin users.
                if not admin_ok:
                    continue

            if admin_ok:
                js.update(compute_job_financial(conn, job_num))
            else:
                for f in FINANCIAL_FIELDS:
                    js[f] = None
                js["labor_rate_saved"] = None
                js["labor_rate_effective"] = None

            cleaned_job_series.append(js)

        job_components = _rowdicts(conn.execute("""
            SELECT *
            FROM job_components
            WHERE job_num IS NOT NULL
              AND job_num != ''
            ORDER BY job_num, sort_order, component_name
        """))

        # Improve estimates in returned job_components when possible without permanently
        # writing unless the normal runtime updater does it.
        for jc in job_components:
            if float(jc.get("estimated_sec") or 0) <= 0:
                hist = similar_completed_component_estimate_sec(
                    conn,
                    job_num=jc.get("job_num", ""),
                    component_name=jc.get("component_name", ""),
                )
                if hist.get("ok") and float(hist.get("estimated_sec") or 0) > 0:
                    jc["estimated_sec"] = float(hist.get("estimated_sec") or 0)
                    jc["estimate_source"] = hist.get("source") or "SIMILAR_COMPLETED_JOB"
                    jc["estimate_basis"] = hist.get("basis") or ""

        job_status_log = _rowdicts(conn.execute("""
            SELECT *
            FROM job_status_log
            ORDER BY timestamp DESC
            LIMIT 300
        """))

        inventory_aliases = _rowdicts(conn.execute("""
            SELECT *
            FROM inventory_aliases
            ORDER BY updated_at DESC
            LIMIT 500
        """))

        catalog = _rowdicts(conn.execute("""
            SELECT *
            FROM part_catalog
            ORDER BY job_id
        """))

        match_history = _rowdicts(conn.execute("""
            SELECT *
            FROM match_history
            ORDER BY timestamp DESC
            LIMIT 100
        """))

        # Efficiency tab/data is hidden unless admin.
        eff_day = []
        eff_month = []
        eff_year = []

        if admin_ok:
            eff_day = efficiency_rows_for_period(conn, "day", _today_str())
            eff_month = efficiency_rows_for_period(conn, "month", datetime.now().strftime("%Y-%m"))
            eff_year = efficiency_rows_for_period(conn, "year", datetime.now().strftime("%Y"))

        # Admin-page-only data hidden unless logged in.
        recent_telemetry = []
        gcode_estimates = []
        steel_sheet_items = []
        steel_sheet_logs = []
        umc_pot_schedule = []
        shelves = []
        autoloader_plans = []
        autoloader_events = []

        if admin_ok:
            recent_telemetry = _rowdicts(conn.execute("""
                SELECT *
                FROM telemetry_samples
                ORDER BY id DESC
                LIMIT 250
            """))

            gcode_estimates = _rowdicts(conn.execute("""
                SELECT *
                FROM gcode_tool_estimates
                ORDER BY analyzed_at DESC, machine, program_key, CAST(tid AS INTEGER)
                LIMIT 500
            """))

            steel_sheet_items = _rowdicts(conn.execute("""
                SELECT *
                FROM steel_sheet_items
                ORDER BY parsed_at DESC, job_num, id
                LIMIT 1000
            """))

            steel_sheet_logs = _rowdicts(conn.execute("""
                SELECT *
                FROM steel_sheet_parse_log
                ORDER BY id DESC
                LIMIT 200
            """))

            umc_pot_schedule = _rowdicts(conn.execute("""
                SELECT *
                FROM machine_schedule_plan
                WHERE machine='UMC-1000'
                  AND status='PLANNED'
                ORDER BY schedule_rank
            """))

            shelves = _rowdicts(conn.execute("""
                SELECT *
                FROM autoloader_shelves
                ORDER BY shelf_level DESC, shelf_col, shelf_code
            """))

            autoloader_plans = _rowdicts(conn.execute("""
                SELECT *
                FROM autoloader_queue
                ORDER BY id DESC
                LIMIT 100
            """))

            autoloader_events = _rowdicts(conn.execute("""
                SELECT *
                FROM autoloader_events
                ORDER BY id DESC
                LIMIT 100
            """))

        shop_board_rows = _rowdicts(conn.execute("""
            SELECT *
            FROM shop_board_rows
            ORDER BY sort_order, due_date='', due_date, job_num, id
        """))

        machines_out = safe_machine_public_state(tool_catalog)

        network_info = {
            "app_name": APP_NAME,
            "short_name": APP_SHORT_NAME,
            "port": APP_PORT,
            "urls": get_local_network_addresses(APP_PORT),
        }

        signature_sync = sync_matching_software_signatures(conn)
        signature_recent_matches = recent_signature_match_summaries(conn)
        conn.commit()

        return {
            "app": {
                "name": APP_NAME,
                "short_name": APP_SHORT_NAME,
                "slug": APP_SLUG,
                "port": APP_PORT,
            },
            "network": network_info,
            "admin": {"authenticated": admin_ok},
            "summary": {"total": total, "low": low, "out": out_count},
            "tools": tools,
            "tool_catalog": tool_catalog,
            "tool_replacement_candidates": pending_tool_candidates if admin_ok else [],
            "machines": machines_out,
            "inventory": inventory,
            "inventory_aliases": inventory_aliases,
            "jobs": jobs,
            "job_series": cleaned_job_series,
            "job_components": job_components,
            "job_status_log": job_status_log,
            "catalog": catalog,
            "match_history": match_history,
            "signature_sync": signature_sync,
            "signature_recent_matches": signature_recent_matches,
            "shop_board_rows": shop_board_rows,
            "eff_day": merge_efficiency_rows(eff_day, "day") if admin_ok else [],
            "eff_month": merge_efficiency_rows(eff_month, "month") if admin_ok else [],
            "eff_year": merge_efficiency_rows(eff_year, "year") if admin_ok else [],
            "telemetry_samples": recent_telemetry,
            "gcode_estimates": gcode_estimates,
            "steel_sheet_items": steel_sheet_items,
            "steel_sheet_logs": steel_sheet_logs,
            "umc_pot_schedule": umc_pot_schedule,
            "autoloader": {
                "mode": "SIMULATION_ONLY",
                "shelves": shelves,
                "plans": autoloader_plans,
                "events": autoloader_events,
            },
        }


# ============================================================
# INVENTORY / QR API
# ============================================================
@app.post("/api/inventory/items")
async def create_inventory_item(req: InventoryItemRequest):
    sku = (req.sku or "").strip().upper()
    desc = (req.description or "").strip()
    cat = (req.category or "Tool").strip()

    if not sku:
        raise HTTPException(400, "SKU required")
    if not desc:
        raise HTTPException(400, "Tool/item name required")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            INSERT INTO inventory (sku, description, category, quantity, min_threshold)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(sku) DO UPDATE SET
                description=excluded.description,
                category=excluded.category,
                quantity=excluded.quantity,
                min_threshold=excluded.min_threshold
        """, (
            sku,
            desc,
            cat,
            max(0, req.quantity),
            max(0, req.min_threshold),
        ))
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", "sku": sku}


@app.post("/api/inventory/qr-resolve")
async def qr_resolve_inventory(req: QRResolveRequest):
    """
    QR workflow:
    scan first, then decide TAKE OUT or ADD/RESTOCK.
    """
    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        result = resolve_qr_inventory(conn, req.barcode)

    return result


@app.post("/api/inventory/qr-action")
async def qr_apply_inventory_action(req: QRActionRequest):
    """
    QR action:
    - Existing QR/item: take out or add.
    - Unknown QR: create/map item, then take out or add.
    - Supports quantity add/subtract.
    - Does not create multiple different QR identities for one tool.
    """
    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        result = apply_qr_inventory_action(conn, req)
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return result


@app.post("/api/inventory/auto-scan")
async def auto_scan_inventory(req: AutoScanRequest):
    """
    Legacy compatibility endpoint.
    Kept so old UI/code still works, but main UI uses qr-resolve + qr-action.
    """
    action_req = QRActionRequest(
        barcode=req.barcode,
        direction=req.direction,
        qty=max(1, int(req.qty or 1)),
        sku="",
        description="",
        category="Tool",
        starting_quantity=0,
        min_threshold=5,
    )

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        result = apply_qr_inventory_action(conn, action_req)
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return result


@app.post("/api/inventory/alias")
async def save_inventory_alias(req: QRAliasRequest):
    alias = normalize_barcode_alias(req.alias)
    sku = (req.sku or "").strip().upper()
    desc = (req.description or "").strip()
    cat = (req.category or "Tool").strip()

    if not alias:
        raise HTTPException(400, "QR/barcode alias required")
    if not sku:
        raise HTTPException(400, "SKU required")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        existing = conn.execute(
            "SELECT sku FROM inventory WHERE sku=?",
            (sku,),
        ).fetchone()

        if not existing:
            conn.execute("""
                INSERT INTO inventory (
                    sku, description, category, quantity, min_threshold
                )
                VALUES (?, ?, ?, ?, ?)
            """, (
                sku,
                desc or sku,
                cat,
                max(0, int(req.quantity or 0)),
                max(0, int(req.min_threshold or 0)),
            ))
        elif desc:
            conn.execute("""
                UPDATE inventory
                SET description=?, category=?
                WHERE sku=?
            """, (desc, cat, sku))

        conn.execute("""
            INSERT INTO inventory_aliases (
                alias, sku, description, created_at, updated_at
            )
            VALUES (?, ?, ?, datetime('now'), datetime('now'))
            ON CONFLICT(alias) DO UPDATE SET
                sku=excluded.sku,
                description=excluded.description,
                updated_at=datetime('now')
        """, (alias, sku, desc))

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", "alias": alias, "sku": sku}


@app.post("/api/scan")
async def process_scan(req: ScanRequest):
    sku = (req.sku or "").strip().upper()
    direction = (req.direction or "OUT").upper()
    qty = max(1, int(req.qty or 1))

    if direction not in ("IN", "OUT"):
        raise HTTPException(400, "direction must be IN or OUT")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        row = conn.execute("SELECT quantity FROM inventory WHERE sku=?", (sku,)).fetchone()

        if not row:
            raise HTTPException(404, "SKU missing")

        new_qty = max(0, row[0] + (qty if direction == "IN" else -qty))

        conn.execute("UPDATE inventory SET quantity=? WHERE sku=?", (new_qty, sku))

        conn.execute(
            "INSERT INTO transactions (sku, direction, qty) VALUES (?, ?, ?)",
            (sku, direction, qty),
        )
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", "sku": sku, "direction": direction, "qty": qty, "new_quantity": new_qty}


# ============================================================
# TOOL DETAIL / REPLACEMENT / INSPECTION API
# ============================================================
@app.post("/api/tools/detail")
async def tool_detail(body: dict = Body(...)):
    machine = body.get("machine", "")
    tid = normalize_tool_num(body.get("tid", ""))

    if not machine or not tid or tid == "0":
        raise HTTPException(400, "machine and tid required")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        health = conn.execute("""
            SELECT
                th.*,
                COALESCE(NULLIF(th.tool_name,''), tc.tool_name, 'Tool ' || th.tid) AS display_tool_name,
                tc.category AS tool_category,
                tc.diameter,
                tc.flutes,
                tc.work_material,
                tc.max_life_min AS catalog_max_life_min
            FROM tool_health th
            LEFT JOIN tool_catalog tc ON th.tid=tc.tid
            WHERE th.machine=? AND th.tid=?
        """, (machine, tid)).fetchone()

        if not health:
            raise HTTPException(404, "Tool health not found")

        health = dict(health)
        health["tid"] = normalize_tool_num(health.get("tid", tid))
        health["tool_name"] = health.get("display_tool_name") or health.get("tool_name") or get_tool_name(tid)
        health["health_color"] = tool_health_color(float(health.get("wear_score") or 0))

        samples = _rowdicts(conn.execute("""
            SELECT *
            FROM tool_cut_samples
            WHERE machine=? AND tid=?
            ORDER BY id DESC
            LIMIT 1500
        """, (machine, tid)))
        samples = list(reversed(samples))

        cycles = _rowdicts(conn.execute("""
            SELECT *
            FROM tool_life_cycles
            WHERE machine=? AND tid=?
            ORDER BY id DESC
            LIMIT 30
        """, (machine, tid)))

        events = _rowdicts(conn.execute("""
            SELECT *
            FROM tool_replacement_events
            WHERE machine=? AND tid=?
            ORDER BY id DESC
            LIMIT 30
        """, (machine, tid)))

        candidates = _rowdicts(conn.execute("""
            SELECT *
            FROM tool_replacement_candidates
            WHERE machine=? AND tid=?
            ORDER BY id DESC
            LIMIT 30
        """, (machine, tid)))

        inspections = _rowdicts(conn.execute("""
            SELECT *
            FROM tool_inspections
            WHERE machine=? AND tid=?
            ORDER BY id DESC
            LIMIT 50
        """, (
            machine,
            tid,
        )))

        comparison = summarize_previous_tool_cycles(conn, machine, tid, 5)

        active_cut = float(health.get("in_cut_sec") or 0)
        if comparison.get("avg_cut_sec", 0) > 0:
            pct_vs_avg = round((active_cut / comparison["avg_cut_sec"]) * 100.0, 1)
        else:
            pct_vs_avg = 0

        comparison["current_vs_previous_avg_pct"] = pct_vs_avg

        return {
            "status": "ok",
            "machine": machine,
            "tid": tid,
            "health": health,
            "samples": samples,
            "cycles": cycles,
            "events": events,
            "candidates": candidates,
            "inspections": inspections,
            "comparison": comparison,
        }


@app.post("/api/tools/replace")
async def manual_replace_tool(request: Request, req: ToolReplaceRequest):
    require_admin_request(request)

    machine = req.machine
    tid = normalize_tool_num(req.tid)

    if not machine or not tid or tid == "0":
        raise HTTPException(400, "machine and tid required")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        result = archive_and_start_new_tool_cycle(
            conn=conn,
            machine=machine,
            tid=tid,
            reason=req.reason or "Manual replacement",
            notes=req.notes or "",
            source="MANUAL",
            candidate_id=0,
        )
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", **result}


@app.post("/api/tools/inspect")
async def inspect_tool(request: Request, req: ToolInspectionRequest):
    require_admin_request(request)

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        result = record_tool_inspection(conn, req)
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", **result}


@app.post("/api/tools/replacement-candidates/{candidate_id}/confirm")
async def confirm_replacement_candidate(
    candidate_id: int,
    request: Request,
    req: ToolCandidateActionRequest,
):
    require_admin_request(request)

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        cand = conn.execute("""
            SELECT *
            FROM tool_replacement_candidates
            WHERE id=?
        """, (int(candidate_id),)).fetchone()

        if not cand:
            raise HTTPException(404, "Candidate not found")

        c = dict(cand)
        if c.get("status") != "PENDING":
            raise HTTPException(409, "Candidate is not pending")

        result = archive_and_start_new_tool_cycle(
            conn=conn,
            machine=c["machine"],
            tid=c["tid"],
            reason=req.reason or "Operator confirmed actual tool replacement after offset change",
            notes=req.notes or "",
            source="OFFSET_CHANGE_CONFIRMED_BY_OPERATOR",
            candidate_id=int(candidate_id),
        )

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", **result}


@app.post("/api/tools/replacement-candidates/{candidate_id}/ignore")
async def ignore_replacement_candidate(
    candidate_id: int,
    request: Request,
    req: ToolCandidateActionRequest,
):
    require_admin_request(request)

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        result = ignore_tool_replacement_candidate(conn, int(candidate_id), req.notes or "")
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", **result}


@app.post("/api/tools/replacement-candidates/clear-pending")
async def clear_pending_replacement_candidates(request: Request):
    require_admin_request(request)

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            UPDATE tool_replacement_candidates
            SET
                status='IGNORED',
                ignored_at=datetime('now'),
                notes=CASE
                    WHEN notes IS NULL OR notes=''
                    THEN 'Bulk cleared as false-positive offset change warning.'
                    ELSE notes || ' | Bulk cleared as false-positive offset change warning.'
                END
            WHERE status='PENDING'
        """)

        conn.execute("""
            UPDATE tool_health
            SET replacement_pending=0
        """)

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok"}


# ============================================================
# G-CODE API
# ============================================================
@app.post("/api/gcode/analyze")
async def api_gcode_analyze(body: dict = Body(...)):
    """
    Non-admin accessible. It does not expose money or perform destructive changes.
    """
    machine = body.get("machine", "")
    prog_raw = body.get("program", "")

    if not machine:
        raise HTTPException(400, "machine required")

    if not prog_raw:
        raise HTTPException(400, "program required")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        result = analyze_and_store_gcode_program(conn, machine, prog_raw)

        rows = _rowdicts(conn.execute("""
            SELECT *
            FROM gcode_tool_estimates
            WHERE machine=? AND program_key=?
            ORDER BY CAST(tid AS INTEGER)
        """, (
            machine,
            normalize_program_key(prog_raw),
        )))

        conn.commit()

    return {
        "status": "ok",
        "analysis": result,
        "rows": rows,
    }


# ============================================================
# STEEL SHEET API
# ============================================================
@app.post("/api/steel-sheet/parse")
async def api_parse_steel_sheet(body: dict = Body(...)):
    """
    Legacy URL parser.
    UI no longer uses this as the primary steel-sheet workflow.
    """
    job_num = validate_new_job_num_or_400(body.get("job_num", ""))
    source_url = body.get("source_url", "") or body.get("steel_sheet_url", "")

    if not source_url:
        with db_lock, sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute("""
                SELECT steel_sheet_url
                FROM job_series
                WHERE job_num=?
            """, (job_num,)).fetchone()

            if row:
                source_url = row["steel_sheet_url"] or ""

    if not source_url:
        raise HTTPException(400, "source_url or steel_sheet_url required")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        result = parse_and_store_steel_sheet(
            conn,
            job_num=job_num,
            source_url=source_url,
            force=True,
        )

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return result


@app.get("/api/steel-sheet/{job_num}")
async def api_get_steel_sheet(job_num: str):
    j = normalize_job_num(job_num)
    if not j:
        raise HTTPException(400, "job_num required")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        items = get_steel_sheet_items_for_job(conn, j)
        summary = steel_sheet_summary_for_job(conn, j)

        logs = _rowdicts(conn.execute("""
            SELECT *
            FROM steel_sheet_parse_log
            WHERE job_num=?
            ORDER BY id DESC
            LIMIT 20
        """, (j,)))

    return {
        "status": "ok",
        "job_num": j,
        "summary": summary,
        "items": items,
        "logs": logs,
    }


@app.post("/api/steel-sheet/parse-all")
async def api_parse_all_steel_sheets():
    """
    Legacy URL parser for old jobs that still have steel_sheet_url saved.
    """
    results = []

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        jobs = _rowdicts(conn.execute("""
            SELECT job_num, steel_sheet_url
            FROM job_series
            WHERE steel_sheet_url IS NOT NULL
              AND steel_sheet_url != ''
            ORDER BY due_date='', due_date, job_num
        """))

        for j in jobs:
            job_num = normalize_job_num(j["job_num"])
            if not job_num or job_num_is_all_zeros(job_num):
                continue

            result = parse_and_store_steel_sheet(
                conn,
                job_num=job_num,
                source_url=j["steel_sheet_url"],
                force=True,
            )
            results.append(result)

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))

    return {
        "status": "ok",
        "count": len(results),
        "results": results,
    }


@app.post("/api/steel-sheet/upload-pdf")
async def api_upload_steel_sheet_pdf(
    job_num: str = Form(...),
    due_date: str = Form(""),
    file: UploadFile = File(...),
):
    """
    PDF upload backend retained.
    UI should only expose this through Job Matching / lookup-with-pdf.
    """
    j = validate_new_job_num_or_400(job_num)

    if due_date:
        try:
            datetime.strptime(due_date[:10], "%Y-%m-%d")
            due_date = due_date[:10]
        except Exception:
            raise HTTPException(400, "Due date must be YYYY-MM-DD")

    pdf_bytes = await file.read()
    if not pdf_bytes:
        raise HTTPException(400, "Uploaded PDF was empty")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        ensure_job_series(conn, j)
        result = save_uploaded_steel_pdf_for_job(
            conn=conn,
            job_num=j,
            pdf_bytes=pdf_bytes,
            original_filename=file.filename or "steel_sheet.pdf",
            due_date=due_date,
        )
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return result


# ============================================================
# UMC-1000 POT SCHEDULER API
# ============================================================
@app.post("/api/schedule/umc1000-pots/optimize")
async def optimize_umc1000_pot_schedule(body: dict = Body(default={})):
    start_at = body.get("start_at", "") or ""

    if start_at:
        try:
            start_dt = datetime.strptime(start_at[:19], "%Y-%m-%d %H:%M:%S")
        except Exception:
            try:
                start_dt = datetime.strptime(start_at[:16], "%Y-%m-%dT%H:%M")
            except Exception:
                start_dt = datetime.now()
    else:
        start_dt = datetime.now()

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        plan = build_umc1000_pot_schedule(conn, start_dt=start_dt, save=True)
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))

    return {
        "status": "ok",
        "machine": UMC_SCHEDULER_MACHINE,
        "count": len(plan),
        "plan": plan,
    }


@app.get("/api/schedule/umc1000-pots/current")
async def get_umc1000_pot_schedule():
    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        rows = _rowdicts(conn.execute("""
            SELECT *
            FROM machine_schedule_plan
            WHERE machine=?
              AND status='PLANNED'
            ORDER BY schedule_rank
        """, (UMC_SCHEDULER_MACHINE,)))

    return {
        "status": "ok",
        "machine": UMC_SCHEDULER_MACHINE,
        "plan": rows,
    }


# ============================================================
# SCHEDULE / JOB API
# ============================================================
@app.post("/api/schedule")
async def add_schedule(body: dict = Body(...)):
    job_id = validate_new_job_num_or_400(body.get("job_id", ""))

    due_date = body.get("due_date", "") or ""
    notes = body.get("notes", "") or ""

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        ensure_job_series(conn, job_id, due_date, "", notes)

        conn.execute("""
            INSERT INTO schedule (job_id, due_date, steel_sheet_url, notes, status)
            VALUES (?, ?, '', ?, 'QUEUED')
        """, (job_id, due_date, notes))

        conn.execute("""
            INSERT INTO shop_board_rows (
                job_num, description, due_date, notes, row_status, updated_at
            )
            SELECT ?, ?, ?, ?, 'OPEN', datetime('now')
            WHERE NOT EXISTS (
                SELECT 1 FROM shop_board_rows WHERE job_num=?
            )
        """, (
            job_id,
            notes or "",
            due_date,
            "Created from calendar schedule.",
            job_id,
        ))

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))

    return {
        "status": "ok",
        "job_id": job_id,
        "due_date": due_date,
    }


@app.post("/api/calendar/jobs/{job_num}/delete")
async def delete_job_from_calendar(job_num: str, request: Request):
    """
    Admin-only calendar delete.

    Removes job from app calendar without deleting the entire job.
    It clears due dates and schedule rows, but keeps job history/components.
    """
    require_admin_request(request)

    j = normalize_job_num(job_num)
    if not j:
        raise HTTPException(400, "Job number required")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        conn.execute("""
            DELETE FROM schedule
            WHERE job_id=?
        """, (j,))

        conn.execute("""
            UPDATE job_series
            SET
                due_date='',
                updated_at=datetime('now')
            WHERE job_num=?
        """, (j,))

        conn.execute("""
            UPDATE part_catalog
            SET due_date=''
            WHERE job_id=?
        """, (j,))

        conn.execute("""
            UPDATE shop_board_rows
            SET
                due_date='',
                updated_at=datetime('now')
            WHERE job_num=?
        """, (j,))

        conn.execute("""
            INSERT INTO job_status_log (
                job_num, status, message
            )
            VALUES (?, 'CALENDAR_REMOVED', ?)
        """, (
            j,
            "Admin removed job from app calendar. Job was not deleted.",
        ))

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))

    return {
        "status": "ok",
        "job_num": j,
        "message": f"J{j} removed from calendar. Job history was kept.",
    }


@app.put("/api/job-series/{job_num}/financials")
async def update_job_financials(job_num: str, request: Request, body: dict = Body(...)):
    require_admin_request(request)

    j = validate_new_job_num_or_400(job_num)

    # Important:
    # Do not use "or 75" on labor rate. Save exactly what admin sends.
    fields = {
        "revenue": _money(body.get("revenue")),
        "steel_cost": _money(body.get("steel_cost")),
        "grinding_cost": _money(body.get("grinding_cost")),
        "pyropel_cost": _money(body.get("pyropel_cost")),
        "labor_rate": _money(body.get("labor_rate")),
        "labor_override_cost": _money(body.get("labor_override_cost")),
    }

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        ensure_job_series(conn, j)

        conn.execute("""
            UPDATE job_series
            SET
                revenue=?,
                steel_cost=?,
                grinding_cost=?,
                pyropel_cost=?,
                labor_rate=?,
                labor_override_cost=?,
                updated_at=datetime('now')
            WHERE job_num=?
        """, (
            fields["revenue"],
            fields["steel_cost"],
            fields["grinding_cost"],
            fields["pyropel_cost"],
            fields["labor_rate"],
            fields["labor_override_cost"],
            j,
        ))

        result = compute_job_financial(conn, j)
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", "job_num": j, **result}


@app.put("/api/job-series/{job_num}/due-date")
async def update_job_due_date(job_num: str, body: dict = Body(...)):
    j = validate_new_job_num_or_400(job_num)

    due_date = str(body.get("due_date") or "").strip()

    if due_date:
        try:
            datetime.strptime(due_date[:10], "%Y-%m-%d")
            due_date = due_date[:10]
        except Exception:
            raise HTTPException(400, "Due date must be YYYY-MM-DD")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        ensure_job_series(conn, j)

        conn.execute("""
            UPDATE job_series
            SET due_date=?, updated_at=datetime('now')
            WHERE job_num=?
        """, (
            due_date,
            j,
        ))

        conn.execute("""
            UPDATE part_catalog
            SET due_date=?
            WHERE job_id=?
        """, (
            due_date,
            j,
        ))

        conn.execute("""
            UPDATE schedule
            SET due_date=?
            WHERE job_id=?
        """, (
            due_date,
            j,
        ))

        conn.execute("""
            UPDATE shop_board_rows
            SET due_date=?, updated_at=datetime('now')
            WHERE job_num=?
        """, (
            due_date,
            j,
        ))

        conn.execute("""
            INSERT INTO job_status_log (
                job_num, status, message
            )
            VALUES (?, 'DUE_DATE_UPDATED', ?)
        """, (
            j,
            f"Due date updated to {due_date or 'blank'}.",
        ))

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))

    return {
        "status": "ok",
        "job_num": j,
        "due_date": due_date,
    }



@app.post("/api/job-signatures/{job_num}/upload")
async def upload_job_signature_csv(job_num: str, request: Request, file: UploadFile = File(...)):
    require_signature_or_admin_request(request)

    j = normalize_job_num(job_num)
    if not j:
        raise HTTPException(400, "Job number required")

    data = await file.read()
    try:
        text = data.decode("utf-8-sig")
    except Exception:
        text = data.decode("latin-1", errors="replace")

    rows = parse_job_signature_csv_text(text, fallback_job_num=j)
    if not rows:
        raise HTTPException(400, "No usable signature rows found in CSV")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        imported = upsert_job_signature_rows(conn, j, rows, source_file=file.filename or "")
        matches = compare_job_signature_to_previous(conn, j, limit=JOB_SIGNATURE_DEFAULT_LIMIT)
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))

    return {
        "status": "ok",
        "job_num": j,
        "imported_components": imported,
        "expected_components": JOB_SIGNATURE_COMPONENTS,
        "matches": matches,
    }


@app.get("/api/job-signatures/{job_num}/matches")
async def get_job_signature_matches(job_num: str, limit: int = JOB_SIGNATURE_DEFAULT_LIMIT):
    j = normalize_job_num(job_num)
    if not j:
        raise HTTPException(400, "Job number required")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        current = load_job_signature(conn, j)
        matches = compare_job_signature_to_previous(conn, j, limit=limit)

    return {
        "job_num": j,
        "has_signature": bool(current),
        "components": JOB_SIGNATURE_COMPONENTS,
        "signature_components": sorted(current.keys()),
        "matches": matches,
    }


@app.delete("/api/job-series/{job_num}")
async def delete_job_series_admin(
    job_num: str,
    request: Request,
    body: JobDeleteRequest = Body(default=JobDeleteRequest()),
):
    require_admin_request(request)

    j = normalize_job_num(job_num)
    if not j:
        raise HTTPException(400, "Job number required")

    expected = f"DELETE J{j}"
    if (body.confirm or "").strip().upper() != expected.upper():
        raise HTTPException(
            400,
            f"Confirmation required. Send confirm='{expected}'"
        )

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        result = delete_job_everywhere(conn, j)
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", **result}


@app.post("/api/job-series/{job_num}/rename")
async def rename_job_series_admin(job_num: str, request: Request, req: JobRenameRequest):
    require_admin_request(request)

    old = normalize_job_num(job_num)
    new = validate_new_job_num_or_400(req.new_job_num)

    if not old:
        raise HTTPException(400, "Old job number required")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        result = rename_or_merge_job_everywhere(
            conn,
            old,
            new,
            merge_if_exists=bool(req.merge_if_exists),
        )

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", **result}


@app.post("/api/job-series/cleanup-fake")
async def cleanup_fake_jobs_admin(request: Request, body: dict = Body(default={})):
    require_admin_request(request)

    jobs = body.get("jobs") or ["000", "0000", "1000"]

    cleaned = []

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        for j in jobs:
            jj = normalize_job_num(j)
            if not jj:
                continue
            cleaned.append(delete_job_everywhere(conn, jj))

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))

    return {
        "status": "ok",
        "cleaned": cleaned,
    }


@app.post("/api/job-series/{job_num}/components/{component}/status")
async def update_component_status(job_num: str, component: str, body: dict = Body(...)):
    j = validate_new_job_num_or_400(job_num)
    c = canonical_component(component)
    status = (body.get("status", "PENDING") or "PENDING").upper()

    if status not in ("PENDING", "RUNNING", "COMPLETE"):
        raise HTTPException(400, "Bad status")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        ensure_job_series(conn, j)

        est_sec = float(DEFAULT_COMPONENT_EST_SEC.get(c, 0) or 0)

        conn.execute("""
            INSERT OR IGNORE INTO job_components (
                job_num, component_name, machine, status, sort_order, estimated_sec,
                estimate_source, estimate_basis, estimate_updated_at
            )
            VALUES (?, ?, ?, 'PENDING', 99, ?, '', '', '')
        """, (j, c, DEFAULT_COMPONENT_MACHINE.get(c, ""), est_sec))

        conn.execute("""
            UPDATE job_components SET
                status=?,
                completed_at=CASE WHEN ?='COMPLETE' THEN datetime('now') ELSE NULL END,
                manual_complete=CASE WHEN ?='COMPLETE' THEN 1 ELSE 0 END,
                last_status_note=?
            WHERE job_num=? AND component_name=?
        """, (
            status,
            status,
            status,
            f"Manually set to {status}",
            j,
            c,
        ))

        log_job_status(conn, j, c, status, "", "", f"Manually set to {status}")
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok"}


# ============================================================
# JOB MATCHING / CATALOG API
# ============================================================
@app.post("/api/catalog")
async def save_catalog(body: dict = Body(...)):
    job_id = validate_new_job_num_or_400(body.get("job_id", ""))

    due_date = body.get("due_date", "") or ""

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            INSERT OR REPLACE INTO part_catalog (
                job_id, weight, com_x, com_y, com_z, due_date, steel_sheet_url, notes
            )
            VALUES (?, ?, ?, ?, ?, ?, '', ?)
        """, (
            job_id,
            body.get("weight", 0),
            body.get("com_x", 0),
            body.get("com_y", 0),
            body.get("com_z", 0),
            due_date,
            body.get("notes", ""),
        ))

        ensure_job_series(conn, job_id, due_date, "")

        conn.execute("""
            INSERT INTO shop_board_rows (
                job_num, description, due_date, steel, notes, row_status, updated_at
            )
            SELECT ?, ?, ?, '', ?, 'OPEN', datetime('now')
            WHERE NOT EXISTS (
                SELECT 1 FROM shop_board_rows WHERE job_num=?
            )
        """, (
            job_id,
            body.get("notes", ""),
            due_date,
            "Created from catalog.",
            job_id,
        ))

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", "job_id": job_id}


@app.delete("/api/catalog/{job_id}")
async def del_catalog(job_id: str, request: Request):
    require_admin_request(request)

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.execute("DELETE FROM part_catalog WHERE job_id=?", (normalize_job_num(job_id),))
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok"}


@app.post("/api/catalog/lookup")
async def catalog_lookup(body: dict = Body(...)):
    job_id = validate_new_job_num_or_400(body.get("job_id", ""))

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        result = catalog_lookup_core(
            conn=conn,
            job_id=job_id,
            weight=float(body.get("weight", 0) or 0),
            com_x=float(body.get("com_x", 0) or 0),
            com_y=float(body.get("com_y", 0) or 0),
            com_z=float(body.get("com_z", 0) or 0),
            tolerance=float(body.get("tolerance", 0.01) or 0.01),
            due_date=body.get("due_date", "") or "",
            notes=body.get("notes", "") or "Reverse Lookup",
        )

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return result


@app.post("/api/catalog/lookup-with-pdf")
async def catalog_lookup_with_pdf(
    job_id: str = Form(...),
    weight: float = Form(0),
    com_x: float = Form(0),
    com_y: float = Form(0),
    com_z: float = Form(0),
    tolerance: float = Form(0.01),
    due_date: str = Form(""),
    notes: str = Form(""),
    file: UploadFile = File(None),
):
    """
    Main workflow:
    - Job Matching tab uploads the steel-sheet PDF here.
    - No URL field required.
    - This endpoint performs matching and attaches/parses the PDF.
    """
    j = validate_new_job_num_or_400(job_id)

    if due_date:
        try:
            datetime.strptime(due_date[:10], "%Y-%m-%d")
            due_date = due_date[:10]
        except Exception:
            raise HTTPException(400, "Due date must be YYYY-MM-DD")

    pdf_result = {
        "status": "no_file",
        "job_num": j,
        "message": "No PDF uploaded.",
        "count": 0,
        "items": [],
    }

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        result = catalog_lookup_core(
            conn=conn,
            job_id=j,
            weight=float(weight or 0),
            com_x=float(com_x or 0),
            com_y=float(com_y or 0),
            com_z=float(com_z or 0),
            tolerance=float(tolerance or 0.01),
            due_date=due_date,
            notes=notes or "Reverse Lookup with steel PDF",
        )

        if file is not None:
            pdf_bytes = await file.read()
            if pdf_bytes:
                pdf_result = save_uploaded_steel_pdf_for_job(
                    conn=conn,
                    job_num=j,
                    pdf_bytes=pdf_bytes,
                    original_filename=file.filename or "steel_sheet.pdf",
                    due_date=result.get("due_date") or due_date or "",
                )

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {
        **result,
        "pdf_upload": pdf_result,
    }


# ============================================================
# SHOP BOARD API
# ============================================================
SHOP_BOARD_FIELDS = [
    "job_num",
    "description",
    "due_date",
    "priority",
    "row_status",
    "row_color",
    "sort_order",
    "finals_mex",
    "finals_john",
    "steel",
    "ht200",
    "xt_sent",
    "pyropel",
    "j_blk",
    "a2_1",
    "req_1",
    "d2",
    "req_2",
    "other",
    "pullcore_cam",
    "pullcore_key",
    "h13",
    "req_3",
    "a2_2",
    "req_4",
    "pdf_mc",
    "prog",
    "notes",
]


@app.get("/api/shop-board")
async def get_shop_board():
    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = _rowdicts(conn.execute("""
            SELECT *
            FROM shop_board_rows
            ORDER BY sort_order, due_date='', due_date, job_num, id
        """))

    return {"status": "ok", "rows": rows}


@app.post("/api/shop-board")
async def create_shop_board_row(req: ShopBoardRowRequest):
    data = req.model_dump()
    if data.get("job_num"):
        data["job_num"] = validate_new_job_num_or_400(data.get("job_num", ""))
    else:
        data["job_num"] = ""

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        placeholders = ",".join("?" for _ in SHOP_BOARD_FIELDS)
        cols = ",".join(SHOP_BOARD_FIELDS)

        cur = conn.execute(f"""
            INSERT INTO shop_board_rows (
                {cols}, created_at, updated_at
            )
            VALUES (
                {placeholders}, datetime('now'), datetime('now')
            )
        """, tuple(data.get(f, "") for f in SHOP_BOARD_FIELDS))

        row_id = cur.lastrowid

        conn.execute("""
            INSERT INTO shop_board_log (
                row_id, job_num, message
            )
            VALUES (?, ?, ?)
        """, (
            row_id,
            data.get("job_num", ""),
            "Created shop board row.",
        ))

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", "id": row_id}


@app.put("/api/shop-board/{row_id}")
async def update_shop_board_row(row_id: int, req: ShopBoardRowRequest):
    data = req.model_dump()
    if data.get("job_num"):
        data["job_num"] = validate_new_job_num_or_400(data.get("job_num", ""))
    else:
        data["job_num"] = ""

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        old = conn.execute("""
            SELECT *
            FROM shop_board_rows
            WHERE id=?
        """, (int(row_id),)).fetchone()

        if not old:
            raise HTTPException(404, "Shop board row not found")

        oldd = dict(old)

        assignments = ",".join(f"{f}=?" for f in SHOP_BOARD_FIELDS)
        conn.execute(f"""
            UPDATE shop_board_rows
            SET
                {assignments},
                updated_at=datetime('now')
            WHERE id=?
        """, tuple(data.get(f, "") for f in SHOP_BOARD_FIELDS) + (int(row_id),))

        for f in SHOP_BOARD_FIELDS:
            old_v = str(oldd.get(f, "") or "")
            new_v = str(data.get(f, "") or "")
            if old_v != new_v:
                conn.execute("""
                    INSERT INTO shop_board_log (
                        row_id, job_num, field_name, old_value, new_value, message
                    )
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (
                    int(row_id),
                    data.get("job_num", ""),
                    f,
                    old_v,
                    new_v,
                    f"Changed {f}.",
                ))

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", "id": row_id}


def delete_shop_board_row_core(conn, row_id: int):
    # Important: make fetched rows dict-like.
    conn.row_factory = sqlite3.Row

    row = conn.execute("""
        SELECT *
        FROM shop_board_rows
        WHERE id=?
    """, (int(row_id),)).fetchone()

    if not row:
        raise HTTPException(404, "Shop board row not found")

    r = dict(row)

    conn.execute("""
        INSERT INTO shop_board_log (
            row_id, job_num, message
        )
        VALUES (?, ?, ?)
    """, (
        int(row_id),
        r.get("job_num", ""),
        "Deleted shop board row.",
    ))

    cur = conn.execute("""
        DELETE FROM shop_board_rows
        WHERE id=?
    """, (int(row_id),))

    return {
        "deleted": cur.rowcount,
        "row_id": int(row_id),
        "job_num": r.get("job_num", ""),
    }


@app.delete("/api/shop-board/{row_id}")
async def delete_shop_board_row(row_id: int, request: Request):
    """
    Original DELETE route kept.
    """
    require_admin_request(request)

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        result = delete_shop_board_row_core(conn, int(row_id))
        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", **result}


@app.post("/api/shop-board/sync-from-jobs")
async def sync_shop_board_from_jobs():
    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        jobs = _rowdicts(conn.execute("""
            SELECT *
            FROM job_series
        """))

        created = 0

        for j in jobs:
            job_num = normalize_job_num(j.get("job_num", ""))
            if not job_num or job_num_is_all_zeros(job_num):
                continue

            exists = conn.execute("""
                SELECT id
                FROM shop_board_rows
                WHERE job_num=?
            """, (job_num,)).fetchone()

            if exists:
                continue

            conn.execute("""
                INSERT INTO shop_board_rows (
                    job_num, description, due_date, notes, row_status, updated_at
                )
                VALUES (?, ?, ?, ?, 'OPEN', datetime('now'))
            """, (
                job_num,
                j.get("notes", "") or "",
                j.get("due_date", "") or "",
                "Synced from jobs.",
            ))
            created += 1

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", "created": created}


# ============================================================
# AUTOLOADER / VIRTUAL SHELF API - SIMULATION ONLY
# ============================================================
@app.get("/api/autoloader/shelves")
async def get_autoloader_shelves():
    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = _rowdicts(conn.execute("""
            SELECT *
            FROM autoloader_shelves
            ORDER BY shelf_level DESC, shelf_col, shelf_code
        """))

    return {
        "shelves": rows,
        "mode": "SIMULATION_ONLY",
        "layout": {"high": 5, "wide": 4},
    }


@app.put("/api/autoloader/shelves/{shelf_code}")
async def upsert_autoloader_shelf(
    shelf_code: str,
    body: ShelfUpdateRequest,
):
    code = normalize_component(shelf_code)
    if not code:
        raise HTTPException(400, "Shelf code required")

    status = normalize_component(body.status or "EMPTY")
    if status not in ("EMPTY", "LOADED", "RESERVED", "BLOCKED", "UNKNOWN"):
        raise HTTPException(400, "Bad shelf status")

    job_num = normalize_job_num(body.assigned_job_num)
    if job_num and job_num_is_all_zeros(job_num):
        raise HTTPException(400, "All-zero job numbers like J0000 are invalid")

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            INSERT INTO autoloader_shelves (
                shelf_code, shelf_row, shelf_col, shelf_level,
                status, assigned_job_num, component_name, material,
                pallet_id, fixture_id, part_present_confidence,
                safe_to_run, locked, notes, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
            ON CONFLICT(shelf_code) DO UPDATE SET
                shelf_row=excluded.shelf_row,
                shelf_col=excluded.shelf_col,
                shelf_level=excluded.shelf_level,
                status=excluded.status,
                assigned_job_num=excluded.assigned_job_num,
                component_name=excluded.component_name,
                material=excluded.material,
                pallet_id=excluded.pallet_id,
                fixture_id=excluded.fixture_id,
                part_present_confidence=excluded.part_present_confidence,
                safe_to_run=excluded.safe_to_run,
                locked=excluded.locked,
                notes=excluded.notes,
                updated_at=datetime('now')
        """, (
            code,
            int(body.shelf_row or 0),
            int(body.shelf_col or 0),
            int(body.shelf_level or 0),
            status,
            job_num,
            canonical_component(body.component_name or ""),
            body.material or "",
            body.pallet_id or "",
            body.fixture_id or "",
            float(body.part_present_confidence or 0),
            1 if body.safe_to_run else 0,
            1 if body.locked else 0,
            body.notes or "",
        ))

        conn.execute("""
            INSERT INTO autoloader_events (
                event_type, shelf_code, job_num, component_name, message
            )
            VALUES ('SHELF_UPDATE', ?, ?, ?, ?)
        """, (
            code,
            job_num,
            canonical_component(body.component_name or ""),
            f"Virtual shelf updated. Status={status}. SIMULATION ONLY.",
        ))

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))
    return {"status": "ok", "shelf_code": code, "mode": "SIMULATION_ONLY"}


@app.post("/api/autoloader/plan")
async def create_autoloader_plan(body: AutoloaderPlanRequest):
    job_num = normalize_job_num(body.requested_job_num)
    if job_num and job_num_is_all_zeros(job_num):
        raise HTTPException(400, "All-zero job numbers like J0000 are invalid")

    comp = canonical_component(body.component_name or "")
    from_shelf = normalize_component(body.from_shelf or "")
    to_machine = body.to_machine or ""
    action = normalize_component(body.requested_action or "LOAD")

    if action not in ("LOAD", "UNLOAD", "SWAP", "INSPECT"):
        raise HTTPException(400, "Bad autoloader action")

    safety_notes = (
        "SIMULATION ONLY. No machine, PLC, robot, gantry, door, vise, pallet, or CNC command "
        "is sent by this software. Before real automation: add hardwired interlocks, safe zones, "
        "door status, chuck/vise confirmation, pallet-present sensors, torque/current limits, "
        "E-stop chain, recovery procedure, and dry-run validation."
    )

    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        shelf = None
        if from_shelf:
            shelf = conn.execute("""
                SELECT *
                FROM autoloader_shelves
                WHERE shelf_code=?
            """, (from_shelf,)).fetchone()

            if not shelf:
                raise HTTPException(404, "Shelf does not exist")

            shelf = dict(shelf)

            if shelf.get("locked"):
                raise HTTPException(409, "Shelf is locked")

            if action in ("LOAD", "SWAP") and shelf.get("status") not in ("LOADED", "RESERVED"):
                raise HTTPException(409, "Shelf is not marked LOADED or RESERVED")

            if action in ("LOAD", "SWAP") and not int(shelf.get("safe_to_run") or 0):
                raise HTTPException(409, "Shelf is not marked safe_to_run")

        cur = conn.execute("""
            INSERT INTO autoloader_queue (
                requested_job_num, component_name, from_shelf, to_machine,
                requested_action, status, safety_notes, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, 'SIMULATION_ONLY', ?, datetime('now'), datetime('now'))
        """, (
            job_num,
            comp,
            from_shelf,
            to_machine,
            action,
            safety_notes,
        ))

        plan_id = cur.lastrowid

        conn.execute("""
            INSERT INTO autoloader_events (
                event_type, shelf_code, machine, job_num, component_name, message
            )
            VALUES ('PLAN_CREATED', ?, ?, ?, ?, ?)
        """, (
            from_shelf,
            to_machine,
            job_num,
            comp,
            f"Autoloader plan {plan_id} created in SIMULATION_ONLY mode.",
        ))

        conn.commit()

    await manager.broadcast(json.dumps({"type": "refresh"}))

    return {
        "status": "ok",
        "plan_id": plan_id,
        "mode": "SIMULATION_ONLY",
        "safety_notes": safety_notes,
    }


@app.get("/api/autoloader/plans")
async def get_autoloader_plans():
    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        plans = _rowdicts(conn.execute("""
            SELECT *
            FROM autoloader_queue
            ORDER BY id DESC
            LIMIT 200
        """))
        events = _rowdicts(conn.execute("""
            SELECT *
            FROM autoloader_events
            ORDER BY id DESC
            LIMIT 200
        """))

    return {
        "mode": "SIMULATION_ONLY",
        "plans": plans,
        "events": events,
    }


# ============================================================
# WEBSOCKET
# ============================================================
@app.websocket("/ws")
async def ws_endpoint(websocket: WebSocket):
    await manager.connect(websocket)

    try:
        while True:
            await websocket.receive_text()
    except WebSocketDisconnect:
        manager.disconnect(websocket)


# ============================================================
# END PART 5
# PART 6 CONTINUES WITH:
# - Full updated HTML/CSS/JS UI
# - Elgin Custom Mold Services branding
# - Hidden Efficiency tab unless admin
# - Job Matching PDF upload only
# - Calendar delete restored
# - Simplified QR scan workflow with quantity add/subtract
# - Checkout timing display
# - Tool inspection modal/tab
# - Root route and startup block
# ============================================================
# ============================================================
# UI
# ============================================================
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Elgin Custom Mold Services</title>
<link href="https://fonts.googleapis.com/css2?family=Share+Tech+Mono&family=Barlow+Condensed:wght@400;600;700;800;900&display=swap" rel="stylesheet">
<script src="https://unpkg.com/html5-qrcode"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script src="https://cdn.jsdelivr.net/npm/qrious@4.0.2/dist/qrious.min.js"></script>

<style>
:root{
  --bg:#05060a;
  --panel:#0c0d12;
  --panel2:#101218;
  --border:#20222e;
  --text:#c8ccd8;
  --dim:#596078;
  --blue:#3d7fff;
  --cyan:#00e5c8;
  --red:#ff3b5c;
  --amber:#ffb627;
  --green:#00e676;
  --purple:#b06cff;
  --paper:#f4efd8;
  --paperline:#a6b3a2;
  --ink:#111;
  --fm:'Share Tech Mono',monospace;
  --fu:'Barlow Condensed',sans-serif
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:var(--fu);min-height:100vh}
.app{max-width:1640px;margin:0 auto;padding:18px 16px 60px}
header{display:flex;justify-content:space-between;align-items:center;padding-bottom:14px;border-bottom:1px solid var(--border);padding-left:160px}
.logo{font-size:1.55rem;font-weight:900;color:var(--blue);font-style:italic;letter-spacing:.06em}
.logo span{color:var(--cyan)}
.tab-bar{display:flex;border-bottom:1px solid var(--border);margin:16px 0 22px;overflow-x:auto;padding-left:150px}
.tab-btn{padding:13px 20px;font-weight:800;font-size:.7rem;letter-spacing:.13em;color:var(--dim);cursor:pointer;border-bottom:2px solid transparent;white-space:nowrap}
.tab-btn.active{color:var(--cyan);border-bottom-color:var(--cyan)}
.tab-content{display:none}
.tab-content.active{display:block}
.panel{background:var(--panel);border:1px solid var(--border);border-radius:7px}
.panel2{background:var(--panel2);border:1px solid var(--border);border-radius:6px}
.phdr{font-size:.62rem;font-weight:800;letter-spacing:.14em;color:var(--dim);padding:12px 16px 10px;border-bottom:1px solid var(--border);text-transform:uppercase}
.lbl{font-size:.58rem;font-weight:800;letter-spacing:.12em;color:var(--dim);text-transform:uppercase}
.mono{font-family:var(--fm)}
.num-xl{font-family:var(--fm);font-size:2rem;line-height:1}
.num-md{font-family:var(--fm);font-size:.95rem}
.inp,.sel,.txtarea{width:100%;background:rgba(0,0,0,.6);border:1px solid var(--border);padding:10px 12px;border-radius:4px;color:var(--text);outline:none}
.txtarea{min-height:72px;resize:vertical;font-family:var(--fm);font-size:.78rem}
.inp{font-family:var(--fm);font-size:.82rem}
.inp-g{background:#d4edda;color:#000}
.btn-p{background:var(--blue);color:#fff;padding:10px 16px;border-radius:4px;font-weight:800;font-size:.68rem;letter-spacing:.1em;cursor:pointer;border:none;text-decoration:none;display:inline-block}
.btn-g{background:transparent;color:var(--cyan);border:1px solid var(--cyan);padding:8px 12px;border-radius:4px;font-weight:800;font-size:.62rem;cursor:pointer;text-decoration:none;display:inline-block}
.btn-r{background:transparent;color:var(--red);border:1px solid var(--red);padding:8px 12px;border-radius:4px;font-weight:800;font-size:.62rem;cursor:pointer;text-decoration:none;display:inline-block}
.btn-y{background:transparent;color:var(--amber);border:1px solid var(--amber);padding:8px 12px;border-radius:4px;font-weight:800;font-size:.62rem;cursor:pointer;text-decoration:none;display:inline-block}
.itable{width:100%;border-collapse:collapse;font-size:.82rem}
.itable th{font-size:.58rem;font-weight:800;letter-spacing:.12em;color:var(--dim);padding:10px 14px;text-align:left;border-bottom:1px solid var(--border)}
.itable td{padding:10px 14px;border-bottom:1px solid rgba(255,255,255,.035);vertical-align:top}
.itable tr:hover td{background:rgba(255,255,255,.018)}
.pill{font-size:.58rem;font-weight:800;padding:2px 9px;border-radius:2px;letter-spacing:.07em;text-transform:uppercase;display:inline-block}
.pill-ok{background:rgba(0,230,118,.08);color:var(--green);border:1px solid rgba(0,230,118,.18)}
.pill-warn{background:rgba(255,182,39,.08);color:var(--amber);border:1px solid rgba(255,182,39,.18)}
.pill-crit{background:rgba(255,59,92,.08);color:var(--red);border:1px solid rgba(255,59,92,.18)}
.pill-blue{background:rgba(61,127,255,.08);color:var(--blue);border:1px solid rgba(61,127,255,.2)}
.pill-dim{background:rgba(255,255,255,.04);color:var(--dim);border:1px solid var(--border)}
.row{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.col{display:flex;flex-direction:column;gap:8px}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
.grid4{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
.sbox{background:rgba(0,0,0,.3);border:1px solid var(--border);border-radius:4px;padding:8px 11px}
.dot{width:6px;height:6px;border-radius:50%;display:inline-block}
.dot-on{background:var(--green)}
.dot-off{background:var(--red)}
.dot-warn{background:var(--amber)}
.prog-bar{height:8px;background:rgba(255,255,255,.05);border-radius:4px;overflow:hidden}
.prog-fill{height:100%;border-radius:4px;transition:width .5s}
.health-bar{height:14px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.08);border-radius:999px;overflow:hidden}
.health-fill{height:100%;border-radius:999px;transition:width .35s ease}
.notice{padding:10px 12px;border:1px solid rgba(255,182,39,.25);background:rgba(255,182,39,.08);color:var(--amber);border-radius:4px;margin-top:10px;font-family:var(--fm);font-size:.72rem}
.goodnote{padding:10px 12px;border:1px solid rgba(0,230,118,.25);background:rgba(0,230,118,.06);color:var(--green);border-radius:4px;margin-top:10px;font-family:var(--fm);font-size:.72rem}
.dangernote{padding:10px 12px;border:1px solid rgba(255,59,92,.25);background:rgba(255,59,92,.06);color:var(--red);border-radius:4px;margin-top:10px;font-family:var(--fm);font-size:.72rem}
.hidden{display:none!important}
a{color:var(--blue)}

.admin-corner{position:fixed;top:8px;left:8px;z-index:3000;width:150px;background:rgba(5,6,10,.94);border:1px solid var(--border);border-radius:7px;padding:7px;box-shadow:0 8px 28px rgba(0,0,0,.45);font-family:var(--fm)}
.admin-corner-title{font-size:.56rem;color:var(--dim);letter-spacing:.12em}
.admin-corner-status{font-size:.65rem;margin:3px 0 5px;font-weight:900}
.admin-corner-login{display:flex;gap:4px}
.admin-corner-login input{width:86px;background:#000;border:1px solid var(--border);color:var(--text);font-size:.58rem;padding:4px;border-radius:3px}
.admin-corner-login button,.corner-logout{background:transparent;color:var(--cyan);border:1px solid var(--cyan);font-size:.52rem;padding:4px 5px;border-radius:3px;cursor:pointer}
.corner-logout{width:100%;margin-top:4px;color:var(--red);border-color:var(--red)}

.mcard{background:var(--panel);border:1px solid var(--border);border-radius:7px;overflow:hidden}
.mcard.online{border-left:3px solid var(--cyan)}
.mcard.offline{border-left:3px solid var(--red);opacity:.65}
.mcard.alarm{border-left:3px solid var(--amber)}

.calendar-layout{display:grid;grid-template-columns:minmax(0,1.55fr) minmax(380px,.45fr);gap:20px;align-items:start}
.cal{display:grid;grid-template-columns:repeat(7,minmax(135px,1fr));gap:10px;width:100%;overflow-x:auto;padding-bottom:4px}
.cal-head{text-align:center;font-family:var(--fm);font-size:.7rem;color:var(--cyan);padding:8px 0;background:rgba(0,229,200,.05);border:1px solid rgba(0,229,200,.12);border-radius:5px;letter-spacing:.08em}
.day{min-height:178px;height:178px;background:linear-gradient(180deg,rgba(255,255,255,.04),rgba(255,255,255,.018));border:1px solid var(--border);border-radius:8px;padding:8px;overflow:hidden;display:flex;flex-direction:column}
.day.blank{visibility:hidden;border:none;background:transparent}
.day.today{border-color:var(--cyan);box-shadow:0 0 0 1px rgba(0,229,200,.22) inset,0 0 22px rgba(0,229,200,.06)}
.daynum{font-family:var(--fm);font-size:.75rem;color:var(--dim);margin-bottom:7px;display:flex;justify-content:space-between;align-items:center;flex-shrink:0}
.daynum strong{font-size:.95rem;color:#fff}
.day-count{font-size:.55rem;color:var(--blue);border:1px solid rgba(61,127,255,.22);background:rgba(61,127,255,.08);padding:2px 6px;border-radius:999px}
.dayjobs{overflow:auto;padding-right:3px;display:flex;flex-direction:column;gap:7px}
.cal-job-card{display:block;border:1px solid rgba(255,255,255,.08);background:rgba(0,0,0,.28);border-radius:6px;padding:6px 7px;cursor:pointer;transition:.15s;position:relative}
.cal-job-card:hover{transform:translateY(-1px);border-color:var(--cyan);background:rgba(0,229,200,.055)}
.cal-job-card.overdue{border-color:rgba(255,59,92,.45);background:rgba(255,59,92,.07)}
.cal-job-card.soon{border-color:rgba(255,182,39,.38);background:rgba(255,182,39,.055)}
.cal-job-top{display:flex;justify-content:space-between;gap:8px;align-items:center}
.cal-job-num{font-family:var(--fm);color:var(--blue);font-weight:900;font-size:.75rem}
.cal-job-pct{font-family:var(--fm);font-size:.62rem;color:var(--cyan)}
.cal-job-sub{margin-top:2px;color:var(--dim);font-size:.58rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-family:var(--fm)}
.cal-mini-bar{height:5px;margin-top:5px;background:rgba(255,255,255,.07);border-radius:999px;overflow:hidden}
.cal-mini-bar span{display:block;height:100%;border-radius:999px}
.cal-del{font-size:.48rem;padding:2px 5px;margin-left:4px}

.paper-wrap{overflow:auto;border:1px solid var(--border);border-radius:6px;background:#16130e;padding:10px}
.paper-board{min-width:1450px;background-color:var(--paper);background-image:linear-gradient(var(--paperline) 1px,transparent 1px),linear-gradient(90deg,var(--paperline) 1px,transparent 1px);background-size:16px 16px;color:var(--ink);border:3px solid #111;font-family:"Comic Sans MS","Bradley Hand",cursive;box-shadow:0 8px 30px rgba(0,0,0,.35)}
.paper-board table{width:100%;border-collapse:collapse;table-layout:fixed}
.paper-board th,.paper-board td{border:2px solid #111;padding:4px 5px;font-size:.72rem;line-height:1.05;height:34px;vertical-align:middle;overflow:hidden}
.paper-board th{font-family:var(--fu);font-size:.65rem;font-weight:900;text-align:center;background:rgba(255,255,255,.28);color:#111;letter-spacing:.05em}
.paper-board .paper-title{font-family:var(--fu);font-size:1.15rem;font-weight:900;letter-spacing:.12em;text-align:center;background:rgba(255,246,160,.62)}
.paper-board input,.paper-board textarea{width:100%;border:none;background:rgba(255,255,255,.16);color:#111;font-family:"Comic Sans MS","Bradley Hand",cursive;font-weight:700;outline:none;padding:2px}
.paper-board textarea{resize:vertical;min-height:28px}
.paper-board .jobcell{font-family:var(--fm);font-size:.76rem;font-weight:900;color:#002a9f}

.tool-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(360px,1fr));gap:14px}
.tool-card{background:linear-gradient(180deg,rgba(255,255,255,.035),rgba(255,255,255,.012));border:1px solid var(--border);border-radius:7px;padding:14px;cursor:pointer;transition:transform .15s,border-color .15s}
.tool-card:hover{transform:translateY(-2px);border-color:var(--cyan)}
.tool-card.pending{border-color:var(--amber);box-shadow:0 0 0 1px rgba(255,182,39,.16) inset}
.tool-title{display:flex;justify-content:space-between;gap:8px;align-items:start;margin-bottom:10px}
.tool-no{font-family:var(--fm);font-size:1.2rem;color:var(--blue);font-weight:900}
.tool-name{font-weight:900;color:#fff;font-size:.95rem;line-height:1.05}

.modal-backdrop{display:none;position:fixed;inset:0;background:rgba(0,0,0,.72);z-index:2500;align-items:flex-start;justify-content:center;padding:32px 12px;overflow:auto}
.modal-backdrop.active{display:flex}
.modal{width:min(1320px,96vw);background:var(--panel);border:1px solid var(--border);border-radius:8px;box-shadow:0 20px 80px rgba(0,0,0,.7);overflow:hidden}
.modal-head{display:flex;justify-content:space-between;align-items:center;padding:14px 18px;border-bottom:1px solid var(--border)}
.modal-body{padding:16px}
.chart-box{height:380px;background:rgba(0,0,0,.28);border:1px solid var(--border);border-radius:6px;padding:12px}
.tool-tabs{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}
.tool-tab-btn{background:rgba(255,255,255,.035);border:1px solid var(--border);color:var(--dim);padding:8px 12px;border-radius:4px;font-weight:900;font-size:.62rem;letter-spacing:.1em;cursor:pointer}
.tool-tab-btn.active{border-color:var(--cyan);color:var(--cyan);background:rgba(0,229,200,.06)}
.tool-tab{display:none}
.tool-tab.active{display:block}

.label-sheet{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:10px;margin-top:12px}
.qr-label{background:#fff;color:#000;border:1px solid #ccc;border-radius:6px;padding:10px;display:flex;gap:10px;align-items:center;min-height:96px}
.qr-label canvas{width:78px!important;height:78px!important;flex:0 0 78px}
.qr-label-title{font-family:Arial,sans-serif;font-weight:900;font-size:18px;line-height:1.05}
.qr-label-desc{font-family:Arial,sans-serif;font-size:10px;line-height:1.1;margin-top:3px}
.qr-label-payload{font-family:monospace;font-size:9px;margin-top:4px;color:#555}
#barcode-reader{display:none;margin-top:12px;border:1px solid var(--blue);border-radius:5px;overflow:hidden;background:#000;max-width:520px}
#barcode-reader video{width:100%!important}
.qr-action-box{display:none;margin-top:12px}
.qr-action-box.active{display:block}

.shelf-grid{display:grid;grid-template-columns:repeat(4,minmax(120px,1fr));gap:10px;max-width:760px}
.shelf-slot{min-height:105px;border:2px solid var(--border);border-radius:8px;padding:10px;background:rgba(255,255,255,.03);cursor:pointer;transition:.15s}
.shelf-slot:hover{transform:translateY(-2px);border-color:var(--cyan)}
.shelf-slot.empty{opacity:.65}
.shelf-slot.loaded{border-color:var(--cyan);background:rgba(0,229,200,.08)}
.shelf-slot.reserved{border-color:var(--amber);background:rgba(255,182,39,.08)}
.shelf-slot.blocked{border-color:var(--red);background:rgba(255,59,92,.08)}
.shelf-slot.safe{box-shadow:0 0 0 2px rgba(0,230,118,.16) inset}
.shelf-code{font-family:var(--fm);font-size:1.1rem;color:#fff;font-weight:900}
.shelf-job{font-family:var(--fm);color:var(--blue);font-size:.82rem;margin-top:4px}

@media(max-width:1100px){
  header,.tab-bar{padding-left:0}
  .admin-corner{position:relative;width:auto;margin:8px}
  .grid2,.grid3,.grid4,.calendar-layout{grid-template-columns:1fr!important}
  .tool-grid{grid-template-columns:1fr}
  .cal{grid-template-columns:repeat(7,minmax(150px,1fr))}
  .day{min-height:170px}
  .shelf-grid{grid-template-columns:repeat(2,1fr)}
}
@media print{
  body *{visibility:hidden}
  #qr-label-print-area,#qr-label-print-area *{visibility:visible}
  #qr-label-print-area{position:absolute;left:0;top:0;width:100%;background:#fff;padding:12px}
  .qr-label{break-inside:avoid;page-break-inside:avoid}
}
</style>
</head>

<body>
<div id="admin-corner" class="admin-corner">
  <div class="admin-corner-title">ADMIN</div>
  <div id="admin-corner-status" class="admin-corner-status">LOCKED</div>
  <div id="admin-corner-login" class="admin-corner-login">
    <input id="corner-admin-pass" type="password" placeholder="password">
    <button onclick="cornerAdminLogin()">LOGIN</button>
  </div>
  <button id="corner-admin-logout" class="corner-logout" onclick="adminLogout()" style="display:none">LOGOUT</button>
</div>

<div class="app">
<header>
  <div class="logo">ELGIN <span>CUSTOM MOLD SERVICES</span></div>
  <div class="mono" id="sys-time">--:--:--</div>
</header>

<div class="tab-bar">
  <div class="tab-btn active" onclick="ST('fleet')" id="btn-fleet">FLEET</div>
  <div class="tab-btn" onclick="ST('jobs')" id="btn-jobs">JOBS</div>
  <div class="tab-btn" onclick="ST('sched')" id="btn-sched">APP CALENDAR</div>
  <div class="tab-btn" onclick="ST('match')" id="btn-match">JOB MATCHING</div>
  <div class="tab-btn" onclick="ST('eff')" id="btn-eff">EFFICIENCY / MONEY</div>
  <div class="tab-btn" onclick="ST('tools')" id="btn-tools">TOOLS</div>
  <div class="tab-btn" onclick="ST('inv')" id="btn-inv">INVENTORY / QR</div>
  <div class="tab-btn" onclick="ST('admin')" id="btn-admin">ADMIN</div>
</div>

<div class="tab-content active" id="tab-fleet">
  <div class="fleet-grid" id="fleet-grid" style="display:grid;grid-template-columns:1fr 1fr;gap:18px"></div>
</div>

<div class="tab-content" id="tab-jobs">
  <div class="panel" style="padding:14px;margin-bottom:18px">
    <div class="lbl" style="color:var(--cyan)">AUTO JOB DETECTION + SMART ESTIMATES</div>
    <div class="mono" style="font-size:.8rem;color:var(--dim);margin-top:5px">
      Program estimates use Q-code cycle timer, G-code analysis, then similar completed jobs with the same component.
      ID and OD are matched separately. Old reused programs can map to current matched jobs.
    </div>
  </div>
  <div id="job-series-board"></div>
  <div id="job-log-board" style="margin-top:20px"></div>
</div>

<div class="tab-content" id="tab-sched">
  <div class="panel" style="padding:18px;margin-bottom:18px">
    <div class="phdr" style="margin:-18px -18px 16px">ADD JOB TO APP CALENDAR</div>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:10px;align-items:end">
      <div class="col"><div class="lbl">Job Number</div><input id="sc-job" class="inp inp-g" placeholder="JXXXX"></div>
      <div class="col"><div class="lbl">Due Date</div><input id="sc-due" type="date" class="inp"></div>
      <div class="col"><div class="lbl">Notes / Description</div><input id="sc-notes" class="inp" placeholder="Optional"></div>
      <button class="btn-p" onclick="addSchedule()">+ ADD TO CALENDAR</button>
    </div>
    <div class="notice">Steel sheet PDFs are uploaded only under Job Matching. Calendar delete removes the job from the calendar only; it does not erase job history.</div>
    <div id="sched-msg" class="mono" style="margin-top:8px;color:var(--green)"></div>
  </div>

  <div class="calendar-layout" style="margin-bottom:20px">
    <div class="panel" style="padding:16px">
      <div style="display:flex;justify-content:space-between;align-items:center;margin:-16px -16px 14px;border-bottom:1px solid var(--border);padding:10px 12px">
        <button class="btn-g" onclick="changeCalMonth(-1)">◀ PREV</button>
        <div class="phdr" style="border:0;padding:0" id="cal-title">CALENDAR</div>
        <button class="btn-g" onclick="changeCalMonth(1)">NEXT ▶</button>
      </div>
      <div class="cal" id="calendar"></div>
    </div>

    <div class="panel" style="overflow:hidden;max-height:760px;overflow-y:auto">
      <div class="phdr">DUE DATE LIST / GROUPED SIMILAR JOBS</div>
      <table class="itable"><tbody id="sched-list"></tbody></table>
    </div>
  </div>

  <div class="panel printable" style="padding:16px;margin-bottom:18px">
    <div style="display:flex;justify-content:space-between;align-items:center;margin:-16px -16px 14px;border-bottom:1px solid var(--border);padding:10px 12px;gap:10px;flex-wrap:wrap">
      <div class="phdr" style="border:0;padding:0">SHOP JOB BOARD / PAPER CALENDAR</div>
      <div class="row">
        <button class="btn-g" onclick="addShopBoardRow()">+ ADD ROW</button>
        <button class="btn-g" onclick="syncShopBoard()">SYNC FROM JOBS</button>
      </div>
    </div>
    <div id="shop-board-msg" class="mono" style="margin:8px 0;color:var(--green)"></div>
    <div class="paper-wrap"><div id="shop-board" class="paper-board"></div></div>
  </div>
</div>

<div class="tab-content" id="tab-match">
  <div class="panel" style="padding:20px;margin-bottom:18px">
    <div class="phdr" style="margin:-20px -20px 18px">REVERSE LOOKUP + STEEL SHEET PDF</div>
    <div class="notice" style="margin-bottom:14px">
      Upload steel-sheet PDFs here only. This matching system is also used to map old/reused CNC programs to the current active job.
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:16px">
      <div class="col"><div class="lbl">Weight</div><input id="rl-w" class="inp inp-g" type="number" step="0.01"></div>
      <div class="col"><div class="lbl">CoM X</div><input id="rl-x" class="inp inp-g" type="number" step="0.01"></div>
      <div class="col"><div class="lbl">CoM Y</div><input id="rl-y" class="inp inp-g" type="number" step="0.01"></div>
      <div class="col"><div class="lbl">CoM Z</div><input id="rl-z" class="inp inp-g" type="number" step="0.01"></div>
    </div>
    <div class="grid2" style="margin-bottom:10px">
      <div class="col"><div class="lbl" style="color:var(--blue)">Job Number</div><input id="rl-job" class="inp" style="background:#cfe2ff;color:#000" placeholder="JXXXX"></div>
      <div class="col"><div class="lbl">Due Date</div><input id="rl-due" class="inp" type="date"></div>
    </div>
    <div class="grid2" style="margin-bottom:10px">
      <div class="col"><div class="lbl">Tolerance</div><input id="rl-tol" class="inp" type="number" step="0.001" value="0.01"></div>
      <div class="col"><div class="lbl">Steel Sheet PDF</div><input id="rl-pdf" class="inp" type="file" accept="application/pdf,.pdf"></div>
    </div>
    <div class="col" style="margin-bottom:14px"><div class="lbl">Notes</div><input id="rl-notes" class="inp" placeholder="Optional"></div>
    <button class="btn-p" style="width:100%" onclick="doLookup()">IDENTIFY PART + UPLOAD/PARSE STEEL PDF</button>
    <div id="rl-result" class="mono" style="margin-top:14px;color:var(--green);min-height:24px"></div>
  </div>

  <div class="panel" style="padding:20px;margin-bottom:18px;border-color:rgba(0,229,200,.35)">
    <div class="phdr" style="margin:-20px -20px 18px">CMS XT SIGNATURE MATCHES</div>
    <div class="notice" style="margin-bottom:14px">
      Run the SolidWorks macro first. It auto-uploads XT_Export_Job_Signature.csv into ELGIN. Then enter that job here to see previous jobs with matching TCP/BCP, holders, and pots by size, mass, and center/COG.
    </div>
    <div class="row" style="gap:10px;align-items:end;margin-bottom:12px">
      <div class="col" style="min-width:180px"><div class="lbl">Job Number</div><input id="sig-job" class="inp inp-g" placeholder="J8423"></div>
      <div class="col" style="min-width:110px"><div class="lbl">Limit</div><input id="sig-limit" class="inp" type="number" min="1" max="50" value="10"></div>
      <button class="btn-p" onclick="loadSignatureMatches()">SHOW SIGNATURE MATCHES</button>
    </div>
    <div id="sig-result" class="mono" style="min-height:32px;color:var(--dim)">No signature job loaded yet.</div>
    <div style="margin-top:16px"><div class="lbl" style="color:var(--cyan);margin-bottom:8px">RECENT AUTO-IMPORTED SIGNATURE MATCHES</div><div id="sig-recent">Loading recent signature matches...</div></div>
  </div>

  <div class="grid2">
    <div class="panel" style="overflow:hidden"><div class="phdr">MATCH HISTORY LOG</div><table class="itable"><thead><tr><th>TIME</th><th>SPECS</th><th>MATCHES</th></tr></thead><tbody id="mh-tbody"></tbody></table></div>
    <div class="panel" style="overflow:hidden"><div class="phdr">CATALOG</div><table class="itable"><thead><tr><th>JOB</th><th>DUE</th><th>W</th><th>X/Y/Z</th><th>NOTES</th></tr></thead><tbody id="cat-tbody"></tbody></table></div>
  </div>
</div>

<div class="tab-content" id="tab-eff">
  <div id="eff-content"></div>
</div>

<div class="tab-content" id="tab-tools">
  <div id="tool-list"></div>
</div>

<div class="tab-content" id="tab-inv">
  <div class="grid4" style="margin-bottom:18px">
    <div class="panel" style="padding:16px"><div class="lbl">Total SKUs</div><div class="num-xl" id="s-total" style="color:var(--blue)">-</div></div>
    <div class="panel" style="padding:16px"><div class="lbl">Low Stock</div><div class="num-xl" id="s-low" style="color:var(--amber)">-</div></div>
    <div class="panel" style="padding:16px"><div class="lbl">Out of Stock</div><div class="num-xl" id="s-out" style="color:var(--red)">-</div></div>
    <div class="panel" style="padding:16px"><div class="lbl">Status</div><div class="num-xl" id="s-status" style="color:var(--green);font-size:1.1rem">STABLE</div></div>
  </div>

  <div class="panel" style="padding:18px;margin-bottom:18px">
    <div class="phdr" style="margin:-18px -18px 16px">QR LABEL MAKER</div>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:10px;align-items:end">
      <div class="col"><div class="lbl">SKU / Tool Code</div><input id="qr-label-sku" class="inp inp-g" placeholder="T182"></div>
      <div class="col"><div class="lbl">Description</div><input id="qr-label-desc" class="inp" placeholder="0.25 GORILLA LONG 5FL"></div>
      <div class="col"><div class="lbl">QR Payload</div><select id="qr-label-style" class="sel"><option value="plain">Plain SKU</option><option value="sku_prefix">SKU:T182</option><option value="mcn_pipe">MCN|SKU|T182|DESC|Name</option></select></div>
      <button class="btn-p" onclick="makeQrLabels()">MAKE ONE QR LABEL</button>
      <button class="btn-g" onclick="printQrLabels()">PRINT LABEL</button>
    </div>
    <div class="notice">Each tool/SKU has one QR identity. You can print another physical copy if needed, but the system does not create multiple different QR codes for the same tool.</div>
    <div id="qr-label-print-area"><div id="qr-label-sheet" class="label-sheet"></div></div>
  </div>

  <div class="panel" style="padding:18px;margin-bottom:18px">
    <div class="phdr" style="margin:-18px -18px 16px">PHONE QR / BARCODE SCANNER</div>
    <div class="notice">
      Scan first, then choose ADD or SUBTRACT and quantity. On iPhone, live camera requires HTTPS or ngrok. If live camera does not work, use PHONE PHOTO SCAN.
    </div>
    <div class="row" style="margin-top:12px">
      <button class="btn-p" onclick="startBarcodeScanner()">START LIVE SCANNER</button>
      <button class="btn-g" onclick="openBarcodePhotoFallback()">PHONE PHOTO SCAN</button>
      <button class="btn-r" onclick="stopBarcodeScanner()">STOP SCANNER</button>
      <input id="manual-barcode" class="inp" style="max-width:260px" placeholder="Manual QR/barcode text">
      <button class="btn-g" onclick="manualBarcodeSubmit()">SCAN TEXT</button>
      <input id="barcode-file-input" type="file" accept="image/*" capture="environment" style="display:none" onchange="decodeBarcodePhoto(this)">
    </div>
    <div id="barcode-reader"></div>
    <div id="barcode-msg" class="mono" style="margin-top:8px;min-height:18px;color:var(--green)"></div>

    <div id="qr-action-box" class="panel2 qr-action-box" style="padding:14px">
      <div class="phdr" style="margin:-14px -14px 12px">SCANNED TOOL ACTION</div>
      <div id="qr-resolved-summary"></div>

      <div id="qr-known-actions" class="hidden" style="margin-top:12px">
        <div class="row">
          <div class="col" style="width:140px"><div class="lbl">Qty</div><input id="qr-action-qty" class="inp" type="number" value="1" min="1"></div>
          <button class="btn-r" onclick="qrTakeOut()">SUBTRACT / TAKE OUT</button>
          <button class="btn-g" onclick="qrAddIn()">ADD / RESTOCK</button>
        </div>
      </div>

      <div id="qr-create-actions" class="hidden" style="margin-top:12px">
        <div class="notice">This QR is not mapped to an inventory tool yet. Enter the tool info, then choose the action and quantity.</div>
        <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:10px;align-items:end;margin-top:10px">
          <div class="col"><div class="lbl">SKU / Tool #</div><input id="qr-create-sku" class="inp inp-g"></div>
          <div class="col"><div class="lbl">Tool Name</div><input id="qr-create-desc" class="inp inp-g"></div>
          <div class="col"><div class="lbl">Category</div><input id="qr-create-cat" class="inp" value="Tool"></div>
          <div class="col"><div class="lbl">Starting Qty</div><input id="qr-create-start" class="inp" type="number" value="0" min="0"></div>
          <div class="col"><div class="lbl">Action Qty</div><input id="qr-create-qty" class="inp" type="number" value="1" min="1"></div>
          <div class="col"><div class="lbl">Min Stock</div><input id="qr-create-min" class="inp" type="number" value="5" min="0"></div>
        </div>
        <div class="row" style="margin-top:10px">
          <button class="btn-r" onclick="qrCreateAndTakeOut()">CREATE + SUBTRACT</button>
          <button class="btn-g" onclick="qrCreateAndAdd()">CREATE + ADD</button>
        </div>
      </div>
    </div>
  </div>

  <div class="panel" style="padding:18px;margin-bottom:18px">
    <div class="phdr" style="margin:-18px -18px 16px">ADD / RENAME TOOL OR INVENTORY ITEM</div>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px;align-items:end">
      <div class="col"><div class="lbl">SKU</div><input id="new-sku" class="inp inp-g" placeholder="T182"></div>
      <div class="col"><div class="lbl">Tool / Item Name</div><input id="new-desc" class="inp inp-g" placeholder="0.25 Flat Endmill"></div>
      <div class="col"><div class="lbl">Category</div><input id="new-cat" class="inp" value="Tool"></div>
      <div class="col"><div class="lbl">Starting Qty</div><input id="new-qty" class="inp" type="number" value="0"></div>
      <div class="col"><div class="lbl">Min Stock</div><input id="new-min" class="inp" type="number" value="5"></div>
      <button class="btn-p" onclick="addInventoryItem()">+ ADD / UPDATE</button>
    </div>
    <div id="new-item-msg" class="mono" style="margin-top:8px;color:var(--green)"></div>
  </div>

  <div class="grid2">
    <div class="panel" style="padding:18px">
      <div class="phdr" style="margin:-18px -18px 16px">MANUAL INVENTORY POST</div>
      <div class="col">
        <input id="sku-input" class="inp" placeholder="SKU CODE">
        <input id="manual-qty" class="inp" type="number" value="1" min="1" placeholder="Quantity">
        <select id="dir-select" class="sel"><option value="OUT">SUBTRACT / CHECK OUT</option><option value="IN">ADD / RESTOCK</option></select>
        <button class="btn-p" onclick="submitScan()">POST ENTRY</button>
        <div id="scan-msg" class="mono"></div>
      </div>
    </div>
    <div class="panel" style="overflow:hidden">
      <div class="phdr">INVENTORY + CHECKOUT TIMING</div>
      <table class="itable">
        <thead>
          <tr><th>SKU</th><th>DESCRIPTION</th><th>QTY</th><th>7D</th><th>30D</th><th>AVG BETWEEN CHECKOUTS</th><th>LAST CHECKOUT</th><th>RUNWAY</th></tr>
        </thead>
        <tbody id="inv-tbody"></tbody>
      </table>
    </div>
  </div>
</div>

<div class="tab-content" id="tab-admin">
  <div class="panel" style="padding:18px;margin-bottom:18px">
    <div class="phdr" style="margin:-18px -18px 16px">ADMIN ONLY</div>
    <div id="admin-msg" class="mono" style="color:var(--dim)"></div>
    <div id="admin-locked-note" class="notice">
      Admin page content is hidden until login. Login using the admin box in the upper-left.
    </div>
  </div>
  <div id="admin-only-content"></div>
</div>
</div>

<div id="tool-modal-backdrop" class="modal-backdrop" onclick="if(event.target.id==='tool-modal-backdrop')closeToolModal()">
  <div class="modal">
    <div class="modal-head">
      <div>
        <div class="lbl">TOOL DETAIL</div>
        <div id="tool-modal-title" style="font-size:1.25rem;font-weight:900;color:#fff">Tool</div>
      </div>
      <button class="btn-r" onclick="closeToolModal()">CLOSE</button>
    </div>
    <div class="modal-body"><div id="tool-modal-content"></div></div>
  </div>
</div>

<script>
let D=null,activeTab='fleet';
const TABS=['fleet','jobs','sched','match','eff','tools','inv','admin'];
let expandedJobs={}, barcodeScanner=null, lastBarcode='', lastBarcodeAt=0;
let calOffset=0, toolChart=null, currentToolDetail=null, currentToolTab='overview';
let lastResolvedQR=null;

function ST(id){
  if(id==='eff' && !adminOK()){
    activeTab='fleet';
  }else{
    activeTab=id;
  }
  TABS.forEach(t=>{
    const tab=document.getElementById('tab-'+t);
    const btn=document.getElementById('btn-'+t);
    if(tab)tab.classList.toggle('active',t===activeTab);
    if(btn)btn.classList.toggle('active',t===activeTab);
  });
  renderData(D);
}

async function fetchData(){
  try{
    const r=await fetch('/api/dashboard',{credentials:'same-origin'});
    D=await r.json();
    renderData(D);
  }catch(e){console.log(e);}
}

const wsScheme=location.protocol==='https:'?'wss':'ws';
try{
  const ws=new WebSocket(wsScheme+'://'+location.host+'/ws');
  ws.onmessage=()=>fetchData();
}catch(e){console.log(e);}

setInterval(fetchData,4000);
setInterval(()=>{document.getElementById('sys-time').textContent=new Date().toLocaleTimeString('en-US',{hour12:false});},1000);
fetchData();

function adminOK(){return !!(D&&D.admin&&D.admin.authenticated);}
function esc(v){return String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));}
function fsec(s){s=Math.round(s||0);const h=Math.floor(s/3600),m=Math.floor((s%3600)/60),sc=s%60;if(s<=0)return'UNKNOWN';return h>0?`${h}h ${m}m`:`${m}m ${String(sc).padStart(2,'0')}s`;}
function fmtHours(h){h=Number(h||0);if(!h||h<0)return'—';if(h<1)return Math.round(h*60)+'m';if(h<48)return h.toFixed(1)+'h';return (h/24).toFixed(1)+'d';}
function fmtDateTime(s){if(!s)return'—';return String(s).replace('T',' ').slice(0,19);}
function dollars(v){return '$'+Number(v||0).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});}
function numVal(v){return Number(v||0).toFixed(2);}
function hcolor(score){score=Number(score||0);return score>=85?'var(--red)':score>=65?'var(--amber)':'var(--green)';}
function dueColor(d){if(!d)return'var(--dim)';const t=new Date();t.setHours(0,0,0,0);const due=new Date(d+'T00:00:00');const diff=Math.ceil((due-t)/86400000);if(diff<0)return'var(--red)';if(diff<=7)return'var(--amber)';if(diff<=14)return'var(--cyan)';return'var(--green)';}
function daysLeft(d){if(!d)return'No due date';const today=new Date();today.setHours(0,0,0,0);const due=new Date(d+'T00:00:00');const diff=Math.ceil((due-today)/86400000);if(diff<0)return`Overdue ${Math.abs(diff)}d`;if(diff===0)return'Due today';return`${diff}d left`;}
function progFor(job, comps){const arr=(comps||[]).filter(c=>c.job_num===job);const done=arr.filter(c=>c.status==='COMPLETE').length;return{done,total:arr.length||0,pct:arr.length?Math.round(done/arr.length*100):0};}
function scl(s){if(!s||s==='OFFLINE')return'offline';const u=String(s).toUpperCase();if(['NO ALARM','NO ALARMS','ALARM: 0','ALARMS: 0','0 ALARM','0 ALARMS'].some(x=>u.includes(x)))return'online';if(u.includes('E-STOP')||u.includes('ESTOP')||u.includes('EMERGENCY')||u.includes('ALARM')||u.includes('FAULT'))return'alarm';return'online';}
function sdot(s){return scl(s)==='online'?'dot-on':scl(s)==='alarm'?'dot-warn':'dot-off';}
function scol(s){return scl(s)==='online'?'var(--cyan)':scl(s)==='alarm'?'var(--amber)':'var(--red)';}

function renderData(d){
  if(!d)return;
  renderAdminCorner();
  updateAdminOnlyTabs();

  if(activeTab==='eff' && !adminOK()){
    ST('fleet');
    return;
  }

  if(activeTab==='fleet')renderFleet(d.machines);
  if(activeTab==='jobs'){renderJobSeries(d.job_series||[],d.job_components||[]);renderJobLog(d.job_status_log||[]);}
  if(activeTab==='sched')renderSchedule(d.job_series||[],d.job_components||[],d.shop_board_rows||[]);
  if(activeTab==='match'){renderMH(d.match_history||[]);renderCat(d.catalog||[]);renderSignatureRecent(d.signature_recent_matches||[],d.signature_sync||{});}
  if(activeTab==='eff')renderEff(d.eff_day,d.eff_month,d.eff_year,d.machines,d.job_series||[],d.job_components||[]);
  if(activeTab==='tools')renderTools(d.tools||[]);
  if(activeTab==='inv')renderInv(d.inventory||[],d.summary||{});
  if(activeTab==='admin')renderAdmin(d);
}

function updateAdminOnlyTabs(){
  const effBtn=document.getElementById('btn-eff');
  const effTab=document.getElementById('tab-eff');
  if(!effBtn||!effTab)return;
  if(adminOK()){
    effBtn.style.display='block';
  }else{
    effBtn.style.display='none';
    effTab.classList.remove('active');
    if(activeTab==='eff')activeTab='fleet';
  }
}

function renderAdminCorner(){
  const status=document.getElementById('admin-corner-status');
  const login=document.getElementById('admin-corner-login');
  const logout=document.getElementById('corner-admin-logout');
  if(!status||!login||!logout)return;
  if(adminOK()){
    status.textContent='UNLOCKED';
    status.style.color='var(--green)';
    login.style.display='none';
    logout.style.display='block';
  }else{
    status.textContent='LOCKED';
    status.style.color='var(--red)';
    login.style.display='flex';
    logout.style.display='none';
  }
}

async function cornerAdminLogin(){
  const pass=document.getElementById('corner-admin-pass').value;
  const r=await fetch('/api/admin/login',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({password:pass})});
  if(r.ok){document.getElementById('corner-admin-pass').value='';fetchData();}else alert('Bad admin password.');
}
async function adminLogout(){await fetch('/api/admin/logout',{method:'POST',credentials:'same-origin'});fetchData();}

function renderFleet(machines){
  document.getElementById('fleet-grid').innerHTML=Object.entries(machines||{}).map(([name,m])=>{
    const sc=scl(m.status),lc=m.load>90?'var(--red)':m.load>70?'var(--amber)':'var(--cyan)';
    const toolLabel=String(m.tool||'0')==='0'?'T—':'T'+m.tool;
    const progBlock=Number(m.program_est_sec||0)>0 ? `
      <div style="padding:10px 16px;border-top:1px solid rgba(255,255,255,.04)">
        <div style="display:flex;justify-content:space-between;gap:10px;align-items:center">
          <div>
            <div class="lbl">PROGRAM PROGRESS</div>
            <div class="mono" style="color:var(--cyan)">${Number(m.program_progress_pct||0).toFixed(1)}% COMPLETE</div>
          </div>
          <div style="text-align:right">
            <div class="lbl">REMAINING</div>
            <div class="mono" style="color:var(--amber)">${fsec(m.program_remaining_sec||0)}</div>
            <div class="lbl">FINISH ${esc(m.program_est_finish||'')}</div>
          </div>
        </div>
        <div class="prog-bar" style="margin-top:6px">
          <div class="prog-fill" style="width:${Math.min(Number(m.program_progress_pct||0),100)}%;background:${Number(m.program_progress_pct||0)>=100?'var(--green)':'var(--cyan)'}"></div>
        </div>
        <div class="lbl" style="margin-top:4px">EST TOTAL ${fsec(m.program_est_sec||0)} · ${esc(m.program_progress_source||'')}</div>
      </div>
    ` : `
      <div style="padding:8px 16px;border-top:1px solid rgba(255,255,255,.04)">
        <div class="lbl" style="color:var(--dim)">PROGRAM PROGRESS: NO Q-CODE / G-CODE / SIMILAR-JOB ESTIMATE YET</div>
      </div>
    `;
    return`<div class="mcard ${sc}">
      <div style="padding:14px 16px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between">
        <div>
          <div class="lbl">${name}</div>
          <div class="mono" style="font-size:1.05rem;color:#fff">${esc(m.prog_display||m.prog)}</div>
          ${m.job_num?`<div class="lbl" style="color:var(--blue)">JOB ${m.job_num}${m.part_name?' · '+esc(m.part_name):''}</div>`:''}
          <div class="lbl">ACTUAL CUTTING: <span style="color:${m.actual_cutting?'var(--green)':'var(--dim)'}">${m.actual_cutting?'YES':'NO'}</span></div>
          <div class="lbl">DATA: <span style="color:${m.mtconnect_ok?'var(--green)':'var(--amber)'}">${esc(m.data_source||'UNKNOWN')}</span></div>
          ${m.mtconnect_error?`<div class="lbl" style="color:var(--red);max-width:420px">MT: ${esc(m.mtconnect_error)}</div>`:''}
        </div>
        <div style="text-align:right">
          <span class="dot ${sdot(m.status)}"></span>
          <span style="font-size:.62rem;font-weight:800;color:${scol(m.status)}">${esc(m.status)}</span>
          <div class="mono" style="font-size:.68rem;color:var(--dim)">${toolLabel}</div>
          <div class="lbl" style="color:var(--cyan);max-width:180px;text-align:right">${esc(m.tool_name||'')}</div>
        </div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(4,1fr);border-bottom:1px solid var(--border)">
        <div style="padding:10px 14px"><div class="lbl">LOAD</div><div class="num-md" style="color:${lc}">${Number(m.load||0).toFixed(1)}%</div></div>
        <div style="padding:10px 14px"><div class="lbl">SPINDLE</div><div class="num-md" style="color:${m.spindle_rpm>0?'var(--blue)':'var(--amber)'}">${Math.round(m.spindle_rpm||0)} RPM</div><div class="lbl" style="font-size:.46rem">${esc(m.rpm_source||'')}</div></div>
        <div style="padding:10px 14px"><div class="lbl">FEED</div><div class="num-md">${Math.round(m.feed_rate||0)} IPM</div></div>
        <div style="padding:10px 14px"><div class="lbl">PART COUNT</div><div class="num-md">${m.part_count||0}</div></div>
      </div>
      <div style="padding:10px 16px"><div class="prog-bar"><div class="prog-fill" style="width:${Math.min(m.load||0,100)}%;background:${lc}"></div></div></div>
      ${progBlock}
      <div style="display:flex;justify-content:space-between;padding:8px 16px;font-family:var(--fm);font-size:.62rem;color:var(--dim);gap:8px;flex-wrap:wrap">
        <span>UP ${esc(m.uptime)}</span><span>CYC ${esc(m.cycle)}</span><span>L ${Number(m.l_wear||0).toFixed(3)} D ${Number(m.d_wear||0).toFixed(3)}</span>
      </div>
    </div>`;
  }).join('');
}

function renderJobSeries(series, comps){
  const el=document.getElementById('job-series-board');
  if(!series.length){el.innerHTML='<div class="panel" style="padding:40px;text-align:center;color:var(--dim)">No job series yet.</div>';return;}
  const used=new Set();
  function isComplete(j){const arr=comps.filter(c=>c.job_num===j.job_num);return arr.length&&arr.every(c=>c.status==='COMPLETE');}
  function isRunning(j){return comps.some(c=>c.job_num===j.job_num&&c.status==='RUNNING');}
  function isOverdue(j){if(!j.due_date)return false;const today=new Date();today.setHours(0,0,0,0);const due=new Date(j.due_date+'T00:00:00');return due<today&&!isComplete(j);}
  function take(title,pred){const arr=series.filter(j=>!used.has(j.job_num)&&pred(j));arr.forEach(j=>used.add(j.job_num));return sectionJobs(title,arr,comps);}
  el.innerHTML=take('RUNNING NOW',isRunning)+take('OVERDUE',isOverdue)+take('SCHEDULED JOBS',j=>j.due_date&&!isComplete(j))+take('NO DUE DATE / AUTO-DETECTED',j=>!j.due_date&&!isComplete(j))+take('COMPLETE',isComplete);
}
function sectionJobs(title, series, comps){if(!series.length)return`<div class="lbl" style="margin:12px 0;color:var(--dim)">${title}: NONE</div>`;return`<div class="lbl" style="margin:12px 0;color:var(--cyan);font-size:.8rem">${title}</div>`+series.map(j=>jobCard(j,comps)).join('');}
function jobCard(j, comps){
  if(expandedJobs[j.job_num]===undefined)expandedJobs[j.job_num]=false;
  const p=progFor(j.job_num,comps);
  const arr=comps.filter(c=>c.job_num===j.job_num);
  const running=arr.find(c=>c.status==='RUNNING');
  return`<div class="panel" style="margin-bottom:16px;overflow:hidden">
    <div onclick="expandedJobs['${j.job_num}']=!expandedJobs['${j.job_num}'];renderData(D)" style="cursor:pointer;padding:14px 18px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap">
      <div>
        <div style="font-size:1.4rem;font-weight:900;color:var(--blue)">${expandedJobs[j.job_num]?'▼':'▶'} J${j.job_num}</div>
        <div class="lbl">Due <span style="color:${dueColor(j.due_date)}">${j.due_date||'—'} · ${daysLeft(j.due_date)}</span></div>
        <div class="row" style="margin-top:6px" onclick="event.stopPropagation()">
          <input id="jobcard-due-${j.job_num}" class="inp" type="date" value="${esc(j.due_date||'')}" style="max-width:150px;padding:6px 8px">
          <button class="btn-g" onclick="saveJobDueDateFromCard('${j.job_num}')">SAVE DATE</button>
        </div>
        ${j.steel_sheet_file?`<div class="lbl" style="color:var(--green)">STEEL PDF ATTACHED</div>`:''}
        ${j.matched_job_num?`<div class="lbl" style="color:var(--purple)">MATCHED FAMILY ROOT J${j.matched_job_num}</div>`:''}
        ${running?`<div class="lbl" style="color:var(--cyan)">RUNNING: ${esc(running.component_name)} on ${esc(running.last_seen_machine||running.machine||'')}</div>`:''}
      </div>
      <div style="text-align:right">
        <div class="mono" style="color:var(--cyan);font-size:1rem">${p.pct}%</div><div class="lbl">Complete</div>
        ${adminOK()?`<div class="row" style="justify-content:flex-end;margin-top:8px"><button class="btn-y" onclick="event.stopPropagation();renameJobAdmin('${j.job_num}')">RENAME</button><button class="btn-r" onclick="event.stopPropagation();deleteJobAdmin('${j.job_num}')">DELETE JOB</button></div>`:''}
      </div>
    </div>
    <div style="padding:14px 18px">
      <div class="lbl">Operation Progress ${p.done}/${p.total} · ${p.pct}%</div>
      <div class="prog-bar" style="margin:5px 0 10px"><div class="prog-fill" style="width:${p.pct}%;background:${p.pct>=100?'var(--green)':'var(--cyan)'}"></div></div>
      <div class="${expandedJobs[j.job_num]?'':'hidden'}">
        <div class="notice">Estimate priority: Q-code cycle timer → G-code → similar completed same-component job → unknown. ID and OD are never mixed.</div>
        <table class="itable"><thead><tr><th>COMPONENT</th><th>MACHINE</th><th>STATUS</th><th>RUNTIME</th><th>ESTIMATE</th><th>SOURCE</th><th>NOTE</th><th>ACTION</th></tr></thead><tbody>${arr.map(c=>{
          const est=Number(c.estimated_sec||0), pct=est>0?Math.max(0,Math.min(100,Math.round(((c.actual_sec||0)/est)*100))):0;
          return `<tr><td style="font-weight:800">${esc(c.component_name)}</td><td>${esc(c.last_seen_machine||c.machine||'—')}</td><td><span class="pill ${c.status==='COMPLETE'?'pill-ok':c.status==='RUNNING'?'pill-blue':'pill-dim'}">${c.status}</span></td><td><div class="lbl">${fsec(c.actual_sec||0)}</div>${est>0?`<div class="prog-bar" style="margin-top:4px"><div class="prog-fill" style="width:${pct}%;background:${pct>=100?'var(--amber)':'var(--cyan)'}"></div></div>`:''}</td><td>${est>0?fsec(est):'UNKNOWN'}</td><td><span class="pill ${c.estimate_source?'pill-blue':'pill-dim'}">${esc(c.estimate_source||'NONE')}</span><div class="lbl">${esc(c.estimate_basis||'')}</div></td><td class="mono" style="font-size:.62rem;color:var(--dim)">${esc(c.last_status_note||'')}</td><td><button class="btn-g" onclick="event.stopPropagation();compStatus('${j.job_num}','${encodeURIComponent(c.component_name)}','COMPLETE')">Done</button><button class="btn-r" onclick="event.stopPropagation();compStatus('${j.job_num}','${encodeURIComponent(c.component_name)}','PENDING')">Undo</button></td></tr>`;
        }).join('')}</tbody></table>
      </div>
    </div>
  </div>`;
}
async function compStatus(j,c,s){await fetch(`/api/job-series/${j}/components/${c}/status`,{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({status:s})});fetchData();}
function renderJobLog(logs){document.getElementById('job-log-board').innerHTML=`<div class="panel" style="overflow:hidden"><div class="phdr">STATUS UPDATE LOG</div><table class="itable"><thead><tr><th>TIME</th><th>JOB</th><th>COMPONENT</th><th>STATUS</th><th>MACHINE</th><th>MESSAGE</th></tr></thead><tbody>${logs.slice(0,80).map(l=>`<tr><td class="mono">${esc(l.timestamp||'')}</td><td>J${l.job_num}</td><td>${esc(l.component_name||'')}</td><td>${esc(l.status||'')}</td><td>${esc(l.machine||'')}</td><td>${esc(l.message||'')}</td></tr>`).join('')}</tbody></table></div>`;}

function jobFamilyRoot(job, allJobs){const jn=String(job.job_num||''),matched=String(job.matched_job_num||'');if(matched)return matched;const hasChildren=(allJobs||[]).some(j=>String(j.matched_job_num||'')===jn);return hasChildren?jn:jn;}
function groupJobsByFamily(jobs, allJobs){const groups={};jobs.forEach(j=>{const root=jobFamilyRoot(j,allJobs);(groups[root]=groups[root]||[]).push(j);});return Object.entries(groups).map(([root,items])=>({root,items:items.sort((a,b)=>String(a.job_num).localeCompare(String(b.job_num),undefined,{numeric:true})),related:items.length>1||items.some(x=>x.matched_job_num)})).sort((a,b)=>String(a.root).localeCompare(String(b.root),undefined,{numeric:true}));}
function jobDueClass(j){if(!j.due_date)return'';const today=new Date();today.setHours(0,0,0,0);const due=new Date(j.due_date+'T00:00:00');const diff=Math.ceil((due-today)/86400000);if(diff<0)return'overdue';if(diff<=7)return'soon';return'';}
function jobPrimaryNote(j){const bits=[];if(j.notes)bits.push(j.notes);if(j.matched_job_num)bits.push('Similar/root J'+j.matched_job_num);if(j.steel_sheet_file)bits.push('Steel PDF attached');if(j.matched_reason)bits.push(String(j.matched_reason).replaceAll('_',' '));return bits.join(' · ')||'Scheduled job';}
function changeCalMonth(n){calOffset+=n;if(D)renderSchedule(D.job_series||[],D.job_components||[],D.shop_board_rows||[]);}

function renderSchedule(series, comps, boardRows){
  renderCalendar(series,comps);renderShopBoard(boardRows||[]);
  const dated=(series||[]).filter(j=>j.due_date).sort((a,b)=>String(a.due_date).localeCompare(String(b.due_date))||String(a.job_num).localeCompare(String(b.job_num),undefined,{numeric:true}));
  const groups=groupJobsByFamily(dated,series||[]);
  document.getElementById('sched-list').innerHTML=groups.map(g=>`<tr><td colspan="6" style="padding:0;border-bottom:none"><div style="border-left:3px solid ${g.related?'var(--purple)':'var(--blue)'};padding-left:9px;margin:8px 8px 12px"><div class="mono" style="color:var(--cyan);font-size:.72rem;margin-bottom:6px">${g.related?`RELATED FAMILY J${g.root} · ${g.items.length} similar jobs`:`J${g.root}`}</div>${g.items.map(j=>{const p=progFor(j.job_num,comps);return`<div style="display:grid;grid-template-columns:70px 82px 1fr 86px;gap:8px;align-items:center;padding:6px 0;border-bottom:1px solid rgba(255,255,255,.04);font-size:.74rem"><div style="font-weight:900;color:var(--blue)">J${j.job_num}</div><div style="color:${dueColor(j.due_date)}">${j.due_date||'—'}</div><div><div style="font-size:.68rem;color:var(--dim)">${daysLeft(j.due_date)}</div><div class="prog-bar" style="margin-top:3px"><div class="prog-fill" style="width:${p.pct}%;background:${p.pct>=100?'var(--green)':'var(--cyan)'}"></div></div></div><div class="row" style="justify-content:flex-end"><span class="mono" style="color:var(--cyan);text-align:right">${p.pct}%</span>${adminOK()?`<button class="btn-r cal-del" onclick="deleteCalendarJob('${j.job_num}',event)">DEL</button>`:''}</div></div>`}).join('')}</div></td></tr>`).join('');
}
function renderCalendar(series,comps){
  const base=new Date();base.setMonth(base.getMonth()+calOffset);
  const y=base.getFullYear(),m=base.getMonth();
  document.getElementById('cal-title').textContent=base.toLocaleString('en-US',{month:'long',year:'numeric'}).toUpperCase();
  const first=new Date(y,m,1),start=first.getDay(),days=new Date(y,m+1,0).getDate();
  const today=new Date(),todayStr=`${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;
  let html='';['SUN','MON','TUE','WED','THU','FRI','SAT'].forEach(d=>html+=`<div class="cal-head">${d}</div>`);
  for(let i=0;i<start;i++)html+='<div class="day blank"></div>';
  for(let d=1;d<=days;d++){
    const ds=`${y}-${String(m+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    const jobs=(series||[]).filter(j=>j.due_date===ds).sort((a,b)=>String(a.job_num).localeCompare(String(b.job_num),undefined,{numeric:true}));
    let dayClass='';if(ds===todayStr)dayClass+=' today';if(jobs.length)dayClass+=' has-jobs';
    html+=`<div class="day${dayClass}"><div class="daynum"><strong>${d}</strong>${jobs.length?`<span class="day-count">${jobs.length} job${jobs.length===1?'':'s'}</span>`:''}${ds===todayStr?'<span style="color:var(--cyan);font-size:.55rem">TODAY</span>':''}</div><div class="dayjobs">${jobs.map(j=>{const p=progFor(j.job_num,comps),dc=jobDueClass(j),barColor=p.pct>=100?'var(--green)':dc==='overdue'?'var(--red)':dc==='soon'?'var(--amber)':'var(--cyan)',note=jobPrimaryNote(j);return`<div class="cal-job-card ${dc}" onclick="openJobFromCalendar('${j.job_num}')" title="${esc(note)}"><div class="cal-job-top"><span class="cal-job-num">J${j.job_num}</span><span>${adminOK()?`<button class="btn-r cal-del" onclick="deleteCalendarJob('${j.job_num}',event)">DEL</button>`:''}<span class="cal-job-pct">${p.pct}%</span></span></div><div class="cal-job-sub">${esc(note)}</div><div class="cal-mini-bar"><span style="width:${p.pct}%;background:${barColor}"></span></div></div>`}).join('')}</div></div>`;
  }
  document.getElementById('calendar').innerHTML=html;
}
function openJobFromCalendar(job){ST('jobs');expandedJobs[job]=true;renderData(D);setTimeout(()=>window.scrollTo({top:0,behavior:'smooth'}),100);}
async function deleteCalendarJob(job,event){
  if(event){event.stopPropagation();event.preventDefault();}
  if(!adminOK())return alert('Admin login required.');
  if(!confirm(`Remove J${job} from the app calendar?\n\nThis clears the due date/schedule but keeps the job history/components.`))return;
  const r=await fetch(`/api/calendar/jobs/${job}/delete`,{method:'POST',credentials:'same-origin'});
  const d=await r.json().catch(()=>({}));
  if(!r.ok){alert('Calendar delete failed: '+(d.detail||r.status));return;}
  document.getElementById('sched-msg').textContent=d.message||('Removed J'+job+' from calendar.');
  document.getElementById('sched-msg').style.color='var(--green)';
  fetchData();
}

function shopInput(row,field,w='70px'){return `<input style="width:${w}" id="sb-${row.id}-${field}" value="${esc(row[field]||'')}">`;}
function shopText(row,field,w='180px'){return `<textarea style="width:${w}" id="sb-${row.id}-${field}">${esc(row[field]||'')}</textarea>`;}
function renderShopBoard(rows){
  const el=document.getElementById('shop-board');if(!el)return;
  const body=rows.map(r=>`<tr><td class="jobcell">${shopInput(r,'job_num','72px')}</td><td>${shopText(r,'description','150px')}</td><td>${shopInput(r,'due_date','96px')}</td><td>${shopInput(r,'finals_mex','55px')}</td><td>${shopInput(r,'finals_john','55px')}</td><td>${shopInput(r,'steel','60px')}</td><td>${shopInput(r,'ht200','60px')}</td><td>${shopInput(r,'xt_sent','68px')}</td><td>${shopInput(r,'pyropel','70px')}</td><td>${shopInput(r,'j_blk','55px')}</td><td>${shopInput(r,'a2_1','50px')}</td><td>${shopInput(r,'req_1','50px')}</td><td>${shopInput(r,'d2','50px')}</td><td>${shopInput(r,'req_2','50px')}</td><td>${shopInput(r,'other','62px')}</td><td>${shopInput(r,'pullcore_cam','55px')}</td><td>${shopInput(r,'pullcore_key','55px')}</td><td>${shopInput(r,'h13','50px')}</td><td>${shopInput(r,'req_3','50px')}</td><td>${shopInput(r,'a2_2','50px')}</td><td>${shopInput(r,'req_4','50px')}</td><td>${shopInput(r,'pdf_mc','70px')}</td><td>${shopInput(r,'prog','70px')}</td><td>${shopText(r,'notes','190px')}</td><td><button class="btn-g" onclick="saveShopBoardRow(${r.id})">SAVE</button>${adminOK()?`<button class="btn-r" onclick="deleteShopBoardRow(${r.id})">DEL</button>`:''}</td></tr>`).join('');
  el.innerHTML=`<table><thead><tr><th colspan="3" class="paper-title">JOB CALENDAR</th><th colspan="2" class="paper-title">FINALS</th><th colspan="10" class="paper-title">SHOP FLOW</th><th colspan="2" class="paper-title">PULLCORE</th><th colspan="7" class="paper-title">REMAINING</th></tr><tr><th>JOB</th><th>DESCRIPTION</th><th>DUE</th><th>MEX</th><th>JOHN</th><th>STEEL</th><th>HT-200</th><th>XT/SENT</th><th>PYROPEL</th><th>J-BLK</th><th>A-2</th><th>REQ</th><th>D-2</th><th>REQ</th><th>OTHER</th><th>CAM</th><th>KEY</th><th>H-13</th><th>REQ</th><th>A-2</th><th>REQ</th><th>PDF/MC</th><th>PROG</th><th>NOTES</th><th>ACTION</th></tr></thead><tbody>${body||'<tr><td colspan="25">No board rows yet.</td></tr>'}</tbody></table>`;
}
function getShopVal(id,field){const el=document.getElementById(`sb-${id}-${field}`);return el?el.value:'';}
async function saveShopBoardRow(id){
  const body={job_num:getShopVal(id,'job_num'),description:getShopVal(id,'description'),due_date:getShopVal(id,'due_date'),priority:5,row_status:'OPEN',row_color:'',sort_order:0,finals_mex:getShopVal(id,'finals_mex'),finals_john:getShopVal(id,'finals_john'),steel:getShopVal(id,'steel'),ht200:getShopVal(id,'ht200'),xt_sent:getShopVal(id,'xt_sent'),pyropel:getShopVal(id,'pyropel'),j_blk:getShopVal(id,'j_blk'),a2_1:getShopVal(id,'a2_1'),req_1:getShopVal(id,'req_1'),d2:getShopVal(id,'d2'),req_2:getShopVal(id,'req_2'),other:getShopVal(id,'other'),pullcore_cam:getShopVal(id,'pullcore_cam'),pullcore_key:getShopVal(id,'pullcore_key'),h13:getShopVal(id,'h13'),req_3:getShopVal(id,'req_3'),a2_2:getShopVal(id,'a2_2'),req_4:getShopVal(id,'req_4'),pdf_mc:getShopVal(id,'pdf_mc'),prog:getShopVal(id,'prog'),notes:getShopVal(id,'notes')};
  const r=await fetch('/api/shop-board/'+id,{method:'PUT',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify(body)});
  const d=await r.json().catch(()=>({}));
  document.getElementById('shop-board-msg').textContent=r.ok?'Saved row '+id:'Save failed: '+(d.detail||r.status);
  document.getElementById('shop-board-msg').style.color=r.ok?'var(--green)':'var(--red)';
  fetchData();
}
async function deleteShopBoardRow(id){
  if(!adminOK())return alert('Admin login required.');
  if(!confirm('Delete this job calendar / shop-board row?'))return;
  const r=await fetch('/api/shop-board/'+id+'/delete',{method:'POST',credentials:'same-origin'});
  const d=await r.json().catch(()=>({}));
  if(!r.ok){alert('Delete failed: '+(d.detail||r.status));return;}
  document.getElementById('shop-board-msg').textContent='Deleted row '+id;
  document.getElementById('shop-board-msg').style.color='var(--green)';
  fetchData();
}
async function addShopBoardRow(){await fetch('/api/shop-board',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({})});fetchData();}
async function syncShopBoard(){const r=await fetch('/api/shop-board/sync-from-jobs',{method:'POST',credentials:'same-origin'});const d=await r.json();document.getElementById('shop-board-msg').textContent='Synced. Created '+(d.created||0)+' rows.';fetchData();}
async function addSchedule(){
  const job=document.getElementById('sc-job').value.trim();if(!job)return;
  const r=await fetch('/api/schedule',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({job_id:job,due_date:document.getElementById('sc-due').value,notes:document.getElementById('sc-notes').value})});
  const d=await r.json().catch(()=>({}));
  document.getElementById('sched-msg').textContent=r.ok?'Added '+job:'Add failed: '+(d.detail||'?');
  document.getElementById('sched-msg').style.color=r.ok?'var(--green)':'var(--red)';
  fetchData();
}

async function doLookup(){
  const form=new FormData();
  form.append('job_id',document.getElementById('rl-job').value);
  form.append('weight',document.getElementById('rl-w').value||0);
  form.append('com_x',document.getElementById('rl-x').value||0);
  form.append('com_y',document.getElementById('rl-y').value||0);
  form.append('com_z',document.getElementById('rl-z').value||0);
  form.append('tolerance',document.getElementById('rl-tol').value||0.01);
  form.append('due_date',document.getElementById('rl-due').value||'');
  form.append('notes',document.getElementById('rl-notes').value||'');
  const f=document.getElementById('rl-pdf').files[0];
  if(f)form.append('file',f);
  const msg=document.getElementById('rl-result');
  msg.textContent='Saving match and uploading PDF...';
  msg.style.color='var(--cyan)';
  const r=await fetch('/api/catalog/lookup-with-pdf',{method:'POST',credentials:'same-origin',body:form});
  const d=await r.json().catch(()=>({}));
  if(r.ok){
    const pdf=d.pdf_upload||{}, parse=pdf.parse||{};
    msg.textContent=`Saved J${d.job_id}. Matches: ${d.match_str}. Steel PDF: ${parse.message||pdf.message||'No PDF uploaded.'}`;
    msg.style.color='var(--green)';
    document.getElementById('rl-pdf').value='';
    fetchData();
  }else{
    msg.textContent='Error: '+(d.detail||'?');
    msg.style.color='var(--red)';
  }
}
function renderMH(data){document.getElementById('mh-tbody').innerHTML=data.map(h=>`<tr><td>${esc(h.timestamp||'')}</td><td>${esc(h.specs||'')}</td><td>${esc(h.matches||'')}</td></tr>`).join('');}
function renderCat(data){document.getElementById('cat-tbody').innerHTML=data.map(r=>`<tr><td style="color:var(--blue);font-weight:800">J${r.job_id}</td><td>${r.due_date||'—'}</td><td>${r.weight||0}</td><td>${r.com_x||0}/${r.com_y||0}/${r.com_z||0}</td><td>${esc(r.notes||'')}</td></tr>`).join('');}

function renderSignatureRecent(items,sync){
  const el=document.getElementById('sig-recent');
  if(!el)return;
  const syncNote=sync&&sync.folder?`<div style="color:var(--dim);margin-bottom:8px">Watching: ${esc(sync.folder)} · imported files ${sync.imported_files||0} · components ${sync.imported_components||0}</div>`:'';
  if(!items.length){el.innerHTML=syncNote+'<div style="color:var(--amber)">No signature files found yet. Run the SolidWorks macro and make sure it saves into the Matching software folder.</div>';return;}
  el.innerHTML=syncNote+items.map(it=>{
    const best=(it.matches||[])[0];
    if(!best)return `<div class="panel" style="padding:10px;margin:8px 0"><b>J${esc(it.job_num)}</b> <span style="color:var(--dim)">No previous signatures to compare yet.</span></div>`;
    return `<div class="panel" style="padding:10px;margin:8px 0;cursor:pointer;border-color:${signatureScoreColor(best.score)}" onclick="document.getElementById('sig-job').value='${esc(it.job_num)}';loadSignatureMatches();">
      <div style="display:flex;justify-content:space-between;gap:10px;align-items:center"><div><b style="color:var(--blue)">J${esc(it.job_num)}</b> best match <b>J${esc(best.job_num)}</b> <span style="color:var(--dim)">${best.matched_components||0}/6 parts</span></div><div style="color:${signatureScoreColor(best.score)};font-weight:900">${Number(best.score||0).toFixed(1)}%</div></div>
    </div>`;
  }).join('');
}

function signatureScoreColor(v){v=Number(v||0);return v>=92?'var(--green)':v>=80?'var(--amber)':'var(--red)';}
function sigDims(r){return `${Number(r?.length||0).toFixed(3)} x ${Number(r?.width||0).toFixed(3)} x ${Number(r?.thickness||0).toFixed(3)}`;}
function sigCenter(r){return `${Number(r?.center_x||0).toFixed(3)}, ${Number(r?.center_y||0).toFixed(3)}, ${Number(r?.center_z||0).toFixed(3)}`;}
function sigMass(r){return `${Number(r?.mass||0).toFixed(3)} lb`;}

function renderSignatureMatches(payload){
  const el=document.getElementById('sig-result');
  if(!el)return;
  if(!payload){el.innerHTML='<span style="color:var(--dim)">No signature job loaded yet.</span>';return;}
  if(!payload.has_signature){el.innerHTML=`<span style="color:var(--amber)">No signature has been uploaded for J${esc(payload.job_num||'')} yet. Run the SolidWorks macro or upload XT_Export_Job_Signature.csv.</span>`;return;}
  const matches=payload.matches||[];
  const have=(payload.signature_components||[]).join(', ');
  if(!matches.length){el.innerHTML=`<div style="color:var(--amber)">Signature found for J${esc(payload.job_num||'')}, but no previous signature jobs were available to compare.</div><div style="color:var(--dim);margin-top:6px">Uploaded parts: ${esc(have||'none')}</div>`;return;}
  el.innerHTML=`<div style="color:var(--dim);margin-bottom:10px">Current job J${esc(payload.job_num||'')} signature parts: ${esc(have||'none')}</div>`+
  matches.map((m,idx)=>{
    const comps=(m.components||[]).map(c=>{
      const cur=c.current||{}, old=c.candidate||{};
      const rowColor=Number(c.similarity||0)>=92?'rgba(0,255,170,.08)':Number(c.similarity||0)>=80?'rgba(255,193,7,.08)':'rgba(255,80,80,.08)';
      return `<tr style="background:${rowColor}"><td style="font-weight:900;color:var(--cyan)">${esc(c.component_role||'')}</td><td style="color:${signatureScoreColor(c.similarity)};font-weight:900">${Number(c.similarity||0).toFixed(1)}%</td><td>${sigDims(cur)}<br><span style="color:var(--dim)">${sigDims(old)}</span></td><td>${sigMass(cur)}<br><span style="color:var(--dim)">${sigMass(old)}</span></td><td>${sigCenter(cur)}<br><span style="color:var(--dim)">${sigCenter(old)}</span></td><td>${Number(c.size_diff_pct||0).toFixed(2)}%</td><td>${Number(c.mass_diff_pct||0).toFixed(2)}%</td><td>${Number(c.center_distance_in||0).toFixed(3)}</td></tr>`;
    }).join('');
    const missing=(m.missing_candidate||[]).length?`<div style="color:var(--amber);margin-top:8px">Missing in old job: ${esc((m.missing_candidate||[]).join(', '))}</div>`:'';
    const coverage=Number(m.coverage_pct||0).toFixed(1);
    return `<div class="panel" style="padding:14px;margin:12px 0;background:rgba(255,255,255,.035);border-color:${signatureScoreColor(m.score)}">
      <div style="display:flex;justify-content:space-between;gap:10px;align-items:center;flex-wrap:wrap">
        <div><span style="color:var(--dim)">#${idx+1}</span> <span style="color:var(--blue);font-weight:900;font-size:1.15rem">J${esc(m.job_num||'')}</span> <span style="color:var(--dim)">components ${m.matched_components||0}/6, coverage ${coverage}%</span></div>
        <div style="font-size:1.35rem;color:${signatureScoreColor(m.score)};font-weight:900">${Number(m.score||0).toFixed(1)}%</div>
      </div>
      <div style="overflow:auto;margin-top:10px"><table class="itable"><thead><tr><th>PART</th><th>MATCH</th><th>CURRENT DIMS / OLD DIMS</th><th>CURRENT MASS / OLD MASS</th><th>CURRENT CENTER / OLD CENTER</th><th>SIZE DIFF</th><th>MASS DIFF</th><th>CENTER DIST</th></tr></thead><tbody>${comps}</tbody></table></div>
      ${missing}
    </div>`;
  }).join('');
}

async function loadSignatureMatches(){
  const job=(document.getElementById('sig-job')?.value||'').trim();
  const limit=Number(document.getElementById('sig-limit')?.value||10);
  const el=document.getElementById('sig-result');
  if(!job){if(el)el.innerHTML='<span style="color:var(--red)">Enter a job number first.</span>';return;}
  if(el)el.innerHTML='<span style="color:var(--dim)">Loading signature matches...</span>';
  try{
    const r=await fetch(`/api/job-signatures/${encodeURIComponent(job)}/matches?limit=${encodeURIComponent(limit)}`,{credentials:'same-origin'});
    const payload=await r.json();
    if(!r.ok)throw new Error(payload.detail||'Signature match lookup failed');
    renderSignatureMatches(payload);
  }catch(e){if(el)el.innerHTML=`<span style="color:var(--red)">${esc(e.message||e)}</span>`;}
}

function renderEff(day,month,year,machines,series,comps){
  if(!adminOK()){document.getElementById('eff-content').innerHTML='';return;}
  document.getElementById('eff-content').innerHTML=`<div style="display:flex;flex-direction:column;gap:24px">${renderMoneyEfficiency(series,comps)}${renderEffTimeBlock(day,'TODAY',machines)}${renderEffTimeBlock(month,'THIS MONTH',machines)}${renderEffTimeBlock(year,'THIS YEAR',machines)}</div>`;
}
function renderEffTimeBlock(rows,title,machines){
  const names=Object.keys(machines||{}),by={};(rows||[]).forEach(r=>by[r.machine]=r);
  return`<div><div class="lbl" style="color:var(--cyan);margin-bottom:10px">${title}</div><div class="grid2">${names.map(n=>{const r=by[n]||{},cut=r.cutting_sec||0,idle=r.idle_sec||0,alarm=r.alarm_sec||0,tot=cut+idle+alarm,p=tot?cut/tot*100:0;return`<div class="panel" style="padding:16px"><div class="lbl">${n}</div><div class="num-xl" style="color:${p>50?'var(--green)':p>20?'var(--amber)':'var(--red)'}">${p.toFixed(0)}%</div><div class="prog-bar"><div class="prog-fill" style="width:${p}%;background:var(--cyan)"></div></div><div class="row" style="margin-top:10px"><span>CUT ${fsec(cut)}</span><span>IDLE ${fsec(idle)}</span><span style="color:var(--red)">ALARM ${fsec(alarm)}</span></div></div>`}).join('')}</div></div>`;
}
function renderMoneyEfficiency(series,comps){
  const rows=(series||[]).slice().sort((a,b)=>String(a.due_date||'9999').localeCompare(String(b.due_date||'9999'))||String(a.job_num).localeCompare(String(b.job_num),undefined,{numeric:true}));
  let totalRevenue=0,totalCost=0,totalProfit=0;rows.forEach(j=>{totalRevenue+=Number(j.revenue||0);totalCost+=Number(j.total_cost||0);totalProfit+=Number(j.profit||0);});
  return`<div class="panel" style="overflow:hidden;border-color:rgba(0,229,200,.35)"><div class="phdr">REVENUE / COST / PROFIT — ADMIN ONLY</div><div class="grid3" style="padding:14px"><div class="sbox"><div class="lbl">Total Amount Paid</div><div class="mono" style="font-size:1.2rem;color:var(--green)">${dollars(totalRevenue)}</div></div><div class="sbox"><div class="lbl">Total Cost</div><div class="mono" style="font-size:1.2rem;color:var(--amber)">${dollars(totalCost)}</div></div><div class="sbox"><div class="lbl">Total Profit</div><div class="mono" style="font-size:1.2rem;color:${totalProfit>=0?'var(--green)':'var(--red)'}">${dollars(totalProfit)}</div></div></div><div style="padding:0 14px 14px"><button class="btn-r" onclick="cleanupFakeJobsAdmin()">CLEANUP FAKE JOBS / J000 / J0000 / J1000</button></div><table class="itable"><thead><tr><th>JOB</th><th>DUE</th><th>AMOUNT PAID</th><th>STEEL</th><th>GRINDING</th><th>PYROPEL</th><th>LABOR</th><th>LABOR RATE</th><th>TOTAL COST</th><th>PROFIT</th><th>ACTION</th><th>ADMIN CLEANUP</th></tr></thead><tbody>${rows.map(j=>{const profit=Number(j.profit||0),pc=profit>=0?'var(--green)':'var(--red)',laborDisplay=Number(j.labor_override_cost||0)>0?Number(j.labor_override_cost||0):Number(j.labor_cost||0);const laborRate=Number(j.labor_rate_saved||j.labor_rate||j.labor_rate_effective||75);return`<tr><td style="font-weight:900;color:var(--blue)">J${j.job_num}</td><td><input id="due-${j.job_num}" class="inp" type="date" value="${esc(j.due_date||'')}"><button class="btn-g" style="margin-top:5px" onclick="saveJobDueDate('${j.job_num}')">SAVE DATE</button></td><td><input id="fin-${j.job_num}-revenue" class="inp" type="number" step="0.01" value="${numVal(j.revenue)}"></td><td><input id="fin-${j.job_num}-steel" class="inp" type="number" step="0.01" value="${numVal(j.steel_cost)}"></td><td><input id="fin-${j.job_num}-grinding" class="inp" type="number" step="0.01" value="${numVal(j.grinding_cost)}"></td><td><input id="fin-${j.job_num}-pyropel" class="inp" type="number" step="0.01" value="${numVal(j.pyropel_cost)}"></td><td><input id="fin-${j.job_num}-labor-override" class="inp" type="number" step="0.01" value="${numVal(laborDisplay)}"><div class="lbl">Runtime ${fsec(j.actual_sec||0)}</div></td><td><input id="fin-${j.job_num}-labor" class="inp" type="number" step="0.01" value="${numVal(laborRate)}"></td><td class="mono" style="color:var(--amber)">${dollars(j.total_cost)}</td><td class="mono" style="color:${pc};font-weight:900">${dollars(j.profit)}<div class="lbl">${Number(j.margin_pct||0).toFixed(1)}%</div></td><td><button class="btn-p" onclick="saveJobFinancials('${j.job_num}')">SAVE</button></td><td><div class="row"><button class="btn-y" onclick="renameJobAdmin('${j.job_num}')">RENAME</button><button class="btn-r" onclick="deleteJobAdmin('${j.job_num}')">DELETE JOB</button></div></td></tr>`}).join('')}</tbody></table><div class="notice" style="margin:12px">Formula: Amount Paid - Steel - Grinding - Pyropel - Labor = Profit. Labor rate saves exactly as entered.</div></div>`;
}
async function saveJobFinancials(job){
  if(!adminOK())return alert('Admin login required.');
  const body={revenue:parseFloat(document.getElementById(`fin-${job}-revenue`).value||0),steel_cost:parseFloat(document.getElementById(`fin-${job}-steel`).value||0),grinding_cost:parseFloat(document.getElementById(`fin-${job}-grinding`).value||0),pyropel_cost:parseFloat(document.getElementById(`fin-${job}-pyropel`).value||0),labor_rate:parseFloat(document.getElementById(`fin-${job}-labor`).value||0),labor_override_cost:parseFloat(document.getElementById(`fin-${job}-labor-override`).value||0)};
  const r=await fetch(`/api/job-series/${job}/financials`,{method:'PUT',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify(body)});
  const d=await r.json().catch(()=>({}));
  if(!r.ok)alert('Save failed: '+(d.detail||r.status));else alert('Financials saved for J'+job+'.');
  fetchData();
}
async function saveJobDueDate(job){
  const due=document.getElementById(`due-${job}`).value||'';
  const r=await fetch(`/api/job-series/${job}/due-date`,{method:'PUT',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({due_date:due})});
  const d=await r.json().catch(()=>({}));
  if(r.ok){fetchData();}else alert('Due date update failed: '+(d.detail||'?'));
}
async function saveJobDueDateFromCard(job){
  const el=document.getElementById(`jobcard-due-${job}`);if(!el)return;
  const due=el.value||'';
  const r=await fetch(`/api/job-series/${job}/due-date`,{method:'PUT',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({due_date:due})});
  const d=await r.json().catch(()=>({}));
  if(r.ok)fetchData();else alert('Due date update failed: '+(d.detail||'?'));
}
async function deleteJobAdmin(job){
  if(!adminOK())return alert('Admin login required.');
  const confirmText=`DELETE J${job}`;
  const entered=prompt(`This will delete J${job} from jobs, components, calendar, shop board, schedules, and parsed steel rows.\n\nType exactly: ${confirmText}`,'');
  if(!entered||entered.trim().toUpperCase()!==confirmText.toUpperCase()){alert('Delete cancelled.');return;}
  const r=await fetch(`/api/job-series/${job}`,{method:'DELETE',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({confirm:confirmText})});
  const d=await r.json().catch(()=>({}));
  if(r.ok){alert(`Deleted J${job}.`);fetchData();}else alert('Delete failed: '+(d.detail||'?'));
}
async function renameJobAdmin(oldJob){
  if(!adminOK())return alert('Admin login required.');
  const newJob=prompt(`Rename J${oldJob} to what job number?`,'');
  if(!newJob)return;
  const merge=confirm(`If J${newJob} already exists, merge J${oldJob} into it?\n\nOK = merge if exists\nCancel = fail if target exists`);
  const r=await fetch(`/api/job-series/${oldJob}/rename`,{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({new_job_num:newJob,merge_if_exists:merge})});
  const d=await r.json().catch(()=>({}));
  if(r.ok){alert(`Renamed J${oldJob} to J${d.new_job_num}.`);fetchData();}else alert('Rename failed: '+(d.detail||'?'));
}
async function cleanupFakeJobsAdmin(){
  if(!adminOK())return alert('Admin login required.');
  const raw=prompt('Enter fake job numbers to remove, comma separated:','000,0000,1000');
  if(!raw)return;
  const jobs=raw.split(',').map(x=>x.trim()).filter(Boolean);
  if(!jobs.length)return;
  if(!confirm('Delete fake jobs: '+jobs.map(j=>'J'+j).join(', ')+'?'))return;
  const r=await fetch('/api/job-series/cleanup-fake',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({jobs})});
  const d=await r.json().catch(()=>({}));
  if(r.ok){alert('Fake jobs cleaned.');fetchData();}else alert('Cleanup failed: '+(d.detail||'?'));
}

function renderTools(tools){
  const el=document.getElementById('tool-list');if(!tools.length){el.innerHTML='<div class="panel" style="padding:40px;text-align:center;color:var(--dim)">No tool data yet.</div>';return;}
  const by={};tools.forEach(t=>{if(String(t.tid||'0')==='0')return;(by[t.machine]=by[t.machine]||[]).push(t)});
  el.innerHTML=Object.entries(by).map(([mach,arr])=>`<div class="panel" style="padding:16px;margin-bottom:18px"><div class="phdr" style="margin:-16px -16px 14px">${mach} · TOOL HEALTH</div><div class="tool-grid">${arr.map(t=>toolCard(t)).join('')}</div></div>`).join('');
}
function toolCard(t){const wear=Number(t.wear_score||0),remaining=Math.max(0,100-wear),col=hcolor(wear),pending=Number(t.replacement_pending||0)&&adminOK()?'pending':'';return`<div class="tool-card ${pending}" onclick="openToolDetail('${esc(t.machine)}','${esc(t.tid)}')"><div class="tool-title"><div><div class="tool-no">T${t.tid}</div><div class="tool-name">${esc(t.tool_name||'')}</div><div class="lbl">${esc(t.tool_category||'Tool')} · DIA ${Number(t.diameter||0).toFixed(3)} · ${t.flutes||''} FL</div></div><div style="text-align:right">${pending?'<span class="pill pill-warn">ADMIN CHECK</span>':''}<div class="mono" style="color:${col};font-size:1.1rem;font-weight:900">${wear.toFixed(0)}%</div><div class="lbl">WEAR</div></div></div><div class="lbl">HEALTH REMAINING ${remaining.toFixed(0)}%</div><div class="health-bar" style="margin:4px 0 10px"><div class="health-fill" style="width:${remaining}%;background:${col}"></div></div><div class="grid3"><div class="sbox"><div class="lbl">In Cut</div><div class="mono">${fsec(t.in_cut_sec||0)}</div></div><div class="sbox"><div class="lbl">Life</div><div class="mono">${Number(t.max_life_min||0).toFixed(1)}m</div></div><div class="sbox"><div class="lbl">Peak Load</div><div class="mono">${Number(t.peak_load||0).toFixed(1)}%</div></div></div></div>`;}

async function openToolDetail(machine,tid){
  const r=await fetch('/api/tools/detail',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({machine,tid})});
  const d=await r.json().catch(()=>({}));
  if(!r.ok){alert('Tool detail failed');return;}
  currentToolDetail=d;
  currentToolTab='overview';
  document.getElementById('tool-modal-title').textContent=`${machine} · T${tid} · ${(d.health||{}).tool_name||''}`;
  document.getElementById('tool-modal-backdrop').classList.add('active');
  renderToolModalContent();
}
function setToolTab(tab){currentToolTab=tab;renderToolModalContent();}
function renderToolModalContent(){
  const d=currentToolDetail;if(!d)return;
  const h=d.health||{},wear=Number(h.wear_score||0),remaining=Math.max(0,100-wear),col=hcolor(wear);
  document.getElementById('tool-modal-content').innerHTML=`
    <div class="tool-tabs">
      <button class="tool-tab-btn ${currentToolTab==='overview'?'active':''}" onclick="setToolTab('overview')">OVERVIEW</button>
      <button class="tool-tab-btn ${currentToolTab==='inspection'?'active':''}" onclick="setToolTab('inspection')">INSPECTION</button>
      <button class="tool-tab-btn ${currentToolTab==='history'?'active':''}" onclick="setToolTab('history')">HISTORY</button>
      <button class="tool-tab-btn ${currentToolTab==='chart'?'active':''}" onclick="setToolTab('chart')">CUT DATA / CHART</button>
    </div>
    <div class="tool-tab ${currentToolTab==='overview'?'active':''}">${renderToolOverview(d,col,wear,remaining)}</div>
    <div class="tool-tab ${currentToolTab==='inspection'?'active':''}">${renderToolInspectionTab(d)}</div>
    <div class="tool-tab ${currentToolTab==='history'?'active':''}">${renderToolHistoryTab(d)}</div>
    <div class="tool-tab ${currentToolTab==='chart'?'active':''}"><div class="chart-box"><canvas id="tool-scatter-chart"></canvas></div></div>
  `;
  if(currentToolTab==='chart')setTimeout(()=>drawToolScatterChart(d.samples||[]),30);
}
function renderToolOverview(d,col,wear,remaining){
  const h=d.health||{},machine=d.machine,tid=d.tid;
  return `<div class="grid4" style="margin-bottom:14px"><div class="sbox"><div class="lbl">Wear</div><div class="mono" style="color:${col};font-size:1.2rem">${wear.toFixed(1)}%</div></div><div class="sbox"><div class="lbl">Health Remaining</div><div class="mono">${remaining.toFixed(1)}%</div></div><div class="sbox"><div class="lbl">In Cut</div><div class="mono">${fsec(h.in_cut_sec||0)}</div></div><div class="sbox"><div class="lbl">Life Limit</div><div class="mono">${Number(h.max_life_min||0).toFixed(1)}m</div></div></div>
  <div class="health-bar" style="height:18px;margin-bottom:16px"><div class="health-fill" style="width:${remaining}%;background:${col}"></div></div>
  <div class="grid3">
    <div class="sbox"><div class="lbl">Category</div><div class="mono">${esc(h.tool_category||'Tool')}</div></div>
    <div class="sbox"><div class="lbl">Diameter</div><div class="mono">${Number(h.diameter||0).toFixed(4)}</div></div>
    <div class="sbox"><div class="lbl">Flutes</div><div class="mono">${esc(h.flutes||'')}</div></div>
  </div>
  ${adminOK()?`<div class="panel2" style="padding:14px;margin-top:14px"><div class="phdr" style="margin:-14px -14px 12px">ADMIN TOOL REPLACEMENT RESET</div><input id="replace-reason" class="inp" value="Manual replacement"><textarea id="replace-notes" class="txtarea" placeholder="Notes"></textarea><button class="btn-r" onclick="manualReplaceTool('${esc(machine)}','${esc(tid)}')">REPLACE TOOL / RESET LIFE</button></div>`:`<div class="notice">Tool replacement reset is admin-only.</div>`}`;
}
function renderToolInspectionTab(d){
  const machine=d.machine,tid=d.tid;
  const latest=(d.inspections||[])[0];
  return `<div class="panel2" style="padding:14px;margin-bottom:14px">
    <div class="phdr" style="margin:-14px -14px 12px">TOOL INSPECTION / QUALITY MARK</div>
    <div class="notice" style="margin-bottom:10px">Marking quality updates visible tool health. Saving inspection/calibration is admin-only.</div>
    ${latest?`<div class="goodnote" style="margin-bottom:10px">Last inspection: ${esc(latest.condition||'')} · ${esc(latest.inspected_by||'')} · ${esc(latest.timestamp||'')}</div>`:''}
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px;align-items:end">
      <div class="col"><div class="lbl">Quality / Condition</div><select id="inspect-condition" class="sel"><option value="GOOD">GOOD</option><option value="OK">OK</option><option value="WATCH">WATCH</option><option value="WORN">WORN</option><option value="CHIPPED">CHIPPED</option><option value="BROKEN">BROKEN</option><option value="REPLACED">REPLACED</option></select></div>
      <div class="col"><div class="lbl">Observed Wear %</div><input id="inspect-observed-wear" class="inp" type="number" step="1" value="25"></div>
      <div class="col"><div class="lbl">Inspected By</div><input id="inspect-by" class="inp" placeholder="Name"></div>
      <div class="col"><div class="lbl">Measured Dia Optional</div><input id="inspect-dia" class="inp" type="number" step="0.0001" value="0"></div>
      <div class="col"><div class="lbl">Measured Length Offset Optional</div><input id="inspect-len" class="inp" type="number" step="0.0001" value="0"></div>
      <div class="col"><div class="lbl">Calibrate Health</div><select id="inspect-calibrate" class="sel"><option value="1">YES</option><option value="0">NO - LOG ONLY</option></select></div>
      <button class="${adminOK()?'btn-p':'btn-y'}" onclick="submitToolInspection('${esc(machine)}','${esc(tid)}')">${adminOK()?'SAVE INSPECTION':'ADMIN REQUIRED'}</button>
    </div>
    <div class="col" style="margin-top:10px"><div class="lbl">Edge / Finish Notes</div><textarea id="inspect-edge-notes" class="txtarea" placeholder="Example: edge looks good, no chips, finish still good"></textarea></div>
    <div class="col" style="margin-top:10px"><div class="lbl">Extra Notes</div><textarea id="inspect-notes" class="txtarea" placeholder="Optional"></textarea></div>
    <div id="inspect-msg" class="mono" style="margin-top:10px;color:var(--green)"></div>
  </div>
  <div class="panel2" style="overflow:hidden">
    <div class="phdr">RECENT INSPECTIONS</div>
    <table class="itable"><thead><tr><th>TIME</th><th>COND</th><th>BY</th><th>OBS %</th><th>OLD → NEW WEAR</th><th>LIFE MIN</th><th>NOTES</th></tr></thead><tbody>${(d.inspections||[]).map(i=>`<tr><td class="mono">${esc(i.timestamp||'')}</td><td><span class="pill ${i.condition==='GOOD'||i.condition==='OK'?'pill-ok':i.condition==='WATCH'?'pill-warn':'pill-crit'}">${esc(i.condition||'')}</span></td><td>${esc(i.inspected_by||'')}</td><td>${Number(i.observed_wear_pct||0).toFixed(0)}%</td><td class="mono">${Number(i.old_wear_score||0).toFixed(1)} → ${Number(i.new_wear_score||0).toFixed(1)}</td><td class="mono">${Number(i.old_max_life_min||0).toFixed(1)} → ${Number(i.new_max_life_min||0).toFixed(1)}</td><td>${esc(i.edge_notes||i.notes||'')}</td></tr>`).join('')}</tbody></table>
  </div>`;
}
function renderToolHistoryTab(d){
  return `<div class="grid2">
    <div class="panel2" style="overflow:hidden"><div class="phdr">LIFE CYCLES</div><table class="itable"><thead><tr><th>ID</th><th>CYCLE</th><th>STATUS</th><th>CUT</th><th>WEAR</th><th>START</th><th>END</th></tr></thead><tbody>${(d.cycles||[]).map(c=>`<tr><td>${c.id}</td><td>${c.cycle_no}</td><td>${esc(c.status||'')}</td><td>${fsec(c.total_cut_sec||0)}</td><td>${Number(c.final_wear_score||0).toFixed(1)}%</td><td class="mono">${esc(c.started_at||'')}</td><td class="mono">${esc(c.ended_at||'')}</td></tr>`).join('')}</tbody></table></div>
    <div class="panel2" style="overflow:hidden"><div class="phdr">REPLACEMENT EVENTS</div><table class="itable"><thead><tr><th>TIME</th><th>SOURCE</th><th>REASON</th><th>OLD→NEW</th></tr></thead><tbody>${(d.events||[]).map(e=>`<tr><td class="mono">${esc(e.timestamp||'')}</td><td>${esc(e.source||'')}</td><td>${esc(e.reason||'')}</td><td>${e.old_cycle_id} → ${e.new_cycle_id}</td></tr>`).join('')}</tbody></table></div>
  </div>`;
}
function closeToolModal(){document.getElementById('tool-modal-backdrop').classList.remove('active');if(toolChart){toolChart.destroy();toolChart=null;}currentToolDetail=null;}
function drawToolScatterChart(samples){const ctx=document.getElementById('tool-scatter-chart');if(!ctx)return;if(toolChart){toolChart.destroy();toolChart=null;}toolChart=new Chart(ctx,{type:'scatter',data:{datasets:[{label:'Load %',data:samples.map((s,i)=>({x:i,y:Number(s.spindle_load||0)})),backgroundColor:'#00e5c8'},{label:'RPM / 100',data:samples.map((s,i)=>({x:i,y:Number(s.spindle_rpm||0)/100})),backgroundColor:'#3d7fff'},{label:'Wear %',data:samples.map((s,i)=>({x:i,y:Number(s.wear_score||0)})),backgroundColor:'#ff3b5c'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#c8ccd8'}}},scales:{x:{ticks:{color:'#596078'},grid:{color:'rgba(255,255,255,.06)'}},y:{ticks:{color:'#596078'},grid:{color:'rgba(255,255,255,.06)'}}}}});}
async function manualReplaceTool(machine,tid){if(!adminOK())return alert('Admin required');if(!confirm('Archive this tool life and start a new cycle?'))return;await fetch('/api/tools/replace',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({machine,tid,reason:document.getElementById('replace-reason')?.value||'Manual replacement',notes:document.getElementById('replace-notes')?.value||''})});closeToolModal();fetchData();}
async function submitToolInspection(machine,tid){
  if(!adminOK())return alert('Admin login required.');
  const msg=document.getElementById('inspect-msg');
  const body={machine,tid,condition:document.getElementById('inspect-condition').value,inspected_by:document.getElementById('inspect-by').value,observed_wear_pct:parseFloat(document.getElementById('inspect-observed-wear').value||0),measured_diameter:parseFloat(document.getElementById('inspect-dia').value||0),measured_length_offset:parseFloat(document.getElementById('inspect-len').value||0),edge_notes:document.getElementById('inspect-edge-notes').value||'',action_taken:'INSPECTED',calibrate_health:parseInt(document.getElementById('inspect-calibrate').value||0),notes:document.getElementById('inspect-notes').value||''};
  msg.textContent='Saving inspection...';msg.style.color='var(--cyan)';
  const r=await fetch('/api/tools/inspect',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify(body)});
  const d=await r.json().catch(()=>({}));
  if(r.ok){msg.textContent=`Saved. Wear ${Number(d.old_wear_score||0).toFixed(1)}% → ${Number(d.new_wear_score||0).toFixed(1)}%.`;msg.style.color='var(--green)';setTimeout(async()=>{await fetchData();openToolDetail(machine,tid);},700);}
  else{msg.textContent='Inspection failed: '+(d.detail||'?');msg.style.color='var(--red)';}
}

function renderInv(inventory,summary){
  document.getElementById('s-total').textContent=summary.total||0;document.getElementById('s-low').textContent=summary.low||0;document.getElementById('s-out').textContent=summary.out||0;
  const st=document.getElementById('s-status');st.textContent=(summary.out||0)>0?'CRITICAL':(summary.low||0)>0?'LOW STOCK':'STABLE';st.style.color=(summary.out||0)>0?'var(--red)':(summary.low||0)>0?'var(--amber)':'var(--green)';
  document.getElementById('inv-tbody').innerHTML=(inventory||[]).map(i=>{
    const burn=i.used_30d/30,days=burn>0?Math.floor(i.quantity/burn):999;
    const qc=i.quantity<=0?'var(--red)':i.quantity<=i.min_threshold?'var(--amber)':'var(--text)';
    const dc=days<=3?'pill-crit':days<=7?'pill-warn':'pill-ok';
    return`<tr><td class="mono" style="color:var(--blue);font-weight:700">${esc(i.sku)}</td><td>${esc(i.description||'-')}<div class="lbl">${esc(i.category||'Tool')}</div></td><td style="color:${qc};font-weight:800">${i.quantity}<div class="lbl">MIN ${i.min_threshold}</div></td><td>${i.used_7d||0}</td><td>${i.used_30d||0}</td><td>${fmtHours(i.avg_hours_between_out)}<div class="lbl">BETWEEN OUTS</div></td><td>${fmtDateTime(i.last_out_timestamp)}<div class="lbl">${fmtHours(i.hours_since_last_out)} AGO</div></td><td><span class="pill ${dc}">${days===999?'STABLE':days+'d'}</span></td></tr>`;
  }).join('');
}
async function addInventoryItem(){const body={sku:document.getElementById('new-sku').value.trim().toUpperCase(),description:document.getElementById('new-desc').value.trim(),category:document.getElementById('new-cat').value.trim()||'Tool',quantity:parseInt(document.getElementById('new-qty').value||0),min_threshold:parseInt(document.getElementById('new-min').value||5)};const r=await fetch('/api/inventory/items',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify(body)});document.getElementById('new-item-msg').textContent=r.ok?'Added/updated '+body.sku:'Error';if(r.ok){document.getElementById('qr-label-sku').value=body.sku;document.getElementById('qr-label-desc').value=body.description;makeQrLabels();}fetchData();}
async function submitScan(){const sku=document.getElementById('sku-input').value.trim().toUpperCase(),dir=document.getElementById('dir-select').value,qty=Math.max(1,parseInt(document.getElementById('manual-qty').value||1));if(!sku)return;const r=await fetch('/api/scan',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({sku,qty,direction:dir})});const d=await r.json().catch(()=>({}));document.getElementById('scan-msg').textContent=r.ok?`OK ${dir} ${qty} ${sku}`:'Error '+(d.detail||'');fetchData();}

async function resolveBarcode(code){
  code=String(code||'').trim();
  if(!code)return;
  const now=Date.now();
  if(code===lastBarcode&&now-lastBarcodeAt<700)return;
  lastBarcode=code;lastBarcodeAt=now;
  const msg=document.getElementById('barcode-msg');
  msg.textContent='Resolving QR...';
  msg.style.color='var(--cyan)';
  const r=await fetch('/api/inventory/qr-resolve',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({barcode:code})});
  const d=await r.json().catch(()=>({}));
  if(!r.ok){msg.textContent='QR resolve failed: '+(d.detail||r.status);msg.style.color='var(--red)';return;}
  lastResolvedQR=d;
  showQrAction(d);
}
function showQrAction(d){
  const box=document.getElementById('qr-action-box');
  const summary=document.getElementById('qr-resolved-summary');
  const known=document.getElementById('qr-known-actions');
  const create=document.getElementById('qr-create-actions');
  box.classList.add('active');
  if(d.status==='found'){
    summary.innerHTML=`<div class="goodnote"><div class="mono" style="font-size:1rem;color:var(--cyan)">${esc(d.sku)} · ${esc(d.description)}</div><div class="lbl">CURRENT QTY: ${d.quantity} · CATEGORY: ${esc(d.category||'Tool')}</div></div>`;
    known.classList.remove('hidden');
    create.classList.add('hidden');
    document.getElementById('qr-label-sku').value=d.sku||'';
    document.getElementById('qr-label-desc').value=d.description||'';
    makeQrLabels();
  }else{
    summary.innerHTML=`<div class="notice"><div class="mono" style="font-size:1rem;color:var(--amber)">NEW QR CODE FOUND</div><div class="lbl">This code is not mapped to a tool yet.</div></div>`;
    known.classList.add('hidden');
    create.classList.remove('hidden');
    document.getElementById('qr-create-sku').value=d.sku||d.sku_hint||'';
    document.getElementById('qr-create-desc').value=d.description||d.description_hint||'';
    document.getElementById('qr-create-cat').value=d.category||'Tool';
    document.getElementById('qr-label-sku').value=d.sku||d.sku_hint||'';
    document.getElementById('qr-label-desc').value=d.description||d.description_hint||'';
  }
  document.getElementById('barcode-msg').textContent='Choose ADD or SUBTRACT and quantity.';
  document.getElementById('barcode-msg').style.color='var(--green)';
}
async function qrAction(direction, createMode=false){
  if(!lastResolvedQR)return alert('Scan a QR first.');
  const body={
    barcode:lastResolvedQR.alias,
    direction,
    qty:createMode?parseInt(document.getElementById('qr-create-qty').value||1):parseInt(document.getElementById('qr-action-qty').value||1),
    sku:createMode?document.getElementById('qr-create-sku').value:'',
    description:createMode?document.getElementById('qr-create-desc').value:'',
    category:createMode?document.getElementById('qr-create-cat').value:'Tool',
    starting_quantity:createMode?parseInt(document.getElementById('qr-create-start').value||0):0,
    min_threshold:createMode?parseInt(document.getElementById('qr-create-min').value||5):5
  };
  body.qty=Math.max(1,body.qty||1);
  const r=await fetch('/api/inventory/qr-action',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify(body)});
  const d=await r.json().catch(()=>({}));
  const msg=document.getElementById('barcode-msg');
  if(r.ok){
    msg.textContent=`${d.direction==='OUT'?'SUBTRACTED':'ADDED'} ${d.qty} · ${d.sku} · Qty ${d.old_quantity} → ${d.new_quantity}`;
    msg.style.color='var(--green)';
    document.getElementById('qr-action-box').classList.remove('active');
    document.getElementById('qr-label-sku').value=d.sku||'';
    document.getElementById('qr-label-desc').value=d.description||'';
    makeQrLabels();
    fetchData();
  }else{
    msg.textContent='QR action failed: '+(d.detail||r.status);
    msg.style.color='var(--red)';
  }
}
function qrTakeOut(){qrAction('OUT',false);}
function qrAddIn(){qrAction('IN',false);}
function qrCreateAndTakeOut(){qrAction('OUT',true);}
function qrCreateAndAdd(){qrAction('IN',true);}
function manualBarcodeSubmit(){const el=document.getElementById('manual-barcode');if(el.value.trim()){resolveBarcode(el.value.trim());el.value='';}}
function browserLikelyNeedsHttpsForCamera(){const host=location.hostname;return location.protocol!=='https:'&&!['localhost','127.0.0.1','::1'].includes(host);}
async function startBarcodeScanner(){
  const el=document.getElementById('barcode-reader'),msg=document.getElementById('barcode-msg');el.style.display='block';
  if(browserLikelyNeedsHttpsForCamera()){
    msg.innerHTML='Live camera requires HTTPS on iPhone/phone browsers. Use PHONE PHOTO SCAN, or open through ngrok/HTTPS.';
    msg.style.color='var(--amber)';
    return;
  }
  try{
    if(barcodeScanner){try{await barcodeScanner.stop();}catch(e){}try{barcodeScanner.clear();}catch(e){}}
    barcodeScanner=new Html5Qrcode('barcode-reader');
    await barcodeScanner.start({facingMode:'environment'},{fps:10,qrbox:{width:300,height:170}},decodedText=>resolveBarcode(decodedText),()=>{});
    msg.textContent='Scanner running. Point camera at QR.';
    msg.style.color='var(--green)';
  }catch(e){
    msg.textContent='Camera failed: '+e+'. Use PHONE PHOTO SCAN.';
    msg.style.color='var(--red)';
  }
}
function openBarcodePhotoFallback(){const input=document.getElementById('barcode-file-input');input.value='';input.click();}
async function decodeBarcodePhoto(input){
  if(!input.files||!input.files[0])return;
  const msg=document.getElementById('barcode-msg'),el=document.getElementById('barcode-reader');
  try{
    el.style.display='block';
    const scanner=new Html5Qrcode('barcode-reader');
    const decoded=await scanner.scanFile(input.files[0],true);
    try{scanner.clear();}catch(e){}
    await resolveBarcode(decoded);
  }catch(e){
    msg.textContent='Could not decode photo. Try closer/brighter photo of the QR.';
    msg.style.color='var(--red)';
  }
}
async function stopBarcodeScanner(){try{if(barcodeScanner){try{await barcodeScanner.stop();}catch(e){}try{barcodeScanner.clear();}catch(e){}barcodeScanner=null;}document.getElementById('barcode-reader').style.display='none';}catch(e){}}
function qrPayloadForLabel(sku,desc){const style=document.getElementById('qr-label-style')?.value||'plain';sku=String(sku||'').trim().toUpperCase();desc=String(desc||'').trim();if(style==='sku_prefix')return 'SKU:'+sku;if(style==='mcn_pipe')return 'ECMS|SKU|'+sku+'|DESC|'+desc;return sku;}
function makeQrLabels(){const sku=document.getElementById('qr-label-sku').value.trim().toUpperCase(),desc=document.getElementById('qr-label-desc').value.trim(),sheet=document.getElementById('qr-label-sheet');if(!sku)return;const payload=qrPayloadForLabel(sku,desc);sheet.innerHTML=`<div class="qr-label"><canvas id="qr-label-canvas-0"></canvas><div><div class="qr-label-title">${esc(sku)}</div><div class="qr-label-desc">${esc(desc||'Tool / Inventory Item')}</div><div class="qr-label-payload">${esc(payload)}</div></div></div>`;new QRious({element:document.getElementById('qr-label-canvas-0'),value:payload,size:144,level:'H'});}
function printQrLabels(){if(!document.getElementById('qr-label-sheet').innerHTML.trim())return alert('Make label first.');window.print();}

function renderAdmin(d){
  const msg=document.getElementById('admin-msg'),el=document.getElementById('admin-only-content'),note=document.getElementById('admin-locked-note');
  if(!adminOK()){
    msg.textContent='LOCKED — Admin login required.';
    msg.style.color='var(--amber)';
    note.style.display='block';
    el.innerHTML='';
    return;
  }
  msg.textContent='Admin authenticated.';
  msg.style.color='var(--green)';
  note.style.display='none';
  el.innerHTML=`${renderNetworkAdmin(d.network||{})}${renderToolCandidateAdmin(d.tool_replacement_candidates||[])}
  ${renderProgramProgressAdmin(d.machines||{})}
  ${renderUmcPotScheduleAdmin(d.umc_pot_schedule||[])}
  ${renderSteelSheetAdmin(d.steel_sheet_items||[], d.steel_sheet_logs||[])}
  ${renderGcodeEstimatesAdmin(d.gcode_estimates||[])}
  <div class="panel" style="padding:18px;margin-bottom:18px"><div class="phdr" style="margin:-18px -18px 16px">TELEMETRY DEBUG</div><div id="telemetry-debug"></div></div>`;
  renderTelemetryDebug(d);
}
function renderNetworkAdmin(n){const urls=n.urls||[];return`<div class="panel" style="padding:18px;margin-bottom:18px"><div class="phdr" style="margin:-18px -18px 16px">LOCAL APP NETWORK ACCESS</div><div class="notice">Run this app on the main computer. Other computers do not need Wi-Fi if they are connected by Ethernet/private LAN and can reach this host computer. Windows Firewall must allow TCP port ${esc(n.port||'2926')}.</div><table class="itable" style="margin-top:12px"><thead><tr><th>IP</th><th>HTTP URL</th><th>HTTPS URL</th><th>TYPE</th></tr></thead><tbody>${urls.map(u=>`<tr><td class="mono">${esc(u.ip)}</td><td><a href="${esc(u.http_url)}" target="_blank">${esc(u.http_url)}</a></td><td>${esc(u.https_url)}</td><td>${u.is_localhost?'THIS COMPUTER':'NETWORK ADAPTER'}</td></tr>`).join('')}</tbody></table></div>`;}
function renderToolCandidateAdmin(candidates){return`<div class="panel" style="padding:14px;margin-bottom:18px;border-color:${candidates.length?'var(--amber)':'var(--border)'}"><div class="lbl" style="color:${candidates.length?'var(--amber)':'var(--dim)'};font-size:.85rem">TOOL OFFSET CHANGE WARNINGS — CONFIRM ONLY IF TOOL WAS ACTUALLY REPLACED</div>${candidates.length?`<div style="margin-top:10px">${candidates.map(c=>`<div class="panel2" style="padding:10px;margin-bottom:8px;display:flex;justify-content:space-between;gap:10px;align-items:center;flex-wrap:wrap"><div><div class="mono" style="color:var(--amber)">T${c.tid} · ${c.machine} · ${esc(c.tool_name||'')}</div><div class="lbl">L Δ ${Number(c.delta_l_wear||0).toFixed(4)} · D Δ ${Number(c.delta_d_wear||0).toFixed(4)}</div><div class="lbl">${esc(c.detected_reason||'')}</div></div><div class="row"><button class="btn-g" onclick="confirmToolCandidate(${c.id})">CONFIRM REPLACED</button><button class="btn-r" onclick="ignoreToolCandidate(${c.id})">IGNORE</button></div></div>`).join('')}</div><button class="btn-y" onclick="clearPendingCandidates()">CLEAR ALL PENDING FALSE WARNINGS</button>`:'<div class="goodnote">No pending offset-change warnings.</div>'}</div>`;}
async function confirmToolCandidate(id){const notes=prompt('Confirm ONLY if the tool was actually replaced/reloaded. Notes:','Operator confirmed actual tool replacement')||'';await fetch(`/api/tools/replacement-candidates/${id}/confirm`,{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({notes,reason:'Operator confirmed actual tool replacement after offset change'})});fetchData();}
async function ignoreToolCandidate(id){const notes=prompt('Ignore notes:','False alarm / offset adjustment only')||'';await fetch(`/api/tools/replacement-candidates/${id}/ignore`,{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({notes})});fetchData();}
async function clearPendingCandidates(){if(!confirm('Clear all pending offset-change warnings as ignored?'))return;await fetch('/api/tools/replacement-candidates/clear-pending',{method:'POST',credentials:'same-origin'});fetchData();}
function renderProgramProgressAdmin(machines){return`<div class="panel" style="padding:18px;margin-bottom:18px"><div class="phdr" style="margin:-18px -18px 16px">LIVE PROGRAM PROGRESS ESTIMATES</div><table class="itable"><thead><tr><th>MACHINE</th><th>PROGRAM</th><th>JOB</th><th>EST TOTAL</th><th>PROGRESS</th><th>REMAINING</th><th>EST FINISH</th><th>SOURCE</th></tr></thead><tbody>${Object.entries(machines||{}).map(([name,m])=>`<tr><td>${esc(name)}</td><td class="mono">${esc(m.prog||'')}</td><td>${m.job_num?'J'+m.job_num:'—'}</td><td>${Number(m.program_est_sec||0)>0?fsec(m.program_est_sec):'—'}</td><td>${Number(m.program_est_sec||0)>0?`<div class="mono" style="color:var(--cyan)">${Number(m.program_progress_pct||0).toFixed(1)}%</div><div class="prog-bar"><div class="prog-fill" style="width:${Math.min(Number(m.program_progress_pct||0),100)}%;background:var(--cyan)"></div></div>`:'—'}</td><td>${Number(m.program_est_sec||0)>0?fsec(m.program_remaining_sec||0):'—'}</td><td>${esc(m.program_est_finish||'')}</td><td>${esc(m.program_progress_source||'')}</td></tr>`).join('')}</tbody></table></div>`;}
function renderUmcPotScheduleAdmin(rows){return`<div class="panel" style="padding:18px;margin-bottom:18px"><div class="phdr" style="margin:-18px -18px 16px">UMC-1000 POT SCHEDULE OPTIMIZER</div><div class="notice">Builds the most efficient UMC-1000 ID/OD POT cut order while respecting due dates first. Runtime comes from Q-code, G-code, or similar completed same-component jobs.</div><div class="row" style="margin-top:12px"><div class="col" style="min-width:240px"><div class="lbl">Schedule Start</div><input id="umc-sched-start" class="inp" type="datetime-local"></div><button class="btn-p" onclick="generateUmcPotSchedule()">GENERATE UMC POT SCHEDULE</button></div><div style="margin-top:14px;overflow:auto"><table class="itable"><thead><tr><th>#</th><th>JOB</th><th>COMPONENT</th><th>DUE</th><th>STEEL</th><th>EST CUT</th><th>SETUP</th><th>TOTAL</th><th>START</th><th>END</th><th>WHY</th><th>EFFICIENCY</th></tr></thead><tbody>${(rows||[]).map(r=>{const steelColor=Number(r.steel_ready||0)?'var(--green)':'var(--amber)',dueCol=dueColor(r.due_date);return`<tr><td class="mono">${r.schedule_rank}</td><td style="font-weight:900;color:var(--blue)">J${r.job_num}</td><td>${esc(r.component_name||'')}</td><td style="color:${dueCol}">${r.due_date||'—'}<div class="lbl">${daysLeft(r.due_date)}</div></td><td style="color:${steelColor}">${Number(r.steel_ready||0)?'READY':'NOT READY'}<div class="lbl">${esc(r.steel_status||'')}</div></td><td>${Number(r.estimated_sec||0)>0?fsec(r.estimated_sec):'UNKNOWN'}</td><td>${fsec(r.setup_sec||0)}</td><td>${Number(r.total_planned_sec||0)>0?fsec(r.total_planned_sec):'UNKNOWN'}</td><td class="mono">${esc(r.planned_start||'')}</td><td class="mono">${esc(r.planned_end||'')}</td><td>${esc(r.priority_reason||'')}</td><td>${esc(r.efficiency_reason||'')}</td></tr>`}).join('')}</tbody></table></div></div>`;}
async function generateUmcPotSchedule(){const start=document.getElementById('umc-sched-start')?.value||'';const r=await fetch('/api/schedule/umc1000-pots/optimize',{method:'POST',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({start_at:start})});const d=await r.json().catch(()=>({}));if(r.ok){alert('UMC-1000 pot schedule generated. Jobs scheduled: '+(d.count||0));fetchData();}else alert('Schedule failed: '+(d.detail||'?'));}
function renderSteelSheetAdmin(items,logs){
  const byJob={};(items||[]).forEach(r=>{const j=r.job_num||'';(byJob[j]=byJob[j]||[]).push(r);});
  const jobBlocks=Object.entries(byJob).map(([job,rows])=>`<div class="panel2" style="padding:12px;margin-bottom:10px"><div><div class="mono" style="color:var(--blue);font-size:1rem">J${esc(job)}</div><div class="lbl">${rows.length} parsed steel item(s)</div></div><table class="itable" style="margin-top:8px"><thead><tr><th>QTY</th><th>ITEM</th><th>DIM 1</th><th>DIM 2</th><th>DIM 3</th><th>MATERIAL</th><th>CONF</th><th>RAW</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${Number(r.qty||0).toFixed(0)}</td><td style="font-weight:800">${esc(r.item_name||'')}</td><td>${Number(r.dim_1||0).toFixed(3)}</td><td>${Number(r.dim_2||0).toFixed(3)}</td><td>${Number(r.dim_3||0).toFixed(3)}</td><td style="color:${String(r.material||'').includes('4140')?'var(--green)':'var(--cyan)'}">${esc(r.material||'')}</td><td>${Number(r.confidence||0).toFixed(2)}</td><td class="mono" style="font-size:.58rem;color:var(--dim)">${esc(r.raw_line||'')}</td></tr>`).join('')}</tbody></table></div>`).join('');
  return`<div class="panel" style="padding:18px;margin-bottom:18px"><div class="phdr" style="margin:-18px -18px 16px">STEEL SHEET READER — VIEW ONLY</div><div class="notice">Steel PDF uploads happen only under Job Matching. This panel only shows parsed results.</div><div style="margin-top:16px">${jobBlocks||'<div class="notice">No parsed steel-sheet items yet.</div>'}</div><div class="panel2" style="padding:12px;margin-top:14px"><div class="phdr" style="margin:-12px -12px 12px">RECENT PARSE LOG</div><table class="itable"><thead><tr><th>TIME</th><th>JOB</th><th>STATUS</th><th>COUNT</th><th>MESSAGE</th></tr></thead><tbody>${(logs||[]).slice(0,50).map(l=>`<tr><td class="mono">${esc(l.timestamp||'')}</td><td>J${esc(l.job_num||'')}</td><td style="color:${l.status==='ok'?'var(--green)':l.status==='error'?'var(--red)':'var(--amber)'}">${esc(l.status||'')}</td><td>${l.parsed_count||0}</td><td>${esc(l.message||'')}</td></tr>`).join('')}</tbody></table></div></div>`;
}
function renderGcodeEstimatesAdmin(rows){return`<div class="panel" style="padding:18px;margin-bottom:18px"><div class="phdr" style="margin:-18px -18px 16px">G-CODE TOOL LIFE ANALYSIS</div><div class="notice">Estimates tool severity and program length from feed moves, arcs, Z plunges, stepdowns, drilling cycles, peck cycles, RPM, feed, and coolant.</div><table class="itable"><thead><tr><th>ANALYZED</th><th>MACHINE</th><th>PROGRAM</th><th>TOOL</th><th>SEVERITY</th><th>EST CUT</th><th>CUT DIST</th><th>AVG FEED</th><th>AVG RPM</th><th>MAX DEPTH</th><th>AVG STEP</th><th>DRILL</th><th>PECK</th></tr></thead><tbody>${(rows||[]).map(r=>{const sev=Number(r.severity_multiplier||1),sevColor=sev>=1.6?'var(--red)':sev>=1.2?'var(--amber)':'var(--green)';return`<tr><td class="mono">${esc(r.analyzed_at||'')}</td><td>${esc(r.machine||'')}</td><td class="mono">${esc(r.program_key||'')}</td><td style="color:var(--blue);font-weight:900">T${esc(r.tid||'')}</td><td class="mono" style="color:${sevColor};font-weight:900">${sev.toFixed(2)}x</td><td>${fsec(r.estimated_cut_sec||0)}</td><td>${Number(r.cut_distance_in||0).toFixed(1)} in</td><td>${Number(r.avg_feed_ipm||0).toFixed(1)}</td><td>${Number(r.avg_rpm||0).toFixed(0)}</td><td>${Number(r.max_depth_in||0).toFixed(3)}</td><td>${Number(r.avg_stepdown_in||0).toFixed(3)}</td><td>${r.drill_cycles||0}</td><td>${r.peck_cycles||0}</td></tr>`}).join('')}</tbody></table></div>`;}
function renderTelemetryDebug(d){const el=document.getElementById('telemetry-debug');if(!el)return;const rows=d.telemetry_samples||[];el.innerHTML=`<table class="itable"><thead><tr><th>TIME</th><th>MACHINE</th><th>SOURCE</th><th>STATUS</th><th>PROGRAM</th><th>TOOL</th><th>RPM</th><th>FEED</th><th>LOAD</th><th>MT ERR</th></tr></thead><tbody>${rows.slice(0,120).map(r=>`<tr><td class="mono">${esc(r.timestamp||'')}</td><td>${esc(r.machine||'')}</td><td>${esc(r.source||'')}</td><td>${esc(r.status||'')}</td><td>${esc(r.prog_raw||'')}</td><td>${String(r.tool||'0')==='0'?'—':'T'+esc(r.tool)}</td><td>${Math.round(r.spindle_rpm||0)}</td><td>${Math.round(r.feed_rate||0)}</td><td>${Number(r.spindle_load||0).toFixed(1)}</td><td style="color:var(--red)">${esc(r.mtconnect_error||'')}</td></tr>`).join('')}</tbody></table>`;}

function safeBind(id,event,fn){const el=document.getElementById(id);if(el)el.addEventListener(event,fn);}
safeBind('sku-input','keydown',e=>{if(e.key==='Enter')submitScan();});
safeBind('manual-barcode','keydown',e=>{if(e.key==='Enter')manualBarcodeSubmit();});
safeBind('corner-admin-pass','keydown',e=>{if(e.key==='Enter')cornerAdminLogin();});
['new-sku','new-desc','new-cat','new-qty','new-min'].forEach(id=>safeBind(id,'keydown',e=>{if(e.key==='Enter')addInventoryItem();}));
</script>
</body>
</html>"""


@app.get("/")
async def root():
    return HTMLResponse(HTML + MATCH_STUDIO_NAV_INJECT)





# ============================================================
# MATCH STUDIO  (added integration)
# Beautiful job-matching workspace: matching list, holder JPG
# overlay compare, quote estimate, steel/quote sheet links, and
# one-click "order steel" email. Reuses the existing matching
# engine, job_series financials and steel_sheet_items.
# ============================================================

import urllib.parse as _ms_urlparse

# ---- Match Studio configuration (override via environment) ----
ELGIN_STEEL_VENDOR_EMAIL = os.environ.get("ELGIN_STEEL_VENDOR_EMAIL", "steel-vendor@example.com")
ELGIN_STEEL_PRICE_PER_LB = float(os.environ.get("ELGIN_STEEL_PRICE_PER_LB", "1.50"))
ELGIN_STEEL_DENSITY_LB_IN3 = float(os.environ.get("ELGIN_STEEL_DENSITY_LB_IN3", "0.283"))
ELGIN_PULLCORE_RATE = float(os.environ.get("ELGIN_PULLCORE_RATE", "77.98"))
ELGIN_JOB_IMAGE_DIRS = [
    p.strip() for p in os.environ.get("ELGIN_JOB_IMAGE_DIRS", ELGIN_MATCHING_SOFTWARE_DIR).split(";") if p.strip()
]
ELGIN_QUOTE_SHEET_DIRS = [
    p.strip() for p in os.environ.get("ELGIN_QUOTE_SHEET_DIRS", ELGIN_MATCHING_SOFTWARE_DIR).split(";") if p.strip()
]
_MS_IMG_EXTS = (".jpg", ".jpeg", ".png", ".gif", ".webp")
_MS_SHEET_EXTS = (".xls", ".xlsx", ".xlsm", ".pdf")
_ms_img_cache = {}


def ms_fmt_num(v) -> str:
    try:
        f = float(v)
    except Exception:
        return str(v)
    if abs(f - round(f)) < 1e-6:
        return str(int(round(f)))
    return ("%.3f" % f).rstrip("0").rstrip(".")


def ms_classify_image(fn: str) -> str:
    u = fn.upper()
    if "HOLDER" in u and ("OD" in u or "BOT" in u):
        return "OD HOLDER"
    if "HOLDER" in u and ("ID" in u or "TOP" in u):
        return "ID HOLDER"
    if "HOLDER" in u:
        return "ID HOLDER"
    if "POT" in u and ("OD" in u or "BOT" in u):
        return "OD POT"
    if "POT" in u and ("ID" in u or "TOP" in u):
        return "ID POT"
    if "TCP" in u or "TOP CLAMP" in u:
        return "TCP"
    if "BCP" in u or "BOTTOM CLAMP" in u:
        return "BCP"
    if "BACK" in u and "ISO" in u:
        return "BACK ISO"
    if "ISO" in u:
        return "ISO"
    return "OTHER"


def ms_find_job_images(job_num: str):
    j = normalize_job_num(job_num) or (job_num or "")
    key = j.upper()
    now = time.time()
    cached = _ms_img_cache.get(key)
    if cached and (now - cached[0]) < 30:
        return cached[1]

    jnorm = re.sub(r"[^A-Z0-9]", "", key)
    found = []
    seen = set()
    if jnorm:
        for base in ELGIN_JOB_IMAGE_DIRS:
            if not base or not os.path.isdir(base):
                continue
            scanned = 0
            for root, dirs, files in os.walk(base):
                depth = root[len(base):].count(os.sep)
                if depth > 4:
                    dirs[:] = []
                    continue
                for fn in files:
                    ext = os.path.splitext(fn)[1].lower()
                    if ext not in _MS_IMG_EXTS:
                        continue
                    fnnorm = re.sub(r"[^A-Z0-9]", "", fn.upper())
                    if jnorm in fnnorm:
                        full = os.path.join(root, fn)
                        if full in seen:
                            continue
                        seen.add(full)
                        found.append({"kind": ms_classify_image(fn), "label": fn, "path": full})
                    scanned += 1
                    if scanned > 8000:
                        break
                if scanned > 8000:
                    break

    order = {"ID HOLDER": 0, "OD HOLDER": 1, "ID POT": 2, "OD POT": 3,
             "TCP": 4, "BCP": 5, "ISO": 6, "BACK ISO": 7, "OTHER": 9}
    found.sort(key=lambda x: (order.get(x["kind"], 8), x["label"]))
    _ms_img_cache[key] = (now, found)
    return found


def ms_locate_sheet(j: str, kind: str) -> str:
    j = normalize_job_num(j) or (j or "")
    jnorm = re.sub(r"[^A-Z0-9]", "", j.upper())
    if kind == "steel":
        try:
            with db_lock, sqlite3.connect(DB_PATH) as conn:
                r = conn.execute("SELECT steel_sheet_file FROM job_series WHERE job_num=?", (j,)).fetchone()
            if r and r[0]:
                cand = os.path.join(STEEL_UPLOAD_DIR, r[0])
                if os.path.isfile(cand):
                    return cand
        except Exception:
            pass
        keywords = ("STEEL", "J000")
    else:
        keywords = ("QUOTE", "GRINDING")

    if not jnorm:
        return ""
    for base in ELGIN_QUOTE_SHEET_DIRS:
        if not base or not os.path.isdir(base):
            continue
        scanned = 0
        for root, dirs, files in os.walk(base):
            depth = root[len(base):].count(os.sep)
            if depth > 4:
                dirs[:] = []
                continue
            for fn in files:
                ext = os.path.splitext(fn)[1].lower()
                if ext not in _MS_SHEET_EXTS:
                    continue
                up = fn.upper()
                fnnorm = re.sub(r"[^A-Z0-9]", "", up)
                if jnorm in fnnorm and any(k in up for k in keywords):
                    return os.path.join(root, fn)
                scanned += 1
                if scanned > 8000:
                    break
            if scanned > 8000:
                break
    return ""


def ms_estimate(conn, j: str) -> dict:
    items = _rowdicts(conn.execute("SELECT * FROM steel_sheet_items WHERE job_num=?", (j,)))
    steel_vol = 0.0
    pc_vol = 0.0
    for it in items:
        try:
            d1 = float(it.get("dim_1") or 0)
            d2 = float(it.get("dim_2") or 0)
            d3 = float(it.get("dim_3") or 0)
            q = float(it.get("qty") or 0) or 1.0
        except Exception:
            continue
        v = d1 * d2 * d3 * q
        if v <= 0:
            continue
        nm = (it.get("item_name") or "").upper()
        if "CAM" in nm or "KEY" in nm:
            pc_vol += v
        else:
            steel_vol += v

    weight = steel_vol * ELGIN_STEEL_DENSITY_LB_IN3
    steel_mat = weight * ELGIN_STEEL_PRICE_PER_LB
    pc_price = pc_vol * ELGIN_PULLCORE_RATE

    js = conn.execute(
        "SELECT revenue, grinding_cost, status, due_date FROM job_series WHERE job_num=?", (j,)
    ).fetchone()
    revenue = _money(js["revenue"]) if js and js["revenue"] is not None else 0.0
    grinding = _money(js["grinding_cost"]) if js and js["grinding_cost"] is not None else 0.0

    fin = compute_job_financial(conn, j) or {}
    labor_cost = float(fin.get("labor_cost") or 0.0)

    computed = round(steel_mat + pc_price + grinding + labor_cost, 2)
    suggested = revenue if revenue > 0 else computed

    return {
        "revenue": round(revenue, 2),
        "suggested_estimate": round(suggested, 2),
        "computed_estimate": computed,
        "steel_volume_cuin": round(steel_vol, 2),
        "steel_weight_lb": round(weight, 2),
        "steel_material_est": round(steel_mat, 2),
        "price_per_lb": ELGIN_STEEL_PRICE_PER_LB,
        "pullcore_volume_cuin": round(pc_vol, 2),
        "pullcore_est": round(pc_price, 2),
        "pullcore_rate": ELGIN_PULLCORE_RATE,
        "grinding_cost": round(grinding, 2),
        "labor_cost": round(labor_cost, 2),
        "total_cost": float(fin.get("total_cost") or 0.0),
        "profit": float(fin.get("profit") or 0.0),
        "margin_pct": float(fin.get("margin_pct") or 0.0),
        "item_count": len(items),
        "status": (js["status"] if js else "") or "",
        "due_date": (js["due_date"] if js else "") or "",
    }


def ms_order_email(conn, j: str) -> dict:
    items = _rowdicts(conn.execute("SELECT * FROM steel_sheet_items WHERE job_num=? ORDER BY id", (j,)))
    lines = []
    for it in items:
        q = it.get("qty") or 0
        nm = it.get("item_name") or ""
        dims = " x ".join(
            ms_fmt_num(d) for d in (it.get("dim_1"), it.get("dim_2"), it.get("dim_3"))
            if (float(d or 0) > 0)
        )
        mat = it.get("material") or ""
        line = ms_fmt_num(q) + " pc  " + str(nm).strip()
        if dims:
            line += "   " + dims
        if mat:
            line += "   [" + str(mat).strip() + "]"
        lines.append(line.strip())

    body = (
        "Hello,\n\nPlease quote / supply the following steel for Job " + j + ":\n\n"
        + ("\n".join(lines) if lines else "(No parsed steel items are on file for this job yet.)")
        + "\n\nPlease confirm pricing and lead time.\n\nThank you,\nElgin Custom Mold Services"
    )
    subject = "Steel Order - Job " + j
    to = ELGIN_STEEL_VENDOR_EMAIL
    mailto = (
        "mailto:" + _ms_urlparse.quote(to)
        + "?subject=" + _ms_urlparse.quote(subject)
        + "&body=" + _ms_urlparse.quote(body)
    )
    return {"to": to, "subject": subject, "body": body, "mailto": mailto, "item_count": len(items)}


# ---------------- Match Studio API ----------------
@app.get("/api/match/overview")
async def ms_overview(limit: int = 80):
    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        try:
            sync_matching_software_signatures(conn)
        except Exception:
            pass
        jobs = _rowdicts(conn.execute(
            """SELECT job_num, MAX(imported_at) AS imported_at
               FROM job_signature_components
               GROUP BY job_num ORDER BY imported_at DESC LIMIT ?""",
            (max(1, int(limit)),),
        ))
        out = []
        for jr in jobs:
            j = normalize_job_num(jr.get("job_num") or "")
            if not j:
                continue
            best_list = compare_job_signature_to_previous(conn, j, limit=1)
            best = best_list[0] if best_list else None
            est = ms_estimate(conn, j)
            out.append({
                "job_num": j,
                "best_match": (best or {}).get("job_num", ""),
                "best_score": (best or {}).get("score", 0),
                "matched_components": (best or {}).get("matched_components", 0),
                "revenue": est.get("revenue", 0),
                "suggested_estimate": est.get("suggested_estimate", 0),
                "steel_material_est": est.get("steel_material_est", 0),
                "status": est.get("status", ""),
                "due_date": est.get("due_date", ""),
            })
    return {"jobs": out, "count": len(out)}


@app.get("/api/match/job/{job_num}")
async def ms_job(job_num: str, limit: int = 12):
    j = normalize_job_num(job_num)
    if not j:
        raise HTTPException(400, "Job number required")
    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        try:
            sync_matching_software_signatures(conn)
        except Exception:
            pass
        current = load_job_signature(conn, j)
        matches = compare_job_signature_to_previous(conn, j, limit=limit)
        est = ms_estimate(conn, j)
        email = ms_order_email(conn, j)
    return {
        "job_num": j,
        "has_signature": bool(current),
        "signature_components": sorted(current.keys()),
        "components": JOB_SIGNATURE_COMPONENTS,
        "matches": matches,
        "estimate": est,
        "order_email": email,
        "sheets": {
            "steel": bool(ms_locate_sheet(j, "steel")),
            "quote": bool(ms_locate_sheet(j, "quote")),
        },
    }


@app.get("/api/match/images/{job_num}")
async def ms_images_list(job_num: str):
    imgs = ms_find_job_images(job_num)
    out = []
    bykind = {}
    for x in imgs:
        k = x["kind"]
        idx = bykind.get(k, 0)
        out.append({"kind": k, "label": x["label"], "index": idx})
        bykind[k] = idx + 1
    return {"job_num": normalize_job_num(job_num), "images": out, "count": len(imgs)}


@app.get("/api/match/image")
async def ms_image(job: str, kind: str = "", i: int = 0):
    from fastapi.responses import FileResponse
    imgs = ms_find_job_images(job)
    if kind:
        filtered = [x for x in imgs if x["kind"] == kind]
        if filtered:
            imgs = filtered
    if not imgs:
        raise HTTPException(404, "No image found for this job")
    idx = max(0, min(int(i), len(imgs) - 1))
    path = imgs[idx]["path"]
    if not os.path.isfile(path):
        raise HTTPException(404, "Image file missing")
    return FileResponse(path)


@app.get("/api/match/file")
async def ms_file(job: str, kind: str = "steel"):
    from fastapi.responses import FileResponse
    path = ms_locate_sheet(job, kind)
    if not path or not os.path.isfile(path):
        raise HTTPException(404, "Sheet not found for this job")
    return FileResponse(path, filename=os.path.basename(path))


@app.get("/api/match/order-email/{job_num}")
async def ms_order_email_route(job_num: str):
    j = normalize_job_num(job_num)
    if not j:
        raise HTTPException(400, "Job number required")
    with db_lock, sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        return ms_order_email(conn, j)


@app.get("/match-studio")
async def ms_page():
    return HTMLResponse(MATCH_STUDIO_HTML)


MATCH_STUDIO_NAV_INJECT = """
<script>
(function(){
  if (document.getElementById('cms-match-studio-fab')) return;
  var a = document.createElement('a');
  a.id = 'cms-match-studio-fab';
  a.href = '/match-studio';
  a.title = 'Open Match Studio';
  a.textContent = '\\u2316 Match Studio';
  a.style.cssText = 'position:fixed;right:18px;bottom:18px;z-index:99999;background:linear-gradient(135deg,#2563eb,#1e40af);color:#fff;padding:12px 18px;border-radius:999px;font:600 14px system-ui,-apple-system,sans-serif;box-shadow:0 8px 24px rgba(30,64,175,.45);text-decoration:none;letter-spacing:.2px;';
  if (document.body) document.body.appendChild(a);
})();
</script>
"""


MATCH_STUDIO_HTML = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>CMS Match Studio</title>
<style>
:root{
  --bg:#0b1220; --panel:#111a2e; --panel2:#0f1729; --line:#22304d;
  --ink:#e8eefc; --muted:#8aa0c6; --accent:#3b82f6; --accent2:#1e40af;
  --good:#22c55e; --warn:#f59e0b; --bad:#ef4444; --chip:#16223c;
}
*{box-sizing:border-box}
body{margin:0;background:radial-gradient(1200px 600px at 80% -10%,#16223f 0%,var(--bg) 55%);color:var(--ink);font:14px/1.45 system-ui,-apple-system,Segoe UI,Roboto,sans-serif}
a{color:var(--accent);text-decoration:none}
.top{display:flex;align-items:center;gap:18px;padding:14px 22px;border-bottom:1px solid var(--line);background:rgba(10,16,30,.7);backdrop-filter:blur(8px);position:sticky;top:0;z-index:10}
.brand{font-weight:800;font-size:18px;letter-spacing:.3px;display:flex;align-items:center;gap:10px}
.brand .dot{width:11px;height:11px;border-radius:3px;background:linear-gradient(135deg,var(--accent),var(--accent2));box-shadow:0 0 14px var(--accent)}
.top .sub{color:var(--muted);font-size:12px}
.spacer{flex:1}
.search{display:flex;gap:8px;align-items:center}
.search input{background:var(--panel2);border:1px solid var(--line);color:var(--ink);padding:9px 12px;border-radius:9px;width:170px;outline:none}
.btn{background:linear-gradient(135deg,var(--accent),var(--accent2));color:#fff;border:0;padding:9px 14px;border-radius:9px;font-weight:600;cursor:pointer}
.btn.ghost{background:var(--chip);border:1px solid var(--line);color:var(--ink)}
.btn.ghost:hover{border-color:var(--accent)}
.btn:disabled{opacity:.4;cursor:not-allowed}
.wrap{display:grid;grid-template-columns:340px 1fr;gap:18px;padding:18px;max-width:1500px;margin:0 auto}
.side{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:14px;height:calc(100vh - 110px);overflow:auto;position:sticky;top:86px}
.side h3{margin:2px 0 10px;font-size:13px;text-transform:uppercase;letter-spacing:.8px;color:var(--muted)}
.filter{width:100%;background:var(--panel2);border:1px solid var(--line);color:var(--ink);padding:9px 11px;border-radius:9px;margin-bottom:10px;outline:none}
.jobcard{border:1px solid var(--line);background:var(--panel2);border-radius:11px;padding:11px 12px;margin-bottom:9px;cursor:pointer;transition:.15s}
.jobcard:hover{border-color:var(--accent);transform:translateY(-1px)}
.jobcard.active{border-color:var(--accent);box-shadow:0 0 0 1px var(--accent) inset}
.jobcard .jn{font-weight:700;font-size:15px}
.jobcard .row{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-top:6px;font-size:12px;color:var(--muted)}
.bar{height:7px;border-radius:6px;background:#1b2741;overflow:hidden;margin-top:7px}
.bar > i{display:block;height:100%;border-radius:6px}
.pill{padding:2px 8px;border-radius:999px;background:var(--chip);border:1px solid var(--line);font-size:11px;color:var(--ink)}
.main{min-width:0}
.card{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px;margin-bottom:16px}
.card h2{margin:0 0 4px;font-size:18px}
.muted{color:var(--muted)}
.empty{display:flex;align-items:center;justify-content:center;height:60vh;color:var(--muted);text-align:center;flex-direction:column;gap:8px}
.estgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-top:14px}
.kpi{background:var(--panel2);border:1px solid var(--line);border-radius:11px;padding:12px 14px}
.kpi .lbl{font-size:11px;text-transform:uppercase;letter-spacing:.6px;color:var(--muted)}
.kpi .val{font-size:20px;font-weight:800;margin-top:3px}
.kpi.big .val{font-size:30px}
.kpi.big{background:linear-gradient(135deg,#16223f,#101a30);border-color:#2a3c63}
.actions{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-top:6px}
select.dd{background:var(--panel2);border:1px solid var(--line);color:var(--ink);padding:9px 12px;border-radius:9px;outline:none}
.matchrow{display:flex;align-items:center;gap:12px;padding:11px 12px;border:1px solid var(--line);background:var(--panel2);border-radius:11px;margin-bottom:9px;cursor:pointer;transition:.15s}
.matchrow:hover{border-color:var(--accent)}
.matchrow.active{border-color:var(--accent);box-shadow:0 0 0 1px var(--accent) inset}
.matchrow .jn{font-weight:700;min-width:90px}
.matchrow .grow{flex:1}
.scorebadge{font-weight:800;min-width:58px;text-align:right}
.tbl{width:100%;border-collapse:collapse;margin-top:6px}
.tbl th,.tbl td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--line);font-size:13px}
.tbl th{color:var(--muted);font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.5px}
.simbar{height:8px;border-radius:6px;background:#1b2741;overflow:hidden;min-width:90px}
.simbar > i{display:block;height:100%}
.cmpbar{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}
.chipbtn{padding:7px 12px;border-radius:9px;border:1px solid var(--line);background:var(--panel2);color:var(--ink);cursor:pointer;font-size:12px}
.chipbtn.active{background:var(--accent);border-color:var(--accent);color:#fff}
.ovstage{display:grid;grid-template-columns:1fr;gap:14px}
.ovwrap{position:relative;width:100%;max-width:680px;margin:0 auto;border:1px solid var(--line);border-radius:12px;background:#0a101e;min-height:240px;display:flex;align-items:center;justify-content:center;overflow:hidden}
.ovimg{max-width:100%;display:block}
.ovtop{position:absolute;inset:0;margin:auto;max-width:100%;max-height:100%;mix-blend-mode:normal}
.sbs{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.imgbox{border:1px solid var(--line);border-radius:12px;background:#0a101e;min-height:220px;display:flex;flex-direction:column}
.imgbox .cap{padding:8px 12px;border-bottom:1px solid var(--line);font-size:12px;color:var(--muted);display:flex;justify-content:space-between}
.imgbox .body{flex:1;display:flex;align-items:center;justify-content:center;padding:8px}
.imgbox img{max-width:100%;max-height:360px}
.ph{color:#56678c;font-size:13px;text-align:center;padding:30px}
.slider{display:flex;align-items:center;gap:12px;max-width:680px;margin:12px auto 0}
.slider input{flex:1}
.toggle{display:flex;gap:6px}
.legend{font-size:11px;color:var(--muted);margin-top:8px}
.note{font-size:12px;color:var(--muted);margin-top:8px}
code{background:#0a101e;border:1px solid var(--line);padding:1px 6px;border-radius:6px}
</style>
</head>
<body>
<div class="top">
  <div class="brand"><span class="dot"></span> CMS Match Studio</div>
  <div class="sub">job matching &middot; holder overlay &middot; quote estimate &middot; steel order</div>
  <div class="spacer"></div>
  <div class="search">
    <input id="jobInput" placeholder="Job # (e.g. C18488)" onkeydown="if(event.key==='Enter')openJob(this.value)">
    <button class="btn" onclick="openJob(document.getElementById('jobInput').value)">Open</button>
  </div>
  <a class="btn ghost" href="/">&larr; App</a>
</div>

<div class="wrap">
  <div class="side">
    <h3>Matching List</h3>
    <input class="filter" id="listFilter" placeholder="Filter jobs..." oninput="renderList()">
    <div id="jobList"><div class="muted">Loading…</div></div>
  </div>
  <div class="main" id="main">
    <div class="empty">
      <div style="font-size:40px">&#9166;</div>
      <div>Pick a job from the matching list, or search a job number above.</div>
    </div>
  </div>
</div>

<script>
var STATE={jobs:[],job:null,data:null,match:null,cmpKind:'ID HOLDER',cmpMode:'overlay',imgCur:{},imgMatch:{}};
var COMPONENTS=['TCP','BCP','ID HOLDER','OD HOLDER','ID POT','OD POT'];

function api(p){return fetch(p).then(function(r){if(!r.ok)throw new Error(r.status);return r.json();});}
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function money(v){var n=Number(v||0);return '$'+n.toLocaleString(undefined,{minimumFractionDigits:0,maximumFractionDigits:0});}
function money2(v){var n=Number(v||0);return '$'+n.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});}
function scoreColor(s){var h=Math.max(0,Math.min(120,(Number(s)||0)*1.2));return 'hsl('+h+',75%,52%)';}
function imgURL(job,kind){return '/api/match/image?job='+encodeURIComponent(job)+'&kind='+encodeURIComponent(kind)+'&i=0';}

function loadOverview(){
  api('/api/match/overview?limit=120').then(function(d){
    STATE.jobs=d.jobs||[];renderList();
  }).catch(function(){document.getElementById('jobList').innerHTML='<div class="muted">No signatures imported yet.</div>';});
}

function renderList(){
  var f=(document.getElementById('listFilter').value||'').toUpperCase();
  var host=document.getElementById('jobList');
  var rows=STATE.jobs.filter(function(j){return !f||j.job_num.toUpperCase().indexOf(f)>=0||(j.best_match||'').toUpperCase().indexOf(f)>=0;});
  if(!rows.length){host.innerHTML='<div class="muted">No matching jobs.</div>';return;}
  host.innerHTML=rows.map(function(j){
    var sc=Number(j.best_score||0);
    var est=j.suggested_estimate||j.steel_material_est||0;
    return '<div class="jobcard'+(STATE.job===j.job_num?' active':'')+'" onclick="openJob(\''+esc(j.job_num)+'\')">'
      +'<div class="jn">'+esc(j.job_num)+'</div>'
      +'<div class="row"><span>'+(j.best_match?('best: '+esc(j.best_match)):'no match')+'</span><span class="pill">'+(est?money(est):'--')+'</span></div>'
      +(j.best_match?('<div class="bar"><i style="width:'+sc+'%;background:'+scoreColor(sc)+'"></i></div>'):'')
      +'</div>';
  }).join('');
}

function openJob(job){
  job=(job||'').trim();if(!job)return;
  STATE.job=job;STATE.match=null;STATE.data=null;renderList();
  document.getElementById('main').innerHTML='<div class="empty">Loading '+esc(job)+'…</div>';
  api('/api/match/job/'+encodeURIComponent(job)).then(function(d){
    STATE.data=d;STATE.job=d.job_num;renderJob();renderList();
  }).catch(function(){document.getElementById('main').innerHTML='<div class="empty">Could not load '+esc(job)+'.</div>';});
}

function renderJob(){
  var d=STATE.data,e=d.estimate||{};
  var sig=d.has_signature?'<span class="pill" style="border-color:#2b6">&#10003; signature ('+d.signature_components.length+'/'+COMPONENTS.length+')</span>':'<span class="pill" style="border-color:#a55">no signature</span>';
  var quoted=e.revenue>0;
  var html=''
  +'<div class="card">'
    +'<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap">'
      +'<div><h2>Job '+esc(d.job_num)+'</h2><div class="muted">'+(e.status?esc(e.status):'')+(e.due_date?(' &middot; due '+esc(e.due_date)):'')+'</div></div>'
      +'<div>'+sig+'</div>'
    +'</div>'
    +'<div class="estgrid">'
      +'<div class="kpi big"><div class="lbl">'+(quoted?'Quoted price':'Estimated price')+'</div><div class="val">'+money(e.suggested_estimate)+'</div></div>'
      +'<div class="kpi"><div class="lbl">Steel material</div><div class="val">'+money(e.steel_material_est)+'</div><div class="muted" style="font-size:11px">'+ (e.steel_weight_lb||0)+' lb @ $'+e.price_per_lb+'/lb</div></div>'
      +'<div class="kpi"><div class="lbl">Pull cores</div><div class="val">'+money(e.pullcore_est)+'</div><div class="muted" style="font-size:11px">'+(e.pullcore_volume_cuin||0)+' in&sup3;</div></div>'
      +'<div class="kpi"><div class="lbl">Total cost</div><div class="val">'+money(e.total_cost)+'</div></div>'
      +'<div class="kpi"><div class="lbl">Profit</div><div class="val" style="color:'+((e.profit||0)>=0?'var(--good)':'var(--bad)')+'">'+money(e.profit)+'</div><div class="muted" style="font-size:11px">'+(e.margin_pct||0)+'% margin</div></div>'
    +'</div>'
    +'<div class="actions" style="margin-top:14px">'
      +'<select class="dd" id="sheetDD"><option value="">Open a sheet…</option>'
        +'<option value="steel"'+(d.sheets.steel?'':' disabled')+'>Steel Order Sheet'+(d.sheets.steel?'':' (not found)')+'</option>'
        +'<option value="quote"'+(d.sheets.quote?'':' disabled')+'>Quote / Steel-Grinding Sheet'+(d.sheets.quote?'':' (not found)')+'</option>'
      +'</select>'
      +'<button class="btn ghost" onclick="openSheet()">Open</button>'
      +'<button class="btn" onclick="orderSteel()">&#9993; Order Steel</button>'
      +'<button class="btn ghost" onclick="copyEmail()">Copy email</button>'
      +'<span class="note" id="actNote"></span>'
    +'</div>'
    +(quoted?'':'<div class="note">No quote saved on this job &mdash; showing a material-based estimate (steel + pull cores + grinding + labor). Set $/lb via <code>ELGIN_STEEL_PRICE_PER_LB</code>.</div>')
  +'</div>';

  // matches
  html+='<div class="card"><h2>Job Matches</h2><div class="muted">Ranked by signature similarity. Click one to compare holders.</div><div id="matchList" style="margin-top:12px"></div></div>';
  html+='<div id="compareHost"></div>';
  document.getElementById('main').innerHTML=html;
  renderMatches();
}

function renderMatches(){
  var ms=STATE.data.matches||[];
  var host=document.getElementById('matchList');
  if(!ms.length){host.innerHTML='<div class="muted">No previous jobs to compare against yet.</div>';return;}
  host.innerHTML=ms.map(function(m){
    var sc=Number(m.score||0);
    return '<div class="matchrow'+(STATE.match===m.job_num?' active':'')+'" onclick="selectMatch(\''+esc(m.job_num)+'\')">'
      +'<div class="jn">'+esc(m.job_num)+'</div>'
      +'<div class="grow"><div class="bar"><i style="width:'+sc+'%;background:'+scoreColor(sc)+'"></i></div>'
      +'<div class="muted" style="font-size:11px;margin-top:4px">'+m.matched_components+'/'+COMPONENTS.length+' parts &middot; '+m.coverage_pct+'% coverage</div></div>'
      +'<div class="scorebadge" style="color:'+scoreColor(sc)+'">'+sc.toFixed(0)+'%</div>'
      +'</div>';
  }).join('');
}

function selectMatch(mj){
  STATE.match=mj;renderMatches();
  var m=(STATE.data.matches||[]).filter(function(x){return x.job_num===mj;})[0];
  if(!m)return;
  // pick default compare component that has data
  var comps=m.components||[];
  var have=comps.map(function(c){return c.component_role;});
  if(have.indexOf(STATE.cmpKind)<0)STATE.cmpKind=have.indexOf('ID HOLDER')>=0?'ID HOLDER':(have[0]||'ID HOLDER');
  renderCompare(m);
}

function renderCompare(m){
  var rows=(m.components||[]).slice().sort(function(a,b){return COMPONENTS.indexOf(a.component_role)-COMPONENTS.indexOf(b.component_role);});
  var tbl='<table class="tbl"><thead><tr><th>Component</th><th>Similarity</th><th>Size &Delta;</th><th>Mass &Delta;</th><th>COG dist</th></tr></thead><tbody>';
  rows.forEach(function(c){
    var s=Number(c.similarity||0);
    tbl+='<tr><td><b>'+esc(c.component_role)+'</b></td>'
      +'<td><div style="display:flex;align-items:center;gap:8px"><div class="simbar"><i style="width:'+s+'%;background:'+scoreColor(s)+'"></i></div><span>'+s.toFixed(1)+'%</span></div></td>'
      +'<td>'+(c.size_diff_pct)+'%</td><td>'+(c.mass_diff_pct)+'%</td><td>'+(c.center_distance_in)+'"</td></tr>';
  });
  tbl+='</tbody></table>';

  var chips=COMPONENTS.map(function(k){
    return '<button class="chipbtn'+(STATE.cmpKind===k?' active':'')+'" onclick="setCmpKind(\''+k+'\')">'+k+'</button>';
  }).join('');

  var html='<div class="card"><div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px">'
    +'<h2>Compare: '+esc(STATE.job)+' vs '+esc(m.job_num)+'</h2>'
    +'<div class="toggle"><button class="chipbtn'+(STATE.cmpMode==='overlay'?' active':'')+'" onclick="setCmpMode(\'overlay\')">Overlay</button>'
    +'<button class="chipbtn'+(STATE.cmpMode==='sbs'?' active':'')+'" onclick="setCmpMode(\'sbs\')">Side by side</button></div></div>'
    +tbl
    +'<div style="margin-top:16px"><div class="muted" style="font-size:12px;text-transform:uppercase;letter-spacing:.6px;margin-bottom:8px">Holder / plate overlay</div>'
    +'<div class="cmpbar">'+chips+'</div>'
    +'<div id="ovStage"></div></div></div>';
  document.getElementById('compareHost').innerHTML=html;
  renderOverlay(m);
}

function setCmpKind(k){STATE.cmpKind=k;var m=curMatch();if(m){renderCompare(m);}}
function setCmpMode(md){STATE.cmpMode=md;var m=curMatch();if(m){renderOverlay(m);document.querySelectorAll('.toggle .chipbtn').forEach(function(b){b.classList.toggle('active',b.textContent.toLowerCase().indexOf(md==='sbs'?'side':'overlay')>=0);});}}
function curMatch(){return (STATE.data.matches||[]).filter(function(x){return x.job_num===STATE.match;})[0];}

function imgTag(job,kind,cls){
  var u=imgURL(job,kind);
  return '<img class="'+cls+'" src="'+u+'" onerror="this.style.display=\'none\';this.parentNode.querySelector(\'.ph\')&&(this.parentNode.querySelector(\'.ph\').style.display=\'block\')">';
}

function renderOverlay(m){
  var k=STATE.cmpKind, stage=document.getElementById('ovStage');
  if(STATE.cmpMode==='sbs'){
    stage.innerHTML='<div class="sbs">'
      +'<div class="imgbox"><div class="cap"><span>'+esc(STATE.job)+' (current)</span><span>'+esc(k)+'</span></div><div class="body">'+imgTag(STATE.job,k,'')+'<div class="ph" style="display:none">no '+esc(k)+' image</div></div></div>'
      +'<div class="imgbox"><div class="cap"><span>'+esc(m.job_num)+' (match)</span><span>'+esc(k)+'</span></div><div class="body">'+imgTag(m.job_num,k,'')+'<div class="ph" style="display:none">no '+esc(k)+' image</div></div></div>'
      +'</div>';
  }else{
    stage.innerHTML='<div class="ovstage">'
      +'<div class="ovwrap">'
        +'<img class="ovimg" src="'+imgURL(m.job_num,k)+'" onerror="this.style.visibility=\'hidden\'">'
        +'<img class="ovimg ovtop" id="ovTop" src="'+imgURL(STATE.job,k)+'" style="opacity:.5" onerror="this.style.visibility=\'hidden\'">'
        +'<div class="ph" id="ovPh" style="position:absolute">'+esc(k)+' images load here when present</div>'
      +'</div>'
      +'<div class="slider"><span class="muted">'+esc(m.job_num)+'</span>'
        +'<input type="range" min="0" max="100" value="50" oninput="document.getElementById(\'ovTop\').style.opacity=this.value/100">'
        +'<span class="muted">'+esc(STATE.job)+'</span></div>'
      +'<div class="legend">Slide to fade between the matched holder (left) and the current holder (right) to judge similarity.</div>'
    +'</div>';
  }
}

function openSheet(){
  var k=document.getElementById('sheetDD').value;if(!k)return;
  window.open('/api/match/file?job='+encodeURIComponent(STATE.job)+'&kind='+k,'_blank');
}
function orderSteel(){
  var em=(STATE.data&&STATE.data.order_email)||{};
  if(!em.mailto){note('No email available');return;}
  note('Opening your email app…');
  window.location.href=em.mailto;
}
function copyEmail(){
  var em=(STATE.data&&STATE.data.order_email)||{};
  var txt='To: '+(em.to||'')+'\nSubject: '+(em.subject||'')+'\n\n'+(em.body||'');
  if(navigator.clipboard){navigator.clipboard.writeText(txt).then(function(){note('Email copied to clipboard');});}
  else note('Copy not supported');
}
function note(t){var n=document.getElementById('actNote');if(n){n.textContent=t;setTimeout(function(){if(n.textContent===t)n.textContent='';},4000);}}

loadOverview();
</script>
</body>
</html>"""


# ============================================================
# STARTUP
# ============================================================
if __name__ == "__main__":
    for cfg in MACHINE_CONFIGS:
        threading.Thread(
            target=poll_single_machine,
            args=(cfg["name"], cfg["ip"], cfg["port"]),
            daemon=True,
        ).start()

    ssl_kwargs = {}

    cert_ok = (
        HTTPS_CERT_FILE
        and HTTPS_KEY_FILE
        and os.path.exists(HTTPS_CERT_FILE)
        and os.path.exists(HTTPS_KEY_FILE)
    )

    if cert_ok:
        ssl_kwargs = {
            "ssl_certfile": HTTPS_CERT_FILE,
            "ssl_keyfile": HTTPS_KEY_FILE,
        }

    launcher = write_windows_launcher_file()
    urls = get_local_network_addresses(APP_PORT)

    print("\n" + "=" * 72)
    print(f"{APP_NAME.upper()} READY")
    print("=" * 72)

    if cert_ok:
        print("\nHTTPS mode enabled.")
        print("Use HTTPS URLs below if the certificate is trusted by the device.")
    else:
        print("\nHTTP mode enabled.")
        print("Normal app pages work over HTTP.")
        print("iPhone live camera scanning requires HTTPS or ngrok.")
        print("PHONE PHOTO SCAN can still be used over HTTP.")

    print("\nOpen on this computer:")
    print(f"  http://127.0.0.1:{APP_PORT}")

    print("\nOpen from other computers on Ethernet/private LAN:")
    any_network = False
    for item in urls:
        if not item.get("is_localhost"):
            any_network = True
            if cert_ok:
                print(f"  {item['https_url']}")
            else:
                print(f"  {item['http_url']}")

    if not any_network:
        print("  No non-localhost network adapter IP detected.")
        print("  Connect the host computer to Ethernet/private LAN and restart if needed.")

    print("\nNetwork notes:")
    print("  - Only the host computer needs to run this Python app.")
    print("  - Other computers can open it in a browser if they can reach this host.")
    print("  - They do NOT need Wi-Fi if connected by Ethernet/private shop network.")
    print(f"  - Allow inbound TCP port {APP_PORT} in Windows Firewall.")

    if launcher:
        print("\nWindows launcher created:")
        print(f"  {launcher}")

    print("\nAdmin password default: cms123")
    print("Change it with:")
    print("  set ELGIN_ADMIN_PASSWORD=yourBetterPasswordHere")
    print("or:")
    print("  set NEXUS_ADMIN_PASSWORD=yourBetterPasswordHere")

    print("\nSystem features:")
    print("  - Haas Q-code + MTConnect hybrid polling enabled.")
    print("  - Smart runtime estimates use Q-code, G-code, then similar completed jobs.")
    print("  - ID and OD estimates are matched separately by component.")
    print("  - Old/reused programs can map to current matched jobs.")
    print("  - Tool checkout timing and usage analytics enabled.")
    print("  - QR scanner supports quantity add/subtract.")
    print("  - Steel sheet PDF upload is exposed in Job Matching only.")
    print("  - Tool inspection/calibration enabled.")
    print("  - UMC-1000 pot scheduler enabled.")
    print("  - Tool offset-change detection is confirm-only.")
    print("  - Autoloader virtual shelf is SIMULATION ONLY.")
    print("  - Financial/profit fields and Efficiency tab are admin-only.")
    print("  - All-zero fake jobs like J0000 are blocked from auto-creation.")
    print("  - Machine idle begins only after RPM has been zero/not valid for 5 minutes.")
    print("=" * 72 + "\n")

    uvicorn.run(
        app,
        host="0.0.0.0",
        port=2926,
        log_level="warning",
        **ssl_kwargs,
    )