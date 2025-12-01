import os
import re
import datetime
import time
import threading
import subprocess
import configparser
import logging
import json
import sqlite3
import hashlib
import uuid
from flask import Flask, render_template, jsonify, redirect, url_for, request, make_response
import ctypes
import sys

app = Flask(__name__)

# ---------- Logging Setup ----------
# Create logs directory if it doesn't exist
logs_dir = "logs"
if not os.path.exists(logs_dir):
    os.makedirs(logs_dir)

# Log file path
log_file = os.path.join(logs_dir, "Licenses_WebUI.log")

# Configure logging to both file and console
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)
logger.info("=" * 80)
logger.info("Licenses WebUI Started")
logger.info("=" * 80)
logger.info(f"Log file: {os.path.abspath(log_file)}")

# Suppress Werkzeug development server warnings but still log to file
werkzeug_log = logging.getLogger('werkzeug')
werkzeug_log.setLevel(logging.ERROR)


# Prevent subprocess console flash on Windows by default; fallback to empty kwargs on other OSes
if os.name == "nt":
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        _SUBP_NO_WINDOW_KW = {
            "creationflags": subprocess.CREATE_NO_WINDOW,
            "startupinfo": startupinfo
        }
    except Exception as e:
        logger.warning(f"Could not set no-window startup info: {e}")
        _SUBP_NO_WINDOW_KW = {}
else:
    _SUBP_NO_WINDOW_KW = {}

# ---------- Constants ----------
SERVICE_START_TIMEOUT = 30
SERVICE_STOP_TIMEOUT = 30
LMSTAT_TIMEOUT = 30
DEFAULT_LMUTIL = r"C:\Program Files (x86)\Leica Geosystems\CLM\lmutil.exe"
DEFAULT_PORT = "27008"
DEFAULT_WEB_PORT = 8080
DEFAULT_REFRESH = 5

# Windows service state codes (from sc output)
STATE_STOPPED = 1
STATE_START_PENDING = 2
STATE_STOP_PENDING = 3
STATE_RUNNING = 4

# ---------- Config ----------
cfg = configparser.ConfigParser(interpolation=None)
# Determine absolute config path next to script or executable (frozen)
if getattr(sys, 'frozen', False):
    _BASE_DIR = os.path.dirname(sys.executable)
else:
    _BASE_DIR = os.path.dirname(__file__)
config_path = os.path.join(_BASE_DIR, "config.ini")
logger.debug(f"Resolved config.ini path: {config_path}")

# Known default configuration to ensure missing keys are auto-filled
DEFAULT_CONFIG = {
    "SETTINGS": {
        "lmutil_path": DEFAULT_LMUTIL,
        "port": DEFAULT_PORT,
        "web_port": str(DEFAULT_WEB_PORT),
        "refresh_minutes": str(DEFAULT_REFRESH),
        "default_locale": "en",
        "hide_maintenance": "yes",
        "hide_list": "",
        "enable_restart": "no",
        "show_eid_info": "no",
    },
    "SERVICE": {
        "service_name": "FLEXnet License Server",
    },
    "TEAMS": {
        "enabled": "no",
        "webhook": "",
        "notify_update": "no",
        "notify_duplicate_checker": "no",
        "notify_extratime": "no",
        "extratime_duration": "72",
        "extratime_exclusion": "",
        "notify_soldout": "no",
        "soldout_exclusion": "",
        "notify_daemon": "no",
    },
}

def ensure_config_defaults(cfg_obj: configparser.ConfigParser, defaults: dict, path: str) -> None:
    """Ensure all known keys exist; write back to disk if anything was missing.

    - Preserves existing values
    - Adds any missing sections/keys with default values
    - Writes atomically and creates a .bak backup of the previous file
    """
    added = []
    changed = False
    # Ensure sections and keys
    for section, keys in defaults.items():
        if not cfg_obj.has_section(section):
            cfg_obj.add_section(section)
            changed = True
            added.append(f"[{section}]")
        for key, val in keys.items():
            if not cfg_obj.has_option(section, key) or cfg_obj.get(section, key) == "":
                cfg_obj.set(section, key, str(val))
                changed = True
                added.append(f"{section}.{key}")

    if not changed:
        logger.debug("ensure_config_defaults: no missing keys detected; no rewrite needed")
        return

    # Only write if the file content would actually change
    import io
    new_buf = io.StringIO()
    cfg_obj.write(new_buf)
    new_content = new_buf.getvalue()
    old_content = None
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            old_content = f.read()
    if old_content is not None and old_content.strip() == new_content.strip():
        logger.debug("ensure_config_defaults: config file content unchanged; no rewrite needed")
        return

    # Prepare atomic write
    tmp_path = path + ".tmp"
    bak_path = path + ".bak"
    try:
        # Backup existing file if present
        if os.path.exists(path):
            try:
                # Use replace to ensure bak is the latest previous copy
                if os.path.exists(bak_path):
                    os.remove(bak_path)
                import shutil
                shutil.copy2(path, bak_path)
                logger.info(f"Backed up existing config to: {os.path.abspath(bak_path)}")
            except Exception as e:
                logger.warning(f"Could not create backup config: {e}")

        with open(tmp_path, "w", encoding="utf-8") as f:
            f.write(new_content)
        os.replace(tmp_path, path)
        logger.info(
            "Updated config.ini with missing defaults (%d additions): %s",
            len(added), ", ".join(added) if added else "<none>"
        )
    except Exception as e:
        # Clean up tmp if write failed
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        logger.warning(f"Failed to update config.ini with defaults: {e}")

def _generate_admin_key() -> str:
    """Generate a random admin key (MD5 hex)."""
    seed = f"{uuid.uuid4()}-{time.time()}-{os.getpid()}-{os.urandom(16).hex()}".encode("utf-8")
    return hashlib.md5(seed).hexdigest()

def ensure_admin_key(cfg_obj: configparser.ConfigParser, path: str) -> str:
    """Ensure SETTINGS.admin_key exists; create and persist if missing. Returns the key."""
    if not cfg_obj.has_section("SETTINGS"):
        cfg_obj.add_section("SETTINGS")
    key = cfg_obj.get("SETTINGS", "admin_key", fallback="").strip()
    if key:
        return key

    key = _generate_admin_key()
    cfg_obj.set("SETTINGS", "admin_key", key)

    # Write atomically with backup
    tmp_path = path + ".tmp"
    bak_path = path + ".bak"
    try:
        if os.path.exists(path):
            try:
                if os.path.exists(bak_path):
                    os.remove(bak_path)
                import shutil
                shutil.copy2(path, bak_path)
                logger.info(f"Backed up existing config to: {os.path.abspath(bak_path)}")
            except Exception as e:
                logger.warning(f"Could not create backup config: {e}")
        with open(tmp_path, "w", encoding="utf-8") as f:
            cfg_obj.write(f)
        os.replace(tmp_path, path)
        logger.info("Generated and saved new admin_key to config.ini")
    except Exception as e:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        logger.warning(f"Failed to persist admin_key to config.ini: {e}")
    return key

if os.path.exists(config_path):
    try:
        cfg.read(config_path)
        logger.info(f"Loaded config from: {os.path.abspath(config_path)}")
        # Ensure any missing keys are added, then rewrite file
        ensure_config_defaults(cfg, DEFAULT_CONFIG, config_path)
        # Ensure admin key exists and is persisted
        ADMIN_KEY = ensure_admin_key(cfg, config_path)
    except Exception as e:
        logger.warning(f"Error reading config.ini: {e}. Using defaults.")
        # Attempt to recover by loading defaults then forcing rewrite
        for section, keys in DEFAULT_CONFIG.items():
            if not cfg.has_section(section):
                cfg.add_section(section)
            for k, v in keys.items():
                if not cfg.has_option(section, k):
                    cfg.set(section, k, v)
        try:
            ADMIN_KEY = ensure_admin_key(cfg, config_path)
        except Exception:
            ADMIN_KEY = _generate_admin_key()
else:
    # Create a sample config.ini with sensible defaults so users can edit it next to the exe.
    sample = """; Licenses WebUI configuration
[SETTINGS]
# Path to lmutil.exe (example)
lmutil_path = C:\\Program Files (x86)\\Leica Geosystems\\CLM\\lmutil.exe

# License manager port
port = 27008

# Web UI port
web_port = 8080

# Auto-refresh interval in minutes
refresh_minutes = 5

# Default UI language (en, fr, de, es)
default_locale = en

# Hide features containing 'maint' (yes/no)
hide_maintenance = yes

# Comma-separated substrings of features to hide
hide_list =

# Enable the Restart service button (yes/no)
enable_restart = no

[SERVICE]
# Windows service name to restart
service_name = FLEXnet License Server

[TEAMS]
# Enable sending notifications to Microsoft Teams (yes/no)
enabled = no

# Incoming webhook URL for Microsoft Teams (set this to your connector URL)
webhook =

# Enable the "update available" notification (yes/no)
notify_update = no

# Enable duplicate user checker notification (yes/no)
notify_duplicate_checker = no

# Enable extratime notification (yes/no)
notify_extratime = no

# Extratime duration threshold in hours (default 72h)
extratime_duration = 72

# Comma-separated list of features to exclude from extratime notifications
extratime_exclusion =

# Enable sold-out notification (yes/no)
notify_soldout = no

# Comma-separated list of features to exclude from sold-out notifications
soldout_exclusion =

# Enable daemon status notifications (yes/no)
notify_daemon = no
"""
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            f.write(sample)
        # Re-read the newly created config so values are available
        cfg.read(config_path)
        logger.info(f"No config.ini found: created sample at {os.path.abspath(config_path)}")
        # After creating sample, still ensure keys align with known defaults
        ensure_config_defaults(cfg, DEFAULT_CONFIG, config_path)
        ADMIN_KEY = ensure_admin_key(cfg, config_path)
    except Exception as e:
        logger.warning(f"Failed to create sample config.ini: {e}. Continuing with defaults.")
        ADMIN_KEY = _generate_admin_key()
        logger.warning("Using ephemeral admin_key (not persisted) due to creation error.")

if 'ADMIN_KEY' not in globals():
    # Fallback in case previous block failed silently
    try:
        ADMIN_KEY = cfg.get("SETTINGS", "admin_key", fallback="")
        if not ADMIN_KEY:
            ADMIN_KEY = _generate_admin_key()
    except Exception:
        ADMIN_KEY = _generate_admin_key()
        logger.debug("Fallback admin_key generated (post-load recovery)")

# Safely read config with fallbacks
try:
    LMUTIL_PATH = cfg.get("SETTINGS", "lmutil_path", fallback=DEFAULT_LMUTIL)
except Exception:
    LMUTIL_PATH = DEFAULT_LMUTIL

try:
    LM_PORT = cfg.get("SETTINGS", "port", fallback=DEFAULT_PORT)
except Exception:
    LM_PORT = DEFAULT_PORT

try:
    REFRESH_MIN = cfg.getint("SETTINGS", "refresh_minutes", fallback=DEFAULT_REFRESH)
except Exception:
    REFRESH_MIN = DEFAULT_REFRESH

try:
    WEB_PORT = cfg.getint("SETTINGS", "web_port", fallback=DEFAULT_WEB_PORT)
except Exception:
    WEB_PORT = DEFAULT_WEB_PORT

# Read default locale preference (normalize like 'en-US' -> 'en')
try:
    DEFAULT_LOCALE_RAW = cfg.get("SETTINGS", "default_locale", fallback="en").strip().lower().split("-")[0]
except Exception:
    DEFAULT_LOCALE_RAW = "en"

try:
    HIDE_MAINT = cfg.get("SETTINGS", "hide_maintenance", fallback="no").lower() in ("1", "yes", "true")
except Exception:
    HIDE_MAINT = False

try:
    hide_list_str = cfg.get("SETTINGS", "hide_list", fallback="")
    HIDE_LIST = [x.strip() for x in hide_list_str.split(",") if x.strip()]
except Exception:
    HIDE_LIST = []

# New: enable_restart config (default True)
try:
    ENABLE_RESTART = cfg.getboolean("SETTINGS", "enable_restart", fallback=False)
    logger.info(f"Service restart enabled: {ENABLE_RESTART}")
except Exception:
    ENABLE_RESTART = True

# Show EID info config (default False)
try:
    SHOW_EID_INFO = cfg.getboolean("SETTINGS", "show_eid_info", fallback=False)
except Exception:
    SHOW_EID_INFO = False

# Safely read SERVICE section
try:
    if cfg.has_section("SERVICE"):
        SERVICE_NAME = cfg.get("SERVICE", "service_name", fallback="FLEXnet License Server")
    else:
        SERVICE_NAME = "FLEXnet License Server"
except Exception:
    SERVICE_NAME = "FLEXnet License Server"

# New Teams config read (defaults: disabled)
try:
    TEAMS_ENABLED = cfg.getboolean("TEAMS", "enabled", fallback=False)
except Exception:
    TEAMS_ENABLED = False

try:
    TEAMS_WEBHOOK = cfg.get("TEAMS", "webhook", fallback="", raw=True).strip()
    # Remove surrounding quotes if present
    if TEAMS_WEBHOOK.startswith('"') and TEAMS_WEBHOOK.endswith('"'):
        TEAMS_WEBHOOK = TEAMS_WEBHOOK[1:-1].strip()
    elif TEAMS_WEBHOOK.startswith("'") and TEAMS_WEBHOOK.endswith("'"):
        TEAMS_WEBHOOK = TEAMS_WEBHOOK[1:-1].strip()
    logger.debug(f"Teams webhook loaded (len={len(TEAMS_WEBHOOK)}): {TEAMS_WEBHOOK[:60]}..." if len(TEAMS_WEBHOOK) > 60 else f"Teams webhook loaded: '{TEAMS_WEBHOOK}'")
except Exception as e:
    TEAMS_WEBHOOK = ""
    logger.debug(f"Teams webhook could not be loaded; exception: {e}")

try:
    TEAMS_NOTIFY_UPDATE = cfg.getboolean("TEAMS", "notify_update", fallback=False)
except Exception:
    TEAMS_NOTIFY_UPDATE = True

try:
    TEAMS_NOTIFY_DUPLICATE_CHECKER = cfg.getboolean("TEAMS", "notify_duplicate_checker", fallback=False)
except Exception:
    TEAMS_NOTIFY_DUPLICATE_CHECKER = False

try:
    TEAMS_NOTIFY_EXTRATIME = cfg.getboolean("TEAMS", "notify_extratime", fallback=False)
except Exception:
    TEAMS_NOTIFY_EXTRATIME = False

try:
    TEAMS_EXTRATIME_DURATION = cfg.getint("TEAMS", "extratime_duration", fallback=72)
except Exception:
    TEAMS_EXTRATIME_DURATION = 72

try:
    extratime_excl_str = cfg.get("TEAMS", "extratime_exclusion", fallback="").strip()
    TEAMS_EXTRATIME_EXCLUSION = [x.strip() for x in extratime_excl_str.split(",") if x.strip()]
except Exception:
    TEAMS_EXTRATIME_EXCLUSION = []

try:
    TEAMS_NOTIFY_SOLDOUT = cfg.getboolean("TEAMS", "notify_soldout", fallback=False)
except Exception:
    TEAMS_NOTIFY_SOLDOUT = False

try:
    soldout_excl_str = cfg.get("TEAMS", "soldout_exclusion", fallback="").strip()
    TEAMS_SOLDOUT_EXCLUSION = [x.strip() for x in soldout_excl_str.split(",") if x.strip()]
except Exception:
    TEAMS_SOLDOUT_EXCLUSION = []

try:
    TEAMS_NOTIFY_DAEMON = cfg.getboolean("TEAMS", "notify_daemon", fallback=False)
except Exception:
    TEAMS_NOTIFY_DAEMON = False

# Teams notification state (avoid duplicate notifications for same release)
_LAST_TEAMS_NOTIFIED_VERSION = None

def send_teams_notification(title, message, link=None):
    locale = DEFAULT_LOCALE
    """Send a message to Microsoft Teams webhook with Adaptive Card payload."""
    logger.debug("send_teams_notification called: title=%s, link=%s", title, bool(link))
    if not TEAMS_ENABLED:
        logger.debug("Teams notifications disabled in config.")
        return False
    if not TEAMS_WEBHOOK:
        logger.warning("Teams webhook URL not configured; cannot send Teams notification.")
        return False

    try:
        import urllib.request
        import urllib.error

        # Build Adaptive Card payload for Power Automate
        payload = {
            "type": "message",
            "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": {
                "type": "AdaptiveCard",
                "body": [
                    {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": title
                    },
                    {
                    "type": "TextBlock",
                    "text": message,
                    "wrap": True
                    }
                ],
                "actions": [
                    {
                    "type": "Action.OpenUrl",
                    "title": TRANSLATIONS.get(locale, {}).get("view_details", "View Details"),
                    "url": link
                    }
                ] if link else [],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2"
                }
            }
            ]
        }

        data_bytes = json.dumps(payload).encode("utf-8")
        masked_webhook = (TEAMS_WEBHOOK[:80] + "...") if len(TEAMS_WEBHOOK) > 80 else TEAMS_WEBHOOK
        logger.debug("Prepared Teams Adaptive Card payload: %s", payload)
        logger.debug("Posting to Teams webhook (masked)=%s", masked_webhook)
        req = urllib.request.Request(
            TEAMS_WEBHOOK,
            data=data_bytes,
            headers={"Content-Type": "application/json", "User-Agent": "Licenses-WebUI-Teams-Notifier"}
        )
        logger.debug("Sending HTTP POST via urllib (timeout=15s)...")
        with urllib.request.urlopen(req, timeout=15) as resp:
            status = resp.getcode()
            try:
                body = resp.read().decode("utf-8", errors="replace")
            except Exception:
                body = "<could not read response body>"
            logger.info("Teams notification sent: status=%s", status)
            logger.debug("Teams response body: %s", body)
            return 200 <= status < 300
    except urllib.error.HTTPError as e:
        try:
            err_body = e.read().decode("utf-8", errors="replace")
        except Exception:
            err_body = "<no body>"
        logger.warning("urllib HTTPError: status=%s, reason=%s, body=%s", e.code, e.reason, err_body)
        return False
    except urllib.error.URLError as e:
        logger.warning("urllib URLError: %s", e.reason, exc_info=True)
        return False
    except Exception as e:
        logger.warning("Failed to send Teams notification: %s", e, exc_info=True)
        return False


# ---------- Internationalization ----------

SUPPORTED_LOCALES = ("en", "fr", "de", "es")
DEFAULT_LOCALE = DEFAULT_LOCALE_RAW if 'DEFAULT_LOCALE_RAW' in globals() and DEFAULT_LOCALE_RAW in SUPPORTED_LOCALES else "en"
TRANSLATIONS = {}

# Feature groups configuration
FEATURE_GROUPS = {}

def load_feature_groups():
    """Load feature groups from feature_groups.json.
    Supports both source-run and PyInstaller bundle via sys._MEIPASS.
    """
    global FEATURE_GROUPS
    base_dir = getattr(sys, '_MEIPASS', None) or os.path.dirname(__file__)
    groups_file = os.path.join(base_dir, "feature_groups.json")
    try:
        with open(groups_file, "r", encoding="utf-8") as f:
            FEATURE_GROUPS = json.load(f)
        logger.info(f"Loaded feature groups configuration from: {groups_file}")
    except FileNotFoundError:
        logger.warning(f"Feature groups file not found: {groups_file}")
        FEATURE_GROUPS = {"groups": [], "other": {"title": "Other", "icon": "static/icons/other.png"}}
    except json.JSONDecodeError as e:
        logger.error(f"Invalid JSON in feature groups file: {e}")
        FEATURE_GROUPS = {"groups": [], "other": {"title": "Other", "icon": "static/icons/other.png"}}

# ---------- Statistics Database ----------
# Database stored next to executable, not in package
def get_db_path():
    """Get database path next to executable (like config.ini)."""
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        base_dir = os.path.dirname(sys.executable)
    else:
        # Running as script
        base_dir = os.path.dirname(__file__)
    return os.path.join(base_dir, "license_stats.db")

DB_PATH = get_db_path()
_stats_last_state = {}  # Track last stored state per feature to detect changes

# EID information cache
_eid_cache = {}
_eid_cache_timestamp = 0
EID_CACHE_DURATION = 24 * 3600  # 24 hours in seconds
_clm_available = None  # None = unknown, True = working, False = failed

def refresh_eid_cache():
    """Refresh the EID cache by running CLM query-features command."""
    global _eid_cache, _eid_cache_timestamp, _clm_available
    try:
        clm_output = try_clm_query_features()
        if clm_output:
            _eid_cache = parse_eid_info(clm_output)
            _eid_cache_timestamp = time.time()
            _clm_available = True
            logger.info(f"EID cache refreshed: {len(_eid_cache)} features")
            return True
        else:
            # CLM command failed or returned nothing
            _clm_available = False
            logger.warning("CLM query-features failed or returned no data - EID features will be hidden")
    except Exception as e:
        _clm_available = False
        logger.error(f"Failed to refresh EID cache: {e} - EID features will be hidden")
    return False

def get_eid_info():
    """Get EID information from cache, refreshing if needed."""
    global _eid_cache, _eid_cache_timestamp
    
    # Check if cache needs refresh (older than 24h or empty)
    cache_age = time.time() - _eid_cache_timestamp
    if not _eid_cache or cache_age > EID_CACHE_DURATION:
        logger.debug(f"EID cache stale (age: {cache_age:.0f}s), refreshing...")
        refresh_eid_cache()
    
    return _eid_cache

def init_stats_db():
    """Initialize SQLite database for statistics tracking.
    Creates database next to executable if it doesn't exist.
    """
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS feature_usage (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp INTEGER NOT NULL,
                feature_name TEXT NOT NULL,
                used INTEGER NOT NULL,
                available INTEGER NOT NULL
            )
        """)
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_feature_timestamp 
            ON feature_usage(feature_name, timestamp)
        """)
        conn.commit()
        conn.close()
        logger.info(f"Stats database initialized at: {os.path.abspath(DB_PATH)}")
    except Exception as e:
        logger.error(f"Failed to initialize stats database: {e}")

def store_feature_stats(licenses_data):
    """Store feature usage stats only when values change.
    Respects HIDE_MAINT and HIDE_LIST exclusions.
    """
    global _stats_last_state
    try:
        timestamp = int(time.time())
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        for feature_name, data in licenses_data.items():
            lname = feature_name.lower()
            
            # Skip maintenance features if configured
            if HIDE_MAINT and 'maint' in lname:
                continue
            
            # Skip features in hide list
            if any(h.lower() in lname for h in HIDE_LIST):
                continue
            
            used = data.get('used')
            total = data.get('total')
            
            # Skip if missing data
            if used is None or total is None:
                continue
            
            # Get the last known state for this feature
            last_state = _stats_last_state.get(feature_name)
            current_state = (used, total)
            
            # Only store when state changes
            if last_state != current_state:
                # Insert the current state at current timestamp
                cursor.execute(
                    "INSERT INTO feature_usage (timestamp, feature_name, used, available) VALUES (?, ?, ?, ?)",
                    (timestamp, feature_name, used, total)
                )
                # Update the tracked state
                _stats_last_state[feature_name] = current_state
        
        conn.commit()
        conn.close()
    except Exception as e:
        logger.error(f"Failed to store feature stats: {e}")

# Application version and GitHub repo for update checks
VERSION = "1.5"
GITHUB_REPO = "nicobubulle/Licenses-WebUI"
GITHUB_RELEASES_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
LATEST_VERSION = VERSION
UPDATE_AVAILABLE = False
LATEST_URL = f"https://github.com/{GITHUB_REPO}"

# check interval in seconds (24 hours)
UPDATE_CHECK_INTERVAL = 24 * 3600

def check_github_latest_version():
    global LATEST_VERSION, UPDATE_AVAILABLE, LATEST_URL, _LAST_TEAMS_NOTIFIED_VERSION
    # Do NOT use Flask request context here (runs in background thread)
    locale = DEFAULT_LOCALE
    try:
        import urllib.request
        req = urllib.request.Request(
            GITHUB_RELEASES_API,
            headers={"User-Agent": "Licenses-WebUI-Version-Checker"}
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.load(resp)
            tag = data.get("tag_name") or data.get("name")
            if tag:
                latest = str(tag).lstrip("vV").strip()
                LATEST_VERSION = latest
                LATEST_URL = data.get("html_url") or f"https://github.com/{GITHUB_REPO}"
                UPDATE_AVAILABLE = (latest != VERSION)
                logger.info(f"GitHub check: latest={latest}, current={VERSION}, update={UPDATE_AVAILABLE}")

                # Debug: log Teams config/state before attempting notify
                try:
                    logger.debug("Teams config: enabled=%s, notify_update=%s, webhook_len=%s, last_notified=%s",
                                 TEAMS_ENABLED, TEAMS_NOTIFY_UPDATE, len(TEAMS_WEBHOOK or ""), _LAST_TEAMS_NOTIFIED_VERSION)
                except Exception:
                    logger.debug("Teams config debug failed")

                # Send Teams notification for new update (once per discovered version)
                if UPDATE_AVAILABLE and TEAMS_ENABLED and TEAMS_NOTIFY_UPDATE:
                    try:
                        if _LAST_TEAMS_NOTIFIED_VERSION != latest:
                            # craft a short message
                            title = TRANSLATIONS.get(locale, {}).get("update_title", "Update available")
                            message_tpl = TRANSLATIONS.get(locale, {}).get(
                                "update_message",
                                "A new version is available: v{latest} (current: v{current})"
                            )
                            try:
                                message = message_tpl.format(latest=latest, current=VERSION)
                            except Exception:
                                message = f"A new version is available: v{latest} (current: v{VERSION})"
                            logger.debug("Attempting to send Teams notification for version %s", latest)
                            sent = send_teams_notification(title, message, link=LATEST_URL)
                            logger.debug("send_teams_notification returned: %s", sent)
                            if sent:
                                logger.info("Teams update notification sent for version %s", latest)
                                _LAST_TEAMS_NOTIFIED_VERSION = latest
                            else:
                                logger.debug("Teams update notification NOT sent for version %s", latest)
                    except Exception as e:
                        logger.debug("Error while attempting to send Teams notification: %s", e, exc_info=True)
                return
    except Exception as e:
        logger.debug(f"GitHub version check failed: {e}", exc_info=True)
    logger.debug("GitHub version check finished with no usable release info.")

def update_check_loop(interval=UPDATE_CHECK_INTERVAL):
    """Run check_github_latest_version immediately and then once every `interval` seconds until shutdown."""
    logger.info("Starting background update-check loop (interval=%s seconds)", interval)
    while not _shutdown_event.is_set():
        try:
            check_github_latest_version()
        except Exception as e:
            logger.debug("Exception in update_check_loop: %s", e)
        # wait for interval but wake early if shutdown event is set
        if _shutdown_event.wait(interval):
            break
    logger.info("Update-check loop exiting due to shutdown.")

def load_translations():
    """Load translation JSON files from i18n folder (tolerate // comment lines)."""
    global TRANSLATIONS
    i18n_dir = os.path.join(os.path.dirname(__file__), "i18n")
    for locale in SUPPORTED_LOCALES:
        filepath = os.path.join(i18n_dir, f"{locale}.json")
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                raw = f.read()
            # Strip lines starting with // (workspace file header comments)
            cleaned = "\n".join(
                line for line in raw.splitlines()
                if not line.strip().startswith("//")
            )
            TRANSLATIONS[locale] = json.loads(cleaned)
            logger.info(f"Loaded translations for {locale}")
        except FileNotFoundError:
            logger.warning(f"Translation file not found: {filepath}")
            TRANSLATIONS[locale] = {}
        except json.JSONDecodeError as e:
            logger.error(f"Invalid JSON in {filepath}: {e}")
            TRANSLATIONS[locale] = {}

load_translations()
load_feature_groups()
init_stats_db()
# Defer initial EID cache refresh until after helper functions are defined later

def get_locale():
    """Get current locale from query param, cookie, or Accept-Language header."""
    # priority: ?lang query -> cookie -> Accept-Language header -> default
    q = request.args.get("lang")
    if q and q in SUPPORTED_LOCALES:
        return q
    c = request.cookies.get("lang")
    if c and c in SUPPORTED_LOCALES:
        return c
    accept = request.headers.get("Accept-Language", "")
    for part in accept.split(","):
        code = part.split(";")[0].strip().lower()
        if not code:
            continue
        code = code.split("-")[0]
        if code in SUPPORTED_LOCALES:
            return code
    return DEFAULT_LOCALE

def translate(key, **kwargs):
    """Translate a key with optional format arguments."""
    locale = get_locale()
    text = TRANSLATIONS.get(locale, {}).get(key) or TRANSLATIONS.get(DEFAULT_LOCALE, {}).get(key) or key
    try:
        return text.format(**kwargs) if kwargs else text
    except Exception:
        return text

@app.context_processor
def inject_i18n():
    """Inject i18n helpers into template context."""
    locale = get_locale()
    i18n = TRANSLATIONS.get(locale, TRANSLATIONS.get(DEFAULT_LOCALE, {}))
    i18n_json = json.dumps(i18n)
    feature_groups_json = json.dumps(FEATURE_GROUPS)
    return dict(
        _=lambda k, **kw: translate(k, **kw),
        i18n=i18n,
        i18n_json=i18n_json,
        lang=locale,
        app_version=VERSION,
        latest_version=LATEST_VERSION,
        update_available=UPDATE_AVAILABLE,
        latest_url=LATEST_URL,
        feature_groups_json=feature_groups_json
    )

# ---------- System Tray ----------

# ---------- Global shutdown flag ----------
_shutdown_event = threading.Event()

def create_systray():
    """Create system tray icon with menu."""
    try:
        from pystray import Icon, Menu, MenuItem
        from PIL import Image, ImageDraw
        import webbrowser

        # Try to load project's favicon.ico from static/ and use it as systray icon.
        icon_path = os.path.join(os.path.dirname(__file__), "static", "favicon.ico")
        img = None
        try:
            if os.path.exists(icon_path):
                img = Image.open(icon_path)
                # ensure RGBA and reasonable size for tray icon
                img = img.convert("RGBA")
                img = img.resize((64, 64), Image.Resampling.LANCZOS)
        except Exception as e:
            logger.warning(f"Failed to load static favicon for systray: {e}")
            img = None

        # Fallback: draw simple blue circle if favicon not available or failed to load
        if img is None:
            img = Image.new('RGBA', (64, 64), color=(255, 255, 255, 0))
            draw = ImageDraw.Draw(img)
            draw.ellipse([8, 8, 56, 56], fill='#007bff', outline='#0056d6')

        def on_open(icon, item):
            """Open web browser."""
            logger.info("User clicked systray: Open in Browser")
            webbrowser.open(f"http://localhost:{WEB_PORT}")

        def on_exit(icon, item):
            """Exit application."""
            logger.info("User clicked systray: Exit")
            # stop icon first (pystray requirement) then shutdown
            try:
                icon.stop()
            except Exception:
                pass
            shutdown_app()

        menu = Menu(
            MenuItem(f"Licenses WebUI - http://localhost:{WEB_PORT}", None, enabled=False),
            MenuItem("Open in Browser", on_open),
            MenuItem("Exit", on_exit)
        )

        tooltip = f"Licenses WebUI\nhttp://localhost:{WEB_PORT}"

        icon = Icon("Licenses WebUI", img, menu=menu, tooltip=tooltip)
        try:
            icon.title = "Licenses WebUI"
        except Exception:
            pass

        return icon
    except ImportError:
        logger.warning("pystray or Pillow not installed; systray disabled. Install with: pip install pystray pillow")
        return None


def run_systray():
    """Run systray icon in background thread."""
    global _systray
    try:
        _systray = create_systray()
        if _systray:
            logger.info("Starting system tray icon (hover for tooltip)...")
            _systray.run()
    except Exception as e:
        logger.error(f"Systray error: {e}", exc_info=True)


def terminate_other_instances(timeout=3):
    """
    Try to find and terminate other running instances of this application/process.
    Uses psutil when available; otherwise logs a warning.
    Matching heuristics:
      - process.exe == sys.executable
      - process.cmdline contains the current script path (when running as .py)
      - process.name matches basename(sys.executable)
    """
    try:
        import psutil
    except Exception:
        logger.warning("psutil not installed - cannot terminate sibling processes automatically. Install with: pip install psutil")
        return

    me_pid = os.getpid()
    my_exe = os.path.abspath(sys.executable)
    my_cmd0 = os.path.abspath(sys.argv[0]) if sys.argv and sys.argv[0] else None
    basename_exe = os.path.basename(my_exe).lower()

    matched = []
    for p in psutil.process_iter(['pid', 'exe', 'cmdline', 'name']):
        try:
            info = p.info
            pid = info.get('pid')
            if pid == me_pid:
                continue

            exe = (info.get('exe') or "") and os.path.abspath(info.get('exe') or "")
            cmdline = info.get('cmdline') or []
            name = (info.get('name') or "").lower()

            is_same = False
            if exe and exe == my_exe:
                is_same = True
            elif my_cmd0 and any(my_cmd0 in (c or "") for c in cmdline):
                is_same = True
            elif name and basename_exe and name == basename_exe:
                # last-resort match by process name
                is_same = True

            if is_same:
                try:
                    logger.info("Found sibling process to terminate: pid=%s name=%s exe=%s cmd=%s", pid, name, exe, cmdline)
                    matched.append(p)
                except Exception:
                    # ignore inspection errors
                    pass
        except Exception:
            # ignore processes we can't inspect
            continue

    if not matched:
        logger.debug("No other matching instances found.")
        return

    # ask processes to terminate
    for p in matched:
        try:
            p.terminate()
        except Exception as e:
            logger.warning("Failed to send terminate to pid %s: %s", p.pid, e)

    # wait up to timeout seconds, then kill remaining
    gone, alive = psutil.wait_procs(matched, timeout=timeout)
    for p in alive:
        try:
            logger.warning("Process pid %s did not exit after terminate, killing...", p.pid)
            p.kill()
        except Exception as e:
            logger.error("Failed to kill pid %s: %s", p.pid, e)

    logger.info("Sibling process termination complete.")

def shutdown_app():
    """Gracefully shutdown the entire application and ensure sibling processes are closed."""
    logger.info("Shutting down application...")

    # Signal threads to stop
    _shutdown_event.set()
    logger.info("Shutdown event set")

    # Stop systray if running
    global _systray
    if _systray:
        try:
            logger.info("Stopping systray...")
            _systray.stop()
        except Exception as e:
            logger.warning("Error stopping systray: %s", e)

    # Give background threads a short time to finish
    logger.info("Waiting briefly for background threads to stop...")
    time.sleep(0.8)

    # Attempt to stop Flask (werkzeug) if possible
    try:
        # safe: only available in werkzeug dev server env
        func = request.environ.get('werkzeug.server.shutdown')
        if func:
            logger.info("Requesting werkzeug server shutdown...")
            try:
                func()
            except Exception as e:
                logger.warning("werkzeug shutdown call failed: %s", e)
    except Exception:
        # request may not be available here; ignore
        pass

    # Attempt to find and terminate other running instances of this app
    try:
        terminate_other_instances(timeout=3)
    except Exception as e:
        logger.warning("terminate_other_instances failed: %s", e)

    logger.info("Exiting application process")
    # final exit
    try:
        sys.exit(0)
    except SystemExit:
        os._exit(0)


# ---------- Service Management ----------

def get_service_state(service_name):
    """
    Query Windows service and return (code:int|None, name:str, raw_output:str).

    Example return values:
      (4, "RUNNING", "<raw sc query output...>")
      (3, "STOP_PENDING", "<raw sc query output...>")
      (None, "UNKNOWN", "<raw sc query output or error>")
    """
    if not service_name or not isinstance(service_name, str):
        return None, "UNKNOWN", ""

    try:
        out = subprocess.check_output(
            ["sc", "query", service_name],
            stderr=subprocess.STDOUT,
            text=True,
            timeout=5,
            **_SUBP_NO_WINDOW_KW
        )
    except subprocess.CalledProcessError as e:
        raw = getattr(e, "output", str(e))
        logger.debug("Service query failed (CalledProcessError): %s", raw)
        return None, "UNKNOWN", raw
    except Exception as e:
        raw = str(e)
        logger.debug("Service query error: %s", raw)
        return None, "UNKNOWN", raw

    # Robustly find the STATE line (sc sometimes emits multiple lines; use multiline search)
    m = re.search(r"STATE\s*:\s*(\d+)\s+([A-Za-z_]+)", out, flags=re.IGNORECASE | re.MULTILINE)
    if not m:
        return None, "UNKNOWN", out
    try:
        code = int(m.group(1))
    except ValueError:
        code = None
    name = m.group(2).strip().upper()
    return code, name, out


def wait_for_state(service_name, target_code, timeout=30):
    """
    Wait for service to reach target numeric state code.

    Returns: True if reached, False on timeout. Logs transitions.
    """
    deadline = time.time() + timeout
    last_seen = None
    while time.time() < deadline:
        code, name, _raw = get_service_state(service_name)
        if code == target_code:
            return True
        if name and name != last_seen:
            logger.debug("Service %s state: %s (code=%s)", service_name, name, code)
            last_seen = name
        time.sleep(0.3)
    return False


def restart_service(service_name):
    """
    Stop and restart a Windows service with reliable state checking.

    Returns: (success: bool, log: str)
    """
    if not service_name or not isinstance(service_name, str):
        raise ValueError("Invalid service name")

    log = []

    def append(msg):
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        log.append(f"[{ts}] {msg}")

    append(f"Requested restart of service: {service_name}")
    logger.info(f"Requested restart of service: {service_name}")

    initial_code, initial_name, initial_raw = get_service_state(service_name)
    append(f"Initial state: {initial_name} (code={initial_code})")
    logger.info(f"Initial state: {initial_name} (code={initial_code})")

    # STOP phase: only issue stop if not already STOPPED
    if initial_code == STATE_STOPPED:
        append("Service already stopped.")
        logger.info("Service already stopped.")
    else:
        append("Sending stop command...")
        logger.info("Sending stop command...")
        try:
            out = subprocess.check_output(
                ["sc", "stop", service_name],
                stderr=subprocess.STDOUT,
                text=True,
                timeout=10,
                **_SUBP_NO_WINDOW_KW
            )
            append(f"sc stop output: {out.strip()}")
            logger.debug(f"sc stop output: {out.strip()}")
        except subprocess.CalledProcessError as e:
            msg = f"sc stop failed (code {e.returncode}): {getattr(e,'output',str(e)).strip()}"
            append(msg)
            logger.error(msg)
        except Exception as e:
            msg = f"sc stop error: {e}"
            append(msg)
            logger.error(msg)

        append(f"Waiting up to {SERVICE_STOP_TIMEOUT}s for service to reach STOPPED...")
        logger.info(f"Waiting up to {SERVICE_STOP_TIMEOUT}s for service to reach STOPPED...")
        stopped = wait_for_state(service_name, STATE_STOPPED, timeout=SERVICE_STOP_TIMEOUT)
        if stopped:
            append("Service reached STOPPED.")
            logger.info("Service reached STOPPED.")
            # capture and show the sc query output when the service is actually STOPPED
            code, name, raw = get_service_state(service_name)
            append(f"sc query output at STOPPED (state={name}, code={code}):")
            append(raw.strip() if raw else "<no output>")
            logger.debug(f"sc query output at STOPPED: {raw.strip() if raw else '<no output>'}")
        else:
            code, name, raw = get_service_state(service_name)
            msg = f"Timed out waiting for STOPPED. Current state: {name} (code={code})"
            append(msg)
            logger.error(msg)
            append("Full sc query output for diagnosis:")
            append(raw.strip() if raw else "<no output>")
            # Do not attempt to start if service never reached STOPPED
            append("Aborting restart: service did not stop cleanly.")
            logger.error("Aborting restart: service did not stop cleanly.")
            return False, "\n".join(log)

    time.sleep(0.8)

    # START phase (only attempted if we confirmed STOPPED above)
    append("Sending start command...")
    logger.info("Sending start command...")
    try:
        out = subprocess.check_output(
            ["sc", "start", service_name],
            stderr=subprocess.STDOUT,
            text=True,
            timeout=10,
            **_SUBP_NO_WINDOW_KW
        )
        append(f"sc start output: {out.strip()}")
        logger.debug(f"sc start output: {out.strip()}")
    except subprocess.CalledProcessError as e:
        msg = f"sc start failed (code {e.returncode}): {getattr(e,'output',str(e)).strip()}"
        append(msg)
        logger.error(msg)
    except Exception as e:
        msg = f"sc start error: {e}"
        append(msg)
        logger.error(msg)

    append(f"Waiting up to {SERVICE_START_TIMEOUT}s for service to reach RUNNING...")
    logger.info(f"Waiting up to {SERVICE_START_TIMEOUT}s for service to reach RUNNING...")
    if wait_for_state(service_name, STATE_RUNNING, timeout=SERVICE_START_TIMEOUT):
        append("Service reached RUNNING.")
        logger.info("Service reached RUNNING.")
        # capture and show the sc query output when the service is actually RUNNING
        code, name, raw = get_service_state(service_name)
        append(f"sc query output at RUNNING (state={name}, code={code}):")
        append(raw.strip() if raw else "<no output>")
        logger.debug(f"sc query output at RUNNING: {raw.strip() if raw else '<no output>'}")
    else:
        code, name, raw = get_service_state(service_name)
        msg = f"Timed out waiting for RUNNING. Current state: {name} (code={code})"
        append(msg)
        logger.error(msg)
        append("Full sc query output for diagnosis:")
        append(raw.strip() if raw else "<no output>")

    final_code, final_name, _ = get_service_state(service_name)
    append(f"Final state: {final_name} (code={final_code})")
    logger.info(f"Final state: {final_name} (code={final_code})")

    success = final_code == STATE_RUNNING
    return success, "\n".join(log)


# ---------- lmstat Execution ----------

def try_clm_query_features():
    """
    Execute CLMCommandLine.exe query-features to get EID information.
    
    Returns: Raw output text or None if command fails
    """
    clm_path = os.path.join(os.path.dirname(LMUTIL_PATH), "CLMCommandLine.exe")
    
    if not os.path.exists(clm_path):
        logger.debug(f"CLMCommandLine.exe not found at: {clm_path}")
        return None
    
    try:
        out = subprocess.check_output(
            [clm_path, "query-features"],
            stderr=subprocess.STDOUT,
            text=True,
            timeout=30,
            **_SUBP_NO_WINDOW_KW
        )
        logger.debug("CLM query-features command succeeded")
        return out
    except Exception as e:
        logger.debug(f"CLM query-features failed: {e}")
        return None


def parse_eid_info(raw_text):
    """
    Parse CLM query-features output to extract EID to feature mapping.
    
    Returns: Dict mapping feature names to set of EIDs
    """
    if not raw_text:
        return {}
    
    feature_to_eids = {}
    current_eid = None
    
    # Pattern: - 00112-15895-00040-08571-EEC92 (Floating):
    eid_pattern = re.compile(r'^\s*-\s*([0-9A-F]{5}-[0-9A-F]{5}-[0-9A-F]{5}-[0-9A-F]{5}-[0-9A-F]{5})\s+\(', re.IGNORECASE)
    # Pattern: - M3D_CLWRX.ArcGIS 0/1
    feature_pattern = re.compile(r'^\s*-\s+([A-Za-z0-9_\.]+)\s+\d+/\d+')
    
    for line in raw_text.splitlines():
        eid_match = eid_pattern.search(line)
        if eid_match:
            current_eid = eid_match.group(1).upper()
            continue
        
        if current_eid:
            feature_match = feature_pattern.search(line)
            if feature_match:
                feature_name = feature_match.group(1)
                if feature_name not in feature_to_eids:
                    feature_to_eids[feature_name] = set()
                feature_to_eids[feature_name].add(current_eid)
    
    logger.debug(f"Parsed EID info for {len(feature_to_eids)} features")
    return feature_to_eids


# Perform initial EID cache refresh now that required functions exist
try:
    refresh_eid_cache()
except Exception as _e:
    logger.debug(f"Initial EID cache refresh failed (non-fatal): {_e}")


def try_lmstat_commands():
    """
    Try multiple lmstat command syntaxes.
    
    Returns: Raw lmstat output text
    Raises: RuntimeError if all attempts fail
    """
    exe_dir = os.path.dirname(LMUTIL_PATH) or None

    commands = [
        [LMUTIL_PATH, "lmstat", "-a", "-c", f"{LM_PORT}@localhost"],
    ]

    last_err = None
    for cmd in commands:
        try:
            out = subprocess.check_output(
                cmd,
                cwd=exe_dir,
                stderr=subprocess.STDOUT,
                text=True,
                timeout=LMSTAT_TIMEOUT,
                **_SUBP_NO_WINDOW_KW
            )
            header = f"(Command: {' '.join(cmd)})\n\n"
            logger.debug(f"lmstat command succeeded: {' '.join(cmd)}")
            return header + out
        except subprocess.CalledProcessError as e:
            last_err = f"Failed: {' '.join(cmd)}\n{getattr(e, 'output', str(e))}"
            logger.debug(last_err)
        except subprocess.TimeoutExpired:
            last_err = f"Timeout: {' '.join(cmd)}"
            logger.debug(last_err)
        except Exception as e:
            last_err = f"Error: {' '.join(cmd)} - {e}"
            logger.debug(last_err)

    raise RuntimeError(last_err or "No lmstat command succeeded")


# ---------- Parsing ----------

def parse_lmstat(raw_text):
    """
    Parse lmstat output into structured data.
    
    Returns: Dict of features with usage and user information
    """
    features = {}
    lines = raw_text.splitlines()

    re_feature = re.compile(
        r"Users of\s+(.+?):\s*\(Total of\s+(\d+).*?Total of\s+(\d+)\s+licenses? in use",
        re.IGNORECASE
    )
    re_expiry = re.compile(r'expiry:\s*([0-9A-Za-z\-\s]+)', re.IGNORECASE)
    re_named_block = re.compile(r'^"([^"]+)"')
    re_user = re.compile(
        r'^\s*([^\s]+)\s+([^\s]+).*?\(v?([0-9A-Za-z\.\-]+)\).*?start\s+(.+)',
        re.IGNORECASE
    )

    current = None

    for line in lines:
        m = re_feature.search(line)
        if m:
            name = m.group(1).strip()
            total = int(m.group(2))
            used = int(m.group(3))
            features.setdefault(name, {"total": total, "used": used, "expiry": None, "users": []})
            current = name
            continue

        m2 = re_named_block.search(line)
        if m2:
            current = m2.group(1).strip()
            features.setdefault(current, {"total": None, "used": None, "expiry": None, "users": []})
            continue

        m3 = re_expiry.search(line)
        if m3 and current:
            features[current]["expiry"] = m3.group(1).strip()
            continue

        m4 = re_user.search(line)
        if m4 and current:
            features[current]["users"].append({
                "user": m4.group(1).strip(),
                "computer": m4.group(2).strip(),
                "version": m4.group(3).strip(),
                "start": m4.group(4).strip()
            })
            continue

    logger.debug(f"Parsed {len(features)} license features")
    return features


# ---------- Background Refresh ----------

_lock = threading.Lock()
_raw_output = ""
_parsed = {}
_last_update = None
_last_error = None
_last_service_msg = None
_systray = None

# Duplicate checker state: set of (feature, user, computer) tuples already notified
_notified_duplicates = set()

# Extratime checker state: users already notified (user, computer)
_notified_extratime = set()

# Sold-out checker state: last known sold-out status per feature
_soldout_last_state = {}
_daemon_last_state = None  # None unknown, True up, False down

def _is_port_open(host: str, port: str | int, timeout: float = 2.0) -> bool:
    """Best-effort TCP connect to check if license port is reachable."""
    try:
        import socket
        p = int(port)
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(timeout)
            s.connect((host, p))
            return True
    except Exception:
        return False

def _lmstat_output_indicates_ok(text: str) -> bool:
    """Heuristic check: return False if lmstat output clearly indicates connection failure.
    Looks for common FlexLM error strings such as "Cannot connect to license server system" or "Error getting status".
    """
    if not isinstance(text, str):
        return False
    t = text.lower()
    patterns = (
        "cannot connect to license server system",
        "error getting status",
        "no such host is known",
        "connection refused",
        "(-15",  # FlexLM cannot connect code
    )
    return not any(p in t for p in patterns)

def check_daemon_state(current_up_hint: bool):
    """Evaluate daemon up/down and notify on transitions.
    If lmstat succeeded (current_up_hint=True), consider daemon up.
    If lmstat failed, verify via service state and port connectivity before notifying down.
    """
    logger.debug(f"check_daemon_state called: current_up_hint={current_up_hint}, TEAMS_ENABLED={TEAMS_ENABLED}, TEAMS_NOTIFY_DAEMON={TEAMS_NOTIFY_DAEMON}")
    if not (TEAMS_ENABLED and TEAMS_NOTIFY_DAEMON):
        logger.debug("Daemon state check skipped (Teams not enabled or notify_daemon=no)")
        return
    global _daemon_last_state

    # Determine current state
    service_running = None
    service_name = "UNKNOWN"
    service_code = None
    try:
        code, name, _ = get_service_state(SERVICE_NAME)
        service_running = (code == STATE_RUNNING) if code is not None else None
        service_name = name or "UNKNOWN"
        service_code = code
    except Exception:
        # leave as None
        pass

    port_ok = None
    try:
        port_ok = _is_port_open("localhost", LM_PORT, timeout=2.0)
    except Exception:
        pass

    logger.debug(f"Daemon check inputs: service_running={service_running}, service_name={service_name}, service_code={service_code}, port_ok={port_ok}")

    if current_up_hint:
        # Even if lmstat looked ok, override to DOWN if service/port clearly indicate down
        if service_running is False and port_ok is False:
            current = False
        elif service_running is False and port_ok is None:
            current = False
        elif port_ok is False and (service_running is None):
            current = False
        else:
            current = True
    else:
        # Only declare down if both checks indicate down or unknown suggests down
        # Conservative: require either explicit not-running AND port closed, else assume up/indeterminate
        if service_running is False and port_ok is False:
            current = False
        elif service_running is False and port_ok is None:
            current = False
        elif port_ok is False and (service_running is None):
            current = False
        else:
            current = bool(service_running or port_ok)

    prev = _daemon_last_state
    _daemon_last_state = current
    
    logger.debug(f"Daemon state determined: prev={prev}, current={current}")

    locale = DEFAULT_LOCALE
    yes_str = TRANSLATIONS.get(locale, {}).get("yes", "yes")
    no_str = TRANSLATIONS.get(locale, {}).get("no", "no")
    port_ok_str = yes_str if port_ok else no_str

    if prev is None:
        # First observation: notify if currently down
        if current is False:
            title = TRANSLATIONS.get(locale, {}).get("daemon_down_title", "License Daemon Down")
            message_tpl = TRANSLATIONS.get(locale, {}).get(
                "daemon_down_message",
                "lmstat failed; daemon appears DOWN.\nService: {service} -> {state} (code={code})\nPort {port} reachable: {port_ok}"
            )
            try:
                message = message_tpl.format(service=SERVICE_NAME, state=service_name, code=service_code, port=LM_PORT, port_ok=port_ok_str)
            except Exception:
                message = (
                    f"lmstat failed; daemon appears DOWN.\n"
                    f"Service: {SERVICE_NAME} -> {service_name} (code={service_code})\n"
                    f"Port {LM_PORT} reachable: {port_ok_str}"
                )
            try:
                send_teams_notification(title, message)
                logger.info("Daemon down notification sent (initial)")
            except Exception as e:
                logger.warning(f"Failed to send daemon down notification: {e}", exc_info=True)
        return

    if prev != current:
        if current:
            title = TRANSLATIONS.get(locale, {}).get("daemon_up_title", "License Daemon Up Again")
            message_tpl = TRANSLATIONS.get(locale, {}).get(
                "daemon_up_message",
                "Daemon appears UP again.\nService: {service} -> {state} (code={code})\nPort {port} reachable: {port_ok}"
            )
        else:
            title = TRANSLATIONS.get(locale, {}).get("daemon_down_title", "License Daemon Down")
            message_tpl = TRANSLATIONS.get(locale, {}).get(
                "daemon_down_message",
                "lmstat failed; daemon appears DOWN.\nService: {service} -> {state} (code={code})\nPort {port} reachable: {port_ok}"
            )
        try:
            message = message_tpl.format(service=SERVICE_NAME, state=service_name, code=service_code, port=LM_PORT, port_ok=port_ok_str)
        except Exception:
            message = (
                ("Daemon appears UP again.\n" if current else "lmstat failed; daemon appears DOWN.\n") +
                f"Service: {SERVICE_NAME} -> {service_name} (code={service_code})\n" +
                f"Port {LM_PORT} reachable: {port_ok_str}"
            )
        try:
            send_teams_notification(title, message)
            logger.info(f"Daemon state change notified: {'UP' if current else 'DOWN'}")
        except Exception as e:
            logger.warning(f"Failed to send daemon state notification: {e}", exc_info=True)

def _parse_start_timestamp(start_str):
    """Best-effort parse of lmstat 'start' field into epoch seconds.
    FLEXlm formats vary; common patterns include:
      'Fri 11/17 12:34' (weekday MM/DD HH:MM) no year
      '11/17/2025 12:34' (MM/DD/YYYY HH:MM)
      'Nov 17 2025 12:34' (Mon DD YYYY HH:MM)
    We assume local time.
    Returns None if parsing fails.
    """
    if not start_str:
        return None
    s = start_str.strip()
    now = datetime.datetime.now()
    year = now.year
    # Try explicit YYYY first
    for fmt in ["%m/%d/%Y %H:%M", "%b %d %Y %H:%M", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y %H:%M:%S"]:
        try:
            dt = datetime.datetime.strptime(s, fmt)
            return dt.timestamp()
        except Exception:
            pass
    # Patterns without year -> append current year
    # weekday MM/DD HH:MM
    m = re.match(r"^[A-Za-z]{3} (\d{1,2})/(\d{1,2}) (\d{1,2}):(\d{2})", s)
    if m:
        try:
            month, day, hour, minute = map(int, m.groups())
            dt = datetime.datetime(year, month, day, hour, minute)
            # If date in future (e.g., year rollover), subtract one year
            if dt > now + datetime.timedelta(days=2):
                dt = dt.replace(year=year-1)
            return dt.timestamp()
        except Exception:
            pass
    # MM/DD HH:MM
    m2 = re.match(r"^(\d{1,2})/(\d{1,2}) (\d{1,2}):(\d{2})", s)
    if m2:
        try:
            month, day, hour, minute = map(int, m2.groups())
            dt = datetime.datetime(year, month, day, hour, minute)
            if dt > now + datetime.timedelta(days=2):
                dt = dt.replace(year=year-1)
            return dt.timestamp()
        except Exception:
            pass
    return None

def check_extratime(parsed_data):
    """Check for users exceeding extratime duration across features and notify once per user.
    Sends one message per user listing all features exceeding the threshold.
    Excludes features in TEAMS_EXTRATIME_EXCLUSION.
    """
    if not (TEAMS_ENABLED and TEAMS_NOTIFY_EXTRATIME):
        return
    threshold_hours = TEAMS_EXTRATIME_DURATION
    exclusions = set(f.lower() for f in TEAMS_EXTRATIME_EXCLUSION)
    now_ts = time.time()
    over_map = {}  # (user, computer) -> list of {feature, hours}
    for feature_name, data in parsed_data.items():
        lname = feature_name.lower()
        if lname in exclusions:
            continue
        if HIDE_MAINT and 'maint' in lname:
            continue
        for u in data.get("users", []):
            start_ts = _parse_start_timestamp(u.get("start"))
            if not start_ts:
                continue
            hours = (now_ts - start_ts) / 3600.0
            if hours >= threshold_hours:
                key = (u.get("user", "?"), u.get("computer", "?"))
                over_map.setdefault(key, []).append({"feature": feature_name, "hours": hours})
    if not over_map:
        return
    for (user, computer), feats in over_map.items():
        if (user, computer) in _notified_extratime:
            continue
        # Build message
        lines = [f"{f['feature']}: {f['hours']:.1f}h" for f in sorted(feats, key=lambda x: -x['hours'])]
        locale = DEFAULT_LOCALE
        title = TRANSLATIONS.get(locale, {}).get("extratime_title", "Extended Usage Detected")
        message_tpl = TRANSLATIONS.get(locale, {}).get(
            "extratime_message",
            "User **{user}@{computer}** has features exceeding {threshold}h:\n{features}"
        )
        try:
            message = message_tpl.format(
            user=user,
            computer=computer,
            threshold=threshold_hours,
            features="\n".join(lines)
            )
        except Exception:
            message = (
            f"User **{user}@{computer}** has features exceeding {threshold_hours}h:\n" +
            "\n".join(lines)
            )
        try:
            send_teams_notification(title, message)
            _notified_extratime.add((user, computer))
            logger.info(f"Extratime notification sent for {user}@{computer} ({len(feats)} features)")
        except Exception as e:
            logger.warning(f"Failed to send extratime notification for {user}@{computer}: {e}", exc_info=True)

def check_soldout(parsed_data):
    """Notify when a feature becomes sold out (used == total) or becomes available again.
    Excludes features in TEAMS_SOLDOUT_EXCLUSION. Sends notifications on state transitions.
    """
    if not (TEAMS_ENABLED and TEAMS_NOTIFY_SOLDOUT):
        return
    exclusions = set(f.lower() for f in TEAMS_SOLDOUT_EXCLUSION)
    global _soldout_last_state
    changes = []  # (feature, new_state, used, total)
    for feature_name, data in parsed_data.items():
        lname = feature_name.lower()
        if lname in exclusions:
            continue
        if HIDE_MAINT and 'maint' in lname:
            continue
        total = data.get("total")
        used = data.get("used")
        if total is None or used is None:
            continue
        sold_out = (used >= total)
        prev = _soldout_last_state.get(feature_name)
        if prev is None:
            # First observation: only notify if currently sold out
            if sold_out:
                changes.append((feature_name, True, used, total))
        else:
            if prev != sold_out:
                changes.append((feature_name, sold_out, used, total))
        _soldout_last_state[feature_name] = sold_out
    if not changes:
        return
    for feature, is_soldout, used, total in changes:
        locale = DEFAULT_LOCALE
        if is_soldout:
            title = TRANSLATIONS.get(locale, {}).get("soldout_title", "Feature Sold Out")
            message_tpl = TRANSLATIONS.get(locale, {}).get(
                "soldout_message",
                "Feature **{feature}** is fully used ({used}/{total})."
            )
            try:
                message = message_tpl.format(feature=feature, used=used, total=total)
            except Exception:
                message = f"Feature **{feature}** is fully used ({used}/{total})."
        else:
            title = TRANSLATIONS.get(locale, {}).get("soldout_available_title", "Feature Available Again")
            message_tpl = TRANSLATIONS.get(locale, {}).get(
                "soldout_available_message",
                "Feature **{feature}** no longer sold out ({used}/{total})."
            )
            try:
                message = message_tpl.format(feature=feature, used=used, total=total)
            except Exception:
                message = f"Feature **{feature}** no longer sold out ({used}/{total})."
        try:
            send_teams_notification(title, message)
            logger.info(f"Soldout state change notified: {feature} -> {'sold out' if is_soldout else 'available'}")
        except Exception as e:
            logger.warning(f"Failed to send soldout notification for {feature}: {e}", exc_info=True)

def check_duplicates(parsed_data):
    """Check for duplicate checkouts (same user@computer with same feature multiple times) and send Teams notifications."""
    global _notified_duplicates
    
    if not TEAMS_ENABLED or not TEAMS_NOTIFY_DUPLICATE_CHECKER:
        return
    
    new_duplicates = []
    
    for feature_name, feature_data in parsed_data.items():
        lname = feature_name.lower()
        if HIDE_MAINT and 'maint' in lname:
            continue
        users = feature_data.get("users", [])
        if len(users) < 2:
            continue
        
        # Count occurrences of each user@computer pair
        user_computer_counts = {}
        for u in users:
            key = (u["user"], u["computer"])
            user_computer_counts[key] = user_computer_counts.get(key, 0) + 1
        
        # Find duplicates (count > 1)
        for (user, computer), count in user_computer_counts.items():
            if count > 1:
                dup_key = (feature_name, user, computer)
                if dup_key not in _notified_duplicates:
                    new_duplicates.append({
                        "feature": feature_name,
                        "user": user,
                        "computer": computer,
                        "count": count
                    })
                    _notified_duplicates.add(dup_key)
                    logger.info(f"Duplicate checkout detected: {user}@{computer} checked out {feature_name} {count} times")
    
    # Send Teams notification for new duplicates
    if new_duplicates:
        for dup in new_duplicates:
            locale = DEFAULT_LOCALE
            title = TRANSLATIONS.get(locale, {}).get("duplicate_title", "Duplicate License Checkout Detected")
            message_tpl = TRANSLATIONS.get(locale, {}).get(
                "duplicate_message",
                "**Feature:** {feature}\n**User:** {user}\n**Computer:** {computer}\n**Checkout count:** {count}"
            )
            try:
                message = message_tpl.format(feature=dup['feature'], user=dup['user'], computer=dup['computer'], count=dup['count'])
            except Exception:
                message = f"**Feature:** {dup['feature']}\n**User:** {dup['user']}\n**Computer:** {dup['computer']}\n**Checkout count:** {dup['count']}"
            try:
                send_teams_notification(title, message)
                logger.info(f"Teams notification sent for duplicate: {dup['user']}@{dup['computer']} - {dup['feature']}")
            except Exception as e:
                logger.warning(f"Failed to send duplicate notification: {e}", exc_info=True)

def refresh_loop():
    """Background thread that periodically refreshes license data."""
    global _raw_output, _parsed, _last_update, _last_error
    
    while not _shutdown_event.is_set():
        try:
            out = try_lmstat_commands()
            parsed = parse_lmstat(out)
            with _lock:
                _raw_output = out
                _parsed = parsed
                _last_update = time.time()
                _last_error = None
                logger.info("License data refreshed successfully")
            
            # Store statistics for changed features
            try:
                store_feature_stats(parsed)
            except Exception as e:
                logger.warning(f"Stats storage failed: {e}", exc_info=True)
            
            # Determine daemon hint based on lmstat output content
            lmstat_ok = _lmstat_output_indicates_ok(out)
            try:
                check_daemon_state(bool(lmstat_ok))
            except Exception:
                pass
            
            # Run duplicate checker after updating parsed data
            try:
                check_duplicates(parsed)
            except Exception as e:
                logger.warning(f"Duplicate checker failed: {e}", exc_info=True)
            # Run extratime checker
            try:
                check_extratime(parsed)
            except Exception as e:
                logger.warning(f"Extratime checker failed: {e}", exc_info=True)
            # Run soldout checker
            try:
                check_soldout(parsed)
            except Exception as e:
                logger.warning(f"Soldout checker failed: {e}", exc_info=True)
                
        except Exception as e:
            with _lock:
                _last_error = str(e)
            logger.error(f"Refresh failed: {e}", exc_info=True)
            # Daemon might be down; verify and notify transitions
            try:
                check_daemon_state(False)
            except Exception:
                pass
        
        # Sleep in small intervals to respond quickly to shutdown
        for _ in range(REFRESH_MIN * 60):
            if _shutdown_event.is_set():
                break
            time.sleep(1)


# Start background thread
threading.Thread(target=refresh_loop, daemon=True).start()


# ---------- Routes ----------

@app.route("/")
def index():
    global _last_service_msg
    show = request.args.get("service_msg")
    service_msg = _last_service_msg if show and _last_service_msg else None
    _last_service_msg = None

    try:
        has_admin = is_admin()
    except Exception:
        has_admin = False
    show_restart = bool(ENABLE_RESTART and has_admin)

    # Application-level admin mode via URL param or cookie
    param_admin = request.args.get("admin")
    cookie_admin = request.cookies.get("admin")
    is_admin_mode = False
    if param_admin and ADMIN_KEY and param_admin == ADMIN_KEY:
        is_admin_mode = True
    elif cookie_admin and ADMIN_KEY and cookie_admin == ADMIN_KEY:
        is_admin_mode = True

    # Only show EID features if CLM is available
    show_eid = (SHOW_EID_INFO or is_admin_mode) and _clm_available is not False
    
    resp = make_response(render_template(
        "index.html",
        refresh_minutes=REFRESH_MIN,
        service_msg=service_msg,
        show_restart=show_restart,
        admin_mode=is_admin_mode,
        show_eid_info=show_eid
    ))

    # Persist ?lang=xx into cookie
    qlang = request.args.get("lang")
    if qlang in SUPPORTED_LOCALES:
        resp.set_cookie("lang", qlang, max_age=30*24*3600)  # 30 days
    # Persist admin mode if provided via query param
    if param_admin and ADMIN_KEY and param_admin == ADMIN_KEY:
        resp.set_cookie("admin", ADMIN_KEY, max_age=7*24*3600)  # 7 days
    return resp


@app.route("/status")
def status():
    with _lock:
        if _last_error:
            return jsonify({"ok": False, "error": _last_error, "last_update": _last_update})
        
        filtered = {}
        for name, item in _parsed.items():
            lname = name.lower()

            if HIDE_MAINT and "maint" in lname:
                continue

            if any(h.lower() in lname for h in HIDE_LIST):
                continue

            filtered[name] = item
        
        # Get EID information from cache (only if CLM is available)
        eid_data = get_eid_info() if _clm_available else {}

        return jsonify({
            "ok": True,
            "last_update": _last_update,
            "licenses": filtered,
            "eid_info": {fname: list(eids) for fname, eids in eid_data.items()} if _clm_available else {}
        })


@app.route("/refresh", methods=["POST"])
def manual_refresh():
    """Force synchronous refresh."""
    global _raw_output, _parsed, _last_update, _last_error
    
    logger.info("Manual refresh requested")
    try:
        out = try_lmstat_commands()
        parsed = parse_lmstat(out)
        with _lock:
            _raw_output = out
            _parsed = parsed
            _last_update = time.time()
            _last_error = None
        logger.info("Manual refresh completed successfully")

        # Note: Do NOT store stats on manual refresh, only on automatic loop

        # Determine daemon hint based on lmstat output content
        lmstat_ok = _lmstat_output_indicates_ok(out)
        try:
            check_daemon_state(bool(lmstat_ok))
        except Exception:
            pass

        # Run same checkers as automatic loop
        try:
            check_duplicates(parsed)
        except Exception as e:
            logger.warning(f"Duplicate checker failed (manual): {e}", exc_info=True)
        try:
            check_extratime(parsed)
        except Exception as e:
            logger.warning(f"Extratime checker failed (manual): {e}", exc_info=True)
        try:
            check_soldout(parsed)
        except Exception as e:
            logger.warning(f"Soldout checker failed (manual): {e}", exc_info=True)

        return redirect(url_for("index"))
    except Exception as e:
        with _lock:
            _last_error = str(e)
        logger.error(f"Manual refresh failed: {e}", exc_info=True)
        # Daemon might be down; verify and notify transitions
        logger.debug("Manual refresh exception handler: calling check_daemon_state(False)")
        try:
            check_daemon_state(False)
        except Exception as ex:
            logger.warning(f"check_daemon_state raised exception: {ex}", exc_info=True)
        return redirect(url_for("index"))


@app.route("/raw")
def raw():
    """Debug route to view raw lmstat output."""
    with _lock:
        return f"<pre>{_raw_output.replace('<', '&lt;')}</pre>"


@app.route("/stats")
def stats():
    """Statistics page with usage graphs."""
    try:
        has_admin = is_admin()
    except Exception:
        has_admin = False
    # App-level admin mode flag (used later for admin-only features)
    admin_mode = False
    try:
        param_admin = request.args.get("admin")
        cookie_admin = request.cookies.get("admin")
        if (param_admin and ADMIN_KEY and param_admin == ADMIN_KEY) or (cookie_admin and ADMIN_KEY and cookie_admin == ADMIN_KEY):
            admin_mode = True
    except Exception:
        admin_mode = False

    return render_template(
        "stats.html",
        show_restart=ENABLE_RESTART and has_admin,
        refresh_minutes=REFRESH_MIN,
        app_version=VERSION,
        update_available=UPDATE_AVAILABLE,
        latest_version=LATEST_VERSION,
        latest_url=LATEST_URL,
        admin_mode=admin_mode
    )

@app.route("/eids")
def eids_page():
    """Application admin EID overview page."""
    # Determine app-level admin mode (cookie or query)
    admin_mode = False
    try:
        param_admin = request.args.get("admin")
        cookie_admin = request.cookies.get("admin")
        if (param_admin and ADMIN_KEY and param_admin == ADMIN_KEY) or (cookie_admin and ADMIN_KEY and cookie_admin == ADMIN_KEY):
            admin_mode = True
    except Exception:
        admin_mode = False

    if not admin_mode:
        # Non-admins are redirected to dashboard
        return redirect(url_for("index"))

    # If CLM is unavailable, redirect to dashboard
    if _clm_available is False:
        logger.warning("EID page accessed but CLM is unavailable - redirecting to dashboard")
        return redirect(url_for("index"))

    # Build EID -> features detailed mapping using cached EID info and current license data
    with _lock:
        eid_feature_map_raw = _eid_cache.copy()
        licenses_snapshot = _parsed.copy()

    # Pre-compute group resolution helpers
    exact_map = {}
    wildcard_patterns = []
    try:
        for g in FEATURE_GROUPS.get("groups", []):
            for f in g.get("features", []):
                fl = str(f).lower()
                if "*" in fl:
                    wildcard_patterns.append((fl, g.get("id")))
                else:
                    exact_map[fl] = g.get("id")
    except Exception:
        pass

    def resolve_group(feature_name: str) -> str:
        lname = str(feature_name).lower()
        if lname in exact_map:
            return exact_map[lname] or "other"
        for pat, gid in wildcard_patterns:
            # Convert wildcard * to regex any chars
            import re as _re
            pattern_re = '^' + _re.escape(pat).replace('\\*', '.*') + '$'
            if _re.match(pattern_re, lname):
                return gid or "other"
        return "other"

    eid_detailed = {}
    for feature, eids in eid_feature_map_raw.items():
        for eid in eids:
            eid_list = eid_detailed.setdefault(eid, [])
            lic = licenses_snapshot.get(feature, {})
            total = lic.get('total')
            used = lic.get('used')
            group_id = resolve_group(feature)
            eid_list.append({
                'feature': feature,
                'total': total if total is not None else None,
                'used': used if used is not None else None,
                'group': group_id,
            })

    # Sort features inside each EID for stable display
    for eid, flist in eid_detailed.items():
        flist.sort(key=lambda x: (x.get('group') != 'other', x.get('feature').lower()))

    eid_json = json.dumps(eid_detailed)

    return render_template(
        "eids.html",
        admin_mode=admin_mode,
        show_restart=False,
        refresh_minutes=REFRESH_MIN,
        app_version=VERSION,
        update_available=UPDATE_AVAILABLE,
        latest_version=LATEST_VERSION,
        latest_url=LATEST_URL,
        eids_json=eid_json
    )


@app.route("/api/stats")
def api_stats():
    """API endpoint to get time-series data for charts.
    Query params:
    - feature: feature name (optional, returns all if not specified)
    - hours: time window in hours (default 24) OR
    - start: start timestamp (unix seconds) with end for absolute range
    - end: end timestamp (unix seconds) with start for absolute range
    """
    feature_filter = request.args.get('feature')
    
    # Support both relative (hours) and absolute (start/end) time ranges
    start_param = request.args.get('start')
    end_param = request.args.get('end')
    
    if start_param and end_param:
        # Absolute time range
        cutoff_time = int(start_param)
        end_time = int(end_param)
    else:
        # Relative time range (hours from now)
        hours = int(request.args.get('hours', 24))
        end_time = int(time.time())
        cutoff_time = end_time - (hours * 3600)
    
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # Calculate backfill time for step transitions
        try:
            backfill_secs = int(REFRESH_MIN) * 60 if isinstance(REFRESH_MIN, (int, float)) else 300
        except Exception:
            backfill_secs = 300
        
        if feature_filter:
            cursor.execute("""
                SELECT timestamp, feature_name, used, available
                FROM feature_usage
                WHERE feature_name = ? AND timestamp <= ?
                ORDER BY timestamp ASC
            """, (feature_filter, end_time))
        else:
            cursor.execute("""
                SELECT timestamp, feature_name, used, available
                FROM feature_usage
                WHERE timestamp <= ?
                ORDER BY feature_name, timestamp ASC
            """, (end_time,))
        
        rows = cursor.fetchall()
        conn.close()
        
        # Process data to add backfill points for step transitions
        data_by_feature = {}
        for ts, fname, used, available in rows:
            if fname not in data_by_feature:
                data_by_feature[fname] = []
            data_by_feature[fname].append({
                'timestamp': ts,
                'used': used,
                'available': available
            })
        
        # Add backfill points: for each feature, insert previous value at REFRESH_MIN before each change
        enhanced_data = {}
        for fname, points in data_by_feature.items():
            enhanced_points = []
            for i, point in enumerate(points):
                # Add backfill point before this point (except for the first point)
                if i > 0:
                    prev_point = points[i - 1]
                    # Add previous value at REFRESH_MIN before current timestamp
                    backfill_ts = point['timestamp'] - backfill_secs
                    # Only add backfill if it's after the previous real point
                    if backfill_ts > prev_point['timestamp']:
                        enhanced_points.append({
                            'timestamp': backfill_ts,
                            'used': prev_point['used'],
                            'available': prev_point['available']
                        })
                
                # Add the actual point
                enhanced_points.append(point)
            
            # Filter to requested time range
            filtered_points = [p for p in enhanced_points if cutoff_time <= p['timestamp'] <= end_time]
            if filtered_points:
                enhanced_data[fname] = filtered_points
        
        return jsonify({
            'ok': True,
            'features': enhanced_data
        })
    
    except Exception as e:
        logger.error(f"Stats API error: {e}", exc_info=True)
        return jsonify({'ok': False, 'error': str(e)})


@app.route("/refresh-eid", methods=["POST"])
def refresh_eid_route():
    """Manually refresh EID cache."""
    logger.info("Manual EID refresh requested")
    try:
        success = refresh_eid_cache()
        if success:
            return jsonify({"ok": True, "message": "EID cache refreshed successfully"})
        else:
            error_msg = "CLM query-features command failed or is unavailable" if _clm_available is False else "Failed to refresh EID cache"
            return jsonify({"ok": False, "error": error_msg}), 500
    except Exception as e:
        logger.error(f"EID refresh error: {e}", exc_info=True)
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/restart", methods=["POST"])
def restart_route():
    """Restart the FLEXnet License Server."""
    global _last_service_msg

    # Block if disabled by config
    if not ENABLE_RESTART:
        logger.warning("Restart attempted but disabled by configuration")
        return {"status": "disabled", "message": "Restart disabled by configuration"}, 403

    # Block if not running as admin
    if not is_admin():
        logger.warning("Restart attempted without admin privileges")
        return {"status": "error", "message": "Administrator privileges required"}, 403

    try:
        success, log = restart_service(SERVICE_NAME)
        _last_service_msg = log
        logger.info("Service restart completed; success=%s", success)

        if success:
            return {"status": "ok", "log": log}
        else:
            return {"status": "error", "message": "Service did not reach RUNNING state", "log": log}, 500
    except Exception as e:
        error_msg = str(e)
        _last_service_msg = error_msg
        logger.error(f"Service restart failed: {e}", exc_info=True)
        return {"status": "error", "message": error_msg}, 500


# ---------- Admin Check ----------

def is_admin():
    """Check if the application is running with administrator privileges."""
    try:
        # use shell32 (not 'shell') which exists on Windows and works when frozen
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception as e:
        logger.warning(f"Admin check failed: {e}")
        return False


def elevate_to_admin():
    """Try to re-run the current executable with admin privileges (UAC)."""
    if is_admin():
        return True

    logger.warning("Not running as admin. Requesting elevation...")
    try:
        # Use shell32.ShellExecuteW; when frozen sys.executable is the exe path
        params = " ".join(f'"{a}"' for a in sys.argv[1:]) if len(sys.argv) > 1 else ""
        ret = ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
        # ShellExecuteW returns >32 on success
        if int(ret) <= 32:
            logger.error(f"ShellExecuteW failed (code={ret})")
            return False
        # If elevation was requested, exit this process; elevated instance will start separately
        logger.info("Elevation requested, exiting current process.")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Failed to elevate to admin: {e}", exc_info=True)
        return False


# ---------- Startup ----------

if __name__ == "__main__":
    logger.info("Starting Licenses WebUI application...")
    
    # Check and elevate to admin FIRST to avoid duplicate background threads
    # Only elevate to admin if restart functionality is enabled
    if ENABLE_RESTART:
        if not elevate_to_admin():
            logger.error("This application requires administrator privileges to restart services.")
            sys.exit(1)
    else:
        logger.info("Service restart disabled in config; skipping admin elevation")
    
    logger.info("Running with administrator privileges")

    # Start background update-check loop (immediate check + daily repeats)
    try:
        t_ver = threading.Thread(target=update_check_loop, daemon=True)
        t_ver.start()
    except Exception as e:
        logger.debug(f"Could not start update-check thread: {e}")
    
    # Perform initial sync fetch
    try:
        out = try_lmstat_commands()
        parsed = parse_lmstat(out)
        with _lock:
            _raw_output = out
            _parsed = parsed
            _last_update = time.time()
            _last_error = None
        logger.info("Initial license data loaded successfully")
    except Exception as e:
        with _lock:
            _last_error = str(e)
        logger.error(f"Initial load failed: {e}", exc_info=True)

    # Start systray in background thread
    systray_thread = threading.Thread(target=run_systray, daemon=True)
    systray_thread.start()
    
    # Auto-open browser after a short delay to allow server to start
    import webbrowser
    def open_browser():
        time.sleep(1.5)  # Wait for Flask to start
        logger.info(f"Opening browser at http://localhost:{WEB_PORT}")
        webbrowser.open(f"http://localhost:{WEB_PORT}")
    
    browser_thread = threading.Thread(target=open_browser, daemon=True)
    browser_thread.start()
    
    logger.info(f"Licenses WebUI running on http://localhost:{WEB_PORT}")
    logger.info("=" * 80)
    app.run(host="0.0.0.0", port=WEB_PORT)