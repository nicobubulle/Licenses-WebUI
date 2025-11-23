import os
import re
import time
import threading
import subprocess
import configparser
import logging
import json
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
cfg = configparser.ConfigParser()
config_path = "config.ini"

if os.path.exists(config_path):
    try:
        cfg.read(config_path)
        logger.info(f"Loaded config from: {os.path.abspath(config_path)}")
    except Exception as e:
        logger.warning(f"Error reading config.ini: {e}. Using defaults.")
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
enable_restart = yes

[SERVICE]
# Windows service name to restart
service_name = FLEXnet License Server

[TEAMS]
# Enable sending notifications to Microsoft Teams (yes/no)
enabled = no

# Incoming webhook URL for Microsoft Teams (set this to your connector URL)
webhook =

# Enable the "update available" notification (yes/no)
notify_update = yes
"""
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            f.write(sample)
        # Re-read the newly created config so values are available
        cfg.read(config_path)
        logger.info(f"No config.ini found: created sample at {os.path.abspath(config_path)}")
    except Exception as e:
        logger.warning(f"Failed to create sample config.ini: {e}. Continuing with defaults.")

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
    ENABLE_RESTART = cfg.getboolean("SETTINGS", "enable_restart", fallback=True)
    logger.info(f"Service restart enabled: {ENABLE_RESTART}")
except Exception:
    ENABLE_RESTART = True

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
    TEAMS_NOTIFY_UPDATE = cfg.getboolean("TEAMS", "notify_update", fallback=True)
except Exception:
    TEAMS_NOTIFY_UPDATE = True

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

# Application version and GitHub repo for update checks
VERSION = "1.0"
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
    return dict(
        _=lambda k, **kw: translate(k, **kw),
        i18n=i18n,
        i18n_json=i18n_json,
        lang=locale,
        app_version=VERSION,
        latest_version=LATEST_VERSION,
        update_available=UPDATE_AVAILABLE,
        latest_url=LATEST_URL
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

def try_lmstat_commands():
    """
    Try multiple lmstat command syntaxes.
    
    Returns: Raw lmstat output text
    Raises: RuntimeError if all attempts fail
    """
    exe_dir = os.path.dirname(LMUTIL_PATH) or None

    commands = [
        [LMUTIL_PATH, "lmstat", "-a", "-c", f"{LM_PORT}@localhost"],
        [LMUTIL_PATH, "lmstat", "-a", "-c", f"localhost:{LM_PORT}"],
        [LMUTIL_PATH, "lmstat", "-a"],
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
        except Exception as e:
            with _lock:
                _last_error = str(e)
            logger.error(f"Refresh failed: {e}", exc_info=True)
        
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

    resp = make_response(render_template(
        "index.html",
        refresh_minutes=REFRESH_MIN,
        service_msg=service_msg,
        show_restart=show_restart
    ))

    # Persist ?lang=xx into cookie
    qlang = request.args.get("lang")
    if qlang in SUPPORTED_LOCALES:
        resp.set_cookie("lang", qlang, max_age=30*24*3600)  # 30 days
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

        return jsonify({"ok": True, "last_update": _last_update, "licenses": filtered})


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
        return redirect(url_for("index"))
    except Exception as e:
        with _lock:
            _last_error = str(e)
        logger.error(f"Manual refresh failed: {e}", exc_info=True)
        return redirect(url_for("index"))


@app.route("/raw")
def raw():
    """Debug route to view raw lmstat output."""
    with _lock:
        return f"<pre>{_raw_output.replace('<', '&lt;')}</pre>"


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