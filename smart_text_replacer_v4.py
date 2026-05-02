"""
Smart Text Replacer v4.0
━━━━━━━━━━━━━━━━━━━━━━━
Privacy-first, open-source Windows text expansion utility.
Replicates macOS Text Replacement — with more power.

Features:
  • Word-boundary aware expansion (no partial-word false triggers)
  • Dynamic variables: %DATE%, %TIME%, %DATETIME%, %CLIP%
  • Password field detection (ES_PASSWORD, credential dialogs)
  • App exclusion blocklist
  • Global hotkey: Ctrl+Alt+P to pause/resume
  • Snippet groups, search, enable/disable per snippet
  • Import/Export CSV
  • Auto-backup (keeps last 10)
  • System tray icon (requires pystray + Pillow)
  • Encrypted snippet storage (Fernet/AES)
  • Starts minimised to tray
  • Auto-update check via GitHub releases API
  • PyInstaller-compatible (single .exe)

Author: Steve Derviniotis
License: MIT
GitHub: https://github.com/sderviniotis/smart-text-replacer
"""

import base64
import ctypes
import ctypes.wintypes
import csv
import json
import logging
import os
import shutil
import sys
import threading
import time
import tkinter as tk
import winreg
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pynput.keyboard as pynput_kb

# ──────────────────────────────────────────────────────────────────
# OPTIONAL IMPORTS
# ──────────────────────────────────────────────────────────────────
try:
    from cryptography.fernet import Fernet
    CRYPTO_AVAILABLE = True
except ImportError:
    CRYPTO_AVAILABLE = False

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

# ──────────────────────────────────────────────────────────────────
# CONSTANTS
# ──────────────────────────────────────────────────────────────────
APP_NAME    = "SmartTextReplacer"
APP_VERSION = "4.0.0"
GITHUB_REPO = "sderviniotis/smart-text-replacer"   # update before publishing

DATA_DIR   = Path(os.environ.get("APPDATA", ".")) / APP_NAME
DATA_DIR.mkdir(parents=True, exist_ok=True)

SNIPPETS_FILE  = DATA_DIR / "snippets.json"
KEY_FILE       = DATA_DIR / ".key"          # encryption key (user-local only)
BACKUP_DIR     = DATA_DIR / "backups"
BLOCKLIST_FILE = DATA_DIR / "blocklist.json"
LOG_FILE       = DATA_DIR / "app.log"

BACKUP_DIR.mkdir(parents=True, exist_ok=True)

STARTUP_REG_KEY  = r"Software\Microsoft\Windows\CurrentVersion\Run"
STARTUP_REG_NAME = APP_NAME

# Word-boundary characters — trigger only fires after one of these
WORD_BOUNDARY_CHARS = set(
    " \t\n\r.,;:!?\"'()[]{}<>/\\|@#$%^&*-+=~`"
)

# Windows API constants
ES_PASSWORD = 0x0020   # Edit control password style

BLOCKED_WINDOW_CLASSES = {
    "CredentialUIBroker",
    "Credential Dialog Xaml Host",
    "NativeHWNDHost",
}

# Dynamic variable tokens
DYN_VARS = {
    "%DATE%":     lambda: datetime.now().strftime("%d/%m/%Y"),
    "%TIME%":     lambda: datetime.now().strftime("%H:%M"),
    "%DATETIME%": lambda: datetime.now().strftime("%d/%m/%Y %H:%M"),
    "%DAY%":      lambda: datetime.now().strftime("%A"),
    "%MONTH%":    lambda: datetime.now().strftime("%B"),
    "%YEAR%":     lambda: datetime.now().strftime("%Y"),
    "%CLIP%":     None,   # handled separately via clipboard
}

# ──────────────────────────────────────────────────────────────────
# LOGGING
# ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(APP_NAME)


# ──────────────────────────────────────────────────────────────────
# ENCRYPTION HELPERS
# ──────────────────────────────────────────────────────────────────
def _get_or_create_key() -> bytes | None:
    """Load or generate a Fernet key stored locally."""
    if not CRYPTO_AVAILABLE:
        return None
    if KEY_FILE.exists():
        return KEY_FILE.read_bytes()
    key = Fernet.generate_key()
    KEY_FILE.write_bytes(key)
    KEY_FILE.chmod(0o600)
    log.info("Encryption key generated.")
    return key


def encrypt_str(text: str, fernet) -> str:
    if not fernet:
        return text
    return fernet.encrypt(text.encode("utf-8")).decode("utf-8")


def decrypt_str(text: str, fernet) -> str:
    if not fernet:
        return text
    try:
        return fernet.decrypt(text.encode("utf-8")).decode("utf-8")
    except Exception:
        # Fallback: return as-is (handles unencrypted legacy data)
        return text


# ──────────────────────────────────────────────────────────────────
# CLIPBOARD HELPER
# ──────────────────────────────────────────────────────────────────
def get_clipboard() -> str:
    """Read current clipboard contents on Windows."""
    try:
        CF_UNICODETEXT = 13
        ctypes.windll.user32.OpenClipboard(0)
        handle = ctypes.windll.user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return ""
        ptr = ctypes.windll.kernel32.GlobalLock(handle)
        if not ptr:
            return ""
        text = ctypes.wstring_at(ptr)
        ctypes.windll.kernel32.GlobalUnlock(handle)
        return text
    except Exception:
        return ""
    finally:
        try:
            ctypes.windll.user32.CloseClipboard()
        except Exception:
            pass


# ──────────────────────────────────────────────────────────────────
# DYNAMIC VARIABLE RESOLUTION
# ──────────────────────────────────────────────────────────────────
def resolve_variables(text: str) -> str:
    """Replace %DATE%, %TIME%, %CLIP% etc. in expansion text."""
    for token, fn in DYN_VARS.items():
        if token in text:
            if token == "%CLIP%":
                text = text.replace(token, get_clipboard())
            else:
                text = text.replace(token, fn())
    return text


# ──────────────────────────────────────────────────────────────────
# WINDOWS SECURITY DETECTION
# ──────────────────────────────────────────────────────────────────
def is_password_field() -> bool:
    """Return True if the currently focused control is a password field."""
    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        focused = ctypes.windll.user32.GetFocus()
        target = focused if focused else hwnd
        # Check ES_PASSWORD style
        style = ctypes.windll.user32.GetWindowLongW(target, -16)
        if style & ES_PASSWORD:
            return True
        # Check class name
        buf = ctypes.create_unicode_buffer(256)
        ctypes.windll.user32.GetClassNameW(target, buf, 256)
        if buf.value in BLOCKED_WINDOW_CLASSES:
            return True
        return False
    except Exception:
        return False


def get_foreground_process_name() -> str:
    """Return the executable name of the current foreground window process."""
    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        pid = ctypes.wintypes.DWORD()
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        PROCESS_QUERY_LIMITED = 0x1000
        handle = ctypes.windll.kernel32.OpenProcess(
            PROCESS_QUERY_LIMITED, False, pid.value
        )
        if not handle:
            return ""
        buf = ctypes.create_unicode_buffer(260)
        ctypes.windll.kernel32.GetModuleFileNameExW(handle, None, buf, 260)
        ctypes.windll.kernel32.CloseHandle(handle)
        return Path(buf.value).name.lower()
    except Exception:
        return ""


# ──────────────────────────────────────────────────────────────────
# BLOCKLIST
# ──────────────────────────────────────────────────────────────────
class Blocklist:
    def __init__(self):
        self.entries: list[str] = []  # lowercase exe names e.g. "cmd.exe"
        self.load()

    def load(self):
        if BLOCKLIST_FILE.exists():
            try:
                self.entries = json.loads(BLOCKLIST_FILE.read_text(encoding="utf-8"))
            except Exception:
                self.entries = []

    def save(self):
        BLOCKLIST_FILE.write_text(
            json.dumps(self.entries, indent=2), encoding="utf-8"
        )

    def is_blocked(self, process_name: str) -> bool:
        return process_name.lower() in self.entries

    def add(self, name: str):
        n = name.lower().strip()
        if n and n not in self.entries:
            self.entries.append(n)
            self.save()

    def remove(self, name: str):
        self.entries = [e for e in self.entries if e != name.lower()]
        self.save()


# ──────────────────────────────────────────────────────────────────
# SNIPPET STORE
# ──────────────────────────────────────────────────────────────────
class SnippetStore:
    def __init__(self):
        self._fernet = None
        if CRYPTO_AVAILABLE:
            key = _get_or_create_key()
            if key:
                self._fernet = Fernet(key)
        self.snippets: list[dict] = []
        self.load()

    # ── Persistence ─────────────────────────────────────────────

    def load(self):
        if not SNIPPETS_FILE.exists():
            self.snippets = []
            return
        try:
            raw = json.loads(SNIPPETS_FILE.read_text(encoding="utf-8"))
            for s in raw:
                s["expansion"] = decrypt_str(s.get("expansion", ""), self._fernet)
            self.snippets = raw
            log.info(f"Loaded {len(self.snippets)} snippets.")
        except Exception as e:
            log.error(f"Failed to load snippets: {e}")
            self.snippets = []

    def save(self):
        try:
            data = []
            for s in self.snippets:
                row = dict(s)
                row["expansion"] = encrypt_str(s["expansion"], self._fernet)
                data.append(row)
            SNIPPETS_FILE.write_text(
                json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8"
            )
        except Exception as e:
            log.error(f"Failed to save snippets: {e}")

    def backup(self) -> Path | None:
        if not SNIPPETS_FILE.exists():
            return None
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dst = BACKUP_DIR / f"snippets_backup_{ts}.json"
        shutil.copy2(SNIPPETS_FILE, dst)
        # Keep last 10
        for old in sorted(BACKUP_DIR.glob("snippets_backup_*.json"))[:-10]:
            old.unlink()
        log.info(f"Backup: {dst.name}")
        return dst

    # ── CRUD ────────────────────────────────────────────────────

    def _new_id(self) -> str:
        return datetime.utcnow().strftime("%Y%m%d%H%M%S%f")

    def add(self, trigger: str, expansion: str, group: str = "General") -> dict:
        s = {
            "id": self._new_id(),
            "trigger": trigger,
            "expansion": expansion,
            "group": group,
            "use_count": 0,
            "created": datetime.utcnow().isoformat(),
            "last_used": None,
            "enabled": True,
        }
        self.snippets.append(s)
        self.save()
        return s

    def update(self, sid: str, trigger: str, expansion: str,
               group: str, enabled: bool):
        for s in self.snippets:
            if s["id"] == sid:
                s.update(trigger=trigger, expansion=expansion,
                          group=group, enabled=enabled)
                self.save()
                return

    def delete(self, sid: str):
        self.snippets = [s for s in self.snippets if s["id"] != sid]
        self.save()

    def record_use(self, sid: str):
        for s in self.snippets:
            if s["id"] == sid:
                s["use_count"] = s.get("use_count", 0) + 1
                s["last_used"] = datetime.utcnow().isoformat()
                break
        self.save()

    def get_enabled(self) -> list[dict]:
        return [s for s in self.snippets if s.get("enabled", True)]

    def groups(self) -> list[str]:
        seen, groups = set(), []
        for s in self.snippets:
            g = s.get("group", "General")
            if g not in seen:
                seen.add(g)
                groups.append(g)
        return groups or ["General"]

    # ── Import / Export ─────────────────────────────────────────

    def export_csv(self, path: str):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(
                f, fieldnames=["trigger", "expansion", "group", "use_count", "enabled"]
            )
            w.writeheader()
            for s in self.snippets:
                w.writerow({
                    "trigger":   s["trigger"],
                    "expansion": s["expansion"],
                    "group":     s.get("group", "General"),
                    "use_count": s.get("use_count", 0),
                    "enabled":   s.get("enabled", True),
                })
        log.info(f"Exported {len(self.snippets)} snippets → {path}")

    def import_csv(self, path: str) -> int:
        count = 0
        with open(path, "r", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                t = row.get("trigger", "").strip()
                e = row.get("expansion", "").strip()
                if t and e:
                    self.add(t, e, row.get("group", "General").strip() or "General")
                    count += 1
        log.info(f"Imported {count} snippets ← {path}")
        return count


# ──────────────────────────────────────────────────────────────────
# KEYBOARD ENGINE
# ──────────────────────────────────────────────────────────────────
class KeyboardEngine:
    MAX_BUFFER = 256

    def __init__(self, store: SnippetStore, blocklist: Blocklist,
                 on_expansion=None):
        self.store      = store
        self.blocklist  = blocklist
        self.on_expansion = on_expansion
        self._buffer: list[str] = []
        self._listener  = None
        self._controller = pynput_kb.Controller()
        self._lock      = threading.Lock()
        self._running   = False

    # ── Lifecycle ───────────────────────────────────────────────

    def start(self):
        if self._running:
            return
        self._running = True
        self._listener = pynput_kb.Listener(on_press=self._on_press, suppress=False)
        self._listener.daemon = True
        self._listener.start()
        log.info("Engine started.")

    def stop(self):
        self._running = False
        if self._listener:
            self._listener.stop()
            self._listener = None
        log.info("Engine stopped.")

    def toggle(self):
        if self._running:
            self.stop()
        else:
            self.start()

    # ── Key Handler ─────────────────────────────────────────────

    def _on_press(self, key):
        if not self._running:
            return
        if is_password_field():
            return
        if self.blocklist.is_blocked(get_foreground_process_name()):
            return

        with self._lock:
            try:
                char = key.char
                if char is None:
                    return
                self._buffer.append(char)
                if len(self._buffer) > self.MAX_BUFFER:
                    self._buffer = self._buffer[-self.MAX_BUFFER:]

            except AttributeError:
                # Special key handling
                if key == pynput_kb.Key.backspace:
                    if self._buffer:
                        self._buffer.pop()

                elif key in (pynput_kb.Key.space, pynput_kb.Key.tab,
                             pynput_kb.Key.enter, pynput_kb.Key.return_):
                    boundary = {
                        pynput_kb.Key.space:  " ",
                        pynput_kb.Key.tab:    "\t",
                        pynput_kb.Key.enter:  "\n",
                        pynput_kb.Key.return_: "\n",
                    }.get(key, " ")
                    self._buffer.append(boundary)
                    if len(self._buffer) > self.MAX_BUFFER:
                        self._buffer = self._buffer[-self.MAX_BUFFER:]
                    self._check_and_expand()

                elif key in (pynput_kb.Key.esc, pynput_kb.Key.delete):
                    self._buffer.clear()

                elif key in (pynput_kb.Key.left, pynput_kb.Key.right,
                             pynput_kb.Key.up, pynput_kb.Key.down,
                             pynput_kb.Key.home, pynput_kb.Key.end,
                             pynput_kb.Key.page_up, pynput_kb.Key.page_down):
                    self._buffer.clear()

    # ── Matching ────────────────────────────────────────────────

    def _check_and_expand(self):
        buf = "".join(self._buffer)
        # Strip trailing boundary to get typed content
        content = buf.rstrip(" \t\n\r")

        # Longest trigger wins
        snippets = sorted(
            self.store.get_enabled(),
            key=lambda s: len(s["trigger"]),
            reverse=True,
        )

        for snippet in snippets:
            trigger = snippet["trigger"]
            if not content.endswith(trigger):
                continue
            idx = len(content) - len(trigger)
            boundary_ok = (
                idx == 0 or content[idx - 1] in WORD_BOUNDARY_CHARS
            )
            if not boundary_ok:
                continue
            self._expand(snippet, trigger)
            return

    def _expand(self, snippet: dict, trigger: str):
        expansion = resolve_variables(snippet["expansion"])
        delete_count = len(trigger) + 1  # trigger + boundary char

        threading.Thread(
            target=self.store.record_use, args=(snippet["id"],), daemon=True
        ).start()

        if self.on_expansion:
            try:
                self.on_expansion(trigger, expansion)
            except Exception:
                pass

        self._buffer.clear()

        def do_expand():
            time.sleep(0.02)
            for _ in range(delete_count):
                self._controller.press(pynput_kb.Key.backspace)
                self._controller.release(pynput_kb.Key.backspace)
                time.sleep(0.004)
            self._controller.type(expansion)
            log.info(f"Expanded '{trigger}' → '{expansion[:60]}'")

        threading.Thread(target=do_expand, daemon=True).start()


# ──────────────────────────────────────────────────────────────────
# GLOBAL HOTKEY  (Ctrl+Alt+P  →  pause / resume)
# ──────────────────────────────────────────────────────────────────
class GlobalHotkey:
    """Listens for Ctrl+Alt+P in a separate listener thread."""

    def __init__(self, callback):
        self._callback  = callback
        self._pressed   = set()
        self._listener  = None

    def start(self):
        self._listener = pynput_kb.Listener(
            on_press=self._on_press, on_release=self._on_release
        )
        self._listener.daemon = True
        self._listener.start()

    def stop(self):
        if self._listener:
            self._listener.stop()

    def _on_press(self, key):
        self._pressed.add(key)
        ctrl  = pynput_kb.Key.ctrl_l  in self._pressed or \
                pynput_kb.Key.ctrl_r  in self._pressed
        alt   = pynput_kb.Key.alt_l   in self._pressed or \
                pynput_kb.Key.alt_r   in self._pressed
        try:
            p = hasattr(key, "char") and key.char == "p"
        except Exception:
            p = False
        if ctrl and alt and p:
            self._callback()

    def _on_release(self, key):
        self._pressed.discard(key)


# ──────────────────────────────────────────────────────────────────
# STARTUP MANAGER
# ──────────────────────────────────────────────────────────────────
class StartupManager:
    @staticmethod
    def _entry() -> str:
        if getattr(sys, "frozen", False):
            return f'"{sys.executable}"'
        pythonw = Path(sys.executable).parent / "pythonw.exe"
        exe = str(pythonw) if pythonw.exists() else sys.executable
        return f'"{exe}" "{Path(__file__).resolve()}"'

    @staticmethod
    def is_enabled() -> bool:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY) as k:
                winreg.QueryValueEx(k, STARTUP_REG_NAME)
                return True
        except FileNotFoundError:
            return False
        except Exception:
            return False

    @staticmethod
    def enable() -> bool:
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY, 0, winreg.KEY_SET_VALUE
            ) as k:
                winreg.SetValueEx(
                    k, STARTUP_REG_NAME, 0, winreg.REG_SZ, StartupManager._entry()
                )
            log.info(f"Startup enabled → {StartupManager._entry()}")
            return True
        except Exception as e:
            log.error(f"Startup enable failed: {e}")
            return False

    @staticmethod
    def disable() -> bool:
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY, 0, winreg.KEY_SET_VALUE
            ) as k:
                winreg.DeleteValue(k, STARTUP_REG_NAME)
            return True
        except FileNotFoundError:
            return True
        except Exception as e:
            log.error(f"Startup disable failed: {e}")
            return False


# ──────────────────────────────────────────────────────────────────
# UPDATE CHECKER
# ──────────────────────────────────────────────────────────────────
def check_for_update(current_version: str, callback):
    """Non-blocking GitHub release check. Calls callback(latest_version, url)
    if a newer version is found."""
    if not REQUESTS_AVAILABLE:
        return

    def _check():
        try:
            url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            r = requests.get(url, timeout=5)
            if r.status_code == 200:
                data = r.json()
                latest = data.get("tag_name", "").lstrip("v")
                html_url = data.get("html_url", "")
                if latest and latest != current_version:
                    callback(latest, html_url)
        except Exception:
            pass  # Silent fail — no network noise

    threading.Thread(target=_check, daemon=True).start()


# ──────────────────────────────────────────────────────────────────
# SYSTEM TRAY
# ──────────────────────────────────────────────────────────────────
def _try_start_tray(app_ref):
    try:
        import pystray
        from PIL import Image, ImageDraw, ImageFont

        def _make_icon(paused: bool = False):
            size = 64
            bg   = (60, 60, 60) if paused else (26, 110, 216)
            img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            d    = ImageDraw.Draw(img)
            d.rounded_rectangle([2, 2, size - 2, size - 2], radius=12, fill=bg)
            d.text((14, 18), "ST" if not paused else "⏸", fill=(255, 255, 255))
            return img

        def show(_i, _it):
            app_ref.root.deiconify()
            app_ref.root.lift()
            app_ref.root.focus_force()

        def toggle_pause(_i, _it):
            app_ref.engine.toggle()
            app_ref._update_status()

        def quit_app(_i, _it):
            _i.stop()
            app_ref.quit()

        menu = pystray.Menu(
            pystray.MenuItem("Open", show, default=True),
            pystray.MenuItem("Pause / Resume  (Ctrl+Alt+P)", toggle_pause),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(f"v{APP_VERSION}", None, enabled=False),
            pystray.MenuItem("Quit", quit_app),
        )

        icon = pystray.Icon(APP_NAME, _make_icon(), APP_NAME, menu)
        threading.Thread(target=icon.run, daemon=True).start()
        log.info("System tray icon active.")
        return icon
    except ImportError:
        log.info("pystray/Pillow not installed — no tray icon.")
        return None


# ──────────────────────────────────────────────────────────────────
# SNIPPET DIALOG
# ──────────────────────────────────────────────────────────────────
class SnippetDialog(tk.Toplevel):
    def __init__(self, app, snippet: dict | None):
        super().__init__(app.root)
        self.app     = app
        self.snippet = snippet
        self.title("New Snippet" if snippet is None else "Edit Snippet")
        self.geometry("580x480")
        self.resizable(True, True)
        self.grab_set()
        self._build()
        if snippet:
            self._populate()

    def _build(self):
        pad = {"padx": 14, "pady": 4}

        # ── Trigger ─────────────────────────────────────────────
        ttk.Label(self, text="Trigger (abbreviation):").pack(anchor=tk.W, **pad)
        self._trigger_var = tk.StringVar()
        ttk.Entry(self, textvariable=self._trigger_var, width=42).pack(
            anchor=tk.W, **pad
        )
        tk.Label(
            self,
            text=(
                "Only fires when preceded by a space, punctuation, or start of input.\n"
                'e.g. trigger "ono" will NOT fire inside "pro-bono".'
            ),
            fg="grey50", font=("Segoe UI", 8), justify=tk.LEFT,
        ).pack(anchor=tk.W, padx=14, pady=(0, 4))

        # ── Expansion ───────────────────────────────────────────
        ttk.Label(self, text="Expansion:").pack(anchor=tk.W, **pad)
        self._exp_text = tk.Text(self, height=9, wrap="word", font=("Segoe UI", 10))
        self._exp_text.pack(fill=tk.BOTH, expand=True, **pad)

        # Variable picker
        var_frm = ttk.Frame(self)
        var_frm.pack(fill=tk.X, padx=14, pady=(0, 4))
        tk.Label(var_frm, text="Insert variable:", fg="grey50",
                 font=("Segoe UI", 8)).pack(side=tk.LEFT)
        for token in DYN_VARS:
            ttk.Button(
                var_frm, text=token, width=len(token) + 1,
                command=lambda t=token: self._insert_var(t),
            ).pack(side=tk.LEFT, padx=2)

        # ── Group ───────────────────────────────────────────────
        grp = ttk.Frame(self)
        grp.pack(fill=tk.X, **pad)
        ttk.Label(grp, text="Group:").pack(side=tk.LEFT)
        self._group_var = tk.StringVar(value="General")
        ttk.Combobox(
            grp, textvariable=self._group_var,
            values=self.app.store.groups(), width=18,
        ).pack(side=tk.LEFT, padx=6)
        tk.Label(grp, text="(type to create new)",
                 fg="grey50", font=("Segoe UI", 8)).pack(side=tk.LEFT)

        # ── Enabled ─────────────────────────────────────────────
        self._enabled_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self, text="Enabled", variable=self._enabled_var).pack(
            anchor=tk.W, **pad
        )

        # ── Buttons ─────────────────────────────────────────────
        btns = ttk.Frame(self)
        btns.pack(fill=tk.X, padx=14, pady=8)
        ttk.Button(btns, text="Save",   command=self._save).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=4)

    def _insert_var(self, token: str):
        self._exp_text.insert(tk.INSERT, token)

    def _populate(self):
        s = self.snippet
        self._trigger_var.set(s.get("trigger", ""))
        self._exp_text.insert("1.0", s.get("expansion", ""))
        self._group_var.set(s.get("group", "General"))
        self._enabled_var.set(s.get("enabled", True))

    def _save(self):
        trigger   = self._trigger_var.get().strip()
        expansion = self._exp_text.get("1.0", tk.END).rstrip("\n")
        group     = self._group_var.get().strip() or "General"
        enabled   = self._enabled_var.get()

        if not trigger:
            messagebox.showwarning("Missing trigger", "Enter a trigger.", parent=self)
            return
        if not expansion:
            messagebox.showwarning("Missing expansion", "Enter expansion text.", parent=self)
            return

        if self.snippet is None:
            self.app.store.add(trigger, expansion, group)
        else:
            self.app.store.update(
                self.snippet["id"], trigger, expansion, group, enabled
            )

        self.app._refresh_tree()
        self.destroy()


# ──────────────────────────────────────────────────────────────────
# BLOCKLIST DIALOG
# ──────────────────────────────────────────────────────────────────
class BlocklistDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.title("App Blocklist — never expand in these apps")
        self.geometry("480x380")
        self.grab_set()
        self._build()
        self._refresh()

    def _build(self):
        tk.Label(
            self,
            text=(
                "Add the executable name (e.g. WindowsTerminal.exe, KeePass.exe).\n"
                "Expansions are silently disabled when these apps are in focus."
            ),
            fg="grey40", justify=tk.LEFT, font=("Segoe UI", 9),
            wraplength=440,
        ).pack(padx=12, pady=8, anchor=tk.W)

        self.lb = tk.Listbox(self, font=("Consolas", 10))
        self.lb.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)

        frm = ttk.Frame(self)
        frm.pack(fill=tk.X, padx=12, pady=6)

        self._entry_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self._entry_var, width=28).pack(side=tk.LEFT, padx=4)
        ttk.Button(frm, text="Add",    command=self._add).pack(side=tk.LEFT, padx=2)
        ttk.Button(frm, text="Remove", command=self._remove).pack(side=tk.LEFT, padx=2)

        # Quick-add current foreground app
        ttk.Button(
            frm, text="Block Current App",
            command=self._block_current,
        ).pack(side=tk.RIGHT, padx=4)

    def _refresh(self):
        self.lb.delete(0, tk.END)
        for e in self.app.blocklist.entries:
            self.lb.insert(tk.END, e)

    def _add(self):
        name = self._entry_var.get().strip()
        if name:
            self.app.blocklist.add(name)
            self._entry_var.set("")
            self._refresh()

    def _remove(self):
        sel = self.lb.curselection()
        if sel:
            name = self.lb.get(sel[0])
            self.app.blocklist.remove(name)
            self._refresh()

    def _block_current(self):
        name = get_foreground_process_name()
        if name:
            self.app.blocklist.add(name)
            self._refresh()
        else:
            messagebox.showinfo("Not found",
                "Could not detect the current foreground application.", parent=self)


# ──────────────────────────────────────────────────────────────────
# MAIN APP
# ──────────────────────────────────────────────────────────────────
class App:
    def __init__(self):
        self.store     = SnippetStore()
        self.blocklist = Blocklist()
        self.engine    = KeyboardEngine(
            self.store, self.blocklist, on_expansion=self._on_expansion
        )
        self._tray     = None
        self._session_expansions = 0

        self.root = tk.Tk()
        self.root.title(f"Smart Text Replacer  v{APP_VERSION}")
        self.root.geometry("960x620")
        self.root.minsize(720, 440)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_ui()
        self._refresh_tree()
        self._update_status()

        # Start minimised to tray
        self.root.withdraw()

        # Engine + hotkey
        self.engine.start()
        self._hotkey = GlobalHotkey(self._toggle_engine)
        self._hotkey.start()

        # Tray
        self._tray = _try_start_tray(self)

        # Auto-backup
        self.store.backup()

        # Update check (non-blocking, silent)
        check_for_update(APP_VERSION, self._on_update_available)

    # ── UI Construction ─────────────────────────────────────────

    def _build_ui(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", padding=4)

        # Menu
        mb = tk.Menu(self.root)
        self.root.config(menu=mb)

        fm = tk.Menu(mb, tearoff=0)
        mb.add_cascade(label="File", menu=fm)
        fm.add_command(label="Import CSV…", command=self._import_csv)
        fm.add_command(label="Export CSV…", command=self._export_csv)
        fm.add_separator()
        fm.add_command(label="Backup Now",           command=self._backup_now)
        fm.add_command(label="Open Backup Folder",   command=self._open_backup_folder)
        fm.add_separator()
        fm.add_command(label="Exit",                 command=self.quit)

        sm = tk.Menu(mb, tearoff=0)
        mb.add_cascade(label="Settings", menu=sm)
        sm.add_command(label="App Blocklist…",       command=self._show_blocklist)
        sm.add_command(label="Open Log File",        command=self._open_log)

        hm = tk.Menu(mb, tearoff=0)
        mb.add_cascade(label="Help", menu=hm)
        hm.add_command(label="Variable Reference",   command=self._show_var_help)
        hm.add_command(label="About",                command=self._show_about)

        # Status bar
        top = tk.Frame(self.root, bg="#1a6ed8", height=38)
        top.pack(fill=tk.X)
        top.pack_propagate(False)

        self._status_var = tk.StringVar(value="● Active")
        tk.Label(top, textvariable=self._status_var,
                 bg="#1a6ed8", fg="white",
                 font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=14, pady=8)

        self._stats_var = tk.StringVar(value="Session expansions: 0")
        tk.Label(top, textvariable=self._stats_var,
                 bg="#1a6ed8", fg="#cce0ff",
                 font=("Segoe UI", 9)).pack(side=tk.RIGHT, padx=14)

        ttk.Button(top, text="Pause / Resume  (Ctrl+Alt+P)",
                   command=self._toggle_engine).pack(side=tk.RIGHT, padx=6, pady=4)

        # Notebook
        nb = ttk.Notebook(self.root)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=(6, 4))

        snip_tab = ttk.Frame(nb)
        nb.add(snip_tab, text="  Snippets  ")
        self._build_snippets_tab(snip_tab)

        stats_tab = ttk.Frame(nb)
        nb.add(stats_tab, text="  Statistics  ")
        self._build_stats_tab(stats_tab)

        settings_tab = ttk.Frame(nb)
        nb.add(settings_tab, text="  Settings  ")
        self._build_settings_tab(settings_tab)

    def _build_snippets_tab(self, parent):
        tb = ttk.Frame(parent)
        tb.pack(fill=tk.X, pady=(6, 2), padx=6)

        ttk.Button(tb, text="＋ New",       command=self._new_snippet).pack(side=tk.LEFT, padx=2)
        ttk.Button(tb, text="✎ Edit",       command=self._edit_snippet).pack(side=tk.LEFT, padx=2)
        ttk.Button(tb, text="✕ Delete",     command=self._delete_snippet).pack(side=tk.LEFT, padx=2)
        ttk.Button(tb, text="⧫ Duplicate",  command=self._duplicate_snippet).pack(side=tk.LEFT, padx=2)

        tk.Label(tb, text="Group:").pack(side=tk.LEFT, padx=(16, 2))
        self._grp_var = tk.StringVar(value="All")
        self._grp_combo = ttk.Combobox(tb, textvariable=self._grp_var,
                                        state="readonly", width=14)
        self._grp_combo.pack(side=tk.LEFT, padx=2)
        self._grp_combo.bind("<<ComboboxSelected>>", lambda _: self._refresh_tree())

        tk.Label(tb, text="Search:").pack(side=tk.LEFT, padx=(16, 2))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._refresh_tree())
        ttk.Entry(tb, textvariable=self._search_var, width=18).pack(side=tk.LEFT)

        frm = ttk.Frame(parent)
        frm.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

        cols = ("trigger", "expansion", "group", "use_count", "enabled")
        self.tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="browse")

        self.tree.heading("trigger",   text="Trigger",   anchor=tk.W)
        self.tree.heading("expansion", text="Expansion", anchor=tk.W)
        self.tree.heading("group",     text="Group",     anchor=tk.W)
        self.tree.heading("use_count", text="Uses",      anchor=tk.CENTER)
        self.tree.heading("enabled",   text="On",        anchor=tk.CENTER)

        self.tree.column("trigger",   width=120, minwidth=80)
        self.tree.column("expansion", width=400, minwidth=200)
        self.tree.column("group",     width=100, minwidth=70)
        self.tree.column("use_count", width=55,  minwidth=40, anchor=tk.CENTER)
        self.tree.column("enabled",   width=45,  minwidth=35, anchor=tk.CENTER)

        vsb = ttk.Scrollbar(frm, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<Double-1>", lambda _: self._edit_snippet())
        self.tree.bind("<Delete>",   lambda _: self._delete_snippet())

    def _build_stats_tab(self, parent):
        self._stats_text = tk.Text(
            parent, state="disabled", wrap="word",
            font=("Consolas", 10), relief="flat", bg=parent.cget("bg"),
        )
        self._stats_text.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        ttk.Button(parent, text="↻ Refresh", command=self._refresh_stats).pack(pady=(0, 8))
        self._refresh_stats()

    def _build_settings_tab(self, parent):
        # Startup
        sf = ttk.LabelFrame(parent, text="Windows Startup", padding=10)
        sf.pack(fill=tk.X, padx=16, pady=10)
        self._startup_var = tk.BooleanVar(value=StartupManager.is_enabled())
        ttk.Checkbutton(
            sf,
            text="Launch Smart Text Replacer automatically when Windows starts",
            variable=self._startup_var,
            command=self._toggle_startup,
        ).pack(anchor=tk.W)
        tk.Label(
            sf,
            text=(
                "Registers in HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run.\n"
                "App starts minimised to the system tray — no window, no clutter."
            ),
            fg="grey40", font=("Segoe UI", 8), justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=(4, 0))

        # Hotkey
        hf = ttk.LabelFrame(parent, text="Global Hotkey", padding=10)
        hf.pack(fill=tk.X, padx=16, pady=4)
        tk.Label(
            hf,
            text="Ctrl + Alt + P  —  Pause / Resume expansion from anywhere, without opening the app.",
            font=("Segoe UI", 9),
        ).pack(anchor=tk.W)

        # Security
        sec = ttk.LabelFrame(parent, text="Security", padding=10)
        sec.pack(fill=tk.X, padx=16, pady=4)
        enc_status = "✓ Fernet/AES-128 encryption active" if CRYPTO_AVAILABLE else \
                     "⚠ cryptography package not installed — snippets stored as plaintext"
        tk.Label(
            sec,
            text=(
                f"• {enc_status}\n"
                "• Expansion disabled automatically in password fields\n"
                "• Expansion disabled in Windows credential dialogs\n"
                "• App blocklist available under Settings → App Blocklist\n"
                "• Zero telemetry — no network requests except optional update check\n"
                f"• Data directory: {DATA_DIR}"
            ),
            justify=tk.LEFT, font=("Segoe UI", 9),
        ).pack(anchor=tk.W)
        ttk.Button(sec, text="Manage App Blocklist…",
                   command=self._show_blocklist).pack(anchor=tk.W, pady=(6, 0))

        # Backup
        bf = ttk.LabelFrame(parent, text="Backup", padding=10)
        bf.pack(fill=tk.X, padx=16, pady=4)
        row = ttk.Frame(bf)
        row.pack(fill=tk.X)
        ttk.Button(row, text="Back Up Now",         command=self._backup_now).pack(side=tk.LEFT, padx=4)
        ttk.Button(row, text="Open Backup Folder",  command=self._open_backup_folder).pack(side=tk.LEFT, padx=4)
        tk.Label(bf, text="Auto-backup runs on every launch. Last 10 kept.",
                 fg="grey40", font=("Segoe UI", 8)).pack(anchor=tk.W, pady=(4, 0))

    # ── Tree ────────────────────────────────────────────────────

    def _refresh_tree(self):
        groups = ["All"] + self.store.groups()
        self._grp_combo["values"] = groups
        if self._grp_var.get() not in groups:
            self._grp_var.set("All")

        search = self._search_var.get().lower()
        gf     = self._grp_var.get()

        self.tree.delete(*self.tree.get_children())
        for s in self.store.snippets:
            if gf != "All" and s.get("group", "General") != gf:
                continue
            if search and search not in s["trigger"].lower() \
                      and search not in s["expansion"].lower():
                continue
            self.tree.insert("", tk.END, iid=s["id"], values=(
                s["trigger"],
                s["expansion"][:80] + ("…" if len(s["expansion"]) > 80 else ""),
                s.get("group", "General"),
                s.get("use_count", 0),
                "✓" if s.get("enabled", True) else "✗",
            ))

    def _selected_snippet(self) -> dict | None:
        sel = self.tree.selection()
        if not sel:
            return None
        sid = sel[0]
        return next((s for s in self.store.snippets if s["id"] == sid), None)

    # ── Snippet Actions ──────────────────────────────────────────

    def _new_snippet(self):
        SnippetDialog(self, None)

    def _edit_snippet(self):
        s = self._selected_snippet()
        if not s:
            messagebox.showinfo("Select snippet", "Please select a snippet to edit.")
            return
        SnippetDialog(self, s)

    def _delete_snippet(self):
        s = self._selected_snippet()
        if not s:
            return
        if messagebox.askyesno("Delete", f"Delete snippet '{s['trigger']}'?"):
            self.store.delete(s["id"])
            self._refresh_tree()

    def _duplicate_snippet(self):
        s = self._selected_snippet()
        if not s:
            return
        self.store.add(s["trigger"] + "2", s["expansion"], s.get("group", "General"))
        self._refresh_tree()

    # ── Import / Export ──────────────────────────────────────────

    def _export_csv(self):
        path = filedialog.asksaveasfilename(
            title="Export Snippets", defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
        )
        if path:
            self.store.export_csv(path)
            messagebox.showinfo("Export", f"Exported to:\n{path}")

    def _import_csv(self):
        path = filedialog.askopenfilename(
            title="Import Snippets", filetypes=[("CSV files", "*.csv")],
        )
        if path:
            n = self.store.import_csv(path)
            self._refresh_tree()
            messagebox.showinfo("Import", f"Imported {n} snippets.")

    # ── Backup ───────────────────────────────────────────────────

    def _backup_now(self):
        dst = self.store.backup()
        if dst:
            messagebox.showinfo("Backup", f"Saved to:\n{dst}")

    def _open_backup_folder(self):
        os.startfile(str(BACKUP_DIR))

    # ── Settings ─────────────────────────────────────────────────

    def _toggle_startup(self):
        if self._startup_var.get():
            if not StartupManager.enable():
                messagebox.showerror("Error", "Could not enable startup.")
                self._startup_var.set(False)
        else:
            StartupManager.disable()

    def _show_blocklist(self):
        BlocklistDialog(self)

    # ── Engine ───────────────────────────────────────────────────

    def _toggle_engine(self):
        self.engine.toggle()
        self._update_status()

    def _update_status(self):
        if self.engine._running:
            self._status_var.set("● Active — expansions enabled")
        else:
            self._status_var.set("⏸ Paused — expansions disabled  (Ctrl+Alt+P to resume)")

    def _on_expansion(self, trigger: str, _expansion: str):
        self._session_expansions += 1
        self._stats_var.set(f"Session expansions: {self._session_expansions}")

    # ── Stats ─────────────────────────────────────────────────────

    def _refresh_stats(self):
        s    = self.store.snippets
        total = len(s)
        enabled = sum(1 for x in s if x.get("enabled", True))
        uses  = sum(x.get("use_count", 0) for x in s)
        top   = sorted(s, key=lambda x: x.get("use_count", 0), reverse=True)[:10]

        lines = [
            f"Smart Text Replacer  v{APP_VERSION}",
            f"Data: {DATA_DIR}",
            f"Encryption: {'Fernet/AES-128 ✓' if CRYPTO_AVAILABLE else 'Not installed ⚠'}",
            "",
            f"Total snippets :  {total}",
            f"Enabled        :  {enabled}",
            f"Groups         :  {len(self.store.groups())}",
            f"All-time uses  :  {uses}",
            f"Session uses   :  {self._session_expansions}",
            "",
            "── Top 10 by usage ─────────────────────────────",
        ]
        for i, x in enumerate(top, 1):
            lines.append(
                f"  {i:2}. [{x.get('use_count',0):>4}×]  "
                f"{x['trigger']:<20}  →  {x['expansion'][:50]}"
            )

        if not top:
            lines.append("  (no snippets yet)")

        self._stats_text.config(state="normal")
        self._stats_text.delete("1.0", tk.END)
        self._stats_text.insert(tk.END, "\n".join(lines))
        self._stats_text.config(state="disabled")

    # ── Misc ──────────────────────────────────────────────────────

    def _open_log(self):
        os.startfile(str(LOG_FILE))

    def _show_var_help(self):
        msg = "Dynamic Variables — use these in any expansion:\n\n"
        for token, fn in DYN_VARS.items():
            example = fn() if fn else "(clipboard contents)"
            msg += f"  {token:<14} →  {example}\n"
        messagebox.showinfo("Variable Reference", msg)

    def _show_about(self):
        messagebox.showinfo(
            f"Smart Text Replacer v{APP_VERSION}",
            f"Smart Text Replacer  v{APP_VERSION}\n\n"
            "Privacy-first, open-source Windows text expansion.\n\n"
            "• Word-boundary aware\n"
            "• Password field safe\n"
            "• Dynamic variables\n"
            "• App blocklist\n"
            "• Encrypted storage\n\n"
            f"GitHub: github.com/{GITHUB_REPO}\n"
            "License: MIT",
        )

    def _on_update_available(self, latest: str, url: str):
        self.root.after(
            0,
            lambda: messagebox.showinfo(
                "Update Available",
                f"Smart Text Replacer v{latest} is available.\n\n"
                f"Download at:\n{url}",
            ),
        )

    def _on_close(self):
        self.root.withdraw()  # Minimise to tray

    def quit(self):
        self.engine.stop()
        self._hotkey.stop()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


# ──────────────────────────────────────────────────────────────────
# ENTRY POINT
# ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    App().run()
