"""
Clippy v14 — Windows Desktop App
Fixes:
  1. Ctrl+D → navigate cards → Enter = PASTE to destination (simulate Ctrl+V), then hide
  2. After Enter-paste, app hides to tray automatically
  3. No keyboard hints anywhere in UI, AI Smart Group removed
  4. Re-copy same word after deletion now re-captures (cooldown cleared on delete)
  5. Redesigned professional logo
  6. Default window size matches screenshot (~660×560)
"""

import sys, json, uuid, os, csv, io, threading, time, traceback, logging, subprocess
from pathlib import Path
from datetime import datetime, timedelta

from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QSystemTrayIcon, QMenu
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebChannel import QWebChannel
from PyQt6.QtCore import QObject, pyqtSlot, pyqtSignal, QUrl, QThread, Qt, QTimer, QPoint, QRectF
from PyQt6.QtGui import QIcon, QPixmap, QColor, QPainter, QBrush, QPen, QRadialGradient, QLinearGradient, QCursor

try:
    import pyperclip; CLIPBOARD_OK = True
except ImportError:
    CLIPBOARD_OK = False

try:
    import keyboard as kb_lib; HAS_KEYBOARD = True
except ImportError:
    HAS_KEYBOARD = False

DATA_FILE  = Path(__file__).parent / "clippy_data.json"
DEFAULT_HOTKEY = "ctrl+d"
BACKUP_DIR = Path.home() / "Documents" / "Clippy Backups"

# ══════════════════════════════════════════════════════════════════════════════
# LOGGING — rotating crash + event log written next to the script
# ══════════════════════════════════════════════════════════════════════════════
LOG_FILE = Path(__file__).parent / "clippy.log"

def _setup_logging():
    """Configure a rotating file logger + console mirror."""
    from logging.handlers import RotatingFileHandler

    fmt = logging.Formatter(
        "%(asctime)s  %(levelname)-8s  %(name)s — %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    root = logging.getLogger()
    root.setLevel(logging.DEBUG)

    # Rotate at 1 MB, keep 3 old files (clippy.log, clippy.log.1, .2)
    fh = RotatingFileHandler(LOG_FILE, maxBytes=1_000_000, backupCount=3,
                             encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    root.addHandler(fh)

    # Mirror WARNING+ to stderr so a terminal run also shows problems
    sh = logging.StreamHandler(sys.stderr)
    sh.setLevel(logging.WARNING)
    sh.setFormatter(fmt)
    root.addHandler(sh)

_setup_logging()
log = logging.getLogger("clippy")

def _log_exc(label: str, exc: BaseException | None = None):
    """Log a full traceback under *label*. Call from any except block."""
    tb = traceback.format_exc()
    log.error("CRASH in %s:\n%s", label, tb)

def _install_global_excepthook():
    """Catch any unhandled exception on the main thread and write a crash log."""
    def _hook(exc_type, exc_value, exc_tb):
        msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        log.critical("UNHANDLED EXCEPTION (main thread):\n%s", msg)
        # Also write a dedicated crash snapshot for easy sharing
        crash_file = Path(__file__).parent / "clippy_crash.log"
        try:
            crash_file.write_text(
                f"=== Clippy crash report ===\n"
                f"Time : {time.strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Python: {sys.version}\n\n"
                f"{msg}",
                encoding="utf-8",
            )
        except Exception:
            pass
        sys.__excepthook__(exc_type, exc_value, exc_tb)

    sys.excepthook = _hook

_install_global_excepthook()
log.info("Clippy starting — log file: %s", LOG_FILE)

# ══════════════════════════════════════════════════════════════════════════════
# STARTUP REGISTRATION — ensure Clippy runs on every Windows login
# ══════════════════════════════════════════════════════════════════════════════
def _ensure_startup():
    """
    Register the lightweight clippy_launcher.py (not clippyo.py itself) to run
    at login.  The launcher starts in <1 second, shows the tray balloon
    immediately, then launches clippyo.py as a detached background subprocess.
    This gives users a <3 second notification even on cold boot.

    Use the HKCU Run key so Clippy launches as soon as the user signs in
    without depending on a Startup-folder batch file. Clean up any old
    Startup-folder launcher to avoid duplicate launches on future logins.
    """
    # Use the launcher script if it exists next to this file.
    # Fall back to registering this script directly if launcher is missing.
    _this    = Path(__file__).resolve()
    launcher = _this.parent / "clippy_launcher.py"
    script   = launcher if launcher.exists() else _this
    log.info("_ensure_startup: registering %s", script.name)

    try:
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"

        pythonw = Path(sys.executable).parent / "pythonw.exe"
        exe = str(pythonw) if pythonw.exists() else sys.executable

        value = f'"{exe}" "{script}"'

        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "Clippy", 0, winreg.REG_SZ, value)
            saved, _ = winreg.QueryValueEx(key, "Clippy")
        if saved == value:
            log.info("Registry Run key OK -> HKCU\\...\\Run\\Clippy = %s", value)
        else:
            log.error("Registry Run key written but read-back mismatch! "
                      "Written=%r  Read=%r", value, saved)
    except Exception as exc:
        log.error("Registry method failed - Clippy will NOT auto-start: %s", exc)
        return

    try:
        startup_dir = Path(os.environ["APPDATA"]) / \
            "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
        for stale_name in ("Clippy.bat", "Clippy.disabled.bat"):
            stale_path = startup_dir / stale_name
            if stale_path.exists():
                stale_path.unlink()
                log.info("Removed stale Startup launcher -> %s", stale_path)
    except Exception as exc:
        log.warning("Could not remove stale Startup launcher: %s", exc)

def _disable_startup():
    """Remove HKCU Run key so Clippy no longer auto-starts at login."""
    try:
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
            winreg.DeleteValue(key, "Clippy")
        log.info("Removed HKCU Run key for Clippy")
    except FileNotFoundError:
        log.info("HKCU Run key for Clippy already absent")
    except Exception as exc:
        _log_exc("_disable_startup()", exc)

def _set_launch_at_startup(enabled: bool) -> bool:
    """Toggle Windows startup registration. Returns True on success."""
    try:
        if enabled:
            _ensure_startup()
        else:
            _disable_startup()
        return True
    except Exception as exc:
        _log_exc("_set_launch_at_startup()", exc)
        return False

# ══════════════════════════════════════════════════════════════════════════════
# STATE
# ══════════════════════════════════════════════════════════════════════════════
def _gid(): return uuid.uuid4().hex[:12]   # 48-bit entropy — collision-safe up to millions of entries
def _entry(text=""): return {"id":_gid(),"text":text,"copy_count":0,"pinned":False}
def _row(ci=0):
    return {"id":_gid(),"color_idx":ci,"entries":[_entry()],"pinned":False,"label":""}

state = {
    "rows": [],
    "cursor": {"row_idx":0,"entry_idx":0},
    "auto_capture": True,
    "last_clip": "",
    "last_clip_time": 0,
    "history": [],
    "history_enabled": True,
    "theme": "daylight",
    "hotkey": DEFAULT_HOTKEY,
    "poll_rate": 500,
    "paste_plain_text": True,
    "launch_at_startup": True,
    "backup_enabled": True,
    "history_limit": 10,
    "last_backup": "",
}

def _normalize_hotkey(combo: str) -> str:
    parts = [p.strip().lower() for p in combo.split("+") if p.strip()]
    if not parts:
        return ""
    norm = []
    for p in parts:
        if p in ("win", "cmd", "meta"):
            p = "windows"
        if p not in norm:
            norm.append(p)
    return "+".join(norm)

def _valid_hotkey(combo: str) -> bool:
    parts = [p for p in combo.split("+") if p]
    if len(parts) < 2 or len(parts) > 4:
        return False
    mods = {"ctrl", "alt", "shift", "windows"}
    has_mod = any(p in mods for p in parts)
    has_trigger = any(p not in mods for p in parts)
    return has_mod and has_trigger

def load():
    if DATA_FILE.exists():
        try:
            d = json.loads(DATA_FILE.read_text("utf-8"))
            state["rows"] = d.get("rows", [])
            state["auto_capture"] = bool(d.get("auto_capture", True))
            state["history_enabled"] = d.get("history_enabled", True)
            state["theme"] = d.get("theme", "daylight")
            state["hotkey"] = _normalize_hotkey(d.get("hotkey", DEFAULT_HOTKEY)) or DEFAULT_HOTKEY
            state["poll_rate"] = int(d.get("poll_rate", 500))
            state["paste_plain_text"] = bool(d.get("paste_plain_text", True))
            state["launch_at_startup"] = bool(d.get("launch_at_startup", True))
            state["backup_enabled"] = bool(d.get("backup_enabled", True))
            state["history_limit"] = int(d.get("history_limit", 10))
            state["last_backup"] = d.get("last_backup", "")

            # Normalize legacy or partial data so older JSON versions always load safely.
            fixed_rows = []
            for i, row in enumerate(state["rows"]):
                if not isinstance(row, dict):
                    continue
                entries = row.get("entries", [])
                if not isinstance(entries, list):
                    entries = []
                fixed_entries = []
                for e in entries:
                    if not isinstance(e, dict):
                        continue
                    fixed_entries.append({
                        "id": e.get("id") or _gid(),
                        "text": str(e.get("text", "")),
                        "copy_count": int(e.get("copy_count", 0)),
                        "pinned": bool(e.get("pinned", False)),
                    })
                if not fixed_entries:
                    fixed_entries = [_entry()]
                fixed_rows.append({
                    "id": row.get("id") or _gid(),
                    "color_idx": int(row.get("color_idx", i)),
                    "entries": fixed_entries[:4],
                    "pinned": bool(row.get("pinned", False)),
                    "label": str(row.get("label", "")),
                })
            state["rows"] = fixed_rows or [_row(0)]
            state["poll_rate"] = 250 if state["poll_rate"] <= 250 else 500 if state["poll_rate"] <= 750 else 1000
            if state["history_limit"] not in (5, 10, 20, 50):
                state["history_limit"] = 10
            state["history"] = [str(x) for x in d.get("history", [])][:state["history_limit"]]
            log.info("State loaded from %s (%d rows, theme=%s)",
                     DATA_FILE, len(state["rows"]), state["theme"])
            return
        except Exception as exc:
            _log_exc("load()", exc)
            # Auto-heal corrupted/empty JSON so startup does not repeatedly fail.
            try:
                bad = DATA_FILE.with_suffix(f".corrupt-{int(time.time())}.json")
                DATA_FILE.replace(bad)
                log.warning("Corrupted state moved to %s", bad)
            except Exception:
                pass
    state["rows"] = [_row(0)]
    state["rows"][0]["entries"][0]["text"] = "Welcome to Clippy!"
    state["auto_capture"] = True
    state["history_enabled"] = True
    state["theme"] = "daylight"
    state["hotkey"] = DEFAULT_HOTKEY
    state["poll_rate"] = 500
    state["paste_plain_text"] = True
    state["launch_at_startup"] = True
    state["backup_enabled"] = True
    state["history_limit"] = 10
    state["last_backup"] = ""
    save()

def save():
    try:
        DATA_FILE.write_text(json.dumps({
            "rows": state["rows"],
            "auto_capture": state.get("auto_capture", True),
            "history_enabled": state["history_enabled"],
            "theme": state.get("theme", "dark"),
            "hotkey": state.get("hotkey", DEFAULT_HOTKEY),
            "poll_rate": state.get("poll_rate", 500),
            "paste_plain_text": state.get("paste_plain_text", True),
            "launch_at_startup": state.get("launch_at_startup", True),
            "backup_enabled": state.get("backup_enabled", True),
            "history_limit": state.get("history_limit", 10),
            "last_backup": state.get("last_backup", ""),
            "history": state.get("history", []),
        }, ensure_ascii=False, indent=2), "utf-8")
    except Exception as exc:
        _log_exc("save()", exc)

def _run_backup(target_dir: Path | None = None, prune_old: bool = True) -> bool:
    try:
        out_dir = Path(target_dir) if target_dir else BACKUP_DIR
        out_dir.mkdir(parents=True, exist_ok=True)
        now = datetime.now()
        backup_file = out_dir / f"clippy_backup_{now.strftime('%Y-%m-%d')}.json"
        with _state_lock:
            payload = {
                "rows": state.get("rows", []),
                "created_at": now.isoformat(timespec="seconds"),
            }
        backup_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2), "utf-8")

        if prune_old:
            cutoff = now - timedelta(days=7)
            for path in out_dir.glob("clippy_backup_*.json"):
                try:
                    stamp = path.stem.replace("clippy_backup_", "")
                    file_date = datetime.strptime(stamp, "%Y-%m-%d")
                    if file_date < cutoff:
                        path.unlink(missing_ok=True)
                except Exception:
                    continue

        with _state_lock:
            state["last_backup"] = now.strftime("%Y-%m-%d %H:%M:%S")
        save()
        log.info("Backup saved: %s", backup_file)
        return True
    except Exception as exc:
        _log_exc("_run_backup()", exc)
        return False

def find_next_capture_slot(rows, ri, ei):
    """
    Determine where the NEXT auto-capture should land after filling slot (ri, ei).

    Rule:
      - A row is exhausted only when it has exactly 4 entries (no Add Variation shown).
      - As long as a row has fewer than 4 entries, the next capture stays in that
        same row — the poller will auto-add a variation slot when it fires.
      - Only when a row is full (4 entries) does capture advance to the next row.
      - User clicking a card overrides everything (handled in setCursorToEntry).
    """
    if ri < 0 or ri >= len(rows):
        return {"row_idx": len(rows), "entry_idx": 0}

    row = rows[ri]
    # Current row still has room — next capture goes here (poller adds the variation)
    if len(row["entries"]) < 4:
        return {"row_idx": ri, "entry_idx": len(row["entries"])}

    # Current row is full — find next row with room, or signal to create new row
    for r in range(ri + 1, len(rows)):
        if len(rows[r]["entries"]) < 4:
            # Find first empty slot in this row, or point to end so poller adds variation
            for i, e in enumerate(rows[r]["entries"]):
                if not e["text"]:
                    return {"row_idx": r, "entry_idx": i}
            return {"row_idx": r, "entry_idx": len(rows[r]["entries"])}

    # All rows full — signal to create a new row
    return {"row_idx": len(rows), "entry_idx": 0}


# ══════════════════════════════════════════════════════════════════════════════
# PASTE GUARD — poller checks this flag before capturing
# ══════════════════════════════════════════════════════════════════════════════
_clippy_is_pasting = False  # True while Clippy is writing+pasting — poller must skip
_clippy_last_pasted = ""   # The text Clippy just pasted — poller must skip this on unfreeze

# Global lock — ALL reads+writes of `state` must hold this lock.
# The clipboard poller runs on a QThread; Bridge slots run on the Qt main thread.
# Without a lock, concurrent access silently corrupts rows/cursor/history.
_last_clipboard_open_warning_ts = 0.0
_last_clipboard_busy_toast_ts = 0.0
_state_lock = threading.Lock()

# Global reference so Bridge slots can reach the window
_win_ref = [None]

# ══════════════════════════════════════════════════════════════════════════════
# BRIDGE
# ══════════════════════════════════════════════════════════════════════════════
class Bridge(QObject):
    stateChanged  = pyqtSignal(str)
    # FIX #1+2: signal to tell the window to paste then hide
    pasteAndHide  = pyqtSignal(str)

    def _push(self, extra=None):
        with _state_lock:
            payload = {
                "rows":       state["rows"],
                "cursor":     state["cursor"],
                "auto":       state["auto_capture"],
                "history":    state["history"],
                "hist_enabled": state["history_enabled"],
                "theme":      state.get("theme", "dark"),
                "hotkey":     state.get("hotkey", DEFAULT_HOTKEY),
                "poll_rate":  state.get("poll_rate", 500),
                "paste_plain_text": state.get("paste_plain_text", True),
                "launch_at_startup": state.get("launch_at_startup", True),
                "backup_enabled": state.get("backup_enabled", True),
                "history_limit": state.get("history_limit", 10),
                "last_backup": state.get("last_backup", ""),
            }
        if extra: payload.update(extra)
        self.stateChanged.emit(json.dumps(payload))

    @pyqtSlot(result=str)
    def getState(self):
        with _state_lock:
            return json.dumps({
                "rows":       state["rows"],
                "cursor":     state["cursor"],
                "auto":       state["auto_capture"],
                "history":    state["history"],
                "hist_enabled": state["history_enabled"],
                "clip_ok":    CLIPBOARD_OK,
                "theme":      state.get("theme", "dark"),
                "hotkey":     state.get("hotkey", DEFAULT_HOTKEY),
                "poll_rate":  state.get("poll_rate", 500),
                "paste_plain_text": state.get("paste_plain_text", True),
                "launch_at_startup": state.get("launch_at_startup", True),
                "backup_enabled": state.get("backup_enabled", True),
                "history_limit": state.get("history_limit", 10),
                "last_backup": state.get("last_backup", ""),
            })

    # ── Copy entry to clipboard ───────────────────────────────────────────────
    @pyqtSlot(str, str)
    def copyEntry(self, entry_id, text):
        if not text: return
        if CLIPBOARD_OK:
            try:
                pyperclip.copy(text)
            except Exception as exc:
                _log_exc("copyEntry/pyperclip.copy", exc)
                self.stateChanged.emit(json.dumps({"toast":"Copy failed — clipboard unavailable","toast_type":"warn"}))
                return
        with _state_lock:
            state["last_clip"] = text
            state["last_clip_time"] = time.time()
            for row in state["rows"]:
                for e in row["entries"]:
                    if e["id"] == entry_id:
                        e["copy_count"] = e.get("copy_count", 0) + 1
        save(); self._push()

    @pyqtSlot(str, str)
    def pasteEntry(self, entry_id, text):
        """Paste exact selected card text to destination — fully deterministic."""
        if not text: return
        with _state_lock:
            for row in state["rows"]:
                for e in row["entries"]:
                    if e["id"] == entry_id:
                        e["copy_count"] = e.get("copy_count", 0) + 1
        save()
        self.pasteAndHide.emit(text)

    @pyqtSlot(str, str)
    def saveEntry(self, entry_id, text):
        with _state_lock:
            for row in state["rows"]:
                for e in row["entries"]:
                    if e["id"] == entry_id:
                        e["text"] = text
        save(); self._push()

    @pyqtSlot(str, str)
    def deleteEntry(self, row_id, entry_id):
        with _state_lock:
            for row in state["rows"]:
                if row["id"] == row_id:
                    row["entries"] = [e for e in row["entries"] if e["id"] != entry_id]
            state["rows"] = [r for r in state["rows"] if r["entries"]]
            # Clamp cursor so it never points at a deleted/shifted row
            n = len(state["rows"])
            if n == 0:
                state["cursor"] = {"row_idx": 0, "entry_idx": 0}
            else:
                ri = max(0, min(state["cursor"]["row_idx"], n - 1))
                ei = max(0, min(state["cursor"]["entry_idx"],
                                len(state["rows"][ri]["entries"]) - 1))
                state["cursor"] = {"row_idx": ri, "entry_idx": ei}
        save(); self._push()

    @pyqtSlot(str)
    def pinEntry(self, entry_id):
        with _state_lock:
            for row in state["rows"]:
                for e in row["entries"]:
                    if e["id"] == entry_id:
                        e["pinned"] = not e.get("pinned", False)
        save(); self._push()

    @pyqtSlot(str)
    def addEntry(self, row_id):
        added = False
        with _state_lock:
            for ri, row in enumerate(state["rows"]):
                if row["id"] == row_id and len(row["entries"]) < 4:
                    new_e = _entry()
                    row["entries"].append(new_e)
                    state["cursor"] = {"row_idx": ri, "entry_idx": len(row["entries"]) - 1}
                    added = True
                    break
        # save() and _push() always called outside the lock — no disk I/O while holding it
        save(); self._push()

    @pyqtSlot(str, str)
    def setCursorToEntry(self, row_id, entry_id):
        with _state_lock:
            for ri, row in enumerate(state["rows"]):
                if row["id"] == row_id:
                    for ei, e in enumerate(row["entries"]):
                        if e["id"] == entry_id:
                            state["cursor"] = {"row_idx": ri, "entry_idx": ei}
                            self._push(); return

    @pyqtSlot(result=str)
    def addRow(self):
        with _state_lock:
            r = _row(len(state["rows"]))
            state["rows"].append(r)
            ri = len(state["rows"]) - 1
            state["cursor"] = {"row_idx": ri, "entry_idx": 0}
        save(); self._push()
        return json.dumps({"new_row_id": r["id"], "new_entry_id": r["entries"][0]["id"]})

    @pyqtSlot(str)
    def deleteRow(self, row_id):
        with _state_lock:
            state["rows"] = [r for r in state["rows"] if r["id"] != row_id]
            # Clamp cursor after row deletion
            n = len(state["rows"])
            if n == 0:
                state["cursor"] = {"row_idx": 0, "entry_idx": 0}
            else:
                ri = max(0, min(state["cursor"]["row_idx"], n - 1))
                state["cursor"] = {"row_idx": ri, "entry_idx": 0}
        save(); self._push()

    @pyqtSlot(str, str)
    def reorderRow(self, from_id, to_id):
        with _state_lock:
            rows = state["rows"]
            fi = next((i for i, r in enumerate(rows) if r["id"] == from_id), None)
            ti = next((i for i, r in enumerate(rows) if r["id"] == to_id), None)
            if fi is not None and ti is not None:
                rows.insert(ti, rows.pop(fi))
        save(); self._push()

    @pyqtSlot(str, str, str)
    def moveEntry(self, entry_id, target_row_id, target_entry_id):
        toast = None
        with _state_lock:
            rows = state["rows"]
            src_row = None
            src_idx = None
            target_row = None

            for row in rows:
                if src_row is None:
                    idx = next((i for i, e in enumerate(row["entries"]) if e["id"] == entry_id), None)
                    if idx is not None:
                        src_row = row
                        src_idx = idx
                if target_row is None and row["id"] == target_row_id:
                    target_row = row
                if src_row is not None and target_row is not None:
                    break

            if src_row is None or src_idx is None or target_row is None:
                return

            if src_row is target_row:
                entries = src_row["entries"]
                src = entries.pop(src_idx)
                ti = next((i for i, e in enumerate(entries) if e["id"] == target_entry_id), None)
                entries.insert(ti if ti is not None else len(entries), src)
            elif len(target_row["entries"]) >= 4:
                toast = "That group is full"
            else:
                src = src_row["entries"].pop(src_idx)
                ti = next((i for i, e in enumerate(target_row["entries"]) if e["id"] == target_entry_id), None)
                target_row["entries"].insert(ti if ti is not None else len(target_row["entries"]), src)

            state["rows"] = [r for r in state["rows"] if r["entries"]]
            n = len(state["rows"])
            if n == 0:
                state["cursor"] = {"row_idx": 0, "entry_idx": 0}
            else:
                ri = max(0, min(state["cursor"]["row_idx"], n - 1))
                ei = max(0, min(state["cursor"]["entry_idx"], len(state["rows"][ri]["entries"]) - 1))
                state["cursor"] = {"row_idx": ri, "entry_idx": ei}
        save(); self._push({"toast": toast, "toast_type": "warn"} if toast else None)

    @pyqtSlot(bool)
    def setAutoCapture(self, enabled):
        with _state_lock:
            state["auto_capture"] = enabled
            if not enabled:
                state["cursor"] = {"row_idx": -1, "entry_idx": -1}
        save(); self._push()

    @pyqtSlot(bool)
    def setHistoryEnabled(self, enabled):
        with _state_lock:
            state["history_enabled"] = enabled
        save(); self._push()

    @pyqtSlot()
    def clearHistory(self):
        with _state_lock:
            state["history"] = []
        save()
        self._push()

    @pyqtSlot(str)
    def setTheme(self, theme):
        """Persist theme selection. Only writes to disk when the value changed.
        Deliberately does NOT call _push() — that would re-send theme to JS,
        which would re-call bridge.setTheme(), creating an infinite save loop."""
        with _state_lock:
            if state.get("theme") == theme:
                return          # already set — nothing to do, no disk write
            state["theme"] = theme
        save()
        log.info("Theme changed to %r", theme)

    @pyqtSlot(str, result=str)
    def setHotkey(self, hotkey):
        combo = _normalize_hotkey(hotkey)
        if not _valid_hotkey(combo):
            return json.dumps({
                "ok": False,
                "error": "Use a valid combo like Ctrl+Alt+H",
            })

        win = _win_ref[0]
        if win is None:
            return json.dumps({"ok": False, "error": "Window is not ready yet"})
        ok, err = win.register_hotkey(combo)
        if not ok:
            return json.dumps({
                "ok": False,
                "error": f"Windows rejected that hotkey: {err}",
            })

        with _state_lock:
            state["hotkey"] = combo
        save()
        self._push({"toast": f"Hotkey set to {combo.upper()}"})
        return json.dumps({"ok": True, "hotkey": combo})

    @pyqtSlot(int)
    def setPollRate(self, ms):
        if ms not in (250, 500, 1000):
            return
        with _state_lock:
            state["poll_rate"] = ms
        save()
        self._push()

    @pyqtSlot(str)
    def pinRow(self, row_id):
        with _state_lock:
            for row in state["rows"]:
                if row["id"] == row_id:
                    row["pinned"] = not row.get("pinned", False)
                    break
        save()
        self._push()

    @pyqtSlot(bool)
    def setPastePlainText(self, enabled):
        with _state_lock:
            state["paste_plain_text"] = bool(enabled)
        save()
        self._push()

    @pyqtSlot(bool)
    def setLaunchAtStartup(self, enabled):
        enabled = bool(enabled)
        ok = _set_launch_at_startup(enabled)
        if not ok:
            self._push({"toast": "Startup setting failed", "toast_type": "warn"})
            return
        with _state_lock:
            state["launch_at_startup"] = enabled
        save()
        self._push()

    @pyqtSlot(bool)
    def setBackupEnabled(self, enabled):
        with _state_lock:
            state["backup_enabled"] = bool(enabled)
        save()
        self._push()

    @pyqtSlot(int)
    def setHistoryLimit(self, limit):
        if limit not in (5, 10, 20, 50):
            return
        with _state_lock:
            state["history_limit"] = limit
            state["history"] = state["history"][:limit]
        save()
        self._push()

    @pyqtSlot()
    def backupNow(self):
        target = QFileDialog.getExistingDirectory(None, "Choose backup folder", str(BACKUP_DIR))
        if not target:
            self._push({"toast": "Backup canceled", "toast_type": "warn"})
            return
        if _run_backup(Path(target), prune_old=False):
            self._push({"toast": "Backup saved"})
        else:
            self._push({"toast": "Backup failed — check clippy.log", "toast_type": "warn"})

    @pyqtSlot()
    def openBackupFolder(self):
        try:
            BACKUP_DIR.mkdir(parents=True, exist_ok=True)
            subprocess.Popen(["explorer", str(BACKUP_DIR)])
        except Exception as exc:
            _log_exc("openBackupFolder()", exc)
            self._push({"toast": "Could not open backup folder", "toast_type": "warn"})

    @pyqtSlot()
    def clearAllData(self):
        with _state_lock:
            state["rows"] = [_row(0)]
            state["history"] = []
            state["cursor"] = {"row_idx": 0, "entry_idx": 0}
        save()
        self._push({"toast": "All data cleared"})

    @pyqtSlot()
    def exportJSON(self):
        path, _ = QFileDialog.getSaveFileName(None, "Save JSON", "clippy-export.json", "JSON (*.json)")
        if path:
            try:
                with _state_lock:
                    data = json.dumps(state["rows"], ensure_ascii=False, indent=2)
                Path(path).write_text(data, "utf-8")
                log.info("exportJSON: wrote %s", path)
                self.stateChanged.emit(json.dumps({"toast": "Exported as JSON!"}))
            except Exception as exc:
                _log_exc("exportJSON", exc)
                self.stateChanged.emit(json.dumps({"toast": f"Export failed: {exc}", "toast_type": "warn"}))

    @pyqtSlot()
    def exportCSV(self):
        path, _ = QFileDialog.getSaveFileName(None, "Save CSV", "clippy-export.csv", "CSV (*.csv)")
        if path:
            try:
                buf = io.StringIO()
                w = csv.writer(buf)
                w.writerow(["row", "text", "copy_count", "pinned"])
                with _state_lock:
                    for ri, row in enumerate(state["rows"]):
                        for e in row["entries"]:
                            w.writerow([ri + 1, e.get("text", ""), e.get("copy_count", 0), e.get("pinned", False)])
                Path(path).write_text(buf.getvalue(), "utf-8")
                log.info("exportCSV: wrote %s", path)
                self.stateChanged.emit(json.dumps({"toast": "Exported as CSV!"}))
            except Exception as exc:
                _log_exc("exportCSV", exc)
                self.stateChanged.emit(json.dumps({"toast": f"Export failed: {exc}", "toast_type": "warn"}))

    @pyqtSlot()
    def hideWindow(self):
        """Called from JS when user clicks outside Clippy."""
        win = _win_ref[0]
        if win is None:
            return
        if getattr(win, '_pasting', False):
            log.debug("hideWindow() suppressed — paste in progress")
            return
        if getattr(win, '_showing', False):
            log.debug("hideWindow() suppressed — show animation in progress")
            return
        log.info("hideWindow() — hiding to tray")
        win.hide()

    @pyqtSlot()
    def importJSON(self):
        path, _ = QFileDialog.getOpenFileName(None, "Open JSON", "", "JSON (*.json)")
        if not path:
            return
        backup = None
        try:
            raw = json.loads(Path(path).read_text("utf-8"))

            # Accept both supported shapes:
            # 1) Export/import format: top-level array of row objects
            # 2) Backup format: {"rows":[...], "created_at":"..."}
            if isinstance(raw, dict) and isinstance(raw.get("rows"), list):
                data = raw["rows"]
            else:
                data = raw

            # Schema validation — must be a list of row dicts with required keys
            if not isinstance(data, list):
                raise ValueError("Top-level value must be a JSON array of groups or an object with a 'rows' array")
            for i, row in enumerate(data):
                if not isinstance(row, dict):
                    raise ValueError(f"Group {i} is not an object")
                if "entries" not in row or not isinstance(row["entries"], list):
                    raise ValueError(f"Group {i} missing 'entries' array")
                for j, e in enumerate(row["entries"]):
                    if not isinstance(e, dict) or "id" not in e:
                        raise ValueError(f"Group {i} entry {j} missing 'id'")

            # Backup existing rows before replacing
            with _state_lock:
                backup = state["rows"].copy()
                normalized = []
                for i, row in enumerate(data):
                    normalized.append({
                        "id": row.get("id") or _gid(),
                        "color_idx": int(row.get("color_idx", i)),
                        "pinned": bool(row.get("pinned", False)),
                        "label": str(row.get("label", "")),
                        "entries": [{
                            "id": e.get("id") or _gid(),
                            "text": str(e.get("text", "")),
                            "copy_count": int(e.get("copy_count", 0)),
                            "pinned": bool(e.get("pinned", False)),
                        } for e in row.get("entries", []) if isinstance(e, dict)] or [_entry()],
                    })
                state["rows"] = normalized

            save()
            log.info("importJSON: imported %d rows from %s", len(data), path)
            self._push()
            self.stateChanged.emit(json.dumps({**json.loads(self.getState()), "toast": "Imported!"}))

        except Exception as ex:
            _log_exc("importJSON", ex)
            # Restore backup if we already replaced state
            try:
                with _state_lock:
                    if backup is not None and state["rows"] is not backup:
                        state["rows"] = backup
            except Exception:
                pass
            self.stateChanged.emit(json.dumps({"toast": f"Import failed: {ex}", "toast_type": "warn"}))


# ══════════════════════════════════════════════════════════════════════════════
# CLIPBOARD POLLER — FIX #4: proper re-copy detection
# ══════════════════════════════════════════════════════════════════════════════
class ClipboardPoller(QThread):
    captured = pyqtSignal(str)

    def run(self):
        last_seen = ""
        MAX_ROWS = 1700  # Default auto-capture ceiling
        clipboard_busy_streak = 0
        def _sleep_tick():
            with _state_lock:
                ms = int(state.get("poll_rate", 500))
            time.sleep(max(0.1, ms / 1000.0))

        def _emit_state(extra=None):
            win = _win_ref[0]
            if win is None:
                return
            with _state_lock:
                payload = {
                    "rows": state["rows"],
                    "cursor": state["cursor"],
                    "auto": state["auto_capture"],
                    "history": state["history"],
                    "hist_enabled": state["history_enabled"],
                    "theme": state.get("theme", "dark"),
                    "hotkey": state.get("hotkey", DEFAULT_HOTKEY),
                    "poll_rate": state.get("poll_rate", 500),
                    "paste_plain_text": state.get("paste_plain_text", True),
                    "launch_at_startup": state.get("launch_at_startup", True),
                    "backup_enabled": state.get("backup_enabled", True),
                    "history_limit": state.get("history_limit", 10),
                    "last_backup": state.get("last_backup", ""),
                }
            if extra:
                payload.update(extra)
            try:
                win.bridge.stateChanged.emit(json.dumps(payload))
            except Exception as exc:
                _log_exc("ClipboardPoller._emit_state()", exc)

        while True:
            # ── Sleep is ALWAYS at the end of the loop body ──────────────────
            # This ensures the very first check fires immediately on thread
            # start (no missed copies at launch) and also fires immediately
            # when auto-capture is re-enabled after being turned off.

            # Read shared flags under lock — even boolean reads can tear on some platforms
            with _state_lock:
                auto_on = state["auto_capture"]
            if not auto_on or not CLIPBOARD_OK:
                _sleep_tick()
                continue

            global _clippy_is_pasting, _clippy_last_pasted
            global _last_clipboard_open_warning_ts, _last_clipboard_busy_toast_ts
            if _clippy_is_pasting:
                _sleep_tick()
                continue

            try:
                try:
                    text = pyperclip.paste()
                except Exception as exc:
                    msg = str(exc)
                    transient_clipboard_lock = (
                        CLIPBOARD_OK and
                        exc.__class__.__name__ == "PyperclipWindowsException" and
                        "OpenClipboard" in msg
                    )
                    if transient_clipboard_lock:
                        clipboard_busy_streak += 1
                        now = time.monotonic()
                        if now - _last_clipboard_open_warning_ts >= 15.0:
                            log.warning("Clipboard busy; retrying capture on next poll")
                            _last_clipboard_open_warning_ts = now
                        if clipboard_busy_streak >= 12:
                            with _state_lock:
                                state["auto_capture"] = False
                                state["cursor"] = {"row_idx": -1, "entry_idx": -1}
                            log.warning("Clipboard busy for an extended period; auto-capture paused")
                            if now - _last_clipboard_busy_toast_ts >= 30.0:
                                _emit_state({
                                    "toast": "Clipboard is busy - auto-capture paused",
                                    "toast_type": "warn",
                                })
                                _last_clipboard_busy_toast_ts = now
                            clipboard_busy_streak = 0
                        _sleep_tick()
                        continue
                    raise

                clipboard_busy_streak = 0

                if not text:
                    last_seen = ""
                    _sleep_tick()
                    continue

                if text == last_seen:
                    _sleep_tick()
                    continue

                if text == _clippy_last_pasted:
                    last_seen = text
                    _sleep_tick()
                    continue

                last_seen = text

                with _state_lock:
                    # History
                    if state["history_enabled"]:
                        lim = int(state.get("history_limit", 10))
                        h = [t for t in state["history"] if t != text]
                        h.insert(0, text)
                        state["history"] = h[:lim]

                    rows = state["rows"]
                    c    = state["cursor"]
                    ri   = c["row_idx"]
                    ei   = c["entry_idx"]

                    if ri < 0 or ri >= len(rows):
                        # FIX C3: only create new row if under cap
                        if len(rows) < MAX_ROWS:
                            rows.append({"id": _gid(), "color_idx": len(rows),
                                         "entries": [_entry(text)], "pinned": False, "label": ""})
                            ri = len(rows) - 1
                            ei = 0
                        else:
                            log.warning("Row cap (%d) reached — recycling oldest non-pinned row", MAX_ROWS)
                            # Recycle oldest row that has no pinned entries
                            for old_ri, old_row in enumerate(rows):
                                if not any(e.get("pinned") for e in old_row["entries"]):
                                    rows[old_ri] = {"id": _gid(), "color_idx": old_ri,
                                                    "entries": [_entry(text)], "pinned": False, "label": ""}
                                    ri = old_ri; ei = 0
                                    break
                            else:
                                _sleep_tick()
                                continue  # All rows pinned — skip capture
                    else:
                        entries = rows[ri]["entries"]

                        if ei < len(entries):
                            entries[ei] = {**entries[ei], "text": text}

                        elif len(entries) < 4:
                            entries.append(_entry(text))
                            ei = len(entries) - 1

                        else:
                            ri += 1
                            if ri >= len(rows):
                                if len(rows) < MAX_ROWS:
                                    rows.append({"id": _gid(), "color_idx": len(rows),
                                                 "entries": [_entry(text)], "pinned": False, "label": ""})
                                    ei = 0
                                else:
                                    log.warning("Row cap reached during overflow — skipping capture")
                                    _sleep_tick()
                                    continue
                            else:
                                ne = rows[ri]["entries"]
                                fi = next((i for i, e in enumerate(ne) if not e["text"]), None)
                                if fi is not None:
                                    ne[fi] = {**ne[fi], "text": text}; ei = fi
                                elif len(ne) < 4:
                                    ne.append(_entry(text)); ei = len(ne) - 1
                                else:
                                    ri += 1
                                    if len(rows) < MAX_ROWS:
                                        rows.append({"id": _gid(), "color_idx": len(rows),
                                                     "entries": [_entry(text)], "pinned": False, "label": ""})
                                        ei = 0
                                    else:
                                        log.warning("Row cap reached — skipping capture")
                                        _sleep_tick()
                                        continue

                    state["cursor"] = find_next_capture_slot(rows, ri, ei)

                save()
                self.captured.emit(text)

            except Exception as exc:
                _log_exc("ClipboardPoller.run()", exc)

            # Single canonical sleep — all non-skip paths reach here
            _sleep_tick()


# ══════════════════════════════════════════════════════════════════════════════
# LOGO — FIX #5: professional SVG-style icon drawn with QPainter
# ══════════════════════════════════════════════════════════════════════════════
def make_icon(size=64):
    pix = QPixmap(size,size)
    pix.fill(QColor(0,0,0,0))
    p = QPainter(pix)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    s = float(size)

    from PyQt6.QtCore import QRectF

    # Clipboard body (tilted yellow card with purple stroke)
    p.save()
    p.translate(s*0.66, s*0.58)
    p.rotate(16)
    clip_rect = QRectF(-s*0.24, -s*0.28, s*0.40, s*0.52)
    p.setBrush(QBrush(QColor("#f6c64b")))
    p.setPen(QPen(QColor("#5c57a8"), max(2, int(s*0.055)), Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
    p.drawRoundedRect(clip_rect, s*0.10, s*0.10)
    # Clip head
    p.setBrush(QBrush(QColor("#ffffff")))
    p.setPen(QPen(QColor("#5c57a8"), max(2, int(s*0.045)), Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
    p.drawRoundedRect(QRectF(-s*0.08, -s*0.35, s*0.22, s*0.12), s*0.05, s*0.05)
    p.restore()

    # Blue tape loop
    ring_rect = QRectF(s*0.08, s*0.18, s*0.56, s*0.56)
    p.setBrush(Qt.BrushStyle.NoBrush)
    p.setPen(QPen(QColor("#3f6ea3"), max(3, int(s*0.08)), Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
    p.drawEllipse(ring_rect)
    p.setPen(QPen(QColor("#e9ad2e"), max(2, int(s*0.05)), Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
    p.drawEllipse(QRectF(s*0.16, s*0.26, s*0.40, s*0.40))

    # Circular arrow inside loop
    p.setPen(QPen(QColor("#2f66a3"), max(2, int(s*0.06)), Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
    p.drawArc(QRectF(s*0.22, s*0.32, s*0.28, s*0.28), 40*16, 260*16)
    p.setBrush(QBrush(QColor("#2f66a3")))
    p.setPen(Qt.PenStyle.NoPen)
    p.drawPolygon(
        QPoint(int(s*0.50), int(s*0.49)),
        QPoint(int(s*0.60), int(s*0.48)),
        QPoint(int(s*0.53), int(s*0.57)),
    )

    # Accent sparkle near clipboard
    p.setPen(QPen(QColor("#5c57a8"), max(1, int(s*0.035)), Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
    p.drawLine(int(s*0.87), int(s*0.44), int(s*0.93), int(s*0.41))
    p.drawLine(int(s*0.89), int(s*0.50), int(s*0.96), int(s*0.50))
    p.drawLine(int(s*0.87), int(s*0.56), int(s*0.93), int(s*0.59))

    p.end()
    return QIcon(pix)



# ══════════════════════════════════════════════════════════════════════════════
# CARET POSITION DETECTION
# Gets the text-cursor (caret) position of the currently focused application
# using Windows Accessibility APIs (UI Automation / AccessibleObjectFromWindow).
# Must be called BEFORE showing Clippy so the destination app still owns focus.
# ══════════════════════════════════════════════════════════════════════════════
def _get_caret_position():
    """
    Return (x, y) of the text cursor in the foreground app, or None on failure.

    Strategy (most-reliable first):
      1. UI Automation (IUIAutomation) — works in modern apps, Office, browsers.
      2. AccessibleObjectFromWindow + OBJID_CARET — works in Notepad, legacy apps.
      3. GetCaretPos on the foreground thread — last resort for Win32 apps.

    All calls are wrapped so any failure falls through to the next method.
    """
    try:
        import ctypes
        import ctypes.wintypes as wt

        # ── Method 1: UI Automation ──────────────────────────────────────────
        try:
            ole32   = ctypes.windll.ole32
            uia_dll = ctypes.windll.UIAutomationCore

            IID_IUIAutomation = "{30CBE57D-D9D0-452A-AB13-7AC5AC4825EE}"
            CLSID_CUIAutomation = "{FF48DBA4-60EF-4201-AA87-54103EEF594E}"

            class POINT(ctypes.Structure):
                _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

            # CoCreateInstance to get IUIAutomation
            # Simpler: use the IAccessible caret approach first since UIA
            # requires COM apartment setup which varies. Fall through to method 2.
            pass
        except Exception:
            pass

        # ── Method 2: AccessibleObjectFromWindow + OBJID_CARET ───────────────
        try:
            user32  = ctypes.windll.user32
            oleacc  = ctypes.windll.oleacc

            OBJID_CARET = ctypes.c_long(-8)   # 0xFFFFFFF8

            class RECT(ctypes.Structure):
                _fields_ = [("left", wt.LONG), ("top", wt.LONG),
                             ("right", wt.LONG), ("bottom", wt.LONG)]

            hwnd_fg = user32.GetForegroundWindow()
            if not hwnd_fg:
                raise RuntimeError("No foreground window")

            # Walk to the focused child window (e.g. the edit control in Notepad)
            hwnd_focus = user32.GetFocus()
            if not hwnd_focus:
                tid_fg = user32.GetWindowThreadProcessId(hwnd_fg, None)
                tid_self = ctypes.windll.kernel32.GetCurrentThreadId()
                user32.AttachThreadInput(tid_self, tid_fg, True)
                hwnd_focus = user32.GetFocus()
                user32.AttachThreadInput(tid_self, tid_fg, False)
            if not hwnd_focus:
                hwnd_focus = hwnd_fg

            # IAccessible GUID
            IID_IAccessible = "{618736E0-3C3D-11CF-810C-00AA00389B71}"
            iid_bytes = (ctypes.c_byte * 16)(*[
                0xE0, 0x36, 0x87, 0x61, 0x3D, 0x3C, 0xCF, 0x11,
                0x81, 0x0C, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71
            ])

            pAcc = ctypes.c_void_p()
            pChild = ctypes.c_void_p()
            hr = oleacc.AccessibleObjectFromWindow(
                hwnd_focus, OBJID_CARET,
                ctypes.byref(iid_bytes),
                ctypes.byref(pAcc)
            )
            if hr == 0 and pAcc:
                rect = RECT()
                # IAccessible::accLocation(pxLeft, pyTop, pcxWidth, pcyHeight, varChild)
                # We use accLocation via vtable — easier with pywin32, but we
                # fall through to Method 3 rather than do raw vtable calls here.
                pass
        except Exception:
            pass

        # ── Method 3: GetCaretPos via thread attachment ───────────────────────
        try:
            user32  = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32

            class POINT(ctypes.Structure):
                _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

            hwnd_fg = user32.GetForegroundWindow()
            if not hwnd_fg:
                return None

            tid_fg   = user32.GetWindowThreadProcessId(hwnd_fg, None)
            tid_self = kernel32.GetCurrentThreadId()

            # Attach to the foreground thread so GetCaretPos sees its caret
            attached = user32.AttachThreadInput(tid_self, tid_fg, True)

            pt = POINT()
            ok = user32.GetCaretPos(ctypes.byref(pt))

            if attached:
                user32.AttachThreadInput(tid_self, tid_fg, False)

            if not ok:
                return None

            # GetCaretPos returns client coords — convert to screen coords
            user32.ClientToScreen(hwnd_fg, ctypes.byref(pt))

            # Sanity check — discard (0,0) which usually means "not found"
            if pt.x == 0 and pt.y == 0:
                return None

            log.debug("Caret detected at (%d, %d) via GetCaretPos", pt.x, pt.y)
            return (pt.x, pt.y)

        except Exception as exc:
            log.debug("GetCaretPos method failed: %s", exc)
            return None

    except Exception as exc:
        log.debug("_get_caret_position() outer failure: %s", exc)
        return None

# ══════════════════════════════════════════════════════════════════════════════
# MAIN WINDOW
# ══════════════════════════════════════════════════════════════════════════════
class ClippyWindow(QMainWindow):
    showRequested = pyqtSignal(int, int)   # carries caret/fallback (x, y)

    def __init__(self):
        super().__init__()
        load()
        # Use only the in-app header (hide native Windows title bar/chrome).
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint, True)
        self.setWindowTitle("Clippy")

        W, H = 640, 460
        self.resize(W, H)
        self.setMinimumSize(500, 380)

        self._icon = make_icon(64)
        self.setWindowIcon(self._icon)
        self._first_show  = True
        self._showing     = False   # True during the show animation — suppresses blur-hide
        self._pasting     = False
        self._last_show_request_ts = 0.0

        log.info("ClippyWindow.__init__ — W=%d H=%d", W, H)

        # Centre on screen for first launch
        screen = QApplication.primaryScreen().geometry()
        self.move((screen.width()-W)//2, (screen.height()-H)//2)

        # ── Step 1: Tray icon — set up IMMEDIATELY before anything else ────────
        # On cold boot QWebEngineView can take 30-60s to initialise (Chromium
        # cold-start, AV scanning, cold disk cache). By setting up the tray
        # FIRST the user sees the icon and balloon notification within ~1 second
        # of the process starting, not after Chromium finishes loading.
        self._tray = QSystemTrayIcon(self._icon, self)
        tray_menu  = QMenu()
        tray_menu.addAction("Show Clippy", lambda: self._show_window(-1, -1))
        tray_menu.addSeparator()
        tray_menu.addAction("Quit", QApplication.quit)
        self._tray.setContextMenu(tray_menu)
        self._tray.setToolTip("Clippy — Clipboard Manager")
        self._tray.activated.connect(
            lambda r: self._show_window(-1, -1)
            if r == QSystemTrayIcon.ActivationReason.DoubleClick else None
        )
        self._tray.show()
        # Show notification immediately — tray icon is already registered above
        self._tray.showMessage(
            "Clippy is running",
            "Clipboard manager active in background.\nPress Ctrl+D to open.",
            QSystemTrayIcon.MessageIcon.Information,
            4000,
        )
        log.info("Tray icon shown and notification fired immediately")

        # ── Step 2: Hotkey registration — also before web init ───────────────
        # Register Ctrl+D now so the user can invoke Clippy immediately.
        # If the web page isn't ready yet, _show_window handles that gracefully.
        self.showRequested.connect(self._show_window)
        if HAS_KEYBOARD:
            def _on_hotkey():
                # Called on keyboard hook thread — destination app still has focus.
                # AllowSetForegroundWindow MUST be called now, while another app
                # owns the foreground lock. Calling it after focus moves to Clippy
                # does nothing because Clippy already owns the lock at that point.
                try:
                    import ctypes
                    ctypes.windll.user32.AllowSetForegroundWindow(0xFFFFFFFF)
                except Exception:
                    pass
                pos = _get_caret_position()
                x, y = pos if pos else (-1, -1)
                self.showRequested.emit(x, y)
            kb_lib.add_hotkey("ctrl+d", _on_hotkey, suppress=True)
            log.info("Hotkey Ctrl+D registered")
        self._custom_hotkey_handle = None
        self._custom_hotkey_combo = ""
        if HAS_KEYBOARD:
            with _state_lock:
                saved_hotkey = state.get("hotkey", DEFAULT_HOTKEY)
            if saved_hotkey and saved_hotkey != DEFAULT_HOTKEY:
                ok, err = self.register_hotkey(saved_hotkey)
                if not ok:
                    with _state_lock:
                        state["hotkey"] = DEFAULT_HOTKEY
                    save()
                    log.warning("Saved custom hotkey rejected (%s). Reverted to %s.", err, DEFAULT_HOTKEY)

        # ── Step 3: WebEngine + Bridge — heavy init ───────────────────────────
        self.web = QWebEngineView()
        self.setCentralWidget(self.web)

        self.channel = QWebChannel()
        self.bridge  = Bridge()
        self.channel.registerObject("bridge", self.bridge)
        self.web.page().setWebChannel(self.channel)

        # Connect paste+hide signal
        self.bridge.pasteAndHide.connect(self._do_paste_and_hide)

        # ── Step 4: Clipboard poller ──────────────────────────────────────────
        self.poller = ClipboardPoller()
        self.poller.captured.connect(self._on_captured)
        self.poller.start()

        _win_ref[0] = self  # give Bridge slots access to this window
        self.web.setHtml(BUILD_UI(), QUrl("about:blank"))
        self._page_ready = False
        # On cold boot Chromium can take 8+ seconds to be ready — use a longer
        # initial timeout so the first hotkey press doesn't hit a blank page.
        QTimer.singleShot(8000, lambda: setattr(self, '_page_ready', True))

        # ── Layer 1: Keep-alive ping every 25s ───────────────────────────────
        self._keepalive = QTimer(self)
        self._keepalive.setInterval(25_000)
        self._keepalive.timeout.connect(self._ping_renderer)
        self._keepalive.start()

        # ── Layer 2: Watchdog every 10s ──────────────────────────────────────
        self._watchdog = QTimer(self)
        self._watchdog.setInterval(10_000)
        self._watchdog.timeout.connect(self._watchdog_check)
        self._watchdog.start()

        QTimer.singleShot(5000, self._run_backup_and_push)
        self._backup_timer = QTimer(self)
        self._backup_timer.setInterval(24 * 60 * 60 * 1000)
        self._backup_timer.timeout.connect(self._run_backup_and_push)
        self._backup_timer.start()

        QTimer.singleShot(0, self._set_native_rounded_corners)

    def _set_native_rounded_corners(self):
        """Request smooth native rounded corners on Windows 11."""
        try:
            import ctypes
            DWMWA_WINDOW_CORNER_PREFERENCE = 33
            DWMWCP_ROUND = 2
            value = ctypes.c_int(DWMWCP_ROUND)
            hwnd = int(self.winId())
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,
                DWMWA_WINDOW_CORNER_PREFERENCE,
                ctypes.byref(value),
                ctypes.sizeof(value),
            )
        except Exception:
            pass

    def _run_backup_and_push(self):
        with _state_lock:
            enabled = bool(state.get("backup_enabled", True))
        if not enabled:
            return
        if _run_backup():
            self.bridge._push()

    def _on_custom_hotkey(self):
        try:
            import ctypes
            ctypes.windll.user32.AllowSetForegroundWindow(0xFFFFFFFF)
        except Exception:
            pass
        pos = _get_caret_position()
        x, y = pos if pos else (-1, -1)
        self.showRequested.emit(x, y)

    def register_hotkey(self, combo: str):
        if not HAS_KEYBOARD:
            return False, "keyboard library missing"
        combo = _normalize_hotkey(combo)
        if combo == DEFAULT_HOTKEY:
            if self._custom_hotkey_handle is not None:
                try:
                    kb_lib.remove_hotkey(self._custom_hotkey_handle)
                except Exception:
                    pass
            self._custom_hotkey_handle = None
            self._custom_hotkey_combo = ""
            return True, ""
        if not _valid_hotkey(combo):
            return False, "invalid format"
        try:
            if self._custom_hotkey_handle is not None:
                kb_lib.remove_hotkey(self._custom_hotkey_handle)
            self._custom_hotkey_handle = kb_lib.add_hotkey(
                combo, self._on_custom_hotkey, suppress=True
            )
            self._custom_hotkey_combo = combo
            log.info("Custom hotkey registered: %s", combo)
            return True, ""
        except Exception as exc:
            _log_exc("register_hotkey()", exc)
            return False, str(exc)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._set_native_rounded_corners()

    def _show_window(self, cx=-1, cy=-1):
        """cx, cy: caret screen coords from the destination app (-1,-1 = unknown)."""
        now = time.monotonic()
        if now - getattr(self, "_last_show_request_ts", 0.0) < 0.45:
            log.debug("_show_window() throttled")
            return
        self._last_show_request_ts = now
        if self.isVisible() and self.isActiveWindow() and not self._pasting:
            log.debug("_show_window() ignored (already foreground)")
            return
        log.debug("_show_window() called — caret=(%d,%d)", cx, cy)
        self._pending_pos = (cx, cy)   # stash so _do_show can read it
        # ── Layer 3: reload if blank before showing ───────────────────────────
        if getattr(self, '_page_ready', False):
            self.web.page().runJavaScript(
                "document.getElementById('main') ? 'ok' : 'blank'",
                lambda r: self._show_after_check(r)
            )
            return
        self._show_after_check('ok')

    def _show_after_check(self, health):
        log.debug("_show_after_check(%r)", health)
        if health != 'ok':
            log.warning("Page blank — triggering UI reload before show")
            self._reload_ui()
            QTimer.singleShot(1200, self._do_show)
            return
        self._do_show()

    def _do_show(self):
        log.info("_do_show() — pasting=%s", self._pasting)
        self._showing = True

        # ── Position Clippy at the caret, or fall back to mouse cursor ────────
        # Priority:
        #   1. Caret coords captured when hotkey fired (destination app in focus)
        #   2. Current mouse cursor position
        #   3. Screen centre (should never reach here)
        cx, cy = getattr(self, '_pending_pos', (-1, -1))
        screen  = QApplication.primaryScreen().geometry()
        W, H    = self.width(), self.height()

        if cx >= 0 and cy >= 0:
            # Place Clippy just below and to the right of the caret so it does
            # not obscure the line the user is editing. Apply a small offset so
            # the caret itself is still visible above the Clippy window.
            tx = cx + 4
            ty = cy + 22   # approximate line-height so caret stays visible
            log.debug("Positioning at caret (%d,%d) → window (%d,%d)", cx, cy, tx, ty)
        else:
            # Fallback: near mouse cursor
            mp = QCursor.pos()
            tx = mp.x() + 15
            ty = mp.y() + 15
            log.debug("No caret — positioning near mouse (%d,%d)", tx, ty)

        # Clamp to screen so Clippy never goes off-edge
        tx = max(0, min(tx, screen.width()  - W - 10))
        ty = max(0, min(ty, screen.height() - H - 10))
        self.move(tx, ty)

        # Momentarily set WindowStaysOnTopHint so the window punches through
        # the foreground lock — it is removed after 400ms so Clippy doesn't
        # permanently float over everything once the user is done with it.
        self.setWindowFlags(
            self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint
        )
        self.show()
        self.setWindowState(self.windowState() & ~Qt.WindowState.WindowMinimized)
        self.raise_()
        self.activateWindow()

        # _force_foreground needs the window to be fully painted before
        # SetForegroundWindow is called — give it one event-loop tick first.
        QTimer.singleShot(0, self._force_foreground)

        # Drop the always-on-top flag after the window has settled.
        # setWindowFlags() hides the window — re-show after dropping the flag.
        def _drop_topmost():
            self.setWindowFlags(
                self.windowFlags() & ~Qt.WindowType.WindowStaysOnTopHint
            )
            self.show()
            self.raise_()
        QTimer.singleShot(400, _drop_topmost)

        # Focus search bar immediately and reset keyboard nav state.
        QTimer.singleShot(120, lambda: self.web.page().runJavaScript("""
            window.__clippyLastShowTs = Date.now();
            kbIdx = 0;
            topNavIdx = -1;
            _topNavItems().forEach(el => el.classList.remove('nav-focus'));
            renderAll();
            var s = document.getElementById('search');
            if (s) { s.blur(); }
        """))

        # Release the showing guard after 1.5 s — slightly longer than before
        # to cover the 400ms topmost drop + any foreground retry delays.
        QTimer.singleShot(1500, self._clear_showing)

    def _clear_showing(self):
        """Release the showing guard — blur events are allowed again."""
        self._showing = False
        log.debug("_showing guard released")

    def _force_foreground(self):
        """Force Clippy to the foreground reliably — including on cold boot.

        Windows enforces a 'foreground lock timeout' (default ~200 seconds after
        last user input) that silently blocks SetForegroundWindow() for any
        process not granted explicit permission.  During the first few minutes
        after login Clippy has no input history so this blocks cold-boot shows.

        Three-layer strategy:
          1. AllowSetForegroundWindow(ASFW_ANY) — already called in the hotkey
             lambda while the destination app still owned the foreground.  We
             repeat it here as belt-and-suspenders.
          2. keybd_event trick — synthesising a keystroke briefly gives Clippy
             'last input' status which lifts the foreground lock for one call.
          3. Retry loop — attempt SetForegroundWindow up to 4 times with 150ms
             gaps.  The first attempt may still fail if Explorer's lock hasn't
             transferred yet; subsequent attempts succeed once it has.
        """
        try:
            import ctypes, ctypes.wintypes
            user32   = ctypes.windll.user32
            ASFW_ANY = 0xFFFFFFFF
            hwnd     = int(self.winId())

            # Layer 1: grant permission (belt-and-suspenders)
            user32.AllowSetForegroundWindow(ASFW_ANY)

            # Layer 2: keybd_event trick — inject a harmless key press/release
            # (VK_MENU = Alt key; KEYEVENTF_KEYUP = 0x0002)
            # This gives the process 'last input' status which lifts the lock.
            KEYEVENTF_KEYUP = 0x0002
            VK_MENU         = 0x12   # Alt
            user32.keybd_event(VK_MENU, 0, 0,             0)
            user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)

            # Layer 3: SW_RESTORE + retry loop
            user32.ShowWindow(hwnd, 9)   # SW_RESTORE = 9

            def _attempt_foreground(attempts_left):
                result = user32.SetForegroundWindow(hwnd)
                if result:
                    user32.BringWindowToTop(hwnd)
                    user32.SetFocus(hwnd)
                    log.debug("_force_foreground() succeeded — hwnd=%d attempts_left=%d",
                              hwnd, attempts_left)
                elif attempts_left > 0:
                    log.debug("SetForegroundWindow failed — retrying in 150ms")
                    QTimer.singleShot(150, lambda: _attempt_foreground(attempts_left - 1))
                else:
                    log.warning("_force_foreground() exhausted retries — hwnd=%d", hwnd)

            _attempt_foreground(3)   # up to 4 total attempts (1 + 3 retries)

        except Exception as exc:
            log.warning("_force_foreground() failed (non-Windows?): %s", exc)

    def _do_paste_and_hide(self, text):
        """
        Deterministic paste — `text` is exactly what the user selected.
        The poller is completely frozen for the entire operation so nothing
        gets recaptured. Clipboard is written immediately before ctrl+v so
        there is zero gap between the write and the keystroke.
        """
        global _clippy_is_pasting
        with _state_lock:
            plain_only = bool(state.get("paste_plain_text", True))
        if plain_only:
            text = str(text or "")
        log.info("_do_paste_and_hide() — text=%r", text[:60] if text else "")

        # Freeze the poller — skips every tick until we release it
        _clippy_is_pasting = True
        self._pasting = True

        # Hide so destination app gets focus back
        self.hide()

        if HAS_KEYBOARD and CLIPBOARD_OK:
            def _paste():
                global _clippy_is_pasting, _clippy_last_pasted
                try:
                    time.sleep(0.4)
                    _clippy_last_pasted = text
                    pyperclip.copy(text)
                    time.sleep(0.05)
                    kb_lib.send("ctrl+v")
                    log.debug("Ctrl+V sent for text=%r", text[:40])
                    time.sleep(1.5)
                except Exception as exc:
                    _log_exc("_do_paste_and_hide._paste()", exc)
                finally:
                    _clippy_is_pasting = False
                    self._pasting = False
                    log.debug("Paste complete — poller unfrozen")
            threading.Thread(target=_paste, daemon=True).start()
        else:
            log.warning("Paste skipped — HAS_KEYBOARD=%s CLIPBOARD_OK=%s",
                        HAS_KEYBOARD, CLIPBOARD_OK)
            _clippy_is_pasting = False
            self._pasting = False

    # ── Layer 1: keep-alive — prevent renderer suspension ───────────────────
    def _ping_renderer(self):
        """Run a no-op JS call so Chromium keeps the renderer process alive."""
        try:
            self.web.page().runJavaScript("1;")
        except Exception as exc:
            _log_exc("_ping_renderer()", exc)

    # ── Layer 2: watchdog — detect and recover blank page ────────────────────
    def _watchdog_check(self):
        """Check every 10s if the page is alive; reload if blank."""
        if not getattr(self, '_page_ready', False):
            return
        try:
            self.web.page().runJavaScript(
                "document.getElementById('main') ? 'ok' : 'blank'",
                lambda r: self._recover_if_blank(r)
            )
        except Exception as exc:
            _log_exc("_watchdog_check()", exc)

    def _recover_if_blank(self, result):
        if result != 'ok':
            log.warning("Watchdog: page blank — reloading UI")
            self._reload_ui()

    def _reload_ui(self):
        """Full UI reload — restores page from blank state."""
        log.info("_reload_ui() triggered")
        self._page_ready = False
        self.web.setHtml(BUILD_UI(), QUrl("about:blank"))
        QTimer.singleShot(4000, lambda: setattr(self, '_page_ready', True))

    def _on_captured(self, text):
        log.debug("_on_captured() — text=%r", text[:60] if text else "")
        payload = json.dumps({
            "rows":    state["rows"],
            "cursor":  state["cursor"],
            "auto":    state["auto_capture"],
            "history": state["history"],
            "hist_enabled": state["history_enabled"],
            "poll_rate": state.get("poll_rate", 500),
            "paste_plain_text": state.get("paste_plain_text", True),
            "launch_at_startup": state.get("launch_at_startup", True),
            "backup_enabled": state.get("backup_enabled", True),
            "history_limit": state.get("history_limit", 10),
            "last_backup": state.get("last_backup", ""),
            "hotkey": state.get("hotkey", DEFAULT_HOTKEY),
            "toast":   "📋 Captured",
        })
        self.web.page().runJavaScript(f"window.__onState({json.dumps(payload)})")

    # changeEvent: no longer auto-hides on focus loss.
    # Hiding is now controlled exclusively by:
    #   1. JS bridge.hideWindow() when user clicks outside
    #   2. _do_paste_and_hide() when user pastes
    #   3. closeEvent (X button)
    def changeEvent(self, event):
        super().changeEvent(event)

    # Close button → hide to tray
    def closeEvent(self, event):
        log.info("closeEvent() — hiding to tray")
        event.ignore()
        self.hide()


# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════
def BUILD_UI():
    return r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<title>Clippy</title>
<script src="qrc:///qtwebchannel/qwebchannel.js"></script>
<script>
// ── Early stub for window.__onState ─────────────────────────────────────────
// The clipboard poller can fire _on_captured() before the QWebChannel
// handshake completes and window.__onState is defined by the main script.
// Without this stub those calls throw "window.__onState is not a function".
// We queue them here and replay them once the bridge is ready.
window.__earlyStateQueue = [];
window.__onState = function(payload) {
  window.__earlyStateQueue.push(payload);
};
</script>
<style>
/* System font stack — no network request, instant render, works fully offline.
   Body: Segoe UI (Win10+), SF Pro Text (macOS/iOS), then best available sans.
   Mono: Cascadia Code (Win Terminal), Consolas (all Windows), then system mono. */
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
html,body{height:100%;font-family:'Poppins','Segoe UI','SF Pro Text',system-ui,-apple-system,BlinkMacSystemFont,sans-serif;margin:0;padding:0;}
html{border-radius:18px;overflow:hidden;background:transparent;}
body{background:var(--appBg);color:var(--text);transition:background .3s,color .3s;
  overflow:hidden;position:relative;border-radius:20px;font-size:100%;}
::-webkit-scrollbar{width:4px;}
::-webkit-scrollbar-thumb{background:var(--accentGlow,rgba(124,58,237,.35));border-radius:10px;}
::placeholder{color:var(--textMuted);}
@keyframes pulse{0%,100%{opacity:1;transform:scale(1);}50%{opacity:.4;transform:scale(.6);}}
@keyframes slideIn{from{opacity:0;transform:translateX(20px);}to{opacity:1;transform:translateX(0);}}
@keyframes fadeUp{from{opacity:0;transform:translateX(-50%) translateY(8px);}to{opacity:1;transform:translateX(-50%) translateY(0);}}

/* ── NIGHT (Editorial Slate) ── */
body.dark,body.night{
  --appBg:linear-gradient(145deg,#090e1a 0%,#0e1426 100%);
  --navBg:rgba(26,37,68,.96);--navBorder:rgba(49,70,111,.9);
  --panelBg:#0f1830;--panelBorder:rgba(49,70,111,.7);
  --rowBg:rgba(15,24,47,.95);--rowBorder:rgba(47,70,130,.8);
  --cardBg:rgba(22,33,63,.92);--cardBorder:rgba(48,70,112,.9);
  --cardActiveBg:rgba(26,38,70,.96);--cardActiveBorder:#7e6bff;
  --text:#eff3ff;--textDim:#94a7d8;--textMuted:#5a6e9a;
  --inputBg:rgba(20,31,58,.98);--inputBorder:rgba(43,61,100,.9);
  --sectionLabel:#94a7d8;--emptyText:#3a4f78;
  --toastBg:rgba(15,18,48,.98);--toastBorder:rgba(110,116,255,.5);
  --toastWarnBg:rgba(80,15,0,.95);--toastWarnBorder:rgba(251,113,133,.4);
  --gearStroke:#b5c4e9;
  --rowNumColor:#94a7d8;
  --accentPrimary:#6e74ff;--accentGlow:rgba(110,116,255,.22);
  --focusGradStart:#8a77ff;--focusGradEnd:#44d6e6;
}
body.dark,body.night{font-family:'Poppins','Segoe UI',system-ui,sans-serif;}
/* ── DAYLIGHT (Aurora Glass) ── */
body.daylight{
  --appBg:linear-gradient(135deg,#f8fbff 0%,#eef3ff 55%,#f7f2ff 100%);
  --navBg:rgba(255,255,255,.82);--navBorder:rgba(208,217,242,.85);
  --panelBg:#f8fbff;--panelBorder:rgba(196,210,240,.94);
  --rowBg:rgba(255,255,255,.56);--rowBorder:rgba(217,211,255,.9);
  --cardBg:rgba(255,255,255,.90);--cardBorder:rgba(214,224,246,.95);
  --cardActiveBg:#f4ecff;--cardActiveBorder:#a886ff;
  --text:#2e2c55;--textDim:#66749a;--textMuted:#98a2bd;
  --inputBg:#f3f6ff;--inputBorder:rgba(198,212,241,.96);
  --sectionLabel:#8390ac;--emptyText:#9ca9c8;
  --toastBg:rgba(255,255,255,.95);--toastBorder:rgba(123,87,255,.28);
  --toastWarnBg:rgba(255,244,230,.97);--toastWarnBorder:rgba(251,113,133,.4);
  --gearStroke:#636d95;
  --rowNumColor:#857dce;
  --accentPrimary:#7a5cff;--accentGlow:rgba(122,92,255,.18);
  --focusGradStart:#7c5cff;--focusGradEnd:#22c7d8;
}
body.daylight{font-family:'Segoe UI','SF Pro Text',system-ui,sans-serif;}
/* Aurora blobs — daylight */
body.daylight::before,body.daylight::after{
  content:"";position:fixed;border-radius:999px;pointer-events:none;z-index:0;filter:blur(58px);
}
body.daylight::before{
  width:440px;height:260px;left:-200px;top:-130px;background:rgba(130,107,255,.33);
}
body.daylight::after{
  width:500px;height:300px;right:-200px;bottom:-150px;background:rgba(42,198,216,.28);
}
/* Third pink aurora blob injected via JS — see auroraBlob div */
#aurora-blob{
  position:fixed;top:-105px;right:-150px;width:340px;height:210px;
  border-radius:999px;pointer-events:none;z-index:0;filter:blur(58px);
  background:rgba(245,140,189,.26);
  display:none;
}
body.daylight #aurora-blob{display:block;}
/* Halo blobs — night */
body.night::before,body.night::after,body.dark::before,body.dark::after{
  content:"";position:fixed;border-radius:999px;pointer-events:none;z-index:0;
}
body.night::before,body.dark::before{
  width:640px;height:360px;left:-260px;top:-170px;
  background:radial-gradient(ellipse at center, rgba(118,100,255,.46), rgba(118,100,255,0));
}
body.night::after,body.dark::after{
  width:760px;height:400px;right:-340px;bottom:-200px;
  background:radial-gradient(ellipse at center, rgba(0,201,214,.34), rgba(0,201,214,0));
}

/* ── NAV ── */
#nav{
  display:flex;align-items:center;padding:0 14px;height:64px;gap:10px;
  background:var(--navBg);border-bottom:1px solid var(--navBorder);
  position:relative;z-index:200;backdrop-filter:blur(18px);flex-shrink:0;
  box-shadow:0 8px 28px rgba(72,89,130,.15);
  border-radius:18px;
  margin:8px 8px 0 8px;
}

/* Logo */
.logo{display:flex;align-items:center;gap:0;flex-shrink:0;text-decoration:none;}
.logo-mark{
  width:34px;height:34px;border-radius:10px;flex-shrink:0;
  background:linear-gradient(140deg,#fef2e1 0%,#ffe2b8 55%,#fff1e3 100%);
  display:flex;align-items:center;justify-content:center;
  box-shadow:0 4px 12px rgba(86,108,155,.2);
}
.logo-mark svg{display:block;width:18px;height:18px;}
.logo-name{
  font-weight:700;font-size:30px;letter-spacing:-.3px;
  color:var(--accentPrimary);
  line-height:1;
}
body.daylight .logo-name{color:#5342a8;}
body.dark .logo-name,body.night .logo-name{color:#eaf0ff;}

.search-wrap{flex:1;position:relative;}
.search-wrap svg{position:absolute;left:10px;top:50%;transform:translateY(-50%);pointer-events:none;color:var(--textMuted);}
#search{
  width:100%;padding:9px 30px 9px 28px;
  background:var(--inputBg);border:1px solid var(--inputBorder);border-radius:100px;
  color:var(--text);font-size:16px;outline:none;font-family:inherit;transition:border-color .06s,box-shadow .06s;
}
#search:focus{border-color:var(--accentPrimary,rgba(123,87,255,.45));box-shadow:0 0 0 3px var(--accentGlow,rgba(123,87,255,.12));}
.search-clear{position:absolute;right:9px;top:50%;transform:translateY(-50%);
  background:none;border:none;color:var(--textMuted);cursor:pointer;font-size:18px;line-height:1;}

.btn-new{
  flex-shrink:0;background:var(--accentPrimary,#6d4bff);
  border:none;border-radius:11px;color:#fff;padding:8px 14px;
  font-size:14px;font-weight:700;cursor:pointer;font-family:inherit;
  box-shadow:0 8px 18px var(--accentGlow,rgba(109,75,255,.28));display:flex;align-items:center;gap:4px;transition:opacity .08s,transform .08s;
}
.btn-new:hover{opacity:.9;transform:translateY(-1px);}

.btn-gear{
  flex-shrink:0;width:35px;height:35px;border-radius:11px;cursor:pointer;
  display:flex;align-items:center;justify-content:center;
  background:var(--inputBg);border:1px solid var(--inputBorder);transition:all .06s;
}
.btn-gear svg{stroke:var(--gearStroke);stroke-width:2.3;transition:stroke .06s;}
.btn-gear.active,.btn-gear:hover{background:var(--accentGlow,rgba(123,87,255,.16));border-color:var(--accentPrimary,rgba(123,87,255,.42));box-shadow:0 0 0 3px var(--accentGlow,rgba(123,87,255,.14));}
.btn-new.nav-focus,.btn-gear.nav-focus,#search.nav-focus{
  box-shadow:0 0 0 3px var(--accentGlow,rgba(123,87,255,.22))!important;
  border-color:var(--accentPrimary,rgba(123,87,255,.46))!important;
}

/* ── MAIN ── */
/* scroll-wrap fills remaining height below nav — rows scroll inside here */
#scroll-wrap{
  height:calc(100vh - 72px);  /* account for nav top margin */
  overflow-y:auto;
  overflow-x:hidden;
  overscroll-behavior:contain;
  border-bottom-left-radius:20px;border-bottom-right-radius:20px;
}
#scroll-wrap::-webkit-scrollbar{width:4px;}
#scroll-wrap::-webkit-scrollbar-thumb{background:rgba(124,58,237,.35);border-radius:10px;}
#main{padding:10px;display:flex;flex-direction:column;gap:8px;position:relative;z-index:1;}
.empty{text-align:center;color:var(--emptyText);padding-top:60px;font-size:18px;}

/* ── ROW ── */
.row-wrap{
  display:flex;align-items:stretch;gap:6px;transition:opacity .08s;
  height:calc((100vh - 64px - 20px - 16px) / 3);
  min-height:126px;
  width:100%;
  min-width:0;
}
.row-wrap.dragging{opacity:.35;}
.row-num{
  color:var(--rowNumColor,var(--textDim));font-size:22px;font-weight:700;
  width:18px;text-align:right;flex-shrink:0;user-select:none;line-height:1;
  text-shadow:none;
}
.row-num.active{color:var(--text);}
.row-group{
  flex:1;position:relative;background:var(--rowBg);border:1px solid var(--rowBorder);
  border-radius:18px;padding:8px 10px 8px 10px;
  display:flex;flex-direction:row;flex-wrap:nowrap;gap:8px;align-items:stretch;
  max-width:100%;
  min-width:0;
  overflow:visible;
  transition:border-color .06s,background .06s,box-shadow .06s;
  backdrop-filter:blur(12px);
  box-shadow:0 4px 18px rgba(0,0,0,.12);
}
.row-group.active-capture{background:rgba(123,87,255,.08);border-color:rgba(123,87,255,.36);}
.row-group.search-match{border-color:rgba(168,85,247,.45);}
.row-btn-strip{
  display:flex;flex-direction:column;justify-content:center;align-items:center;gap:6px;
  flex-shrink:0;min-width:0;
  width:0;overflow:hidden;
  opacity:0;pointer-events:none;
  transition:width .15s ease,opacity .15s ease,min-width .15s ease;
}
.row-group > .row-del,.row-group > .row-pin{display:none;}
.row-del{
  width:24px;height:24px;border-radius:50%;flex-shrink:0;padding:0;
  background:var(--inputBg);border:1px solid var(--inputBorder);
  color:var(--textDim);cursor:pointer;display:flex;align-items:center;justify-content:center;
  font-size:13px;font-weight:700;line-height:1;
  transition:background .05s,border-color .05s,color .05s;
  box-shadow:0 2px 8px rgba(109,122,156,.18);
}
.row-del:hover{background:rgba(239,68,68,.18);border-color:rgba(239,68,68,.4);color:#c33232;}
.row-pin{
  width:24px;height:24px;border-radius:50%;flex-shrink:0;padding:0;
  background:var(--inputBg);border:1px solid var(--inputBorder);
  color:var(--textDim);cursor:pointer;display:flex;align-items:center;justify-content:center;
  font-size:11px;line-height:1;
  transition:background .05s,border-color .05s,color .05s;
  box-shadow:0 2px 8px rgba(109,122,156,.18);
}
.row-pin.active{background:rgba(109,40,217,.2);border-color:rgba(139,92,246,.55);color:var(--accentPrimary);}
.row-pin:hover{border-color:rgba(139,92,246,.55);color:var(--text);}

/* ── CARD ── */
.card{
  flex:1 1 0;min-width:0;max-width:none;aspect-ratio:auto;height:100%;position:relative;
  background:var(--cardBg);border:1px solid var(--cardBorder);border-radius:14px;
  padding:9px 10px 34px;display:flex;flex-direction:column;justify-content:space-between;
  overflow:hidden;transition:all .05s linear;cursor:pointer;
}
.card:hover{
  border-color:var(--accentPrimary,rgba(123,87,255,.55));background:rgba(123,87,255,.08);
  box-shadow:0 0 0 1px var(--accentGlow,rgba(123,87,255,.18)),0 8px 20px var(--accentGlow,rgba(123,87,255,.12));
}
.card.editing{border-color:rgba(139,92,246,.8)!important;background:rgba(109,40,217,.12)!important;box-shadow:0 0 0 3px rgba(109,40,217,.14)!important;}
.card.kb-focus{
  border-color:var(--focusGradStart,#a855f7)!important;background:rgba(168,85,247,.10)!important;
  box-shadow:0 0 0 2px rgba(168,85,247,.3),0 6px 18px rgba(109,40,217,.16)!important;
}
.card.search-hit{border-color:rgba(168,85,247,.6)!important;background:rgba(109,40,217,.08)!important;}
.card.active-slot{border-color:var(--cardActiveBorder,rgba(139,92,246,.5))!important;background:var(--cardActiveBg,rgba(109,40,217,.08))!important;box-shadow:0 0 0 2px var(--accentGlow,rgba(139,92,246,.2))!important;}
.card.user-target{border-color:#fbbf24!important;background:rgba(251,191,36,.05)!important;box-shadow:0 0 0 2px rgba(251,191,36,.28)!important;}
.card.dragging{opacity:.35;}

.card-dot{position:absolute;top:6px;right:6px;width:5px;height:5px;border-radius:50%;background:var(--accentPrimary,#a855f7);box-shadow:0 0 5px var(--accentPrimary,#a855f7);animation:pulse 1.4s ease-in-out infinite;}
.card-kb-bar{position:absolute;left:0;top:10%;bottom:10%;width:3px;border-radius:0 3px 3px 0;background:linear-gradient(180deg,var(--focusGradStart,#6d28d9),var(--focusGradEnd,#a855f7));box-shadow:0 0 8px rgba(168,85,247,.5);}
.card-pin{position:absolute;top:4px;right:5px;font-size:8px;color:#fbbf24;}
.heatbar{position:absolute;bottom:0;left:0;height:2px;border-radius:0 2px 0 0;opacity:.8;transition:width .4s;}
.card-text{flex:1;overflow:hidden;font-size:15px;line-height:1.42;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;}
.card-text.empty-slot{color:var(--textMuted);font-style:italic;font-size:15px;}
.card-text.code-font{font-family:'Cascadia Code','Cascadia Mono',Consolas,'Courier New',monospace;font-size:15px;}
.card-actions{
  display:flex;gap:5px;justify-content:flex-end;flex-shrink:0;
  position:absolute;right:6px;bottom:6px;z-index:6;
  background:var(--inputBg);border:1px solid var(--inputBorder);border-radius:8px;padding:2px;
  opacity:0;visibility:hidden;transform:translateY(3px);transition:opacity .03s linear,transform .03s linear;pointer-events:none;
}
.card:hover .card-actions,.card.editing .card-actions{
  opacity:1;visibility:visible;transform:translateY(0);pointer-events:auto;
}
.card.hover-actions .card-actions{
  opacity:1;visibility:visible;transform:translateY(0);pointer-events:auto;
}
.card-textarea{width:100%;height:100%;background:transparent;border:none;color:var(--text);font-size:17px;line-height:1.42;padding:0;outline:none;resize:none;font-family:'Cascadia Code','Cascadia Mono',Consolas,'Courier New',monospace;}
mark{background:rgba(168,85,247,.35);color:#fff;border-radius:3px;padding:0 2px;}

/* ── CARD BUTTONS ── */
.cbtn{
  width:22px;height:22px;border-radius:6px;padding:0;flex-shrink:0;
  cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:13px;
  transition:all .04s linear;background:var(--cardBg);border:1px solid var(--cardBorder);color:var(--textDim);
}
.cbtn:hover{background:var(--inputBg);border-color:rgba(123,87,255,.5);color:var(--text);}
.cbtn.active{background:rgba(109,40,217,.22);border-color:rgba(139,92,246,.5);color:var(--accentPrimary,#a78bfa);}
.cbtn.danger{border-color:rgba(239,68,68,.18);color:#ef4444;}
.cbtn.danger:hover{background:rgba(239,68,68,.2);border-color:rgba(239,68,68,.42);color:#f87171;}
.save-btn{background:var(--accentGlow,rgba(109,40,217,.28));border:1px solid var(--accentPrimary,rgba(139,92,246,.5));border-radius:5px;color:var(--accentPrimary,#a78bfa);font-size:14px;padding:2px 10px;cursor:pointer;font-family:inherit;font-weight:600;}

/* ── ADD VARIATION ── */
.card-add{
  flex:1 1 0;min-width:0;max-width:none;aspect-ratio:auto;height:100%;cursor:pointer;
  background:rgba(123,87,255,.03);border:1.5px dashed rgba(123,87,255,.28);
  border-radius:11px;display:flex;flex-direction:column;align-items:center;
  justify-content:center;gap:3px;color:#8b7fc0;font-size:15px;font-family:inherit;transition:all .06s;
}
.card-add:hover{background:rgba(123,87,255,.11);border-color:rgba(123,87,255,.5);}

/* ── SETTINGS ── */
#settings{position:fixed;top:0;right:0;bottom:0;width:300px;background:var(--panelBg);border-left:1px solid var(--panelBorder);z-index:500;display:flex;flex-direction:column;box-shadow:-12px 0 40px rgba(0,0,0,.3);animation:slideIn .08s ease;backdrop-filter:none;}
#settings-backdrop{position:fixed;inset:0;z-index:499;background:rgba(10,15,30,.14);}
.settings-head{display:flex;align-items:center;justify-content:space-between;padding:16px 16px 12px;border-bottom:1px solid var(--navBorder);}
.settings-title{font-size:38px;font-weight:700;letter-spacing:-.02em;}
.settings-close{width:32px;height:32px;border-radius:8px;border:1px solid var(--inputBorder);background:var(--inputBg);color:var(--textMuted);cursor:pointer;font-size:22px;line-height:1;}
.settings-close.nav-focus,.settings-tab.nav-focus,.theme-btn.nav-focus,.toggle-switch.nav-focus,.range-slider.nav-focus,.pill-btn.nav-focus,.hotkey-input.nav-focus,.small-search.nav-focus,.hist-item.nav-focus,.action-btn.nav-focus,.hist-clear.nav-focus{
  box-shadow:0 0 0 3px var(--accentGlow,rgba(123,87,255,.22))!important;
  border-color:var(--accentPrimary,rgba(123,87,255,.46))!important;
}
.settings-body{flex:1;overflow-y:auto;padding:12px 14px;}
.settings-body::-webkit-scrollbar{width:8px;}
.settings-body::-webkit-scrollbar-thumb{background:var(--accentGlow,rgba(124,58,237,.35));border-radius:10px;}
.settings-tabs{display:flex;gap:6px;background:var(--inputBg);border:1px solid var(--inputBorder);border-radius:22px;padding:4px;margin-bottom:12px;position:static;z-index:1;}
.settings-tab{flex:1;height:32px;border:none;border-radius:16px;background:transparent;color:var(--textDim);font-size:13px;font-weight:700;cursor:pointer;transition:all .08s;}
.settings-tab.active{background:linear-gradient(135deg,var(--accentPrimary,#6d5efc),#8b5cf6);color:#fff;box-shadow:0 5px 14px rgba(109,94,252,.25);}
.settings-panel{display:none;}
.settings-panel.active{display:block;}
.settings-card{background:var(--cardBg);border:1px solid var(--inputBorder);border-radius:12px;padding:12px;margin-bottom:12px;}
.settings-section{margin-bottom:12px;background:var(--cardBg);border:1px solid var(--inputBorder);border-radius:12px;padding:12px;}
.s-title{font-size:13px;color:var(--sectionLabel);font-weight:700;letter-spacing:.08em;text-transform:uppercase;margin-bottom:9px;}
.theme-grid{display:flex;gap:7px;}
.theme-btn{flex:1;padding:9px 6px;border-radius:9px;cursor:pointer;border:2px solid var(--inputBorder);background:var(--inputBg);transition:all .08s;display:flex;flex-direction:column;align-items:center;gap:4px;}
.theme-btn.active{border-color:rgba(139,92,246,.7);background:rgba(109,40,217,.14);}
.theme-swatch{width:26px;height:26px;border-radius:7px;border:1px solid rgba(139,92,246,.25);}
.theme-label{font-size:15px;font-weight:500;color:var(--textDim);}
.theme-btn.active .theme-label{color:#a78bfa;}
.hotkey-row{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:8px;align-items:center;}
.hotkey-head{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:8px;}
.hotkey-input{
  width:100%;min-width:0;height:34px;border-radius:8px;border:1px solid var(--inputBorder);
  background:var(--inputBg);color:var(--text);padding:0 10px;font-size:14px;outline:none;
}
.hotkey-input:focus{border-color:rgba(123,87,255,.62);box-shadow:0 0 0 2px rgba(123,87,255,.18);}
.hotkey-save{width:auto;flex-shrink:0;margin:0;padding:8px 12px;height:34px;min-width:66px;font-size:14px;font-weight:700;border-radius:8px;}
.hotkey-note{margin-top:7px;font-size:13px;color:var(--textMuted);}
.hotkey-reset-mini{width:28px;padding:0;font-size:0;height:28px;border-radius:7px;min-width:28px;white-space:nowrap;margin:0;justify-content:center;}
.toggle-row{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:9px;}
.toggle-info{flex:1;}
.toggle-label{font-size:16px;font-weight:700;}
.toggle-sub{font-size:13px;color:var(--textMuted);margin-top:1px;}
.toggle-switch{width:36px;height:19px;border-radius:10px;position:relative;cursor:pointer;background:rgba(128,128,128,.22);border:1px solid rgba(128,128,128,.28);transition:all .08s;flex-shrink:0;}
.toggle-switch.on{background:rgba(109,40,217,.75);border-color:rgba(139,92,246,.6);}
.toggle-thumb{position:absolute;top:2px;left:2px;width:13px;height:13px;border-radius:50%;background:rgba(200,200,200,.8);transition:all .08s;}
.toggle-switch.on .toggle-thumb{left:17px;background:#fff;box-shadow:0 0 5px rgba(168,85,247,.55);}
.history-list{background:var(--inputBg);border:1px solid var(--inputBorder);border-radius:9px;overflow:hidden;}
.hist-item{width:100%;background:transparent;border:none;border-bottom:1px solid var(--navBorder);padding:7px 11px;text-align:left;color:var(--text);font-size:14px;cursor:pointer;font-family:'Cascadia Code','Cascadia Mono',Consolas,'Courier New',monospace;display:flex;justify-content:space-between;align-items:center;gap:7px;}
.hist-item:last-child{border-bottom:none;}
.hist-item:hover{background:rgba(139,92,246,.09);}
.hist-item-text{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;}
.hist-empty{padding:10px 12px;font-size:15px;color:var(--textMuted);font-style:italic;}
.hist-clear{margin-top:5px;background:none;border:none;color:var(--textMuted);font-size:15px;cursor:pointer;font-family:inherit;}
.action-btn{width:100%;padding:9px 11px;border-radius:10px;cursor:pointer;background:rgba(128,128,128,.05);border:1px solid var(--inputBorder);color:var(--textDim);font-size:16px;font-family:inherit;font-weight:600;text-align:left;transition:all .08s;display:flex;align-items:center;gap:7px;margin-bottom:6px;}
.action-btn:hover{background:rgba(128,128,128,.11);}
.pill-group{display:flex;gap:6px;flex-wrap:wrap;}
.pill-btn{
  min-width:66px;padding:7px 10px;border-radius:999px;cursor:pointer;
  border:1px solid var(--inputBorder);background:var(--inputBg);color:var(--textDim);
  font-size:13px;font-weight:600;transition:all .08s;
}
.pill-btn.active{background:rgba(109,40,217,.2);border-color:rgba(139,92,246,.55);color:var(--text);}
.range-row{display:flex;align-items:center;gap:10px;}
.range-slider{
  flex:1;appearance:none;height:8px;border-radius:999px;background:var(--inputBg);
  border:1px solid var(--inputBorder);outline:none;
}
.range-slider::-webkit-slider-thumb{
  appearance:none;width:16px;height:16px;border-radius:50%;
  background:var(--accentPrimary);border:none;cursor:pointer;
}
.range-val{min-width:42px;text-align:right;color:var(--textDim);font-size:13px;font-weight:600;}
.small-search{
  width:100%;height:32px;border-radius:999px;border:1px solid var(--inputBorder);
  background:var(--inputBg);color:var(--text);padding:0 12px;font-size:13px;outline:none;margin:8px 0;
}
.small-search:focus{border-color:var(--accentPrimary,rgba(123,87,255,.45));box-shadow:0 0 0 2px var(--accentGlow,rgba(123,87,255,.12));}
.backup-status{font-size:13px;color:var(--textMuted);margin-bottom:8px;}
.danger-btn{border-color:rgba(239,68,68,.45)!important;color:#ef4444!important;}
.danger-action{color:#b4232d!important;border-color:#ef9aa3!important;font-weight:800!important;}
.slider-hints{display:flex;justify-content:space-between;margin-top:5px;color:var(--textMuted);font-size:11px;font-weight:600;}

/* ── TOAST ── */
#toast{position:fixed;bottom:20px;left:50%;transform:translateX(-50%);padding:7px 20px;border-radius:100px;font-size:16px;font-weight:500;box-shadow:0 6px 24px rgba(0,0,0,.3);z-index:999;white-space:nowrap;animation:fadeUp .18s ease;pointer-events:none;display:none;}
#toast.ok{background:var(--toastBg);border:1px solid var(--toastBorder);color:var(--text);}
#toast.warn{background:var(--toastWarnBg);border:1px solid var(--toastWarnBorder);color:var(--text);}
</style>
</head>
<body class="daylight">

<!-- Aurora pink blob (daylight only) -->
<div id="aurora-blob"></div>

<!-- NAV -->
<nav id="nav">
  <!-- FIX #5: new logo -->
  <div class="logo">
    <span class="logo-name">Clippy</span>
  </div>

  <div class="search-wrap">
    <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
    <input id="search" placeholder="Search entries…" oninput="onSearch(this.value)" onkeydown="onSearchKey(event)"/>
    <button class="search-clear" id="search-clear" style="display:none" onclick="clearSearch()">×</button>
  </div>

  <button class="btn-new" onclick="addRow()">+ New Group</button>

  <button class="btn-gear" id="gear-btn" onclick="toggleSettings()">
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="12" cy="12" r="3"/>
      <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/>
    </svg>
  </button>
</nav>

<!-- MAIN (scroll-wrap prevents card rows intercepting page scroll) -->
<div id="scroll-wrap"><div id="main"></div></div>

<!-- SETTINGS BACKDROP -->
<div id="settings-backdrop" style="display:none" onclick="closeSettings()"></div>

<!-- SETTINGS -->
<div id="settings" style="display:none">
  <div class="settings-head">
    <span class="settings-title">Settings</span>
    <button class="settings-close" onclick="closeSettings()">×</button>
  </div>
  <div class="settings-body">
    <div class="settings-tabs">
      <button class="settings-tab active" id="stab-general" onclick="switchSettingsTab('general')">General</button>
      <button class="settings-tab" id="stab-capture" onclick="switchSettingsTab('capture')">Capture</button>
      <button class="settings-tab" id="stab-data" onclick="switchSettingsTab('data')">Data</button>
    </div>

    <div class="settings-panel active" data-stab="general">

    <div class="settings-section">
      <div class="s-title">General</div>
      <div class="toggle-row">
        <div class="toggle-info">
          <div class="toggle-label">Launch at Windows startup</div>
          <div class="toggle-sub">Start Clippy automatically after login</div>
        </div>
        <div class="toggle-switch on" id="startup-toggle" tabindex="0" onclick="toggleStartup()">
          <div class="toggle-thumb"></div>
        </div>
      </div>
    </div>

    <div class="settings-section">
      <div class="s-title">Display</div>
      <div class="theme-grid">
        <button class="theme-btn" id="theme-night" onclick="setTheme('night')">
          <div class="theme-swatch" style="background:linear-gradient(135deg,#090e1a,#1a2544)"></div>
          <span class="theme-label">Night</span>
        </button>
        <button class="theme-btn active" id="theme-daylight" onclick="setTheme('daylight')">
          <div class="theme-swatch" style="background:linear-gradient(135deg,#f8fbff,#eef3ff,#f7f2ff)"></div>
          <span class="theme-label">Daylight</span>
        </button>
      </div>
    </div>

    </div>
    <div class="settings-panel" data-stab="capture">

    <div class="settings-section">
      <div class="s-title">Capture</div>
      <div class="toggle-row">
        <div class="toggle-info">
          <div class="toggle-label">Auto-capture</div>
          <div class="toggle-sub">Reads clipboard at selected speed</div>
        </div>
        <div class="toggle-switch on" id="auto-toggle" tabindex="0" onclick="toggleCapture()">
          <div class="toggle-thumb"></div>
        </div>
      </div>
      <div class="toggle-sub" style="margin-bottom:6px">Capture speed</div>
      <div class="range-row">
        <input id="poll-slider" class="range-slider" type="range" min="0" max="2" step="1" value="1" oninput="setPollRateFromSlider(this.value)" />
      </div>
      <div class="slider-hints"><span>Fast</span><span>Normal</span><span>Relaxed</span></div>
      <div style="height:10px"></div>
      <div class="toggle-sub" style="margin-bottom:6px">Paste format</div>
      <div class="pill-group">
        <button class="pill-btn active" id="paste-plain" onclick="setPasteMode(true)">Plain text only</button>
        <button class="pill-btn" id="paste-rich" onclick="setPasteMode(false)">Preserve formatting</button>
      </div>
    </div>

    <div class="settings-section">
      <div class="hotkey-head">
        <div class="s-title" style="margin-bottom:0">Hotkey</div>
        <button class="action-btn hotkey-reset-mini" id="hotkey-reset" onclick="resetHotkey()" title="Reset hotkey">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M3 12a9 9 0 1 0 3-6.7"/>
            <path d="M3 3v6h6"/>
          </svg>
        </button>
      </div>
      <div class="hotkey-row">
        <input id="hotkey-input" class="hotkey-input" value="Ctrl + D" placeholder="Press your keys" tabindex="-1" onmousedown="onHotkeyInputMouseDown(event)" onkeydown="onHotkeyInputKey(event)" onblur="hotkeyCaptureArmed=false" oninput="this.dataset.value=''" />
        <button class="action-btn hotkey-save" id="hotkey-save" onclick="saveHotkey()">Save</button>
      </div>
      <div class="hotkey-note">Default: Ctrl + D</div>
    </div>

    <div class="settings-section">
      <div class="s-title">Paste History</div>
      <div class="toggle-row">
        <div class="toggle-info">
          <div class="toggle-label">Record history</div>
          <div class="toggle-sub">Save captured items for quick reuse</div>
        </div>
        <div class="toggle-switch on" id="hist-toggle" tabindex="0" onclick="toggleHistory()">
          <div class="toggle-thumb"></div>
        </div>
      </div>
      <div class="toggle-sub" style="margin-bottom:6px">History size</div>
      <div class="range-row">
        <input id="hist-limit-slider" class="range-slider" type="range" min="0" max="3" step="1" value="1" oninput="setHistoryLimitFromSlider(this.value)" />
      </div>
      <div class="slider-hints"><span>5</span><span>10</span><span>20</span><span>50</span></div>
      <input id="history-search" class="small-search" placeholder="Search history..." oninput="onHistorySearch(this.value)" />
      <div id="history-section">
        <div class="history-list" id="history-list"><div class="hist-empty">Nothing captured yet</div></div>
        <button class="hist-clear danger-action" id="hist-clear" style="display:none" onclick="clearHistory()">Clear history</button>
      </div>
    </div>

    </div>
    <div class="settings-panel" data-stab="data">

    <div class="settings-section">
      <div class="s-title">Backup</div>
      <div class="toggle-row">
        <div class="toggle-info">
          <div class="toggle-label">Auto backup</div>
          <div class="toggle-sub">Run daily backup automatically</div>
        </div>
        <div class="toggle-switch on" id="backup-toggle" tabindex="0" onclick="toggleBackupEnabled()">
          <div class="toggle-thumb"></div>
        </div>
      </div>
      <div class="backup-status" id="backup-status">Last backup: Never</div>
      <button class="action-btn" onclick="bridge.openBackupFolder()">Open Backup Folder</button>
      <button class="action-btn" onclick="bridge.backupNow()">Backup Now</button>
    </div>

    <div class="settings-section">
      <div class="s-title">Data</div>
      <button class="action-btn" onclick="bridge.exportJSON()">⬇ Export JSON</button>
      <button class="action-btn" onclick="bridge.exportCSV()">⬇ Export CSV</button>
      <button class="action-btn" onclick="bridge.importJSON()">⬆ Import JSON</button>
      <button class="action-btn danger-action" id="clear-all-btn" onclick="clearAllDataConfirm()">Clear All Data</button>
    </div>

    </div>

  </div>
</div>

<div id="toast"></div>

<script>
// ── State ─────────────────────────────────────────────────────────────────────
let bridge = null;
let rows = [], cursor = {row_idx:0,entry_idx:0};
let autoCapture = true, history = [], histEnabled = true;
let hotkey = 'ctrl+d';
let pollRate = 500, pastePlainText = true;
let launchAtStartup = true, backupEnabled = true, historyLimit = 10, lastBackup = '';
let historySearch = '';
let hotkeyCaptureArmed = false;
let editingId = null, kbIdx = -1, searchTerm = '';
let settingsOpen = false, dragCard = null, dragRow = null;
let userTargetId = null;
let settingsNavIdx = -1;
let settingsTab = 'general';
let topNavIdx = -1;
let clearAllConfirmTimer = null;
let clearAllArmed = false;

// Colors match mockups: violet, sky-blue, teal, coral, amber
const TAG_CYCLE = [
  {dot:"#857dce",num:"#857dce",border:"rgba(217,211,255,.9)",darkNum:"#94a7d8",darkBorder:"rgba(142,107,255,.55)"},
  {dot:"#7cb3ec",num:"#7cb3ec",border:"rgba(191,230,255,.9)",darkNum:"#94a7d8",darkBorder:"rgba(47,90,138,.75)"},
  {dot:"#6dcbb0",num:"#6dcbb0",border:"rgba(187,238,221,.9)",darkNum:"#94a7d8",darkBorder:"rgba(42,106,117,.75)"},
  {dot:"#fb7185",num:"#f87171",border:"rgba(251,113,133,.38)",darkNum:"#f87171",darkBorder:"rgba(251,113,133,.38)"},
  {dot:"#fbbf24",num:"#fbbf24",border:"rgba(251,191,36,.38)",darkNum:"#fbbf24",darkBorder:"rgba(251,191,36,.38)"},
];
function _isDark(){ return document.body.classList.contains('dark')||document.body.classList.contains('night'); }
function tag(i) {
  const t = TAG_CYCLE[i % TAG_CYCLE.length];
  return _isDark()
    ? {dot:t.dot, num:t.darkNum, border:t.darkBorder}
    : {dot:t.dot, num:t.num,     border:t.border};
}

// ── Init ──────────────────────────────────────────────────────────────────────
new QWebChannel(qt.webChannelTransport, function(ch) {
  bridge = ch.objects.bridge;
  bridge.stateChanged.connect(window.__onState);
  bridge.getState(function(json) { applyState(JSON.parse(json)); });
});

// Hide Clippy when the WebEngine window loses focus (user clicked elsewhere).
// The delay must be LONGER than the Python _showing guard (1500ms) so that
// blur events from the show-animation focus dance (including the WindowStaysOnTopHint
// drop at 400ms and foreground retries) are suppressed before hideWindow() fires.
window.__clippyPointerInside = false;
window.__clippyLastShowTs = 0;
window.addEventListener('mouseenter', function(){ window.__clippyPointerInside = true; });
window.addEventListener('mouseleave', function(){ window.__clippyPointerInside = false; });
window.addEventListener('blur', function() {
  setTimeout(function() {
    const sinceShow = Date.now() - (window.__clippyLastShowTs || 0);
    if (!document.hasFocus() && !window.__clippyPointerInside && !settingsOpen && sinceShow > 1700 && bridge) {
      bridge.hideWindow();
    }
  }, 500);
});

// Real __onState — replaces the early stub set in <head>.
// After defining it, drain any events that arrived during bridge init.
window.__onState = function(json) {
  const d = typeof json==='string' ? JSON.parse(json) : json;
  if (d.toast) showToast(d.toast, d.toast_type||'ok');
  if (d.rows !== undefined) applyState(d);
};

// Drain the early queue — replay any captures that arrived before bridge ready
if (window.__earlyStateQueue && window.__earlyStateQueue.length) {
  window.__earlyStateQueue.forEach(function(payload) {
    window.__onState(payload);
  });
  window.__earlyStateQueue = [];
}

function applyState(d) {
  rows        = d.rows        || [];
  cursor      = d.cursor      || {row_idx:0,entry_idx:0};
  autoCapture = d.auto        ?? true;
  history     = d.history     || [];
  histEnabled = d.hist_enabled ?? true;
  hotkey      = (d.hotkey || 'ctrl+d').toLowerCase();
  pollRate    = d.poll_rate ?? 500;
  pastePlainText = d.paste_plain_text ?? true;
  launchAtStartup = d.launch_at_startup ?? true;
  backupEnabled = d.backup_enabled ?? true;
  historyLimit = d.history_limit ?? 10;
  lastBackup = d.last_backup || '';
  // Restore persisted theme — use _applyThemeUI (DOM only) so we never call
  // bridge.setTheme() here, which would trigger save() → _push() → applyState
  // again — an infinite loop that was causing Clippy to freeze.
  if (d.theme) _applyThemeUI(d.theme);

  // ── Auto-save on capture ─────────────────────────────────────────────────
  // When a card is open in edit mode (editingId is set) and the clipboard
  // poller has just filled that slot with captured text, auto-save and exit
  // edit mode immediately — no need for the user to click Save.
  //
  // Two sub-cases:
  //   A) User opened the box but typed nothing (pendingText empty):
  //      the captured text is already written into state by the poller,
  //      so just clearing editingId is enough — the card displays correctly.
  //   B) User had started typing manually (pendingText has content):
  //      honour their input — save the pending text, then exit edit mode.
  if (editingId) {
    const pending = pendingText[editingId];
    if (pending) {
      // User typed something — commit it before the render overwrites the box
      bridge.saveEntry(editingId, pending);
      delete pendingText[editingId];
    } else {
      // No manual input — check if the poller already wrote text into this slot
      const captured = rows.flatMap(r => r.entries).find(e => e.id === editingId);
      if (captured && captured.text) {
        // Poller filled the slot — nothing more to save, just exit edit mode
      } else {
        // Slot still empty (capture hasn't arrived yet) — keep edit mode open
        renderAll();
        if (settingsOpen) updateSettings();
        return;
      }
    }
    editingId = null;
  }

  renderAll();
  if (settingsOpen) updateSettings();
}

// ── Toast ─────────────────────────────────────────────────────────────────────
let toastTimer=null;
function showToast(msg, type='ok') {
  const t = document.getElementById('toast');
  t.textContent=msg; t.className=type; t.style.display='block';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(()=>{ t.style.display='none'; }, 2200);
}

// ── Theme ─────────────────────────────────────────────────────────────────────
// _applyThemeUI — update DOM only, no bridge call, no state write.
// Used by applyState so restoring a persisted theme never triggers a round-trip
// back to Python (which would re-push state → re-trigger applyState → loop).
let _currentTheme = 'daylight';
function _applyThemeUI(t) {
  if (t === 'dark') t = 'night'; // Backward compatibility with older saved theme values.
  if (!t) return;
  const changed = t !== _currentTheme;
  _currentTheme = t;
  if (changed) document.body.className = t;
  document.querySelectorAll('.theme-btn').forEach(b => b.classList.remove('active'));
  const btn = document.getElementById('theme-' + t);
  if (btn) btn.classList.add('active');
}

// setTheme — called by the user clicking a theme button.
// Updates the DOM AND persists the choice to Python (only if actually changed).
function setTheme(t) {
  if (t === 'dark') t = 'night';
  if (t === _currentTheme) return;        // user clicked the already-active theme
  _applyThemeUI(t);
  if (bridge) bridge.setTheme(t);
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function escHtml(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function escAttr(s){ return String(s).replace(/'/g,"&#39;").replace(/"/g,'&quot;'); }

function highlight(text, term) {
  if (!text) return '<span class="empty-slot">Waiting for copy…</span>';
  if (!term)  return escHtml(text);
  const i = text.toLowerCase().indexOf(term.toLowerCase());
  if (i===-1) return escHtml(text);
  return escHtml(text.slice(0,i))+'<mark>'+escHtml(text.slice(i,i+term.length))+'</mark>'+escHtml(text.slice(i+term.length));
}

function heatBar(count) {
  if (!count) return '';
  const pct = Math.min(100,count*16);
  const col  = count>=7?'#f87171':count>=3?'#fbbf24':'#34d399';
  return `<div class="heatbar" style="width:${pct}%;background:${col}"></div>`;
}

function _buildVisible() {
  // Single authoritative sort+filter used by both renderAll and flatEntries.
  // Returns {sorted, visible} so callers can share the same arrays.
  const sorted = rows
    .map((r,i)=>({...r,__idx:i,entries:[...r.entries].sort((a,b)=>(b.pinned?1:0)-(a.pinned?1:0))}))
    .sort((a,b)=>{
      const ap = a.pinned ? 1 : 0;
      const bp = b.pinned ? 1 : 0;
      if (bp !== ap) return bp - ap;
      return a.__idx - b.__idx;
    });
  const visible = searchTerm
    ? sorted.filter(r=>r.entries.some(e=>e.text&&e.text.toLowerCase().includes(searchTerm)))
    : sorted;
  return {sorted, visible};
}

function flatEntries(visibleRows) {
  // Must match the sorted order used by renderAll (pinned first) so kbIdx lines up with what's on screen.
  // Accepts pre-built visibleRows to avoid recomputing sort/filter when called from renderAll.
  const src = visibleRows || _buildVisible().visible;
  const out=[];
  src.forEach((row,rowIndex)=>row.entries.forEach((e,ei)=>out.push({ei,e,rowId:row.id,rowIndex})));
  return out;
}

function activeId() {
  if (!autoCapture||cursor.row_idx<0) return null;
  return rows[cursor.row_idx]?.entries[cursor.entry_idx]?.id ?? null;
}

// ── Render ────────────────────────────────────────────────────────────────────
function renderAll() {
  const main  = document.getElementById('main');
  // Build sorted+visible ONCE — pass into flatEntries so sort/filter never runs twice
  const {visible} = _buildVisible();
  const flat  = flatEntries(visible);
  const actId = activeId();

  if (visible.length===0) {
    main.innerHTML=`<div class="empty">${searchTerm?'No matches for "'+escHtml(searchTerm)+'"':'No groups yet — click + New Group or copy something!'}</div>`;
    return;
  }

  main.innerHTML = visible.map((row, visIdx) => {
    const stateIdx  = rows.findIndex(r=>r.id===row.id);
    const isActiveRow = stateIdx===cursor.row_idx && autoCapture;
    const t         = tag(row.color_idx??stateIdx);
    const hasMatch  = searchTerm && row.entries.some(e=>e.text.toLowerCase().includes(searchTerm));

    const cards = row.entries.map(entry => {
      const fi     = flat.findIndex(f=>f.e.id===entry.id);
      const isAct  = entry.id===actId;
      const isKb   = fi===kbIdx;
      const isHit  = searchTerm && entry.text.toLowerCase().includes(searchTerm);
      const isUser = entry.id===userTargetId;
      const isCode = entry.text && /[{};()=>]/.test(entry.text);

      let cls='card';
      if (editingId===entry.id)    cls+=' editing';
      else if (isKb)               cls+=' kb-focus';
      else if (isUser&&isAct)      cls+=' user-target';
      else if (isHit)              cls+=' search-hit';
      else if (isAct)              cls+=' active-slot';

      const dot   = autoCapture&&isAct&&editingId!==entry.id ? '<div class="card-dot"></div>' : '';
      const kbBar = isKb ? '<div class="card-kb-bar"></div>' : '';
      const pin   = entry.pinned ? '<div class="card-pin">📌</div>' : '';

      let body;
      if (editingId===entry.id) {
        // IMPORTANT: textarea value is set via DOM (.value=) after innerHTML,
        // NOT via escHtml() inside the template. Setting innerHTML with escHtml()
        // double-encodes entities (&amp; → &amp;amp;) which permanently corrupts
        // text containing & < > on save.
        body = `<div style="flex:1;overflow:hidden"><textarea class="card-textarea" id="ta-${entry.id}"
          onkeydown="onCardKey(event,'${entry.id}')"
          oninput="syncTA('${entry.id}',this.value)"></textarea></div>
          <div class="card-actions" style="opacity:1;pointer-events:auto">
            <button class="save-btn" onclick="saveCard('${entry.id}')">✓ Save</button>
          </div>`;
      } else {
        body = `<div style="flex:1;overflow:hidden;padding-left:${isKb?6:0}px;transition:padding .05s">
            <div class="card-text${isCode?' code-font':''}">${highlight(entry.text,searchTerm)}</div>
          </div>
          <div class="card-actions">
            <button class="cbtn" title="Edit" onclick="event.stopPropagation();editCard('${entry.id}')">${penIco()}</button>
            <button class="cbtn" title="Copy" onclick="event.stopPropagation();copyCard('${entry.id}','${escAttr(entry.text)}')">${copyIco()}</button>
            <button class="cbtn${entry.pinned?' active':''}" title="Pin" onclick="event.stopPropagation();bridge.pinEntry('${entry.id}')">📌</button>
            <button class="cbtn danger" title="Delete" onclick="event.stopPropagation();delCard('${row.id}','${entry.id}')">${trashIco()}</button>
          </div>`;
      }

      return `<div class="${cls}" id="card-${entry.id}"
        onclick="onCardClick(event,'${row.id}','${entry.id}')"
        onmouseenter="setCardHover(this,true)"
        onmouseleave="setCardHover(this,false)"
        draggable="true"
        ondragstart="onCardDragStart(event,'${row.id}','${entry.id}')"
        ondragover="event.preventDefault()"
        ondrop="onCardDrop(event,'${row.id}','${entry.id}')">
        ${dot}${kbBar}${pin}${heatBar(entry.copy_count||0)}${body}
      </div>`;
    }).join('');

    const addCard = row.entries.length<4
      ? `<button class="card-add" onclick="addEntry('${row.id}')"><span style="font-size:16px">+</span><span>Add variation</span></button>`
      : '';

    return `<div class="row-wrap" id="rowwrap-${row.id}" draggable="true"
      ondragstart="onRowDragStart(event,'${row.id}')"
      ondragover="onRowDragOver(event,'${row.id}')"
      ondrop="onRowDrop(event,'${row.id}')"
      ondragend="onRowDragEnd()">
      <span class="row-num${isActiveRow?' active':''}" style="${isActiveRow?'':'color:'+t.num}">${visIdx+1}</span>
      <div class="row-group${isActiveRow?' active-capture':''}${hasMatch?' search-match':''}" style="border-color:${t.border}">
        <button class="row-del" onclick="delRow('${row.id}')">✕</button>
        <button class="row-pin${row.pinned ? ' active' : ''}" onclick="bridge.pinRow('${row.id}')">📌</button>
        ${cards}${addCard}
        <div class="row-btn-strip">
          <button class="row-del" onclick="delRow('${row.id}')" title="Close group">&times;</button>
          <button class="row-pin${row.pinned ? ' active' : ''}" onclick="bridge.pinRow('${row.id}')" title="Pin group">&#128204;</button>
        </div>
      </div>
    </div>`;
  }).join('');

  if (editingId) {
    const ta=document.getElementById('ta-'+editingId);
    if (ta) {
      // Set value via DOM — NOT via innerHTML/escHtml — so special chars are never double-encoded
      const entry = rows.flatMap(r=>r.entries).find(e=>e.id===editingId);
      ta.value = entry ? (pendingText[editingId] ?? entry.text) : '';
      ta.focus();
      ta.selectionStart = ta.selectionEnd = ta.value.length;
    }
  }
  _attachRowHoverListeners();
}

// ── Card click — set user target ──────────────────────────────────────────────
function _attachRowHoverListeners(){
  document.querySelectorAll('.row-wrap').forEach(wrap => {
    const strip = wrap.querySelector('.row-btn-strip');
    if (!strip) return;
    wrap.addEventListener('mouseenter', () => {
      strip.style.width = '30px';
      strip.style.minWidth = '30px';
      strip.style.opacity = '1';
      strip.style.pointerEvents = 'auto';
    });
    wrap.addEventListener('mouseleave', () => {
      strip.style.width = '0';
      strip.style.minWidth = '0';
      strip.style.opacity = '0';
      strip.style.pointerEvents = 'none';
    });
  });
}

function onCardClick(event, rowId, entryId) {
  if (event.target.closest('button')||event.target.tagName==='TEXTAREA') return;
  userTargetId = entryId;
  bridge.setCursorToEntry(rowId, entryId);
  renderAll();
}

// ── Editing ───────────────────────────────────────────────────────────────────
function setCardHover(el, on){
  if(!el) return;
  el.classList.toggle('hover-actions', !!on);
}
const pendingText={};
function syncTA(id,val){pendingText[id]=val;}
function editCard(id){editingId=id;renderAll();}
function saveCard(id){
  const text=pendingText[id]??rows.flatMap(r=>r.entries).find(e=>e.id===id)?.text??'';
  bridge.saveEntry(id,text); editingId=null; delete pendingText[id];
}

// ── Copy (manual, no paste) ───────────────────────────────────────────────────
function copyCard(id, text) {
  userTargetId=null;
  bridge.copyEntry(id, text);
  showToast('📋 Copied!');
}

// ── Add / delete ──────────────────────────────────────────────────────────────
function addEntry(rowId){ bridge.addEntry(rowId); }
function delCard(rowId, entryId){ if(userTargetId===entryId)userTargetId=null; bridge.deleteEntry(rowId,entryId); }
function addRow(){
  bridge.addRow(function(json){
    const d = JSON.parse(json);
    editingId = d.new_entry_id;
    // Scroll the new group into view — it is always appended at the bottom
    // and may be off-screen if the list is long.
    requestAnimationFrame(function() {
      const el = document.getElementById('rowwrap-' + d.new_row_id);
      if (el) el.scrollIntoView({block:'nearest', behavior:'auto'});
    });
  });
}
function delRow(rowId){ bridge.deleteRow(rowId); showToast('Group removed','warn'); }

// ── Search ────────────────────────────────────────────────────────────────────
function onSearch(val){
  searchTerm=val.trim().toLowerCase();
  document.getElementById('search-clear').style.display=val?'block':'none';
  // Auto-highlight first match so Enter immediately pastes it
  if (searchTerm) {
    const {visible} = _buildVisible();
    const flat = flatEntries(visible);
    const mi = flat.findIndex(f=>f.e.text&&f.e.text.toLowerCase().includes(searchTerm));
    kbIdx = mi >= 0 ? mi : -1;
    const inp = document.getElementById('search');
    inp.placeholder = mi >= 0
      ? `${flat.filter(f=>f.e.text&&f.e.text.toLowerCase().includes(searchTerm)).length} · Enter to paste`
      : 'No matches';
  } else {
    kbIdx = -1;
    document.getElementById('search').placeholder = 'Search entries…';
  }
  renderAll();
}
function clearSearch(){document.getElementById('search').value='';onSearch('');}
function onSearchKey(e){
  if (e.key!=='Enter') return;
  // Always paste the kbIdx-highlighted entry (arrow keys move it).
  // Fall back to first match only when kbIdx is unset or off-screen.
  const flat=flatEntries();
  const matches=flat.filter(f=>f.e.text.toLowerCase().includes(searchTerm));
  const highlighted=(kbIdx>=0 && kbIdx<flat.length) ? flat[kbIdx] : null;
  const isMatch=highlighted && highlighted.e.text.toLowerCase().includes(searchTerm);
  const target=isMatch ? highlighted : matches[0];
  if (target && target.e.text) {
    clearSearch();
    bridge.pasteEntry(target.e.id, target.e.text);
  }
}

// ── Keyboard navigation ───────────────────────────────────────────────────────
function _settingsNavItems() {
  const ordered = [
    document.querySelector('#settings .settings-close'),
    ...Array.from(document.querySelectorAll('#settings .settings-tab')),
    ...Array.from(document.querySelectorAll(
      '#settings .theme-btn, ' +
      '#settings #startup-toggle, ' +
      '#settings #hotkey-input, ' +
      '#settings #hotkey-save, ' +
      '#settings #hotkey-reset, ' +
      '#settings #auto-toggle, ' +
      '#settings #backup-toggle, ' +
      '#settings #hist-toggle, ' +
      '#settings #poll-slider, ' +
      '#settings #hist-limit-slider, ' +
      '#settings .pill-btn, ' +
      '#settings #history-search, ' +
      '#settings #hist-clear, ' +
      '#settings .hist-item, ' +
      '#settings .action-btn'
    ))
  ];
  return ordered.filter(el => _isSettingsNavVisible(el) && !el.disabled);
}

function _settingsTabs() {
  return Array.from(document.querySelectorAll('#settings .settings-tab')).filter(el => _isSettingsNavVisible(el) && !el.disabled);
}

function _isSettingsNavVisible(el) {
  if (!el) return false;
  const style = window.getComputedStyle(el);
  if (style.display === 'none' || style.visibility === 'hidden') {
    return false;
  }
  const panel = el.closest('.settings-panel');
  if (panel) {
    const panelStyle = window.getComputedStyle(panel);
    if (panelStyle.display === 'none' || panelStyle.visibility === 'hidden') {
      return false;
    }
  }
  if (el.classList?.contains('settings-tab') || el.classList?.contains('settings-close')) {
    return true;
  }
  return true;
}

function _settingsPanelItems() {
  return _settingsNavItems().filter(el =>
    !el.classList.contains('settings-tab') &&
    !el.classList.contains('settings-close')
  );
}

function _firstSettingsPanelItem() {
  return _settingsPanelItems()[0] || null;
}

function _focusSettingsEl(el) {
  if (!el) return false;
  const items = _settingsNavItems();
  const idx = items.indexOf(el);
  if (idx < 0) return false;
  _focusSettingsItem(idx);
  return true;
}

function _settingsItemIndex(el) {
  if (!el) return -1;
  return _settingsNavItems().indexOf(el);
}

function _focusSettingsSibling(active, direction) {
  if (!active) return false;
  const container = active.closest('.settings-tabs, .theme-grid, .pill-group, .hotkey-row');
  if (!container) return false;
  const items = _settingsNavItems().filter(el => el.closest('.settings-tabs, .theme-grid, .pill-group, .hotkey-row') === container);
  const idx = items.indexOf(active);
  if (idx < 0) return false;
  const next = direction === 'left'
    ? items[Math.max(0, idx - 1)]
    : items[Math.min(items.length - 1, idx + 1)];
  if (!next || next === active) return false;
  return _focusSettingsEl(next);
}

function _focusSettingsItem(idx) {
  const items = _settingsNavItems();
  if (!items.length) {
    settingsNavIdx = -1;
    return;
  }
  const n = items.length;
  settingsNavIdx = ((idx % n) + n) % n;
  items.forEach(el => el.classList.remove('nav-focus'));
  const el = items[settingsNavIdx];
  el.classList.add('nav-focus');
  if (typeof el.focus === 'function') el.focus({preventScroll:true});
  const behavior = el.classList?.contains('settings-tab') ? 'smooth' : 'auto';
  el.scrollIntoView({block:'nearest', inline:'nearest', behavior});
}

function _topNavItems() {
  return [
    document.getElementById('search'),
    document.querySelector('.btn-new'),
    document.getElementById('gear-btn'),
  ].filter(Boolean);
}

function _focusTopNav(idx) {
  const items = _topNavItems();
  if (!items.length) return;
  const n = items.length;
  topNavIdx = ((idx % n) + n) % n;
  items.forEach(el => el.classList.remove('nav-focus'));
  const el = items[topNavIdx];
  el.classList.add('nav-focus');
  if (typeof el.focus === 'function') el.focus({preventScroll:true});
}

function _focusSettingsByPosition(direction) {
  const items = _settingsNavItems();
  if (!items.length) return false;
  const current = document.activeElement;
  const base = items.includes(current) ? current : items[Math.max(0, settingsNavIdx)];
  if (!base) return false;
  const baseRect = base.getBoundingClientRect();
  let best = null;
  let bestScore = Infinity;
  for (const item of items) {
    if (item === base) continue;
    const rect = item.getBoundingClientRect();
    const centerX = rect.left + rect.width / 2;
    const centerY = rect.top + rect.height / 2;
    const baseX = baseRect.left + baseRect.width / 2;
    const baseY = baseRect.top + baseRect.height / 2;
    const dx = centerX - baseX;
    const dy = centerY - baseY;
    let valid = false;
    let score = Infinity;
    if (direction === 'down' && dy > 4) { valid = true; score = dy * 10 + Math.abs(dx); }
    if (direction === 'up' && dy < -4) { valid = true; score = Math.abs(dy) * 10 + Math.abs(dx); }
    if (direction === 'right' && dx > 4) { valid = true; score = dx * 10 + Math.abs(dy); }
    if (direction === 'left' && dx < -4) { valid = true; score = Math.abs(dx) * 10 + Math.abs(dy); }
    if (valid && score < bestScore) {
      best = item;
      bestScore = score;
    }
  }
  if (!best) return false;
  const idx = items.indexOf(best);
  if (idx >= 0) {
    _focusSettingsItem(idx);
    return true;
  }
  return false;
}

function _handleSettingsKeys(e) {
  if (!settingsOpen) return false;
  if (e.defaultPrevented) return true;
  const key = e.key;
  if (!['ArrowDown', 'ArrowUp', 'ArrowLeft', 'ArrowRight', 'Enter', ' ', 'Escape', 'Home', 'End'].includes(key)) {
    return false;
  }
  e.preventDefault();
  e.stopPropagation();

  if (key === 'Escape') {
    closeSettings();
    document.getElementById('gear-btn')?.focus();
    return true;
  }
  const active = document.activeElement;
  const onTab = active?.classList?.contains('settings-tab');
  const onClose = active?.classList?.contains('settings-close');
  const tabs = _settingsTabs();
  const firstTab = tabs[0] || null;
  const activeTab = document.querySelector('#settings .settings-tab.active') ||
    tabs.find(tab => tab.id === 'stab-' + settingsTab) ||
    tabs[tabs.length - 1] ||
    null;
  const activeIdx = _settingsItemIndex(active);
  if (key === 'ArrowDown') {
    if (onTab) {
      const panelItems = _settingsPanelItems();
      if (panelItems.length) {
        const idx = _settingsNavItems().indexOf(panelItems[0]);
        _focusSettingsItem(idx);
      }
    } else if (active?.id === 'startup-toggle' && settingsTab === 'general') {
      _focusSettingsEl(document.getElementById('theme-night'));
    } else if (onClose) {
      _focusSettingsEl(firstTab);
    } else if (activeIdx >= 0) {
      _focusSettingsItem(activeIdx + 1);
    }
    return true;
  }
  if (key === 'ArrowUp') {
    if (onClose) {
      const items = _settingsPanelItems();
      const last = items[items.length - 1];
      if (!_focusSettingsEl(last)) {
        _focusSettingsEl(activeTab);
      }
    } else if (onTab && active === firstTab) {
      _focusSettingsEl(document.querySelector('#settings .settings-close'));
    } else if (active && active === _firstSettingsPanelItem()) {
      _focusSettingsEl(activeTab);
    } else if (activeIdx >= 0) {
      _focusSettingsItem(activeIdx - 1);
    }
    return true;
  }
  if (key === 'ArrowLeft') {
    if (onTab) {
      const ti = tabs.indexOf(active);
      const next = tabs[Math.max(0, ti - 1)];
      next?.click();
      next?.focus({preventScroll:true});
    } else if (active?.id === 'theme-daylight') {
      _focusSettingsEl(document.getElementById('theme-night'));
    } else if (onClose) {
      _focusSettingsEl(firstTab);
    } else if (!_focusSettingsSibling(active, 'left') && activeIdx >= 0) {
      _focusSettingsItem(activeIdx - 1);
    }
    return true;
  }
  if (key === 'ArrowRight') {
    if (onTab) {
      const ti = tabs.indexOf(active);
      const next = tabs[Math.min(tabs.length - 1, ti + 1)];
      next?.click();
      next?.focus({preventScroll:true});
    } else if (active?.id === 'theme-night') {
      _focusSettingsEl(document.getElementById('theme-daylight'));
    } else if (onClose) {
      const panelItems = _settingsPanelItems();
      if (panelItems.length) {
        const idx = _settingsNavItems().indexOf(panelItems[0]);
        _focusSettingsItem(idx);
      }
    } else if (!_focusSettingsSibling(active, 'right') && activeIdx >= 0) {
      _focusSettingsItem(activeIdx + 1);
    }
    return true;
  }
  if (key === 'Home') {
    if (onTab) {
      tabs[0]?.focus({preventScroll:true});
    } else {
      const items = _settingsPanelItems();
      const first = items[0] || _settingsNavItems()[0];
      const idx = _settingsNavItems().indexOf(first);
      if (idx >= 0) _focusSettingsItem(idx);
    }
    return true;
  }
  if (key === 'End') {
    if (onTab) {
      tabs[tabs.length - 1]?.focus({preventScroll:true});
    } else {
      const items = _settingsPanelItems();
      const last = items[items.length - 1] || _settingsNavItems()[_settingsNavItems().length - 1];
      const idx = _settingsNavItems().indexOf(last);
      if (idx >= 0) _focusSettingsItem(idx);
    }
    return true;
  }

  if (onTab) {
    const panelItems = _settingsPanelItems();
    if (panelItems.length) {
      const idx = _settingsNavItems().indexOf(panelItems[0]);
      _focusSettingsItem(idx);
    }
    return true;
  }

  const items = _settingsNavItems();
  if (!items.length) return true;
  const idx = settingsNavIdx >= 0 ? settingsNavIdx : 0;
  items[idx]?.click();
  return true;
}

document.addEventListener('keydown', e => {
  if (_handleSettingsKeys(e)) return;

  const isArrow = ['ArrowRight','ArrowLeft','ArrowUp','ArrowDown'].includes(e.key);
  if (topNavIdx >= 0 && isArrow) {
    e.preventDefault();
    if (e.key === 'ArrowLeft') { _focusTopNav(topNavIdx - 1); return; }
    if (e.key === 'ArrowRight') { _focusTopNav(topNavIdx + 1); return; }
    if (e.key === 'ArrowDown') {
      topNavIdx = -1;
      _topNavItems().forEach(el => el.classList.remove('nav-focus'));
      kbIdx = 0;
      renderAll();
      return;
    }
  }

  // Ctrl+D: arm nav, focus search so user can type immediately
  if ((e.ctrlKey||e.metaKey) && e.key==='d') {
    e.preventDefault();
    topNavIdx = -1;
    _topNavItems().forEach(el => el.classList.remove('nav-focus'));
    kbIdx = 0;
    renderAll(); return;
  }

  // Never steal from textarea (editing a card)
  if (document.activeElement?.tagName==='TEXTAREA') return;

  const isSearch = document.activeElement?.id==='search';

  // If the user starts typing a non-arrow, non-special key and search is not
  // already focused, redirect keystrokes into the search bar automatically.
  if (!isSearch && e.key.length===1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
    const s = document.getElementById('search');
    if (s) { s.focus(); }
    // Let the keystroke fall through so the character lands in the search input
    return;
  }

  // Arrow keys: if kbIdx is -1 (no card highlighted yet), auto-arm to card 0
  // so the first arrow press immediately highlights the first card rather than
  // doing nothing. This makes navigation available without any prior Ctrl+D.
  if (isArrow && kbIdx===-1) {
    kbIdx = 0;
  }

  // Only process nav keys when in nav mode OR search is active
  if (kbIdx===-1 && !isSearch) return;

  const flat=flatEntries();
  const scrollKbTarget = (block='nearest') => {
    const target = flat[kbIdx];
    if (!target) return;
    const rowEl = document.getElementById('rowwrap-' + target.rowId);
    const cardEl = document.getElementById('card-' + target.e.id);
    // For vertical row jumps, align the row cleanly with the top edge.
    // This avoids the "90% visible" stuck state when navigating upward.
    if (block === 'start' && rowEl) {
      rowEl.scrollIntoView({block:'start', inline:'nearest', behavior:'auto'});
      return;
    }
    cardEl?.scrollIntoView({block, inline:'nearest', behavior:'auto'});
  };

  if (e.key==='Escape') {
    e.preventDefault(); kbIdx=-1;
    document.getElementById('search')?.focus();
    renderAll(); return;
  }
  // ArrowRight = next card in row, ArrowDown = same column in next row when possible
  // ArrowLeft  = prev card in row, ArrowUp   = same column in previous row when possible
  if (e.key==='ArrowRight') {
    e.preventDefault();
    kbIdx=Math.min(kbIdx+1, flat.length-1);
    renderAll();
    scrollKbTarget('nearest');
  } else if (e.key==='ArrowLeft') {
    e.preventDefault();
    kbIdx=Math.max(kbIdx-1, 0);
    renderAll();
    scrollKbTarget('nearest');
  } else if (e.key==='ArrowDown') {
    e.preventDefault();
    const cur = flat[kbIdx];
    if (cur) {
      const nextRowItems = flat.filter(f => f.rowIndex === cur.rowIndex + 1);
      if (nextRowItems.length) {
        const preferred = nextRowItems.find(f => f.ei === cur.ei) || nextRowItems[Math.min(cur.ei, nextRowItems.length - 1)];
        const nextIdx = flat.findIndex(f => f.e.id === preferred.e.id);
        if (nextIdx !== -1) kbIdx = nextIdx;
      }
    }
    renderAll();
    scrollKbTarget('nearest');
  } else if (e.key==='ArrowUp') {
    e.preventDefault();
    const firstRowId = flat[0]?.rowId;
    const currentRowId = flat[kbIdx]?.rowId;
    if (currentRowId && currentRowId === firstRowId) {
      kbIdx = -1;
      renderAll();
      _focusTopNav(0);
      return;
    }
    const cur = flat[kbIdx];
    if (cur) {
      const prevRowItems = flat.filter(f => f.rowIndex === cur.rowIndex - 1);
      if (prevRowItems.length) {
        const preferred = prevRowItems.find(f => f.ei === cur.ei) || prevRowItems[Math.min(cur.ei, prevRowItems.length - 1)];
        const prevIdx = flat.findIndex(f => f.e.id === preferred.e.id);
        if (prevIdx !== -1) kbIdx = prevIdx;
      }
    }
    renderAll();
    scrollKbTarget('start');
  } else if (e.key==='Enter') {
    e.preventDefault();
    const selectedIdx = kbIdx;
    const f = flat[selectedIdx];
    if (f && f.e.text) {
      kbIdx = -1;
      renderAll();
      bridge.pasteEntry(f.e.id, f.e.text);
    }
  }
});

function onCardKey(e,entryId){
  if (e.key==='Enter'&&!e.shiftKey){e.preventDefault();saveCard(entryId);}
  if (e.key==='Escape'){editingId=null;renderAll();}
}

// ── Settings ──────────────────────────────────────────────────────────────────
function switchSettingsTab(tab){
  settingsTab = tab;
  document.querySelectorAll('.settings-tab').forEach(b => b.classList.toggle('active', b.id === 'stab-' + tab));
  document.querySelectorAll('.settings-panel').forEach(p => p.classList.toggle('active', p.dataset.stab === tab));
  settingsNavIdx = -1;
}

function toggleSettings(){
  settingsOpen=!settingsOpen;
  document.getElementById('settings').style.display=settingsOpen?'flex':'none';
  document.getElementById('settings-backdrop').style.display=settingsOpen?'block':'none';
  document.getElementById('gear-btn').classList.toggle('active',settingsOpen);
  if (settingsOpen) {
    switchSettingsTab(settingsTab || 'general');
    updateSettings();
    requestAnimationFrame(() => {
      const activeTab = document.querySelector('#settings .settings-tab.active');
      if (!activeTab) {
        _focusSettingsItem(0);
        return;
      }
      const idx = _settingsNavItems().indexOf(activeTab);
      if (idx >= 0) _focusSettingsItem(idx);
    });
  } else {
    settingsNavIdx = -1;
  }
}
function closeSettings(){
  settingsOpen=false;
  settingsNavIdx=-1;
  document.querySelectorAll('#settings .nav-focus').forEach(el => el.classList.remove('nav-focus'));
  document.getElementById('settings').style.display='none';
  document.getElementById('settings-backdrop').style.display='none';
  document.getElementById('gear-btn').classList.remove('active');
  const btn = document.getElementById('clear-all-btn');
  if (btn) {
    clearAllArmed = false;
    clearTimeout(clearAllConfirmTimer);
    btn.textContent = 'Clear All Data';
    btn.classList.remove('danger-btn');
  }
}

function updateSettings(){
  document.getElementById('startup-toggle').classList.toggle('on',launchAtStartup);
  document.getElementById('backup-toggle')?.classList.toggle('on',backupEnabled);
  document.getElementById('auto-toggle').classList.toggle('on',autoCapture);
  document.getElementById('hist-toggle').classList.toggle('on',histEnabled);
  const pollSlider = document.getElementById('poll-slider');
  if (pollSlider) pollSlider.value = String([250,500,1000].indexOf(pollRate));
  document.getElementById('paste-plain')?.classList.toggle('active', !!pastePlainText);
  document.getElementById('paste-rich')?.classList.toggle('active', !pastePlainText);
  const histSlider = document.getElementById('hist-limit-slider');
  if (histSlider) histSlider.value = String([5,10,20,50].indexOf(historyLimit));

  const backupStatus = document.getElementById('backup-status');
  if (backupStatus) {
    backupStatus.textContent = `Last backup: ${lastBackup || 'Never'}`;
  }

  const hk = document.getElementById('hotkey-input');
  if (hk) {
    hk.value = formatHotkey(hotkey);
    hk.dataset.value = hotkey;
  }
  document.getElementById('history-section').style.opacity=histEnabled?'1':'0.45';
  const hl=document.getElementById('history-list');
  const hc=document.getElementById('hist-clear');
  const histFiltered = (history || [])
    .filter(item => !historySearch || item.toLowerCase().includes(historySearch))
    .slice(0, historyLimit);
  if (!histEnabled||histFiltered.length===0){
    hl.innerHTML=`<div class="hist-empty">${histEnabled?'Nothing captured yet':'History is off'}</div>`;
    hc.style.display='none';
  } else {
    hl.innerHTML=histFiltered.map(item=>`<button class="hist-item" onclick="copyCard('__hist__','${escAttr(item)}')"><span class="hist-item-text">${escHtml(item)}</span><span style="font-size:9px;color:var(--textMuted)">copy</span></button>`).join('');
    hc.style.display='block';
  }
}

function normalizeHotkeyInput(v){
  const p=(v||'').toLowerCase().replace(/\s+/g,'').split('+').filter(Boolean);
  const out=[];
  p.forEach(k=>{
    if(k==='win'||k==='cmd'||k==='meta') k='windows';
    if(!out.includes(k)) out.push(k);
  });
  return out.join('+');
}
function isValidHotkey(v){
  const parts=v.split('+').filter(Boolean);
  if(parts.length<2||parts.length>4) return false;
  const mods=['ctrl','alt','shift','windows'];
  const hasMod=parts.some(p=>mods.includes(p));
  const hasKey=parts.some(p=>!mods.includes(p));
  return hasMod&&hasKey;
}
function formatHotkey(v){
  const map={ctrl:'Ctrl',alt:'Alt',shift:'Shift',windows:'Win',space:'Space'};
  return (v||'ctrl+d').split('+').filter(Boolean).map(p=>map[p]||p.toUpperCase()).join(' + ');
}
function comboFromEvent(e){
  const mods=[];
  if(e.ctrlKey) mods.push('ctrl');
  if(e.altKey) mods.push('alt');
  if(e.shiftKey) mods.push('shift');
  if(e.metaKey) mods.push('windows');
  const skip=['control','shift','alt','meta','os','win','windows'];
  let key=(e.key||'').toLowerCase();
  if(skip.includes(key)) return null;
  if(key===' ') key='space';
  if(!key) return null;
  return normalizeHotkeyInput([...mods,key].join('+'));
}
function onHotkeyInputKey(e){
  if(!hotkeyCaptureArmed){
    return;
  }
  if(e.key==='Tab') return;
  e.preventDefault();
  const combo=comboFromEvent(e);
  if(!combo) return;
  const i=document.getElementById('hotkey-input');
  i.dataset.value=combo;
  i.value=formatHotkey(combo);
}
function onHotkeyInputMouseDown(e){
  // Only a left mouse click can arm hotkey capture.
  hotkeyCaptureArmed = e.button === 0;
}
function saveHotkey(){
  const i=document.getElementById('hotkey-input');
  const combo=normalizeHotkeyInput(i?.dataset?.value||i?.value||'');
  if(!isValidHotkey(combo)){showToast('Use a valid combo like Ctrl + Alt + H','warn');return;}
  if(!bridge||!bridge.setHotkey){showToast('Hotkey bridge unavailable','warn');return;}
  bridge.setHotkey(combo,function(res){
    const d=typeof res==='string'?JSON.parse(res):res;
    if(d&&d.ok){
      hotkey=d.hotkey||combo;
      i.dataset.value=hotkey;
      i.value=formatHotkey(hotkey);
      showToast('Hotkey saved','ok');
    }else{
      showToast(d?.error||'Could not save hotkey','warn');
      i.dataset.value=hotkey;
      i.value=formatHotkey(hotkey);
    }
  });
}
function resetHotkey(){
  const i=document.getElementById('hotkey-input');
  if(!i) return;
  i.dataset.value='ctrl+d';
  i.value='Ctrl + D';
  saveHotkey();
}

function toggleStartup(){
  launchAtStartup = !launchAtStartup;
  bridge.setLaunchAtStartup(launchAtStartup);
  showToast(launchAtStartup ? 'Startup enabled' : 'Startup disabled', 'ok');
  updateSettings();
}

function toggleBackupEnabled(){
  backupEnabled = !backupEnabled;
  bridge.setBackupEnabled(backupEnabled);
  showToast(backupEnabled ? 'Auto backup enabled' : 'Auto backup disabled', 'ok');
  updateSettings();
}

function setPollRate(ms){
  if (pollRate === ms) return;
  pollRate = ms;
  bridge.setPollRate(ms);
  showToast(`Capture speed set to ${ms}ms`, 'ok');
  updateSettings();
}

function setPollRateFromSlider(idx){
  const opts = [250, 500, 1000];
  const ms = opts[Math.max(0, Math.min(2, Number(idx) || 0))];
  setPollRate(ms);
}

function setPasteMode(plain){
  pastePlainText = !!plain;
  bridge.setPastePlainText(pastePlainText);
  showToast(pastePlainText ? 'Paste mode: plain text' : 'Paste mode: preserve formatting', 'ok');
  updateSettings();
}

function setHistoryLimit(limit){
  if (historyLimit === limit) return;
  historyLimit = limit;
  bridge.setHistoryLimit(limit);
  showToast(`History size set to ${limit}`, 'ok');
  updateSettings();
}

function setHistoryLimitFromSlider(idx){
  const opts = [5, 10, 20, 50];
  const lim = opts[Math.max(0, Math.min(3, Number(idx) || 0))];
  setHistoryLimit(lim);
}

function onHistorySearch(value){
  historySearch = (value || '').trim().toLowerCase();
  updateSettings();
}

function clearAllDataConfirm(){
  const btn = document.getElementById('clear-all-btn');
  if (!btn) return;
  if (!clearAllArmed){
    clearAllArmed = true;
    btn.textContent = 'Tap again to confirm';
    btn.classList.add('danger-btn');
    clearTimeout(clearAllConfirmTimer);
    clearAllConfirmTimer = setTimeout(()=>{
      clearAllArmed = false;
      btn.textContent = 'Clear All Data';
      btn.classList.remove('danger-btn');
    }, 3000);
    return;
  }
  clearTimeout(clearAllConfirmTimer);
  clearAllArmed = false;
  btn.textContent = 'Clear All Data';
  btn.classList.remove('danger-btn');
  bridge.clearAllData();
}

function toggleCapture(){
  autoCapture=!autoCapture;
  bridge.setAutoCapture(autoCapture);
  showToast(autoCapture?'Auto-capture on':'Auto-capture off',autoCapture?'ok':'warn');
  updateSettings(); renderAll();
}
function toggleHistory(){
  histEnabled=!histEnabled;
  bridge.setHistoryEnabled(histEnabled);
  showToast(histEnabled?'History on':'History off',histEnabled?'ok':'warn');
  updateSettings();
}
function clearHistory(){ bridge.clearHistory(); }

// ── Drag & drop ───────────────────────────────────────────────────────────────
function onCardDragStart(e,rowId,entryId){dragCard={rowId,entryId};e.stopPropagation();}
function onCardDrop(e,targetRowId,targetEntryId){
  e.stopPropagation();
  if(!dragCard||dragCard.entryId===targetEntryId){dragCard=null;return;}
  bridge.moveEntry(dragCard.entryId,targetRowId,targetEntryId);dragCard=null;
}
function onRowDragStart(e,rowId){if(dragCard){e.preventDefault();return;}dragRow=rowId;document.getElementById('rowwrap-'+rowId)?.classList.add('dragging');}
function onRowDragOver(e,rowId){e.preventDefault();}
function onRowDrop(e,targetRowId){if(!dragRow||dragRow===targetRowId){dragRow=null;return;}bridge.reorderRow(dragRow,targetRowId);dragRow=null;}
function onRowDragEnd(){if(dragRow)document.getElementById('rowwrap-'+dragRow)?.classList.remove('dragging');dragRow=null;}

// ── Icons ─────────────────────────────────────────────────────────────────────
function penIco(){return'<svg width="9" height="9" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>';}
function copyIco(){return'<svg width="9" height="9" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>';}
function trashIco(){return'<svg width="9" height="9" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>';}
</script>
</body>
</html>"""


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    # Catch unhandled exceptions on *any* thread (e.g. ClipboardPoller)
    def _thread_excepthook(args):
        msg = "".join(traceback.format_exception(
            args.exc_type, args.exc_value, args.exc_traceback))
        log.critical("UNHANDLED EXCEPTION (thread %s):\n%s",
                     args.thread.name if args.thread else "?", msg)
    threading.excepthook = _thread_excepthook

    # Load persisted settings before startup registration so the user's
    # Launch-at-startup preference is respected on every app run.
    load()
    with _state_lock:
        launch_on_login = bool(state.get("launch_at_startup", True))
    _set_launch_at_startup(launch_on_login)

    # MUST be set before QApplication() — Qt reads these flags during
    # GPU process initialisation which happens inside QApplication.__init__.
    # Setting them after QApplication is created has no effect, causing
    # the "Failed to create GLES3 context" errors logged to stderr.
    os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = (
        "--disable-gpu --disable-software-rasterizer "
        "--disable-gpu-sandbox --no-sandbox --disable-dev-shm-usage"
    )
    os.environ.setdefault("QTWEBENGINE_DISABLE_SANDBOX", "1")
    log.info("QTWEBENGINE_CHROMIUM_FLAGS set before QApplication()")

    app = QApplication(sys.argv)
    app.setApplicationName("Clippy")
    app.setQuitOnLastWindowClosed(False)

    try:
        win = ClippyWindow()
        # Do NOT call win.show() — Clippy starts hidden in the system tray.
        # The tray balloon notification tells the user it is running.
        # The window only appears when the user presses Ctrl+D or double-clicks the tray icon.
        log.info("Entering Qt event loop — window starts hidden in tray")
        exit_code = app.exec()
        log.info("Qt event loop exited with code %d", exit_code)
        sys.exit(exit_code)
    except Exception as exc:
        _log_exc("__main__ top-level", exc)
        sys.exit(1)
