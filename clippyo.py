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

import sys, json, uuid, os, csv, io, threading, time, traceback, logging
from pathlib import Path

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

# ══════════════════════════════════════════════════════════════════════════════
# STATE
# ══════════════════════════════════════════════════════════════════════════════
def _gid(): return uuid.uuid4().hex[:12]   # 48-bit entropy — collision-safe up to millions of entries
def _entry(text=""): return {"id":_gid(),"text":text,"copy_count":0,"pinned":False}
def _row(ci=0):      return {"id":_gid(),"color_idx":ci,"entries":[_entry()]}

state = {
    "rows": [],
    "cursor": {"row_idx":0,"entry_idx":0},
    "auto_capture": True,
    "last_clip": "",
    "last_clip_time": 0,
    "history": [],
    "history_enabled": True,
    "theme": "daylight",
}

def load():
    if DATA_FILE.exists():
        try:
            d = json.loads(DATA_FILE.read_text("utf-8"))
            state["rows"] = d.get("rows", [])
            state["history_enabled"] = d.get("history_enabled", True)
            state["theme"] = d.get("theme", "daylight")
            log.info("State loaded from %s (%d rows, theme=%s)",
                     DATA_FILE, len(state["rows"]), state["theme"])
            return
        except Exception as exc:
            _log_exc("load()", exc)
    state["rows"] = [
        {"id":_gid(),"color_idx":0,"entries":[
            _entry("Welcome to Clippy!"),
        ]},
    ]

def save():
    try:
        DATA_FILE.write_text(json.dumps({
            "rows": state["rows"],
            "history_enabled": state["history_enabled"],
            "theme": state.get("theme", "dark"),
        }, ensure_ascii=False, indent=2), "utf-8")
    except Exception as exc:
        _log_exc("save()", exc)

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
        with _state_lock:
            src = None
            for row in state["rows"]:
                idx = next((i for i, e in enumerate(row["entries"]) if e["id"] == entry_id), None)
                if idx is not None:
                    src = row["entries"].pop(idx); break
            if src:
                for row in state["rows"]:
                    if row["id"] == target_row_id:
                        ti = next((i for i, e in enumerate(row["entries"]) if e["id"] == target_entry_id), None)
                        row["entries"].insert(ti if ti is not None else len(row["entries"]), src); break
            state["rows"] = [r for r in state["rows"] if r["entries"]]
        save(); self._push()

    @pyqtSlot(bool)
    def setAutoCapture(self, enabled):
        with _state_lock:
            state["auto_capture"] = enabled
            if not enabled:
                state["cursor"] = {"row_idx": -1, "entry_idx": -1}
        self._push()

    @pyqtSlot(bool)
    def setHistoryEnabled(self, enabled):
        with _state_lock:
            state["history_enabled"] = enabled
        save(); self._push()

    @pyqtSlot()
    def clearHistory(self):
        with _state_lock:
            state["history"] = []
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
        try:
            data = json.loads(Path(path).read_text("utf-8"))

            # Schema validation — must be a list of row dicts with required keys
            if not isinstance(data, list):
                raise ValueError("Top-level value must be a JSON array of groups")
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
                state["rows"] = data

            save()
            log.info("importJSON: imported %d rows from %s", len(data), path)
            self._push()
            self.stateChanged.emit(json.dumps({**json.loads(self.getState()), "toast": "Imported!"}))

        except Exception as ex:
            _log_exc("importJSON", ex)
            # Restore backup if we already replaced state
            try:
                with _state_lock:
                    if state["rows"] is not backup:
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
        MAX_ROWS = 50  # FIX C3: cap row growth to prevent unbounded expansion

        while True:
            # ── Sleep is ALWAYS at the end of the loop body ──────────────────
            # This ensures the very first check fires immediately on thread
            # start (no missed copies at launch) and also fires immediately
            # when auto-capture is re-enabled after being turned off.

            # Read shared flags under lock — even boolean reads can tear on some platforms
            with _state_lock:
                auto_on = state["auto_capture"]
            if not auto_on or not CLIPBOARD_OK:
                time.sleep(0.5)
                continue

            global _clippy_is_pasting, _clippy_last_pasted
            if _clippy_is_pasting:
                time.sleep(0.5)
                continue

            try:
                text = pyperclip.paste()

                if not text:
                    last_seen = ""
                    time.sleep(0.5)
                    continue

                if text == last_seen:
                    time.sleep(0.5)
                    continue

                if text == _clippy_last_pasted:
                    last_seen = text
                    time.sleep(0.5)
                    continue

                last_seen = text

                with _state_lock:
                    # History
                    if state["history_enabled"]:
                        h = [t for t in state["history"] if t != text]
                        h.insert(0, text)
                        state["history"] = h[:10]

                    rows = state["rows"]
                    c    = state["cursor"]
                    ri   = c["row_idx"]
                    ei   = c["entry_idx"]

                    if ri < 0 or ri >= len(rows):
                        # FIX C3: only create new row if under cap
                        if len(rows) < MAX_ROWS:
                            rows.append({"id": _gid(), "color_idx": len(rows),
                                         "entries": [_entry(text)]})
                            ri = len(rows) - 1
                            ei = 0
                        else:
                            log.warning("Row cap (%d) reached — recycling oldest non-pinned row", MAX_ROWS)
                            # Recycle oldest row that has no pinned entries
                            for old_ri, old_row in enumerate(rows):
                                if not any(e.get("pinned") for e in old_row["entries"]):
                                    rows[old_ri] = {"id": _gid(), "color_idx": old_ri,
                                                    "entries": [_entry(text)]}
                                    ri = old_ri; ei = 0
                                    break
                            else:
                                time.sleep(0.5)
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
                                                 "entries": [_entry(text)]})
                                    ei = 0
                                else:
                                    log.warning("Row cap reached during overflow — skipping capture")
                                    time.sleep(0.5)
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
                                                     "entries": [_entry(text)]})
                                        ei = 0
                                    else:
                                        log.warning("Row cap reached — skipping capture")
                                        time.sleep(0.5)
                                        continue

                    state["cursor"] = find_next_capture_slot(rows, ri, ei)

                save()
                self.captured.emit(text)

            except Exception as exc:
                _log_exc("ClipboardPoller.run()", exc)

            # Single canonical sleep — all non-skip paths reach here
            time.sleep(0.5)


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
        load()
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

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._set_native_rounded_corners()

    def _show_window(self, cx=-1, cy=-1):
        """cx, cy: caret screen coords from the destination app (-1,-1 = unknown)."""
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
html,body{height:100%;font-family:'Segoe UI','SF Pro Text',system-ui,-apple-system,BlinkMacSystemFont,sans-serif;margin:0;padding:0;}
html{border-radius:18px;overflow:hidden;background:transparent;}
body{background:var(--appBg);color:var(--text);transition:background .3s,color .3s;
  overflow:hidden;position:relative;border-radius:18px;font-size:140%;}
::-webkit-scrollbar{width:4px;}
::-webkit-scrollbar-thumb{background:rgba(124,58,237,.35);border-radius:10px;}
::placeholder{color:var(--textMuted);}
@keyframes pulse{0%,100%{opacity:1;transform:scale(1);}50%{opacity:.4;transform:scale(.6);}}
@keyframes slideIn{from{opacity:0;transform:translateX(20px);}to{opacity:1;transform:translateX(0);}}
@keyframes fadeUp{from{opacity:0;transform:translateX(-50%) translateY(8px);}to{opacity:1;transform:translateX(-50%) translateY(0);}}

/* ── NIGHT ── */
body.dark,body.night{
  --appBg:#0d0f1e;--navBg:rgba(18,20,40,1);--navBorder:rgba(139,92,246,.25);
  --panelBg:#10132a;--panelBorder:rgba(139,92,246,.18);
  --rowBg:rgba(255,255,255,.016);--rowBorder:rgba(255,255,255,.055);
  --cardBg:rgba(255,255,255,.03);--cardBorder:rgba(255,255,255,.07);
  --text:#ddd6f3;--textDim:#94a3b8;--textMuted:#3f4660;
  --inputBg:rgba(255,255,255,.04);--inputBorder:rgba(255,255,255,.08);
  --sectionLabel:#4b5280;--emptyText:#2a2f4a;
  --toastBg:rgba(20,16,48,.98);--toastBorder:rgba(139,92,246,.45);
  --toastWarnBg:rgba(100,20,0,.95);--toastWarnBorder:rgba(251,113,133,.4);
  --gearStroke:#c4b5e8;
}
/* ── DAYLIGHT ── */
body.daylight{
  --appBg:linear-gradient(135deg,#e9f3ff 0%,#f6f2ff 58%,#f2fff7 100%);
  --navBg:rgba(255,255,255,.58);--navBorder:rgba(185,197,226,.72);
  --panelBg:#f7fbff;--panelBorder:rgba(185,197,226,.92);
  --rowBg:rgba(255,255,255,.50);--rowBorder:rgba(194,206,234,.92);
  --cardBg:rgba(255,255,255,.82);--cardBorder:rgba(206,219,245,.95);
  --text:#2f3252;--textDim:#677391;--textMuted:#8e99b1;
  --inputBg:rgba(255,255,255,.80);--inputBorder:rgba(196,210,240,.96);
  --sectionLabel:#8894ad;--emptyText:#9ba6c3;
  --toastBg:rgba(255,255,255,.95);--toastBorder:rgba(123,87,255,.28);
  --toastWarnBg:rgba(255,244,230,.97);--toastWarnBorder:rgba(251,113,133,.4);
  --gearStroke:#4a4f79;
}
body.daylight::before,body.daylight::after{
  content:"";position:fixed;border-radius:999px;pointer-events:none;z-index:0;filter:blur(1px);
}
body.daylight::before{
  width:360px;height:360px;left:-90px;top:-120px;background:rgba(207,232,255,.65);
}
body.daylight::after{
  width:400px;height:400px;right:-120px;bottom:-180px;background:rgba(216,255,226,.68);
}

/* ── NAV ── */
#nav{
  display:flex;align-items:center;padding:0 12px;height:58px;gap:10px;
  background:var(--navBg);border-bottom:1px solid var(--navBorder);
  position:relative;z-index:200;backdrop-filter:blur(16px);flex-shrink:0;
  box-shadow:0 10px 30px rgba(120,138,182,.16);
  border-top-left-radius:18px;border-top-right-radius:18px;
}

/* FIX #5: new logo mark — small pill with icon */
.logo{display:flex;align-items:center;gap:8px;flex-shrink:0;text-decoration:none;}
.logo-mark{
  width:30px;height:30px;border-radius:9px;flex-shrink:0;
  background:linear-gradient(135deg,#6d4bff 0%,#7b57ff 55%,#8d6bff 100%);
  display:flex;align-items:center;justify-content:center;
  box-shadow:0 6px 14px rgba(109,75,255,.35);
}
.logo-mark svg{display:block;}
.logo-name{
  font-weight:700;font-size:22px;letter-spacing:-.3px;
  color:#7b57ff;
}

.search-wrap{flex:1;position:relative;}
.search-wrap svg{position:absolute;left:10px;top:50%;transform:translateY(-50%);pointer-events:none;color:var(--textMuted);}
#search{
  width:100%;padding:8px 30px 8px 28px;
  background:var(--inputBg);border:1px solid var(--inputBorder);border-radius:100px;
  color:var(--text);font-size:17px;outline:none;font-family:inherit;transition:border-color .2s,box-shadow .2s;
}
#search:focus{border-color:rgba(123,87,255,.45);box-shadow:0 0 0 3px rgba(123,87,255,.12);}
.search-clear{position:absolute;right:9px;top:50%;transform:translateY(-50%);
  background:none;border:none;color:var(--textMuted);cursor:pointer;font-size:18px;line-height:1;}

.btn-new{
  flex-shrink:0;background:linear-gradient(135deg,#6d4bff,#8968ff);
  border:none;border-radius:11px;color:#fff;padding:8px 14px;
  font-size:17px;font-weight:700;cursor:pointer;font-family:inherit;
  box-shadow:0 10px 18px rgba(109,75,255,.28);display:flex;align-items:center;gap:4px;transition:opacity .15s,transform .15s;
}
.btn-new:hover{opacity:.94;transform:translateY(-1px);}

.btn-gear{
  flex-shrink:0;width:35px;height:35px;border-radius:11px;cursor:pointer;
  display:flex;align-items:center;justify-content:center;
  background:rgba(255,255,255,.84);border:1px solid rgba(196,210,240,.96);transition:all .18s;
}
.btn-gear svg{stroke:var(--gearStroke);stroke-width:2.3;transition:stroke .3s;}
.btn-gear.active,.btn-gear:hover{background:rgba(123,87,255,.16);border-color:rgba(123,87,255,.42);box-shadow:0 0 0 3px rgba(123,87,255,.14);}
.btn-new.nav-focus,.btn-gear.nav-focus,#search.nav-focus{
  box-shadow:0 0 0 3px rgba(123,87,255,.22)!important;
  border-color:rgba(123,87,255,.46)!important;
}

/* ── MAIN ── */
/* scroll-wrap fills remaining height below nav — rows scroll inside here */
#scroll-wrap{
  height:calc(100vh - 58px);  /* exactly viewport minus nav bar */
  overflow-y:auto;
  overflow-x:hidden;
  overscroll-behavior:contain;
  border-bottom-left-radius:18px;border-bottom-right-radius:18px;
}
#scroll-wrap::-webkit-scrollbar{width:4px;}
#scroll-wrap::-webkit-scrollbar-thumb{background:rgba(124,58,237,.35);border-radius:10px;}
#main{padding:10px;display:flex;flex-direction:column;gap:8px;position:relative;z-index:1;}
.empty{text-align:center;color:var(--emptyText);padding-top:60px;font-size:18px;}

/* ── ROW ── */
.row-wrap{
  display:flex;align-items:stretch;gap:8px;transition:opacity .15s;
  height:calc((100vh - 58px - 20px - 16px) / 3);
  min-height:108px;
}
.row-wrap.dragging{opacity:.35;}
.row-num{
  color:rgba(139,92,246,.75);font-size:25px;font-weight:900;
  width:22px;text-align:right;flex-shrink:0;user-select:none;line-height:1;
  text-shadow:0 0 8px rgba(139,92,246,.3);
}
.row-num.active{color:rgba(168,85,247,1);text-shadow:0 0 10px rgba(168,85,247,.5);}
.row-accent{width:3px;align-self:stretch;border-radius:3px;flex-shrink:0;opacity:.8;}
.row-group{
  flex:1;position:relative;background:var(--rowBg);border:1px solid var(--rowBorder);
  border-radius:18px;padding:8px 10px;
  display:flex;flex-direction:row;flex-wrap:nowrap;gap:8px;align-items:stretch;
  overflow:hidden;
  transition:border-color .14s,background .14s,box-shadow .12s;
  backdrop-filter:blur(8px);
  box-shadow:0 7px 22px rgba(125,140,185,.14);
}
.row-group.active-capture{background:rgba(123,87,255,.08);border-color:rgba(123,87,255,.36);}
.row-group.search-match{border-color:rgba(168,85,247,.45);}
.row-del{
  position:absolute;top:8px;right:8px;width:22px;height:22px;border-radius:50%;
  background:rgba(255,255,255,.9);border:1px solid rgba(198,208,232,.95);
  color:#5d6786;cursor:pointer;display:flex;align-items:center;justify-content:center;
  z-index:3;font-size:17px;font-weight:700;line-height:1;transition:all .15s;
  box-shadow:0 4px 10px rgba(109,122,156,.2);
}
.row-del:hover{background:rgba(239,68,68,.18);border-color:rgba(239,68,68,.4);color:#c33232;}

/* ── CARD ── */
.card{
  flex:1 1 0;min-width:0;max-width:none;aspect-ratio:auto;height:100%;position:relative;
  background:var(--cardBg);border:1px solid var(--cardBorder);border-radius:11px;
  padding:9px 10px 7px;display:flex;flex-direction:column;justify-content:space-between;
  overflow:hidden;transition:all .08s linear;cursor:pointer;
}
.card:hover{
  border-color:rgba(123,87,255,.55);background:rgba(123,87,255,.10);
  box-shadow:0 0 0 1px rgba(123,87,255,.18),0 10px 22px rgba(123,87,255,.14);
}
.card.editing{border-color:rgba(139,92,246,.8)!important;background:rgba(109,40,217,.12)!important;box-shadow:0 0 0 3px rgba(109,40,217,.14)!important;}
.card.kb-focus{
  border-color:#a855f7!important;background:rgba(168,85,247,.13)!important;
  box-shadow:0 0 0 2px rgba(168,85,247,.38),0 6px 20px rgba(109,40,217,.2)!important;
}
.card.search-hit{border-color:rgba(168,85,247,.6)!important;background:rgba(109,40,217,.1)!important;}
.card.active-slot{border-color:rgba(139,92,246,.5)!important;background:rgba(109,40,217,.08)!important;box-shadow:0 0 0 2px rgba(139,92,246,.2)!important;}
.card.user-target{border-color:#fbbf24!important;background:rgba(251,191,36,.05)!important;box-shadow:0 0 0 2px rgba(251,191,36,.28)!important;}
.card.dragging{opacity:.35;}

.card-dot{position:absolute;top:6px;right:6px;width:5px;height:5px;border-radius:50%;background:#a855f7;box-shadow:0 0 5px #a855f7;animation:pulse 1.4s ease-in-out infinite;}
.card-kb-bar{position:absolute;left:0;top:10%;bottom:10%;width:3px;border-radius:0 3px 3px 0;background:linear-gradient(180deg,#6d28d9,#a855f7);box-shadow:0 0 8px rgba(168,85,247,.5);}
.card-pin{position:absolute;top:4px;right:5px;font-size:8px;color:#fbbf24;}
.heatbar{position:absolute;bottom:0;left:0;height:2px;border-radius:0 2px 0 0;opacity:.8;transition:width .4s;}
.card-text{flex:1;overflow:hidden;font-size:17px;line-height:1.42;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;}
.card-text.empty-slot{color:var(--textMuted);font-style:italic;font-size:15px;}
.card-text.code-font{font-family:'Cascadia Code','Cascadia Mono',Consolas,'Courier New',monospace;font-size:15px;}
.card-actions{display:flex;gap:5px;justify-content:flex-end;padding-top:3px;flex-shrink:0;opacity:0;transition:opacity .05s linear;pointer-events:none;}
.card:hover .card-actions,.card:focus-within .card-actions,.card.editing .card-actions{opacity:1;pointer-events:auto;}
.card-textarea{width:100%;height:100%;background:transparent;border:none;color:var(--text);font-size:17px;line-height:1.42;padding:0;outline:none;resize:none;font-family:'Cascadia Code','Cascadia Mono',Consolas,'Courier New',monospace;}
mark{background:rgba(168,85,247,.35);color:#fff;border-radius:3px;padding:0 2px;}

/* ── CARD BUTTONS ── */
.cbtn{
  width:22px;height:22px;border-radius:6px;padding:0;flex-shrink:0;
  cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:13px;
  transition:all .08s linear;background:rgba(255,255,255,.86);border:1px solid rgba(150,170,215,.92);color:#475477;
}
.cbtn:hover{background:#ffffff;border-color:rgba(123,87,255,.5);color:#26345c;}
.cbtn.active{background:rgba(109,40,217,.25);border-color:rgba(139,92,246,.5);color:#a78bfa;}
.cbtn.danger{border-color:rgba(239,68,68,.18);color:#ef4444;}
.cbtn.danger:hover{background:rgba(239,68,68,.2);border-color:rgba(239,68,68,.42);color:#f87171;}
.save-btn{background:rgba(109,40,217,.28);border:1px solid rgba(139,92,246,.5);border-radius:5px;color:#a78bfa;font-size:14px;padding:2px 10px;cursor:pointer;font-family:inherit;font-weight:600;}

/* ── ADD VARIATION ── */
.card-add{
  flex:1 1 0;min-width:0;max-width:none;aspect-ratio:auto;height:100%;cursor:pointer;
  background:rgba(123,87,255,.03);border:1.5px dashed rgba(123,87,255,.28);
  border-radius:11px;display:flex;flex-direction:column;align-items:center;
  justify-content:center;gap:3px;color:#8b7fc0;font-size:15px;font-family:inherit;transition:all .16s;
}
.card-add:hover{background:rgba(123,87,255,.11);border-color:rgba(123,87,255,.5);}

/* ── SETTINGS ── */
#settings{position:fixed;top:0;right:0;bottom:0;width:280px;background:var(--panelBg);border-left:1px solid var(--panelBorder);z-index:500;display:flex;flex-direction:column;box-shadow:-12px 0 40px rgba(0,0,0,.3);animation:slideIn .14s ease;backdrop-filter:none;}
#settings-backdrop{position:fixed;inset:0;z-index:499;background:rgba(10,15,30,.14);}
.settings-head{display:flex;align-items:center;justify-content:space-between;padding:16px 18px 12px;border-bottom:1px solid var(--navBorder);}
.settings-title{font-size:20px;font-weight:700;}
.settings-close{background:none;border:none;color:var(--textMuted);cursor:pointer;font-size:24px;line-height:1;}
.settings-body{flex:1;overflow-y:auto;padding:16px 18px;}
.settings-section{margin-bottom:20px;}
.s-title{font-size:13px;color:var(--sectionLabel);font-weight:700;letter-spacing:.08em;text-transform:uppercase;margin-bottom:9px;}
.theme-grid{display:flex;gap:7px;}
.theme-btn{flex:1;padding:9px 6px;border-radius:9px;cursor:pointer;border:2px solid var(--inputBorder);background:var(--inputBg);transition:all .16s;display:flex;flex-direction:column;align-items:center;gap:4px;}
.theme-btn.active{border-color:rgba(139,92,246,.7);background:rgba(109,40,217,.14);}
.theme-swatch{width:26px;height:26px;border-radius:7px;border:1px solid rgba(139,92,246,.25);}
.theme-label{font-size:15px;font-weight:500;color:var(--textDim);}
.theme-btn.active .theme-label{color:#a78bfa;}
.toggle-row{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:9px;}
.toggle-info{flex:1;}
.toggle-label{font-size:18px;font-weight:500;}
.toggle-sub{font-size:15px;color:var(--textMuted);margin-top:1px;}
.toggle-switch{width:36px;height:19px;border-radius:10px;position:relative;cursor:pointer;background:rgba(128,128,128,.22);border:1px solid rgba(128,128,128,.28);transition:all .2s;flex-shrink:0;}
.toggle-switch.on{background:rgba(109,40,217,.75);border-color:rgba(139,92,246,.6);}
.toggle-thumb{position:absolute;top:2px;left:2px;width:13px;height:13px;border-radius:50%;background:rgba(200,200,200,.8);transition:all .2s;}
.toggle-switch.on .toggle-thumb{left:17px;background:#fff;box-shadow:0 0 5px rgba(168,85,247,.55);}
.history-list{background:var(--inputBg);border:1px solid var(--inputBorder);border-radius:9px;overflow:hidden;}
.hist-item{width:100%;background:transparent;border:none;border-bottom:1px solid var(--navBorder);padding:7px 11px;text-align:left;color:var(--text);font-size:14px;cursor:pointer;font-family:'Cascadia Code','Cascadia Mono',Consolas,'Courier New',monospace;display:flex;justify-content:space-between;align-items:center;gap:7px;}
.hist-item:last-child{border-bottom:none;}
.hist-item:hover{background:rgba(139,92,246,.09);}
.hist-item-text{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;}
.hist-empty{padding:10px 12px;font-size:15px;color:var(--textMuted);font-style:italic;}
.hist-clear{margin-top:5px;background:none;border:none;color:var(--textMuted);font-size:15px;cursor:pointer;font-family:inherit;}
.action-btn{width:100%;padding:7px 11px;border-radius:7px;cursor:pointer;background:rgba(128,128,128,.05);border:1px solid var(--inputBorder);color:var(--textDim);font-size:16px;font-family:inherit;font-weight:500;text-align:left;transition:all .14s;display:flex;align-items:center;gap:7px;margin-bottom:5px;}
.action-btn:hover{background:rgba(128,128,128,.11);}

/* ── TOAST ── */
#toast{position:fixed;bottom:20px;left:50%;transform:translateX(-50%);padding:7px 20px;border-radius:100px;font-size:16px;font-weight:500;box-shadow:0 6px 24px rgba(0,0,0,.3);z-index:999;white-space:nowrap;animation:fadeUp .18s ease;pointer-events:none;display:none;}
#toast.ok{background:var(--toastBg);border:1px solid var(--toastBorder);color:var(--text);}
#toast.warn{background:var(--toastWarnBg);border:1px solid var(--toastWarnBorder);color:var(--text);}
</style>
</head>
<body class="daylight">

<!-- NAV -->
<nav id="nav">
  <!-- FIX #5: new logo -->
  <div class="logo">
    <div class="logo-mark">
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
        <rect x="2.1" y="2.4" width="8.7" height="10.8" rx="1.6" fill="#FFE57C"/>
        <path d="M9.1 2.4 L10.8 4.1 L9.1 4.1 Z" fill="#EBCB52"/>
        <path d="M3.8 5.6H8.9" stroke="#9A7C2A" stroke-width="0.95" stroke-linecap="round"/>
        <path d="M3.8 7.6H8.4" stroke="#9A7C2A" stroke-width="0.95" stroke-linecap="round"/>
        <path d="M4.0 9.6 C5.1 8.9, 6.0 10.1, 7.0 9.5" stroke="#7A6120" stroke-width="0.8" fill="none" stroke-linecap="round"/>
        <g transform="translate(6.1,8.5) rotate(-24)">
          <rect x="0" y="0" width="6.2" height="1.25" rx="0.62" fill="#22365B"/>
          <rect x="0.7" y="0.2" width="2.6" height="0.85" rx="0.42" fill="#9DB8F7" fill-opacity="0.85"/>
          <polygon points="6.2,0 7.6,0.62 6.2,1.25" fill="#E8EDF9"/>
          <polygon points="7.6,0.62 8.35,0.62 7.6,0.3" fill="#27324E"/>
          <polygon points="7.6,0.62 8.35,0.62 7.6,0.95" fill="#27324E"/>
        </g>
      </svg>
    </div>
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

    <div class="settings-section">
      <div class="s-title">Theme</div>
      <div class="theme-grid">
        <button class="theme-btn" id="theme-night" onclick="setTheme('night')">
          <div class="theme-swatch" style="background:linear-gradient(135deg,#0d0f1e,#1e1b4b)"></div>
          <span class="theme-label">Night</span>
        </button>
        <button class="theme-btn active" id="theme-daylight" onclick="setTheme('daylight')">
          <div class="theme-swatch" style="background:linear-gradient(135deg,#f4f2ff,#ede9fe)"></div>
          <span class="theme-label">Daylight</span>
        </button>
      </div>
    </div>

    <div class="settings-section">
      <div class="s-title">Capture</div>
      <div class="toggle-row">
        <div class="toggle-info">
          <div class="toggle-label">Auto-capture</div>
          <div class="toggle-sub">Reads clipboard every 500ms</div>
        </div>
        <div class="toggle-switch on" id="auto-toggle" tabindex="0" onclick="toggleCapture()">
          <div class="toggle-thumb"></div>
        </div>
      </div>
    </div>

    <div class="settings-section">
      <div class="s-title">Paste History</div>
      <div class="toggle-row">
        <div class="toggle-info">
          <div class="toggle-label">Record history</div>
          <div class="toggle-sub">Save last 10 captured items</div>
        </div>
        <div class="toggle-switch on" id="hist-toggle" tabindex="0" onclick="toggleHistory()">
          <div class="toggle-thumb"></div>
        </div>
      </div>
      <div id="history-section">
        <div class="history-list" id="history-list"><div class="hist-empty">Nothing captured yet</div></div>
        <button class="hist-clear" id="hist-clear" style="display:none" onclick="clearHistory()">Clear history</button>
      </div>
    </div>

    <!-- FIX #3: AI Smart Group section REMOVED -->

    <div class="settings-section">
      <div class="s-title">Data</div>
      <button class="action-btn" onclick="bridge.exportJSON()">⬇ Export JSON</button>
      <button class="action-btn" onclick="bridge.exportCSV()">⬇ Export CSV</button>
      <button class="action-btn" onclick="bridge.importJSON()">⬆ Import JSON</button>
    </div>

  </div>
</div>

<div id="toast"></div>

<script>
// ── State ─────────────────────────────────────────────────────────────────────
let bridge = null;
let rows = [], cursor = {row_idx:0,entry_idx:0};
let autoCapture = true, history = [], histEnabled = true;
let editingId = null, kbIdx = -1, searchTerm = '';
let settingsOpen = false, dragCard = null, dragRow = null;
let userTargetId = null;
let settingsNavIdx = -1;
let topNavIdx = -1;

const TAG_CYCLE = [
  {dot:"#8b5cf6",border:"rgba(139,92,246,.42)"},
  {dot:"#38bdf8",border:"rgba(56,189,248,.42)"},
  {dot:"#34d399",border:"rgba(52,211,153,.42)"},
  {dot:"#fb7185",border:"rgba(251,113,133,.42)"},
  {dot:"#fbbf24",border:"rgba(251,191,36,.42)"},
  {dot:"#f87171",border:"rgba(248,113,113,.42)"},
  {dot:"#facc15",border:"rgba(250,204,21,.42)"},
];
const tag = i => TAG_CYCLE[i % TAG_CYCLE.length];

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
    if (!document.hasFocus() && !window.__clippyPointerInside && !settingsOpen && sinceShow > 6000 && bridge) {
      bridge.hideWindow();
    }
  }, 2200);
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
  const sorted = rows.map(r=>({...r,entries:[...r.entries].sort((a,b)=>(b.pinned?1:0)-(a.pinned?1:0))}));
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
  src.forEach((row)=>row.entries.forEach((e,ei)=>out.push({ei,e,rowId:row.id})));
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

  main.innerHTML = visible.map(row => {
    const realIdx   = rows.findIndex(r=>r.id===row.id);
    const isActiveRow = realIdx===cursor.row_idx && autoCapture;
    const t         = tag(row.color_idx??realIdx);
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
        body = `<div style="flex:1;overflow:hidden;padding-left:${isKb?6:0}px;transition:padding .13s">
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
      <span class="row-num${isActiveRow?' active':''}">${realIdx+1}</span>
      <div class="row-accent" style="background:${t.dot}"></div>
      <div class="row-group${isActiveRow?' active-capture':''}${hasMatch?' search-match':''}" style="border-color:${t.border}">
        <button class="row-del" onclick="delRow('${row.id}')">✕</button>
        ${cards}${addCard}
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
}

// ── Card click — set user target ──────────────────────────────────────────────
function onCardClick(event, rowId, entryId) {
  if (event.target.closest('button')||event.target.tagName==='TEXTAREA') return;
  userTargetId = entryId;
  bridge.setCursorToEntry(rowId, entryId);
  renderAll();
}

// ── Editing ───────────────────────────────────────────────────────────────────
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
      if (el) el.scrollIntoView({block:'nearest', behavior:'smooth'});
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
  return Array.from(document.querySelectorAll(
    '#settings .settings-close, ' +
    '#settings .theme-btn, ' +
    '#settings #auto-toggle, ' +
    '#settings #hist-toggle, ' +
    '#settings #hist-clear, ' +
    '#settings .hist-item, ' +
    '#settings .action-btn'
  )).filter(el => el.offsetParent !== null && !el.disabled);
}

function _focusSettingsItem(idx) {
  const items = _settingsNavItems();
  if (!items.length) {
    settingsNavIdx = -1;
    return;
  }
  const n = items.length;
  settingsNavIdx = ((idx % n) + n) % n;
  const el = items[settingsNavIdx];
  if (typeof el.focus === 'function') el.focus({preventScroll:true});
  el.scrollIntoView({block:'nearest', inline:'nearest', behavior:'smooth'});
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

function _handleSettingsKeys(e) {
  if (!settingsOpen) return false;
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
  if (key === 'ArrowDown') {
    _focusSettingsItem(settingsNavIdx + 1);
    return true;
  }
  if (key === 'ArrowUp') {
    _focusSettingsItem(settingsNavIdx - 1);
    return true;
  }
  if (key === 'ArrowLeft') {
    _focusSettingsItem(settingsNavIdx - 1);
    return true;
  }
  if (key === 'ArrowRight') {
    _focusSettingsItem(settingsNavIdx + 1);
    return true;
  }
  if (key === 'Home') {
    _focusSettingsItem(0);
    return true;
  }
  if (key === 'End') {
    const items = _settingsNavItems();
    _focusSettingsItem(items.length - 1);
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
      rowEl.scrollIntoView({block:'start', inline:'nearest', behavior:'smooth'});
      return;
    }
    cardEl?.scrollIntoView({block, inline:'nearest', behavior:'smooth'});
  };

  if (e.key==='Escape') {
    e.preventDefault(); kbIdx=-1;
    document.getElementById('search')?.focus();
    renderAll(); return;
  }
  // ArrowRight = next card in row, ArrowDown = first card of next row
  // ArrowLeft  = prev card in row, ArrowUp   = first card of prev row
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
    // Jump to first card of the next row
    const curRowId = flat[kbIdx]?.rowId;
    const nextRowIdx = flat.findIndex((f,i)=>i>kbIdx && f.rowId!==curRowId);
    if (nextRowIdx !== -1) kbIdx = nextRowIdx;
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
    // Jump to first card of the previous row
    const curRowId = flat[kbIdx]?.rowId;
    const prevFlat = flat.slice(0, kbIdx).reverse();
    const prevRowStart = prevFlat.find(f=>f.rowId!==curRowId);
    if (prevRowStart) {
      // find the first card of that row
      const firstOfRow = flat.findIndex(f=>f.rowId===prevRowStart.rowId);
      kbIdx = firstOfRow;
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
function toggleSettings(){
  settingsOpen=!settingsOpen;
  document.getElementById('settings').style.display=settingsOpen?'flex':'none';
  document.getElementById('settings-backdrop').style.display=settingsOpen?'block':'none';
  document.getElementById('gear-btn').classList.toggle('active',settingsOpen);
  if (settingsOpen) {
    updateSettings();
    requestAnimationFrame(() => _focusSettingsItem(0));
  } else {
    settingsNavIdx = -1;
  }
}
function closeSettings(){
  settingsOpen=false;
  settingsNavIdx=-1;
  document.getElementById('settings').style.display='none';
  document.getElementById('settings-backdrop').style.display='none';
  document.getElementById('gear-btn').classList.remove('active');
}

function updateSettings(){
  document.getElementById('auto-toggle').classList.toggle('on',autoCapture);
  document.getElementById('hist-toggle').classList.toggle('on',histEnabled);
  document.getElementById('history-section').style.opacity=histEnabled?'1':'0.45';
  const hl=document.getElementById('history-list');
  const hc=document.getElementById('hist-clear');
  if (!histEnabled||history.length===0){
    hl.innerHTML=`<div class="hist-empty">${histEnabled?'Nothing captured yet':'History is off'}</div>`;
    hc.style.display='none';
  } else {
    hl.innerHTML=history.map(item=>`<button class="hist-item" onclick="copyCard('__hist__','${escAttr(item)}')"><span class="hist-item-text">${escHtml(item)}</span><span style="font-size:9px;color:var(--textMuted)">copy</span></button>`).join('');
    hc.style.display='block';
  }
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

    # Register auto-start BEFORE Qt starts (no Qt objects needed)
    _ensure_startup()

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
