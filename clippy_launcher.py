"""
clippy_launcher.py — instant startup shim for Clippy.

This script is what the Startup folder / registry run-key launches.
It has NO heavy imports (no Qt, no Chromium, no PyQt6).
It boots in under 1 second, shows a Windows balloon immediately,
then starts clippy.py as a detached background subprocess.
"""

import sys, os, subprocess, time, ctypes, ctypes.wintypes, threading
from pathlib import Path

# ── Tray balloon via pure Win32 ctypes ───────────────────────────────────────
NIIF_INFO = 0x00000001
NIF_ICON  = 0x00000002
NIF_TIP   = 0x00000004
NIF_INFO  = 0x00000010
NIM_ADD    = 0x00000000
NIM_DELETE = 0x00000002

class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ("cbSize",           ctypes.c_ulong),
        ("hWnd",             ctypes.wintypes.HWND),
        ("uID",              ctypes.c_uint),
        ("uFlags",           ctypes.c_uint),
        ("uCallbackMessage", ctypes.c_uint),
        ("hIcon",            ctypes.wintypes.HANDLE),
        ("szTip",            ctypes.c_wchar * 128),
        ("dwState",          ctypes.c_ulong),
        ("dwStateMask",      ctypes.c_ulong),
        ("szInfo",           ctypes.c_wchar * 256),
        ("uTimeout",         ctypes.c_uint),
        ("szInfoTitle",      ctypes.c_wchar * 64),
        ("dwInfoFlags",      ctypes.c_ulong),
    ]

def _show_balloon(title, message, duration_ms=5000):
    try:
        shell32 = ctypes.windll.shell32
        user32  = ctypes.windll.user32
        hwnd = user32.CreateWindowExW(0, "STATIC", "ClippyTray", 0,
                                      0, 0, 0, 0, None, None, None, None)
        if not hwnd:
            return
        hIcon = user32.LoadIconW(None, ctypes.c_wchar_p(32512))  # IDI_APPLICATION
        nid = NOTIFYICONDATAW()
        nid.cbSize      = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd        = hwnd
        nid.uID         = 1
        nid.uFlags      = NIF_ICON | NIF_TIP | NIF_INFO
        nid.hIcon       = hIcon
        nid.szTip       = "Clippy"
        nid.szInfo      = message
        nid.szInfoTitle = title
        nid.dwInfoFlags = NIIF_INFO
        nid.uTimeout    = duration_ms
        shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid))
        time.sleep(duration_ms / 1000.0 + 0.5)
        shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(nid))
        user32.DestroyWindow(hwnd)
    except Exception:
        pass

if __name__ == "__main__":
    script_dir  = Path(__file__).parent.resolve()

    # Find the main Clippy script in the app folder.
    clippy_main = None
    for name in ("clippyo.py", "clippy.py", "clippysrc.py", "clippy_main.py"):
        candidate = script_dir / name
        if candidate.exists():
            clippy_main = candidate
            break

    if clippy_main is None:
        sys.exit(1)

    # Find pythonw.exe — no console window on launch
    pythonw = Path(sys.executable).parent / "pythonw.exe"
    if not pythonw.exists():
        pythonw = Path(sys.executable).parent.parent / "pythonw.exe"
    if not pythonw.exists():
        pythonw = Path(sys.executable)

    # ── Step 1: Fire balloon notification immediately (non-blocking) ─────────
    threading.Thread(
        target=_show_balloon,
        args=("Clippy is running",
              "Clipboard manager active.\nPress Ctrl+D to open."),
        daemon=True,
    ).start()

    # ── Step 2: Launch full Clippy process fully detached ────────────────────
    # DETACHED_PROCESS + CREATE_NO_WINDOW: subprocess is independent of this
    # launcher — when launcher exits Clippy keeps running.
    DETACHED_PROCESS = 0x00000008
    CREATE_NO_WINDOW = 0x08000000

    try:
        subprocess.Popen(
            [str(pythonw), str(clippy_main)],
            cwd=str(script_dir),
            creationflags=DETACHED_PROCESS | CREATE_NO_WINDOW,
            close_fds=True,
        )
    except Exception:
        try:
            os.startfile(str(clippy_main))
        except Exception:
            pass

    # Launcher exits immediately — Clippy continues independently
    sys.exit(0)
