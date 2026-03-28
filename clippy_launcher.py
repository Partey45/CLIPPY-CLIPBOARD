"""
clippy_launcher.py - lightweight startup shim for Clippy.

This script is what the Windows Run key launches at login.
It keeps startup fast by avoiding heavy UI imports and simply
starting clippyo.py as a detached background subprocess.

The full Clippy app is responsible for the single user-facing
"ready" notification so users do not get duplicate startup balloons.
"""

import os
import subprocess
import sys
from pathlib import Path


if __name__ == "__main__":
    script_dir = Path(__file__).parent.resolve()

    clippy_main = None
    for name in ("clippyo.py", "clippysrc.py", "clippy_main.py"):
        candidate = script_dir / name
        if candidate.exists():
            clippy_main = candidate
            break

    if clippy_main is None:
        sys.exit(1)

    pythonw = Path(sys.executable).parent / "pythonw.exe"
    if not pythonw.exists():
        pythonw = Path(sys.executable).parent.parent / "pythonw.exe"
    if not pythonw.exists():
        pythonw = Path(sys.executable)

    detached_process = 0x00000008
    create_no_window = 0x08000000

    try:
        subprocess.Popen(
            [str(pythonw), str(clippy_main)],
            cwd=str(script_dir),
            creationflags=detached_process | create_no_window,
            close_fds=True,
        )
    except Exception:
        try:
            os.startfile(str(clippy_main))
        except Exception:
            pass

    sys.exit(0)
