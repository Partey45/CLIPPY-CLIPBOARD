# Clippy Clipboard

Clippy is a Windows clipboard manager built with Python + PyQt6.  
It helps users capture, organize, search, and paste clipboard entries quickly through a keyboard-first workflow.

## Features

- Group-based clipboard organization
- Multiple entry variations per group
- Fast search and keyboard navigation
- Tray-based background app behavior
- Global `Ctrl + D` hotkey to open Clippy
- Import/export support (JSON/CSV)
- Auto-start on user login (Windows)

## Platform

- Windows (primary target)

## Tech Stack

- Python
- PyQt6 / Qt WebEngine
- PyInstaller
- `pyperclip` and `keyboard` for clipboard + hotkey behavior

## Project Files

- `clippyo.py`: Main desktop application (UI, hotkeys, data logic)
- `clippy_launcher.py`: Lightweight launcher for startup/login flow
- `build_exe.bat`: Build helper for executable packaging
- `Clippy.spec`: PyInstaller spec file
- `clippy_logo_official.svg`: Official logo asset

## Quick Start (Run From Source)

1. Create and activate a Python environment.
2. Install dependencies.
3. Run the app.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python .\clippyo.py
```

## Build Executable

```powershell
build_exe.bat
```

If needed, you can also build directly with PyInstaller using `Clippy.spec`.

## Usage

- Press `Ctrl + D` to open Clippy near the caret.
- Use arrow keys to navigate entries.
- Press `Enter` to paste selected content.
- Use Settings for theme, capture/history toggles, and import/export.

## Startup Behavior

Clippy is designed to run in the background after user sign-in.  
The launcher (`clippy_launcher.py`) is used for quick startup and background activation.

## Troubleshooting

- If global hotkey does not respond, run Clippy as the same privilege level as the target app.
- If clipboard capture fails, verify required Python packages installed correctly.
- If startup does not trigger, verify Windows login-start entry points to `clippy_launcher.py`.
- For GPU/WebEngine issues, check logs in `clippy.log`.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Open a pull request

## License

MIT License
