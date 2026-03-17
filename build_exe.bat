@echo off
title Clippy v13 — EXE Builder (Full Fix)
color 0A
setlocal enabledelayedexpansion

echo.
echo  ==========================================
echo   Clippy v13 EXE Builder — Full Qt Fix
echo  ==========================================
echo.

:: ── Find Python 3.11 ──────────────────────────────────────────────────────────
py -3.11 --version >nul 2>&1
if not errorlevel 1 (
    set PYTHON=py -3.11
    set PIP=py -3.11 -m pip
    goto :BUILD
)

set PY311=%LOCALAPPDATA%\Programs\Python\Python311\python.exe
if exist "%PY311%" (
    set PYTHON="%PY311%"
    set PIP="%PY311%" -m pip
    goto :BUILD
)

set PY311B=%PROGRAMFILES%\Python311\python.exe
if exist "%PY311B%" (
    set PYTHON="%PY311B%"
    set PIP="%PY311B%" -m pip
    goto :BUILD
)

echo  Python 3.11 not found. Downloading...
set PY_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe
set PY_INSTALLER=%TEMP%\python311_installer.exe
powershell -Command "Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%PY_INSTALLER%' -UseBasicParsing"
if not exist "%PY_INSTALLER%" (
    echo  Download failed. Install Python 3.11 from https://python.org
    pause & exit /b 1
)
"%PY_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=0 Include_test=0
set PYTHON="%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
set PIP="%LOCALAPPDATA%\Programs\Python\Python311\python.exe" -m pip

:BUILD
echo  Using:
%PYTHON% --version
echo.

echo  [1/5] Upgrading pip...
%PIP% install --upgrade pip --quiet

echo  [2/5] Installing dependencies...
%PIP% install "pyinstaller==6.10.0" "PyQt6==6.7.1" "PyQt6-WebEngine==6.7.0" "PyQt6-WebEngine-Qt6==6.7.0" pyperclip anthropic keyboard --quiet --upgrade

echo  [3/5] Cleaning old build...
if exist build  rmdir /s /q build
if exist dist   rmdir /s /q dist
if exist clippy_hook.py del clippy_hook.py

echo  [4/5] Writing PyInstaller hook...

:: Write a hook file that copies all Qt WebEngine pak/locale/resource files
(
echo import os, glob, shutil
echo from PyInstaller.utils.hooks import collect_data_files, get_package_paths
echo.
echo # Collect ALL PyQt6 WebEngine data files ^(pak, locales, resources^)
echo datas = collect_data_files^('PyQt6', includes=['**/*.pak', '**/*.dat', '**/*.bin', '**/*.so', '**/*.pyd']^)
echo datas += collect_data_files^('PyQt6.Qt6', includes=['**']^)
) > clippy_hook.py

echo  [5/5] Building Clippy.exe...
echo  ^(This takes 2-5 minutes, please wait^)
echo.

%PYTHON% -m PyInstaller ^
  --name "Clippy" ^
  --noconsole ^
  --noconfirm ^
  --clean ^
  --hidden-import "PyQt6" ^
  --hidden-import "PyQt6.QtWidgets" ^
  --hidden-import "PyQt6.QtWebEngineWidgets" ^
  --hidden-import "PyQt6.QtWebEngineCore" ^
  --hidden-import "PyQt6.QtWebChannel" ^
  --hidden-import "PyQt6.QtCore" ^
  --hidden-import "PyQt6.QtGui" ^
  --hidden-import "PyQt6.QtNetwork" ^
  --hidden-import "PyQt6.QtPrintSupport" ^
  --hidden-import "PyQt6.sip" ^
  --hidden-import "pyperclip" ^
  --hidden-import "anthropic" ^
  --hidden-import "keyboard" ^
  --collect-all "PyQt6" ^
  --collect-all "PyQt6-WebEngine" ^
  --collect-data "PyQt6" ^
  clippy.py

echo.

:: ── Manually copy any missing .pak files ──────────────────────────────────────
echo  Patching missing Qt WebEngine resource files...

:: Find where PyQt6 is installed in this Python
for /f "delims=" %%i in ('%PYTHON% -c "import PyQt6; import os; print(os.path.dirname(PyQt6.__file__))"') do set PYQT6_PATH=%%i

echo  PyQt6 source: %PYQT6_PATH%
echo  Copying resources to dist...

set DEST=dist\Clippy\_internal\PyQt6\Qt6\resources
if not exist "%DEST%" mkdir "%DEST%"

:: Copy all .pak files from the installed PyQt6
if exist "%PYQT6_PATH%\Qt6\resources\" (
    xcopy /s /y "%PYQT6_PATH%\Qt6\resources\*" "%DEST%\" >nul 2>&1
    echo  Copied resources from %PYQT6_PATH%\Qt6\resources\
) else (
    echo  WARNING: Could not find Qt6 resources folder automatically.
)

:: Copy locales too
set DEST_LOC=dist\Clippy\_internal\PyQt6\Qt6\translations
if not exist "%DEST_LOC%" mkdir "%DEST_LOC%"
if exist "%PYQT6_PATH%\Qt6\translations\" (
    xcopy /s /y "%PYQT6_PATH%\Qt6\translations\*" "%DEST_LOC%\" >nul 2>&1
    echo  Copied translations.
)

echo.

if exist "dist\Clippy\Clippy.exe" (
    echo  ==========================================
    echo   SUCCESS!
    echo   Location: dist\Clippy\Clippy.exe
    echo.
    echo   IMPORTANT: Share the ENTIRE dist\Clippy\
    echo   folder — not just the .exe file.
    echo  ==========================================
    echo.
    explorer "dist\Clippy"
) else (
    echo  !! Build failed. See errors above.
    echo.
    echo  Try:
    echo    1. Right-click this .bat and Run as Administrator
    echo    2. Disable antivirus temporarily
    echo    3. Make sure clippy.py is in this same folder
)

pause
