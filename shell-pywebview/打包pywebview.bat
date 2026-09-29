@echo off
rem ============================================================
rem Build the unified pywebview shell EXE (modern Windows, WebView2).
rem Output: dist\<exe>.exe + dist\core\
rem
rem Shares the repo-root core\: the ONLY business core for every shell and
rem build script. A stale per-shell copy would silently ship old code
rem (real incident: this snapshot lagged 3 minutes and shipped pre-fix code).
rem ASCII-only + CRLF on purpose (cmd parses batch in the ANSI codepage).
rem ============================================================
cd /d "%~dp0"

if not exist "%~dp0..\core\src\base_audit" (
    echo [ERROR] Repo-root core not found: %~dp0..\core\src\base_audit
    echo         Keep shell-pywebview inside the repository root next to core\.
    pause
    exit /b 1
)
echo [core] sharing repo-root core: %~dp0..\core

set PY=
for %%P in (
    "%~dp0..\.venv\Scripts\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
) do (
    if not defined PY if exist "%%~P" (
        "%%~P" -c "import tkinter, webview, PyInstaller" >nul 2>nul
        if not errorlevel 1 set "PY=%%~P"
    )
)
if not defined PY (python -c "import sys" >nul 2>nul && set PY=python)
if not defined PY (py -3 -c "import sys" >nul 2>nul && set PY=py -3)
if not defined PY (
    echo [ERROR] Python not found. Install Python and add it to PATH.
    pause
    exit /b 1
)

"%PY%" -c "import tkinter, webview, PyInstaller" >nul 2>nul
if errorlevel 1 (
    echo [ERROR] tkinter / pywebview / PyInstaller missing in:
    "%PY%" -c "import sys; print(sys.executable)"
    echo Install them first:
    echo   "%PY%" -m pip install -r requirements.txt
    pause
    exit /b 1
)

"%PY%" build_exe.py
if errorlevel 1 (
    pause
    exit /b 1
)
pause
