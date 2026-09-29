@echo off
rem ============================================================
rem V3 basic-data audit tool - pywebview launcher (dev: source first).
rem
rem Binds the repo-root core\: edit code, rerun this script - no packing.
rem ASCII-only + CRLF on purpose: cmd parses batch files before chcp takes
rem effect, so CN comments would be executed as commands (flash-close).
rem
rem Launcher priority:
rem   1. source run.py  (default; interpreter = one that already has deps)
rem   2. packed EXE under dist\  only with --packed
rem
rem Deps: webview, openpyxl, xlrd, lxml, pythoncom/pywintypes (win32com).
rem If the chosen interpreter misses deps we print the pip command instead
rem of auto-installing (do not pollute a user's environment silently).
rem ============================================================
setlocal EnableExtensions
cd /d "%~dp0"

set "WANT_PACKED=0"
if /i "%~1"=="--packed" set "WANT_PACKED=1"

set "ROOT_CORE=%~dp0..\core"
if not exist "%ROOT_CORE%\src\base_audit" (
    echo [ERROR] Repo-root core not found: %ROOT_CORE%\src\base_audit
    echo         Keep shell-pywebview inside the repo root, next to core\.
    pause
    exit /b 1
)

if "%WANT_PACKED%"=="1" goto :TRY_PACKED

rem ---- interpreter search: prefer one that already has the deps ----
set "PY_CMD="
for %%P in (
    "%~dp0..\.venv\Scripts\python.exe"
    "%~dp0..\venv\Scripts\python.exe"
    "%~dp0.venv\Scripts\python.exe"
    "%~dp0venv\Scripts\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
) do (
    if not defined PY_CMD if exist "%%~P" (
        "%%~P" -c "import webview, openpyxl, xlrd, lxml, pythoncom, pywintypes, win32com.client" >nul 2>nul
        if not errorlevel 1 set "PY_CMD=%%~P"
    )
)
if not defined PY_CMD (
    python -c "import sys" >nul 2>nul && set "PY_CMD=python"
)
if not defined PY_CMD (
    py -3 -c "import sys" >nul 2>nul && set "PY_CMD=py -3"
)
if not defined PY_CMD (
    echo [ERROR] Python not found. Install Python 3 and add it to PATH.
    pause
    exit /b 1
)

rem ---- dependency check (pywin32 included) ----
"%PY_CMD%" -c "import webview, openpyxl, xlrd, lxml, pythoncom, pywintypes, win32com.client" >nul 2>nul
if errorlevel 1 (
    echo.
    echo [ERROR] current interpreter misses pywebview runtime deps:
    "%PY_CMD%" -c "import sys; print(sys.executable)"
    echo.
    echo Install them with:
    echo   "%PY_CMD%" -m pip install -r "%~dp0requirements.txt"
    echo.
    pause
    exit /b 1
)

echo [source mode] interpreter: %PY_CMD%
"%PY_CMD%" -c "import sys; print('               ' + sys.executable)"
echo [source mode] core        : %ROOT_CORE%
echo.
"%PY_CMD%" run.py %*
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
    echo.
    echo [ERROR] exited with code %RC%
    pause
)
exit /b %RC%

:TRY_PACKED
set "PACKED="
for %%F in ("%~dp0dist\*.exe") do if not defined PACKED set "PACKED=%%~fF"
if defined PACKED (
    echo [start] packed: %PACKED%
    start "" "%PACKED%" %*
    exit /b 0
)
echo [ERROR] packed EXE not found under dist\.
pause
exit /b 1
