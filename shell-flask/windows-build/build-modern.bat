@echo off
rem ============================================================
rem V3 basic-data audit tool - build the Flask shell EXE
rem Target: Windows 10/11 64-bit (recommended). Output: dist\win10-TIMESTAMP\
rem
rem Shares the repo-root core\: do NOT collect sources from anywhere else.
rem The root .venv is intentionally separate from build-win7-venv so a
rem large workbook can use a modern 64-bit Python without weakening Win7.
rem ASCII-only + CRLF on purpose (cmd parses batch in the ANSI codepage).
rem ============================================================
setlocal EnableExtensions EnableDelayedExpansion
set "SHELL_DIR=%~dp0..\"
cd /d "%SHELL_DIR%"

if not exist "%SHELL_DIR%..\core\src\base_audit" (
    echo [ERROR] Repo-root core not found: %SHELL_DIR%..\core\src\base_audit
    echo         Keep shell-flask inside the repository root next to core\.
    pause
    exit /b 1
)
echo [core] sharing repo-root core: %SHELL_DIR%..\core

set "VENV_DIR=%SHELL_DIR%..\.venv"
set "PY=%VENV_DIR%\Scripts\python.exe"
if exist "%PY%" (
    "%PY%" -c "import struct,sys;sys.exit(0 if struct.calcsize('P')==8 and sys.version_info[:2] in ((3,11),(3,12)) else 1)" >nul 2>nul
    if errorlevel 1 (
        set "VENV_DIR=%SHELL_DIR%..\.venv-flask-build"
        set "PY=%SHELL_DIR%..\.venv-flask-build\Scripts\python.exe"
    )
)
if not exist "%PY%" (
    echo [BOOT] Creating an isolated Windows 10/11 64-bit build environment ...
    set "BOOTSTRAP="
    for %%P in (
        "%PYTHON_MODERN%"
        "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
        "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
        "%USERPROFILE%\miniconda3\envs\HZPBC\python.exe"
    ) do (
        if not defined BOOTSTRAP if exist "%%~P" (
            "%%~P" -c "import struct,sys;sys.exit(0 if struct.calcsize('P')==8 and sys.version_info[:2] in ((3,11),(3,12)) else 1)" >nul 2>nul
            if not errorlevel 1 set "BOOTSTRAP=%%~P"
        )
    )
    if not defined BOOTSTRAP (
        echo [ERROR] 64-bit Python 3.11 or 3.12 was not found.
        echo Set PYTHON_MODERN to its python.exe, or install Python 3.11/3.12.
        pause
        exit /b 1
    )
    "!BOOTSTRAP!" -m venv "!VENV_DIR!"
    if errorlevel 1 (
        echo [ERROR] Failed to create .venv.
        pause
        exit /b 1
    )
)

"%PY%" -c "import struct, sys; sys.exit(0 if struct.calcsize('P') == 8 else 1)" >nul 2>nul
if errorlevel 1 (
    echo [ERROR] .venv must be 64-bit Python 3.10 or newer for the Win10/11 build.
    pause
    exit /b 1
)

"%PY%" -c "import flask, openpyxl, lxml, xlrd, win32com.client, PyInstaller" >nul 2>nul
if errorlevel 1 (
    echo [BOOT] Installing modern 64-bit build dependencies into .venv ...
    "%PY%" -m pip install --disable-pip-version-check --upgrade pip
    "%PY%" -m pip install --disable-pip-version-check -r requirements-win10win11-build.txt
    if errorlevel 1 (
        echo [ERROR] Could not install modern build dependencies.
        pause
        exit /b 1
    )
)

echo [environment] Windows 10/11 64-bit: %PY%
for /f %%T in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "BUILD_STAMP=%%T"
if not defined BUILD_STAMP set "BUILD_STAMP=%RANDOM%-%RANDOM%"
if not defined FLASK_DIST_DIR set "FLASK_DIST_DIR=%SHELL_DIR%dist\win10-%BUILD_STAMP%"
echo [output] %FLASK_DIST_DIR%
"%PY%" build_exe.py
if errorlevel 1 (
    pause
    exit /b 1
)
echo [DONE] Ship this whole folder: %FLASK_DIST_DIR%
pause
