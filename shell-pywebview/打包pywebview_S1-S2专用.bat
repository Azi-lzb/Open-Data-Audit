@echo off
rem Build the S1/S2 pywebview release into a separate backup folder.
setlocal EnableExtensions
cd /d "%~dp0"
set "PYWEBVIEW_DIST_DIR=%~dp0dist\S1-S2"

if not exist "%~dp0..\core\src\base_audit" (
    echo [ERROR] Repo-root core not found: %~dp0..\core\src\base_audit
    echo         Keep shell-pywebview inside the repository root next to core\.
    pause
    exit /b 1
)
echo [core] sharing repo-root core: %~dp0..\core

set PY=
set PY_ARGS=
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
if not defined PY (
    py -3 -c "import sys" >nul 2>nul
    if not errorlevel 1 (
        set "PY=py"
        set "PY_ARGS=-3"
    )
)
if not defined PY (
    echo [ERROR] Python not found. Install Python and add it to PATH.
    pause
    exit /b 1
)

"%PY%" %PY_ARGS% -c "import tkinter, webview, PyInstaller" >nul 2>nul
if errorlevel 1 (
    echo [ERROR] tkinter / pywebview / PyInstaller missing in:
    "%PY%" %PY_ARGS% -c "import sys; print(sys.executable)"
    echo Install them first:
    echo   "%PY%" %PY_ARGS% -m pip install -r requirements.txt
    pause
    exit /b 1
)

"%PY%" %PY_ARGS% build_exe.py --profile s1s2
if errorlevel 1 (
    pause
    exit /b 1
)
pause
exit /b 0
