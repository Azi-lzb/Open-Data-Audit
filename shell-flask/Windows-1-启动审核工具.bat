@echo off
rem ============================================================
rem V3 basic-data audit tool - Flask launcher (dev: source first).
rem Binds the repo-root core/: edit code, rerun this script - no build.
rem ASCII-only + CRLF on purpose (cmd parses batch in the ANSI codepage).
rem
rem Launcher priority:
rem   1. source run.py (default; interpreter = one that already has deps)
rem   2. packed EXE under dist\ only with --packed [win7]
rem      --packed          -> Win10/11 64-bit recommended EXE
rem      --packed win7     -> Win7-compatible 32-bit EXE
rem ============================================================
setlocal EnableExtensions
cd /d "%~dp0"

set "WANT_PACKED=0"
set "PACKED_KIND=MODERN"
if /i "%~1"=="--packed" set "WANT_PACKED=1"
if /i "%~2"=="win7" set "PACKED_KIND=WIN7"

set "ROOT_CORE=%~dp0..\core"
if not exist "%ROOT_CORE%\src\base_audit" (
    echo [ERROR] Repo-root core not found: %ROOT_CORE%\src\base_audit
    echo         Keep shell-flask inside the repository root next to core\.
    pause
    exit /b 1
)

if "%WANT_PACKED%"=="1" goto :TRY_PACKED

rem ---- interpreter search: prefer one that already has the deps ----
set "PY="
set "PY_ARGS="
set "CANDIDATES=%~dp0..\.venv\Scripts\python.exe;%~dp0..\venv\Scripts\python.exe;%~dp0build-win7-venv\Scripts\python.exe;%~dp0runtime\python38-full\python.exe;%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
for %%P in ("%CANDIDATES:;=" "%") do (
    if not defined PY if exist "%%~P" (
        "%%~P" -c "import flask, openpyxl, win32com.client" >nul 2>nul
        if not errorlevel 1 set "PY=%%~P"
    )
)
if not defined PY (
    python -c "import flask, openpyxl, win32com.client" >nul 2>nul && set "PY=python"
)
if not defined PY (
    py -3 -c "import flask, openpyxl, win32com.client" >nul 2>nul && set "PY=py" && set "PY_ARGS=-3"
)
if not defined PY (
    echo [ERROR] No working Python with Flask, openpyxl and pywin32 was found.
    echo         A copied venv may still point to its old Python location.
    echo         Install deps into a working Python 3 with:
    echo         python -m pip install -r "%~dp0requirements.txt"
    pause
    exit /b 1
)

"%PY%" %PY_ARGS% -c "import flask, openpyxl, win32com.client" >nul 2>nul
if errorlevel 1 (
    echo.
    echo [ERROR] current interpreter misses Flask runtime deps:
    "%PY%" %PY_ARGS% -c "import sys; print(sys.executable)"
    echo.
    echo Fix:
    echo   "%PY%" %PY_ARGS% -m pip install -r "%~dp0requirements.txt"
    echo.
    pause
    exit /b 1
)

echo [source mode] interpreter: %PY%
"%PY%" %PY_ARGS% -c "import sys; print('               ' + sys.executable)"
echo [source mode] core        : %ROOT_CORE%
echo.
"%PY%" %PY_ARGS% run.py %*
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
    echo.
    echo [ERROR] exited with code %RC%
    pause
)
exit /b %RC%

:TRY_PACKED
set "PACKED="
if /i "%PACKED_KIND%"=="WIN7" (
    for /f "delims=" %%D in ('dir /b /ad /o-d "%~dp0dist\windows-*" 2^>nul') do (
        for %%F in ("%~dp0dist\%%D\*Win7*.exe") do if not defined PACKED if exist "%%~fF" set "PACKED=%%~fF"
    )
    for /f "delims=" %%D in ('dir /b /ad /o-d "%~dp0dist\win7-*" 2^>nul') do (
        for %%F in ("%~dp0dist\%%D\*Win7*.exe") do if not defined PACKED if exist "%%~fF" set "PACKED=%%~fF"
    )
    rem ASCII-only wildcard avoids encoding a Chinese filename in this BAT.
    for %%F in ("%~dp0dist\*Win7*.exe") do if not defined PACKED if exist "%%~fF" set "PACKED=%%~fF"
) else (
    for /f "delims=" %%D in ('dir /b /ad /o-d "%~dp0dist\windows-*" 2^>nul') do (
        for %%F in ("%~dp0dist\%%D\*_Flask_V*.exe") do if not defined PACKED if exist "%%~fF" set "PACKED=%%~fF"
    )
    for /f "delims=" %%D in ('dir /b /ad /o-d "%~dp0dist\win10-*" 2^>nul') do (
        for %%F in ("%~dp0dist\%%D\*_Flask_V*.exe") do if not defined PACKED if exist "%%~fF" set "PACKED=%%~fF"
    )
    for %%F in ("%~dp0dist\*_Flask_V*.exe") do if not defined PACKED if exist "%%~fF" set "PACKED=%%~fF"
    rem Backward compatibility for the previous lower-case v version marker.
    for /f "delims=" %%D in ('dir /b /ad /o-d "%~dp0dist\windows-*" 2^>nul') do (
        for %%F in ("%~dp0dist\%%D\*_Flask_v*.exe") do if not defined PACKED if exist "%%~fF" set "PACKED=%%~fF"
    )
    for /f "delims=" %%D in ('dir /b /ad /o-d "%~dp0dist\win10-*" 2^>nul') do (
        for %%F in ("%~dp0dist\%%D\*_Flask_v*.exe") do if not defined PACKED if exist "%%~fF" set "PACKED=%%~fF"
    )
    for %%F in ("%~dp0dist\*_Flask_v*.exe") do if not defined PACKED if exist "%%~fF" set "PACKED=%%~fF"
    rem Backward compatibility for releases made before filenames included the version.
    for /f "delims=" %%D in ('dir /b /ad /o-d "%~dp0dist\windows-*" 2^>nul') do (
        for %%F in ("%~dp0dist\%%D\*Flask.exe") do if not defined PACKED if exist "%%~fF" set "PACKED=%%~fF"
    )
    for /f "delims=" %%D in ('dir /b /ad /o-d "%~dp0dist\win10-*" 2^>nul') do (
        for %%F in ("%~dp0dist\%%D\*Flask.exe") do if not defined PACKED if exist "%%~fF" set "PACKED=%%~fF"
    )
    for %%F in ("%~dp0dist\*Flask.exe") do if not defined PACKED if exist "%%~fF" set "PACKED=%%~fF"
)
if defined PACKED (
    echo [start] packed: %PACKED%
    start "" "%PACKED%"
    exit /b 0
)
echo [ERROR] requested packed EXE not found under dist\.
echo         Run the modern build script or the Win7 build script first.
pause
exit /b 1
