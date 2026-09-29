@echo off
rem ==============================================================
rem  Build the Flask shell for Windows 7 SP1+ (max compatibility)
rem
rem  Target: Windows 7 SP1 / later, 32-bit and 64-bit.
rem  Interpreter: Python 3.8.10 win32 preferred - the LAST official
rem  release that runs on Windows 7 (3.9+ dropped Win7). Python 3.7
rem  is also accepted. If no suitable interpreter is found, this
rem  script auto-downloads the official python-3.8.10-embed-win32
rem  runtime into runtime\python38 - no admin rights needed.
rem
rem  Usage:
rem    this script -> dist\win7-TIMESTAMP\ with onefile EXE, core\, config\
rem                   and runtime-repair\ (the ONLY shipping form; the green
rem                   portable folder was removed on 2026-09-21 per the
rem                   user's request - do not re-add it here).
rem
rem  ASCII-only + CRLF on purpose: cmd parses batch files in the ANSI
rem  codepage, so Chinese text or LF endings break the script.
rem ==============================================================
setlocal EnableExtensions
cd /d "%~dp0"

rem ---- [0/4] Shared core: the repo-root core\ is the ONLY source ----------
rem All shells (flask / pywebview) and all build scripts use this one core.
rem A stale per-shell copy would silently ship old code (real incident:
rem the pywebview snapshot lagged 3 minutes and shipped the pre-fix build).
if not exist "%~dp0..\core\src\base_audit" (
  echo [ERROR] Repo-root core not found: %~dp0..\core\src\base_audit
  echo         Keep shell-flask inside the repository root, next to core\.
  pause
  exit /b 1
)
echo [core] sharing repo-root core: %~dp0..\core

set "PY="
set "RT=%~dp0runtime\python38"
set "RTF=%~dp0runtime\python38-full"

rem ---- [1/4] Locate a Win7-capable interpreter: 3.8 best, 3.7 ok ----
rem runtime\python38-full (complete win32) is the PyInstaller build env:
rem PyInstaller cannot run on the embeddable distribution, so EXE mode
rem always uses the full one. runtime\python38 is only kept for the
rem bootstrap cache and is NOT shipped any more.
if defined PYTHON_WIN7 if exist "%PYTHON_WIN7%" set "PY=%PYTHON_WIN7%"
if not defined PY if exist "%RTF%\python.exe" (
  "%RTF%\python.exe" -c "import sys,struct,pyexpat,pip;sys.exit(0 if sys.version_info[:2] in ((3,7),(3,8)) and struct.calcsize('P')==4 else 1)" >nul 2>nul
  if not errorlevel 1 set "PY=%RTF%\python.exe"
)
if not defined PY if exist "%USERPROFILE%\miniconda3\envs\HZPBCwin7x86\python.exe" (
  "%USERPROFILE%\miniconda3\envs\HZPBCwin7x86\python.exe" -c "import sys,struct,pyexpat,pip;sys.exit(0 if sys.version_info[:2] in ((3,7),(3,8)) and struct.calcsize('P')==4 else 1)" >nul 2>nul
  if not errorlevel 1 set "PY=%USERPROFILE%\miniconda3\envs\HZPBCwin7x86\python.exe"
)
if not defined PY if exist "%RT%\python.exe" set "PY=%RT%\python.exe"
if not defined PY (py -3.8-32 -c "import sys" >nul 2>nul && set "PY=py -3.8-32")
if not defined PY (py -3.8 -c "import sys" >nul 2>nul && set "PY=py -3.8")

if defined PY (
  %PY% -c "import sys; sys.exit(0 if sys.version_info[:2] in ((3,8),(3,7)) else 1)" >nul 2>nul
  if errorlevel 1 (
    echo [ERROR] PYTHON_WIN7 points to a Python other than 3.7/3.8.
    echo         Win7 builds need Python 3.8.10 win32 - the last official
    echo         release that runs on Windows 7. Set PYTHON_WIN7 or let
    echo         this script auto-download the runtimes.
    pause
    exit /b 1
  )
  if /i "%PY%"=="%RT%\python.exe" (
    echo [BOOT] PyInstaller cannot run on the embeddable runtime - preparing python38-full ...
    call :bootstrap_python38_full
    if errorlevel 1 (
      pause
      exit /b 1
    )
  )
) else (
  call :bootstrap_python38
  if errorlevel 1 (
    pause
    exit /b 1
  )
)

echo [1/4] Interpreter: %PY%

rem The embeddable runtime hides site-packages via python38._pth by default,
rem which would make every pip-installed package invisible. Ensure it is
rem rewritten whenever this runtime is used - not only right after extraction
rem (an earlier partial run may have left the default file in place).
if exist "%RT%\python38._pth" findstr /c:"site-packages" "%RT%\python38._pth" >nul 2>nul || (
  (echo python38.zip & echo . & echo Lib\site-packages & echo ..\..\shell & echo import site) > "%RT%\python38._pth"
)

%PY% -c "import sys; sys.exit(0 if sys.maxsize <= 2**32 else 1)" >nul 2>nul
if errorlevel 1 echo [WARN] 64-bit Python: the EXE will NOT run on 32-bit Win7. Prefer Python 3.8.10 win32.

rem ---- [2/4] Isolated env + pinned dependencies ----
rem A copied venv can still point to the old checkout. Keep it for diagnosis;
rem create a separate environment when its interpreter is unusable.
set "VPY=%PY%"
set "VENV_DIR=%~dp0build-win7-venv"
if exist "%VENV_DIR%\Scripts\python.exe" (
  "%VENV_DIR%\Scripts\python.exe" -c "import sys,struct,pyexpat;sys.exit(0 if sys.version_info[:2] in ((3,7),(3,8)) and struct.calcsize('P')==4 else 1)" >nul 2>nul
  if errorlevel 1 set "VENV_DIR=%~dp0build-win7-venv-rebuilt"
)
if not exist "%VENV_DIR%\Scripts\python.exe" (
  %PY% -m venv "%VENV_DIR%"
  if errorlevel 1 (
    echo [ERROR] Failed to create a working Win7 build venv: %VENV_DIR%
    pause
    exit /b 1
  )
)
if exist "%VENV_DIR%\Scripts\python.exe" set "VPY=%VENV_DIR%\Scripts\python.exe"
%VPY% -c "import sys,struct,pyexpat;sys.exit(0 if sys.version_info[:2] in ((3,7),(3,8)) and struct.calcsize('P')==4 else 1)" >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Win7 build Python is missing stdlib modules or has the wrong architecture: %VPY%
  pause
  exit /b 1
)

%VPY% -m pip --version >nul 2>nul
if errorlevel 1 (
  if /i "%VPY%"=="%RT%\python.exe" (
    call :bootstrap_getpip
    if errorlevel 1 (
      pause
      exit /b 1
    )
  )
)

rem OFFLINE FIRST: the wheels\ folder ships with the repo, so a build machine
rem without network can build. Online is only a fallback for missing wheels;
rem whatever gets downloaded online is then cached into wheels\ (--find-links
rem does not block index access) so the NEXT build is fully offline again.
echo [2/4] Installing pinned dependencies - offline wheels first ...
%VPY% -m pip install --disable-pip-version-check --no-index --find-links wheels -r requirements-win7.txt
if errorlevel 1 (
  echo [WARN] Offline wheels incomplete - trying online, then caching into wheels\ ...
  %VPY% -m pip install --disable-pip-version-check --find-links wheels -r requirements-win7.txt
  if errorlevel 1 (
    echo [ERROR] Dependency install failed. Fill wheels\ with the win32 wheels listed in requirements-win7.txt.
    pause
    exit /b 1
  )
  %VPY% -m pip download --disable-pip-version-check --find-links wheels -r requirements-win7.txt -d wheels
)

rem pywin32 DLLs sit inside site-packages\pywin32_system32\DLLs and are NOT on
rem the DLL search path of a bare interpreter. Copy them beside python.exe so
rem the build interpreter can import win32com directly (PyInstaller must see
rem a working win32com to bundle it into the EXE).
if exist "%RT%\Lib\site-packages\pywin32_system32\DLLs" (
  copy /y "%RT%\Lib\site-packages\pywin32_system32\DLLs\*.dll" "%RT%\" >nul
  %VPY% -c "import win32com.client, pythoncom; print('[OK] win32com ready')" >nul 2>nul
  if errorlevel 1 echo [WARN] win32com still failing after DLL copy - COM engine may be unavailable.
) else (
  %VPY% -c "import win32com.client" >nul 2>nul
  if errorlevel 1 echo [WARN] pywin32 missing: COM engine - Excel/WPS - unavailable; check config check still works without it.
)
%VPY% -c "import PyInstaller" >nul 2>nul
if errorlevel 1 (
  echo [ERROR] pyinstaller 4.10 unavailable - put pyinstaller and its deps ^(altgraph, pefile, pywin32-ctypes^) into wheels\ for a fully offline build.
  pause
  exit /b 1
)

rem ---- [3/4] Build ----
for /f %%T in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "BUILD_STAMP=%%T"
if not defined BUILD_STAMP set "BUILD_STAMP=%RANDOM%-%RANDOM%"
if not defined FLASK_DIST_DIR set "FLASK_DIST_DIR=%~dp0dist\win7-%BUILD_STAMP%"
echo [output] %FLASK_DIST_DIR%
echo [3/4] Building onefile EXE with PyInstaller 4.10 ...
%VPY% build_exe.py --win7
if errorlevel 1 (
  pause
  exit /b 1
)
%VPY% build_win7.py exe-extras
if errorlevel 1 (
  pause
  exit /b 1
)
echo [4/4] Done. Output: %FLASK_DIST_DIR% - ship the whole folder.
echo.
pause
exit /b 0

:bootstrap_python38_full
rem Full win32 Python 3.8.10 (per-user silent install, no admin needed) for
rem PyInstaller: it does NOT work on the embeddable distribution. Installer
rem is cached at build\python-3.8.10.exe so later builds stay offline.
if exist "%RTF%\python.exe" (
  set "PY=%RTF%\python.exe"
  exit /b 0
)
echo [BOOT] Preparing full Python 3.8.10 win32 for PyInstaller ...
if not exist "build" mkdir "build"
set "FSETUP=%~dp0build\python-3.8.10.exe"
if not exist "%FSETUP%" (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('https://www.python.org/ftp/python/3.8.10/python-3.8.10.exe','%FSETUP%')"
  if errorlevel 1 (
    echo [ERROR] Download failed. Manual fix: download python-3.8.10.exe ^(win32 full installer^) from python.org to build\ , then re-run.
    exit /b 1
  )
)
"%FSETUP%" /quiet InstallAllUsers=0 TargetDir="%RTF%" Include_doc=0 Include_test=0 Include_launcher=0 InstallLauncherAllUsers=0 SimpleInstall=1
if not exist "%RTF%\python.exe" (
  echo [ERROR] Silent install failed. Run build\python-3.8.10.exe manually, install per-user to %RTF% , then re-run.
  exit /b 1
)
set "PY=%RTF%\python.exe"
exit /b 0

:bootstrap_python38
echo [BOOT] No Python 3.7/3.8 found - downloading Python 3.8.10 embeddable win32 ...
if not exist "build" mkdir "build"
set "RTZIP=%~dp0build\python-3.8.10-embed-win32.zip"
if not exist "%RTZIP%" (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('https://www.python.org/ftp/python/3.8.10/python-3.8.10-embed-win32.zip','%RTZIP%')"
  if errorlevel 1 (
    echo [ERROR] Download failed. Manual fix: download python-3.8.10-embed-win32.zip from python.org and place it at build\ , then re-run.
    exit /b 1
  )
)
if not exist "%RT%\python.exe" (
  mkdir "%RT%" 2>nul
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -Force '%RTZIP%' '%RT%'"
  if errorlevel 1 (
    echo [ERROR] Extract failed - PowerShell 5 or newer required. Extract the zip into runtime\python38 manually and re-run.
    exit /b 1
  )
)
rem The embeddable runtime hides site-packages via python38._pth by default,
rem which would make pip-installed packages invisible. Rewrite it once to
rem enable site-packages; ..\..\shell also lets this runtime run app sources.
findstr /c:"site-packages" "%RT%\python38._pth" >nul 2>nul || (
  (echo python38.zip & echo . & echo Lib\site-packages & echo ..\..\shell & echo import site) > "%RT%\python38._pth"
)
set "PY=%RT%\python.exe"
call :bootstrap_getpip
exit /b %errorlevel%

:bootstrap_getpip
%PY% -m pip --version >nul 2>nul
if not errorlevel 1 exit /b 0
set "GETPIP=%~dp0build\get-pip-38.py"
echo [BOOT] Installing pip via get-pip.py for Python 3.8 ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('https://bootstrap.pypa.io/pip/3.8/get-pip.py','%GETPIP%')"
if errorlevel 1 (
  echo [ERROR] get-pip download failed. Manual fix: download https://bootstrap.pypa.io/pip/3.8/get-pip.py to build\get-pip-38.py and re-run.
  exit /b 1
)
%PY% "%GETPIP%" --no-warn-script-location
exit /b %errorlevel%
