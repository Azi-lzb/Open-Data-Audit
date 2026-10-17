@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Open Data Audit - Full Tests

for /f "tokens=2 delims=:" %%C in ('chcp') do set "OLD_CODEPAGE=%%C"
set "OLD_CODEPAGE=%OLD_CODEPAGE: =%"
chcp 65001 >nul

set "NO_PAUSE=0"
set "RUN_WINDOWS=1"
set "RUN_WSL=0"
set "RUN_ARM64_QEMU=0"
set "WSL_DISTRO="

:READ_ARGS
if "%~1"=="" goto ARGS_DONE
if /i "%~1"=="--no-pause" (
    set "NO_PAUSE=1"
    shift
    goto READ_ARGS
)
if /i "%~1"=="--wsl-linux" (
    set "RUN_WINDOWS=0"
    set "RUN_WSL=1"
    shift
    goto READ_ARGS
)
if /i "%~1"=="--with-wsl-linux" (
    set "RUN_WSL=1"
    shift
    goto READ_ARGS
)
if /i "%~1"=="--arm64-qemu" (
    set "RUN_WINDOWS=0"
    set "RUN_ARM64_QEMU=1"
    shift
    goto READ_ARGS
)
if /i "%~1"=="--with-arm64-qemu" (
    set "RUN_ARM64_QEMU=1"
    shift
    goto READ_ARGS
)
if /i "%~1"=="--all-platforms" (
    set "RUN_WSL=1"
    set "RUN_ARM64_QEMU=1"
    shift
    goto READ_ARGS
)
if /i "%~1"=="--wsl-distro" (
    if "%~2"=="" (
        echo [ERROR] --wsl-distro requires a WSL distribution name.
        set "RC=2"
        goto FINISH
    )
    set "WSL_DISTRO=%~2"
    shift
    shift
    goto READ_ARGS
)
if /i "%~1"=="--help" goto HELP
echo [ERROR] Unknown option: %~1
goto SHOW_USAGE_ERROR

:HELP
set "NO_PAUSE=1"
set "RC=0"
echo Usage: %~nx0 [--no-pause] [--wsl-linux^|--with-wsl-linux] [--arm64-qemu^|--with-arm64-qemu^|--all-platforms] [--wsl-distro NAME]
echo   No platform option: run the Windows full test suite.
echo   --wsl-linux:         run the Linux suite in the default WSL distro only.
echo   --with-wsl-linux:    run Windows, then Linux in WSL.
echo   --arm64-qemu:        run the ARM64 QEMU suite only (requires Docker in WSL).
echo   --with-arm64-qemu:   run Windows, then the ARM64 QEMU suite.
echo   --all-platforms:     run Windows, WSL Linux, then ARM64 QEMU.
echo   --wsl-distro NAME:   select a WSL distro, for example Ubuntu.
goto FINISH

:SHOW_USAGE_ERROR
set "RC=2"
echo Usage: %~nx0 [--no-pause] [--wsl-linux^|--with-wsl-linux] [--arm64-qemu^|--with-arm64-qemu^|--all-platforms] [--wsl-distro NAME]
goto FINISH

:ARGS_DONE
set "RC=0"
if not exist "core\tests\conftest.py" (
    echo [ERROR] Project tests were not found at: %CD%\core\tests
    set "RC=1"
    goto FINISH
)

if "%RUN_WINDOWS%"=="1" call :RUN_WINDOWS_TEST
if "%RUN_WSL%"=="1" call :RUN_WSL_TEST
if "%RUN_ARM64_QEMU%"=="1" call :RUN_ARM64_QEMU_TEST

echo.
echo [result] Overall exit code: %RC%
goto FINISH

:RUN_WINDOWS_TEST
echo ============================================================
echo Running Windows full core test suite
echo ============================================================

rem Prefer the project venv, then Python on PATH, then the Python launcher.
rem Use an interpreter only when it can import pytest.
set "PY_EXE="
set "PY_ARGS="
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import pytest" >nul 2>&1
    if not errorlevel 1 set "PY_EXE=%CD%\.venv\Scripts\python.exe"
)
if not defined PY_EXE (
    python -c "import pytest" >nul 2>&1
    if not errorlevel 1 set "PY_EXE=python"
)
if not defined PY_EXE (
    py -3 -c "import pytest" >nul 2>&1
    if not errorlevel 1 (
        set "PY_EXE=py"
        set "PY_ARGS=-3"
    )
)
if not defined PY_EXE (
    echo [ERROR] Python 3 with pytest was not found.
    echo Install pytest into the Python interpreter you use for this project:
    echo   python -m pip install pytest
    set "WINDOWS_RC=1"
    goto WINDOWS_DONE
)

where powershell >nul 2>&1
if errorlevel 1 (
    echo [ERROR] PowerShell is required to create the timestamped test report.
    set "WINDOWS_RC=1"
    goto WINDOWS_DONE
)

set "REPORT_DIR=%~dp0test-output\full-test-reports"
if not exist "%REPORT_DIR%" mkdir "%REPORT_DIR%" >nul 2>&1
if not exist "%REPORT_DIR%" (
    echo [ERROR] Could not create report directory: %REPORT_DIR%
    set "WINDOWS_RC=1"
    goto WINDOWS_DONE
)
for /f "delims=" %%T in ('powershell -NoLogo -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "RUN_STAMP=%%T"
if not defined RUN_STAMP set "RUN_STAMP=%RANDOM%_%RANDOM%"
for /f "delims=" %%A in ('powershell -NoLogo -NoProfile -Command "$a=$env:PROCESSOR_ARCHITEW6432; if (-not $a) {$a=$env:PROCESSOR_ARCHITECTURE}; if ($a -match 'ARM64') {'arm64'} elseif ($a -match 'AMD64') {'amd64'} else {'x86'}"') do set "PLATFORM_ARCH=%%A"
if not defined PLATFORM_ARCH set "PLATFORM_ARCH=unknown"
set "REPORT_FILE=%REPORT_DIR%\FullTest_windows-%PLATFORM_ARCH%_%RUN_STAMP%.txt"

echo Working directory: %CD%
echo Platform: Windows %PLATFORM_ARCH%
echo Python interpreter:
"%PY_EXE%" %PY_ARGS% -c "import sys; print(sys.executable); print(sys.version.split()[0])"
echo Report: %REPORT_FILE%
echo.
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\run_full_test_report.ps1" -PythonExe "%PY_EXE%" -PythonArgs "%PY_ARGS%" -PlatformArch "%PLATFORM_ARCH%" -ReportFile "%REPORT_FILE%"
set "WINDOWS_RC=%ERRORLEVEL%"

:WINDOWS_DONE
if not "%WINDOWS_RC%"=="0" if "%RC%"=="0" set "RC=%WINDOWS_RC%"
echo [Windows] Exit code: %WINDOWS_RC%
exit /b

:RUN_WSL_TEST
echo ============================================================
echo Running Linux full core test suite in WSL
echo ============================================================
if not exist "run_full_test_linux.sh" (
    echo [ERROR] Linux test runner was not found: %CD%\run_full_test_linux.sh
    set "WSL_RC=2"
    goto WSL_DONE
)
where wsl.exe >nul 2>&1
if errorlevel 1 (
    echo [ERROR] wsl.exe was not found. Enable WSL and install a Linux distribution.
    set "WSL_RC=2"
    goto WSL_DONE
)
where powershell >nul 2>&1
if errorlevel 1 (
    echo [ERROR] PowerShell is required to pass the repository path to WSL.
    set "WSL_RC=2"
    goto WSL_DONE
)
if defined WSL_DISTRO (
    powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\run_full_test_linux_wsl.ps1" -Distribution "%WSL_DISTRO%"
) else (
    powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\run_full_test_linux_wsl.ps1"
)
set "WSL_RC=%ERRORLEVEL%"

:WSL_DONE
if not "%WSL_RC%"=="0" if "%RC%"=="0" set "RC=%WSL_RC%"
echo [WSL Linux] Exit code: %WSL_RC%
exit /b

:RUN_ARM64_QEMU_TEST
echo ============================================================
echo Running ARM64 full test suite in Docker/QEMU via WSL
echo ============================================================
if not exist "tools\run_full_test_arm64_qemu.sh" (
    echo [ERROR] ARM64 QEMU test runner was not found: %CD%\tools\run_full_test_arm64_qemu.sh
    set "ARM64_RC=2"
    goto ARM64_DONE
)
where wsl.exe >nul 2>&1
if errorlevel 1 (
    echo [ERROR] wsl.exe was not found. Enable WSL and install a Linux distribution.
    set "ARM64_RC=2"
    goto ARM64_DONE
)
where powershell >nul 2>&1
if errorlevel 1 (
    echo [ERROR] PowerShell is required to pass the repository path to WSL.
    set "ARM64_RC=2"
    goto ARM64_DONE
)
if defined WSL_DISTRO (
    powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\run_full_test_linux_wsl.ps1" -Arm64Qemu -Distribution "%WSL_DISTRO%"
) else (
    powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\run_full_test_linux_wsl.ps1" -Arm64Qemu
)
set "ARM64_RC=%ERRORLEVEL%"

:ARM64_DONE
if not "%ARM64_RC%"=="0" if "%RC%"=="0" set "RC=%ARM64_RC%"
echo [ARM64 QEMU] Exit code: %ARM64_RC%
exit /b

:FINISH
if defined OLD_CODEPAGE chcp %OLD_CODEPAGE% >nul
if not "%NO_PAUSE%"=="1" (
    echo.
    pause
)
exit /b %RC%
