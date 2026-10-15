@echo off
rem Double-click to build the S1/S2 delivery for Windows 10/11 and Windows 7.
setlocal EnableExtensions
cd /d "%~dp0"
python "%~dp0build_windows_all.py" %* --profile s1s2
if errorlevel 1 (
    echo [ERROR] S1/S2 Windows build failed. See the output above and BUILD_INCOMPLETE.txt.
    pause
    exit /b 1
)
pause
exit /b 0
