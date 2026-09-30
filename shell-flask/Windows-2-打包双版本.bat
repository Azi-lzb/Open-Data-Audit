@echo off
rem Build both Windows EXEs into one timestamped folder under dist\.
rem Internal builders are in windows-build\.
rem ASCII-only content for cmd.exe compatibility.
setlocal EnableExtensions
cd /d "%~dp0"
python "%~dp0build_windows_all.py"
if errorlevel 1 (
  echo [ERROR] Combined Windows build failed. Check BUILD_INCOMPLETE.txt in the output folder.
  pause
  exit /b 1
)
pause
exit /b 0
