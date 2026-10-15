@echo off
chcp 65001 >nul
call "%~dp0Linux-7-Docker双架构打包.bat" --s1-s2-only
exit /b %ERRORLEVEL%
