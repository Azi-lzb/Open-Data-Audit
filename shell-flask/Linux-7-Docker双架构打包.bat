@echo off
setlocal EnableExtensions
chcp 65001 >nul

set "WSL_DISTRO=Ubuntu"
set "DOCKER_DESKTOP=%ProgramFiles%\Docker\Docker\Docker Desktop.exe"
set "LINUX_SCRIPT=Linux-7-Docker双架构打包.sh"
set "BUILD_OPTIONS="
if "%~1"=="--s1-s2-only" set "BUILD_OPTIONS=--s1-s2-only"
if "%~1"=="--profile" (
    if "%~2"=="" goto invalid_options
    if not "%~3"=="" goto invalid_options
    set "BUILD_OPTIONS=--profile %~2"
)
if not "%~1"=="" if not "%~1"=="--s1-s2-only" if not "%~1"=="--profile" goto invalid_options
if "%~1"=="--s1-s2-only" if not "%~2"=="" goto invalid_options

if not exist "%~dp0%LINUX_SCRIPT%" (
    echo [错误] 找不到 %LINUX_SCRIPT%
    pause
    exit /b 1
)

where wsl.exe >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 WSL，请先安装并配置 Ubuntu。
    pause
    exit /b 1
)

if not exist "%DOCKER_DESKTOP%" (
    echo [错误] 找不到 Docker Desktop：
    echo %DOCKER_DESKTOP%
    pause
    exit /b 1
)

echo [1/3] 检查 Ubuntu WSL……
wsl.exe --distribution "%WSL_DISTRO%" --exec true >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到可用的 WSL 发行版：%WSL_DISTRO%
    echo 请确认 Ubuntu 已安装，且 Docker Desktop 已启用 Ubuntu 的 WSL 集成。
    pause
    exit /b 1
)

echo [2/3] 启动 Docker Desktop 并等待引擎就绪……
tasklist /FI "IMAGENAME eq Docker Desktop.exe" 2>nul | find /I "Docker Desktop.exe" >nul
if errorlevel 1 start "" /min "%DOCKER_DESKTOP%"

set /a WAIT_COUNT=0
:wait_for_docker
wsl.exe --distribution "%WSL_DISTRO%" --exec docker info >nul 2>&1
if not errorlevel 1 goto docker_ready
set /a WAIT_COUNT+=1
if %WAIT_COUNT% geq 90 goto docker_timeout
timeout /t 2 /nobreak >nul
goto wait_for_docker

:docker_timeout
echo [错误] 等待 Docker 引擎超时。
echo 请检查 Docker Desktop 是否已启动，以及 Ubuntu 的 WSL 集成是否已开启。
pause
exit /b 1

:docker_ready
echo [3/3] 在 Ubuntu WSL 中构建 amd64 和 arm64 交付包……
if "%~1"=="--s1-s2-only" (echo [交付范围] S1/S2 专用版)
if "%~1"=="--profile" (echo [交付方案] %~2)
if "%~1"=="" (echo [交付范围] 完整版)
set "WSL_SCRIPT_DIR="
for /f "usebackq delims=" %%P in (`wsl.exe --distribution "%WSL_DISTRO%" --exec wslpath -u "%~dp0." 2^>nul`) do set "WSL_SCRIPT_DIR=%%P"
if not defined WSL_SCRIPT_DIR (
    echo [错误] 无法把脚本目录转换为 WSL 路径。
    pause
    exit /b 1
)

wsl.exe --distribution "%WSL_DISTRO%" --cd "%WSL_SCRIPT_DIR%" --exec sh "./%LINUX_SCRIPT%" %BUILD_OPTIONS%
set "BUILD_EXIT=%ERRORLEVEL%"
if not "%BUILD_EXIT%"=="0" goto build_failed

echo.
echo 双架构打包完成。交付目录位于 shell-flask\dist。
pause
exit /b 0

:build_failed
echo.
echo [错误] 打包失败，退出代码：%BUILD_EXIT%
pause
exit /b %BUILD_EXIT%

:invalid_options
echo [错误] 用法：直接运行此文件，或传 --profile 方案标识。
echo        --s1-s2-only 是旧参数兼容写法。
pause
exit /b 2
