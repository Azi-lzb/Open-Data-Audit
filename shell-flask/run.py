from __future__ import annotations

import argparse
import atexit
import json
import os
import signal
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path

# 旧实例 PID 锁文件（core/data 是程序运行数据目录，升级保留）。
# UOS DEB 将其放在用户数据目录，避免向 /opt 写入；源码布局保持原路径。
PROJECT_ROOT = Path(
    os.environ.get("BASE_AUDIT_PROJECT_ROOT")
    or (Path(__file__).resolve().parent.parent / "core")
).expanduser()
PID_FILE = PROJECT_ROOT / "data" / "app.pid"

_KNOWN_COMMAND_MARKERS = (
    "\\shell-flask\\", "/shell-flask/", "基础数据审核工具_flask", "run.py", "base-audit-v3",
)


def _process_is_alive(pid: int) -> bool:
    """跨平台检查 PID 是否仍存在；不能仅因锁文件存在就视为正在运行。"""
    if pid <= 0:
        return False
    if sys.platform == "win32":
        import ctypes

        process_query_limited_information = 0x1000
        handle = ctypes.windll.kernel32.OpenProcess(
            process_query_limited_information, False, pid
        )
        if not handle:
            return False
        ctypes.windll.kernel32.CloseHandle(handle)
        return True
    try:
        os.kill(pid, 0)
    except OSError:
        return False
    return True


def _process_matches(pid: int) -> bool:
    """校验 pid 是否真的是本程序的旧实例（只关自己，不误伤其它进程）。"""
    if sys.platform == "win32":
        if not _process_is_alive(pid):
            return False
        # 仅当该 PID 正在监听本工具端口时才会用此结果结束进程。PowerShell/CIM
        # 不可用时宁可保守地返回 False，随后改用备用端口，不误杀其它服务。
        command = (
            "$p=Get-CimInstance Win32_Process -Filter 'ProcessId=" + str(pid)
            + "'; if($p){$p.CommandLine}"
        )
        try:
            result = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
                capture_output=True, text=True, encoding="utf-8", errors="ignore",
                timeout=3, check=False,
            )
        except (OSError, subprocess.TimeoutExpired):
            return False
        cmdline = result.stdout.casefold()
        return any(marker.casefold() in cmdline for marker in _KNOWN_COMMAND_MARKERS)
    try:
        with open(f"/proc/{pid}/cmdline", "rb") as fh:
            cmdline = fh.read().decode("utf-8", "ignore")
    except OSError:
        return False
    return any(marker in cmdline for marker in _KNOWN_COMMAND_MARKERS)


def _read_lock_pid() -> int:
    """读取新旧格式的锁文件；旧版本只保存一个 PID。"""
    try:
        raw = PID_FILE.read_text(encoding="utf-8").strip()
        payload = json.loads(raw) if raw.startswith("{") else raw
        return int(payload.get("pid", 0) if isinstance(payload, dict) else payload)
    except (OSError, ValueError, json.JSONDecodeError):
        return 0


def _remove_lock_if_owned(pid: int) -> None:
    """仅删除属于指定 PID 的锁，避免新实例启动后被旧实例清理。"""
    if _read_lock_pid() != pid:
        return
    try:
        PID_FILE.unlink()
    except OSError:
        pass


def _listening_pid(port: int) -> int:
    """返回监听本机 TCP 端口的 PID；解析 netstat，避免引入 psutil。"""
    if sys.platform != "win32":
        return 0
    try:
        result = subprocess.run(
            ["netstat", "-ano", "-p", "tcp"], capture_output=True,
            text=True, encoding="mbcs", errors="ignore", check=False,
        )
    except OSError:
        return 0
    suffixes = (f":{port}", f"]:{port}")
    for line in result.stdout.splitlines():
        parts = line.split()
        # TCP / 本地地址 / 远程地址 / 状态（语言无关）/ PID
        if len(parts) >= 5 and parts[0].upper() == "TCP" and parts[1].endswith(suffixes):
            try:
                return int(parts[-1])
            except ValueError:
                continue
    return 0


def _can_bind(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            probe.bind(("127.0.0.1", port))
        except OSError:
            return False
    return True


def _first_available_port(requested_port: int) -> int:
    for port in range(requested_port, requested_port + 20):
        if _can_bind(port):
            return port
    raise RuntimeError(f"端口 {requested_port}-{requested_port + 19} 均被占用")


def _prepare_instance(requested_port: int) -> int:
    """回收死亡锁/可确认的旧 Flask 实例，并保证本次选到可监听端口。"""
    locked_pid = _read_lock_pid()
    if locked_pid and not _process_is_alive(locked_pid):
        _remove_lock_if_owned(locked_pid)
        print(f"[提示] 已清理死进程遗留的应用锁（pid {locked_pid}）。")
        locked_pid = 0

    listener_pid = _listening_pid(requested_port)
    if listener_pid and _process_matches(listener_pid):
        # PID 正在监听指定端口，且命令行可识别为本工具。即使此前一次失败
        # 启动覆盖了 app.pid，也能回收遗留 Flask 服务；其它程序不受影响。
        try:
            os.kill(listener_pid, signal.SIGTERM)
        except OSError:
            pass
        deadline = time.monotonic() + 3.0
        while time.monotonic() < deadline and not _can_bind(requested_port):
            time.sleep(0.1)
        if _can_bind(requested_port):
            _remove_lock_if_owned(listener_pid)
            print(f"[提示] 已关闭旧 Flask 实例（pid {listener_pid}），当前实例接管端口。")

    selected_port = _first_available_port(requested_port)
    if selected_port != requested_port:
        print(f"[提示] 端口 {requested_port} 被其它程序占用，已改用 {selected_port}。")
    return selected_port


def _write_pid(port: int) -> None:
    try:
        PID_FILE.parent.mkdir(parents=True, exist_ok=True)
        PID_FILE.write_text(
            json.dumps({"pid": os.getpid(), "port": port}, ensure_ascii=False),
            encoding="utf-8",
        )
    except OSError:
        pass


def main() -> int:
    import multiprocessing

    multiprocessing.freeze_support()   # 打包 EXE 中多进程渲染的必需入口

    parser = argparse.ArgumentParser(description="基础数据审核工具（Flask 版）")
    parser.add_argument("--port", type=int, default=8750, help="本地服务端口")
    parser.add_argument("--no-browser", action="store_true", help="不自动打开浏览器窗口")
    args = parser.parse_args()

    try:
        args.port = _prepare_instance(args.port)
    except RuntimeError as exc:
        print(f"[错误] {exc}")
        return 1
    _write_pid(args.port)
    atexit.register(_remove_lock_if_owned, os.getpid())

    from app import app

    url = f"http://127.0.0.1:{args.port}/"
    if not args.no_browser:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=args.port, debug=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
