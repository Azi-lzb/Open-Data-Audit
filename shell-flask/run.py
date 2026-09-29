from __future__ import annotations

import argparse
import atexit
import json
import os
import re
import signal
import socket
import subprocess
import sys
import threading
import time
from pathlib import Path

# 旧实例 PID 锁文件（core/data 是程序运行数据目录，升级保留）。冻结版由启动器
# 指向每位用户的数据目录；源码版沿用仓库内 core/data。
_configured_project = os.environ.get("BASE_AUDIT_PROJECT_ROOT", "").strip()
_configured_core = (
    os.environ.get("BASE_AUDIT_CORE_DIR", "").strip()
    or os.environ.get("BASE_AUDIT_CORE_ROOT", "").strip()
)
if _configured_project:
    PID_CORE = Path(_configured_project).expanduser().resolve()
elif _configured_core:
    PID_CORE = Path(_configured_core).expanduser().resolve()
else:
    PID_CORE = Path(__file__).resolve().parent.parent / "core"
PID_FILE = PID_CORE / "data" / "app.pid"

_KNOWN_COMMAND_MARKERS = (
    "\\shell-flask\\", "/shell-flask/", "基础数据审核工具_flask", "run.py", "/app/base-audit-v3",
)


def _system_command_env() -> dict[str, str]:
    """PyInstaller 会设置 LD_LIBRARY_PATH；系统子进程应恢复目标机原值。"""
    env = os.environ.copy()
    if "LD_LIBRARY_PATH_ORIG" in env:
        original = env.pop("LD_LIBRARY_PATH_ORIG")
        if original:
            env["LD_LIBRARY_PATH"] = original
        else:
            env.pop("LD_LIBRARY_PATH", None)
    return env


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
                timeout=3, check=False, env=_system_command_env(),
            )
        except (OSError, subprocess.TimeoutExpired):
            return False
        cmdline = result.stdout.casefold()
        return any(marker.casefold() in cmdline for marker in _KNOWN_COMMAND_MARKERS)
    try:
        proc_dir = Path("/proc") / str(pid)
        command = proc_dir.joinpath("cmdline").read_bytes().split(b"\0")
        cwd = Path(os.readlink(proc_dir / "cwd")).resolve()
        executable = Path(os.readlink(proc_dir / "exe")).resolve()
    except OSError:
        return False
    runs_entrypoint = any(Path(arg.decode("utf-8", "ignore")).name == "run.py" for arg in command if arg)
    try:
        proc_env = proc_dir.joinpath("environ").read_bytes().split(b"\0")
    except OSError:
        proc_env = []
    runs_frozen_entrypoint = executable.name.startswith("base-audit-v3") and any(
        item == b"BASE_AUDIT_FROZEN=1" for item in proc_env
    )
    if not (runs_entrypoint or runs_frozen_entrypoint) or not cwd.name.startswith("shell-flask"):
        return False
    if runs_frozen_entrypoint:
        return True
    try:
        source = (cwd / "app.py").read_text(encoding="utf-8", errors="ignore")
    except OSError:
        source = ""
    if "base_audit.path_browser" in source and "class FlaskApi" in source:
        return True
    # 已删除到回收站的旧验收目录无法再读取 app.py；只清理目录名明确属于本项目的旧副本。
    path_text = str(cwd)
    return cwd.name.startswith("shell-flask (deleted)") and any(
        marker in path_text for marker in ("基础数据审核程序", "基础数据工具", "UOS_V3验收包_")
    )


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


def _listening_pids() -> dict[int, int]:
    """返回本机 TCP 监听端口到 PID 的映射；Linux 用 ss，Windows 用 netstat。"""
    listeners: dict[int, int] = {}
    if sys.platform == "win32":
        try:
            result = subprocess.run(
                ["netstat", "-ano", "-p", "tcp"], capture_output=True,
                text=True, encoding="mbcs", errors="ignore", check=False,
                env=_system_command_env(),
            )
        except OSError:
            return listeners
        for line in result.stdout.splitlines():
            parts = line.split()
            if len(parts) >= 5 and parts[0].upper() == "TCP" and parts[3].upper() == "LISTENING":
                try:
                    port = int(parts[1].rsplit(":", 1)[-1].rstrip("]"))
                    listeners[port] = int(parts[-1])
                except ValueError:
                    continue
        return listeners

    try:
        result = subprocess.run(
            ["ss", "-H", "-ltnp"], capture_output=True,
            text=True, encoding="utf-8", errors="ignore", timeout=3, check=False,
            env=_system_command_env(),
        )
    except (OSError, subprocess.TimeoutExpired):
        return listeners
    for line in result.stdout.splitlines():
        parts = line.split()
        if len(parts) < 6 or parts[0] != "LISTEN":
            continue
        match = re.search(r"\bpid=(\d+)\b", line)
        if not match:
            continue
        try:
            port = int(parts[3].rsplit(":", 1)[-1].rstrip("]"))
        except ValueError:
            continue
        listeners[port] = int(match.group(1))
    return listeners


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

    listeners = _listening_pids()
    for port in range(requested_port, requested_port + 20):
        listener_pid = listeners.get(port, 0)
        if not listener_pid or not _process_matches(listener_pid):
            continue
        # 接管本工具旧副本在 8750-8769 上的监听；不触碰其它程序。
        try:
            os.kill(listener_pid, signal.SIGTERM)
        except OSError:
            pass
        deadline = time.monotonic() + 3.0
        while time.monotonic() < deadline and not _can_bind(port):
            time.sleep(0.1)
        if _can_bind(port):
            _remove_lock_if_owned(listener_pid)
            print(f"[提示] 已清理旧审核服务（pid {listener_pid}，端口 {port}），当前实例重新使用首选端口。")

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
        def _open_browser() -> None:
            if sys.platform == "win32":
                import webbrowser
                webbrowser.open(url)
            else:
                opener = shutil.which("xdg-open")
                if opener:
                    subprocess.Popen([opener, url], env=_system_command_env(),
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        import shutil
        threading.Timer(1.0, _open_browser).start()
    app.run(host="127.0.0.1", port=args.port, debug=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
