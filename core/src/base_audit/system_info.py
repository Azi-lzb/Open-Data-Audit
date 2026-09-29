"""本机运行环境信息（高级设置展示与 S1F1 预检日志共用）。

用户报障时可直接复制环境摘要一行，免去逐项询问系统配置；glibc/Python/
LibreOffice 三项正是离线包选型与"引擎找得到但起不动"类问题的定位依据。
"""

from __future__ import annotations

import os
import platform
import sys

_SUMMARY: dict[str, str] | None = None


def _os_pretty_name() -> str:
    try:
        with open("/etc/os-release", encoding="utf-8") as fh:
            for line in fh:
                if line.startswith("PRETTY_NAME="):
                    return line.split("=", 1)[1].strip().strip('"')
    except OSError:
        pass
    return f"{platform.system()} {platform.release()}"


def _glibc_version() -> str:
    """本机 glibc 版本；Windows 无此概念。

    优先 os.confstr；deepin/UOS 的 Python 常不注册该配置名（实测 ValueError），
    回退解析 `ldd --version` 首行。
    """
    if sys.platform == "win32":
        return "不适用（Windows）"
    try:
        value = os.confstr("CS_GNU_GLIBC_VERSION")
        if value and value.strip():
            return value.split()[-1]
    except (AttributeError, OSError, ValueError):
        pass
    import re
    import subprocess
    try:
        completed = subprocess.run(
            ["ldd", "--version"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, timeout=10, check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return "未知（可运行 ldd --version 查看）"
    match = re.search(r"([0-9]+\.[0-9]+)", completed.stdout.splitlines()[0] if completed.stdout else "")
    if match:
        # 首行形如 "ldd (Debian GLIBC 2.38-6deepin27) 2.38"，取最后一个版本号。
        tail = completed.stdout.splitlines()[0].rstrip().split()[-1]
        return tail if re.fullmatch(r"[0-9]+(\.[0-9]+)+", tail) else match.group(1)
    return "未知（可运行 ldd --version 查看）"


def _python_description() -> str:
    """系统内置 python3（PATH 上的）版本；随包运行时版本固定，不占此行。"""
    import re
    import subprocess
    try:
        completed = subprocess.run(
            ["python3", "--version"], stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT, text=True, timeout=10, check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return f"{sys.version.split()[0]}（当前进程）"
    match = re.search(r"([0-9]+\.[0-9]+\.[0-9]+)", completed.stdout or "")
    if match:
        return f"{match.group(1)}（系统 python3）"
    return f"{sys.version.split()[0]}（当前进程）"


def _probe_libreoffice_version(command_prefix: tuple[str, ...]) -> str:
    """启动一次 --version 取真实版本号；失败返回空串（不阻塞启动）。"""
    import re
    import subprocess

    from .engines.libreoffice import _system_command_env
    try:
        completed = subprocess.run(
            [*command_prefix, "--version"], stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT, text=True, timeout=15, check=False,
            env=_system_command_env(),
        )
    except (OSError, subprocess.TimeoutExpired):
        return ""
    for line in (completed.stdout or "").splitlines():
        match = re.search(r"LibreOffice\s+[0-9][0-9.]*", line)
        if match:
            return match.group(0).split()[-1]
    return ""


def _libreoffice_summary(probe_version: bool = False) -> str:
    """引擎摘要；probe_version=True 时实启一次 --version 取版本号。

    玲珑版 ll-cli 冷启动可达 10 秒，默认只显示来源（秒回），由调用方
    在后台线程里以 probe_version=True 补全，避免阻塞应用启动。
    """
    if sys.platform == "win32":
        return "Windows 由 Excel/WPS 提供（运行时实测）"
    try:
        from .engines.libreoffice import find_calc_engine
        engine = find_calc_engine()
    except Exception:
        engine = None
    if engine is None:
        return "未找到"
    source = engine.source or "已安装"
    if not probe_version:
        return f"LibreOffice（{source}，版本检测中…）"
    version = _probe_libreoffice_version(engine.command_prefix)
    # display 是启动命令（如 ll-cli run …），不适合给人看；改为 版本+来源。
    return f"LibreOffice {version}（{source}）" if version else f"LibreOffice（{source}，版本未识别）"


def system_environment_summary(probe_version: bool = False) -> dict[str, str]:
    """进程内缓存的环境四项：os / glibc / python / libreoffice。

    probe_version=True 时重探 LibreOffice 版本并更新缓存（慢，供后台线程用）。
    """
    global _SUMMARY
    if _SUMMARY is None or probe_version:
        _SUMMARY = {
            "os": _os_pretty_name(),
            "glibc": _glibc_version(),
            "python": _python_description(),
            "libreoffice": _libreoffice_summary(probe_version=probe_version),
        }
    return dict(_SUMMARY)


def system_environment_text() -> str:
    info = system_environment_summary()
    return (
        f"系统：{info['os']}；glibc：{info['glibc']}；"
        f"Python：{info['python']}；LibreOffice：{info['libreoffice']}"
    )
