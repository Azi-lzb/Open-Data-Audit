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
    frozen = "随包运行时" if getattr(sys, "frozen", False) else "系统 Python"
    return f"{sys.version.split()[0]}（{frozen}）"


def _libreoffice_summary() -> str:
    if sys.platform == "win32":
        return "Windows 由 Excel/WPS 提供（运行时实测）"
    try:
        from .engines.libreoffice import find_calc_engine
        engine = find_calc_engine()
    except Exception:
        engine = None
    if engine is None:
        return "未找到"
    return engine.display


def system_environment_summary() -> dict[str, str]:
    """进程内缓存的环境四项：os / glibc / python / libreoffice。"""
    global _SUMMARY
    if _SUMMARY is None:
        _SUMMARY = {
            "os": _os_pretty_name(),
            "glibc": _glibc_version(),
            "python": _python_description(),
            "libreoffice": _libreoffice_summary(),
        }
    return dict(_SUMMARY)


def system_environment_text() -> str:
    info = system_environment_summary()
    return (
        f"系统：{info['os']}；glibc：{info['glibc']}；"
        f"Python：{info['python']}；LibreOffice：{info['libreoffice']}"
    )
