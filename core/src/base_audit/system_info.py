"""本机运行环境信息（高级设置展示与 S1F1 预检日志共用）。

用户报障时可直接复制环境摘要一行，免去逐项询问系统配置；glibc/Python/
LibreOffice 三项正是离线包选型与"引擎找得到但起不动"类问题的定位依据。
Windows 上补充 Excel/WPS/COM 接管/pywin32/WebView2/openpyxl/硬件七项：
其中 COM 接管必须实启 Dispatch 才能看到真相——WPS 进程运行时会通过运行时
COM 类对象注册临时接管 Excel.Application 的激活，注册表三层（ProgID→CLSID→
LocalServer32）全部看不出差异。慢探测统一由调用方在后台线程补全，
避免阻塞应用启动（与 LibreOffice 玲珑冷启动的处理一致）。
"""

from __future__ import annotations

import os
import platform
import re
import sys
from pathlib import Path

_SUMMARY: dict[str, str] | None = None

# 卸载表里排除的 Office 周边组件（语言包、Click-to-Run 辅件等），避免误报主程序版本。
_OFFICE_COMPONENT_PATTERN = re.compile(
    r"组件|Component|MUI|Localization|Extensibility|Proof(ing)?|校对|语言包|Language Pack|加载项"
)


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
        return "Windows 由 Excel/WPS 提供（见 Excel/WPS 两行）"
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


def _read_registry(root: int, path: str, name: str | None = None) -> str | None:
    """读取单个注册表值；name=None 取默认值，失败返回 None。"""
    try:
        import winreg
        with winreg.OpenKey(root, path) as key:
            value, _ = winreg.QueryValueEx(key, name or "")
        return value if isinstance(value, str) else str(value)
    except Exception:
        return None


def _app_paths_exe(exe_name: str) -> str | None:
    """App Paths 注册的可执行文件路径；未注册返回 None。"""
    if sys.platform != "win32":
        return None
    import winreg
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe_name}",
        ) as key:
            return winreg.QueryValue(key, None) or None
    except OSError:
        return None


def _uninstall_display(match: re.Pattern[str]) -> tuple[str, str] | None:
    """在卸载表里找 DisplayName 匹配 match 的条目，返回 (名称, 版本)。

    遍历 HKLM/HKCU 的 64 位与 WOW6432Node 两套 Uninstall 键；同名主程序可能在
    两个视图各登记一次（64 位键常缺 DisplayVersion），优先返回带版本号的那个。
    只读注册表，不触发任何程序启动。
    """
    if sys.platform != "win32":
        return None
    import winreg
    fallback: tuple[str, str] | None = None
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for base in (
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        ):
            try:
                with winreg.OpenKey(root, base) as hive_key:
                    count, _, _ = winreg.QueryInfoKey(hive_key)
                    sub_keys = [winreg.EnumKey(hive_key, index) for index in range(count)]
            except OSError:
                continue
            # 实测同一句柄先读 DisplayName 再读 DisplayVersion 会偶发 WinError 6
            # （句柄无效），因此每个值都用独立的新句柄读取，不做句柄复用。
            for sub_key in sub_keys:
                full_path = rf"{base}\{sub_key}"
                display = _read_registry(root, full_path, "DisplayName")
                if not display:
                    continue
                if not match.search(display) or _OFFICE_COMPONENT_PATTERN.search(display):
                    continue
                version = _read_registry(root, full_path, "DisplayVersion") or ""
                if version:
                    return display, version
                fallback = fallback or (display, version)
    return fallback


def _excel_summary() -> str:
    """Microsoft Excel 安装版本（卸载表）；S1F1 的首选计算引擎。"""
    if sys.platform != "win32":
        return "不适用（Windows 专属）"
    found = _uninstall_display(re.compile(r"Microsoft (Office|365)"))
    if found:
        display, version = found
        display = re.sub(r"\s*-\s*[a-zA-Z]{2}(-[a-zA-Z]{2})?$", "", display)
        return f"{display}（{version}）" if version and version not in display else display
    if _app_paths_exe("excel.exe"):
        return "已安装（卸载表未登记版本）"
    return "未找到（S1F1 引擎不可用）"


def _wps_summary() -> str:
    """WPS 表格安装版本；Windows 取卸载表，UOS/Linux 只报是否在 PATH。"""
    if sys.platform == "win32":
        found = _uninstall_display(re.compile(r"^WPS Office"))
        if found:
            display, version = found
            return f"{display}（{version}）" if version and version not in display else display
        if _app_paths_exe("et.exe"):
            return "已安装（卸载表未登记版本）"
        return "未找到"
    import shutil
    exe = shutil.which("wps")
    return f"Linux 版：{exe}" if exe else "未找到"


def _resolve_progid_server(progid: str) -> str | None:
    """按 COM 真实解析顺序取 ProgID 的服务器路径（ProgID→CLSID→LocalServer32）。

    TreatAs 优先于 ProgID 自身的 CLSID。注意：这只反映注册表层，WPS 进程
    运行时接管（见 _probe_com_activation）在此处不可见。
    """
    if sys.platform != "win32":
        return None
    import winreg
    clsid = _read_registry(winreg.HKEY_CLASSES_ROOT, rf"{progid}\TreatAs")
    if not clsid:
        clsid = _read_registry(winreg.HKEY_CLASSES_ROOT, rf"{progid}\CLSID")
    if not clsid:
        return None
    for sub_key in ("LocalServer32", "InprocServer32"):
        server = _read_registry(winreg.HKEY_CLASSES_ROOT, rf"CLSID\{clsid}\{sub_key}")
        if server:
            return server
    return None


def _com_registry_summary() -> str:
    """COM 接管初始值：注册表层解析结果，秒回、无副作用。

    WPS 运行中临时接管 Excel.Application 时本值仍显示真 Excel；实启结果由
    web_app 后台线程以 probe_com=True 补全覆盖。
    """
    if sys.platform != "win32":
        return "不适用（Windows 专属）"
    server = _resolve_progid_server("Excel.Application") or ""
    upper = server.upper()
    if "KINGSOFT" in upper or "WPS" in upper:
        target = "WPS 表格"
    elif "EXCEL.EXE" in upper or "MICROSOFT OFFICE" in upper:
        target = "Microsoft Excel"
    elif server:
        return f"注册表解析到未知来源（{server}）"
    else:
        return "未找到 Excel.Application 注册（COM 引擎不可用）"
    return f"注册表解析到 {target}"


def _probe_com_activation() -> str:
    """实启一次 Excel.Application，返回真实激活目标；带 pywin32 时约 1 秒。

    WPS 进程运行时会通过运行时 COM 类对象注册临时接管 Excel.Application，
    注册表三层全部看不出差异，只有实启才见真章。解析到 WPS 时可能复用
    用户已开的实例，不调用 Quit 以免关掉用户窗口；解析到 Microsoft Excel
    时仅当实例不可见（本次自动化新建）才退出，避免误关用户打开的 Excel。
    """
    try:
        import pythoncom
        import win32com.client
    except Exception:
        return "COM 组件不可用（缺 pywin32）"
    pythoncom.CoInitialize()
    try:
        try:
            app = win32com.client.Dispatch("Excel.Application")
            path = str(app.Path or "")
            version = str(app.Version or "")
            build = str(app.Build or "")
        except Exception as exc:
            return f"实启失败：{exc}"
        upper = path.upper()
        if "KINGSOFT" in upper or "WPS" in upper:
            return f"实启解析到 WPS 表格（Build {build}；WPS 运行中临时接管 Excel COM 类）"
        try:
            visible = bool(app.Visible)
        except Exception:
            visible = True
        if not visible:
            try:
                app.Quit()
            except Exception:
                pass
        return f"实启解析到 Microsoft Excel（{version}.{build}）" if version else "实启解析到 Microsoft Excel"
    finally:
        pythoncom.CoUninitialize()


def _pywin32_summary() -> str:
    """Excel/WPS COM 自动化的 Python 依赖；缺失时 COM 引擎不可用。

    pywin32 的版本是构建号（306/308/…/312），恰好形似 Python 版本号，
    加 build 前缀避免误读成"Python 3.12 掉了点"。
    """
    if sys.platform != "win32":
        return "不适用（Windows 专属）"
    try:
        from importlib.metadata import version
        return f"build {version('pywin32')}"
    except Exception:
        return "未安装（Excel/WPS COM 引擎不可用）"


_WEBVIEW2_CLIENT_GUID = "{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"


def _webview2_summary() -> str:
    """WebView2 运行时版本；pywebview 桌面外壳的渲染内核。"""
    if sys.platform != "win32":
        return "不适用（Windows 专属）"
    import winreg
    for base in (
        rf"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{_WEBVIEW2_CLIENT_GUID}",
        rf"SOFTWARE\Microsoft\EdgeUpdate\Clients\{_WEBVIEW2_CLIENT_GUID}",
    ):
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            value = _read_registry(root, base, "pv")
            if value and value not in ("", "0.0.0.0"):
                return value
    return "未找到（pywebview 桌面外壳需安装 WebView2 运行时）"


def _openpyxl_summary() -> str:
    """通用 xlsx 静态读写库版本；跨平台组件，双端都报真实版本。"""
    try:
        from importlib.metadata import version
        return version("openpyxl")
    except Exception:
        return "未安装"


def _hardware_summary() -> str:
    """CPU / 内存 / 显卡一行；性能类问题定位用。"""
    if sys.platform == "win32":
        import winreg
        cpu = (
            _read_registry(
                winreg.HKEY_LOCAL_MACHINE,
                r"HARDWARE\DESCRIPTION\System\CentralProcessor\0",
                "ProcessorNameString",
            )
            or "未知 CPU"
        ).strip()
        mem_gb = None
        try:
            import ctypes

            class _MemoryStatusEx(ctypes.Structure):
                _fields_ = [
                    ("dwLength", ctypes.c_uint32), ("dwMemoryLoad", ctypes.c_uint32),
                    ("ullTotalPhys", ctypes.c_uint64), ("ullAvailPhys", ctypes.c_uint64),
                    ("ullTotalPageFile", ctypes.c_uint64), ("ullAvailPageFile", ctypes.c_uint64),
                    ("ullTotalVirtual", ctypes.c_uint64), ("ullAvailVirtual", ctypes.c_uint64),
                    ("ullAvailExtendedVirtual", ctypes.c_uint64),
                ]

            status = _MemoryStatusEx()
            status.dwLength = ctypes.sizeof(_MemoryStatusEx)
            if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
                mem_gb = round(status.ullTotalPhys / (1024 ** 3))
        except Exception:
            mem_gb = None
        gpu = None
        # 显卡注册表类键 0000-0005；跳过 Todesk 之类的虚拟显卡。
        for index in range(6):
            desc = _read_registry(
                winreg.HKEY_LOCAL_MACHINE,
                rf"SYSTEM\CurrentControlSet\Control\Class\{{4d36e968-e325-11ce-bfc1-08002be10318}}\{index:04d}",
                "DriverDesc",
            )
            if desc and "virtual" not in desc.lower():
                gpu = desc
                if re.search(r"Radeon|NVIDIA|GeForce|Intel|Arc|UHD", desc, re.IGNORECASE):
                    break
        parts = [cpu]
        if mem_gb:
            parts.append(f"{mem_gb}GB")
        if gpu:
            parts.append(gpu)
        return " / ".join(parts)
    cpu = None
    mem_gb = None
    try:
        for line in Path("/proc/cpuinfo").read_text(encoding="utf-8", errors="replace").splitlines():
            if line.startswith("model name"):
                cpu = line.split(":", 1)[1].strip()
                break
    except OSError:
        pass
    try:
        for line in Path("/proc/meminfo").read_text(encoding="utf-8", errors="replace").splitlines():
            if line.startswith("MemTotal:"):
                match = re.search(r"([0-9]+)", line)
                if match:
                    mem_gb = round(int(match.group(1)) / (1024 * 1024))
                break
    except OSError:
        pass
    parts = [cpu or "未知 CPU"]
    if mem_gb:
        parts.append(f"{mem_gb}GB")
    return " / ".join(parts)


def _app_summary() -> str:
    """产品版本/构建/外壳一行；报障定位的第一要素。

    外壳按进程内已加载模块判断（pywebview 壳加载 webview、Flask 壳加载
    flask），直接进程（测试等）如实标注"未识别外壳"。
    """
    try:
        from .app_identity import load_product_info
        version = f"V{load_product_info()['version']}"
    except Exception:
        version = "版本清单不可用"
    build = "打包 exe" if getattr(sys, "frozen", False) else "源码运行"
    if "webview" in sys.modules:
        shell = "pywebview 桌面外壳"
    elif "flask" in sys.modules:
        shell = "Flask 浏览器外壳"
    else:
        shell = "未识别外壳"
    return f"{version}（{build} · {shell}）"


def _locale_summary() -> str:
    """区域与编码；中文文件名/内容乱码、subprocess 解码类问题的头号定位依据。

    ANSI 代码页 65001 表示系统开了"Beta: 使用 UTF-8 提供全球语言支持"，
    该开关会改变所有程序读写中文文本的默认编码行为。
    """
    parts: list[str] = []
    if sys.platform == "win32":
        import ctypes
        acp = int(ctypes.windll.kernel32.GetACP())
        parts.append("ANSI UTF-8（系统级已开）" if acp == 65001 else f"ANSI CP{acp}")
    import locale
    parts.append(f"Python偏好编码：{locale.getpreferredencoding(False)}")
    if sys.flags.utf8_mode:
        parts.append("Python UTF-8模式：开")
    lang = os.environ.get("LANG")
    if lang:
        parts.append(f"LANG={lang}")
    return "；".join(parts)


def system_environment_summary(probe_version: bool = False, probe_com: bool = False) -> dict[str, str]:
    """进程内缓存的环境摘要：os / glibc / python / libreoffice 与平台扩展项。

    probe_version=True 时重探 LibreOffice 版本并更新缓存（Linux，慢，后台线程用）；
    probe_com=True 时实启 Excel.Application 确认 COM 接管（Windows，约 1 秒起，
    后台线程用）。首次调用只做注册表/文件级快探，不启动任何 Office 进程。
    """
    global _SUMMARY
    if _SUMMARY is None:
        _SUMMARY = {
            "app": _app_summary(),
            "os": _os_pretty_name(),
            "glibc": _glibc_version(),
            "python": _python_description(),
            "libreoffice": _libreoffice_summary(probe_version=False),
            "excel": _excel_summary(),
            "wps": _wps_summary(),
            "com": _com_registry_summary(),
            "pywin32": _pywin32_summary(),
            "webview2": _webview2_summary(),
            "openpyxl": _openpyxl_summary(),
            "hardware": _hardware_summary(),
            "locale": _locale_summary(),
        }
    if probe_version:
        _SUMMARY["libreoffice"] = _libreoffice_summary(probe_version=True)
    if probe_com:
        _SUMMARY["com"] = _probe_com_activation()
    return dict(_SUMMARY)


def system_environment_text() -> str:
    info = system_environment_summary()
    return "；".join(
        (
            f"产品：{info['app']}",
            f"系统：{info['os']}",
            f"区域编码：{info['locale']}",
            f"glibc：{info['glibc']}",
            f"Python：{info['python']}",
            f"LibreOffice：{info['libreoffice']}",
            f"Excel：{info['excel']}",
            f"WPS：{info['wps']}",
            f"COM接管：{info['com']}",
            f"pywin32：{info['pywin32']}",
            f"WebView2：{info['webview2']}",
            f"openpyxl：{info['openpyxl']}",
            f"硬件：{info['hardware']}",
        )
    )
