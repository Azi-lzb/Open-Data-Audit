"""Flask 外壳打包入口（unified 统一核心）。

用法：python build_exe.py [--win7]
由 windows-build/build-modern.bat / build-win7.bat 调用；bat 负责选择解释器。

产物布局（由 FLASK_DIST_DIR 指定，可为独立目录或 dist/windows-时间戳）：
  产品名_外壳标识_V产品版本.exe
  core/                              ← 统一业务核心（src + frontend + 空 data）
  config/                            ← 正式配置副本
模板、历史库、用户设置均为 EXE 同级外部目录，升级不覆盖。
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
# 统一核心的唯一定义：仓库根目录 core\（与 shell-flask 同级）。
# 不再回退到 shell 自己的 core\：那样会在升级/换目录时静默打包到陈旧副本，
# 两个发行包从此分叉（真实事故：pywebview 快照落后 3 分钟导致旧代码发行）。
_REPO_CORE = ROOT.parent / "core"
if not (_REPO_CORE / "src" / "base_audit").is_dir():
    raise SystemExit(
        f"[错误] 未找到仓库根目录统一核心：{_REPO_CORE}\\src\\base_audit\n"
        "       打包脚本必须与 core\\ 处于同一仓库根目录下。")
CORE = _REPO_CORE
DIST = Path(os.environ.get("FLASK_DIST_DIR") or ROOT / "dist")
CONFIG_TEMPLATE = ROOT.parent / "config"
if str(CORE / "src") not in sys.path:
    sys.path.insert(0, str(CORE / "src"))

from base_audit.app_identity import app_title, executable_stem, icon_path, load_product_info, manifest_path

COMMON_HIDDEN = [
    "base_audit.excel_com",
    "base_audit.service",
    "base_audit.merge_org",
    "base_audit.web_app",
    "base_audit.region_summary",
    "base_audit.engines.libreoffice_adapter",
    "openpyxl",
    "pythoncom",
    "pywintypes",
    "flask",
    # 文件选择方式可在高级设置切到 Tk；动态导入时也必须把 Tk runtime 带入 EXE。
    "tkinter",
    "tkinter.filedialog",
]

COMMON_EXCLUDES = [
    "doctest", "pydoc", "pytest", "unittest", "numpy", "pandas",
    "lxml", "lxml.etree", "lxml.objectify", "lxml.html", "lxml.isoschematron",
    "pythonwin", "win32ui",
]


def _pywin32_args() -> list[str]:
    """把 pywin32 的 Python 包树经 add-data 直接送进冻结包（_MEIPASS）。

    原因：本工程构建环境下 PyInstaller 4.10 的 modulegraph 对 `win32com`
    一律报 Hidden import not found（连最小 hello 构建、显式 --paths 指到
    site-packages、--collect-all 都不行；见 GPT修改记录 20260921_123000），
    而 add-data 是纯文件复制不经模块分析，运行期 _MEIPASS 在 sys.path 上，
    `import win32com` 直接命中。pythoncom/pywintypes 的 DLL 本体已由
    hidden-import 正常进入（warn 文件之外单独验证过），这里只补 Python 树。
    """
    try:
        import win32com

        sp = Path(win32com.__file__).resolve().parent.parent
    except ImportError:
        return []
    args: list[str] = []
    for item in ("win32com", "win32comext"):
        if (sp / item).is_dir():
            args += ["--add-data", f"{sp / item}{os.pathsep}{item}"]
    win32_dir = sp / "win32"
    if win32_dir.is_dir():
        import glob as _glob

        for pattern, tool in (("*.pyd", "--add-binary"), ("*.py", "--add-data")):
            if _glob.glob(str(win32_dir / pattern)):
                args += [tool, f"{win32_dir / pattern}{os.pathsep}."]
        # pywin32.pth 在正常安装里把 win32\lib 加进 sys.path——那里的模块
        # （winerror、win32timezone 等）都是顶层导入语义；冻结包里平铺到
        # _MEIPASS 根即可获得同样的可见性。
        if (win32_dir / "lib").is_dir():
            args += ["--add-data", f"{win32_dir / 'lib' / '*.py'}{os.pathsep}."]
    loader = sp / "pythoncom.py"
    if loader.is_file():
        args += ["--add-data", f"{loader}{os.pathsep}."]
    return args


def build(win7: bool) -> int:
    if not (CORE / "src" / "base_audit").is_dir():
        raise SystemExit(f"[错误] 未找到统一核心：{CORE}")
    if not (CORE / "frontend" / "web" / "index.html").is_file():
        raise SystemExit("[错误] 未找到前端 index.html")
    if win7 and (
        not os.environ.get("FLASK_DIST_DIR")
        or DIST.resolve().parent != (ROOT / "dist").resolve()
        or not DIST.name.startswith(("win7-", "windows-"))
    ):
        raise SystemExit("[错误] Win7 构建必须输出到 dist\\win7-时间戳\\ 或 dist\\windows-时间戳\\。")

    product = load_product_info()
    name = executable_stem("flask", win7=win7)
    icon = icon_path(".ico")
    manifest = manifest_path()
    assets = CORE / "frontend" / "assets"
    for label, path in (("应用图标", icon), ("版本清单", manifest), ("应用资产目录", assets)):
        if not path.exists():
            raise SystemExit(f"[错误] 未找到{label}：{path}")
    DIST.mkdir(parents=True, exist_ok=True)

    args = [
        # 保留控制台：Flask 服务端日志（端口占用、异常 traceback）直接可见。
        "--onefile", "--clean", "--noconfirm",
        # 整包收集 base_audit：workflow 适配层（audit_handlers）虽被 runner
        # 引用，PyInstaller 4.10 的模块分析仍漏掉了它的顶层依赖
        # external_sheet_writer（真机 Win7 运行 DAG 报 No module named）。
        # 逐个 hidden-import 太脆，base_audit 全是纯 Python，直接整包收集。
        "--collect-submodules", "base_audit",
        "--name", name,
        "--paths", str(CORE / "src"),
        "--distpath", str(DIST),
        "--workpath", str(ROOT / "build"),
        "--specpath", str(ROOT / "build"),
        "--icon", str(icon),
        "--add-data", f"{manifest}{os.pathsep}版本管理",
        "--add-data", f"{assets}{os.pathsep}assets",
    ]
    if sys.platform == "win32":
        version_file = write_version_file(name, product)
        args += ["--version-file", str(version_file)]
    for mod in COMMON_HIDDEN:
        args += ["--hidden-import", mod]
    args += _pywin32_args()
    for mod in COMMON_EXCLUDES:
        args += ["--exclude-module", mod]
    args.append(str(ROOT / "run.py"))

    import PyInstaller.__main__

    PyInstaller.__main__.run(args)

    assemble_core()
    assemble_config_template()
    if win7:
        legacy_name = DIST / "基础数据审核工具_Flask_Win7.exe"
        if legacy_name.is_file():
            try:
                legacy_name.unlink()
                print(f"已移除旧命名的 Win7 EXE：{legacy_name.name}")
            except OSError as exc:
                print(f"[提示] 无法移除旧命名 Win7 EXE：{legacy_name.name}（{exc}）")
    print(f"\n打包完成：{DIST / (name + '.exe')}")
    print("发布时只发送当前输出目录（exe 必须与 core/ 同级）。")
    return 0


def assemble_core() -> None:
    """把统一核心复制到 dist/core（外部目录：升级 EXE 不覆盖用户数据）。"""
    target = DIST / "core"
    target.mkdir(parents=True, exist_ok=True)
    for item in ("src", "frontend"):
        src = CORE / item
        if src.is_dir():
            existing = target / item
            if existing.exists():
                shutil.rmtree(existing)
            shutil.copytree(src, target / item, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    # 新发行包创建空的外部数据目录；重建已有发行目录时保留用户数据。
    (target / "data").mkdir(exist_ok=True)
    # tests 属于开发基线，不进发行包。


def _numeric_version(version: str) -> tuple[int, int, int, int]:
    parts = version.split(".")
    if len(parts) != 4 or any(not part.isdecimal() for part in parts):
        raise SystemExit(f"[错误] Windows PE 版本必须是四段数字：{version}")
    numeric = tuple(int(part) for part in parts)
    if any(value > 65535 for value in numeric):
        raise SystemExit(f"[错误] 产品版本超出 Windows PE 范围：{version}")
    return numeric


def write_version_file(name: str, product: dict) -> Path:
    """生成 PyInstaller PE 版本资源；版本只读取统一产品清单。"""
    version = str(product["version"])
    numeric = _numeric_version(version)
    strings = {
        "CompanyName": "V3",
        "FileDescription": app_title(),
        "FileVersion": version,
        "InternalName": name,
        "OriginalFilename": name + ".exe",
        "ProductName": str(product["name"]),
        "ProductVersion": version,
    }
    rows = ",\n        ".join(
        f"StringStruct({key!r}, {value!r})"
        for key, value in strings.items()
    )
    source = (
        "VSVersionInfo(\n"
        f"    ffi=FixedFileInfo(filevers={numeric!r}, prodvers={numeric!r}),\n"
        "    kids=[\n"
        "        StringFileInfo([\n"
        "            StringTable('040904B0', [\n"
        f"                {rows}\n"
        "            ])\n"
        "        ]),\n"
        "        VarFileInfo([VarStruct('Translation', [1033, 1200])])\n"
        "    ]\n"
        ")\n"
    )
    target = ROOT / "build" / f"version-resource-{name}.txt"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(source, encoding="utf-8")
    return target


def assemble_config_template() -> None:
    """把 config 的节点配置模板复制到发行包同级 config/。"""
    if not CONFIG_TEMPLATE.is_dir():
        print(f"[提示] 未找到节点配置目录，跳过：{CONFIG_TEMPLATE}")
        return
    target = DIST / "config"
    if target.exists():
        shutil.rmtree(target)
    shutil.copytree(
        CONFIG_TEMPLATE,
        target,
        ignore=shutil.ignore_patterns("__pycache__", "*.pyc", "node_modules"),
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--win7", action="store_true", help="Win7 兼容版命名")
    raise SystemExit(build(parser.parse_args().win7))
