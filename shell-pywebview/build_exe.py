"""pywebview 外壳打包入口（unified 统一核心）。

用法：python build_exe.py [--win7]
由 打包pywebview.bat / 打包pywebview-Win7.bat 调用；bat 负责选择解释器。

产物布局（由 PYWEBVIEW_DIST_DIR 指定，默认 dist/）：
  产品名[_Win7兼容]_V产品版本.exe ← PyInstaller onefile，内嵌前端页面
  core/                         ← 统一业务核心（src + frontend + 空 data）
  config/                       ← 配置副本（优先 config-real，回退公开空表头）
模板、历史库、用户设置均为 EXE 同级外部目录，升级不覆盖。
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
# 统一核心的唯一定义：仓库根目录 core\（与 shell-pywebview 同级）。
# 不再回退到 shell 自己的 core\：那样会在升级/换目录时静默打包到陈旧副本，
# 两个发行包从此分叉（真实事故：pywebview 快照落后 3 分钟导致旧代码发行）。
_REPO_CORE = ROOT.parent / "core"
if not (_REPO_CORE / "src" / "base_audit").is_dir():
    raise SystemExit(
        f"[错误] 未找到仓库根目录统一核心：{_REPO_CORE}\\src\\base_audit\n"
        "       打包脚本必须与 core\\ 处于同一仓库根目录下。")
CORE = _REPO_CORE
DIST = Path(os.environ.get("PYWEBVIEW_DIST_DIR") or ROOT / "dist")
# 打包优先携带本机私有正式配置；无 config-real（如公开快照构建）时回退公开空表头。
CONFIG_TEMPLATE = next(
    (d for d in (ROOT.parent / "config-real", ROOT.parent / "config") if d.is_dir()),
    ROOT.parent / "config",
)
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
    # S2/S3 在读取旧版 .xls 时按后缀延迟导入；PyInstaller 静态分析无法发现。
    "xlrd",
    "pythoncom",
    "pywintypes",
    "win32com",
    "win32com.client",
]

COMMON_EXCLUDES = [
    "doctest", "pydoc", "pytest", "unittest", "numpy", "pandas",
    "lxml", "lxml.etree", "lxml.objectify", "lxml.html", "lxml.isoschematron",
    "pythonwin", "win32ui",
    "webview.platforms.android", "webview.platforms.cocoa",
    "webview.platforms.gtk", "webview.platforms.qt",
]


def build(win7: bool) -> int:
    if not (CORE / "src" / "base_audit").is_dir():
        raise SystemExit(f"[错误] 未找到统一核心：{CORE}")
    if not (CORE / "frontend" / "web" / "index.html").is_file():
        raise SystemExit("[错误] 未找到前端 index.html")

    product = load_product_info()
    name = executable_stem("pywebview", win7=win7)
    icon = icon_path(".ico")
    manifest = manifest_path()
    assets = CORE / "frontend" / "assets"
    for label, path in (("应用图标", icon), ("版本清单", manifest), ("应用资产目录", assets)):
        if not path.exists():
            raise SystemExit(f"[错误] 未找到{label}：{path}")
    DIST.mkdir(parents=True, exist_ok=True)

    args = [
        "--onefile", "--noconsole", "--clean",
        "--name", name,
        "--paths", str(CORE / "src"),
        # 冻结态 launch_web 从 _MEIPASS/web 读前端页面。
        "--add-data", f"{CORE / 'frontend' / 'web'}{os_pathSep()}web",
        "--add-data", f"{manifest}{os_pathSep()}版本管理",
        "--add-data", f"{assets}{os_pathSep()}assets",
        "--icon", str(icon),
        "--distpath", str(DIST),
        "--workpath", str(ROOT / "build"),
        "--specpath", str(ROOT / "build"),
    ]
    if sys.platform == "win32":
        version_file = write_version_file(name, product)
        args += ["--version-file", str(version_file)]
    # Tk 是高级设置中可选的文件选择器，不能因为默认使用系统对话框而被漏打包。
    for mod in COMMON_HIDDEN + ["webview", "webview.platforms.edgechromium", "tkinter", "tkinter.filedialog"]:
        args += ["--hidden-import", mod]
    for mod in COMMON_EXCLUDES:
        args += ["--exclude-module", mod]
    args.append(str(ROOT / "run.py"))

    import PyInstaller.__main__

    PyInstaller.__main__.run(args)

    assemble_core()
    assemble_config_template()
    print(f"\n打包完成：{DIST / (name + '.exe')}")
    print("发布时将 dist/ 整个目录发给用户（exe 必须与 core/ 同级）。")
    return 0


def os_pathSep() -> str:
    return ";" if sys.platform == "win32" else ":"


def assemble_core() -> None:
    """更新发行核心代码，同时保留已有的外部用户数据。"""
    target = DIST / "core"
    target.mkdir(parents=True, exist_ok=True)
    for item in ("src", "frontend"):
        src = CORE / item
        if src.is_dir():
            existing = target / item
            if existing.exists():
                shutil.rmtree(existing)
            shutil.copytree(src, target / item, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    # New release folders receive an empty data directory; existing data is never copied or removed.
    (target / "data").mkdir(exist_ok=True)
    history = CORE / "历史审核配置.xlsx"
    if history.is_file() and not (target / history.name).exists():
        shutil.copy2(history, target / history.name)
    config = CORE / "跨期比较配置.xlsx"
    if config.is_file() and not (target / config.name).exists():
        shutil.copy2(config, target / config.name)
    # tests 属于开发基线，不进发行包。


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


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--win7", action="store_true", help="Win7 兼容版命名")
    raise SystemExit(build(parser.parse_args().win7))
