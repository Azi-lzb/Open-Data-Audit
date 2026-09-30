"""各外壳共用的产品名称、发布版本、图标和产物命名。"""

from __future__ import annotations

import argparse
from html import escape
import json
from pathlib import Path
import re
import sys


def manifest_path() -> Path:
    """源码和冻结包均读取构建时的同一份版本清单。"""
    if getattr(sys, "frozen", False):
        root = Path(sys._MEIPASS)
    else:
        root = Path(__file__).resolve().parents[3]
    return root / "版本管理" / "VERSION_MANIFEST.json"


def load_product_info(path: Path | None = None) -> dict[str, str]:
    source = Path(path) if path is not None else manifest_path()
    try:
        product = json.loads(source.read_text(encoding="utf-8"))["product"]
        product_id, name, version = (product[key] for key in ("id", "name", "version"))
        if not all(isinstance(value, str) and value for value in (product_id, name, version)):
            raise ValueError("产品字段必须为非空文字")
        if not re.fullmatch(r"\d{2}\.\d+\.\d+\.\d+", version):
            raise ValueError("产品版本必须为 年份.大版本.中版本.小版本（年份两位）")
        if any(int(part) > 65535 for part in version.split(".")):
            raise ValueError("产品版本的各段不能超过 65535")
        if re.search(r'[<>:"/\\|?*\x00-\x1f]', name) or name.endswith((".", " ")):
            raise ValueError("产品名称含文件名不支持的字符")
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise RuntimeError(
            f"产品版本清单不可用：{source.name}。请使用包含完整版本清单的程序包。"
        ) from exc
    return {"id": product_id, "name": name, "version": version, "title": f"{name} V{version}"}


def app_title() -> str:
    return load_product_info()["title"]


def executable_stem(shell: str, win7: bool = False) -> str:
    if shell not in ("pywebview", "flask", "linux"):
        raise ValueError(f"未知程序外壳：{shell}")
    if shell == "linux" and win7:
        raise ValueError("Linux 程序不支持 Win7 标记")
    info = load_product_info()
    parts = [info["name"]]
    if shell == "flask":
        parts.append("Flask")
    if win7:
        parts.append("Win7兼容")
    parts.append("V" + info["version"])
    return "_".join(parts)


def icon_path(suffix: str = ".png") -> Path:
    if suffix not in (".png", ".ico"):
        raise ValueError(f"不支持的图标格式：{suffix}")
    if getattr(sys, "frozen", False):
        root = Path(sys._MEIPASS) / "assets"
    else:
        root = Path(__file__).resolve().parents[2] / "frontend" / "assets"
    return root / ("app-icon" + suffix)


def render_app_page(source: str, icon_url: str | None = None) -> str:
    """为同一前端页面注入名称/版本，临时 pywebview 页面使用绝对图标 URI。"""
    info = load_product_info()
    return source.replace("__AUDIT_APP_TITLE__", escape(info["title"])).replace(
        "__AUDIT_ICON_URL__", escape(icon_url or icon_path().as_uri(), quote=True)
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="读取审核工具的统一产品信息")
    field = parser.add_mutually_exclusive_group(required=True)
    for option in ("version", "title", "linux-name", "manifest"):
        field.add_argument("--" + option, action="store_true")
    args = parser.parse_args()
    if args.manifest:
        value = str(manifest_path())
    elif args.linux_name:
        value = executable_stem("linux")
    else:
        value = load_product_info()["version" if args.version else "title"]
    print(value)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
