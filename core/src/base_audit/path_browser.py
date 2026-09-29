"""受控的本地路径浏览数据源，供浏览器内置选择器使用。"""

from __future__ import annotations

from pathlib import Path
from typing import Any


BROWSE_FILTERS: dict[str, tuple[str, ...]] = {
    # 兼容别名：历史调用方（excel 含 .xls、data 无 .xls）
    "excel": (".xlsx", ".xlsm", ".xls"),
    "data": (".csv", ".xlsx"),
    "config": (".xlsx",),
    # 精确模式（UOS 验收 ISSUE-F：与系统原生选择器口径一致）
    "excel_ooxml": (".xlsx", ".xlsm"),          # openpyxl 读取入口，不放行 .xls
    "central_data": (".csv", ".xlsx", ".xls"),  # 大集中查询导出（xlrd 纯读取兼容 .xls）
    "report_data": (".xlsx",),                 # 报表采集·大集中数据（.xls/.csv 会丢「参照表」，仅放行 .xlsx）
}


def _clean_pasted_path(path: str) -> str:
    """清洗用户粘贴的路径：去首尾空白与成对引号。

    资源管理器「复制文件地址」给出的是带 ASCII 引号的路径
    （如 "D:\待处理"）；不剥引号会被当成文件名的一部分，
    is_dir 失败后浏览器内置选择器就会退回项目根目录。
    全角引号按「开引号+闭引号」配对剥离（“…”、‘…’）。
    """
    text = str(path or "").strip()
    pairs = (('"', '"'), ("'", "'"), ("“", "”"), ("‘", "’"))
    for _ in range(3):          # 最多剥三层，防循环
        stripped = False
        for opening, closing in pairs:
            if len(text) >= 2 and text.startswith(opening) and text.endswith(closing):
                text = text[1:-1].strip()
                stripped = True
        if not stripped:
            break
    return text


def browse_directory(project_root: Path, path: str = "", mode: str = "") -> dict[str, Any]:
    """列出指定目录，或定位用户粘贴的合规文件完整路径。

    此函数只返回文件系统元数据，不打开文件，也不执行用户提供的路径。
    """
    root_path = project_root.resolve()
    root = str(root_path)
    cleaned = _clean_pasted_path(path)
    target = Path(cleaned).expanduser() if cleaned else root_path
    if not target.is_absolute():
        target = root_path / target
    try:
        target = target.resolve()
    except OSError:
        target = root_path
    extensions = BROWSE_FILTERS.get(mode or "", ())
    selected_file = ""
    if target.is_file():
        if target.name.startswith((".", "~$")) or (extensions and target.suffix.lower() not in extensions):
            return _empty(target.parent, root, "该文件不符合当前选择类型")
        selected_file = str(target)
        target = target.parent
    elif not target.is_dir():
        return _empty(root_path, root, "路径不存在或不是可访问的目录")

    directories: list[dict[str, str]] = []
    files: list[dict[str, str]] = []
    try:
        for entry in sorted(target.iterdir(), key=lambda item: item.name.casefold()):
            if entry.name.startswith((".", "~$")) or entry.name == "__pycache__":
                continue
            try:
                if entry.is_dir():
                    directories.append({"name": entry.name, "path": str(entry)})
                elif not extensions or entry.suffix.lower() in extensions:
                    size = entry.stat().st_size
                    files.append({
                        "name": entry.name,
                        "path": str(entry),
                        "size": f"{size / 1024:.0f} KB" if size >= 1024 else f"{size} B",
                    })
            except OSError:
                continue
    except OSError as exc:
        return _empty(target, root, str(exc), selected_file=selected_file)
    parent = str(target.parent) if target.parent != target else ""
    return {"path": str(target), "parent": parent, "dirs": directories, "files": files,
            "root": root, "selectedFile": selected_file, "error": ""}


def _empty(target: Path, root: str, error: str, *, selected_file: str = "") -> dict[str, Any]:
    parent = str(target.parent) if target.parent != target else ""
    return {"path": str(target), "parent": parent, "dirs": [], "files": [],
            "root": root, "selectedFile": selected_file, "error": error}
