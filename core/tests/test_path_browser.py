"""path_browser 测试：粘贴路径清洗（引号/空白）与目录列举契约。"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from base_audit.path_browser import browse_directory  # noqa: E402


def _make_tree(tmp_path: Path) -> tuple[Path, Path, Path]:
    """project_root / 目标目录(含一个 xlsx) / 目标目录路径。"""
    project = tmp_path / "project"
    (project / "core").mkdir(parents=True)
    target = tmp_path / "待处理数据"
    target.mkdir()
    (target / "单位贷款202609.xlsx").write_bytes(b"x")
    (target / "说明.txt").write_bytes(b"x")
    return project, target, target / "单位贷款202609.xlsx"


def test_quoted_directory_path_is_cleaned(tmp_path: Path) -> None:
    """资源管理器「复制文件地址」带 ASCII 引号：必须剥掉后正常列举。"""
    project, target, _ = _make_tree(tmp_path)
    result = browse_directory(project, f'"{target}"')
    assert result["error"] == ""
    assert Path(result["path"]) == target
    assert any(item["name"] == "单位贷款202609.xlsx" for item in result["files"])


def test_curly_quoted_and_padded_path_is_cleaned(tmp_path: Path) -> None:
    """全角引号与首尾空白同样剥除（输入法/网页复制常见）。"""
    project, target, _ = _make_tree(tmp_path)
    result = browse_directory(project, f'  “{target}”  ')
    assert result["error"] == ""
    assert Path(result["path"]) == target


def test_quoted_file_path_selects_file(tmp_path: Path) -> None:
    project, _, file_path = _make_tree(tmp_path)
    result = browse_directory(project, f'"{file_path}"', mode="excel")
    assert result["selectedFile"] == str(file_path)


def test_unquoted_wrong_extension_file_is_rejected(tmp_path: Path) -> None:
    project, target, _ = _make_tree(tmp_path)
    txt = target / "说明.txt"
    result = browse_directory(project, str(txt), mode="excel")
    assert result["files"] == []
    assert "不符合当前选择类型" in result["error"]


def test_missing_path_falls_back_to_root_with_error(tmp_path: Path) -> None:
    project, _, _ = _make_tree(tmp_path)
    result = browse_directory(project, str(project / "不存在"))
    assert Path(result["path"]) == project.resolve()
    assert "路径不存在" in result["error"]
