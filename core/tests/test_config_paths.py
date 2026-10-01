"""测试配置来源优先级；用临时文件验证，不依赖本机真实业务内容。"""

from __future__ import annotations

from pathlib import Path

import pytest

from config_paths import resolve_test_config
from base_audit.node_flow_config import release_config_dir


def _file(root: Path, directory: str, relative_path: str) -> Path:
    path = root / directory / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(directory.encode())
    return path


def test_real_config_takes_priority_and_files_are_not_changed(tmp_path: Path) -> None:
    private = _file(tmp_path, "config-real", "配置.xlsx")
    public = _file(tmp_path, "config", "配置.xlsx")
    before = {path: path.read_bytes() for path in (private, public)}

    assert resolve_test_config("配置.xlsx", repo_root=tmp_path) == private
    assert {path: path.read_bytes() for path in before} == before


@pytest.mark.parametrize("private_directory_exists", [False, True])
def test_missing_real_file_falls_back_to_public(
    tmp_path: Path, private_directory_exists: bool,
) -> None:
    if private_directory_exists:
        (tmp_path / "config-real").mkdir()
    public = _file(tmp_path, "config", "配置.xlsx")

    assert resolve_test_config("配置.xlsx", repo_root=tmp_path) == public


def test_default_snapshot_is_resolved_independently(tmp_path: Path) -> None:
    private_active = _file(tmp_path, "config-real", "配置.xlsx")
    public_default = _file(tmp_path, "config", "默认配置/配置.xlsx")
    assert resolve_test_config("配置.xlsx", repo_root=tmp_path) == private_active
    assert resolve_test_config("默认配置/配置.xlsx", repo_root=tmp_path) == public_default

    private_default = _file(tmp_path, "config-real", "默认配置/配置.xlsx")
    assert resolve_test_config(Path("默认配置") / "配置.xlsx", repo_root=tmp_path) == private_default


def test_directory_with_workbook_name_does_not_count_as_real_file(tmp_path: Path) -> None:
    (tmp_path / "config-real" / "配置.xlsx").mkdir(parents=True)
    public = _file(tmp_path, "config", "配置.xlsx")

    assert resolve_test_config("配置.xlsx", repo_root=tmp_path) == public


def test_missing_both_paths_returns_public_path_without_creating_files(tmp_path: Path) -> None:
    expected = tmp_path / "config" / "配置.xlsx"

    assert resolve_test_config("配置.xlsx", repo_root=tmp_path) == expected
    assert not expected.exists()
    assert list(tmp_path.iterdir()) == []


@pytest.mark.parametrize("relative_path", ["../配置.xlsx", Path(__file__).resolve()])
def test_paths_outside_configuration_directory_are_rejected(
    tmp_path: Path, relative_path: str | Path,
) -> None:
    with pytest.raises(ValueError, match="相对路径"):
        resolve_test_config(relative_path, repo_root=tmp_path)


def test_release_config_dir_prefers_private_config_real(tmp_path: Path) -> None:
    core = tmp_path / "core"
    core.mkdir()
    (tmp_path / "config-real").mkdir()
    (tmp_path / "config").mkdir()

    # 源码布局：project_root=core，上级同时存在 config-real 与 config。
    assert release_config_dir(core) == tmp_path / "config-real"
    # 仓库根布局：root 下就有公开 config，也不得抢占上级/同级 config-real。
    assert release_config_dir(tmp_path) == tmp_path / "config-real"


def test_release_config_dir_falls_back_to_public_config(tmp_path: Path) -> None:
    core = tmp_path / "core"
    core.mkdir()
    (tmp_path / "config").mkdir()

    assert release_config_dir(core) == tmp_path / "config"
    assert release_config_dir(tmp_path) == tmp_path / "config"

    # 全部缺失时返回 root/config 默认路径，不创建目录（上级也不能有 config）。
    isolated = tmp_path / "isolated"
    empty = isolated / "pkg"
    empty.mkdir(parents=True)
    assert release_config_dir(empty) == empty / "config"
    assert not (empty / "config").exists()
