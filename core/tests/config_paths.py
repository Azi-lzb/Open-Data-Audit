"""现成测试配置逐文件优先读取本机真实配置，缺失时回退公开配置。"""

from __future__ import annotations

from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]


def resolve_test_config(relative_path: str | Path, *, repo_root: Path = REPO_ROOT) -> Path:
    """保留相对目录（含默认配置）；只选择路径，不创建或修改工作簿。"""
    relative = Path(relative_path)
    if relative.is_absolute() or ".." in relative.parts:
        raise ValueError("测试配置必须使用配置目录内的相对路径")
    private = repo_root / "config-real" / relative
    if private.is_file():
        return private
    return repo_root / "config" / relative
