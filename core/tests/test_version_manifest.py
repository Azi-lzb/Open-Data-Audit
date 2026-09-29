# -*- coding: utf-8 -*-
"""版本管理基线校验：VERSION_MANIFEST.json 与 CHANGELOG 目录树的一致性。

守护「V3 组件编号与版本管理」体系（rules/11）：清单合法、组件齐全、
版本号形态正确、CHANGELOG 与实现/配置路径真实存在。
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
VM_ROOT = ROOT / "版本管理"
MANIFEST = VM_ROOT / "VERSION_MANIFEST.json"

EXPECTED_FEATURES = {
    "S1-F01", "S1-F02", "S1-F03", "S1-F04", "S1-F05",
    "S2-F01",
    "S3-F01", "S3-F02", "S3-F03", "S3-F04", "S3-F05",
}
EXPECTED_CONFIGS = {
    "S1-CFG-01", "S2-CFG-01",
    "S3-CFG-00", "S3-CFG-01", "S3-CFG-02", "S3-CFG-03",
}
VERSION_RE = re.compile(r"^\d+\.\d+\.\d+$")


@pytest.fixture(scope="module")
def manifest() -> dict:
    assert MANIFEST.is_file(), f"缺少版本清单：{MANIFEST}"
    return json.loads(MANIFEST.read_text(encoding="utf-8"))


def test_manifest_is_valid_json_with_core_sections(manifest: dict) -> None:
    for section in ("product", "systems", "features", "configs"):
        assert section in manifest, section
    assert manifest["product"]["id"] == "V3"
    assert set(manifest["systems"]) == {"S1", "S2", "S3"}


def test_all_11_features_and_6_configs_registered(manifest: dict) -> None:
    assert set(manifest["features"]) == EXPECTED_FEATURES
    assert set(manifest["configs"]) == EXPECTED_CONFIGS


EXPECTED_VERSIONS = {
    # v1.0.1：源码按 S1/S2/S3 归位（Refactor）；S1-F05 为实测发现的
    # COM 路径 NameError Bugfix；S1-F01 v1.2.0 为日志设置拆分（MINOR），并保留
    # v1.1.1 条件格式多区域判定修复；S1-F02/F03/F04 v1.1.0 为日志导出体验兼容；
    # S2-F01 为八段占位符取数 Bugfix（用户报障）；
    # 其余维持基线。
    # S2-F01 v7.1.0：配置检查逐项展示已检查内容与范围说明；
    # 前序 v7.0.0：外部核对支持六种比较方式，明确左右值语义并以输出单位容差判断；
    # S2-CFG-01 v8.0.0：外部核对规则删除“外部系统”列并调整为新协议顺序；
    # 前序 S2-F01 v6.1.0：按通用设置可选隐藏两类正常结果行；v6.0.2 从主配置读取单位设置；
    # v6.0.0：源数据与外部金额先精确换算到输出单位后再计算；
    # S1/S3 各配置 PATCH：同步补充工作簿内字段说明与高级版教程。
    "S1-F01": "1.2.0", "S1-F02": "1.1.1", "S1-F03": "1.1.0",
    "S1-F04": "1.1.0", "S1-F05": "1.0.1",
    "S2-F01": "7.1.0", "S3-F01": "1.5.1", "S3-F02": "1.0.1",
    "S3-F03": "1.0.1", "S3-F04": "1.2.1", "S3-F05": "1.1.0",
    "S1-CFG-01": "1.0.5", "S2-CFG-01": "8.0.2",
    "S3-CFG-00": "2.0.2", "S3-CFG-01": "3.0.2",
    "S3-CFG-02": "1.0.3", "S3-CFG-03": "1.0.4",
}


def test_component_ids_are_unique_and_versions_wellformed(manifest: dict) -> None:
    ids = [*manifest["features"], *manifest["configs"]]
    assert len(ids) == len(set(ids)) == 17
    for cid, component in [*manifest["features"].items(), *manifest["configs"].items()]:
        assert VERSION_RE.match(component["version"]), component
        assert component["version"] == EXPECTED_VERSIONS.get(cid, "1.0.0"), cid


def test_every_component_has_changelog(manifest: dict) -> None:
    for component in [*manifest["features"].values(), *manifest["configs"].values()]:
        path = ROOT / component["changelog"]
        assert path.is_file(), path
        text = path.read_text(encoding="utf-8")
        assert "v1.0.0" in text and "Baseline" in text


EXPECTED_CONFIG_VERSIONS = {
    "S1-CFG-01": "1.0.5",
    "S2-CFG-01": "8.0.2",
    "S3-CFG-00": "2.0.2",
    # v3.0.0：3.1 表达式规则输入列从“触发方式”升级为“是否取反（是/否）”。
    "S3-CFG-01": "3.0.2",
    "S3-CFG-02": "1.0.3",
    "S3-CFG-03": "1.0.4",
}


def test_config_files_and_implementation_paths_exist(manifest: dict) -> None:
    for cid, component in manifest["configs"].items():
        assert (ROOT / component["file"]).is_file(), component["file"]
        if cid == "S2-CFG-01":
            assert "companion_files" not in component
        for companion in component.get("companion_files", []):
            assert (ROOT / companion).is_file(), companion
        assert component["version"] == EXPECTED_CONFIG_VERSIONS.get(cid, "1.0.0"), cid
    for component in manifest["features"].values():
        implementations = component["implementation"]
        if isinstance(implementations, str):
            implementations = [implementations]
        for impl in implementations:
            assert (ROOT / impl).exists(), impl


def test_s2_f01_implementation_is_systems_path(manifest: dict) -> None:
    """S2-F01 正式实现路径已去版本化：systems/s2_report_collection。"""
    entry = manifest["features"]["S2-F01"]
    assert entry["name"] == "执行报表规则"
    assert entry["implementation"] == (
        "core/src/base_audit/systems/s2_report_collection")
    assert (ROOT / "core/src/base_audit/systems/s2_report_collection").is_dir()
    # 旧 period_v2 目录必须已删除，不得留下兼容代理包。
    assert not (ROOT / "core/src/base_audit/period_v2").exists()


def test_version_management_readme_exists() -> None:
    readme = VM_ROOT / "README.md"
    assert readme.is_file()
    text = readme.read_text(encoding="utf-8")
    for keyword in ("V3", "S1", "S2", "S3", "VERSION_MANIFEST.json",
                    "CHANGELOG.md", "GPT修改记录"):
        assert keyword in text, keyword


def test_no_backups_or_outputs_inside_version_management() -> None:
    """版本管理目录只存元数据：不得混入 xlsx/zip/py 等实体文件。"""
    allowed_suffixes = {".md", ".json"}
    for path in VM_ROOT.rglob("*"):
        if path.is_file():
            assert path.suffix in allowed_suffixes, path


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-q"]))
