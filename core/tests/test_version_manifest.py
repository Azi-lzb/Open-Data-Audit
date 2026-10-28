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
    "S5-F01", "S5-F02", "S5-F03", "S5-F04",
}
EXPECTED_CONFIGS = {
    "S1-CFG-01", "S2-CFG-01", "S2-CFG-02",
    "S3-CFG-00", "S3-CFG-01", "S3-CFG-02", "S3-CFG-03",
    "S5-CFG-01",
}
VERSION_RE = re.compile(r"^\d+\.\d+\.\d+$")
PRODUCT_VERSION_RE = re.compile(r"^\d{2}\.\d+\.\d+\.\d+$")


@pytest.fixture(scope="module")
def manifest() -> dict:
    assert MANIFEST.is_file(), f"缺少版本清单：{MANIFEST}"
    return json.loads(MANIFEST.read_text(encoding="utf-8"))


def test_manifest_is_valid_json_with_core_sections(manifest: dict) -> None:
    for section in ("product", "systems", "features", "configs"):
        assert section in manifest, section
    product = manifest["product"]
    assert product["id"] == "V3"
    assert product["name"] == "审核工具"
    assert PRODUCT_VERSION_RE.fullmatch(product["version"])
    assert product["changelog"] == "版本管理/产品发布/CHANGELOG.md"
    assert (ROOT / product["changelog"]).is_file()
    release_log = (ROOT / product["changelog"]).read_text(encoding="utf-8")
    assert any(line.startswith(f"## V{product['version']} —") for line in release_log.splitlines())
    assert set(manifest["systems"]) == {"S1", "S2", "S3", "S5"}


def test_all_15_features_and_8_configs_registered(manifest: dict) -> None:
    assert set(manifest["features"]) == EXPECTED_FEATURES
    assert set(manifest["configs"]) == EXPECTED_CONFIGS


EXPECTED_VERSIONS = {
    # 2026-10-13：补齐未登记变更。S1 13 列/七段式协议升 MAJOR；
    # S3-F05/CFG-03 空白默认禁用升 MAJOR，其他 S3 按各自影响升级。
    # v1.0.1：源码按 S1/S2/S3 归位（Refactor）；S1-F05 为实测发现的
    # COM 路径 NameError Bugfix；S1-F01 v1.2.0 为日志设置拆分（MINOR），并保留
    # v1.1.1 条件格式多区域判定修复；S1-F02/F03/F04 v1.1.0 为日志导出体验兼容；
    # S2-F01 为八段占位符取数 Bugfix（用户报障）；
    # 其余维持基线。
    # S2-F01 v10.6.0：新增 python-calamine 0.3.1 快速读取引擎（双引擎回退），1098 文件读取 4.26s→1.02s；
    # S2-F01 v10.6.1：普通报表支持显式“行名称”表头；按“数据日期”+“行名称”共同识别时序格式。
    # v10.6.2：匹配行空白/非数值按0并给出软性提示，空行名称不阻断，完成状态显示提示数量；
    # v10.6.3：时序布局无效/缺失日期行提示后忽略，仍精确匹配本期完整日期。
    # v10.6.4：行名称为空时仅指标列有值才提示；时序元数据不再误报为空项目行。
    # S2-F01 v10.5.0：环比结果在指标代码后输出数据属性列；通用设置新增可选“文字变动默认级别”
    # （文字指标变化与文字单侧缺失，缺省硬性）及“本期缺失默认级别”“上期缺失默认级别”
    # （仅非文字指标，按“本期无，上期有”/“本期有，上期无”方向分别控制，含该侧无数值/为0按无处理）；
    # S2-CFG-01 v8.4.0：通用设置新增上述三行可选设置并带硬性/软性下拉（B8:B10）；
    # 全部设置行缺失时按默认执行（单位=元、隐藏=否、级别=硬性）并提示补齐，不再必填；
    # S2-F01 v10.4.2：运行记录新增阶段/分段结果；固定三表导出切换Direct OOXML并对拍；v10.4.1移除独立进度框；
    # 前序 v10.4.0：实际阶段/数量进度与耗时、固定列表表头和批量导入/导出性能优化；
    # 前序 v10.3.1：外部核对缺数审核说明区分已映射报表项目和缺少机构参照映射；
    # 前序 v10.3.0：对比对象新增单边及缺少期别筛选，缺失位置明确标注本期/上期；
    # 前序 v10.2.0：对比对象新增机构、信用代码、表单的组合显示筛选，不改变审核范围；
    # 前序 v10.1.0：配对预览保留并显示独立信用代码，名称缺失或单边仍可定位；
    # 前序 v10.0.0：外部核对结果在信用代码后插入报表项目，原后续输出列顺移；
    # 前序 v9.3.1：预览及正式导入共用文件名/表名身份，空表和正文占位符不再丢失机构字段；
    # 前序 v9.3.0：外部数据按表头兼容时序布局，按本期报送完整日期匹配且区分缺数/真实0；
    # 前序 v9.2.0：集中系统数据缺少可识别机构行标签列时按0兜底并继续外部核对；
    # 前序 v9.1.0：外部机构映射优先读取机构参照表，并兼容旧参照表；
    # 前序 v9.0.0：环比结果删除规则编号列；未触发规则或未满足门槛时说明留空，计算过程不再给出命中/通过判定；
    # 前序 v8.0.0：审核级别输出契约统一为硬性/软性；旧配置值继续兼容读取；
    # S2-CFG-01 v8.3.0：三张规则表增加硬性/软性下拉与筛选；
    # 前序 S2-F01 v7.1.0：配置检查逐项展示已检查内容与范围说明；
    # 前序 v7.0.0：外部核对支持六种比较方式，明确左右值语义并以输出单位容差判断；
    # S2-CFG-01 v8.0.0：外部核对规则删除“外部系统”列并调整为新协议顺序；
    # 前序 S2-F01 v6.1.0：按通用设置可选隐藏两类正常结果行；v6.0.2 从主配置读取单位设置；
    # v6.0.0：源数据与外部金额先精确换算到输出单位后再计算；
    # S1/S3 各配置 PATCH：同步补充工作簿内字段说明与高级版教程。
    # S1-F01 v2.4.0：源表被模板覆盖即可匹配；v2.5.0 增加待处理文件筛选；v2.6.0 增加审核结果 Rust/OpenPyXL 写入路由。
    # v2.6.3：外部文件快路径不支持复杂外部表时降级 openpyxl，写入器全失败也只标记该文件并继续链路。
    # S1-F02 v1.2.2：模板多出的表不显示部分匹配；源表区分兜底可用性；v1.3.0 增加 Calamine/Rust 读写路由；v1.3.1 修正模板外隐藏工作表提示。
    "S1-F01": "2.6.3", "S1-F02": "1.3.1", "S1-F03": "1.1.0",
    "S1-F04": "1.1.0", "S1-F05": "1.0.1",
    # S5-F01 v2.2.x unified external-layout reads, row diagnostics, and the F1-F4 card.
    # v2.3.0-v2.7.1 add configured-path, configuration-check, and matching improvements;
    # v3.0.0 changes F1 output to one consolidated workbook and removes the ledger/copy protocol;
    # v3.0.1 caches immutable audit-expression ASTs and dictionary/path validation indexes.
    "S5-F01": "3.0.1",
    # S5-F02/F03/F04 are separately registered hidden configuration tools.
    "S5-F02": "1.0.0", "S5-F03": "1.0.0", "S5-F04": "1.0.0",
    "S2-F01": "10.8.1", "S3-F01": "1.6.0", "S3-F02": "1.1.0",
    "S3-F03": "1.1.0", "S3-F04": "1.3.0", "S3-F05": "2.0.0",
    "S1-CFG-01": "2.0.0", "S2-CFG-01": "8.4.0", "S2-CFG-02": "1.0.0",
    "S5-CFG-01": "2.0.0",
    "S3-CFG-00": "2.0.2", "S3-CFG-01": "3.0.3",
    "S3-CFG-02": "1.0.3", "S3-CFG-03": "2.0.0",
}


def test_component_ids_are_unique_and_versions_wellformed(manifest: dict) -> None:
    ids = [*manifest["features"], *manifest["configs"]]
    assert len(ids) == len(set(ids)) == 23
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
    "S1-CFG-01": "2.0.0",
    "S2-CFG-01": "8.4.0",
    "S2-CFG-02": "1.0.0",
    "S5-CFG-01": "2.0.0",
    "S3-CFG-00": "2.0.2",
    # v3.0.0：3.1 表达式规则输入列从“触发方式”升级为“是否取反（是/否）”。
    "S3-CFG-01": "3.0.3",
    "S3-CFG-02": "1.0.3",
    "S3-CFG-03": "2.0.0",
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
