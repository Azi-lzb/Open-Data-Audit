"""固定 DAG 不再读取 config.xlsx 的回归测试。"""

from pathlib import Path

from openpyxl import Workbook

from base_audit.node_flow_config import apply_node_flow_config, load_dag_feature_mappings
from base_audit.workflow.defaults import build_audit_workflow, default_workflows
from base_audit.workflow.runner import build_registry
from base_audit.workflow.validation import validate_workflow


def test_legacy_node_workbook_cannot_override_fixed_dag(tmp_path: Path) -> None:
    """旧 config.xlsx 即使仍在磁盘，也不得改变节点、区域或失败策略。"""
    legacy = tmp_path / "config" / "config.xlsx"
    legacy.parent.mkdir()
    book = Workbook()
    sheet = book.active
    sheet.title = "节点配置"
    sheet.append(["流程", "顺序", "节点名称", "启用", "失败后处理", "区域组"])
    sheet.append(["汇总核查表校验", 1, "公式校验复制", "否", "停止", "自定义校验区域"])
    book.save(legacy)
    book.close()

    definition = build_audit_workflow()
    configured = apply_node_flow_config(definition, tmp_path / "config" / "逐笔统计系统_配置.xlsx")

    assert configured is definition
    formula_node = configured.node("copy-formulas")
    assert formula_node is not None
    assert formula_node.enabled is True
    assert formula_node.parameters.get("region_group") is None


def test_dag_named_ranges_always_use_code_defaults(tmp_path: Path) -> None:
    mappings = {item.name: item.range_names for item in load_dag_feature_mappings(tmp_path / "anything.xlsx")}

    assert mappings["检查校验区域"] == ("校验区域",)
    assert mappings["检查表结构区域"] == ("表结构区域",)
    assert mappings["检查汇总区域"] == ("任意行汇总区域",)
    assert mappings["检查表头区域"] == ("表头区域",)
    assert mappings["任意行汇总"] == ("任意行汇总区域",)


def test_summary_workflow_is_granular_but_user_cannot_rewire_it() -> None:
    workflow = default_workflows()["dag:汇总校验结果说明"]
    names = [node.display_name for node in workflow.nodes]
    assert names == [
        "报送文件来源", "模板装载与体检", "检查表结构区域", "检查汇总区域",
        "检查表头区域", "表结构比对", "任意行汇总", "历史说明富化",
    ]
    enrich = next(node for node in workflow.nodes if node.display_name == "历史说明富化")
    assert enrich.module_id == "summary.history_enrich"


def test_fixed_workflows_validate_for_com_and_native() -> None:
    workflow = default_workflows()["dag:汇总校验结果说明"]
    assert validate_workflow(workflow, build_registry("com"), ("com",)) == []
    assert validate_workflow(workflow, build_registry("native"), ("native",)) == []
