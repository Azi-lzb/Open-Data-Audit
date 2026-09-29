"""四条默认 DAG 工作流（新 schemaVersion=2；不迁移旧流程配置）。

与旧“执行流程”默认值逐项对应，节点参数即旧步骤的列语义：
- dag:汇总核查表校验 —— 审核主链（结构比对 → 外部表 → 公式复制 → 双提取
  → 汇聚 → 历史富化 → 结果工作簿 → 运行日志）；
- dag:汇总校验结果说明 —— preflight 过滤 + 区域汇总；
- dag:组合联合核查表 / dag:组合工作表 —— 共用组合工作表模块；前者绑定固定正则方案。
"""

from __future__ import annotations

from typing import Dict

from .graph import WorkflowBinding, WorkflowDefinition, WorkflowEdge, WorkflowNode
from .types import FailurePolicy

# 与旧默认流程一致的失败后处理：各步骤均为“跳过”（对应 SKIP_ITEM）。
_SKIP = FailurePolicy.SKIP_ITEM


def build_audit_workflow() -> WorkflowDefinition:
    definition = WorkflowDefinition(
        workflow_id="dag:汇总核查表校验",
        name="汇总核查表校验（DAG）",
        description="结构比对 → 外部表 → 公式复制重算 → 双路结果提取 → 历史富化 → 结果输出。",
        settings={
            "feature_names": [
                "检查校验区域", "检查表结构区域", "表结构比对", "外部文件添加",
                "公式校验复制", "校验结果提取", "条件格式结果提取", "审核结果输出",
            ],
        },
    )
    definition.nodes = [
        WorkflowNode("source-files", "source.excel_files", display_name="报送文件来源"),
        WorkflowNode("template", "source.excel_template", display_name="模板装载与体检",
                     parameters={"keep_preflight_report": False}),
        WorkflowNode("check-validation-ranges", "excel.named_range_check",
                     display_name="检查校验区域",
                     parameters={"feature_names": ["检查校验区域"]}, failure_policy=_SKIP),
        WorkflowNode("check-structure-ranges", "excel.named_range_check",
                     display_name="检查表结构区域",
                     parameters={"feature_names": ["检查表结构区域"]}, failure_policy=_SKIP),
        WorkflowNode("compare-structure", "excel.structure_compare", display_name="表结构比对",
                     parameters={"feature_name": "表结构比对"}, failure_policy=_SKIP),
        WorkflowNode("copy-files", "audit.copy_files", display_name="复制审核副本",
                     parameters={"order": 50, "feature_name": "公式校验复制",
                                 "output_name": "机构审核副本"}),
        WorkflowNode("add-external-sheets", "excel.external_sheet_copy",
                     display_name="外部文件添加",
                     parameters={"order": 40, "feature_name": "外部文件添加", "output_name": ""},
                     failure_policy=_SKIP),
        WorkflowNode("copy-formulas", "excel.formula_copy_recalculate",
                     display_name="公式校验复制",
                     parameters={"order": 50, "feature_name": "公式校验复制"}, failure_policy=_SKIP),
        WorkflowNode("extract-formula-issues", "excel.formula_issue_extract",
                     display_name="校验结果提取", parameters={"feature_name": "校验结果提取"},
                     failure_policy=_SKIP),
        WorkflowNode("extract-conditional-issues", "excel.conditional_format_issue_extract",
                     display_name="条件格式结果提取",
                     parameters={"feature_name": "条件格式结果提取"}, failure_policy=_SKIP),
        WorkflowNode("merge-audit-results", "audit.result_merge", display_name="审核结果汇聚"),
        WorkflowNode("history-enrich", "audit.history_enrich", display_name="历史说明富化"),
        WorkflowNode("write-result-workbook", "audit.result_workbook_write",
                     display_name="审核结果输出",
                     parameters={"order": 80, "feature_name": "审核结果输出",
                                 "output_name": "汇总核查表校验", "result_set": "本期审核结果"}),
        WorkflowNode("log-workbook-write", "log.workbook_write", display_name="运行日志输出",
                     parameters={"issue_feature_name": "校验结果提取"}),
    ]
    definition.edges = [
        WorkflowEdge("template", "template", "check-validation-ranges", "template"),
        WorkflowEdge("template", "template", "check-structure-ranges", "template"),
        WorkflowEdge("source-files", "source_files", "compare-structure", "source_files"),
        WorkflowEdge("template", "template", "compare-structure", "template"),
        WorkflowEdge("source-files", "source_files", "extract-conditional-issues", "source_files"),
        WorkflowEdge("template", "template", "extract-conditional-issues", "template"),
        WorkflowEdge("compare-structure", "accepted_files", "copy-files", "accepted_files"),
        WorkflowEdge("copy-files", "workbooks", "add-external-sheets", "workbooks"),
        WorkflowEdge("template", "template", "add-external-sheets", "template"),
        WorkflowEdge("add-external-sheets", "workbooks", "copy-formulas", "workbooks"),
        WorkflowEdge("template", "template", "copy-formulas", "template"),
        WorkflowEdge("copy-formulas", "workbooks", "extract-formula-issues", "workbooks"),
        WorkflowEdge("template", "template", "extract-formula-issues", "template"),
        WorkflowEdge("extract-formula-issues", "formula_issues", "merge-audit-results", "formula_issues"),
        WorkflowEdge("extract-conditional-issues", "conditional_issues", "merge-audit-results", "conditional_issues"),
        WorkflowEdge("copy-formulas", "workbooks", "merge-audit-results", "workbooks"),
        WorkflowEdge("merge-audit-results", "merged", "history-enrich", "merged"),
        WorkflowEdge("history-enrich", "enriched", "write-result-workbook", "enriched"),
        WorkflowEdge("history-enrich", "enriched", "log-workbook-write", "enriched"),
        WorkflowEdge("template", "template", "log-workbook-write", "template"),
    ]
    definition.input_bindings = [
        WorkflowBinding("history", "history-enrich", "history"),
    ]
    definition.output_bindings = [
        WorkflowBinding("summary_workbook", "write-result-workbook", "summary_workbook"),
        WorkflowBinding("issues", "history-enrich", "enriched"),
        WorkflowBinding("final_workbooks", "copy-formulas", "workbooks"),
        WorkflowBinding("structure_report", "compare-structure", "structure_report"),
    ]
    return definition


def build_summary_workflow() -> WorkflowDefinition:
    """汇总校验结果说明：细粒度节点（区域检查 → 表结构比对 → 区域汇总 → 历史富化）。

    这里刻意把旧版被 `preflight` 大包大揽的步骤拆成独立节点：检查表结构区域、
    检查汇总区域、检查表头区域、表结构比对、任意行汇总、历史说明富化，让 DAG
    真正表达业务步骤，而不是两个黑盒。
    """
    definition = WorkflowDefinition(
        workflow_id="dag:汇总校验结果说明",
        name="汇总校验结果说明（DAG）",
        description="结构比对通过的文件做任意行汇总，输出校验结果与报送说明汇总。",
        settings={
            "flow_name": "汇总校验结果说明",
            "output_name": "汇总校验结果说明",
            "feature_names": ["检查表结构区域", "检查汇总区域", "检查表头区域", "表结构比对", "任意行汇总"],
        },
    )
    definition.nodes = [
        WorkflowNode("source-files", "source.excel_files", display_name="报送文件来源"),
        # 汇总模板不含可复制的校验公式，节点走 summary_template 模式：只按汇总
        # 区域 + 表头区域建立定义，不要求“校验区域”。
        WorkflowNode("template", "source.excel_template", display_name="模板装载与体检",
                     parameters={"keep_preflight_report": False, "summary_template": True}),
        WorkflowNode("check-structure-range", "excel.named_range_check",
                     display_name="检查表结构区域",
                     parameters={"feature_names": ["检查表结构区域"]}, failure_policy=_SKIP),
        WorkflowNode("check-summary-range", "excel.named_range_check",
                     display_name="检查汇总区域",
                     parameters={"feature_names": ["检查汇总区域"]}, failure_policy=_SKIP),
        WorkflowNode("check-header-range", "excel.named_range_check",
                     display_name="检查表头区域",
                     parameters={"feature_names": ["检查表头区域"]}, failure_policy=_SKIP),
        WorkflowNode("compare-structure", "excel.structure_compare", display_name="表结构比对",
                     parameters={"feature_name": "表结构比对"}, failure_policy=_SKIP),
        WorkflowNode("region-summary", "summary.region_summary", display_name="任意行汇总",
                     parameters={"flow_name": "汇总校验结果说明", "output_name": "汇总校验结果说明",
                                 "summary_feature_names": ["任意行汇总"]}),
        WorkflowNode("history-enrich", "summary.history_enrich", display_name="历史说明富化",
                     parameters={"flow_name": "汇总校验结果说明"}, failure_policy=_SKIP),
    ]
    definition.edges = [
        WorkflowEdge("template", "template", "check-structure-range", "template"),
        WorkflowEdge("template", "template", "check-summary-range", "template"),
        WorkflowEdge("template", "template", "check-header-range", "template"),
        WorkflowEdge("source-files", "source_files", "compare-structure", "source_files"),
        WorkflowEdge("template", "template", "compare-structure", "template"),
        WorkflowEdge("compare-structure", "accepted_files", "region-summary", "accepted_files"),
        WorkflowEdge("compare-structure", "structure_report", "region-summary", "structure_report"),
        WorkflowEdge("region-summary", "summary_workbook", "history-enrich", "summary_workbook"),
    ]
    definition.input_bindings = [
        WorkflowBinding("runtime", "source-files", "runtime"),
        WorkflowBinding("runtime", "template", "runtime"),
        WorkflowBinding("runtime", "region-summary", "runtime"),
        WorkflowBinding("runtime", "history-enrich", "runtime"),
    ]
    definition.output_bindings = [
        WorkflowBinding("summary_result", "region-summary", "result"),
        WorkflowBinding("summary_workbook", "history-enrich", "summary_workbook"),
    ]
    return definition


def build_merge_org_workflow() -> WorkflowDefinition:
    definition = WorkflowDefinition(
        workflow_id="dag:组合联合核查表",
        name="组合工作表（核查表）",
        description="组合工作表的固定预置：按标准文件名正则合并同一机构文件。",
        settings={"flow_name": "组合联合核查表", "preset": "standard-combine-regex"},
    )
    definition.nodes = [
        WorkflowNode("combine-sheets", "excel.combine_sheets",
                     display_name="组合工作表（联合核查表预置）",
                     parameters={
                         "flow_name": "组合联合核查表",
                         "plan_id": "组合工作表（核查表）",
                     }),
    ]
    definition.input_bindings = [WorkflowBinding("runtime", "combine-sheets", "runtime")]
    definition.output_bindings = [
        WorkflowBinding("result", "combine-sheets", "result"),
        WorkflowBinding("output_files", "combine-sheets", "output_files"),
    ]
    return definition


def build_combine_workflow() -> WorkflowDefinition:
    definition = WorkflowDefinition(
        workflow_id="dag:组合工作表",
        name="组合工作表（按配置）",
        description="按“组合工作表分组方案”将文件组合为工作簿。",
        settings={"flow_name": "组合工作表"},
    )
    definition.nodes = [
        WorkflowNode("combine-sheets", "excel.combine_sheets",
                     display_name="组合工作表",
                     parameters={"flow_name": "组合工作表"}),
    ]
    definition.input_bindings = [WorkflowBinding("runtime", "combine-sheets", "runtime")]
    definition.output_bindings = [
        WorkflowBinding("result", "combine-sheets", "result"),
        WorkflowBinding("output_files", "combine-sheets", "output_files"),
    ]
    return definition


def default_workflows() -> Dict[str, WorkflowDefinition]:
    workflows = [
        build_audit_workflow(),
        build_summary_workflow(),
        build_merge_org_workflow(),
        build_combine_workflow(),
    ]
    return {item.workflow_id: item for item in workflows}
