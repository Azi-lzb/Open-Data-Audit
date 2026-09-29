"""业务功能与普通流程模板。

Module 是代码能力；FunctionDefinition 是面向实施人员的业务预置；
BusinessFlowTemplate 决定普通编辑器允许基于哪些已验收 DAG 创建流程。
调度器只执行编译后的 WorkflowDefinition，不感知本文件中的 UI 概念。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Tuple


@dataclass(frozen=True)
class FunctionDefinition:
    function_id: str
    name: str
    module_id: str
    description: str = ""
    category: str = ""
    default_parameters: Dict[str, Any] = field(default_factory=dict)
    config_schema: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    allow_disable: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return {
            "functionId": self.function_id,
            "name": self.name,
            "moduleId": self.module_id,
            "description": self.description,
            "category": self.category,
            "defaultParameters": dict(self.default_parameters),
            "configSchema": {key: dict(value) for key, value in self.config_schema.items()},
            "allowDisable": self.allow_disable,
        }


@dataclass(frozen=True)
class BusinessFlowTemplate:
    template_id: str
    name: str
    description: str
    base_workflow_id: str
    step_functions: Tuple[Tuple[str, str], ...]

    def to_dict(self, functions: Dict[str, FunctionDefinition]) -> Dict[str, Any]:
        return {
            "templateId": self.template_id,
            "name": self.name,
            "description": self.description,
            "baseWorkflowId": self.base_workflow_id,
            "steps": [
                {
                    "nodeId": node_id,
                    "functionId": function_id,
                    "name": functions[function_id].name,
                    "description": functions[function_id].description,
                    "configSchema": functions[function_id].config_schema,
                    "allowDisable": functions[function_id].allow_disable,
                }
                for node_id, function_id in self.step_functions
            ],
        }


def _text(title: str, description: str = "") -> Dict[str, Any]:
    return {"type": "text", "title": title, "description": description}


def build_function_catalog() -> Dict[str, FunctionDefinition]:
    items = (
        FunctionDefinition("source.files", "报送文件来源", "source.excel_files", "读取本次勾选的报送文件。", "来源"),
        FunctionDefinition("source.template", "模板装载与体检", "source.excel_template", "只读装载模板并完成模板体检。", "来源"),
        FunctionDefinition("check.validation_ranges", "检查校验区域", "excel.named_range_check", "检查模板所需校验区域。", "检查"),
        FunctionDefinition("check.structure_ranges", "检查表结构区域", "excel.named_range_check", "检查模板所需表结构区域。", "检查"),
        FunctionDefinition("check.structure", "表结构比对", "excel.structure_compare", "过滤表结构不匹配的报送文件。", "检查"),
        FunctionDefinition("modify.audit_copy", "准备审核副本", "audit.copy_files", "生成只写审核副本，保留原文件。", "修改"),
        FunctionDefinition("modify.external_sheets", "添加外部工作表", "excel.external_sheet_copy", "把公式需要的外部工作表加入审核副本；未选择外部文件时自动跳过。", "修改"),
        FunctionDefinition("modify.formula_recalculate", "复制公式并重算", "excel.formula_copy_recalculate", "复制模板公式和格式，并调用当前办公套件重算。", "修改"),
        FunctionDefinition("extract.formula_issues", "提取公式问题", "excel.formula_issue_extract", "从重算后的审核副本提取公式问题。", "提取"),
        FunctionDefinition("extract.conditional_issues", "提取条件格式问题", "excel.conditional_format_issue_extract", "提取实际触发的条件格式单元格。", "提取"),
        FunctionDefinition("aggregate.audit_results", "汇聚审核问题", "audit.result_merge", "合并公式和条件格式问题。", "汇总"),
        FunctionDefinition("aggregate.history", "带入历史说明", "audit.history_enrich", "匹配历史状态和人工说明。", "汇总"),
        FunctionDefinition(
            "output.audit_result", "生成审核结果", "audit.result_workbook_write",
            "生成最终审核结果工作簿。", "输出",
            config_schema={"output_name": _text("输出名称", "不填写时沿用流程默认名称。")},
        ),
        FunctionDefinition("output.run_log", "生成运行日志", "log.workbook_write", "生成逐功能运行日志。", "输出"),
        FunctionDefinition("summary.region", "区域汇总", "summary.region_summary", "按模板区域汇总多个文件。", "汇总",
            config_schema={"output_name": _text("输出名称")},
        ),
        FunctionDefinition("summary.history", "历史说明富化", "summary.history_enrich", "按复合键把历史表的人工说明列追加到汇总输出。", "汇总"),
        FunctionDefinition(
            "merge.sheets", "组合工作表", "excel.combine_sheets", "按正则或关键字方案组合工作表。", "合并",
            config_schema={"plan_id": {"type": "combine_plan", "title": "分组方案"}},
        ),
    )
    return {item.function_id: item for item in items}


def build_business_flow_templates() -> Dict[str, BusinessFlowTemplate]:
    items = (
        BusinessFlowTemplate(
            "audit", "汇总核查表校验", "完整审核主流程；关键顺序保持已验收配置。", "dag:汇总核查表校验",
            (
                ("source-files", "source.files"), ("template", "source.template"),
                ("check-validation-ranges", "check.validation_ranges"),
                ("check-structure-ranges", "check.structure_ranges"),
                ("compare-structure", "check.structure"), ("copy-files", "modify.audit_copy"),
                ("add-external-sheets", "modify.external_sheets"),
                ("copy-formulas", "modify.formula_recalculate"),
                ("extract-formula-issues", "extract.formula_issues"),
                ("extract-conditional-issues", "extract.conditional_issues"),
                ("merge-audit-results", "aggregate.audit_results"),
                ("history-enrich", "aggregate.history"),
                ("write-result-workbook", "output.audit_result"),
                ("log-workbook-write", "output.run_log"),
            ),
        ),
        BusinessFlowTemplate(
            "summary", "汇总校验结果说明", "结构检查后执行区域汇总，再富化历史说明。", "dag:汇总校验结果说明",
            (
                ("source-files", "source.files"), ("template", "source.template"),
                ("check-structure-range", "check.structure_ranges"),
                ("compare-structure", "check.structure"),
                ("region-summary", "summary.region"),
                ("history-enrich", "summary.history"),
            ),
        ),
        BusinessFlowTemplate(
            "combine_sheets", "组合工作表", "使用指定分组方案组合工作表。", "dag:组合工作表",
            (("combine-sheets", "merge.sheets"),),
        ),
    )
    return {item.template_id: item for item in items}
