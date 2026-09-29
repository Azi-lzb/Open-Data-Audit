"""汇总/组合流程的 DAG Handler：直接包装 AuditService 公共方法。"""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from typing import Any, Dict

from ..artifacts import FileRecord, FileSetArtifact, MetadataArtifact, ResultWorkbookArtifact
from ..context import NodeContext
from ..modules import ModuleCategory, ModuleDefinition
from ..ports import InputPortDefinition, OutputPortDefinition
from ..scheduler import ModuleExecutionResult
from ..types import ArtifactSubtype, ArtifactType, ExecutionScope, Retention


def _records(artifact) -> list:
    return list(artifact.records)


def _params(inputs: Dict[str, Any]) -> Dict[str, Any]:
    """Runtime business parameters are an explicit workflow input Artifact."""
    runtime = inputs.get("runtime")
    if not isinstance(runtime, MetadataArtifact) or not isinstance(runtime.payload, dict):
        raise ValueError("节点缺少运行参数 Artifact")
    return runtime.payload


def _output_file_set(result: Any, runtime: MetadataArtifact) -> FileSetArtifact:
    records = []
    for item in getattr(result, "items", ()):
        output_path = getattr(item, "output_path", None)
        if output_path is None:
            continue
        sources = tuple(getattr(item, "source_files", ()) or ())
        records.append(FileRecord(
            source=str(sources[0]) if sources else str(output_path),
            audit=str(output_path),
            org_name=str(getattr(item, "organisation", "")),
            version=1,
            stage_paths=(str(output_path),),
        ))
    return FileSetArtifact(
        records=tuple(records), subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET,
        lineage=(runtime.artifact_id,), retention=Retention.PERSISTENT,
    )


def handle_summary_preflight(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """summary.source_preflight：模板体检 + 表结构比对（沿用 AuditService.preflight）。"""
    params = {**_params(inputs), **parameters}
    service = ctx.resource("service")
    from ...node_flow_config import load_dag_feature_mappings
    check = service.preflight(
        template_path=Path(params["template_path"]),
        input_dir=Path(params["input_dir"]),
        output_dir=Path(params["output_dir"]),
        selected_files=[Path(p) for p in params["selected_files"]] if params.get("selected_files") is not None else None,
        external_path=Path(params["external_path"]) if params.get("external_path") else None,
        recursive=params.get("recursive", True),
        summary_feature_names=tuple(params.get("summary_feature_names") or ()),
        write_report=False,
        on_step=ctx.run.log,
        named_range_features=tuple(params.get("named_range_features") or ()),
        feature_mappings=load_dag_feature_mappings(),
        feature_log=ctx.run.resources.get("feature_log"),
    )
    artifact = MetadataArtifact(
        payload={"check": check}, lineage=(inputs["runtime"].artifact_id,),
    )
    return ModuleExecutionResult.ok(preflight=artifact)


def handle_region_summary(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """summary.region_summary：对结构比对通过的文件执行区域汇总（含汇总表合并分派）。

    输入直接来自前驱「表结构比对」节点的通过文件集，因此本节点不再内嵌
    结构过滤；历史说明富化也已拆到独立节点，此处只产出基础汇总列。
    """
    params = {**_params(inputs), **parameters}
    service = ctx.resource("service")
    accepted = _records(inputs["accepted_files"])
    matched_files = [Path(record.source) for record in accepted]
    structure = inputs.get("structure_report")
    matches = list(getattr(structure, "matches", ()) or ())
    skipped_files = tuple((path, match) for path, match in matches if not match.matched)
    if not matched_files:
        names = "、".join(path.name for path, _ in skipped_files)
        raise ValueError(
            f"表结构比对后没有可汇总文件，已跳过 {len(skipped_files)} 个文件：{names}"
        )
    # 汇总区域由代码内置的命名区域协议决定；不再由用户编辑 DAG 区域组。
    from ...node_flow_config import load_dag_feature_mappings as _load_mappings

    summary_names = set(params.get("summary_feature_names") or ())
    summary_features = tuple(
        mapping for mapping in _load_mappings(None)
        if mapping.name in summary_names
    )
    result = service.summarize_regions(
        template_path=Path(params["template_path"]),
        input_dir=Path(params["input_dir"]),
        output_dir=Path(params["output_dir"]),
        selected_files=matched_files,
        feature_names=tuple(params.get("summary_feature_names") or ()),
        flow_name=params.get("flow_name") or "汇总校验结果说明",
        recursive=params.get("recursive", True),
        on_step=ctx.run.log,
        output_name=params.get("output_name"),
        feature_log=ctx.run.resources.get("feature_log"),
        with_history=False,
        summary_features=summary_features,
    )
    result = replace(result, skipped_files=skipped_files)
    artifact = MetadataArtifact(
        payload={"result": result},
        lineage=(inputs["accepted_files"].artifact_id, inputs["runtime"].artifact_id),
    )
    artifact.retention = Retention.PERSISTENT
    workbook = ResultWorkbookArtifact(
        path=str(result.output_path), subtype=ArtifactSubtype.RESULT_WORKBOOK,
        lineage=(artifact.artifact_id,), retention=Retention.PERSISTENT,
    )
    return ModuleExecutionResult.ok(result=artifact, summary_workbook=workbook)


def handle_summary_history_enrich(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """summary.history_enrich：把历史表的人工说明列追加到汇总工作簿（原地）。

    独立节点，与「任意行汇总」解耦：读取汇总输出，按内置追加规则的复合键
    匹配历史工作簿，把缺失的说明列补到右侧；历史表只读。
    """
    service = ctx.resource("service")
    workbook_input = inputs["summary_workbook"]
    path = Path(str(workbook_input.path))
    service.enrich_summary_history(summary_path=path, on_step=ctx.run.log)
    artifact = MetadataArtifact(
        payload={"summary_path": str(path)},
        lineage=(workbook_input.artifact_id, inputs["runtime"].artifact_id),
    )
    artifact.retention = Retention.PERSISTENT
    workbook = ResultWorkbookArtifact(
        path=str(path), subtype=ArtifactSubtype.RESULT_WORKBOOK,
        lineage=(artifact.artifact_id,), retention=Retention.PERSISTENT,
    )
    return ModuleExecutionResult.ok(summary_workbook=workbook)


def handle_combine_sheets(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    params = {**_params(inputs), **parameters}
    service = ctx.resource("service")
    common = {
        "input_dir": Path(params["input_dir"]),
        "output_dir": Path(params["output_dir"]),
        "period": params.get("period") or "",
        "selected_files": [Path(p) for p in params["selected_files"]] if params.get("selected_files") is not None else None,
        "flow_name": params.get("flow_name") or "组合工作表",
        "recursive": params.get("recursive", True),
        "on_step": ctx.run.log,
        "feature_log": ctx.run.resources.get("feature_log"),
    }
    # 每个 DAG 节点可绑定自己的分组方案；未指定时沿用“当前启用方案”。
    # 文件组合不依赖公式计算：native/UOS 仅走 openpyxl，Windows 可保留
    # Excel/WPS 兼容回退，以保障复杂工作簿的既有兼容性。
    from ...merge_org import run_combine_sheets
    from ...name_config import DEFAULT_COMBINE_SHEETS_PLAN
    plan_id = str(parameters.get("plan_id") or "")
    # “组合联合核查表”只是组合工作表的标准预置。它不读取用户当前方案，
    # 以确保入口的正则分组口径固定；自定义流程仍按保存的 plan_id 读取。
    is_union_preset = plan_id == "组合工作表（核查表）"
    if is_union_preset:
        plan = {**DEFAULT_COMBINE_SHEETS_PLAN, "name": "联合核查表默认规则", "groups": []}
        output_marker = "核查表合并"
    else:
        plan = service.get_combine_sheets_plan(plan_id or None)
        output_marker = str(plan.get("output_marker") or "组合")
    native = "native" in ctx.run.platform_capabilities
    ctx.run.log(
        f"组合分组方案：{plan['name']}（{plan['mode']}，"
        + ("openpyxl）" if native else "Excel/WPS 兼容模式）")
    )
    result = run_combine_sheets(
        **common,
        engine_preference=service.engine_preference,
        grouping_plan=plan,
        # 保留联合核查表入口的历史输出目录、日志页和文件后缀；其余逻辑
        # 完全复用组合工作表。
        operation_name="组合联合核查表" if is_union_preset else "组合工作表",
        output_file_marker=output_marker,
        allow_com_fallback=not native,
    )
    artifact = MetadataArtifact(
        payload={"result": result}, lineage=(inputs["runtime"].artifact_id,),
    )
    artifact.retention = Retention.PERSISTENT
    return ModuleExecutionResult.ok(
        result=artifact, output_files=_output_file_set(result, inputs["runtime"]),
    )


SUMMARY_MODULE_DEFINITIONS = (
    ModuleDefinition(
        module_id="summary.source_preflight", version="1", name="模板体检与结构比对",
        category=ModuleCategory.CHECK, execution_scope=ExecutionScope.PER_RUN,
        input_ports=(InputPortDefinition("runtime", ArtifactType.METADATA),),
        output_ports=(OutputPortDefinition("preflight", ArtifactType.METADATA),),
        description="沿用 AuditService.preflight：模板体检 + 表结构比对（不单独输出报告）。",
        side_effects=("写运行日志",),
    ),
    ModuleDefinition(
        module_id="summary.region_summary", version="1", name="任意行汇总",
        category=ModuleCategory.AGGREGATE, execution_scope=ExecutionScope.PER_RUN,
        input_ports=(
            InputPortDefinition("accepted_files", ArtifactType.FILE_SET,
                                accepted_subtypes=(ArtifactSubtype.EXCEL_WORKBOOK_SET,)),
            InputPortDefinition("structure_report", ArtifactType.STRUCTURE_MATCH_RESULT, required=False),
            InputPortDefinition("runtime", ArtifactType.METADATA),
        ),
        output_ports=(
            OutputPortDefinition("result", ArtifactType.METADATA, retention_default=Retention.PERSISTENT),
            OutputPortDefinition("summary_workbook", ArtifactType.RESULT_WORKBOOK, subtype=ArtifactSubtype.RESULT_WORKBOOK, retention_default=Retention.PERSISTENT),
        ),
        description="对结构比对通过的文件做任意行/固定行汇总与汇总表合并（不含历史富化）。",
        side_effects=("输出汇总工作簿", "写运行日志"),
    ),
    ModuleDefinition(
        module_id="summary.history_enrich", version="1", name="历史说明富化",
        category=ModuleCategory.MODIFY, execution_scope=ExecutionScope.PER_RUN,
        input_ports=(
            InputPortDefinition("summary_workbook", ArtifactType.RESULT_WORKBOOK,
                                accepted_subtypes=(ArtifactSubtype.RESULT_WORKBOOK,)),
            InputPortDefinition("runtime", ArtifactType.METADATA),
        ),
        output_ports=(
            OutputPortDefinition("summary_workbook", ArtifactType.RESULT_WORKBOOK,
                                 subtype=ArtifactSubtype.RESULT_WORKBOOK, retention_default=Retention.PERSISTENT),
        ),
        description="按内置追加规则的复合键，把历史表的人工说明列追加到汇总工作簿。",
        side_effects=("修改汇总工作簿", "写运行日志"),
    ),
    ModuleDefinition(
        module_id="excel.combine_sheets", version="1", name="组合工作表",
        category=ModuleCategory.MODIFY, execution_scope=ExecutionScope.PER_RUN,
        input_ports=(InputPortDefinition("runtime", ArtifactType.METADATA),),
        output_ports=(
            OutputPortDefinition("result", ArtifactType.METADATA, retention_default=Retention.PERSISTENT),
            OutputPortDefinition("output_files", ArtifactType.FILE_SET, subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET, retention_default=Retention.PERSISTENT),
        ),
        description="按分组方案组合工作表；UOS 对 .xlsx 使用 openpyxl，不回退 COM。",
    ),
)

SUMMARY_HANDLERS = {
    "summary.source_preflight": handle_summary_preflight,
    "summary.region_summary": handle_region_summary,
    "summary.history_enrich": handle_summary_history_enrich,
    "excel.combine_sheets": handle_combine_sheets,
}
