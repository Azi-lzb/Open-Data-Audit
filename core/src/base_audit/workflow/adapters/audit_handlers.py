"""DAG handlers：把现有审核能力包装为显式节点（阶段 3）。

编排语义与 service.AuditService.run() 的阶段块逐行对齐——同一批底层原语
（ExcelSession / preflight_xlsx / history / 临时助手），仅把“编排层、数据
传递层”换成显式 DAG：文件集走 FileSetArtifact，问题集走
AuditResultSetArtifact。禁止复制 Excel 业务实现。
"""

from __future__ import annotations

import shutil
import time
from dataclasses import replace
from pathlib import Path
from typing import Any, Dict, List

from openpyxl import load_workbook

from ..artifacts import (
    AuditResultSetArtifact,
    FileRecord,
    FileSetArtifact,
    LogEventArtifact,
    MetadataArtifact,
    OperationResultArtifact,
    ResultWorkbookArtifact,
    StructureMatchArtifact,
)
from ..context import NodeContext
from ..modules import ModuleCategory, ModuleDefinition
from ..ports import InputPortDefinition, OutputPortDefinition
from ..scheduler import ModuleExecutionResult
from ..types import (
    ArtifactSubtype,
    ArtifactType,
    ExecutionScope,
    FailurePolicy,
    Retention,
    OLD_FAILURE_POLICY_MAP,
)
from ...external import make_external_sheet_plan
from ...external_sheet_writer import (
    EXTERNAL_SHEET_WRITER_DIRECT_OOXML,
    EXTERNAL_SHEET_WRITERS,
    DirectOoxmlExternalSheetWriter,
)
from ...models import FileAuditResult, PreflightItem, SourceMatch
from ...name_config import (
    USED_RANGE_SUMMARY_FUNCTION,
    FeatureMapping,
)
from ...native.openpyxl_workbook import (
    add_external_sheets,
    extract_issues as extract_saved_issues,
    require_named_ranges,
    template_structure_values,
)
from ...preflight_xlsx import validate_source_xlsx
from ...template_snapshot import TemplateSnapshot, load_template_snapshot
from ...service import (
    _available_output_path,
    _organisation_from_name,
    _output_prefix,
    _sha256,
)


def _records(artifact) -> List[FileRecord]:
    return list(artifact.records)


def _item_failure_policy(parameters: Dict[str, Any]) -> FailurePolicy:
    """兼容旧配置中文值；节点枚举值则直接接受。"""
    raw = str(parameters.get("on_failure") or "跳过")
    if raw in OLD_FAILURE_POLICY_MAP:
        return OLD_FAILURE_POLICY_MAP[raw]
    return FailurePolicy(raw)


def _copy_version(record: FileRecord, parameters: Dict[str, Any], stage: str) -> FileRecord:
    """创建真正独立的下一版本文件，禁止修改上游 Artifact 指向的文件。"""
    folder = Path(parameters["stage_dir"]) / stage
    folder.mkdir(parents=True, exist_ok=True)
    target = _available_output_path(folder, Path(record.source))
    shutil.copy2(record.audit, target)
    return FileRecord(
        source=record.source, audit=str(target), org_code=record.org_code,
        org_name=record.org_name, version=record.version + 1,
        stage_paths=record.stage_paths + (str(target),), stage_errors=record.stage_errors,
    )


def _copy_final_formula_version(record: FileRecord, parameters: Dict[str, Any]) -> FileRecord:
    """创建持久的最终公式副本；审核副本不再附加“审核导航”工作表。"""
    folder = Path(parameters["audit_output_dir"])
    folder.mkdir(parents=True, exist_ok=True)
    target = _available_output_path(folder, Path(record.source))
    shutil.copy2(record.audit, target)
    return FileRecord(
        source=record.source, audit=str(target), org_code=record.org_code,
        org_name=record.org_name, version=record.version + 1,
        stage_paths=record.stage_paths + (str(target),), stage_errors=record.stage_errors,
    )


# ---------------------------------------------------------------------------
# 节点 Handler
# ---------------------------------------------------------------------------

def handle_source_files(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """source.excel_files：源目录扫描 + 待处理清单过滤（与旧 _source_files 一致）。"""
    from ...service import _source_files

    sources = _source_files(
        Path(parameters["input_dir"]),
        [Path(p) for p in parameters["selected_files"]] if parameters.get("selected_files") is not None else None,
        recursive=parameters.get("recursive", False),
        extra_files=[Path(p) for p in parameters["extra_files"]] if parameters.get("extra_files") is not None else None,
    )
    if not sources:
        raise ValueError("源数据目录中没有可审核的 .xlsx 文件")
    records = []
    for path in sources:
        org_code, org_name = _organisation_from_name(path)
        records.append(FileRecord(
            source=str(path), org_code=org_code, org_name=org_name,
            stage_paths=(str(path),),
        ))
    artifact = FileSetArtifact(records=tuple(records), subtype=ArtifactSubtype.EXCEL_SOURCE_SET)
    ctx.log(f"待处理文件 {len(records)} 个")
    return ModuleExecutionResult.ok(source_files=artifact)


def handle_template_load(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """source.excel_template：只读打开模板一次，读取模板定义与体检结果。

    复刻 run() 模板阶段的语义：读取/体检/外部表计划失败 → 体检报告（按需）
    + TemplateError；体检存在“错误”级条目 → 拒绝执行。

    ``summary_template=True``（汇总流程使用）时只按汇总区域建立定义，不要求
    “校验区域”或启用的校验规则——汇总模板本就不含可复制的校验公式。
    """
    if parameters.get("summary_template"):
        return _load_summary_template(ctx, parameters)

    from ...template import TemplateError
    from ...node_flow_config import load_dag_feature_mappings

    template_path = Path(parameters["template_path"])
    output_dir = Path(parameters["output_dir"])
    batch_id = parameters["batch_id"]
    external_path = parameters.get("external_path")
    feature_log = ctx.run.resources.get("feature_log")
    preflight_path = output_dir / f"审核前检查_{batch_id}.xlsx"
    keep_preflight_report = parameters.get("keep_preflight_report", True)
    if feature_log is not None:
        keep_preflight_report = False

    external_plan = None
    try:
        try:
            mappings = load_dag_feature_mappings()
            snapshot = load_template_snapshot(template_path, mappings)
            definition = snapshot.definition
            template_items = _inspect_template_health_snapshot(snapshot)
            if external_path and parameters.get("include_external", True):
                external_book = load_workbook(Path(external_path), read_only=True, data_only=False, keep_links=False)
                try:
                    external_plan = make_external_sheet_plan(
                        formulas=snapshot.formulas,
                        available_sheets=external_book.sheetnames,
                    )
                    template_items.append(
                        PreflightItem(
                            "外部文件", "提示", "通过", "", "、".join(external_plan.sheet_names),
                            f"将复制 {len(external_plan.sheet_names)} 个工作表（{external_plan.source}）",
                        )
                    )
                finally:
                    external_book.close()
            if feature_log is not None:
                feature_log.add_template_health(template_path, template_items)
        except Exception as exc:
            template_items = [
                PreflightItem("总体", "错误", "不通过", "", "", str(exc)),
            ]
            if feature_log is not None:
                feature_log.add_template_health(template_path, template_items)
            if keep_preflight_report:
                from ...preflight_xlsx import write_preflight_report_xlsx

                write_preflight_report_xlsx(preflight_path, template_path, template_items)
            raise TemplateError(
                f"模板体检不通过：{exc}"
                + (f"。体检报告：{preflight_path}" if keep_preflight_report else "")
            ) from exc

        if any(item.level == "错误" for item in template_items):
            if keep_preflight_report:
                from ...preflight_xlsx import write_preflight_report_xlsx

                write_preflight_report_xlsx(preflight_path, template_path, template_items)
            raise TemplateError("模板体检不通过，请先处理错误项")

        metadata = MetadataArtifact(payload={
            "definition": definition,
            "external_plan": external_plan,
            "template_items": template_items,
            "mappings": mappings,
            "snapshot": snapshot,
            "template_path": str(template_path),
            "preflight_path": str(preflight_path),
        })
        return ModuleExecutionResult.ok(template=metadata)
    except TemplateError:
        raise
    except Exception as exc:
        raise TemplateError(f"模板体检不通过：{exc}") from exc


def _inspect_template_health_snapshot(snapshot: TemplateSnapshot) -> list[PreflightItem]:
    """仅检查模板协议与静态公式文本，不为体检启动 Excel/WPS。"""
    definition = snapshot.definition
    enabled = [rule for rule in definition.rules if rule.enabled]
    result_rules = [rule for rule in enabled if rule.result_mode != "error_only"]
    items = [
        PreflightItem(
            "规则", "错误" if not enabled else "提示", "需处理" if not enabled else "信息", "", "",
            "模板没有启用的校验规则" if not enabled else (
                f"启用公式 {len(enabled)} 个，其中问题规则 {len(result_rules)} 个、"
                f"公式健康检查 {len(enabled) - len(result_rules)} 个"
            ),
        ),
        PreflightItem(
            "规则", "错误" if not result_rules else "提示", "需处理" if not result_rules else "信息", "", "",
            "没有识别到会返回问题标记的公式，审核结果可能始终为零"
            if not result_rules else "问题规则识别正常",
        ),
        PreflightItem(
            "表结构区域", "提示" if definition.structure_ranges else "错误",
            "通过" if definition.structure_ranges else "需处理", "", "",
            f"已设置 {len(definition.structure_ranges)} 个“表结构区域”命名区域"
            if definition.structure_ranges else "模板没有“表结构区域”命名区域，无法精确核对报送文件结构",
        ),
    ]
    broken = [
        f"{rule.sheet_name}!{rule.formula_cell}"
        for rule in enabled
        if "#REF!" in str(snapshot.formula_cells.get((rule.sheet_name, rule.formula_cell), "")).upper()
    ]
    if broken:
        items.append(PreflightItem(
            "公式", "警告", "关注", "", "、".join(broken[:10]),
            f"发现 {len(broken)} 个公式含 #REF! 引用；实际计算触发错误时会写入问题结果",
        ))
    has_error = any(item.level == "错误" for item in items)
    return [PreflightItem(
        "总体", "错误" if has_error else "提示", "不通过" if has_error else "通过", "", "",
        "模板存在必须处理的问题" if has_error else "模板结构可用于审核",
    )] + items


def _load_summary_template(ctx: NodeContext, parameters: Dict[str, Any]) -> ModuleExecutionResult:
    """汇总流程的模板装载：只按汇总区域 + 表头区域建立定义。

    与审核流程不同，汇总模板只用来定位“任意行汇总区域/固定行汇总区域”和
    “表头区域”，既不需要“校验区域”，也不含可复制的校验公式。定义中的
    ``copy_ranges`` 用汇总区域（供表结构比对核对报送文件含哪些工作表），
    ``structure_ranges`` 用表头区域（供逐格比对固定表头）。
    """
    from ...models import TemplateDefinition, PreflightItem
    from ...template import TemplateError
    from ...node_flow_config import load_dag_feature_mappings

    excel = ctx.resource("excel")
    template_workbook = ctx.resource("template_workbook")
    template_path = Path(parameters["template_path"])
    feature_log = ctx.run.resources.get("feature_log")
    summary_names = set(parameters.get("summary_feature_names") or ("任意行汇总",))

    mappings = load_dag_feature_mappings()
    summary_mappings = [
        mapping for mapping in mappings
        if mapping.name in summary_names and mapping.range_names
    ]
    if not summary_mappings:
        raise TemplateError(f"汇总功能未找到：{'、'.join(summary_names)}")

    template_items: list[PreflightItem] = []
    data_ranges = []
    for mapping in summary_mappings:
        data_ranges.extend(excel._named_ranges_for_features(template_workbook, [mapping]))
    headers = excel._named_ranges_for_features(
        template_workbook,
        [FeatureMapping("汇总表头", USED_RANGE_SUMMARY_FUNCTION, ("表头区域",), False, "", "")],
    )
    if not data_ranges:
        raise TemplateError("汇总功能未找到命名区域：" + "、".join(
            name for mapping in summary_mappings for name in mapping.range_names
        ))
    if not headers:
        raise TemplateError("汇总功能缺少“表头区域”命名区域")
    template_items.append(
        PreflightItem("汇总结构", "提示", "通过", "", "",
                     f"将按 {len(headers)} 个表头区域核对汇总文件结构")
    )

    external_plan = None
    external_path = parameters.get("external_path")
    if external_path and parameters.get("include_external", True):
        external_workbook = excel.open_workbook(Path(external_path), read_only=True)
        try:
            external_plan = _external_sheet_plan(
                excel, template_workbook, TemplateDefinition(
                    rules=[], copy_ranges=data_ranges, structured=False,
                    structure_ranges=headers,
                ), external_workbook,
            )
            template_items.append(
                excel.inspect_external_workbook(
                    external_workbook, external_plan.sheet_names,
                    plan_source=external_plan.source,
                )
            )
        finally:
            excel.close_workbook(external_workbook)

    if feature_log is not None:
        feature_log.add_template_health(template_path, template_items)
    definition = TemplateDefinition(
        rules=[], copy_ranges=data_ranges, structured=False, structure_ranges=headers,
    )
    metadata = MetadataArtifact(payload={
        "definition": definition,
        "external_plan": external_plan,
        "template_items": template_items,
        "mappings": mappings,
        "preflight_path": "",
    })
    return ModuleExecutionResult.ok(template=metadata)


def handle_named_range_checks(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """excel.named_range_check：先调查再强校验，区域缺失记日志后失败。"""
    feature_log = ctx.run.resources.get("feature_log")
    template = inputs["template"].payload
    snapshot = template.get("snapshot")
    if snapshot is not None:
        for feature_name in parameters.get("feature_names") or []:
            mapping = _mapping_of(template, feature_name)
            if mapping is None:
                raise ValueError(f"执行流程引用了“模块化功能”中不存在的功能：{feature_name}")
            ctx.log(f"正在执行：{feature_name}")
            areas = snapshot.require_feature_ranges(feature_name)
            survey = [
                (area.sheet_name, "、".join(mapping.range_names), area.address, "通过")
                for area in areas
            ]
            if feature_log is not None:
                feature_log.add_sheet(feature_name, ("工作表", "命名区域名", "覆盖区域", "结果"), survey)
            per_sheet: Dict[str, int] = {}
            for area in areas:
                per_sheet[area.sheet_name] = per_sheet.get(area.sheet_name, 0) + 1
            ctx.log(f"{feature_name}：" + "、".join(f"{sheet} {count} 处" for sheet, count in per_sheet.items()))
            ctx.log(f"完成：{feature_name}")
        return ModuleExecutionResult.ok()

    book = load_workbook(template["template_path"], read_only=True, data_only=False, keep_links=False)
    try:
        for feature_name in parameters.get("feature_names") or []:
            mapping = _mapping_of(template, feature_name)
            if mapping is None:
                raise ValueError(f"执行流程引用了“模块化功能”中不存在的功能：{feature_name}")
            ctx.log(f"正在执行：{feature_name}")
            areas = require_named_ranges(book, mapping)
            survey = [
                (area.sheet_name, "、".join(mapping.range_names), area.address, "通过")
                for area in areas
            ]
            if feature_log is not None:
                feature_log.add_sheet(feature_name, ("工作表", "命名区域名", "覆盖区域", "结果"), survey)
            per_sheet: Dict[str, int] = {}
            for area in areas:
                per_sheet[area.sheet_name] = per_sheet.get(area.sheet_name, 0) + 1
            detail = "、".join(f"{sheet} {count} 处" for sheet, count in per_sheet.items())
            ctx.log(f"{feature_name}：{detail}")
            ctx.log(f"完成：{feature_name}")
    finally:
        book.close()
    return ModuleExecutionResult.ok()


def handle_structure_compare(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """excel.structure_compare：逐文件 openpyxl 精确比对固定表头。"""
    template = inputs["template"].payload
    definition = template["definition"]
    external_plan = template.get("external_plan")
    feature_log = ctx.run.resources.get("feature_log")
    source_files = inputs["source_files"]

    snapshot = template.get("snapshot")
    expected_structure = (
        [(item, [list(row) for row in values]) for item, values in snapshot.structure_values]
        if snapshot is not None
        else template_structure_values(Path(template["template_path"]), definition)
    )
    source_matches = []
    processable = []
    rejected_results: List[FileAuditResult] = []
    for record in _records(source_files):
        source_path = Path(record.source)
        try:
            match_result = validate_source_xlsx(
                source_path, definition, expected_structure,
                external_sheet_names=external_plan.sheet_names if external_plan else (),
            )
        except Exception as exc:
            match_result = SourceMatch(False, 0.0, 0, 0, (), f"无法完成结构匹配：{exc}")
        source_matches.append((source_path, match_result))
        if match_result.matched:
            processable.append(record)
        else:
            rejected_results.append(
                FileAuditResult(source_path=source_path, audit_path=None,
                                org_code=record.org_code, org_name=record.org_name,
                                error="报送文件与模板不匹配：" + match_result.details)
            )
    matched_count = sum(1 for _, match in source_matches if match.matched)
    ctx.log(f"表结构比对：{matched_count} 个文件通过，{len(source_matches) - matched_count} 个跳过")
    if feature_log is not None:
        structure_name = parameters.get("feature_name") or "表结构比对"
        feature_log.add_sheet(
            structure_name,
            ("报送文件", "匹配结果", "匹配率", "匹配标签数", "检查标签数", "缺少工作表", "说明"),
            [
                (
                    str(source_path),
                    "通过" if result.matched else "不通过",
                    f"{result.score:.0%}",
                    result.matched_labels,
                    result.checked_labels,
                    "、".join(result.missing_sheets) or "无",
                    result.details,
                )
                for source_path, result in source_matches
            ],
        )
    accepted = FileSetArtifact(records=tuple(processable), subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET)
    structure = StructureMatchArtifact(
        matches=tuple(source_matches),
        accepted=tuple(str(path) for path, match in source_matches if match.matched),
        rejected=tuple(str(path) for path, match in source_matches if not match.matched),
        metadata={"rejected_file_results": tuple(rejected_results)},
    )
    return ModuleExecutionResult.ok(accepted_files=accepted, structure_report=structure)


def handle_files_copy(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """复制审核副本快照（每文件复制前后各一次 SHA-256 校验）。"""
    flow_name = parameters.get("flow_name", "")
    stage_dir = Path(parameters["stage_dir"]) / "v1-copy"
    stage_dir.mkdir(parents=True, exist_ok=True)
    records = []
    for record in _records(inputs["accepted_files"]):
        audit_path = _available_output_path(stage_dir, Path(record.source))
        source_hash = _sha256(Path(record.source))
        shutil.copy2(record.source, audit_path)
        if _sha256(Path(record.source)) != source_hash:
            try:
                audit_path.unlink(missing_ok=True)
            except OSError:
                pass
            raise RuntimeError(
                f"执行流程“{flow_name}”处理“{Path(record.source).name}”失败，已按“停止”中止："
                "原始报送文件在审核过程中发生变化"
            )
        records.append(FileRecord(
            source=record.source, audit=str(audit_path),
            org_code=record.org_code, org_name=record.org_name, version=record.version + 1,
            stage_paths=record.stage_paths + (str(audit_path),),
        ))
    artifact = FileSetArtifact(records=tuple(records), subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET)
    artifact.lineage = (inputs["accepted_files"].artifact_id,)
    return ModuleExecutionResult.ok(workbooks=artifact)


def handle_external_copy(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """excel.external_sheet_copy：把外部工作表复制进审核副本（先于公式复制）。"""
    template = inputs["template"].payload
    external_plan = template.get("external_plan")
    output_dir = Path(parameters["output_dir"])
    batch_id = parameters["batch_id"]
    source_files = inputs["workbooks"]

    records = [_copy_version(record, parameters, "v2-external") for record in _records(source_files)]
    if external_plan is None:
        return ModuleExecutionResult.ok(workbooks=FileSetArtifact(
            records=tuple(records), subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET,
            lineage=(source_files.artifact_id,),
        ))

    service = ctx.run.services.get("service")
    writer_mode = str(getattr(service, "external_sheet_writer", "") or "OPENPYXL").upper()
    if writer_mode not in EXTERNAL_SHEET_WRITERS:
        writer_mode = "OPENPYXL"
    direct_writer = None
    prepare_error: Exception | None = None
    prepare_seconds = 0.0
    if writer_mode == EXTERNAL_SHEET_WRITER_DIRECT_OOXML:
        try:
            direct_writer = DirectOoxmlExternalSheetWriter(
                Path(parameters["external_path"]), external_plan.sheet_names,
            )
            prepared = direct_writer.prepare()
            prepare_seconds = prepared.prepare_seconds
            for message in prepared.messages:
                ctx.log(message)
        except Exception as exc:
            prepare_error = exc
    ctx.log(f"外部文件添加方式：{writer_mode}")
    log_rows: List[tuple] = []
    write_seconds = 0.0
    for record in records:
        external_ok = True
        try:
            if prepare_error is not None:
                raise prepare_error
            if direct_writer is not None:
                written = direct_writer.write(Path(record.audit))
                write_seconds += written.write_seconds
                for message in written.messages:
                    ctx.log(f"  {message}")
            else:
                add_external_sheets(
                    Path(record.audit), Path(parameters["external_path"]), external_plan.sheet_names,
                )
            checkpoint_dir = parameters.get("checkpoint_dir_name") or ""
            if checkpoint_dir:
                checkpoint = output_dir / f"{checkpoint_dir}_{batch_id}"
                checkpoint.mkdir(parents=True, exist_ok=True)
                shutil.copy2(record.audit, checkpoint / Path(record.audit).name)
        except Exception as exc:
            if _item_failure_policy(parameters) != FailurePolicy.SKIP_ITEM:
                try:
                    Path(record.audit).unlink(missing_ok=True)
                except OSError:
                    pass
                raise RuntimeError(
                    f"执行流程“{parameters.get('flow_name', '')}”处理“{Path(record.source).name}”失败，"
                    f"已按“停止”中止：{exc}"
                ) from exc
            external_ok = False
            detail = f"{parameters.get('feature_name', '外部文件添加')}：{exc}"
            record.stage_errors = tuple(record.stage_errors) + (detail,)
            ctx.log(f"已跳过：{Path(record.source).name}｜{detail}")
        log_rows.append((
            Path(record.source).name,
            "、".join(external_plan.sheet_names) if external_plan else "",
            "成功" if external_ok else "；".join(record.stage_errors),
        ))
        ctx.log(f"为{Path(record.source).name}添加了{external_plan.sheet_names[0] if external_plan else '外部工作表'}")
    artifact = FileSetArtifact(
        records=tuple(FileRecord(
            source=r.source, audit=r.audit, org_code=r.org_code,
            org_name=r.org_name, version=r.version,
            stage_paths=r.stage_paths,
            stage_errors=r.stage_errors,
        ) for r in records),
        subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET,
        lineage=(source_files.artifact_id,),
    )
    sheet_name = parameters.get("feature_name") or "外部文件添加"
    log = LogEventArtifact(entries=((sheet_name, ("报送文件", "复制的外部工作表", "结果"), tuple(log_rows)),))
    ctx.run.resources.setdefault("log_entries", []).append(log)
    result = ModuleExecutionResult.ok(workbooks=artifact)
    if direct_writer is not None:
        result.metrics.update({
            "external_prepare_seconds": prepare_seconds,
            "external_direct_write_seconds": write_seconds,
        })
    return result


def _raw_formula_ranges(template_book: Any, definition) -> Dict[tuple[str, str], Any]:
    """一次读取 OOXML 原始公式矩阵，供 COM 在暂停计算期间写入。"""
    from openpyxl.utils.cell import range_boundaries

    result: Dict[tuple[str, str], Any] = {}
    for item in definition.copy_ranges:
        identity = (item.sheet_name, item.address.upper())
        if identity in result:
            continue
        min_col, min_row, max_col, max_row = range_boundaries(item.address)
        source_sheet = template_book[item.sheet_name]
        matrix = tuple(
            tuple(source_sheet.cell(row, column).value for column in range(min_col, max_col + 1))
            for row in range(min_row, max_row + 1)
        )
        result[identity] = matrix[0][0] if len(matrix) == 1 and len(matrix[0]) == 1 else matrix
    return result


def handle_formula_copy(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """excel.formula_copy_recalculate：COM 区域复制并逐副本计算。"""
    excel = ctx.resource("excel")
    definition = inputs["template"].payload["definition"]
    feature_log = ctx.run.resources.get("feature_log")
    source_files = inputs["workbooks"]
    feature_name = parameters.get("feature_name") or "公式校验复制"

    log_rows: List[tuple] = []
    records = [_copy_final_formula_version(record, parameters) for record in _records(source_files)]
    template_path = Path(inputs["template"].payload["template_path"])
    copy_seconds = 0.0
    recalculate_seconds = 0.0
    snapshot = inputs["template"].payload.get("snapshot")
    template_book = None
    template_workbook = None
    # 写入后端由设置中心「公式校验复制方式」决定（formula_region_writer）：
    # COM_RANGE=稳定基线；DIRECT_OOXML=ZIP 级高速写入（无回退）。计算策略
    # 两者完全一致：单实例、单副本 open→算→存→关，不做全局批量重算。
    from ...formula_region_writer import (
        WRITER_COM_RANGE, DirectOoxmlFormulaRegionWriter,
    )
    service = ctx.run.services.get("service")
    writer_mode = str(getattr(service, "formula_region_writer", "") or WRITER_COM_RANGE).upper()
    if writer_mode not in (WRITER_COM_RANGE, "DIRECT_OOXML"):
        writer_mode = WRITER_COM_RANGE
    direct_writer = DirectOoxmlFormulaRegionWriter() if writer_mode == "DIRECT_OOXML" else None
    ctx.log(f"公式校验复制方式：{writer_mode}")
    try:
        if snapshot is not None:
            formula_overrides = snapshot.formula_overrides
        else:
            template_book = load_workbook(template_path, data_only=False, keep_links=False)
            formula_overrides = _raw_formula_ranges(template_book, definition)
        if writer_mode != "DIRECT_OOXML":
            template_workbook = excel.open_workbook(
                template_path, read_only=True, purpose="template",
            )
        for record in records:
            formula_ok = True
            workbook = None
            try:
                ctx.log(f"正在复制公式并重算：{Path(record.source).name}")
                copy_started = time.monotonic()
                if direct_writer is not None:
                    # DIRECT_OOXML：不经 COM 写公式（副本保持关闭），写入后
                    # 再打开→只算当前工作簿→保存→关闭（计算策略与基线一致）。
                    write_result = direct_writer.write(
                        template_path=template_path, audit_path=Path(record.audit),
                        definition=definition, formula_overrides=formula_overrides,
                    )
                    for message in write_result.messages:
                        ctx.log(f"  {message}")
                    workbook = excel.open_workbook(Path(record.audit), read_only=False)
                    recalculate_started = time.monotonic()
                    excel.calculate_workbook(workbook)
                    workbook.Save()
                    recalculate_seconds += time.monotonic() - recalculate_started
                    copy_seconds += write_result.seconds
                else:
                    workbook = excel.open_workbook(Path(record.audit), read_only=False)
                    excel.apply_rules(
                        template_workbook, workbook, definition,
                        calculate=False, formula_overrides=formula_overrides,
                    )
                    copy_seconds += time.monotonic() - copy_started

                    recalculate_started = time.monotonic()
                    excel.calculate_workbook(workbook)
                    workbook.Save()
                    recalculate_seconds += time.monotonic() - recalculate_started
            except Exception as exc:
                if _item_failure_policy(parameters) != FailurePolicy.SKIP_ITEM:
                    try:
                        Path(record.audit).unlink(missing_ok=True)
                    except OSError:
                        pass
                    raise RuntimeError(
                        f"执行流程“{parameters.get('flow_name', '')}”处理“{Path(record.source).name}”失败，"
                        f"已按“停止”中止：{exc}"
                    ) from exc
                formula_ok = False
                detail = f"{feature_name}：{exc}"
                record.stage_errors = tuple(record.stage_errors) + (detail,)
                ctx.log(f"已跳过：{Path(record.source).name}｜{detail}")
            finally:
                if workbook is not None:
                    try:
                        excel.close_workbook(workbook)
                    except Exception:
                        pass
            log_rows.append((
                Path(record.source).name, str(record.audit), sum(1 for rule in definition.rules if rule.enabled),
                "成功" if formula_ok else "；".join(record.stage_errors),
            ))
    finally:
        if template_workbook is not None:
            try:
                excel.close_workbook(template_workbook)
            except Exception:
                pass
        if template_book is not None:
            template_book.close()

    ctx.log(f"公式写入 {copy_seconds:.2f} 秒；逐副本重算与保存 {recalculate_seconds:.2f} 秒")
    artifact = FileSetArtifact(
        records=tuple(FileRecord(
            source=r.source, audit=r.audit, org_code=r.org_code,
            org_name=r.org_name, version=r.version, stage_paths=r.stage_paths,
            stage_errors=r.stage_errors,
        ) for r in records),
        subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET,
        lineage=(source_files.artifact_id,),
    )
    log = LogEventArtifact(entries=(
        (feature_name, ("报送文件", "审核副本", "复制公式数", "结果"), tuple(log_rows)),
    ))
    ctx.run.resources.setdefault("log_entries", []).append(log)
    result = ModuleExecutionResult.ok(workbooks=artifact)
    result.metrics.update({"formula_copy_seconds": copy_seconds, "formula_recalculate_seconds": recalculate_seconds})
    return result


def handle_formula_issue_extract(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """excel.formula_issue_extract：直接读取已重算并保存的审核副本公式缓存。"""
    definition = inputs["template"].payload["definition"]
    period = parameters["period"]
    batch_id = parameters["batch_id"]
    source_files = inputs["workbooks"]

    issues = []
    for record in _records(source_files):
        stage_errors = record.stage_errors
        if stage_errors:
            # 旧口径：该文件此前的修改阶段按“停止”已中止；按“跳过”则本阶段照常。
            pass
        ctx.log(f"正在读取公式缓存：{Path(record.source).name}")
        extracted = extract_saved_issues(
            Path(record.audit), definition.rules, period=period, batch_id=batch_id,
            org_code=record.org_code, org_name=record.org_name,
            source_file=Path(record.source), extraction_ranges=definition.extraction_ranges,
        )
        for issue in extracted:
            issue.audit_time = parameters["audit_time"]
        issues.extend(extracted)
    artifact = AuditResultSetArtifact(issues=tuple(issues))
    artifact.lineage = (source_files.artifact_id,)
    return ModuleExecutionResult.ok(formula_issues=artifact)


def handle_conditional_issue_extract(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """excel.conditional_format_issue_extract：读取原始报送文件的条件格式触发项。

    检测模式由设置中心「条件格式检测模式」决定（经 AuditService 传入）：
    NATIVE（默认）= Excel/WPS COM 真实渲染对比；OOXML = 直接求值报送文件内的
    条件格式规则（不启动办公软件读取颜色）。两条路径共用同一业务输出。
    """
    excel = ctx.resource("excel")
    template = inputs["template"].payload
    definition = template["definition"]
    period = parameters["period"]
    batch_id = parameters["batch_id"]
    audit_time = parameters["audit_time"]
    mapping = _mapping_of(template, parameters.get("feature_name") or "条件格式结果提取")
    if mapping is None:
        raise ValueError("执行流程引用了“模块化功能”中不存在的功能：条件格式结果提取")
    source_files = inputs["source_files"]
    service = ctx.run.services.get("service")
    # 三求值器由设置中心决定（旧口径 OOXML/NATIVE 已在设置层映射）。
    from ...conditional_evaluators import (
        EVALUATOR_COM,
        EVALUATOR_COM_DISPLAY,
        EVALUATOR_PYTHON,
        ConditionalFormatService,
        effective_evaluator_mode,
    )

    mode = effective_evaluator_mode(
        str(getattr(service, "conditional_format_evaluator", "") or "COM_DISPLAY"))
    rule_reader = str(getattr(service, "conditional_format_rule_reader", "") or "DIRECT_OOXML")

    # 三种判定方式统一走 ConditionalFormatService：
    # PYTHON/COM_EVALUATE 求值规则真假；COM_DISPLAY 渲染判定 + RuleReader 解释。
    return _conditional_issues_via_service(
        ctx, template, definition, mapping, source_files, excel,
        mode=mode, service=service, rule_reader=rule_reader,
        period=period, batch_id=batch_id, audit_time=audit_time,
    )


def _conditional_issues_via_service(ctx: NodeContext, template: Dict[str, Any], definition,
                                   mapping, source_files, excel, *, mode: str, service,
                                   rule_reader: str = "DIRECT_OOXML",
                                   period: str, batch_id: str, audit_time: str) -> ModuleExecutionResult:
    """PYTHON / COM_EVALUATE：Direct OOXML 规则读取 + 所选求值器（无回退）。"""
    from ...conditional_evaluators import ConditionalFormatService
    from ...native.openpyxl_workbook import named_ranges

    ctx.log(f"条件格式判定方式：{mode} | 规则读取方式：{rule_reader}（规则统一来自 RuleReader）")
    book = load_workbook(
        Path(template["template_path"]), read_only=True, data_only=False, keep_links=False)
    try:
        ranges = named_ranges(book, (mapping,))
    finally:
        book.close()

    cf_service = ConditionalFormatService(mode, session=excel)
    issues = []
    stat_lines = []
    for record in _records(source_files):
        ctx.log(f"正在求值条件格式规则：{Path(record.source).name}")
        com_workbook = None
        if mode in ("COM_EVALUATE", "COM_DISPLAY"):
            com_workbook = excel.open_workbook(Path(record.source), read_only=True)
        try:
            file_issues, stats = cf_service.extract_issues(
                workbook_path=Path(record.source), ranges=ranges,
                structure_ranges=definition.structure_ranges,
                period=period, batch_id=batch_id, audit_time=audit_time,
                source_file=Path(record.source), mapping=mapping,
                com_workbook=com_workbook,
                org_code=record.org_code, org_name=record.org_name,
            )
        finally:
            if com_workbook is not None:
                excel.close_workbook(com_workbook)
        issues.extend(file_issues)
        if mode == "COM_DISPLAY":
            line = ("  rule_read {rule_read_time:.2f}s | display_read {display_format_read_time:.2f}s"
                    " | total {total_time:.2f}s | 规则匹配 {match_counts} | 触发 {triggered}"
                    ).format(**stats, triggered=len(file_issues))
        else:
            line = ("  rule_read {rule_read_time:.2f}s | evaluate {evaluation_time:.2f}s"
                    " | build {result_build_time:.2f}s | total {total_time:.2f}s"
                    " | 规则 {rules_total}（不支持 {unsupported_rule_count}）"
                    " | 触发 {triggered}").format(**stats)
            if "com_call_count" in stats:
                line += f" | COM 调用 {stats['com_call_count']} 次（{stats.get('com_evaluate_time', 0):.2f}s）"
        stat_lines.append(f"{Path(record.source).name}: {line}")
        if stats.get("unsupported_rule_details"):
            detail = "、".join(
                f"{name}×{count}" for name, count in
                sorted(stats["unsupported_rule_details"].items(), key=lambda kv: -kv[1])[:5])
            ctx.log(f"  UNSUPPORTED 规则：{detail}")
    for line in stat_lines:
        ctx.log(line)
    artifact = AuditResultSetArtifact(issues=tuple(issues))
    artifact.lineage = (source_files.artifact_id,)
    return ModuleExecutionResult.ok(conditional_issues=artifact)


def handle_results_merge(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """audit.result_merge：按文件交错合并——每文件先公式结果、后条件格式结果。"""
    formula = inputs["formula_issues"]
    conditional = inputs["conditional_issues"]
    records = _records(inputs["workbooks"])
    merged: list = []
    for record in records:
        audit_key = Path(record.audit).resolve()
        source_key = Path(record.source).resolve()
        merged.extend(
            item for item in formula.issues if Path(item.audit_file).resolve() == audit_key
        )
        merged.extend(
            replace(item, audit_file=str(record.audit))
            for item in conditional.issues if Path(item.source_file).resolve() == source_key
        )
    artifact = AuditResultSetArtifact(issues=tuple(merged))
    artifact.lineage = (formula.artifact_id, conditional.artifact_id)
    return ModuleExecutionResult.ok(merged=artifact)


def handle_history_enrich(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """audit.history_enrich：带入历史人工说明与审核意见（历史表只读）。"""
    from ...history import merge_history

    history = inputs["history"].payload.get("history") or []
    merged = list(inputs["merged"].issues)
    merge_history(history, merged)
    artifact = AuditResultSetArtifact(issues=tuple(merged))
    artifact.lineage = (inputs["merged"].artifact_id,)
    return ModuleExecutionResult.ok(enriched=artifact)


def handle_navigation_write(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """excel.audit_navigation_write：把带历史说明的问题写进审核副本导航表。"""
    excel = ctx.resource("excel")
    source_files = inputs["workbooks"]
    enriched = list(inputs["enriched"].issues)
    by_source: Dict[str, list] = {}
    for item in enriched:
        by_source.setdefault(str(Path(item.source_file).resolve()), []).append(item)
    file_results: List[FileAuditResult] = []
    final_records: List[FileRecord] = []
    final_by_input: Dict[str, Path] = {}
    audit_folder = Path(parameters["audit_output_dir"])
    audit_folder.mkdir(parents=True, exist_ok=True)
    for record in _records(source_files):
        file_issues = by_source.get(str(Path(record.source).resolve()), [])
        final_path = _available_output_path(audit_folder, Path(record.source))
        shutil.copy2(record.audit, final_path)
        final_by_input[str(Path(record.audit).resolve())] = final_path
        workbook = excel.open_workbook(final_path, read_only=False)
        try:
            excel.write_navigation(workbook, file_issues)
            workbook.Save()
        finally:
            excel.close_workbook(workbook)
        file_results.append(FileAuditResult(
            source_path=Path(record.source), audit_path=final_path,
            org_code=record.org_code, org_name=record.org_name,
            issues=list(file_issues),
            error="；".join(record.stage_errors),
        ))
        final_records.append(FileRecord(
            source=record.source, audit=str(final_path), org_code=record.org_code,
            org_name=record.org_name, version=record.version + 1,
            stage_paths=record.stage_paths + (str(final_path),), stage_errors=record.stage_errors,
        ))
    artifact = OperationResultArtifact(ok=True, message="审核导航写入完成", rows=tuple(file_results))
    artifact.lineage = (source_files.artifact_id, inputs["enriched"].artifact_id)
    final_workbooks = FileSetArtifact(
        records=tuple(final_records),
        subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET,
        lineage=(source_files.artifact_id,),
    )
    final_issues = AuditResultSetArtifact(issues=tuple(
        replace(item, audit_file=final_by_input[str(Path(item.audit_file).resolve())])
        if item.audit_file and str(Path(item.audit_file).resolve()) in final_by_input else item
        for item in enriched
    ))
    final_issues.lineage = (inputs["enriched"].artifact_id, final_workbooks.artifact_id)
    return ModuleExecutionResult.ok(
        navigated=artifact, workbooks=final_workbooks, final_issues=final_issues,
    )


def handle_result_workbook(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """audit.result_workbook_write：输出本期审核结果工作簿（12 列 + 超链接）。"""
    from ...native.result_writer import write_audit_summary

    output_dir = Path(parameters["output_dir"])
    batch_id = parameters["batch_id"]
    result_set = parameters.get("result_set") or "本期审核结果"
    path = output_dir / (
        f"{_output_prefix(int(parameters.get('order', 80)), parameters.get('feature_name', '审核结果输出'), parameters.get('output_name') or '')}_{batch_id}.xlsx"
    )
    write_audit_summary(path, list(inputs["enriched"].issues), sheet_name=result_set)
    artifact = ResultWorkbookArtifact(
        path=str(path), sheet_name=result_set,
        subtype=ArtifactSubtype.RESULT_WORKBOOK,
    )
    return ModuleExecutionResult.ok(summary_workbook=artifact)


def handle_log_write(ctx: NodeContext, parameters: Dict[str, Any], inputs: Dict[str, Any]) -> ModuleExecutionResult:
    """log.workbook_write：把各节点收集的日志分片写入运行日志工作簿。"""
    feature_log = ctx.run.resources.get("feature_log")
    enriched = inputs.get("enriched")
    if feature_log is not None:
        for log in ctx.run.resources.get("log_entries", []):
            for title, headers, rows in log.entries:
                feature_log.add_sheet(title, headers, rows)
        if enriched is not None:
            issue_name = parameters.get("issue_feature_name") or "校验结果提取"
            feature_log.add_sheet(
                issue_name,
                ("报送文件", "工作表", "定位单元格", "级别", "校验字段", "问题说明", "当前值", "对比值", "差值", "规则编号"),
                [
                    (
                        Path(item.source_file).name if item.source_file else "",
                        item.sheet_name,
                        item.target_cell,
                        item.severity,
                        item.check_field,
                        item.detail or item.message,
                        item.target_value,
                        item.comparison_value,
                        item.difference_value,
                        item.rule_id,
                    )
                    for item in enriched.issues
                ],
            )
            for _step, mapping in _conditional_mappings(inputs["template"].payload):
                feature_log.add_sheet(
                    mapping.name,
                    ("报送文件", "工作表", "定位单元格", "错误类型", "校验指标", "描述", "当前值"),
                    [
                        (
                            Path(item.source_file).name if item.source_file else "",
                            item.sheet_name, item.target_cell, item.severity,
                            item.check_field, item.detail or item.message, item.target_value,
                        )
                        for item in enriched.issues if item.rule_id == "条件格式填充"
                    ],
                )
    return ModuleExecutionResult.ok()


def _mapping_of(template: Dict[str, Any], feature_name: str):
    for mapping in template.get("mappings") or []:
        if mapping.name == feature_name:
            return mapping
    return None


def _conditional_mappings(template: Dict[str, Any]):
    return [
        (None, mapping) for mapping in (template.get("mappings") or [])
        if mapping.feature_type == "汇总_条件格式结果提取"
    ]


# ---------------------------------------------------------------------------
# 模块定义（代码注册表；与“流程配置.json”解耦）
# ---------------------------------------------------------------------------

FILE_SET_IN = InputPortDefinition("workbooks", ArtifactType.FILE_SET, accepted_subtypes=(ArtifactSubtype.EXCEL_WORKBOOK_SET,))
SOURCE_SET_IN = InputPortDefinition(
    "source_files", ArtifactType.FILE_SET, accepted_subtypes=(ArtifactSubtype.EXCEL_SOURCE_SET,),
)
SET_OUT = OutputPortDefinition("workbooks", ArtifactType.FILE_SET, subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET)


AUDIT_MODULE_DEFINITIONS = (
    ModuleDefinition(
        module_id="source.excel_files", version="1", name="报送文件来源",
        category=ModuleCategory.SOURCE, execution_scope=ExecutionScope.PER_RUN,
        output_ports=(OutputPortDefinition("source_files", ArtifactType.FILE_SET, subtype=ArtifactSubtype.EXCEL_SOURCE_SET),),
        description="扫描源数据目录，归一化待处理文件集。",
    ),
    ModuleDefinition(
        module_id="source.excel_template", version="1", name="模板装载与体检",
        category=ModuleCategory.SOURCE, execution_scope=ExecutionScope.PER_RUN,
        output_ports=(OutputPortDefinition("template", ArtifactType.METADATA),),
        description="只读打开模板一次，读取审核规则/命名区域并完成模板体检。",
        side_effects=("写模板体检运行日志",),
    ),
    ModuleDefinition(
        module_id="excel.named_range_check", version="1", name="命名区域检查",
        category=ModuleCategory.CHECK, execution_scope=ExecutionScope.PER_RUN,
        input_ports=(InputPortDefinition("template", ArtifactType.METADATA),),
        description="检查模板命名区域存在性；先调查后强校验。",
        side_effects=("写运行日志",),
    ),
    ModuleDefinition(
        module_id="excel.structure_compare", version="1", name="表结构比对",
        category=ModuleCategory.CHECK, execution_scope=ExecutionScope.PER_RUN,
        input_ports=(
            InputPortDefinition("source_files", ArtifactType.FILE_SET, accepted_subtypes=(ArtifactSubtype.EXCEL_SOURCE_SET,)),
            InputPortDefinition("template", ArtifactType.METADATA),
        ),
        output_ports=(
            OutputPortDefinition("accepted_files", ArtifactType.FILE_SET, subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET),
            OutputPortDefinition("structure_report", ArtifactType.STRUCTURE_MATCH_RESULT),
        ),
        description="逐格比对固定表头，输出通过/跳过清单。",
        side_effects=("写运行日志",),
    ),
    ModuleDefinition(
        module_id="audit.copy_files", version="1", name="审核副本复制",
        category=ModuleCategory.MODIFY, execution_scope=ExecutionScope.PER_FILE,
        input_ports=(InputPortDefinition("accepted_files", ArtifactType.FILE_SET, accepted_subtypes=(ArtifactSubtype.EXCEL_WORKBOOK_SET,)),),
        output_ports=(SET_OUT,),
        description="复制报送文件为审核副本快照，前后 SHA-256 校验。",
        current_file_policy={"promote_to_current": True},
    ),
    ModuleDefinition(
        module_id="excel.external_sheet_copy", version="1", name="外部工作表复制",
        category=ModuleCategory.MODIFY, execution_scope=ExecutionScope.PER_FILE,
        input_ports=(FILE_SET_IN, InputPortDefinition("template", ArtifactType.METADATA)),
        output_ports=(SET_OUT,),
        description="把外部辅助工作簿的工作表复制进审核副本，必须先于公式复制。",
        current_file_policy={"promote_to_current": True},
    ),
    ModuleDefinition(
        module_id="excel.formula_copy_recalculate", version="1", name="公式校验复制",
        category=ModuleCategory.MODIFY, execution_scope=ExecutionScope.PER_FILE,
        input_ports=(FILE_SET_IN, InputPortDefinition("template", ArtifactType.METADATA)),
        output_ports=(SET_OUT,),
        description="复制模板公式与格式并调用办公套件重算。",
        current_file_policy={"promote_to_current": True},
    ),
    ModuleDefinition(
        module_id="excel.formula_issue_extract", version="1", name="公式结果提取",
        category=ModuleCategory.CHECK, execution_scope=ExecutionScope.PER_FILE,
        input_ports=(FILE_SET_IN, InputPortDefinition("template", ArtifactType.METADATA)),
        output_ports=(OutputPortDefinition("formula_issues", ArtifactType.AUDIT_RESULT_SET),),
        description="读取重算后的审核副本缓存，提取公式校验问题。",
    ),
    ModuleDefinition(
        module_id="excel.conditional_format_issue_extract", version="1", name="条件格式结果提取",
        category=ModuleCategory.CHECK, execution_scope=ExecutionScope.PER_FILE,
        input_ports=(SOURCE_SET_IN, InputPortDefinition("template", ArtifactType.METADATA)),
        output_ports=(OutputPortDefinition("conditional_issues", ArtifactType.AUDIT_RESULT_SET),),
        description="读取原始报送文件中条件格式实际触发的单元格。",
    ),
    ModuleDefinition(
        module_id="audit.result_merge", version="1", name="审核结果汇聚",
        category=ModuleCategory.AGGREGATE, execution_scope=ExecutionScope.PER_RUN,
        input_ports=(
            InputPortDefinition("formula_issues", ArtifactType.AUDIT_RESULT_SET),
            InputPortDefinition("conditional_issues", ArtifactType.AUDIT_RESULT_SET),
            InputPortDefinition("workbooks", ArtifactType.FILE_SET, accepted_subtypes=(ArtifactSubtype.EXCEL_WORKBOOK_SET,)),
        ),
        output_ports=(OutputPortDefinition("merged", ArtifactType.AUDIT_RESULT_SET),),
        description="两类审核结果按文件序交错合并为同一结果集。",
    ),
    ModuleDefinition(
        module_id="audit.history_enrich", version="1", name="历史说明富化",
        category=ModuleCategory.MODIFY, execution_scope=ExecutionScope.PER_RUN,
        input_ports=(
            InputPortDefinition("merged", ArtifactType.AUDIT_RESULT_SET),
            InputPortDefinition("history", ArtifactType.METADATA),
        ),
        output_ports=(OutputPortDefinition("enriched", ArtifactType.AUDIT_RESULT_SET),),
        description="带入历史审核说明与审核意见；历史表只读。",
    ),
    ModuleDefinition(
        module_id="audit.result_workbook_write", version="1", name="审核结果输出",
        category=ModuleCategory.SINK, execution_scope=ExecutionScope.SINK,
        input_ports=(InputPortDefinition("enriched", ArtifactType.AUDIT_RESULT_SET),),
        output_ports=(OutputPortDefinition("summary_workbook", ArtifactType.RESULT_WORKBOOK, subtype=ArtifactSubtype.RESULT_WORKBOOK, retention_default=Retention.PERSISTENT),),
        description="输出本期审核结果工作簿（超链接定位到审核副本）。",
    ),
    ModuleDefinition(
        module_id="log.workbook_write", version="1", name="运行日志输出",
        category=ModuleCategory.SINK, execution_scope=ExecutionScope.SINK,
        input_ports=(
            InputPortDefinition("enriched", ArtifactType.AUDIT_RESULT_SET, required=False),
            InputPortDefinition("template", ArtifactType.METADATA, required=False),
        ),
        output_ports=(OutputPortDefinition("log_workbook", ArtifactType.LOG_EVENT_SET),),
        description="把各节点日志分片与问题清单写入运行日志工作簿。",
    ),
)

AUDIT_HANDLERS = {
    "source.excel_files": handle_source_files,
    "source.excel_template": handle_template_load,
    "excel.named_range_check": handle_named_range_checks,
    "excel.structure_compare": handle_structure_compare,
    "audit.copy_files": handle_files_copy,
    "excel.external_sheet_copy": handle_external_copy,
    "excel.formula_copy_recalculate": handle_formula_copy,
    "excel.formula_issue_extract": handle_formula_issue_extract,
    "excel.conditional_format_issue_extract": handle_conditional_issue_extract,
    "audit.result_merge": handle_results_merge,
    "audit.history_enrich": handle_history_enrich,
    "audit.result_workbook_write": handle_result_workbook,
    "log.workbook_write": handle_log_write,
}
