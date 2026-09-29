"""UOS/native 的审核 DAG 节点实现。

这里只组合 native 的细粒度工作簿原语（旧入口 ``AuditService.run_flow`` 与
``native.run_native_audit`` 已删除）。这样同一张 DAG 在 UOS 上也有真实的端口、版本
和谱系语义，旧执行器仍可独立保留给回归对比。
"""

from __future__ import annotations

import hashlib
import shutil
from dataclasses import replace
from pathlib import Path

from openpyxl import load_workbook

from ..artifacts import (AuditResultSetArtifact, FileRecord, FileSetArtifact,
                         LogEventArtifact, MetadataArtifact, OperationResultArtifact,
                         ResultWorkbookArtifact, StructureMatchArtifact)
from ..context import NodeContext
from ..scheduler import ModuleExecutionResult
from ..types import ArtifactSubtype
from ...discovery import source_workbooks
from ...engines.libreoffice import LibreOfficeCalculator
from ...external import make_external_sheet_plan
from ...external_sheet_writer import (
    EXTERNAL_SHEET_WRITER_DIRECT_OOXML,
    EXTERNAL_SHEET_WRITERS,
    DirectOoxmlExternalSheetWriter,
)
from ...models import FileAuditResult, SourceMatch
from ...name_config import (CONDITIONAL_FORMAT_EXTRACT_FUNCTION, FeatureMapping,
                            USED_RANGE_SUMMARY_FUNCTION)
from ...node_flow_config import load_dag_feature_mappings
from ...native.conditional_scan import extract_conditional_format_issues
from ...native.openpyxl_workbook import (add_external_sheets, copy_formula_ranges_from_template,
                                         extract_issues, named_ranges,
                                         require_named_ranges, template_formulas,
                                         template_structure_values)
from ...preflight_xlsx import validate_source_xlsx
from ...template import TemplateError
from ...template_snapshot import load_template_snapshot


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _organisation(path: Path) -> tuple[str, str]:
    parts = [part.strip() for part in path.stem.split("_") if part.strip()]
    return (parts[0], parts[1]) if len(parts) >= 2 else (path.stem, path.stem)


def _next_path(folder: Path, source: Path) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    candidate = folder / f"{source.stem}_审核版.xlsx"
    if not candidate.exists():
        return candidate
    return folder / f"{source.stem}_审核版_{len(list(folder.glob(source.stem + '_审核版*.xlsx'))):03d}.xlsx"


def _records(artifact):
    return list(artifact.records)


def _copy_version(record: FileRecord, parameters: dict, stage: str) -> FileRecord:
    target = _next_path(Path(parameters["stage_dir"]) / stage, Path(record.source))
    shutil.copy2(record.audit, target)
    return FileRecord(record.source, str(target), record.org_code, record.org_name,
                      record.version + 1, record.stage_paths + (str(target),), record.stage_errors)


def _copy_final_formula_version(record: FileRecord, parameters: dict) -> FileRecord:
    """将公式复制节点直接产出的 v3 持久化为最终审核副本。"""
    target = _next_path(Path(parameters["audit_output_dir"]), Path(record.source))
    shutil.copy2(record.audit, target)
    return FileRecord(record.source, str(target), record.org_code, record.org_name,
                      record.version + 1, record.stage_paths + (str(target),), record.stage_errors)


def handle_source_files(ctx: NodeContext, parameters: dict, inputs: dict) -> ModuleExecutionResult:
    folder = Path(parameters["input_dir"])
    sources = source_workbooks(folder, recursive=bool(parameters.get("recursive")))
    if parameters.get("selected_files") is not None:
        selected = {Path(value).resolve() for value in parameters["selected_files"]}
        sources = [path for path in sources if path.resolve() in selected]
    for value in parameters.get("extra_files") or ():
        path = Path(value).resolve()
        if path.is_file() and path not in sources:
            sources.append(path)
    sources.sort(key=lambda item: str(item).casefold())
    if not sources:
        raise ValueError("源数据目录中没有可审核的 .xlsx 文件")
    records = []
    for path in sources:
        code, name = _organisation(path)
        records.append(FileRecord(str(path), org_code=code, org_name=name, stage_paths=(str(path),)))
    ctx.log(f"待处理文件 {len(records)} 个")
    return ModuleExecutionResult.ok(source_files=FileSetArtifact(
        records=tuple(records), subtype=ArtifactSubtype.EXCEL_SOURCE_SET))


def handle_template_load(ctx: NodeContext, parameters: dict, inputs: dict) -> ModuleExecutionResult:
    template_path = Path(parameters["template_path"])
    mappings = load_dag_feature_mappings()
    if parameters.get("summary_template"):
        # 汇总流程：只按汇总区域 + 表头区域建立定义，不要求“校验区域”或校验规则。
        from ...models import TemplateDefinition
        from ...name_config import USED_RANGE_SUMMARY_FUNCTION

        summary_names = set(parameters.get("summary_feature_names") or ("任意行汇总",))
        book = load_workbook(template_path, read_only=False, data_only=False, keep_links=False)
        try:
            summary_mappings = [
                mapping for mapping in mappings
                if mapping.name in summary_names and mapping.range_names
            ]
            data_ranges = []
            for mapping in summary_mappings:
                data_ranges.extend(named_ranges(book, (mapping,)))
            headers = named_ranges(
                book,
                (FeatureMapping("汇总表头", USED_RANGE_SUMMARY_FUNCTION, ("表头区域",), False, "", ""),),
)
        finally:
            book.close()
        if not data_ranges:
            raise TemplateError("汇总功能未找到命名区域：" + "、".join(
                name for mapping in summary_mappings for name in mapping.range_names
            ))
        if not headers:
            raise TemplateError("汇总功能缺少“表头区域”命名区域")
        definition = TemplateDefinition(
            rules=[], copy_ranges=data_ranges, structured=False, structure_ranges=headers,
        )
        plan = None
        if parameters.get("external_path"):
            external = load_workbook(parameters["external_path"], read_only=True, data_only=False, keep_links=False)
            try:
                plan = make_external_sheet_plan(
                    formulas=template_formulas(template_path), available_sheets=external.sheetnames,
                )
            finally:
                external.close()
        return ModuleExecutionResult.ok(template=MetadataArtifact(payload={
            "definition": definition, "mappings": mappings, "external_plan": plan,
            "template_path": str(template_path),
        }))

    snapshot = load_template_snapshot(template_path, mappings)
    definition = snapshot.definition
    external_plan = None
    if parameters.get("external_path"):
        external = load_workbook(parameters["external_path"], read_only=True, data_only=False, keep_links=False)
        try:
            external_plan = make_external_sheet_plan(
                formulas=snapshot.formulas, available_sheets=external.sheetnames,
            )
        finally:
            external.close()
    return ModuleExecutionResult.ok(template=MetadataArtifact(payload={
        "definition": definition, "mappings": mappings, "external_plan": external_plan,
        "snapshot": snapshot,
        "template_path": str(template_path),
    }))


def handle_named_range_checks(ctx: NodeContext, parameters: dict, inputs: dict) -> ModuleExecutionResult:
    template = inputs["template"].payload
    requested = set(parameters.get("feature_names") or ())
    snapshot = template.get("snapshot")
    if snapshot is not None:
        for mapping in template["mappings"]:
            if mapping.name not in requested:
                continue
            areas = snapshot.require_feature_ranges(mapping.name)
            feature_log = ctx.run.resources.get("feature_log")
            if feature_log is not None:
                feature_log.add_sheet(
                    mapping.name, ("工作表", "命名区域名", "覆盖区域", "结果"),
                    [(area.sheet_name, "、".join(mapping.range_names), area.address, "通过") for area in areas],
                )
            ctx.log(f"完成：{mapping.name}")
        return ModuleExecutionResult.ok()

    book = load_workbook(template["template_path"], read_only=True, data_only=False, keep_links=False)
    try:
        for mapping in template["mappings"]:
            if mapping.name in requested:
                require_named_ranges(book, mapping)
                feature_log = ctx.run.resources.get("feature_log")
                if feature_log is not None:
                    areas = named_ranges(book, (mapping,))
                    feature_log.add_sheet(
                        mapping.name,
                        ("工作表", "命名区域名", "覆盖区域", "结果"),
                        [
                            (area.sheet_name, mapping.range_names[0], area.address, "通过")
                            for area in areas
                        ],
                    )
                ctx.log(f"完成：{mapping.name}")
    finally:
        book.close()
    return ModuleExecutionResult.ok()


def handle_structure_compare(ctx: NodeContext, parameters: dict, inputs: dict) -> ModuleExecutionResult:
    template = inputs["template"].payload
    definition = template["definition"]
    snapshot = template.get("snapshot")
    expected = (
        [(item, [list(row) for row in values]) for item, values in snapshot.structure_values]
        if snapshot is not None
        else template_structure_values(Path(template["template_path"]), definition)
    )
    external = template.get("external_plan")
    accepted, matches, rejected = [], [], []
    for record in _records(inputs["source_files"]):
        path = Path(record.source)
        try:
            match = validate_source_xlsx(path, definition, expected,
                                         external_sheet_names=external.sheet_names if external else ())
        except Exception as exc:
            match = SourceMatch(False, 0, 0, 0, (), f"无法完成结构匹配：{exc}")
        matches.append((path, match))
        if match.matched:
            accepted.append(record)
        else:
            rejected.append(FileAuditResult(path, None, record.org_code, record.org_name,
                                            error="报送文件与模板不匹配：" + match.details))
    structure = StructureMatchArtifact(matches=tuple(matches),
        accepted=tuple(str(path) for path, match in matches if match.matched),
        rejected=tuple(str(path) for path, match in matches if not match.matched),
        metadata={"rejected_file_results": tuple(rejected)})
    feature_log = ctx.run.resources.get("feature_log")
    if feature_log is not None:
        feature_log.add_sheet(
            "表结构比对",
            ("报送文件", "匹配结果", "匹配率", "匹配标签数", "检查标签数", "缺少工作表", "说明"),
            [
                (
                    str(path), "通过" if match.matched else "不通过", f"{match.score:.0%}",
                    match.matched_labels, match.checked_labels,
                    "、".join(match.missing_sheets) or "无", match.details,
                )
                for path, match in matches
            ],
        )
    return ModuleExecutionResult.ok(accepted_files=FileSetArtifact(records=tuple(accepted), subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET),
                                   structure_report=structure)


def handle_files_copy(ctx: NodeContext, parameters: dict, inputs: dict) -> ModuleExecutionResult:
    records = []
    for record in _records(inputs["accepted_files"]):
        source = Path(record.source); before = _sha256(source)
        target = _next_path(Path(parameters["stage_dir"]) / "v1-copy", source)
        shutil.copy2(source, target)
        if _sha256(source) != before:
            raise RuntimeError(f"原始报送文件在审核过程中发生变化：{source.name}")
        records.append(FileRecord(record.source, str(target), record.org_code, record.org_name,
                                  record.version + 1, record.stage_paths + (str(target),)))
    return ModuleExecutionResult.ok(workbooks=FileSetArtifact(records=tuple(records), subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET))


def handle_external_copy(ctx: NodeContext, parameters: dict, inputs: dict) -> ModuleExecutionResult:
    template = inputs["template"].payload; plan = template.get("external_plan")
    records = [_copy_version(item, parameters, "v2-external") for item in _records(inputs["workbooks"])]
    service = ctx.run.services.get("service")
    writer_mode = str(getattr(service, "external_sheet_writer", "") or "OPENPYXL").upper()
    if writer_mode not in EXTERNAL_SHEET_WRITERS:
        writer_mode = "OPENPYXL"
    if plan:
        direct_writer = None
        prepare_seconds = 0.0
        if writer_mode == EXTERNAL_SHEET_WRITER_DIRECT_OOXML:
            direct_writer = DirectOoxmlExternalSheetWriter(Path(parameters["external_path"]), plan.sheet_names)
            prepared = direct_writer.prepare()
            prepare_seconds = prepared.prepare_seconds
            for message in prepared.messages:
                ctx.log(message)
        ctx.log(f"外部文件添加方式：{writer_mode}")
        write_seconds = 0.0
        for record in records:
            if direct_writer is not None:
                written = direct_writer.write(Path(record.audit))
                write_seconds += written.write_seconds
            else:
                add_external_sheets(Path(record.audit), Path(parameters["external_path"]), plan.sheet_names)
    result = ModuleExecutionResult.ok(workbooks=FileSetArtifact(records=tuple(records), subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET,
                                                                lineage=(inputs["workbooks"].artifact_id,)))
    if plan and writer_mode == EXTERNAL_SHEET_WRITER_DIRECT_OOXML:
        result.metrics.update({"external_prepare_seconds": prepare_seconds,
                               "external_direct_write_seconds": write_seconds})
    return result


def handle_formula_copy(ctx: NodeContext, parameters: dict, inputs: dict) -> ModuleExecutionResult:
    definition = inputs["template"].payload["definition"]
    template_path = Path(inputs["template"].payload["template_path"])
    calculator = LibreOfficeCalculator()
    records = [_copy_final_formula_version(item, parameters) for item in _records(inputs["workbooks"])]
    template_book = load_workbook(template_path, data_only=False, keep_links=False)
    try:
        for record in records:
            copy_formula_ranges_from_template(template_book, Path(record.audit), definition)
            ctx.log(f"正在通过 LibreOffice Calc 计算：{Path(record.source).name}")
            calculator.recalculate(Path(record.audit))
    finally:
        template_book.close()
    return ModuleExecutionResult.ok(workbooks=FileSetArtifact(records=tuple(records), subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET,
                                                                lineage=(inputs["workbooks"].artifact_id,)))


def handle_formula_issue_extract(ctx: NodeContext, parameters: dict, inputs: dict) -> ModuleExecutionResult:
    definition = inputs["template"].payload["definition"]; issues = []
    for record in _records(inputs["workbooks"]):
        issues.extend(extract_issues(Path(record.audit), definition.rules, period=parameters["period"],
            batch_id=parameters["batch_id"], org_code=record.org_code, org_name=record.org_name,
            source_file=Path(record.source), extraction_ranges=definition.extraction_ranges))
    return ModuleExecutionResult.ok(formula_issues=AuditResultSetArtifact(issues=tuple(issues)))


def handle_conditional_issue_extract(ctx: NodeContext, parameters: dict, inputs: dict) -> ModuleExecutionResult:
    template = inputs["template"].payload
    mapping = next((m for m in template["mappings"] if m.feature_type == CONDITIONAL_FORMAT_EXTRACT_FUNCTION), None)
    if mapping is None:
        return ModuleExecutionResult.ok(conditional_issues=AuditResultSetArtifact())
    book = load_workbook(template["template_path"], read_only=True, data_only=False, keep_links=False)
    try:
        ranges = named_ranges(book, (mapping,))
    finally:
        book.close()
    # UOS 管线：仅支持 PYTHON 求值器（Direct OOXML 规则读取，统一三路架构）；
    # LibreOffice 只负责工作簿公式计算，不做条件格式显示读取（UNO 限制）。
    from ...conditional_evaluators import EVALUATOR_PYTHON, ConditionalFormatService

    ctx.log(f"条件格式求值器：{EVALUATOR_PYTHON}（UOS 仅支持规则求值；规则统一来自 Direct OOXML）")
    cf_service = ConditionalFormatService(EVALUATOR_PYTHON)
    issues = []
    for record in _records(inputs["source_files"]):
        file_issues, stats = cf_service.extract_issues(
            workbook_path=Path(record.source), ranges=ranges,
            structure_ranges=template["definition"].structure_ranges,
            period=parameters["period"], batch_id=parameters["batch_id"],
            audit_time=parameters.get("audit_time") or "", source_file=Path(record.source),
        )
        ctx.log(
            f"  {Path(record.source).name}: rule_read {stats['rule_read_time']:.2f}s | "
            f"evaluate {stats['evaluation_time']:.2f}s | total {stats['total_time']:.2f}s | "
            f"规则 {stats['rules_total']}（不支持 {stats['unsupported_rule_count']}）| 触发 {stats['triggered']}"
        )
        issues.extend(file_issues)
    return ModuleExecutionResult.ok(conditional_issues=AuditResultSetArtifact(issues=tuple(issues)))


def handle_navigation_write(ctx: NodeContext, parameters: dict, inputs: dict) -> ModuleExecutionResult:
    """将导航表写入 v4 文件并重算，确保最终交付副本含公式缓存。"""
    enriched = list(inputs["enriched"].issues); by_source = {}
    for issue in enriched:
        by_source.setdefault(str(Path(issue.source_file).resolve()), []).append(issue)
    final_records, file_results, redirects_by_source = [], [], {}
    # openpyxl 写入导航表会丢弃公式缓存。必须对最终 v4 再重算一次；不能把
    # 仅公式文本正确、data_only 缓存为空的中间版本交付给人工复核。
    calculator = LibreOfficeCalculator()
    for record in _records(inputs["workbooks"]):
        # v4 是用户可复核的持久版本；v1-v3 仍只保留在本次运行的临时谱系中。
        target = _next_path(Path(parameters["audit_output_dir"]), Path(record.source))
        shutil.copy2(record.audit, target)
        final = FileRecord(record.source, str(target), record.org_code, record.org_name,
                           record.version + 1, record.stage_paths + (str(target),), record.stage_errors)
        book = load_workbook(final.audit)
        try:
            if "审核导航" in book.sheetnames: del book["审核导航"]
            sheet = book.create_sheet("审核导航", 0)
            headers = ("状态", "级别", "规则编号", "报表", "问题位置", "校验字段", "当前值", "对比值", "参考值", "差值", "详细说明", "连续期数", "首次出现期", "上次出现期", "问题说明", "公式结果", "机构反馈", "审核意见")
            sheet.append(headers)
            rows = by_source.get(str(Path(record.source).resolve()), [])
            for item in rows:
                sheet.append((item.status, item.severity, item.rule_id, item.report_code, f"{item.sheet_name}!{item.target_cell}", item.check_field, item.target_value, item.comparison_value, item.reference_value, item.difference_value, item.detail, item.consecutive_count, item.first_seen_period, item.previous_seen_period, item.message, item.formula_result, item.institution_feedback, item.auditor_opinion))
            if not rows: sheet.append(("本期未发现问题",) + ("",) * (len(headers) - 1))
            sheet.freeze_panes = "A2"; sheet.auto_filter.ref = sheet.dimensions
            book.save(final.audit)
        finally:
            book.close()
        ctx.log(f"正在通过 LibreOffice Calc 保存最终审核副本：{Path(record.source).name}")
        calculator.recalculate(Path(final.audit))
        redirects_by_source[str(Path(record.source).resolve())] = final.audit
        final_records.append(final); file_results.append(FileAuditResult(Path(record.source), Path(final.audit), record.org_code, record.org_name, list(rows), "；".join(record.stage_errors)))
    # 条件格式问题本身来自源工作簿；汇总超链接仍应统一落到可复核的最终 v4 副本。
    final_issues = AuditResultSetArtifact(issues=tuple(
        replace(item, audit_file=redirects_by_source.get(str(Path(item.source_file).resolve()), item.audit_file))
        for item in enriched
    ))
    return ModuleExecutionResult.ok(navigated=OperationResultArtifact(ok=True, rows=tuple(file_results)),
        workbooks=FileSetArtifact(records=tuple(final_records), subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET), final_issues=final_issues)


def handle_history_enrich(ctx: NodeContext, parameters: dict, inputs: dict) -> ModuleExecutionResult:
    """合并人工字段并在导航/结果输出之前完成本期历史分类。"""
    from ...history import classify_current_issues, merge_history

    history = inputs["history"].payload.get("history") or []
    issues = list(inputs["merged"].issues)
    merge_history(history, issues)
    current, resolved = classify_current_issues(parameters["period"], issues, history,
                                                  batch_id=parameters["batch_id"])
    artifact = AuditResultSetArtifact(issues=tuple(current), metadata={"resolved_issues": tuple(resolved)})
    artifact.lineage = (inputs["merged"].artifact_id,)
    return ModuleExecutionResult.ok(enriched=artifact)


def handle_result_workbook(ctx: NodeContext, parameters: dict, inputs: dict) -> ModuleExecutionResult:
    from ...native.result_writer import write_audit_summary
    path = Path(parameters["output_dir"]) / f"{parameters.get('output_name') or '汇总核查表校验'}_{parameters['batch_id']}.xlsx"
    write_audit_summary(path, list(inputs["enriched"].issues), sheet_name="本期审核结果")
    return ModuleExecutionResult.ok(summary_workbook=ResultWorkbookArtifact(path=str(path), sheet_name="本期审核结果", subtype=ArtifactSubtype.RESULT_WORKBOOK))


# 平台无关的汇聚和日志只处理 Artifact，不需要 COM。
from .audit_handlers import handle_log_write, handle_results_merge

AUDIT_NATIVE_HANDLERS = {
    "source.excel_files": handle_source_files, "source.excel_template": handle_template_load,
    "excel.named_range_check": handle_named_range_checks, "excel.structure_compare": handle_structure_compare,
    "audit.copy_files": handle_files_copy, "excel.external_sheet_copy": handle_external_copy,
    "excel.formula_copy_recalculate": handle_formula_copy, "excel.formula_issue_extract": handle_formula_issue_extract,
    "excel.conditional_format_issue_extract": handle_conditional_issue_extract,
    "audit.result_merge": handle_results_merge, "audit.history_enrich": handle_history_enrich,
    "audit.result_workbook_write": handle_result_workbook,
    "log.workbook_write": handle_log_write,
}
