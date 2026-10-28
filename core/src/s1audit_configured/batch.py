"""Batch configured audits: resolve identity per file and consolidate one result workbook."""
from __future__ import annotations

import uuid
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from time import perf_counter

from .engine import AuditEngine, _result_dict
from .conditional_audit import extract_conditional_triggers
from .formula_error_scan import scan_validation_formula_errors
from .institutions import InstitutionCatalog, resolve_institution
from .model import AuditRunResult, Configuration, RuleResult, RuntimeContext, SourceBinding
from .outputs import _write_result_workbook
from .source_reader import file_hashing_enabled, shared_readers, reader_diagnostics
from .runtime import RuntimePaths, AuditCancelled, check_cancel


@shared_readers
def inspect_institutions(paths: list[Path], catalog: InstitutionCatalog, *,
                         identity_sheet: str = "", identity_cell: str = "", on_progress=None, cancel_event=None) -> list[dict]:
    records = []
    for index, path in enumerate(paths, 1):
        check_cancel(cancel_event)
        if on_progress:
            on_progress({"phase": "identity", "file_index": index, "files_total": len(paths), "source": str(path), "message": "识别机构"})
        record = {"source": str(path.resolve()), "status": "IDENTIFIED"}
        try:
            identity = resolve_institution(path, catalog, identity_sheet=identity_sheet, identity_cell=identity_cell)
            record["identity"] = identity.to_dict()
        except Exception as exc:
            record.update(status="IDENTITY_FAILED", error_code=getattr(exc, "code", "SOURCE_READ_FAILED"), message=str(exc))
        records.append(record)
    return records


@shared_readers
def run_batch(configuration: Configuration, files: list[Path], *, business: str,
              period: str, output_root: Path, catalog: InstitutionCatalog,
              previous_files: list[Path] | None = None, previous_period: str = "",
              external: Path | None = None, external_period: str = "",
              current_id: str = "CUR", external_id: str = "EXT",
              identity_sheet: str = "", identity_cell: str = "",
              runtime_paths: RuntimePaths | None = None, on_progress=None, cancel_event=None,
              allow_external_time_series: bool = False) -> dict:
    total_started = perf_counter()
    timings: dict[str, float] = {}
    check_cancel(cancel_event)
    runtime_paths = runtime_paths or RuntimePaths()
    if not files:
        raise ValueError("请至少选择一个本期源文件")
    if previous_files and not previous_period:
        raise ValueError("选择上期文件时必须填写上期报告期")
    output_root = output_root.resolve()
    output_root.mkdir(parents=True, exist_ok=True)
    batch_id = uuid.uuid4().hex
    verify_input_integrity = file_hashing_enabled()
    watched = dict(catalog.hashes) if verify_input_integrity else {}
    if verify_input_integrity and runtime_paths.template_path:
        from .sources import file_sha256
        watched[str(runtime_paths.template_path)] = file_sha256(runtime_paths.template_path)
    if verify_input_integrity and configuration.path:
        watched[str(configuration.path)] = configuration.sha256
    if verify_input_integrity and external:
        from .sources import file_sha256
        watched.setdefault(str(external), file_sha256(external))
    files = list(dict.fromkeys(path.resolve() for path in files))
    previous_files = list(dict.fromkeys(Path(path).resolve() for path in (previous_files or [])))
    if verify_input_integrity:
        from .sources import file_sha256
        for path in [*files, *previous_files]:
            if path.is_file():
                watched[str(path)] = file_sha256(path)
    stage_started = perf_counter()
    records = inspect_institutions(files, catalog, identity_sheet=identity_sheet, identity_cell=identity_cell, on_progress=on_progress, cancel_event=cancel_event)
    previous_records = inspect_institutions(previous_files, catalog, identity_sheet=identity_sheet, identity_cell=identity_cell, on_progress=on_progress, cancel_event=cancel_event)
    timings["file_recognition_seconds"] = perf_counter() - stage_started
    _report_timing(on_progress, "文件识别", timings["file_recognition_seconds"], perf_counter() - total_started)
    previous_by_code = defaultdict(list)
    for record in previous_records:
        if record["status"] == "IDENTIFIED":
            previous_by_code[record["identity"]["code"].casefold()].append(record)
    engine = AuditEngine(configuration)
    runs: list[AuditRunResult] = []
    errors = {"RETRIEVAL_FAILED", "STRUCTURE_ERROR", "CALCULATION_FAILED", "CONFIG_ERROR"}
    timings.update(rule_audit_seconds=0.0, formula_scan_seconds=0.0, conditional_format_scan_seconds=0.0)
    for index, record in enumerate(records, 1):
        check_cancel(cancel_event)
        if on_progress:
            on_progress({"phase": "audit", "file_index": index, "files_total": len(records), "source": record["source"], "message": "审核机构文件"})
        if record["status"] != "IDENTIFIED":
            continue
        identity = record["identity"]
        code, name = identity["code"], identity["name"]
        try:
            matched_previous = previous_by_code.get(code.casefold(), [])
            if len(matched_previous) > 1:
                record.update(status="FAILED", error_code="PREVIOUS_SOURCE_AMBIGUOUS", message="同一机构匹配到多个上期文件，请只保留本业务的一份上期文件")
                continue
            previous = {}
            if matched_previous:
                previous_path = matched_previous[0]["source"]
                record["previous"] = previous_path
                previous[current_id] = SourceBinding(current_id, Path(previous_path), business, previous_period, code, name)
            context = RuntimeContext(
                period, code, name,
                {current_id: SourceBinding(current_id, Path(record["source"]), business, period, code, name)},
                previous_sources=previous,
                external_sources={external_id: SourceBinding(external_id, external, business, external_period or period, code, name)} if external else {},
                default_external_source=external_id if external else "",
                allow_external_time_series=allow_external_time_series,
            )
            stage_started = perf_counter()
            run = engine.run(context, businesses=[business], persist_schema_snapshots=False,
                             runtime_paths=runtime_paths, cancel_event=cancel_event)
            timings["rule_audit_seconds"] += perf_counter() - stage_started
            layout_paths = [runtime_paths.template_path] if runtime_paths.template_path else []
            stage_started = perf_counter()
            try:
                formula_scan = scan_validation_formula_errors(Path(record["source"]), layout_paths)
                run.ledger["template_formula_scan"] = {
                    "status": formula_scan["status"],
                    "reason": formula_scan.get("reason", ""),
                    "template_paths": [str(path) for path in layout_paths],
                    "formula_cells_scanned": formula_scan["formula_cells_scanned"],
                    "missing_sheets": formula_scan["missing_sheets"],
                    "formula_errors": formula_scan["errors"],
                }
                for error in formula_scan["errors"]:
                    origin = error.get("origin", "source_cached_value")
                    if origin == "template_formula":
                        error_message = (f"模板校验公式本身含 Excel 错误值或引用 {error['value']}；"
                                         f"原公式：{error['formula']}")
                    elif origin == "template_formula_and_source_cached_value":
                        error_message = (f"模板校验公式及源文件结果均包含 Excel 错误值 {error['value']}；"
                                         f"原公式：{error['formula']}")
                    else:
                        error_message = (f"模板校验公式在源文件中返回 Excel 错误值 {error['value']}；"
                                         f"原公式：{error['formula']}")
                    run.results.append(RuleResult(
                        rule_id=f"TEMPLATE_FORMULA_ERROR:{error['sheet']}!{error['cell']}",
                        revision=1, business=business, period=period,
                        institution_code=code, institution_name=name,
                        status="CALCULATION_FAILED", triggered=None, severity="错误",
                        check_type="校验公式错误", metric="模板校验公式结果",
                        current_value=error["value"], error_code="TEMPLATE_FORMULA_ERROR",
                        error_message=error_message,
                        run_id=run.run_id,
                        rule_name=f"模板校验公式错误 {error['sheet']}!{error['cell']}",
                        error_evidence={
                            "sheet": error["sheet"], "cell": error["cell"],
                            "source_path": error["source"], "excel_error": error["value"],
                            "template_formula": error["formula"], "origin": origin,
                            "template_paths": error.get("template_paths", []),
                            "formula_error_tokens": error.get("formula_error_tokens", []),
                        },
                        sheet_name=error["sheet"],
                    ))
                for sheet_name in formula_scan["missing_sheets"]:
                    run.results.append(RuleResult(
                        rule_id=f"TEMPLATE_FORMULA_SCAN:{sheet_name}", revision=1,
                        business=business, period=period, institution_code=code,
                        institution_name=name, status="CALCULATION_FAILED", triggered=None,
                        severity="错误", check_type="校验公式扫描失败", metric="模板校验公式",
                        error_code="TEMPLATE_VALIDATION_SHEET_MISSING",
                        error_message=f"本次源文件缺少模板校验公式所在工作表：{sheet_name}",
                        run_id=run.run_id, rule_name="模板校验公式扫描失败",
                        error_evidence={"sheet": sheet_name, "cell": ""}, sheet_name=sheet_name,
                    ))
            except Exception as exc:
                run.ledger["template_formula_scan"] = {
                    "status": "FAILED", "template_paths": [str(path) for path in layout_paths],
                    "error": f"{type(exc).__name__}: {exc}",
                }
                run.results.append(RuleResult(
                    rule_id="TEMPLATE_FORMULA_SCAN:FAILED", revision=1,
                    business=business, period=period, institution_code=code,
                    institution_name=name, status="CALCULATION_FAILED", triggered=None,
                    severity="错误", check_type="校验公式扫描失败", metric="模板校验公式",
                    error_code="TEMPLATE_FORMULA_SCAN_FAILED",
                    error_message=f"模板校验公式错误扫描未完成：{type(exc).__name__}: {exc}",
                    run_id=run.run_id, rule_name="模板校验公式扫描失败",
                    error_evidence={"sheet": "", "cell": ""},
                ))
            finally:
                timings["formula_scan_seconds"] += perf_counter() - stage_started
            run.ledger["rules"] = [_result_dict(item) for item in run.results]
            run.ledger["status_counts"] = dict(Counter(item.status for item in run.results))
            stage_started = perf_counter()
            try:
                # Source rules and values are authoritative; the explicit
                # layout template only bounds the conditional-format scan.
                cf_template = layout_paths[0] if layout_paths else Path(record["source"])
                run.conditional_format_results.extend(extract_conditional_triggers(
                    Path(record["source"]), institution_name=name,
                    report_item=identity.get("report_item", ""), institution_code=code,
                    template_path=cf_template,
                ))
            except Exception as exc:
                run.conditional_format_results.append({
                    "机构名称": name, "报表项目": identity.get("report_item", ""),
                    "统一社会信用代码": code, "工作表名": "", "定位单元格": "",
                    "单元格值": None, "判定状态": "无法判定", "条件格式类型": "读取失败", "规则公式": "",
                    "应用范围": "", "优先级": None, "来源文件": Path(record["source"]).name,
                    "判定说明": str(exc),
                })
            finally:
                timings["conditional_format_scan_seconds"] += perf_counter() - stage_started
            check_cancel(cancel_event)
            run.ledger["conditional_format_results"] = run.conditional_format_results
            runs.append(run)
            counts = run.ledger["status_counts"]
            record.update(status="COMPLETED" if run.status == "COMPLETED" and not errors.intersection(counts) else "COMPLETED_WITH_ISSUES", run_id=run.run_id,
                          counts=counts)
            if previous_files and not matched_previous:
                record["message"] = "未匹配到同机构上期文件；涉及 V_PRE 的规则会报告取数失败"
        except AuditCancelled:
            raise
        except Exception as exc:
            record.update(status="FAILED", error_code=getattr(exc, "code", "AUDIT_FAILED"), message=str(exc))
    for key, label in (
        ("rule_audit_seconds", "规则审核"),
        ("formula_scan_seconds", "模板公式扫描"),
        ("conditional_format_scan_seconds", "条件格式扫描"),
    ):
        _report_timing(on_progress, label, timings[key], perf_counter() - total_started)
    check_cancel(cancel_event)
    if verify_input_integrity:
        from .sources import file_sha256
        hashes = {path: {"before": before, "after": file_sha256(Path(path)) if Path(path).is_file() else "MISSING"}
                  for path, before in watched.items()}
        unchanged = all(value["before"] == value["after"] for value in hashes.values())
    else:
        hashes = {}
        unchanged = True
    if not unchanged:
        for run in runs:
            run.status = "INPUT_CHANGED_DURING_RUN"
            for result in run.results:
                if result.status in {"PASSED", "TRIGGERED"} or result.error_code == "TEMPLATE_FORMULA_ERROR":
                    result.status = "CALCULATION_FAILED"
                    result.triggered = None
                    result.error_code = "INPUT_CHANGED_DURING_RUN"
                    result.error_message = "批量审核期间输入或机构参照发生变化，请重新运行"
            run.ledger.update(run_status=run.status, rules=[_result_dict(item) for item in run.results],
                              status_counts=dict(Counter(item.status for item in run.results)),
                              batch_input_integrity="CHANGED_DURING_RUN")
        for record in records:
            if record.get("run_id"):
                record.update(status="FAILED", error_code="INPUT_CHANGED_DURING_RUN", message="批量审核期间输入或机构参照发生变化，请重新运行")
                record["counts"] = next(run.ledger["status_counts"] for run in runs if run.run_id == record["run_id"])
        for run in runs:
            run.conditional_format_results.clear()
    all_results = [result for run in runs for result in run.results]
    counts = dict(Counter(result.status for result in all_results))
    processed = sum(record["status"].startswith("COMPLETED") for record in records)
    issue = any(record["status"] != "COMPLETED" for record in records) or any(record["status"] != "IDENTIFIED" for record in previous_records)
    status = "INPUT_CHANGED_DURING_RUN" if not unchanged else "FAILED" if not processed else "COMPLETED_WITH_ISSUES" if issue else "COMPLETED"
    result_path = _result_path(output_root)
    diagnostics = reader_diagnostics()
    reader_warnings = [f"{Path(item['path']).name}：已回退 openpyxl（{item['fallback_reason']}）"
                       for item in diagnostics if item["fallback_reason"]]
    data_quality_warnings = []
    warning_keys = set()
    for run in runs:
        for warning in run.ledger.get("data_quality_warnings", []):
            key = tuple(sorted((str(name), str(value)) for name, value in warning.items()))
            if key not in warning_keys:
                warning_keys.add(key)
                data_quality_warnings.append(warning)
    ledger = {
        "batch_id": batch_id, "run_status": status, "period": period,
        "files_total": len(records), "files_processed": processed,
        "files_failed": len(records) - processed, "status_counts": counts,
        "institutions": records, "previous_institutions": previous_records,
        "input_integrity": ("UNCHANGED" if unchanged else "CHANGED_DURING_RUN") if verify_input_integrity else "NOT_CHECKED",
        "input_hashes": {path: {"sha256_before": value["before"], "sha256_after": value["after"]} for path, value in hashes.items()},
        "structure_checks": [event for run in runs for event in run.ledger.get("structure_checks", [])],
        "source_reader_diagnostics": diagnostics, "reader_warnings": reader_warnings,
        "data_quality_warnings": data_quality_warnings,
    }
    conditional_results = [row for run in runs for row in run.conditional_format_results]
    source_paths = [*files, *previous_files, external, runtime_paths.template_path]
    source_map = {str(Path(path).resolve()): Path(path).resolve() for path in source_paths if path}
    combined = AuditRunResult(batch_id, period, configuration.version,
                              configuration.sha256 if verify_input_integrity else "", status,
                              all_results, ledger, conditional_format_results=conditional_results)
    check_cancel(cancel_event)
    if on_progress:
        on_progress({"phase": "export", "file_index": len(records), "files_total": len(records), "source": "", "message": "写出汇总审核结果"})
    stage_started = perf_counter()
    _write_result_workbook(combined, result_path, source_map, batch_records=records, previous_records=previous_records)
    timings["result_export_seconds"] = perf_counter() - stage_started
    timings["batch_seconds"] = perf_counter() - total_started
    ledger["stage_timings"] = {key: round(value, 3) for key, value in timings.items()}
    _report_timing(on_progress, "结果导出", timings["result_export_seconds"], timings["batch_seconds"])
    _report_timing(on_progress, "审核与结果生成合计", timings["batch_seconds"], timings["batch_seconds"])
    return {"ok": status == "COMPLETED", "status": status, "run_id": batch_id,
            "result_path": str(result_path), "output_dir": str(output_root), "counts": counts,
            "conditional_format_count": len(conditional_results),
            "conditional_triggered": sum(row.get("判定状态") == "已触发" for row in conditional_results),
            "conditional_unresolved": sum(row.get("判定状态") == "无法判定" for row in conditional_results),
            "files_total": len(records), "files_processed": processed, "files_failed": len(records) - processed,
            "institutions": records, "previous_institutions": previous_records,
            "source_reader_diagnostics": diagnostics, "reader_warnings": reader_warnings,
            "data_quality_warnings": data_quality_warnings,
            "input_integrity": ledger["input_integrity"], "stage_timings": ledger["stage_timings"]}


def _result_path(output_root: Path) -> Path:
    stem = f"汇总核查表校验_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    candidate = output_root / f"{stem}.xlsx"
    suffix = 2
    while candidate.exists():
        candidate = output_root / f"{stem}_{suffix:02d}.xlsx"
        suffix += 1
    return candidate


def _report_timing(on_progress, label: str, seconds: float, total_seconds: float) -> None:
    if on_progress:
        on_progress({"phase": "timing", "stage": label, "elapsed_seconds": round(seconds, 3),
                     "total_elapsed_seconds": round(total_seconds, 3),
                     "message": f"{label}耗时 {seconds:.2f} 秒"})
