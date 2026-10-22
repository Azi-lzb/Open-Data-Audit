from __future__ import annotations

import re
import uuid
from collections import defaultdict
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Mapping

from .dsl import evaluate, parse_dictionary, parse_expression, validate_dictionary_dependencies, variables_in
from .errors import ConfigurationError, EvaluationError, RetrievalError
from .model import (
    AuditRunResult, Configuration, EvalValue, Evidence, RuleDefinition, RuleResult, RuntimeContext,
)
from .sources import SourceRepository, file_sha256
from .workbook_config import load_configuration
from .source_reader import shared_readers
from .runtime import RuntimePaths, check_cancel


_STRUCTURE_ERROR_CODES = {
    "SHEET_NOT_FOUND", "STRUCTURE_ANCHOR_NOT_FOUND", "ROW_PATH_NOT_FOUND",
    "COLUMN_PATH_NOT_FOUND", "STRUCTURE_ERROR", "METRIC_AMBIGUOUS",
    "MAPPING_SHEET_MISSING", "MAPPING_FIELD_NOT_FOUND", "MAPPING_FIELD_AMBIGUOUS",
    "METRIC_NOT_FOUND", "REPORT_ITEM_NOT_FOUND", "REPORT_ITEM_DUPLICATE",
    "NUMERIC_REGION_NOT_FOUND", "NUMERIC_REGION_AMBIGUOUS",
}
_CONFIG_ERROR_CODES = {
    "SOURCE_CONTRACT_MISSING", "CATALOG_METRIC_NOT_FOUND", "METRIC_NOT_CONFIRMED",
    "UNIT_CONTRACT_MISSING", "PERIOD_CONTRACT_MISSING", "INSTITUTION_CONTRACT_MISSING",
    "SOURCE_AMBIGUOUS",
}


class AuditEngine:
    """Independent XLSX-configured audit engine; does not depend on the legacy S1F1 runtime."""

    def __init__(self, configuration: Configuration | str | Path):
        self.configuration = load_configuration(configuration) if isinstance(configuration, (str, Path)) else configuration

    @shared_readers
    def run(self, context: RuntimeContext, *, output_dir: str | Path | None = None,
            run_id: str | None = None, businesses: tuple[str, ...] | list[str] | None = None,
            schema_snapshot_dir: str | Path | None = None, persist_schema_snapshots: bool = True,
            runtime_paths: RuntimePaths | None = None, cancel_event=None) -> AuditRunResult:
        if not context.period:
            raise ConfigurationError("运行上下文必须提供报告期")
        run_id = run_id or uuid.uuid4().hex
        started_at = datetime.now(timezone.utc)
        repository = SourceRepository(self.configuration, context, schema_snapshot_dir=schema_snapshot_dir, runtime_paths=runtime_paths)
        repository.snapshot_selected_hashes()
        chosen = self._select_rules(context.period, businesses=businesses)
        if not _within_span(context.period, self.configuration.effective_from, self.configuration.effective_to):
            chosen = [(rule, ("NOT_APPLICABLE", "本次报告期不在配置版本生效范围内"))
                      for rule, _ in chosen]
        results: list[RuleResult] = []
        try:
            for rule, selection_status in chosen:
                check_cancel(cancel_event)
                if selection_status:
                    results.append(self._status_result(rule, context, run_id, selection_status[0], selection_status[1]))
                else:
                    scope_paths = self._scope_paths(rule)
                    if not scope_paths:
                        results.append(self._status_result(
                            rule, context, run_id, "CONFIG_ERROR", "适用范围没有匹配到任何模板行路径",
                        ))
                    else:
                        for scope_path in scope_paths:
                            check_cancel(cancel_event)
                            results.append(self._evaluate_rule(rule, context, repository, run_id, scope_path=scope_path))
            unchanged = repository.hashes_unchanged()
            configuration_after = ""
            configuration_unchanged = True
            if self.configuration.path:
                configuration_after = file_sha256(self.configuration.path) if self.configuration.path.is_file() else "MISSING"
                configuration_unchanged = configuration_after == self.configuration.sha256
                unchanged = unchanged and configuration_unchanged
            hashes = {
                key: {"sha256_before": start, "sha256_after": repository.end_hashes.get(key, "")}
                for key, start in repository.start_hashes.items()
            }
            if self.configuration.path:
                hashes["configuration"] = {
                    "sha256_before": self.configuration.sha256,
                    "sha256_after": configuration_after or "MISSING",
                }
            if not unchanged:
                for result in results:
                    if result.status in {"PASSED", "TRIGGERED"}:
                        result.status = "CALCULATION_FAILED"
                        result.triggered = None
                        result.error_code = "INPUT_CHANGED_DURING_RUN"
                        result.error_message = "审核运行期间输入文件哈希发生变化，结论不可用"
            structure_events = repository.compare_structure_snapshots()
            if not unchanged and output_dir is None:
                for event in structure_events:
                    if event["status"] in {"BASELINE_CREATED", "UNCHANGED", "POSITION_CHANGED"}:
                        event["status"] = "SNAPSHOT_NOT_SAVED"
                        event["detail"] = "输入或配置在运行期间发生变化，未更新结构快照"
            run_status = "COMPLETED" if unchanged else "INPUT_CHANGED_DURING_RUN"
            finished_at = datetime.now(timezone.utc)
            ledger = self._make_ledger(run_id, context, results, hashes, unchanged, run_status,
                                       configuration_unchanged, started_at, finished_at)
            ledger["structure_checks"] = structure_events
            ledger["data_quality_warnings"] = list(repository.data_quality_warnings)
            output_files: dict[str, str] = {}
            result = AuditRunResult(
                run_id=run_id, period=context.period, configuration_version=self.configuration.version,
                configuration_sha256=self.configuration.sha256, status=run_status,
                results=results, ledger=ledger, output_files=output_files,
            )
            check_cancel(cancel_event)
            if output_dir is not None:
                from .outputs import write_run_outputs
                output_files = write_run_outputs(result, self.configuration, repository, Path(output_dir))
                result.output_files = output_files
                result.ledger["output_files"] = output_files
            elif unchanged and persist_schema_snapshots:
                result.ledger["structure_checks"] = repository.persist_structure_snapshots()
            return result
        finally:
            repository.close()

    def _select_rules(self, period: str, *, businesses: tuple[str, ...] | list[str] | None = None
                      ) -> list[tuple[RuleDefinition, tuple[str, str] | None]]:
        date_value = _parse_date(period, field="报告期")
        groups: dict[str, list[RuleDefinition]] = defaultdict(list)
        for rule in self.configuration.rules:
            if businesses and rule.business not in set(businesses):
                continue
            groups[rule.rule_id].append(rule)
        chosen: list[tuple[RuleDefinition, tuple[str, str] | None]] = []
        for rule_id in sorted(groups):
            rules = groups[rule_id]
            eligible = [rule for rule in rules if _within_rule_period(rule, date_value)]
            if not eligible:
                rule = max(rules, key=lambda item: item.revision)
                chosen.append((rule, ("NOT_APPLICABLE", "本次报告期不在规则生效范围内")))
                continue
            selected = max(eligible, key=lambda item: item.revision)
            if not selected.enabled:
                chosen.append((selected, ("NOT_EXECUTED", "所选规则修订未启用")))
            else:
                chosen.append((selected, None))
        return chosen

    def _scope_paths(self, rule: RuleDefinition) -> tuple[tuple[str, ...], ...]:
        if not rule.applicability_scope:
            return ((),)
        scope = next((item for item in self.configuration.applicability_scopes
                      if item.scope_id == rule.applicability_scope), None)
        if scope is None:
            return ()
        paths = {
            metric.row_path for metric in self.configuration.metrics
            if metric.business == scope.business and metric.sheet == scope.sheet
            and metric.row_path[:len(scope.row_path_prefix)] == scope.row_path_prefix
        }
        return tuple(sorted(paths))

    def _evaluate_rule(self, rule: RuleDefinition, context: RuntimeContext,
                       repository: SourceRepository, run_id: str,
                       scope_path: tuple[str, ...] = ()) -> RuleResult:
        base = self._new_result(rule, context, run_id)
        base.scope_path = " / ".join(scope_path)
        cache: dict[str, EvalValue] = {}
        variable_stack: list[str] = []
        traced_calls: set[tuple[Any, ...]] = set()
        try:
            dictionary_nodes = parse_dictionary(rule.value_dictionary)
            validate_dictionary_dependencies(dictionary_nodes)
            retrieval_cache: dict[tuple[Any, ...], EvalValue | Exception] = {}
            used_variables: set[str] = set()

            def resolve_variable(name: str) -> EvalValue:
                if name in cache:
                    used_variables.add(name)
                    return cache[name]
                if name not in dictionary_nodes:
                    raise EvaluationError(f"UNDEFINED_VARIABLE：未定义变量 {name}")
                used_variables.add(name)
                variable_stack.append(name)
                try:
                    value = evaluate(dictionary_nodes[name], resolve_variable=resolve_variable,
                                     call_value=call_value, numeric_mode=rule.numeric_mode)
                finally:
                    variable_stack.pop()
                cache[name] = value
                return value

            def call_value(name: str, args: tuple[Any, ...], keywords: Mapping[str, str]) -> EvalValue:
                if name in {"V_CUR", "V_PRE"} and len(args) == 2 and args[0] == "@当前行":
                    if not scope_path:
                        raise ConfigurationError("取值函数使用 @当前行，但当前规则没有展开适用范围")
                    args = (scope_path, *args[1:])
                elif name in {"V_CUR", "V_PRE"} and len(args) == 3 and args[1] == "@当前行":
                    if not scope_path:
                        raise ConfigurationError("取值函数使用 @当前行，但当前规则没有展开适用范围")
                    args = (args[0], scope_path, args[2])
                key = (name, args, tuple(sorted(keywords.items())))
                if key in retrieval_cache:
                    existing = retrieval_cache[key]
                    if isinstance(existing, Exception):
                        raise existing
                    return existing
                try:
                    value = repository.retrieve(name, args, dict(keywords), rule)
                except Exception as exc:
                    retrieval_cache[key] = exc
                    if key not in traced_calls:
                        traced_calls.add(key)
                        base.retrieval_trace.extend(_retrieval_trace(
                            name, args, keywords, variable_stack,
                            source_id=_trace_source_id(name, keywords, rule, context), error=str(exc),
                        ))
                    raise
                retrieval_cache[key] = value
                if key not in traced_calls:
                    traced_calls.add(key)
                    base.retrieval_trace.extend(_retrieval_trace(
                        name, args, keywords, variable_stack,
                        source_id=_trace_source_id(name, keywords, rule, context), value=value,
                    ))
                return value

            if rule.applicable_expression:
                applicable = self._run_expression(rule.applicable_expression, resolve_variable, call_value,
                                                   numeric_mode=rule.numeric_mode)
                if applicable.kind != "boolean" or not isinstance(applicable.value, bool):
                    raise ConfigurationError("适用表达式必须返回布尔值")
                if not applicable.value:
                    base.status = "NOT_APPLICABLE"
                    base.error_message = "规则适用表达式判定为不适用"
                    base.evidence = list(applicable.evidence)
                    base.unused_variables = sorted(set(dictionary_nodes) - used_variables)
                    return base

            audit_value = self._run_expression(rule.audit_expression, resolve_variable, call_value,
                                               numeric_mode=rule.numeric_mode)
            if audit_value.kind != "boolean" or not isinstance(audit_value.value, bool):
                raise EvaluationError("TYPE_ERROR：审核表达式必须返回布尔值")
            audit_passes = not audit_value.value if rule.audit_expression_inverted else audit_value.value
            current_value = self._run_expression(rule.current_expression, resolve_variable, call_value,
                                                 numeric_mode=rule.numeric_mode) if rule.current_expression else None
            comparison_value = self._run_expression(rule.comparison_expression, resolve_variable, call_value,
                                                    numeric_mode=rule.numeric_mode) if rule.comparison_expression else None
            difference_value = self._run_expression(rule.difference_expression, resolve_variable, call_value,
                                                    numeric_mode=rule.numeric_mode) if rule.difference_expression else None
            if difference_value is not None and difference_value.kind not in {"number", "blank"}:
                raise EvaluationError("TYPE_ERROR：差值表达式必须返回数值或空值")
            description = _render_description(rule.description, resolve_variable)

            base.status = "TRIGGERED" if not audit_passes else "PASSED"
            base.triggered = not audit_passes
            base.description = description
            base.current_value = current_value.value if current_value else None
            base.comparison_value = comparison_value.value if comparison_value else None
            base.difference_value = difference_value.value if difference_value else None
            base.current_unit = current_value.unit or "" if current_value else ""
            base.comparison_unit = comparison_value.unit or "" if comparison_value else ""
            base.difference_unit = difference_value.unit or "" if difference_value else ""
            base.current_precision = current_value.precision if current_value else None
            base.comparison_precision = comparison_value.precision if comparison_value else None
            base.difference_precision = difference_value.precision if difference_value else None
            base.variables = {name: value.value for name, value in cache.items()}
            base.unused_variables = sorted(set(dictionary_nodes) - used_variables)
            base.evidence = _unique_evidence(
                list(audit_value.evidence)
                + (list(current_value.evidence) if current_value else [])
                + (list(comparison_value.evidence) if comparison_value else [])
                + (list(difference_value.evidence) if difference_value else [])
                + [evidence for value in cache.values() for evidence in value.evidence]
            )
            return base
        except RetrievalError as exc:
            base.status = (
                "CONFIG_ERROR" if exc.code in _CONFIG_ERROR_CODES
                else "STRUCTURE_ERROR" if exc.code in _STRUCTURE_ERROR_CODES
                else "RETRIEVAL_FAILED"
            )
            base.error_code = exc.code
            base.error_message = str(exc)
            base.error_evidence = dict(exc.evidence)
            base.evidence = _unique_evidence([evidence for value in cache.values() for evidence in value.evidence])
            base.variables = {name: value.value for name, value in cache.items()}
            return base
        except ConfigurationError as exc:
            base.status = "CONFIG_ERROR"
            base.error_code = "CONFIG_ERROR"
            base.error_message = str(exc)
            base.evidence = _unique_evidence([evidence for value in cache.values() for evidence in value.evidence])
            base.variables = {name: value.value for name, value in cache.items()}
            return base
        except EvaluationError as exc:
            base.status = "CALCULATION_FAILED"
            base.error_code = _error_code(str(exc), "CALCULATION_FAILED")
            base.error_message = str(exc)
            base.evidence = _unique_evidence([evidence for value in cache.values() for evidence in value.evidence])
            base.variables = {name: value.value for name, value in cache.items()}
            return base
        except Exception as exc:
            base.status = "CALCULATION_FAILED"
            base.error_code = "UNEXPECTED_EVALUATION_ERROR"
            base.error_message = f"规则执行出现未预期错误：{type(exc).__name__}: {exc}"
            base.evidence = _unique_evidence([evidence for value in cache.values() for evidence in value.evidence])
            base.variables = {name: value.value for name, value in cache.items()}
            return base

    def _run_expression(self, text: str, resolve_variable, call_value, *, numeric_mode="decimal") -> EvalValue:
        if not text or not text.strip():
            raise ConfigurationError("需要执行的表达式为空")
        node = parse_expression(text)
        return evaluate(node, resolve_variable=resolve_variable, call_value=call_value,
                        numeric_mode=numeric_mode)

    def _new_result(self, rule: RuleDefinition, context: RuntimeContext, run_id: str) -> RuleResult:
        return RuleResult(
            rule_id=rule.rule_id, revision=rule.revision, business=rule.business,
            period=context.period, institution_code=context.institution_code,
            institution_name=context.institution_name, status="NOT_EXECUTED", triggered=None,
            severity=rule.severity, check_type=rule.check_type, metric=rule.metric,
            description=rule.description, configuration_version=self.configuration.version,
            run_id=run_id, rule_name=rule.name, sheet_name=rule.default_sheet,
            formula_cell=rule.formula_cell,
        )

    def _status_result(self, rule: RuleDefinition, context: RuntimeContext, run_id: str,
                       status: str, message: str) -> RuleResult:
        result = self._new_result(rule, context, run_id)
        result.status = status
        result.error_message = message
        return result

    def _make_ledger(self, run_id: str, context: RuntimeContext, results: list[RuleResult],
                     hashes: dict[str, Any], unchanged: bool, run_status: str,
                     configuration_unchanged: bool, started_at: datetime,
                     finished_at: datetime) -> dict[str, Any]:
        counts: dict[str, int] = defaultdict(int)
        for result in results:
            counts[result.status] += 1
        return {
            "run_id": run_id,
            "started_at": started_at.isoformat(),
            "finished_at": finished_at.isoformat(),
            "duration_seconds": (finished_at - started_at).total_seconds(),
            "period": context.period,
            "institution_code": context.institution_code,
            "institution_name": context.institution_name,
            "configuration_version": self.configuration.version,
            "configuration_status": self.configuration.version_status,
            "configuration_release_date": self.configuration.release_date,
            "configuration_release_note": self.configuration.release_note,
            "configuration_sha256": self.configuration.sha256,
            "configuration_path": str(self.configuration.path) if self.configuration.path else "",
            "configuration_integrity": "UNCHANGED" if configuration_unchanged else "CHANGED_DURING_RUN",
            "run_status": run_status,
            "input_integrity": "UNCHANGED" if unchanged else "CHANGED_DURING_RUN",
            "input_hashes": hashes,
            "status_counts": dict(sorted(counts.items())),
            "rules": [_result_dict(result) for result in results],
        }


def _parse_date(value: Any, *, field: str) -> date:
    text = str(value or "").strip()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for pattern, fmt in ((r"^\d{8}$", "%Y%m%d"), (r"^\d{4}-\d{2}-\d{2}$", "%Y-%m-%d"),
                         (r"^\d{6}$", "%Y%m"), (r"^\d{4}-\d{2}$", "%Y-%m")):
        if re.fullmatch(pattern, text):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError as exc:
                raise ConfigurationError(f"{field} 日期格式或日期值无效：{text}") from exc
    raise ConfigurationError(f"{field} 日期格式必须为 YYYYMMDD、YYYY-MM-DD、YYYYMM 或 YYYY-MM")


def _within_rule_period(rule: RuleDefinition, period: date) -> bool:
    if rule.effective_from and period < _parse_date(rule.effective_from, field="规则生效起始日"):
        return False
    if rule.effective_to and period > _parse_date(rule.effective_to, field="规则生效截止日"):
        return False
    return True


def _within_span(period_value: str, start_value: str, end_value: str) -> bool:
    period = _parse_date(period_value, field="报告期")
    if start_value and period < _parse_date(start_value, field="配置版本生效起始日"):
        return False
    if end_value and period > _parse_date(end_value, field="配置版本生效截止日"):
        return False
    return True


def _error_code(message: str, fallback: str) -> str:
    if "：" in message:
        prefix = message.split("：", 1)[0]
        if re.fullmatch(r"[A-Z][A-Z0-9_]+", prefix):
            return prefix
    return fallback


def _render_description(template: str, resolve_variable) -> str:
    result: list[str] = []
    index = 0
    while index < len(template):
        if template.startswith("{{", index):
            result.append("{")
            index += 2
            continue
        if template.startswith("}}", index):
            result.append("}")
            index += 2
            continue
        if template[index] == "{":
            end = template.find("}", index + 1)
            if end < 0:
                raise ConfigurationError("规则描述中的变量占位符缺少结束大括号")
            name = template[index + 1:end]
            if not name or not re.fullmatch(r"[\w\u3400-\u9fff]+", name):
                raise ConfigurationError(f"规则描述中的变量名无效：{name}")
            value = resolve_variable(name)
            result.append(_display_value(value))
            index = end + 1
            continue
        if template[index] == "}":
            raise ConfigurationError("规则描述中的右大括号需要写成 }}")
        result.append(template[index])
        index += 1
    return "".join(result)


def _path_argument(value: Any) -> str:
    return "@".join(str(part) for part in value) if isinstance(value, (tuple, list)) else str(value or "")


def _trace_source_id(name, keywords, rule, context):
    requested = keywords.get("source")
    if requested:
        return requested
    if name == "V_EXT":
        return context.default_external_source or ""
    if name == "V_PRE":
        return rule.default_source or next(iter(context.previous_sources), "")
    return rule.default_source or next(iter(context.current_sources), "")


def _retrieval_trace(name, args, keywords, variable_stack, *, source_id="", value=None, error=""):
    source_sheet = str(keywords.get("sheet", ""))
    row_path = column_path = metric = ""
    if name in {"V_CUR", "V_PRE"}:
        if len(args) == 3:
            source_sheet = str(args[0])
            row_path, column_path = _path_argument(args[1]), _path_argument(args[2])
        elif len(args) == 2:
            row_path, column_path = _path_argument(args[0]), _path_argument(args[1])
    elif name == "V_EXT" and args:
        metric = str(args[0])
        column_path = metric
    evidence = value.evidence if value is not None else ()
    if not evidence:
        evidence = (None,)
    rows = []
    for item in evidence:
        rows.append({
            "variable": variable_stack[-1] if variable_stack else "审核表达式",
            "function": name, "source_id": item.source_id if item else source_id,
            "sheet": item.sheet if item else source_sheet, "requested_sheet": source_sheet,
            "row_path": row_path, "column_path": column_path, "metric": metric,
            "cell": item.cell if item else "", "raw_value": _json_value(item.raw_value) if item else "",
            "resolved_value": _json_value(value.value) if value is not None else "",
            "source_unit": item.source_unit if item else "", "standard_unit": item.standard_unit if item else "",
            "conversion": item.conversion if item else "", "report_item": item.report_item if item else "",
            "mapping_cell": item.mapping_cell if item else "", "institution_code": item.institution_code if item else "",
            "status": "已取到" if not error else "取数失败", "error": error,
        })
    return rows


def _display_value(value: EvalValue) -> str:
    if value.value is None:
        return "（空值）"
    if value.kind == "number":
        if value.numeric_mode == "excel_binary64":
            return format(value.value, ".15g").replace("e", "E")
        return format(value.value, "f")
    if value.kind == "boolean":
        return "是" if value.value else "否"
    return str(value.value)


def _unique_evidence(evidence: list[Evidence]) -> list[Evidence]:
    seen: set[tuple[str, str, str, str]] = set()
    unique: list[Evidence] = []
    for item in evidence:
        key = (item.source_id, item.sheet, item.cell, item.source_sha256)
        if key not in seen:
            seen.add(key)
            unique.append(item)
    return unique


def _json_value(value: Any) -> Any:
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if hasattr(value, "as_tuple"):
        return str(value)
    return value


def _result_dict(result: RuleResult) -> dict[str, Any]:
    return {
        "rule_id": result.rule_id, "revision": result.revision, "rule_name": result.rule_name,
        "business": result.business, "period": result.period,
        "institution_code": result.institution_code, "institution_name": result.institution_name,
        "scope_path": result.scope_path,
        "status": result.status, "triggered": result.triggered, "severity": result.severity,
        "check_type": result.check_type, "metric": result.metric, "description": result.description,
        "sheet_name": result.sheet_name,
        "formula_cell": result.formula_cell,
        "current_value": _json_value(result.current_value), "current_unit": result.current_unit,
        "current_precision": result.current_precision,
        "comparison_value": _json_value(result.comparison_value), "comparison_unit": result.comparison_unit,
        "comparison_precision": result.comparison_precision,
        "difference_value": _json_value(result.difference_value), "difference_unit": result.difference_unit,
        "difference_precision": result.difference_precision,
        "error_code": result.error_code, "error_message": result.error_message,
        "error_evidence": result.error_evidence,
        "variables": {name: _json_value(value) for name, value in result.variables.items()},
        "unused_variables": result.unused_variables,
        "configuration_version": result.configuration_version, "run_id": result.run_id,
        "evidence": [{**item.__dict__, "raw_value": _json_value(item.raw_value)} for item in result.evidence],
    }
