"""条件格式：三个求值/验证后端与统一服务（规则读取与求值彻底分离）。

规则来源统一为 :class:`cf_reader.OoxmlConditionalRuleReader`（Direct OOXML），
产出 :class:`cf_reader.NormalizedConditionalRule`；本模块只回答「规则是真是假」：

- ``PythonConditionalRuleEvaluator``（PYTHON）：复用共享求值器
  ``conditional_format.evaluate_cellis_rule / evaluate_expression_condition``，
  不造第二套解析器；四态 TRUE/FALSE/UNSUPPORTED/ERROR。
- ``ComFormulaEvaluator``（COM_EVALUATE）：规则仍来自 Direct OOXML；把标准化
  公式按锚点平移后交给 Excel/WPS 的 ``Worksheet.Evaluate`` 求值；cellIs 构造
  语义等价表达式；按 (规则, 平移后表达式) 去重以减少 COM 往返并计数。
- ``NativeDisplayVerifier``（DISPLAY_FORMAT_NATIVE）：保留现有 Windows 真值
  路径（``DisplayFormat.Interior.Color`` vs 基础 ``Interior.Color``）。
  **语义是 Native Render Baseline（观察最终渲染），不是布尔求值器。**

三路统一输出 :class:`ConditionalFormatResult` → 统一结果构建（conditional_context）→ Issue。
无 AUTO、无自动回退：高速/所选路径失败直接抛错。
"""

from __future__ import annotations

import re
import time
from dataclasses import dataclass, field
from pathlib import Path

from .cf_reader import (
    CellValueProvider,
    CfRuleReadError,
    CfRuleReadResult,
    NormalizedConditionalRule,
    OoxmlConditionalRuleReader,
    _col_letter,
    _col_number,
)
from .conditional_engine import (
    CONFIDENCE_FULL,
    CONFIDENCE_PARTIAL,
    _OPERATOR_LABELS,
    SOURCE_EXCEL_NATIVE,
    SOURCE_OOXML,
    SOURCE_WPS_NATIVE,
    ConditionalFormatResult,
    ConditionalFormatUnsupported,
    cellis_operand,
)
from .conditional_format import (
    UnsupportedFormulaError,
    evaluate_cellis_rule,
    evaluate_expression_condition,
    translate_formula_refs,
)
from .conditional_context import ConditionalFormatContext, build_conditional_issues
from .models import CopyRange

EVALUATOR_PYTHON = "PYTHON"
EVALUATOR_COM = "COM_EVALUATE"
EVALUATOR_COM_DISPLAY = "COM_DISPLAY"
EVALUATOR_MODES = (EVALUATOR_PYTHON, EVALUATOR_COM, EVALUATOR_COM_DISPLAY)

# 旧口径映射：OOXML（规则求值）→ PYTHON；NATIVE/DISPLAY_FORMAT_NATIVE → COM_DISPLAY。
LEGACY_EVALUATOR_MAP = {
    "OOXML": EVALUATOR_PYTHON,
    "NATIVE": EVALUATOR_COM_DISPLAY,
    "DISPLAY_FORMAT_NATIVE": EVALUATOR_COM_DISPLAY,
}

SOURCE_COM_EVALUATE = "COM_EVALUATE"


def effective_evaluator_mode(mode: str) -> str:
    """把设置口径映射为三求值器之一；未知值回退到当前默认（COM 渲染基线）。"""
    text = str(mode or "").strip().upper()
    if text in EVALUATOR_MODES:
        return text
    if text in LEGACY_EVALUATOR_MAP:
        return LEGACY_EVALUATOR_MAP[text]
    return EVALUATOR_COM_DISPLAY


# ---------------------------------------------------------------------------
# 区域限制（业务层职责）：命名“条件格式区域”限定扫描范围；空则回退 AppliesTo
# ---------------------------------------------------------------------------


def _scanned_areas(reader_result: CfRuleReadResult, ranges) -> dict[str, list[tuple[int, int, int, int]]]:
    """业务层决定哪些 Sheet/Cell 进入求值：与既有引擎相同的回退约定。"""
    scanned: dict[str, list[tuple[int, int, int, int]]] = {}
    known_sheets = set(reader_result.sheet_parts)
    for area in ranges or []:
        if area.sheet_name not in known_sheets:
            continue
        bounds = _bounds_of_address(area.address)
        if bounds:
            scanned.setdefault(area.sheet_name, []).append(bounds)
    if not scanned:
        # 回退：每张表条件格式自身的 AppliesTo（与现有引擎口径一致）
        for rule in reader_result.rules:
            for bounds in rule.bounds:
                scanned.setdefault(rule.sheet, []).append(bounds)
    return scanned


def _bounds_of_address(address: str) -> tuple[int, int, int, int] | None:
    match = re.fullmatch(
        r"\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?", str(address or "").strip(), re.I)
    if not match:
        return None

    def col(text: str) -> int:
        number = 0
        for ch in text.upper():
            number = number * 26 + (ord(ch) - 64)
        return number

    c1, r1 = col(match.group(1)), int(match.group(2))
    c2 = col(match.group(3)) if match.group(3) else c1
    r2 = int(match.group(4)) if match.group(4) else r1
    return (min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2))


def _rules_by_sheet(reader_result: CfRuleReadResult) -> dict[str, list[NormalizedConditionalRule]]:
    """Excel 语义排序：priority 数值小者优先；同优先级按文件内顺序。"""
    by_sheet: dict[str, list[NormalizedConditionalRule]] = {}
    for rule in reader_result.rules:
        by_sheet.setdefault(rule.sheet, []).append(rule)
    for rules in by_sheet.values():
        rules.sort(key=lambda rule: (rule.priority, rule.raw_metadata.get("order", 0)))
    return by_sheet


def _inside(row: int, col: int, bounds_list) -> bool:
    return any(lo_col <= col <= hi_col and lo_row <= row <= hi_row
               for lo_col, lo_row, hi_col, hi_row in bounds_list)


# ---------------------------------------------------------------------------
# RuleMatcher：COM_DISPLAY 的规则解释（判定来自渲染，解释来自 RuleReader）
# ---------------------------------------------------------------------------

MATCH_EXACT = "EXACT"
MATCH_MULTIPLE = "MULTIPLE_CANDIDATES"
MATCH_UNKNOWN = "UNKNOWN"


@dataclass
class RuleMatch:
    status: str                      # EXACT / MULTIPLE_CANDIDATES / UNKNOWN
    rules: list = field(default_factory=list)


def match_rules_for_cell(row: int, col: int,
                         rules: list[NormalizedConditionalRule]) -> RuleMatch:
    """把 DisplayFormat 判定的触发格关联到 AppliesTo 覆盖它的规则。

    DisplayFormat 只证明「显示状态变化」，不能天然证明由哪条规则导致；
    本匹配不伪造唯一结论：覆盖规则恰一条 → EXACT，多条 → MULTIPLE_CANDIDATES
    （全部候选都返回），零条 → UNKNOWN（保留 COM 原生描述）。
    """
    covering = [rule for rule in rules if _inside(row, col, rule.bounds)]
    if not covering:
        return RuleMatch(MATCH_UNKNOWN, [])
    if len(covering) == 1:
        return RuleMatch(MATCH_EXACT, covering)
    return RuleMatch(MATCH_MULTIPLE, covering)


# ---------------------------------------------------------------------------
# Evaluator A：PYTHON（复用共享求值器；值来源 = Direct OOXML 值缓存）
# ---------------------------------------------------------------------------


class _GridCell:
    __slots__ = ("value", "comment")

    def __init__(self, value, comment_text: str = "") -> None:
        self.value = value
        self.comment = None if not comment_text else type("C", (), {"text": comment_text})()


class _ValueGridSheet:
    """openpyxl 兼容的最小 ws 协议（cell(row=,column=).value），供共享求值器使用。"""

    def __init__(self, provider: CellValueProvider, sheet: str) -> None:
        self._grid = provider.values(sheet)

    def cell(self, *, row: int, column: int):
        return _GridCell(self._grid.get((column, row)))


class PythonConditionalRuleEvaluator:
    """PYTHON 求值器：规则是否成立（triggered），颜色只作附加信息。"""

    source = SOURCE_OOXML
    name = "PYTHON（OOXML 规则求值）"

    def evaluate_sheet(
        self, sheet: str, rules: list[NormalizedConditionalRule],
        areas: list[tuple[int, int, int, int]], reader_result: CfRuleReadResult,
    ) -> tuple[list[ConditionalFormatResult], dict]:
        provider = reader_result.cell_values
        grid = _ValueGridSheet(provider, sheet)
        stats = {"evaluated_cells": 0, "unsupported_cells": {}, "errors": []}
        results: list[ConditionalFormatResult] = []
        for lo_col, lo_row, hi_col, hi_row in areas:
            for row in range(lo_row, hi_row + 1):
                for col in range(lo_col, hi_col + 1):
                    stats["evaluated_cells"] += 1
                    value = provider.get(sheet, row, col)
                    partial_reason = ""
                    for rule in rules:
                        if not _inside(row, col, rule.bounds):
                            continue
                        if not rule.supported:
                            # 视觉类规则：离线无法判定，不当作未触发；降级置信度。
                            reason = "暂不支持的条件格式类型：{}".format(
                                rule.raw_metadata.get("raw_type") or rule.rule_type)
                            if not partial_reason:
                                partial_reason = reason
                            stats["unsupported_cells"][(sheet, row, col)] = reason
                            continue
                        evaluation = self._evaluate_rule(rule, value, grid, row, col)
                        if evaluation.status == "TRUE":
                            address = f"{_col_letter(col)}{row}"
                            results.append(ConditionalFormatResult(
                                sheet=sheet, cell=address,
                                range=rule.applies_to,
                                rule_type=rule.rule_type,
                                formula=rule.formula1,
                                triggered=True, source=self.source,
                                color=_fill_colour(rule),
                                message=rule_label_normalized(rule, target_cell=address),
                                confidence=CONFIDENCE_PARTIAL if partial_reason else CONFIDENCE_FULL,
                                unsupported_reason=partial_reason,
                            ))
                            break   # 最高优先级命中规则定案
                        if evaluation.status == "UNSUPPORTED" and not partial_reason:
                            partial_reason = evaluation.reason
                            stats["unsupported_cells"][(sheet, row, col)] = evaluation.reason
                        elif evaluation.status == "ERROR":
                            stats["errors"].append((sheet, row, col, rule.rule_id, evaluation.reason))
        return results, stats

    def _evaluate_rule(self, rule: NormalizedConditionalRule, value, grid, row: int, col: int):
        anchor = _anchor_parts(rule.anchor_cell)
        if rule.rule_type == "cellis":
            formulas = [rule.formula1] + ([rule.formula2] if rule.formula2 else [])
            resolver = lambda formulas, index: _resolve_operand(grid, formulas, index, anchor, row, col)
            return evaluate_cellis_rule(rule.operator, formulas, value, resolve_operand=resolver)
        if rule.rule_type == "expression":
            if not rule.formula1.strip():
                return type("R", (), {"status": "ERROR", "reason": "expression 规则公式为空"})()
            try:
                return evaluate_expression_condition(
                    rule.formula1, grid, anchor[0], anchor[1], row, col)
            except UnsupportedFormulaError as exc:
                return type("R", (), {"status": "UNSUPPORTED", "reason": str(exc)})()
            except ValueError as exc:
                return type("R", (), {"status": "ERROR", "reason": str(exc)})()
        return type("R", (), {"status": "UNSUPPORTED",
                              "reason": "暂不支持的条件格式类型：" + rule.rule_type})()


def _anchor_parts(anchor_cell: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", str(anchor_cell or ""), re.I)
    if not match:
        return (1, 1)
    row = int(match.group(2))

    def col(text: str) -> int:
        number = 0
        for ch in text.upper():
            number = number * 26 + (ord(ch) - 64)
        return number

    return (row, col(match.group(1)))


def _resolve_operand(grid, formulas: list[str], index: int, anchor, row: int, col: int):
    """cellis 操作数解析：字面量或单格引用（按锚点平移）。"""
    from .conditional_format import _expr_number

    raw = str(formulas[index] if index < len(formulas) else "").strip()
    if not raw:
        raise UnsupportedFormulaError("cellIs 操作数为空")
    number = _expr_number(raw)
    if number is not None:
        return number
    if raw.upper() in ("TRUE", "FALSE"):
        return raw.upper() == "TRUE"
    if re.fullmatch(r"\$?[A-Z]{1,3}\$?\d+", raw, re.I):
        # 单格引用：按 Excel 条件格式语义相对锚点平移。
        match = re.fullmatch(r"(\$?)([A-Z]{1,3})(\$?)(\d+)", raw, re.I)
        ref_col = _anchor_parts(match.group(2) + "1")[1]
        ref_row = int(match.group(4))
        if not match.group(1):
            ref_col += col - anchor[1]
        if not match.group(3):
            ref_row += row - anchor[0]
        return grid.cell(row=ref_row, column=ref_col).value
    raise UnsupportedFormulaError("cellIs 操作数无法解析：{}".format(raw))


def rule_label_normalized(rule: NormalizedConditionalRule, target_cell: str = "") -> str:
    """生成规则说明；有目标格时把相对引用平移到该格。

    ``formula1`` 始终保留为标准化规则的原始公式；这里只生成展示文本，
    避免 Office 渲染路径把目标格说明退回到 AppliesTo 锚点公式。
    """
    target = str(target_cell or "").strip().replace("$", "")
    if target:
        target_row, target_col = _cell_rowcol(target)
        anchor_row, anchor_col = _anchor_parts(rule.anchor_cell)
        dr, dc = target_row - anchor_row, target_col - anchor_col
    else:
        dr = dc = 0
        target = str(rule.anchor_cell or "").replace("$", "")

    if rule.rule_type == "cellis":
        op_label = _OPERATOR_LABELS.get(str(rule.operator or "").casefold(), rule.operator or "?")
        operand = translate_formula_refs(str(rule.formula1 or "").lstrip("="), dr, dc).replace("$", "")
        return "条件格式规则：{}{}{}".format(target, op_label, operand)
    if rule.rule_type == "expression":
        formula = translate_formula_refs(
            str(rule.formula1 or "").lstrip("="), dr, dc).replace("$", "")
        return "条件格式规则：{}".format(formula)
    return "条件格式规则：{}".format(rule.rule_type or "未知")


def _fill_colour(rule: NormalizedConditionalRule):
    rgb = rule.raw_metadata.get("fill_rgb")
    if not rgb:
        return None
    try:
        value = int(str(rgb)[-6:], 16)
    except (TypeError, ValueError):
        return None
    from .conditional_engine import NO_FILL_COLOURS

    return None if value in NO_FILL_COLOURS else value


# ---------------------------------------------------------------------------
# Evaluator B：COM_EVALUATE（OOXML 规则 + 办公套件公式引擎求值）
# ---------------------------------------------------------------------------

_REF_RE = re.compile(r"(?<![A-Za-z0-9_!$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![0-9A-Za-z_(])")

_CELLIS_TEMPLATE = {
    "greaterthan": "{cell}>{op}",
    "greaterthanorequal": "{cell}>={op}",
    "lessthan": "{cell}<{op}",
    "lessthanorequal": "{cell}<={op}",
    "equal": "{cell}={op}",
    "notequal": "{cell}<>{op}",
    "between": "AND({cell}>={op1},{cell}<={op2})",
    "notbetween": "OR({cell}<{op1},{cell}>{op2})",
}


def _shift_refs(formula: str, dr: int, dc: int) -> str:
    """把公式里的相对引用按 (dr, dc) 平移；$ 锁定与跨表引用保持原样。"""
    if not formula:
        return formula

    def replace(match: re.Match) -> str:
        prefix = match.group(0)
        col_locked, letters, row_locked, digits = match.groups()
        col = _col_number(letters.upper())
        row = int(digits)
        if not col_locked:
            col += dc
        if not row_locked:
            row += dr
        if col < 1 or row < 1:
            return "#REF!"
        return f"{'$' if col_locked else ''}{_col_letter(col)}{'$' if row_locked else ''}{row}"

    return _REF_RE.sub(replace, formula)


def _coerce_com_truth(value) -> bool:
    """把 Office Evaluate 的标量结果安全转换为条件真假。

    Excel/WPS 对 TRUE 可能返回 Python ``bool``，也可能返回数值 1/-1；
    COM 错误对象、数组、空值和任意其他对象不能通过 ``bool(value)`` 误报。
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value in (-1, 1)
    return False


class ComFormulaEvaluator:
    """COM_EVALUATE：把标准化公式交给 Excel/WPS 的公式引擎求值。

    COM 不重新读取条件格式规则；只对平移后的表达式做 ``Worksheet.Evaluate``。
    性能策略：同一 (规则, 平移后表达式) 只求值一次（无相对引用的规则天然去重），
    逐次计数 ``com_call_count``。
    """

    source = SOURCE_COM_EVALUATE
    name = "COM_EVALUATE（办公套件公式求值）"

    def __init__(self, session, com_workbook) -> None:
        self.session = session
        self.workbook = com_workbook
        engine_name = str(getattr(session, "engine_name", "") or "").upper()
        self.is_wps = "WPS" in engine_name

    def evaluate_sheet(
        self, sheet: str, rules: list[NormalizedConditionalRule],
        areas: list[tuple[int, int, int, int]], reader_result: CfRuleReadResult,
    ) -> tuple[list[ConditionalFormatResult], dict]:
        stats = {"evaluated_cells": 0, "com_call_count": 0, "com_evaluate_time": 0.0,
                 "unsupported_cells": {}, "errors": []}
        results: list[ConditionalFormatResult] = []
        worksheet = self.workbook.Worksheets(sheet)
        cache: dict[tuple[str, str, str, str], bool] = {}
        provider = reader_result.cell_values
        for lo_col, lo_row, hi_col, hi_row in areas:
            for row in range(lo_row, hi_row + 1):
                for col in range(lo_col, hi_col + 1):
                    stats["evaluated_cells"] += 1
                    partial_reason = ""
                    for rule in rules:
                        if not _inside(row, col, rule.bounds):
                            continue
                        if not rule.supported:
                            reason = "暂不支持的条件格式类型：{}".format(
                                rule.raw_metadata.get("raw_type") or rule.rule_type)
                            if not partial_reason:
                                partial_reason = reason
                            stats["unsupported_cells"][(sheet, row, col)] = reason
                            continue
                        expression, build_error = self._build_expression(
                            rule, row, col, provider)
                        if build_error:
                            if not partial_reason:
                                partial_reason = build_error
                            stats["unsupported_cells"][(sheet, row, col)] = build_error
                            continue
                        address = f"{_col_letter(col)}{row}"
                        cache_key = (sheet, rule.rule_id, address, expression)
                        truth = cache.get(cache_key)
                        if truth is None:
                            started = time.monotonic()
                            try:
                                # Evaluate 接受带等号的完整公式；保留原表达式用于
                                # 日志/缓存键，避免不同目标格共享规则级结果。
                                evaluate_text = expression if expression.startswith("=") else "=" + expression
                                raw = worksheet.Evaluate(evaluate_text)
                                truth = _coerce_com_truth(raw)
                            except Exception as exc:
                                stats["errors"].append((sheet, row, col, rule.rule_id, str(exc)))
                                truth = False
                            stats["com_evaluate_time"] += time.monotonic() - started
                            stats["com_call_count"] += 1
                            cache[cache_key] = truth
                        if truth:
                            results.append(ConditionalFormatResult(
                                sheet=sheet, cell=address,
                                range=rule.applies_to,
                                rule_type=rule.rule_type,
                                formula=rule.formula1,
                                triggered=True, source=self.source,
                                color=_fill_colour(rule),
                                message=rule_label_normalized(rule, target_cell=address),
                                confidence=CONFIDENCE_PARTIAL if partial_reason else CONFIDENCE_FULL,
                                unsupported_reason=partial_reason,
                            ))
                            break
        return results, stats

    def _build_expression(self, rule: NormalizedConditionalRule, row: int, col: int,
                          provider: CellValueProvider) -> tuple[str, str]:
        """构造语义等价的求值表达式；返回 (表达式, 错误)。"""
        anchor = _anchor_parts(rule.anchor_cell)
        dr, dc = row - anchor[0], col - anchor[1]
        cell_ref = f"{_col_letter(col)}{row}"
        if rule.rule_type == "expression":
            shifted = _shift_refs(rule.formula1, dr, dc)
            return shifted, ""
        if rule.rule_type == "cellis":
            template = _CELLIS_TEMPLATE.get(str(rule.operator or "").strip().casefold())
            if template is None:
                return "", "未知 cellIs 运算符：{}".format(rule.operator or "空")
            op1 = _shift_refs(rule.formula1, dr, dc)
            op2 = _shift_refs(rule.formula2, dr, dc) if rule.formula2 else ""
            if "{op2}" in template and not op2:
                return "", "cellIs {} 缺少第二操作数".format(rule.operator)
            return template.format(cell=cell_ref, op=op1, op1=op1, op2=op2), ""
        return "", "暂不支持的条件格式类型：" + rule.rule_type


# ---------------------------------------------------------------------------
# Evaluator C：DISPLAY_FORMAT_NATIVE（Native Render Baseline，语义是渲染观察）
# ---------------------------------------------------------------------------


class NativeDisplayVerifier:
    """Native Render Verification：DisplayFormat 与基础填充对比（Windows 真值基准）。

    它回答「最终显示填充是否因条件格式而变化」，与 PYTHON/COM_EVALUATE 的
    「规则真假」语义不同；报告与代码中不得称其为布尔求值器。UOS 不提供。
    """

    name = "DISPLAY_FORMAT_NATIVE（Native 渲染基线）"

    def __init__(self, session) -> None:
        self.session = session
        engine_name = str(getattr(session, "engine_name", "") or "").upper()
        self.source = SOURCE_WPS_NATIVE if "WPS" in engine_name else SOURCE_EXCEL_NATIVE

    def extract_issues(self, *, com_workbook, mapping, structure_ranges, period, batch_id,
                       audit_time, org_code, org_name, source_file, workbook_path):
        started = time.monotonic()
        issues = self.session.extract_conditional_format_issues(
            com_workbook, mapping=mapping, structure_ranges=structure_ranges,
            period=period, batch_id=batch_id, audit_time=audit_time,
            org_code=org_code, org_name=org_name, source_file=source_file,
            workbook_path=workbook_path,
        )
        stats = {
            "display_format_read_time": time.monotonic() - started,
            "evaluated_cells": int(getattr(self.session, "last_conditional_evaluated_cells", 0) or 0),
            "unsupported_rule_count": int(getattr(self.session, "last_conditional_unsupported_count", 0) or 0),
            "unsupported_rule_details": [],
        }
        return issues, stats


# ---------------------------------------------------------------------------
# ConditionalFormatService：DAG 唯一入口（求值器由设置中心决定）
# ---------------------------------------------------------------------------


class ConditionalFormatService:
    """规则读取（Direct OOXML）→ 求值（按设置）→ ConditionalFormatResult → Issue。

    无 AUTO、无自动回退；PYTHON/COM_EVALUATE 读取失败或求值失败直接抛错。
    """

    def __init__(self, mode: str, session=None,
                 rule_reader: str = "DIRECT_OOXML") -> None:
        self.mode = effective_evaluator_mode(mode)
        self.rule_reader = str(rule_reader or "DIRECT_OOXML").strip().upper()
        self.session = session

    def extract_issues(
        self, *, workbook_path: Path, ranges, structure_ranges,
        period: str, batch_id: str, audit_time: str, source_file: Path,
        mapping=None, com_workbook=None, org_code: str = "", org_name: str = "",
    ) -> tuple[list, dict]:
        from openpyxl import load_workbook

        total_started = time.monotonic()
        if self.mode == EVALUATOR_COM_DISPLAY:
            return self._extract_via_com_display(
                workbook_path=workbook_path, ranges=ranges,
                structure_ranges=structure_ranges, period=period,
                batch_id=batch_id, audit_time=audit_time, source_file=source_file,
                mapping=mapping, com_workbook=com_workbook,
                org_code=org_code, org_name=org_name,
                total_started=total_started,
            )

        reader_result = self._read_rules(workbook_path)
        by_sheet = _rules_by_sheet(reader_result)
        scanned = _scanned_areas(reader_result, ranges)

        rules_supported = sum(1 for rule in reader_result.rules if rule.supported)
        unsupported_rules = [rule for rule in reader_result.rules if not rule.supported]
        unsupported_details: dict[str, int] = {}
        for rule in unsupported_rules:
            key = rule.raw_metadata.get("raw_type") or rule.rule_type
            unsupported_details[key] = unsupported_details.get(key, 0) + 1

        evaluation_started = time.monotonic()
        results: list[ConditionalFormatResult] = []
        merged: dict = {"evaluated_cells": 0, "unsupported_cells": {}, "errors": [],
                        "com_call_count": 0, "com_evaluate_time": 0.0}
        if self.mode == EVALUATOR_PYTHON:
            evaluator = PythonConditionalRuleEvaluator()
        elif self.mode == EVALUATOR_COM:
            if self.session is None or com_workbook is None:
                raise RuntimeError("COM_EVALUATE 需要已打开的 Excel/WPS COM 工作簿")
            evaluator = ComFormulaEvaluator(self.session, com_workbook)
        else:
            raise RuntimeError(f"未知条件格式求值器：{self.mode}")
        for sheet, areas in scanned.items():
            rules = by_sheet.get(sheet, [])
            if not rules:
                continue
            sheet_results, sheet_stats = evaluator.evaluate_sheet(
                sheet, rules, areas, reader_result)
            results.extend(sheet_results)
            merged["evaluated_cells"] += sheet_stats.get("evaluated_cells", 0)
            merged["unsupported_cells"].update(sheet_stats.get("unsupported_cells", {}))
            merged["errors"].extend(sheet_stats.get("errors", []))
            merged["com_call_count"] += sheet_stats.get("com_call_count", 0)
            merged["com_evaluate_time"] += sheet_stats.get("com_evaluate_time", 0.0)
        evaluation_time = time.monotonic() - evaluation_started

        # 每单元格的 UNSUPPORTED（函数不支持等）并入 details 统计
        for reason in set(merged["unsupported_cells"].values()):
            unsupported_details[reason] = unsupported_details.get(reason, 0) + 1

        build_started = time.monotonic()
        static_book = load_workbook(workbook_path, data_only=True, keep_links=False)
        try:
            areas_as_ranges = [
                type("R", (), {"sheet_name": sheet, "address": _address_of(bounds)})()
                for sheet, bounds_list in scanned.items() for bounds in bounds_list
            ]
            context_started = time.monotonic()
            context = ConditionalFormatContext.build(
                workbook_path=workbook_path, static_book=static_book,
                structure_ranges=structure_ranges)
            context_build_time = time.monotonic() - context_started
            issues, build_stats = build_conditional_issues(
                results, context=context,
                ranges=areas_as_ranges or list(ranges or []),
                workbook_path=workbook_path, source_file=source_file,
                period=period, batch_id=batch_id, audit_time=audit_time,
                org_code=org_code, org_name=org_name,
            )
        finally:
            static_book.close()
        result_build_time = time.monotonic() - build_started

        stats = {
            "mode": self.mode,
            "rule_read_time": reader_result.read_seconds,
            "rule_normalize_time": reader_result.normalize_seconds,
            "evaluation_time": evaluation_time,
            "result_build_time": result_build_time,
            "context_build_time": context_build_time,
            "indicator_resolve_time": build_stats["indicator_resolve_time"],
            "comment_read_time": build_stats["comment_read_time"],
            "message_build_time": build_stats["message_build_time"],
            "comment_backend": build_stats["comment_backend"],
            "unsupported_notes": build_stats["unsupported_notes"],
            "total_time": time.monotonic() - total_started,
            "rules_total": len(reader_result.rules),
            "rules_supported": rules_supported,
            "unsupported_rule_count": len(unsupported_rules),
            "unsupported_rule_details": unsupported_details,
            "unsupported_cells": len(merged["unsupported_cells"]),
            "errors": merged["errors"],
            "evaluated_cells": merged["evaluated_cells"],
            "triggered": len(results),
        }
        if self.mode == EVALUATOR_COM:
            stats["com_call_count"] = merged["com_call_count"]
            stats["com_evaluate_time"] = merged["com_evaluate_time"]
        return issues, stats

    def _read_rules(self, workbook_path: Path) -> CfRuleReadResult:
        """按设置分派 RuleReader：DIRECT_OOXML（默认）或 OPENPYXL（对照）。"""
        if self.rule_reader == "OPENPYXL":
            from .cf_reader import OpenPyxlConditionalRuleReader

            return OpenPyxlConditionalRuleReader().read(workbook_path)
        if self.rule_reader == "DIRECT_OOXML":
            return OoxmlConditionalRuleReader().read(workbook_path)
        raise CfRuleReadError(f"未知的条件格式规则读取方式：{self.rule_reader}")

    def _extract_via_com_display(
        self, *, workbook_path: Path, ranges, structure_ranges,
        period: str, batch_id: str, audit_time: str, source_file: Path,
        mapping, com_workbook, org_code: str, org_name: str,
        total_started: float,
    ) -> tuple[list, dict]:
        """COM_DISPLAY：判定来自 DisplayFormat 渲染，结果构建走统一 Builder。

        RuleReader 读到的规则【不参与触发判定】；只用于生成无批注时的兜底
        描述（RuleMatcher 三态，无法唯一确定时保留 COM 原生描述）。触发格
        的 check_field / message / detail 全部由共享
        ConditionalFormatContext（IndicatorResolver + CommentReader +
        MessageResolver）统一产出，与 PYTHON / COM_EVALUATE 同一套实现。
        """
        from openpyxl import load_workbook

        if self.session is None or com_workbook is None:
            raise RuntimeError("COM_DISPLAY 需要已打开的 Excel/WPS COM 工作簿")
        verifier = NativeDisplayVerifier(self.session)
        raw_issues, verifier_stats = verifier.extract_issues(
            com_workbook=com_workbook, mapping=mapping,
            structure_ranges=structure_ranges, period=period,
            batch_id=batch_id, audit_time=audit_time,
            org_code=org_code, org_name=org_name,
            source_file=source_file, workbook_path=workbook_path,
        )

        # ---- RuleReader 解释层（读取耗时单列，不混入判定耗时） ----
        # 只产生无批注时的兜底文案；触发格与业务结果由统一 Builder 重建。
        reader_result = self._read_rules(workbook_path)
        by_sheet = _rules_by_sheet(reader_result)
        match_counts = {MATCH_EXACT: 0, MATCH_MULTIPLE: 0, MATCH_UNKNOWN: 0}
        results = []
        for issue in raw_issues:
            row, col = _cell_rowcol(issue.target_cell)
            if not row:
                continue
            match = match_rules_for_cell(row, col, by_sheet.get(issue.sheet_name, []))
            match_counts[match.status] += 1
            if match.status == MATCH_EXACT:
                fallback = rule_label_normalized(
                    match.rules[0], target_cell=issue.target_cell)
            elif match.status == MATCH_MULTIPLE:
                fallback = "触发规则存在多个候选：" + " || ".join(
                    rule_label_normalized(rule, target_cell=issue.target_cell)
                    for rule in match.rules)
            else:
                fallback = str(issue.message or "")   # COM 原生描述，不伪造
            results.append(ConditionalFormatResult(
                sheet=issue.sheet_name, cell=issue.target_cell, triggered=True,
                message=fallback,
            ))

        # ---- 统一结果构建：与 PYTHON / COM_EVALUATE 完全同一套实现 ----
        static_book = load_workbook(workbook_path, data_only=True, keep_links=False)
        try:
            context = ConditionalFormatContext.build(
                workbook_path=workbook_path, static_book=static_book,
                structure_ranges=structure_ranges)
            areas = [
                CopyRange(issue.sheet_name, issue.target_cell) for issue in raw_issues
            ]
            issues, build_stats = build_conditional_issues(
                results, context=context, ranges=areas,
                workbook_path=workbook_path, source_file=source_file,
                period=period, batch_id=batch_id, audit_time=audit_time,
                org_code=org_code, org_name=org_name,
            )
        finally:
            static_book.close()
        stats = dict(verifier_stats)
        stats.update({
            "mode": self.mode,
            "rule_read_time": reader_result.read_seconds,
            "rule_normalize_time": reader_result.normalize_seconds,
            "match_counts": match_counts,
            "context_build_time": context.context_build_seconds,
            "indicator_resolve_time": build_stats["indicator_resolve_time"],
            "comment_read_time": build_stats["comment_read_time"],
            "message_build_time": build_stats["message_build_time"],
            "comment_backend": build_stats["comment_backend"],
            "unsupported_notes": build_stats["unsupported_notes"],
            "total_time": time.monotonic() - total_started,
        })
        return issues, stats


def _cell_rowcol(cell_ref: str) -> tuple[int, int]:
    """'D26' → (26, 4)；解析失败返回 (0, 0)。"""
    match = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", str(cell_ref or "").strip(), re.I)
    if not match:
        return (0, 0)
    return (int(match.group(2)), _col_number(match.group(1).upper()))


def _address_of(bounds: tuple[int, int, int, int]) -> str:
    lo_col, lo_row, hi_col, hi_row = bounds
    return f"{_col_letter(lo_col)}{lo_row}:{_col_letter(hi_col)}{hi_row}"
