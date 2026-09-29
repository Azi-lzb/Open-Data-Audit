"""兼容现有 [本期指标] / {上期指标} 语法的 V2 规则引擎。"""

from __future__ import annotations

import ast
import hashlib
import re
from collections import defaultdict
from decimal import Decimal, localcontext
from typing import Any

from .models import AuditFinding, IndicatorRecord, RelatedRecord, ReportAuditConfig, RuleConfig, STATUS_ABNORMAL
from .unit_conversion import AMOUNT_DATA_TYPES


_PLACEHOLDER = re.compile(r"(\[|\{)([^\]}]+)[\]}]")


def _id(*parts: str) -> str:
    return hashlib.sha1("|".join(parts).encode()).hexdigest()[:16]


def _num(value: Any) -> float | Decimal:
    try:
        if isinstance(value, Decimal):
            return value
        return float(value) if value not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def _AND(*values): return all(bool(item) for item in values)
def _OR(*values): return any(bool(item) for item in values)
def _NOT(value): return not bool(value)


class _DecimalLiteralTransformer(ast.NodeTransformer):
    """将表达式中的数值字面量按原始文本转换为 Decimal。"""

    def __init__(self, source: str) -> None:
        self.source = source

    def visit_Constant(self, node: ast.Constant) -> ast.AST:  # noqa: N802 - ast API name
        if isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
            # 从 AST 源位置取回字面量，不能先走 Python float 再 repr，
            # 否则 15 位以上的小数会在 Decimal 构造前丢失原始位数。
            literal = ast.get_source_segment(self.source, node) or repr(node.value)
            return ast.copy_location(
                ast.Call(
                    func=ast.Name(id="_D", ctx=ast.Load()),
                    args=[ast.Constant(value=literal)],
                    keywords=[],
                ),
                node,
            )
        return node


def _decimal_expression(text: str) -> Any:
    """Compile a restricted rule expression with Decimal numeric literals."""
    tree = ast.parse(text, mode="eval")
    tree = _DecimalLiteralTransformer(text).visit(tree)
    ast.fix_missing_locations(tree)
    return compile(tree, "<s2-rule>", "eval")


def _decimal_literal(value: str) -> Decimal:
    """Build a Decimal directly from the original numeric token, without rounding."""
    return Decimal(value)


def _indicator_code(inner: str) -> str:
    """占位符内部文本 → 指标代码。

    兼容旧 VBA 八段语法 ``[,,,20203045,]``（逗号分段，第 4 段为指标代码；
    V1 period_compare 同口径）与简写 ``[20201028]``。此前直接拿整段
    （含逗号）作 lookup key，导致 8 段规则永远“缺少指标”。
    """
    text = inner.strip()
    if "," not in text:
        return text
    parts = [segment.strip() for segment in text.split(",")]
    if len(parts) >= 4 and parts[3]:
        return parts[3]
    non_empty = [segment for segment in parts if segment]
    return non_empty[-1] if non_empty else text


def _evaluate(rule: RuleConfig, current: dict[str, IndicatorRecord], previous: dict[str, IndicatorRecord], *, config: ReportAuditConfig, amount_codes: set[str] | None = None) -> tuple[bool, str, str, set[str]]:
    """返回 命中、代入表达式、错误、涉及指标。"""
    involved: set[str] = set()
    errors: list[str] = []
    def repl(match: re.Match) -> str:
        code = _indicator_code(match.group(2))
        involved.add(code)
        record = (current if match.group(1) == "[" else previous).get(code)
        if record is None:
            errors.append(f"缺少指标 {code} 的当期数据")
            return "0.0"
        value = record.normalized_value
        # 行存在但没报有效数值（空串/文字）：与缺失同等对待，禁止拿 0
        # 参与比较——否则“充足率低于监管线”这类规则会把空报机构全部
        # 误判成“低于监管线 异常”。
        if value is None or (isinstance(value, str) and not value.strip()):
            errors.append(f"缺少指标 {code} 的有效数值")
            return "0.0"
        if isinstance(value, str):
            try:
                number = float(value.replace(",", "").strip())
            except ValueError:
                errors.append(f"指标 {code} 的值不是数值：{value!r}")
                return "0.0"
            return str(_num(number))
        parsed = _num(value)
        return format(parsed, "f") if isinstance(parsed, Decimal) else repr(parsed)
    rendered = _PLACEHOLDER.sub(repl, rule.expression)
    if errors:
        return False, rendered, "；".join(sorted(set(errors))), involved
    text = rendered.replace("===", "==").replace("<>", "!=")
    text = re.sub(r"(?<![<>=!])=(?!=)", "==", text)
    text = re.sub(r"\bAND\s*\(", "_AND(", text, flags=re.I)
    text = re.sub(r"\bOR\s*\(", "_OR(", text, flags=re.I)
    text = re.sub(r"\bNOT\s*\(", "_NOT(", text, flags=re.I)
    if re.search(r"\bThd\b", text, flags=re.I):
        if rule.threshold is None:
            return False, rendered, "表达式使用 Thd 但规则未填写 Thd值(万元)", involved
        threshold = Decimal(str(rule.threshold))
        # 配置列仍按万元填写；只在表达式依赖金额指标时，先把阈值精确
        # 换算到本次统一使用的输出单位，再参与规则计算。
        if amount_codes and involved.intersection(amount_codes):
            from .unit_conversion import convert_amount

            threshold = convert_amount(
                threshold, "万元", config.unit_settings.output_file_unit,
            )
        text = re.sub(r"\bThd\b", format(threshold, "f"), text, flags=re.I)
    try:
        with localcontext() as context:
            context.prec = 50
            value = bool(eval(  # noqa: S307 - 配置表达式仅在受限命名空间运行
                _decimal_expression(text),
                {"__builtins__": {}, "_D": _decimal_literal},
                {"_AND": _AND, "_OR": _OR, "_NOT": _NOT},
            ))
    except Exception as exc:
        return False, rendered, f"表达式无法执行：{exc}", involved
    return (not value if rule.invert else value), rendered, "", involved


def _related(records: dict[str, IndicatorRecord], codes: set[str]) -> tuple[RelatedRecord, ...]:
    return tuple(RelatedRecord(code, record.indicator_name, record.form_code, record.period, record.normalized_value, record.source) for code, record in sorted(records.items()) if code in codes)


def _display_value(record: IndicatorRecord | None, *, missing: str = "缺少指标") -> str:
    if record is None:
        return missing
    if record.normalized_value in (None, ""):
        return "无数据"
    return str(record.normalized_value)


def _expression_value_details(expression: str, current: dict[str, IndicatorRecord], previous: dict[str, IndicatorRecord]) -> str:
    """Render every referenced indicator with its explicit period side."""
    entries: list[str] = []
    seen: set[tuple[str, str]] = set()
    for match in _PLACEHOLDER.finditer(expression):
        period_label = "本期" if match.group(1) == "[" else "上期"
        code = _indicator_code(match.group(2))
        key = (period_label, code)
        if key in seen:
            continue
        seen.add(key)
        record = (current if period_label == "本期" else previous).get(code)
        name = record.indicator_name if record else ""
        label = f"{period_label} {code}"
        if name:
            label += f" {name}"
        entries.append(f"{label}={_display_value(record)}")
    return "；".join(entries)


def _rule_type(rule: RuleConfig) -> str:
    if rule.rule_type.startswith("累计不降"):
        return "累计不降"
    if rule.rule_type == "表达式":
        return "表达式"
    return "未识别"


class RuleEngine:
    def run(self, current: dict[tuple[str, str, str], IndicatorRecord], previous: dict[tuple[str, str, str], IndicatorRecord], config: ReportAuditConfig) -> tuple[list[AuditFinding], list[AuditFinding]]:
        cur_by: dict[str, dict[str, IndicatorRecord]] = defaultdict(dict)
        pre_by: dict[str, dict[str, IndicatorRecord]] = defaultdict(dict)
        for record in current.values(): cur_by[record.institution_id][record.indicator_code] = record
        for record in previous.values(): pre_by[record.institution_id][record.indicator_code] = record
        findings: list[AuditFinding] = []
        quality: list[AuditFinding] = []
        for rule in config.rules:
            if rule.disabled:
                continue
            for institution_id in sorted(set(cur_by) | set(pre_by)):
                cur, pre = cur_by[institution_id], pre_by[institution_id]
                anchor = next(iter(cur.values()), next(iter(pre.values()), None))
                if anchor is None:
                    continue
                if rule.rule_type.startswith("累计不降"):
                    codes = [item for item in re.split(r"[,，、;；\s]+", rule.expression) if item]
                    violated = [code for code in codes if code in cur and code in pre and _num(cur[code].normalized_value) < _num(pre[code].normalized_value)]
                    if not violated:
                        continue
                    # 累计不降是一组逐指标规则：每个违反的指标均独立形成审核结果，
                    # 不能压缩为“一机构 + 一规则”一行，否则本期/上期值无法追溯。
                    for code in violated:
                        current_record, previous_record = cur[code], pre[code]
                        current_value = current_record.normalized_value
                        previous_value = previous_record.normalized_value
                        difference = _num(current_value) - _num(previous_value)
                        findings.append(AuditFinding(
                            finding_id=_id("规则", rule.rule_id, institution_id, code, current_record.period),
                            audit_type="规则", status=STATUS_ABNORMAL, severity=rule.severity,
                            region=current_record.region,
                            handling_branch=current_record.handling_branch, institution_type=current_record.institution_type,
                            institution_name=current_record.institution_name, social_credit_code=current_record.social_credit_code,
                            form_code=current_record.form_code,
                            form_codes=(current_record.form_code,) if current_record.form_code else (), period=current_record.period,
                            indicator_code=code, indicator_name=current_record.indicator_name,
                            # 审核说明必须是稳定的规则文案，便于按同类规则筛选；逐笔值放在计算过程。
                            rule_id=rule.rule_id, rule_description=rule.description, description=rule.description,
                            current_value=current_value, previous_value=previous_value, difference_value=difference,
                            calculation_trace=f"{code}: {current_value} < {previous_value}",
                            rule_type=_rule_type(rule),
                            value_details=f"{code} {current_record.indicator_name}: 本期={current_value}；上期={previous_value}",
                            source_refs=(current_record.source, rule.source) if rule.source else (current_record.source,),
                            related_records=_related({**pre, **cur}, {code}),
                        ))
                    continue
                else:
                    amount_codes = {
                        code for code, indicator in config.indicators.items()
                        if indicator.data_type in AMOUNT_DATA_TYPES
                    }
                    hit, rendered, error, involved = _evaluate(
                        rule, cur, pre, config=config, amount_codes=amount_codes,
                    )
                    if error:
                        # 审核说明必须写真实原因（缺少哪些指标的数据），
                        # 不得用规则描述冒充——否则“缺少指标”会显示成
                        # “低于监管线 异常”，完全误导。
                        quality.append(AuditFinding(_id("规则配置", rule.rule_id, institution_id, error), "校验规则", STATUS_ABNORMAL, "关注", institution_name=anchor.institution_name, social_credit_code=anchor.social_credit_code, period=anchor.period, rule_id=rule.rule_id, rule_description="", description=f"规则无法执行：{error}（规则：{rule.description}）", calculation_trace=rendered, rule_type=_rule_type(rule), value_details=_expression_value_details(rule.expression, cur, pre)))
                        continue
                    if not hit:
                        continue
                related = _related({**pre, **cur}, involved)
                forms = tuple(sorted({item.form_code for item in related if item.form_code}))
                primary = related[0] if related else anchor
                findings.append(AuditFinding(
                    finding_id=_id("规则", rule.rule_id, institution_id, anchor.period), audit_type="规则", status=STATUS_ABNORMAL,
                    severity=rule.severity, region=anchor.region,
                    handling_branch=anchor.handling_branch, institution_type=anchor.institution_type,
                    institution_name=anchor.institution_name, social_credit_code=anchor.social_credit_code,
                    form_code=primary.form_code, form_codes=forms, period=anchor.period,
                    indicator_code=primary.indicator_code, indicator_name=primary.indicator_name,
                    rule_id=rule.rule_id, rule_description=rule.description, description=rule.description,
                    calculation_trace=rendered + ("（取反）" if rule.invert else ""),
                    rule_type=_rule_type(rule), value_details=_expression_value_details(rule.expression, cur, pre),
                    source_refs=(primary.source, rule.source) if rule.source else (primary.source,), related_records=related,
                ))
        return findings, quality
