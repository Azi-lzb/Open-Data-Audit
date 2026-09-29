"""大集中统计系统：复杂校验引擎（单期/跨期/年报/结转/自定义）。

对应 VBA ``DoCompare_AddMyComplexExplain``（B导数比较.bas 1583-1763、
1992-2141）。装载时按规则内出现的 ``[...]`` 裸 token 的指标代码注册
（1340-1378）；行匹配用 token 的 8 段（0/4/5/6/7 相等、1/2 为正则，
IsThisTargetRule 1992-2043）；求值用完整公式：先 {...}（上期）后 [...]
（本期）替换，值 × GetUnitVal(目标单位)（豁免指标不乘），0 → 0.01，
缺失 → 0，Thd → Thd值(万元)×10000（空则 0.01）。命中写「是否说明」（描述
原文），过程写「计算过程」（``计算过程：【原式】->【替换式】``），分隔符
4 空格双竖线；软性描述命中时该列填软性色。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .comparison_engine import get_unit_factor
from .config import CentralConfig
from .models import ComparisonRow

JOIN = "    ||    "

# VBA InitComplexRef_ClearRule（B导数比较.bas 1512-1523）装载期字符归一：
# 全角括号/逗号/引号括号在进入求值前替换为半角；VBA Evaluate 亦原生兼容全角。
_CLEAR_CHAR_MAP = str.maketrans({
    "【": "[", "】": "]", "（": "(", "）": ")", "｛": "{", "｝": "}",
    "，": ",", "\u3000": " ",
})


def _clear_rule(expression: str) -> str:
    return expression.translate(_CLEAR_CHAR_MAP)


class _Rule:
    __slots__ = (
        "rule_id", "description", "inverted", "thd", "soft", "expression",
        "variants", "source_group", "required_frequency_keys",
    )

    def __init__(self, rule_id: str, description: str, inverted: bool, thd: float,
                 soft: bool, expression: str, variants: list[list[str]],
                 source_group: str = "",
                 required_frequency_keys: tuple[str, ...] = ()) -> None:
        self.rule_id = rule_id
        self.description = description
        self.inverted = inverted
        self.thd = thd
        self.soft = soft
        self.expression = expression
        self.variants = variants
        self.source_group = source_group
        self.required_frequency_keys = required_frequency_keys


@dataclass(frozen=True)
class ComplexRuleCompileIssue:
    """一条启用表达式规则无法进入正式执行链时的只读诊断。"""

    rule_id: str
    sheet: str
    excel_row: str
    field: str
    error_code: str
    message: str
    original_expression: str = ""
    offending_token: str = ""
    suggestion: str = ""
    severity: str = "错误"

    def format_user_message(self) -> str:
        location = f"「{self.sheet}」第{self.excel_row or '?'}行"
        rule = self.rule_id or "未编号规则"
        parts = [f"{location} {rule} / {self.field}：{self.message}"]
        if self.offending_token:
            parts.append(f"原始 token：{self.offending_token}")
        if self.suggestion:
            parts.append(f"建议人工复核：{self.suggestion}")
        parts.append("程序未自动修改配置。")
        return "；".join(parts)


@dataclass
class ComplexRuleCompileResult:
    """复杂规则预编译结果；统计按规则身份去重，不按索引桶数量计算。"""

    total_rows: int = 0
    disabled_rows: int = 0
    enabled_rows: int = 0
    compiled_rules: int = 0
    failed_rules: int = 0
    issues: list[ComplexRuleCompileIssue] = field(default_factory=list)
    index: dict[str, list[_Rule]] = field(default_factory=dict)


class ComplexRuleCompileError(ValueError):
    """存在启用但无法执行的复杂规则；正式运行必须阻断而不是静默漏跑。"""

    def __init__(self, result: ComplexRuleCompileResult) -> None:
        self.result = result
        ids = [issue.rule_id or f"第{issue.excel_row}行" for issue in result.issues]
        detail = "、".join(dict.fromkeys(ids))
        issue_lines = "\n".join(f"- {issue.format_user_message()}" for issue in result.issues)
        message = (
            f"配置存在 {result.failed_rules} 条启用但无法执行的表达式规则"
            + (f"：{detail}" if detail else "")
            + "。请先运行配置检查并人工修正；程序未自动修改配置。"
        )
        if issue_lines:
            message += "\n" + issue_lines
        super().__init__(message)


def _is_disabled(value) -> bool:
    return str(value or "").strip() in {"是", "1", "true", "True", "Y", "y"}


def _token_suggestion(token: str, segments: list[str]) -> str:
    """只生成高置信人工建议；绝不修改 token 或继续执行修补后的表达式。"""
    if len(segments) == 7 and len(segments) > 3:
        indicator_and_attr = segments[3].strip()
        for attr in ("发生额", "余额"):
            if indicator_and_attr.endswith(attr) and indicator_and_attr != attr:
                indicator = indicator_and_attr[:-len(attr)].strip()
                fixed = list(segments)
                fixed[3:4] = [indicator, attr]
                return (
                    "疑似“指标”与“数据属性”之间缺少逗号；如业务含义确为此，"
                    f"候选写法为 [{','.join(fixed)}]"
                )
    return (
        "请按 8 段格式人工核对："
        "[业务类,机构类,地区,指标,数据属性,币种,频度,批次]"
    )


def _issue(row: dict, *, field_name: str, error_code: str, message: str,
           expression: str = "", token: str = "", suggestion: str = "") -> ComplexRuleCompileIssue:
    return ComplexRuleCompileIssue(
        rule_id=str(row.get("规则编号") or "").strip(),
        sheet=str(row.get("__工作表__") or "表达式校验").strip() or "表达式校验",
        excel_row=str(row.get("__行号__") or "").strip(),
        field=field_name,
        error_code=error_code,
        message=message,
        original_expression=expression,
        offending_token=token,
        suggestion=suggestion,
    )


def inspect_complex_rules(config: CentralConfig) -> ComplexRuleCompileResult:
    """用正式运行时语法只读预编译全部表达式规则。

    该函数是配置体检与正式执行的共同事实来源。启用规则只允许两种结果：
    成功进入索引，或生成明确 issue；不存在静默 ``continue``。
    """
    from .expression_parser import ExpressionError, parse_cached

    result = ComplexRuleCompileResult(total_rows=len(config.complex_rules))
    for row in config.complex_rules:
        if _is_disabled(row.get("禁用")):
            result.disabled_rows += 1
            continue
        result.enabled_rows += 1
        expression = _clear_rule(str(row.get("校验规则") or "").strip())
        row_issues: list[ComplexRuleCompileIssue] = []
        if not expression:
            row_issues.append(_issue(
                row, field_name="校验表达式", error_code="EMPTY_EXPRESSION",
                message="启用规则的校验表达式为空",
                suggestion="请人工补充正式表达式，或确认该规则是否应停用。",
            ))

        thd_raw = str(row.get("Thd值(万元)") or "").strip()
        if thd_raw:
            try:
                thd = float(thd_raw) * 10000.0
            except ValueError:
                thd = 0.01
                row_issues.append(_issue(
                    row, field_name="Thd值(万元)", error_code="INVALID_THD",
                    message=f"非空 Thd 不是有效数值：{thd_raw}", expression=expression,
                    suggestion="请人工改为有效数值；如业务无需阈值，请按规则约定留空。",
                ))
        else:
            thd = 0.01

        # 所有本期/上期 token 都做结构检查；只有本期 [...] token 用于注册执行索引。
        all_tokens = list(re.finditer(r"\[[^\[\]]*\]|\{[^{}]*\}", expression)) if expression else []
        current_variants: list[list[str]] = []
        required_frequency_keys: list[str] = []
        for match in all_tokens:
            token = match.group(0)
            segments = [seg.strip() for seg in token[1:-1].split(",")]
            if len(segments) != 8:
                row_issues.append(_issue(
                    row, field_name="校验表达式", error_code="TOKEN_SEGMENT_COUNT",
                    message=f"token 共 {len(segments) - 1} 个逗号（8 段式必须有 7 个逗号，即恰好 8 段），大概率取不到数据：{token}",
                    expression=expression, token=token,
                    suggestion=_token_suggestion(token, segments),
                ))
                continue
            if not segments[3].replace("'", "").strip():
                row_issues.append(_issue(
                    row, field_name="校验表达式", error_code="EMPTY_INDICATOR",
                    message="token 第4段“指标”为空", expression=expression, token=token,
                    suggestion="请人工填写指标代码，并核对 8 段槽位位置。",
                ))
                continue
            for position, label in ((1, "机构类"), (2, "地区")):
                pattern = segments[position]
                if not pattern:
                    continue
                try:
                    re.compile(pattern.replace("左中", "[").replace("右中", "]"))
                except re.error as exc:
                    row_issues.append(_issue(
                        row, field_name="校验表达式", error_code="INVALID_REGEX",
                        message=f"{label}正则无法编译：{exc}", expression=expression, token=token,
                        suggestion=f"请人工核对 token 第{position + 1}段“{label}”正则。",
                    ))
            frequency_key = segments[6] + segments[7]
            if frequency_key not in required_frequency_keys:
                required_frequency_keys.append(frequency_key)
            if token.startswith("[") and segments not in current_variants:
                current_variants.append(segments)

        if expression and not current_variants:
            row_issues.append(_issue(
                row, field_name="校验表达式", error_code="NO_CURRENT_TARGET",
                message="规则没有可注册的本期 [...] 目标 token，正式执行时无法挂接到任何指标",
                expression=expression,
                suggestion="请人工确认规则至少包含一个作为触发目标的本期 [...] token。",
            ))

        if expression:
            # 结构 token 先替换为哑值，只做正式 parser 的语法编译，不执行数值计算，
            # 避免合法规则因为探针值恰好触发除零等运行期错误而被误判。
            probe = re.sub(r"\[[^\[\]]*\]|\{[^{}]*\}", "1", expression)
            try:
                parse_cached(probe)
            except ExpressionError as exc:
                row_issues.append(_issue(
                    row, field_name="校验表达式", error_code="EXPRESSION_PARSE_ERROR",
                    message=f"正式表达式解析器无法编译：{exc}", expression=expression,
                    suggestion="请人工对照 VBA 原规则或同类规则核对运算符、函数、括号和逗号。",
                ))

        # 同一规则可能从多个角度发现问题，但 failed_rules 只按规则计 1。
        if row_issues:
            result.issues.extend(row_issues)
            result.failed_rules += 1
            continue

        description = str(row.get("校验描述") or "").strip()
        inverted = bool(str(row.get("取反标识") or "").strip())
        rule = _Rule(
            rule_id=str(row.get("规则编号") or ""), description=description,
            inverted=inverted, thd=thd, soft="软性" in description,
            expression=expression, variants=current_variants,
            source_group=str(row.get("来源") or row.get("来源分组") or "").strip(),
            required_frequency_keys=tuple(required_frequency_keys),
        )
        result.compiled_rules += 1
        for segments in current_variants:
            code = segments[3].replace("'", "").strip()
            result.index.setdefault(code, []).append(rule)

    return result


def compile_complex_rules(config: CentralConfig, *, strict: bool = True) -> dict[str, list[_Rule]]:
    """按指标代码索引全部启用规则；默认严格模式，禁止启用规则静默丢弃。"""
    result = inspect_complex_rules(config)
    if strict and result.failed_rules:
        raise ComplexRuleCompileError(result)
    return result.index

def _target_key(token: str, row: ComparisonRow) -> str:
    """VBA ChangeTarge（2097-2138）：空段继承当前行；含 ^ 的机构/地区段不继承。"""
    body = token.strip().strip("[]{}").replace(" ", "")
    segments = [seg.strip() for seg in body.split(",")]
    while len(segments) < 8:
        segments.append("")
    record = row.record

    def pick(position: int, inherited: str) -> str:
        value = segments[position]
        if value and "^" not in value:
            return value.replace("'", "")
        return inherited

    return "".join((
        segments[0] or record.biz_class,
        pick(1, record.org_code),
        pick(2, record.region_code),
        segments[3].replace("'", "") or record.indicator,
        segments[4] or record.data_attr,
        segments[5] or record.currency,
        segments[6] or record.frequency,
        segments[7] or record.batch,
    ))


def _variant_matches(variant: list[str], row: ComparisonRow) -> bool:
    """VBA IsThisTargetRule：0/4/5/6/7 段相等（空=不限），1/2 段为正则（左中/右中→[]）。"""
    record = row.record
    for position, actual in ((0, record.biz_class), (4, record.data_attr),
                             (5, record.currency), (6, record.frequency), (7, record.batch)):
        expected = variant[position]
        if expected and expected != actual:
            return False
    for position, actual in ((1, record.org_code), (2, record.region_code)):
        pattern = variant[position]
        if not pattern:
            continue
        pattern = pattern.replace("左中", "[").replace("右中", "]")
        try:
            if not re.search(pattern, actual):
                return False
        except re.error:
            return False
    return True


def _expression_stats(
    expr_mode: str, expr_backend: str, office_evaluator, lae_context,
) -> dict:
    """汇总表达式求值观测数据（写入运行日志，见 V3 计划 §十）。"""
    from . import expression_parser

    stats: dict = {
        "mode": expr_mode,
        "backend": expr_backend,
        "provider": (getattr(office_evaluator, "provider", "") or "PYTHON")
        if expr_backend == "OFFICE" else "PYTHON",
        "office_evaluations": getattr(office_evaluator, "evaluation_count", 0),
        "office_cell_path": getattr(office_evaluator, "cell_path_count", 0),
        "expression_parse_count": sum(expression_parser.cache_stats().values()),
        "expression_parse_cache_hits": expression_parser.cache_stats()["hits"],
    }
    if expr_mode == "LAE":
        from . import expression_lae

        stats["ast_cache_hit_count"] = expression_lae.cache_stats()["hits"]
        stats["ast_cache_miss_count"] = expression_lae.cache_stats()["misses"]
        if lae_context is not None and getattr(lae_context, "stats", None):
            stats.update(lae_context.stats)
    return stats


def _cstr(value: float) -> str:
    """VBA CStr 数值文本：最多 15 位有效数字，尾零去除（3204999999.9999995 → 3205000000）。"""
    formatted = f"{value:.15g}"
    if formatted.endswith(".0"):
        formatted = formatted[:-2]
    return formatted


def _token_indicator(token: str) -> str:
    body = token.strip().strip("[]{}")
    segments = [seg.strip() for seg in body.split(",")]
    return segments[3].replace("'", "").strip() if len(segments) > 3 else ""


def _substitute(
    expression: str, row: ComparisonRow, cur_index: dict, pre_index: dict,
    thd: float, unit: float, exempt: set[str],
    registered: set[str] | None = None,
) -> str:
    """先 {...}（上期）再 [...]（本期）：存在非 0 → ×单位因子；存在为 0 → 0.01；缺失 → 0。

    ``registered`` 非 None 时收集每个 token 的目标键（VBA AddHadChecked）。
    """

    def substitute_side(text: str, index: dict, pattern: str) -> str:
        def replace(match: re.Match) -> str:
            key = _target_key(match.group(0), row)
            if registered is not None:
                registered.add(key)
            record = index.get(key)
            if record is None:
                return "0"
            value = record.value if isinstance(record.value, (int, float)) else 0.0
            if value == 0:
                return "0.01"
            factor = 1.0 if _token_indicator(match.group(0)) in exempt else unit
            return _cstr(value * factor)

        return re.sub(pattern, replace, text)

    result = substitute_side(expression, pre_index, r"\{[^{}]*\}")
    result = substitute_side(result, cur_index, r"\[[^\[\]]*\]")
    result = result.replace("Thd", repr(thd)).replace("thd", repr(thd))
    return result


#: 3.1 表达式规则语法：LEGACY_8=旧 8 段（读取「表达式校验」）；
#: FIVE_SEGMENT_V1=新 5 段（读取「表达式校验5段式」，机构/地区为规则级作用域列）。
SCHEMA_LEGACY_8 = "LEGACY_8"
SCHEMA_FIVE_SEGMENT_V1 = "FIVE_SEGMENT_V1"
CENTRAL_EXPRESSION_SCHEMAS = (SCHEMA_LEGACY_8, SCHEMA_FIVE_SEGMENT_V1)

_FIVE_TOKEN_RE = re.compile(r"\[[^\[\]]*\]|\{[^{}]*\}")


class FiveSegmentKeyConflict(ValueError):
    """5 段取数键在数据层不唯一：阻止 FIVE_SEGMENT_V1 正式执行（8 段不受影响）。"""


@dataclass
class FiveSegmentRule:
    """一条 5 段式规则；求值前已 8 段化（复用既有 SBE/LAE 求值器）。"""

    rule_id: str
    description: str
    inverted: bool
    thd: float
    soft: bool
    expression: str            # 5 段原文（展示用）
    expression_8: str          # 8 段化文本（业务类恢复、机构/地区置空；求值用）
    variants: list[list[str]]  # 8 段化挂接 variants（段0=业务类绑定，段1/2=空）
    org_pattern: str
    region_pattern: str
    source_group: str = ""
    required_frequency_keys: tuple[str, ...] = ()


def _five_segment_trigger_inverted(row: dict[str, str]) -> bool:
    """解析 5 段式“是否取反”，并兼容旧“触发方式”。

    正式配置以“是/否”表达是否在表达式为假时触发；旧配置的
    “表达式成立 / 表达式不成立”仍可读取。保留“取反标识”仅用于内存调用
    兼容。
    """
    inverted = str(row.get("是否取反") or "").strip().casefold()
    if inverted:
        return inverted in {"是", "1", "true", "y"}
    trigger = str(row.get("触发方式") or "").strip()
    if trigger:
        return trigger == "表达式不成立"
    return str(row.get("取反标识") or "").strip().casefold() in {
        "是", "1", "true", "y",
    }


def legacy_business_bindings(config: CentralConfig) -> dict[str, list[str]]:
    """旧 8 段表逐规则提取每个 token 的业务类段（段0）序列，供 5 段 8 段化配对。"""
    bindings: dict[str, list[str]] = {}
    for row in config.complex_rules:
        rule_id = str(row.get("规则编号") or "").strip()
        if not rule_id:
            continue
        expression = _clear_rule(str(row.get("校验规则") or "").strip())
        bindings[rule_id] = [
            token[1:-1].split(",")[0].strip()
            for token in _FIVE_TOKEN_RE.findall(expression)
        ]
    return bindings


def _eight_segment_expression(expression: str, biz_sequence: list[str]) -> str:
    """5 段表达式 → 8 段化：业务类恢复为旧表绑定值，机构/地区置空（规则级列管）。"""
    counter = {"i": 0}

    def repl(match: re.Match) -> str:
        position = counter["i"]
        counter["i"] += 1
        biz = biz_sequence[position] if position < len(biz_sequence) else ""
        segments = [part.strip() for part in match.group(0)[1:-1].split(",")]
        while len(segments) < 5:
            segments.append("")
        opener, closer = match.group(0)[0], match.group(0)[-1]
        return opener + ",".join((biz, "", "", *segments)) + closer

    return _FIVE_TOKEN_RE.sub(repl, expression)


def audit_five_segment_keys(dataset) -> list[dict]:
    """机构+地区桶内 5 段取数键唯一性审计（只读；字段与索引键同源）。

    返回冲突列表：同一 (机构, 地区, 指标, 属性, 币种, 频度, 批次) 对应
    多条记录（含同业务类重复——无论取哪条都缺乏唯一性依据）。
    """
    groups: dict[tuple, list] = {}
    for record in dataset.records:
        key5 = (record.org_code, record.region_code, record.indicator,
                record.data_attr, record.currency, record.frequency, record.batch)
        groups.setdefault(key5, []).append(record)
    conflicts: list[dict] = []
    for key5, records in sorted(groups.items()):
        if len(records) < 2:
            continue
        conflicts.append({
            "org": key5[0], "region": key5[1], "indicator": key5[2],
            "attr": key5[3], "currency": key5[4], "frequency": key5[5],
            "batch": key5[6],
            "biz_classes": sorted({r.biz_class for r in records}),
            "count": len(records),
            "values": [r.value for r in records],
        })
    return conflicts


def inspect_five_segment_rules(config: CentralConfig) -> ComplexRuleCompileResult:
    """5 段式规则只读预编译：token 结构/作用域正则/左中右中残留/配对绑定。

    业务类绑定来自旧 8 段表同编号规则（按 token 顺序配对）——5 段 token
    砍掉业务类段后，挂接与跨业务类取数语义由 8 段化恢复（见 enrich）。
    启用规则只允许两种结果：进入索引，或生成明确 issue。
    """
    from .expression_parser import ExpressionError, parse_cached

    bindings = legacy_business_bindings(config)
    result = ComplexRuleCompileResult(total_rows=len(config.five_segment_rules))
    for row in config.five_segment_rules:
        rule_id = str(row.get("规则编号") or "").strip()
        # 「启用」列语义：是=启用、否=停用（与 8 段表一致；迁移表无空值）。
        if str(row.get("启用") or "").strip() in {"否", "0", "false", "False"}:
            result.disabled_rows += 1
            continue
        result.enabled_rows += 1
        expression = _clear_rule(str(row.get("校验表达式") or "").strip())
        row_issues: list[ComplexRuleCompileIssue] = []
        if not expression:
            row_issues.append(_issue(
                row, field_name="校验表达式", error_code="EMPTY_EXPRESSION",
                message="启用规则的校验表达式为空",
                suggestion="请人工补充正式表达式，或确认该规则是否应停用。"))

        thd_raw = str(row.get("容差值(万元)") or "").strip()
        thd = 0.01
        if thd_raw:
            try:
                thd = float(thd_raw) * 10000.0
            except ValueError:
                row_issues.append(_issue(
                    row, field_name="容差值(万元)", error_code="INVALID_THD",
                    message=f"非空容差不是有效数值：{thd_raw}", expression=expression,
                    suggestion="请人工改为有效数值；如业务无需阈值，请按规则约定留空。"))

        # 新表使用直观的“是否取反”：是=表达式不成立时触发，否=成立时触发。
        # 旧“触发方式”表头仍由加载器兼容，不在这里误报旧配置。
        if "是否取反" in row:
            inversion_text = str(row.get("是否取反") or "").strip()
            if inversion_text not in {"是", "否"}:
                row_issues.append(_issue(
                    row, field_name="是否取反", error_code="INVALID_INVERSION_FLAG",
                    message=f"是否取反只能填写“是”或“否”，当前：{inversion_text!r}",
                    expression=expression,
                    suggestion="“是”表示表达式不成立时触发；“否”表示表达式成立时触发。"))

        # 作用域列：机构/地区必须为标准正则；禁止左中/右中转写。
        scope_patterns: dict[str, str] = {}
        for column, label in (("机构类代码", "机构类"), ("地区代码", "地区")):
            pattern = str(row.get(column) or "").strip()
            if not pattern:
                continue
            if "左中" in pattern or "右中" in pattern:
                row_issues.append(_issue(
                    row, field_name=column, error_code="LEGACY_REGEX_TRANSLATION",
                    message=f"{label}作用域包含「左中/右中」转写：{pattern}",
                    suggestion="5 段式配置应直接使用标准正则表达式。"))
                continue
            try:
                scope_patterns[column] = pattern
                re.compile(pattern)
            except re.error as exc:
                row_issues.append(_issue(
                    row, field_name=column, error_code="INVALID_REGEX",
                    message=f"{label}正则无法编译：{exc}", expression=expression,
                    suggestion=f"请人工核对「{column}」列的正则写法。"))

        tokens = _FIVE_TOKEN_RE.findall(expression) if expression else []
        binding = bindings.get(rule_id)
        if binding is None:
            row_issues.append(_issue(
                row, field_name="规则编号", error_code="LEGACY_BINDING_MISSING",
                message="5 段式规则在兼容 8 段表「表达式校验」中缺少同编号规则（业务类绑定缺失）",
                suggestion="两表应一一对应；请检查 3.1 配置迁移完整性。"))
        elif len(binding) != len(tokens):
            row_issues.append(_issue(
                row, field_name="校验表达式", error_code="TOKEN_COUNT_MISMATCH",
                message=f"与兼容 8 段表的 token 数不一致：5 段 {len(tokens)} 个，8 段 {len(binding)} 个",
                suggestion="两表应逐 token 一一对应；请检查 3.1 配置迁移完整性。"))

        current_variants: list[list[str]] = []
        biz_sequence: list[str] = []
        required_frequency_keys: list[str] = []
        for index, token in enumerate(tokens):
            segments = [part.strip() for part in token[1:-1].split(",")]
            if len(segments) != 5:
                row_issues.append(_issue(
                    row, field_name="校验表达式", error_code="TOKEN_SEGMENT_COUNT",
                    message=f"token 共 {len(segments) - 1} 个逗号（5 段式必须有 4 个逗号，即恰好 5 段），大概率取不到数据：{token}",
                    expression=expression, token=token,
                    suggestion="5 段式 token 应为 [指标代码,数据属性,币种,频度,批次]。"))
                continue
            if not segments[0].replace("'", "").strip():
                row_issues.append(_issue(
                    row, field_name="校验表达式", error_code="EMPTY_INDICATOR",
                    message="token 第1段“指标代码”为空", expression=expression, token=token,
                    suggestion="请人工填写指标代码。"))
            if "左中" in token or "右中" in token:
                row_issues.append(_issue(
                    row, field_name="校验表达式", error_code="LEGACY_REGEX_TRANSLATION",
                    message=f"token 含「左中/右中」转写：{token}",
                    suggestion="5 段式配置应直接使用标准正则表达式。"))
                continue
            biz = binding[index] if binding and index < len(binding) else ""
            biz_sequence.append(biz)
            frequency_key = segments[3] + segments[4]
            if frequency_key not in required_frequency_keys:
                required_frequency_keys.append(frequency_key)
            if token.startswith("[") and [biz, "", "", *segments] not in current_variants:
                current_variants.append([biz, "", "", *segments])

        if expression and not current_variants:
            row_issues.append(_issue(
                row, field_name="校验表达式", error_code="NO_CURRENT_TARGET",
                message="规则没有可注册的本期 [...] 目标 token，正式执行时无法挂接到任何指标",
                expression=expression,
                suggestion="请人工确认规则至少包含一个作为触发目标的本期 [...] token。"))
        if expression:
            probe = re.sub(r"\[[^\[\]]*\]|\{[^{}]*\}", "1", expression)
            try:
                parse_cached(probe)
            except ExpressionError as exc:
                row_issues.append(_issue(
                    row, field_name="校验表达式", error_code="EXPRESSION_PARSE_ERROR",
                    message=f"正式表达式解析器无法编译：{exc}", expression=expression,
                    suggestion="请人工对照运算符、函数、括号和逗号。"))
        if row_issues:
            result.issues.extend(row_issues)
            result.failed_rules += 1
            continue

        description = str(row.get("规则说明") or "").strip()
        rule = FiveSegmentRule(
            rule_id=rule_id,
            description=description,
            inverted=_five_segment_trigger_inverted(row),
            thd=thd, soft="软性" in description,
            expression=expression,
            expression_8=_eight_segment_expression(expression, biz_sequence),
            variants=current_variants,
            org_pattern=scope_patterns.get("机构类代码", ""),
            region_pattern=scope_patterns.get("地区代码", ""),
            source_group=str(row.get("来源分组") or row.get("来源") or "").strip(),
            required_frequency_keys=tuple(required_frequency_keys),
        )
        result.compiled_rules += 1
        for variant in current_variants:
            result.index.setdefault(variant[3], []).append(rule)
    return result


def compile_five_segment_rules(config: CentralConfig, *, strict: bool = True):
    """按指标代码索引全部启用的 5 段式规则；strict 下失败即阻断。"""
    result = inspect_five_segment_rules(config)
    if strict and result.failed_rules:
        raise ComplexRuleCompileError(result)
    return result.index


def _safe_scope_match(pattern: str, text: str) -> bool:
    """规则级作用域正则匹配；空 pattern=不限制。非法正则按不匹配处理
    （编译期检查已在 5 段体检中报 ERROR，运行期兜底不抛）。"""
    if not pattern:
        return True
    try:
        return re.search(pattern, text) is not None
    except re.error:
        return False


def _filter_rules_by_date(rules_by_code: dict[str, list], current_date: str) -> dict[str, list]:
    """按本期结转日选择结转/非结转表达式，与 VBA 的表级装载互斥。"""
    january_first = str(current_date or "").endswith("-01-01")
    filtered = {
        code: [rule for rule in rules if (rule.source_group == "结转") == january_first]
        for code, rules in rules_by_code.items()
    }
    return {code: rules for code, rules in filtered.items() if rules}


def enrich_complex_rules(
    rows: list[ComparisonRow],
    config: CentralConfig,
    *,
    current_index: dict,
    previous_index: dict,
    target_unit: str = "亿元",
    exempt: set[str] | None = None,
    soft_color: int = 46,
    fres: set[str] | None = None,
    expr_mode: str = "SBE",
    expr_backend: str = "PYTHON",
    expression_dates: tuple[str, str] = ("", ""),
    office_evaluator=None,
    stats_out: dict | None = None,
    expression_schema: str = SCHEMA_LEGACY_8,
) -> None:
    """对全部比较结果行执行复杂校验；原地写 是否说明/计算过程/填充。

    ``fres`` 为两期数据的频度键（VBA dFres：{"", 各频度, 频度+批次}）；
    规则内任一 token 的 频度+批次 不在其中时整条跳过（VBA notfreNum）。

    ``expr_mode``/``expr_backend`` 为表达式 2×2 执行组合（默认 SBE+PYTHON =
    现行正式行为）；LAE 走 expression_lae 惰性 AST 求值；OFFICE 须由调用方
    注入任务级 office_eval.OfficeEvaluationAdapter（provider 一次解析、全程固定）。
    """
    from . import expression_backend
    from .expression_parser import ExpressionError, evaluate_expression

    five_segment_mode = expression_schema == SCHEMA_FIVE_SEGMENT_V1
    if expression_schema not in CENTRAL_EXPRESSION_SCHEMAS:
        raise ValueError(
            f"未知 3.1 表达式规则语法：{expression_schema!r}"
            f"（应为 {'/'.join(CENTRAL_EXPRESSION_SCHEMAS)}）")
    if five_segment_mode:
        # FIVE_SEGMENT_V1：读取「表达式校验5段式」；规则已 8 段化
        # （业务类绑定恢复、机构/地区置空由规则级列过滤），求值器零改动复用。
        rules_by_code = compile_five_segment_rules(config)
    else:
        rules_by_code = compile_complex_rules(config)
    if not rules_by_code:
        return
    rules_by_code = _filter_rules_by_date(rules_by_code, expression_dates[0])
    if not rules_by_code:
        return
    if expr_backend == expression_backend.BACKEND_OFFICE and office_evaluator is None:
        raise expression_backend.ExpressionBackendError(
            "表达式求值方式=OFFICE：未提供 OfficeEvaluationAdapter；不静默回退 Python。")
    if fres:
        kept = {}
        for code, rules in rules_by_code.items():
            filtered = [
                rule for rule in rules
                if all(key in fres for key in rule.required_frequency_keys)
            ]
            if filtered:
                kept[code] = filtered
        rules_by_code = kept
    unit = get_unit_factor(target_unit)
    exempt = exempt or set()
    lae_context: dict[str, object] | None = None
    if expr_mode == expression_backend.MODE_LAE:
        from .expression_lae import ExpressionContext

        lae_context = ExpressionContext(
            current_date=expression_dates[0], previous_date=expression_dates[1],
            current_index=current_index, previous_index=previous_index,
            unit_factor=unit, exempt_indicators=frozenset(exempt),
            stats={} if stats_out is not None else None,
        )
    soft_checked: dict[str, set[str]] = {}
    for row in rows:
        applicable = rules_by_code.get(row.record.indicator)
        if not applicable:
            continue
        seen_rules: set[int] = set()
        for rule in applicable:
            if id(rule) in seen_rules:
                continue
            seen_rules.add(id(rule))
            matched = next((v for v in rule.variants if _variant_matches(v, row)), None)
            if matched is None:
                continue
            if five_segment_mode:
                # 5 段式规则级作用域：机构/地区列正则过滤当前比较行
                # （标准正则；不匹配的机构整行跳过，禁止跨机构笛卡尔积）。
                if not _safe_scope_match(rule.org_pattern, row.record.org_code):
                    continue
                if not _safe_scope_match(rule.region_pattern, row.record.region_code):
                    continue
            evaluation_text = (
                rule.expression_8 if five_segment_mode else rule.expression)
            # 软性规则去重（VBA AddHadChecked/IsNotChecked）：行键已被本规则
            # 的 token 目标登记过时不再执行。
            registered: set[str] | None = None
            if rule.soft:
                registered = soft_checked.setdefault(evaluation_text, set())
                if row.record.key() in registered:
                    continue
            if lae_context is not None:
                lae_context.row = row
                lae_context.thd = rule.thd
            result = expression_backend.evaluate_expression_rule(
                evaluation_text,
                trigger_inverted=rule.inverted,
                mode=expr_mode, backend=expr_backend,
                thd=rule.thd,
                substitute=lambda text, _row=row, _registered=registered: _substitute(
                    text, _row, current_index, previous_index, rule.thd, unit, exempt,
                    registered=_registered,
                ).replace(" ", "").lower(),
                evaluate_text=lambda text: evaluate_expression(text, {"Thd": rule.thd}),
                context=lae_context,
                office_evaluator=office_evaluator,
            )
            if registered is not None:
                registered.add(row.record.key())
            hit = result.hit
            if expr_mode == expression_backend.MODE_SBE:
                # VBA RoundRule 为 ByRef：调用后 tRule 被小写化并去掉空格（round 包装
                # 只影响求值副本）；此处同步该行为，求值与过程文本保持一致。
                substituted = result.substituted
                display_text = (
                    rule.expression if five_segment_mode else rule.expression)
                process = f"计算过程：【{display_text}】->【{substituted}】"
                if result.status == "ERROR":
                    process = f" 计算【{rule.description}】校验时发生错误--{result.reason}。" + process
                elif result.status == "OK" and result.reason:
                    process = f" 计算【{rule.description}】校验{result.reason}--。" + process
            else:
                process = (
                    f"计算过程：【{rule.expression}】->【LAE+{expr_backend}"
                    f" status={result.status}】"
                )
                if result.status != "OK":
                    process = f" 计算【{rule.description}】{result.reason}。" + process
            if not hit:
                continue
            row.need_explain = (
                rule.description if not row.need_explain
                else row.need_explain + JOIN + rule.description
            )
            process = process.replace("左中", "[").replace("右中", "]")
            row.process = process if not row.process else row.process + JOIN + process
            if rule.soft and (not row.fill_color or row.fill_scope != "row"):
                row.fill_color = soft_color
                row.fill_scope = "remark"
    if stats_out is not None:
        stats_out.update(_expression_stats(expr_mode, expr_backend, office_evaluator, lae_context))
