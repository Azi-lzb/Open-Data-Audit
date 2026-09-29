"""大集中统计系统：特殊指标与累计指标引擎。

对应 VBA ``DoCompare_AddMyExplain``（B导数比较.bas 2150-2420）与
``当年累计指标比上期不应减少``（1768-1792、2445-2481）。阈值单位为亿元
（豁免指标 u=1）；命中文案写「是否说明」，多条以 4 空格双竖线连接；
全部命中均含「软性」时该列填软性色。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import Decimal

from .comparison_engine import get_unit_factor
from .config import CentralConfig
from .models import CentralRecord, ComparisonRow, make_record_key

JOIN = "    ||    "
SPECIAL_UNIT = "亿元"

# 配置检查与正式动作分派共用同一套可识别动作口径。末尾三个动作在
# apply_special_action 中按关键词分派，其余均为精确匹配。
_EXACT_SPECIAL_ACTIONS = frozenset({
    "指标不应有数，需说明。", "指标有变动需说明。", "指标变动或变幅超过阈值需说明。",
    "指标变幅异常需说明。", "当年累计指标比上期不应减少。",
    "历史累计指标比上期不应减少。", "指标比上期增加需说明。",
    "指标比上期减少需说明。", "指标新增需说明。", "指标结清需说明。",
    "指标为负数需核实。", "指标应为整数。", "指标应能被某数整除。",
    "指标本期余额应小于（上期余额+本期发生额）。",
    "指标本期余额应小于（上期余额+本期当年发生额-上期当年累计发生额）。",
    "指标需要对应存在。", "指标应相等核查。",
})


def supports_special_action(action: str) -> bool:
    """正式分派器能否识别动作；不判断数据条件是否命中。"""
    text = str(action or "").strip()
    return text in _EXACT_SPECIAL_ACTIONS or "相除" in text or "太小核查" in text or "相减" in text


def _is_disabled(value) -> bool:
    return str(value or "").strip() in {"是", "1", "true", "True", "Y", "y"}


def _split_multi_codes(raw: str) -> list[str]:
    return [part.replace("'", "").strip() for part in str(raw or "").split("/") if part.strip()]


def _to_float(value) -> "float | None":
    try:
        return float(str(value or "").strip())
    except (TypeError, ValueError):
        return None


def _display_number(value: float | Decimal) -> str:
    """保留配置数值的十进制显示，不把 0 写成空白或科学计数法。"""
    shown = format(Decimal(str(value)), "f")
    return shown.rstrip("0").rstrip(".") if "." in shown else shown


def _special_threshold_text(rule: dict, action: str, facts: _RowFacts) -> str:
    """只展示动作实际使用的阈值维度；不从说明文字反向解析阈值。"""
    rmb = _to_float(rule.get("人民币阀值(亿元)")) or 0.0
    usd = _to_float(rule.get("美元阀值(亿美元)")) or 0.0
    amplitude = _to_float(rule.get("绝对值变幅(%)")) or 0.0
    amount = usd if facts.is_usd and usd else rmb
    amount_unit = "亿美元" if facts.is_usd and (usd or not rmb) else "亿元"
    amount_text = f"阈值{_display_number(amount)}{amount_unit}"
    amplitude_text = f"变动幅度{_display_number(Decimal(str(amplitude)) * 100)}%"
    if action == "指标变幅异常需说明。":
        return amplitude_text
    if action in {"指标有变动需说明。", "指标变动或变幅超过阈值需说明。"}:
        parts = [amount_text] if amount or not amplitude else []
        if amplitude:
            parts.append(amplitude_text)
        return "、".join(parts)
    if action == "指标应能被某数整除。":
        return f"阈值{_display_number(rmb)}（除数）"
    if "相除" in action or "太小核查" in action:
        return f"比值阈值{_display_number(amount)}"
    if action in {"指标新增需说明。", "指标结清需说明。"}:
        # 这两个动作的现行判定仅使用人民币阈值列，不读取美元阈值。
        return f"阈值{_display_number(rmb)}亿元"
    if action in {
        "指标比上期增加需说明。", "指标比上期减少需说明。",
    } or "相减" in action:
        return amount_text
    return "阈值不适用"


def _special_output_message(rule: dict, hit_message: str, facts: _RowFacts) -> str:
    """规则编号－配置动作－有效阈值－配置备注；保留动态判定细节。"""
    action = str(rule.get("规则动作") or rule.get("备注") or "").strip()
    if rule.get("类型") == "累计不应下降" and not rule.get("规则动作"):
        action = "当年累计指标比上期不应减少。"
    rule_id = str(rule.get("规则编号") or "").strip() or "未编号"
    note = str(rule.get("配置备注") or "").strip()
    if hit_message != action:
        detail = f"判定：{hit_message.strip()}"
        note = f"{note}；{detail}" if note else detail
    head = f"{rule_id}-{action}-{_special_threshold_text(rule, action, facts)}"
    return f"{head}-{note}" if note else head


@dataclass
class PreparedActionRules:
    """已装载的动作规则集：现行路径与 V3 路径共用同一形状。

    ``index`` 非 None 时按指标代码取候选（V3 启动过滤 + 索引路径）；
    None 时线性扫描全部特殊规则（现行路径，行为不变）。
    """

    accu_keys: set[tuple[str, str, str]] = field(default_factory=set)
    accu_rules: dict[tuple[str, str, str], list[dict]] = field(default_factory=dict)
    special_rules: list[dict] = field(default_factory=list)
    index: dict[str, list[dict]] | None = None
    stats: dict = field(default_factory=dict)


@dataclass
class ActionContext:
    """动作执行上下文（一次流程构建一次）。"""

    current_index: dict[str, CentralRecord]
    previous_index: dict[str, CentralRecord]
    current_date: str
    previous_date: str
    exempt: set[str]
    target_unit: str
    unit: float


@dataclass
class _RowFacts:
    """当前比较行的数值快照（动作分派只读）。"""

    cur_number: float
    pre_number: float
    change: float
    ratio: float
    factor: float
    is_usd: bool


def _accu_active(current_date: str, previous_date: str) -> bool:
    """跨年或上期为 1 月 1 日时，累计指标启用（VBA 2445-2481）。"""
    try:
        cur = date.fromisoformat(current_date)
        pre = date.fromisoformat(previous_date)
    except (TypeError, ValueError):
        return False
    return cur.year != pre.year or (pre.month == 1 and pre.day == 1)


def accu_semantics(row: ComparisonRow, current_date: str, previous_date: str) -> set[str]:
    """累计指标判定的**语义集合**（生产与对拍共用的同一纯函数）。

    VBA ``当年累计指标比上期不应减少``（2445-2481）会在负数命中的同时继续走
    日期分支，两个语义可能同轮触发；消息组装（`_accu_message`）只取主文案，
    因此对拍按本函数的语义集合对齐，而不是按消息文本反推。

    语义：negative=累计值为负；mono=同年度累计不下降（本期月份>上期时同样
    要求不下降，VBA 2461-2465）；cross_year_down / cross_year_up = 跨年按月序
    双分支（VBA 2466-2477，**业务口径待确认**，V2 报告单列
    CROSS_YEAR_SEMANTICS）；1月1日结转：Legacy/VBA 跳过日期分支（无语义），
    业务口径「结转应为0」与跳过口径的差异由 Expression V2 报告单列
    JAN1_CARRYOVER_SEMANTICS。
    """
    try:
        cur = date.fromisoformat(current_date)
        pre = date.fromisoformat(previous_date)
    except (TypeError, ValueError):
        return set()
    cur_number = row.record.value if isinstance(row.record.value, (int, float)) else 0.0
    change = row.change if isinstance(row.change, (int, float)) else 0.0
    flags: set[str] = set()
    if cur_number < 0:
        flags.add("negative")
    if cur.month == 1 and cur.day == 1:
        return flags   # 结转数不判断（负数已判）
    same_year = cur.year == pre.year
    if same_year and change < 0:
        flags.add("mono")
    if not same_year:
        if cur.month >= pre.month and change < 0:
            flags.add("cross_year_down")
        if cur.month < pre.month and change >= 0:
            flags.add("cross_year_up")
    return flags


def _accu_message(row: ComparisonRow, current_date: str, previous_date: str) -> str:
    """累计指标主文案（与 accu_semantics 同一判定流；负数优先展示）。"""
    semantics = accu_semantics(row, current_date, previous_date)
    if "negative" in semantics:
        return "发生额指标不应为负数，需说明。"
    if "mono" in semantics:
        return " 当年累计指标比上期不应减少，需说明。"
    try:
        pre = date.fromisoformat(previous_date)
    except (TypeError, ValueError):
        return ""
    if "cross_year_down" in semantics:
        return f"{row.record.indicator_name}数据比{previous_date[:7]}的小，请核实。"
    if "cross_year_up" in semantics:
        return f"{row.record.indicator_name}数据 >= {previous_date[:7]}的，请核实。"
    return ""


def _as_date(value: str) -> date | None:
    try:
        return date.fromisoformat(value)
    except (TypeError, ValueError):
        return None


def _frequency_months(frequency: str) -> int:
    return {"月": 1, "季": 3, "半年": 6, "年": 12}.get(str(frequency or "").strip(), 1)


def _previous_month_end(current: date, months: int) -> date:
    month = current.month - months + 1
    year = current.year
    while month <= 0:
        month += 12
        year -= 1
    # 当月 1 日减一天即上月最后一天。
    return date(year, month, 1) - timedelta(days=1)


def _period_start(current: date, months: int) -> date:
    month = current.month - months + 1
    year = current.year
    while month <= 0:
        month += 12
        year -= 1
    return date(year, month, 1)


def _counterpart_value(
    record: CentralRecord,
    *,
    data_attr: str,
    data_index: dict[str, CentralRecord],
) -> float:
    counterpart = data_index.get(make_record_key(
        record.biz_class, record.org_code, record.region_code, record.indicator,
        data_attr, record.currency, record.frequency, record.batch,
    ))
    return counterpart.value if counterpart and isinstance(counterpart.value, (int, float)) else 0.0


def _balance_error_message(delta: float, max_value: float) -> str:
    """复刻 VBA「发生额与余额误差判断」的分档口径；输入已换算为元。"""
    if round(delta, 2) >= 0:
        return ""
    absolute = abs(delta)
    if round(absolute, 2) < 150:
        return "精度误差：上期余额+本期发生额<本期余额，偏差150元内。"
    if max_value < 10_000:
        if round(absolute, 2) < max_value:
            return "精度误差：上期余额+本期发生额<本期余额，超过150元，未超过1万元。"
        return "上期余额+本期发生额>本期余额，超过该业务量万分之一。"
    if round(absolute, 2) < 10_000:
        return "精度误差：上期余额+本期发生额<本期余额，超过150元，未超过10000元。"
    return "上期余额+本期发生额>本期余额，超过10000元。"


def _balance_and_flow_message(
    row: ComparisonRow,
    *,
    current_index: dict[str, CentralRecord],
    previous_index: dict[str, CentralRecord],
    current_date: str,
    previous_date: str,
    exempt: set[str],
    target_unit: str,
    accumulated: bool,
) -> str:
    """迁移 VBA 两类「余额与发生额」校验，保留日期连续性与跨年处理。"""
    current = _as_date(current_date)
    previous = _as_date(previous_date)
    if not current or not previous:
        return "不进行余额与发生额校验，因数据日期无法识别。"
    record = row.record
    counterpart_attr = "发生额" if record.data_attr == "余额" else "余额"
    frequency_months = _frequency_months(record.frequency)

    if accumulated:
        continuous = current.year == previous.year or (
            current.year - previous.year == 1 and previous.month == 12 and previous.day == 31
        )
        discontinuity = "年份不连续"
    else:
        continuous = previous == _previous_month_end(current, frequency_months) or (
            previous.month == 1 and previous.day == 1 and previous == _period_start(current, frequency_months)
        )
        discontinuity = "不连续"
    if not continuous:
        title = "余额与当年累计发生额" if accumulated else "余额与本期发生额"
        return f"不进行{title}校验，因{current_date}与{previous_date}{discontinuity}。"

    if counterpart_attr == "发生额":
        current_flow = _counterpart_value(record, data_attr="发生额", data_index=current_index)
        previous_flow = _counterpart_value(record, data_attr="发生额", data_index=previous_index) if accumulated else 0.0
        current_balance = record.value if isinstance(record.value, (int, float)) else 0.0
        previous_balance = row.prev_value if isinstance(row.prev_value, (int, float)) else 0.0
    else:
        current_balance = _counterpart_value(record, data_attr="余额", data_index=current_index)
        previous_balance = _counterpart_value(record, data_attr="余额", data_index=previous_index)
        current_flow = record.value if isinstance(record.value, (int, float)) else 0.0
        previous_flow = row.prev_value if accumulated and isinstance(row.prev_value, (int, float)) else 0.0

    if accumulated and ((previous.month == 12 and previous.day == 31) or (previous.month == 1 and previous.day == 1)):
        previous_flow = 0.0
    to_yuan = 1.0 if record.indicator in exempt else get_unit_factor(target_unit)
    values = [current_balance * to_yuan, previous_balance * to_yuan, current_flow * to_yuan, previous_flow * to_yuan]
    current_balance, previous_balance, current_flow, previous_flow = values
    maximum = max(values)
    delta = previous_balance + current_flow - previous_flow - current_balance if accumulated else previous_balance + current_flow - current_balance
    if round(delta, 2) >= 0:
        message = ""
    elif accumulated and round(current_flow - previous_flow, 10) == 0:
        message = "本期无发生额，余额增加了，请核实"
    else:
        message = _balance_error_message(delta, maximum)
    if not accumulated and current_flow < 0:
        message = "发生额指标不应为负数。" + message
    return message


def _divisible_message(
    row: ComparisonRow,
    divisor: float,
    *,
    detail: str = "",
    exempt: set[str],
    target_unit: str,
) -> str:
    divisor_as_integer = int(divisor)
    if divisor_as_integer == 0:
        return "除数为0，特殊指标表设置有误。"
    current = row.record.value if isinstance(row.record.value, (int, float)) else 0.0
    to_yuan = 1.0 if row.record.indicator in exempt else get_unit_factor(target_unit)
    quotient = abs(current) * to_yuan / divisor_as_integer
    if abs(quotient - round(quotient)) > 1e-10:
        return f"指标应能被{divisor_as_integer}整除。{detail}"
    return ""


def build_prepared_from_config(config: CentralConfig) -> PreparedActionRules:
    """现行路径：从旧三表（累计/特殊）装载，不过滤、不索引（行为与历史版本一致）。"""
    prepared = PreparedActionRules()
    for row in config.rules:
        if _is_disabled(row.get("禁用")):
            continue
        kind = row.get("类型")
        if kind == "累计不应下降":
            code = str(row.get("指标代码") or "").replace("'", "").strip()
            if code:
                key = (code, str(row.get("数据属性") or "").strip(),
                       str(row.get("频度") or "").strip())
                prepared.accu_keys.add(key)
                prepared.accu_rules.setdefault(key, []).append(row)
        elif kind == "特殊阈值":
            codes = _split_multi_codes(row.get("指标代码"))
            if codes:
                entry = dict(row)
                entry["__codes__"] = codes
                prepared.special_rules.append(entry)
    return prepared


def apply_special_action(rule: dict, row: ComparisonRow, facts: _RowFacts, ctx: ActionContext) -> str:
    """特殊指标动作分派（16 种）；现行路径与 V3 路径共用，保证同文同判。"""
    text = str(rule.get("备注") or "").strip()
    if not supports_special_action(text):
        return ""
    rmb = _to_float(rule.get("人民币阀值(亿元)")) or 0.0
    usd = _to_float(rule.get("美元阀值(亿美元)")) or 0.0
    amplitude = _to_float(rule.get("绝对值变幅(%)")) or 0.0
    threshold = usd if facts.is_usd and usd else rmb
    cur_number = facts.cur_number
    pre_number = facts.pre_number
    change = facts.change
    ratio = facts.ratio
    factor = facts.factor
    record = row.record

    if text == "指标不应有数，需说明。":
        if record.value is not None and str(record.value) != "":
            return text
        return ""
    if text == "指标有变动需说明。":
        # 这里的“金额”指两期变动额，而非本期余额。两期相同但余额非零
        # （例如实收资本）不属于“有变动”。
        changed_amount = abs(change) * factor
        if round(amplitude, 6) != 0:
            # 有金额阈值时，金额和变幅必须同时达到；金额阈值留空时，
            # 仍需确有变动且变幅达到。
            amount_threshold = usd if facts.is_usd and usd else rmb
            amount_hit = (
                change != 0
                if amount_threshold == 0
                else changed_amount >= amount_threshold
            )
            return text if amount_hit and abs(ratio) > amplitude else ""
        if rmb == 0 and usd == 0:
            hit = change != 0
        else:
            hit = change != 0 and changed_amount >= threshold
        return text if hit else ""
    if text == "指标变动或变幅超过阈值需说明。":
        amount_threshold = usd if facts.is_usd and usd else rmb
        value_hit = amount_threshold > 0 and abs(cur_number) * factor >= amount_threshold
        range_hit = amplitude > 0 and abs(ratio) > amplitude
        return text if value_hit or range_hit else ""
    if text == "指标变幅异常需说明。":
        # 这是纯变幅动作，不得再借用“人民币阈值(亿元)”存比例。
        if amplitude > 0 and abs(ratio) >= amplitude:
            return f"{ratio * 100:.2f}%，{text}"
        return ""
    if text in ("当年累计指标比上期不应减少。", "历史累计指标比上期不应减少。"):
        if text.startswith("当年"):
            if _accu_active(ctx.current_date, ctx.previous_date):
                return _accu_message(row, ctx.current_date, ctx.previous_date)
            if change < 0:
                return text
        else:
            if cur_number < 0 or change < 0:
                return text
        return ""
    if text in ("指标比上期增加需说明。", "指标比上期减少需说明。"):
        scaled = change * factor
        if (((text.endswith("增加需说明。") and scaled > 0) or
                (text.endswith("减少需说明。") and scaled < 0)) and abs(scaled) > threshold):
            return text
        return ""
    if text in ("指标新增需说明。", "指标结清需说明。"):
        if text.startswith("指标新增"):
            hit = cur_number > 0 and pre_number == 0 and abs(change) * factor >= abs(rmb)
        else:
            hit = cur_number == 0 and pre_number > 0 and abs(change) * factor >= abs(rmb)
        return text if hit else ""
    if text == "指标为负数需核实。":
        return text if cur_number < 0 else ""
    if text == "指标应为整数。":
        # VBA 用 CStr(Abs(curD)) 与 CStr(CLng(Abs(curD))) 比较；
        # Python 的 str(1) / str(1.0) 文本不同，不能照搬文本比较。
        return "" if float(abs(cur_number)).is_integer() else text
    if text == "指标应能被某数整除。":
        return _divisible_message(
            row, rmb, detail=str(rule.get("详细说明") or "").strip(),
            exempt=ctx.exempt, target_unit=ctx.target_unit,
        )
    if text == "指标本期余额应小于（上期余额+本期发生额）。":
        return _balance_and_flow_message(
            row, current_index=ctx.current_index, previous_index=ctx.previous_index,
            current_date=ctx.current_date, previous_date=ctx.previous_date,
            exempt=ctx.exempt, target_unit=ctx.target_unit, accumulated=False,
        )
    if text == "指标本期余额应小于（上期余额+本期当年发生额-上期当年累计发生额）。":
        return _balance_and_flow_message(
            row, current_index=ctx.current_index, previous_index=ctx.previous_index,
            current_date=ctx.current_date, previous_date=ctx.previous_date,
            exempt=ctx.exempt, target_unit=ctx.target_unit, accumulated=True,
        )
    if text == "指标需要对应存在。":
        opponent = _opponent_value(row, rule, ctx.current_index, action=text)
        if cur_number != 0 and opponent == 0:
            return f"{record.indicator_name}有数，对应指标无数，需说明。"
        if cur_number == 0 and opponent != 0:
            return f"{record.indicator_name}无数，对应指标有数，需说明。"
        return ""
    if text == "指标应相等核查。":
        target = str(rule.get("说明/比较指标") or "").replace("'", "").strip()
        opponent = _opponent_value(row, rule, ctx.current_index, action=text)
        # VBA 直接在比较结果的统一输出单位上比较，未再进行单位换算。
        # 仅用极小误差屏蔽浮点尾差，避免把 0.1 + 0.2 的表示误差当作业务异常。
        difference = cur_number - opponent
        if target and abs(difference) > 2.220446049250313e-13:
            detail = str(rule.get("详细说明") or "").strip()
            return f"{record.indicator_name}不等于{target}偏差{difference:.4f},请核实。{detail}"
        return ""
    if "相除" in text or "太小核查" in text:
        target = str(rule.get("说明/比较指标") or "").replace("'", "").strip()
        opponent = _opponent_value(row, rule, ctx.current_index, action=text)
        opponent_factor = 1.0 if target in ctx.exempt else ctx.unit
        if opponent == 0:
            if cur_number != 0:
                return f"{record.indicator_name}有数，比较指标{target}为空或为0，需核实。"
            return ""
        quotient = abs(cur_number * factor / (opponent * opponent_factor))
        limit = usd if (facts.is_usd and usd) else rmb
        if ("太小" in text and quotient < limit) or ("太小" not in text and quotient > limit):
            if "太小" in text:
                return f"{record.indicator_name}={quotient:.4f},小于{limit}，需说明。"
            return f"{record.indicator_name}={quotient:.4f},超过{limit}，需说明。"
        return ""
    if "相减" in text:
        target = str(rule.get("说明/比较指标") or "").replace("'", "").strip()
        opponent = _opponent_value(row, rule, ctx.current_index, action=text)
        opponent_factor = 1.0 if target in ctx.exempt else ctx.unit
        limit = usd if (facts.is_usd and usd) else rmb
        if cur_number * factor - opponent * opponent_factor < abs(limit):
            return text
        return ""
    return ""


def _rule_applies_to_record(rule: dict, record: CentralRecord, current_date: str) -> bool:
    """规则行的行级过滤（频度/批次/属性/币种/结转场景），现行与 V3 共用。"""
    rule_frequency = str(rule.get("频度") or "").strip()
    rule_batch = str(rule.get("批次") or "").strip()
    scene = str(rule.get("场景") or "").strip()
    rule_attr = str(rule.get("数据属性") or "").strip()
    rule_currency = str(rule.get("币种") or "").strip()
    if rule_frequency and rule_frequency != record.frequency:
        return False
    if rule_batch and rule_batch != record.batch:
        return False
    # VBA dataColor：数据属性空/0 为通用，否则须包含行属性（InStr）。
    if rule_attr and rule_attr != "0" and record.data_attr not in rule_attr:
        return False
    if rule_currency and rule_currency != record.currency:
        return False
    if scene == "结转" and not current_date.endswith("-01-01"):
        return False
    return True


def enrich_indicator_rules(
    rows: list[ComparisonRow],
    config: CentralConfig,
    *,
    current_index: dict[str, CentralRecord],
    previous_index: dict[str, CentralRecord] | None = None,
    current_date: str,
    previous_date: str,
    target_unit: str = "亿元",
    exempt: set[str] | None = None,
    soft_color: int = 46,
    prepared: PreparedActionRules | None = None,
) -> dict:
    """执行累计指标与特殊阈值规则；原地写 是否说明/填充。

    ``prepared`` 缺省时从旧三表装载（现行路径）；传入 rule_action_engine
    构造的已过滤/已索引规则集即为 V3 路径，动作语义与现行路径完全同源。
    返回统计（用于运行日志）。
    """
    exempt = exempt or set()
    previous_index = previous_index or {}
    if prepared is None:
        prepared = build_prepared_from_config(config)
    if not prepared.accu_keys and not prepared.special_rules:
        return {"loaded": 0, "used": 0}
    unit = get_unit_factor(target_unit) / get_unit_factor(SPECIAL_UNIT)
    ctx = ActionContext(
        current_index=current_index, previous_index=previous_index,
        current_date=current_date, previous_date=previous_date,
        exempt=exempt, target_unit=target_unit, unit=unit,
    )
    used_rules = 0
    # “对应存在”规则的主指标可能根本没有报送行。为这种情况建立反向索引，
    # 让已存在的比较指标承接提示；不制造缺失指标的虚假数据行。
    reverse_correspondence: dict[str, list[dict]] = {}
    for rule in prepared.special_rules:
        if str(rule.get("备注") or "").strip() != "指标需要对应存在。":
            continue
        target = str(rule.get("说明/比较指标") or "").replace("'", "").strip()
        if target:
            reverse_correspondence.setdefault(target, []).append(rule)

    for row in rows:
        record = row.record
        messages: list[str] = []
        cur_number = record.value if isinstance(record.value, (int, float)) else 0.0
        pre_number = row.prev_value if isinstance(row.prev_value, (int, float)) else 0.0
        change = row.change if isinstance(row.change, (int, float)) else 0.0
        ratio = row.ratio if isinstance(row.ratio, (int, float)) else 0.0
        factor = 1.0 if record.indicator in exempt else unit
        facts = _RowFacts(
            cur_number=cur_number, pre_number=pre_number, change=change,
            ratio=ratio, factor=factor, is_usd=record.currency == "美元合计",
        )

        # ---- 累计指标 ----
        accu_key = (record.indicator, record.data_attr, record.frequency)
        if accu_key in prepared.accu_keys:
            if _accu_active(current_date, previous_date):
                message = _accu_message(row, current_date, previous_date)
            elif change < 0:
                message = " 当年累计指标比上期不应减少，需说明。"
            else:
                message = ""
            if message:
                rules = prepared.accu_rules.get(accu_key, ())
                if rules:
                    messages.extend(_special_output_message(rule, message, facts) for rule in rules)
                else:
                    messages.append(message)

        # ---- 特殊阈值 ----
        if prepared.index is not None:
            candidates = prepared.index.get(record.indicator, ())
        else:
            candidates = prepared.special_rules
        for rule in candidates:
            if record.indicator not in rule["__codes__"]:
                continue
            if not _rule_applies_to_record(rule, record, current_date):
                continue
            used_rules += 1
            hit_message = apply_special_action(rule, row, facts, ctx)
            if hit_message:
                messages.append(_special_output_message(rule, hit_message, facts))

        if isinstance(record.value, (int, float)) and record.value != 0:
            for rule in reverse_correspondence.get(record.indicator, ()):
                if not _rule_applies_to_record(rule, record, current_date):
                    continue
                for source in rule["__codes__"]:
                    if source == record.indicator:
                        continue
                    source_key = make_record_key(
                        record.biz_class, record.org_code, record.region_code,
                        source, record.data_attr, record.currency,
                        record.frequency, record.batch,
                    )
                    if source_key in current_index:
                        continue
                    used_rules += 1
                    hit_message = f"{record.indicator}有数,{source}无数,请核实。"
                    messages.append(_special_output_message(rule, hit_message, facts))

        if not messages:
            continue
        joined = JOIN.join(messages)
        row.need_explain = joined if not row.need_explain else row.need_explain + JOIN + joined
        soft = row.fill_color is None or row.fill_scope != "row"
        if all(("软性" in message) for message in messages):
            row.fill_color = soft_color
            row.fill_scope = "remark"
        _ = soft
    return {
        "loaded": len(prepared.accu_keys) + len(prepared.special_rules),
        "used": used_rules,
        "accu": len(prepared.accu_keys),
        "special": len(prepared.special_rules),
    }


def _opponent_value(
    row: ComparisonRow,
    rule: dict,
    data_index: dict[str, CentralRecord],
    *,
    action: str,
) -> float:
    """按 VBA ``dCurDataRef`` 的八维键读取本期比较指标；缺失按 0 处理。"""
    target = str(rule.get("说明/比较指标") or "").replace("'", "").strip()
    if not target:
        return 0.0
    currency = row.record.currency
    usd = _to_float(rule.get("美元阀值(亿美元)")) or 0.0
    # VBA 特殊协议：对应存在规则在配置美元阈值时、相除规则在美元阈值为
    # 999 时，用“人民币”指标作为“美元合计”行的比较对象。
    if currency == "美元合计":
        if (action == "指标需要对应存在。" and usd != 0) or (
            ("相除" in action or "太小核查" in action) and usd == 999
        ):
            currency = "人民币"
    record = data_index.get(
        make_record_key(
            row.record.biz_class, row.record.org_code, row.record.region_code,
            target, row.record.data_attr, currency, row.record.frequency, row.record.batch,
        )
    )
    if record is None:
        return 0.0
    return record.value if isinstance(record.value, (int, float)) else 0.0
