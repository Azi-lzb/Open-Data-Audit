"""大集中统计系统：V3 规则动作引擎（启动过滤 + 指标索引）。

读「规则动作」表（V3 合并计划 §三 表 1），在流程开始前整批剔除不适用规则
（VBA ``B导数比较.bas`` 1121-1143 的装载过滤语义）并按指标代码建索引，
替代逐行全量扫描；动作执行复用 ``indicator_rule_engine`` 的共享分派函数，
判定义与文案与现行路径完全同源（构造保证，不翻译）。

启动过滤口径（VBA 对齐）：
- 场景=结转的规则仅在本期为 1 月 1 日时保留；非结转规则在 1 月 1 日整批剔除；
- 「当年累计」来源分组默认不受 1 月 1 日门控；若显式标记适用场景=结转，
  仍只在本期 1 月 1 日执行；
- 规则带适用频度时，频度(+批次) 必须出现在两期数据的 fres 集合中。
"""

from __future__ import annotations

from .config import CentralConfig
from .indicator_rule_engine import (
    PreparedActionRules,
    _is_disabled,
    _split_multi_codes,
)

ACCUM_SOURCE_GROUP = "当年累计"


def _is_jan1(current_date: str) -> bool:
    return str(current_date or "").endswith("-01-01")


def filter_applicable(rules: list[dict], *, current_date: str, fres: set[str] | None) -> list[dict]:
    """V3 启动过滤：1 月 1 日门控 + 频度(批次) 存在性；返回保留的规则。"""
    jan1 = _is_jan1(current_date)
    fres = fres or set()
    kept: list[dict] = []
    for rule in rules:
        scene = str(rule.get("场景") or "").strip()
        if scene == "结转" and not jan1:
            continue
        if rule.get("类型") == "特殊阈值":
            if scene != "结转" and jan1:
                continue
        frequency = str(rule.get("频度") or "").strip()
        if frequency and fres:
            batch = str(rule.get("批次") or "").strip()
            key = frequency + batch if batch else frequency
            if key not in fres:
                continue
        kept.append(rule)
    return kept


def build_index(rules: list[dict]) -> dict[str, list[dict]]:
    """指标代码 → 规则列表（多代码规则按 ``/`` 拆分登记）。"""
    index: dict[str, list[dict]] = {}
    for rule in rules:
        for code in rule["__codes__"]:
            index.setdefault(code, []).append(rule)
    return index


def prepare_action_rules(
    config: CentralConfig,
    *,
    current_date: str,
    fres: set[str] | None = None,
    use_filter: bool = True,
) -> PreparedActionRules:
    """装载「规则动作」表 → 启动过滤 → 建索引；供 enrich_indicator_rules 消费。

    配置无「规则动作」表时返回空集（调用方应回退现行路径）。
    """
    prepared = PreparedActionRules()
    if not config.action_rules:
        return prepared
    total = 0
    filtered_out = 0
    for row in config.action_rules:
        if _is_disabled(row.get("禁用")):
            continue
        total += 1
        if row.get("类型") == "累计不应下降":
            # 累计来源组通常不受 1 月 1 日特殊阈值门控；但若该行明确
            # 配置“适用场景=结转”，仍必须只在本期 1 月 1 日执行。
            if use_filter and not filter_applicable([row], current_date=current_date, fres=None):
                filtered_out += 1
                continue
            code = str(row.get("指标代码") or "").replace("'", "").strip()
            if code:
                key = (code, str(row.get("数据属性") or "").strip(),
                       str(row.get("频度") or "").strip())
                prepared.accu_keys.add(key)
                prepared.accu_rules.setdefault(key, []).append(row)
        elif row.get("类型") == "特殊阈值":
            codes = _split_multi_codes(row.get("指标代码"))
            if codes:
                entry = dict(row)
                entry["__codes__"] = codes
                prepared.special_rules.append(entry)
    if use_filter:
        kept = filter_applicable(
            prepared.special_rules, current_date=current_date, fres=fres,
        )
        filtered_out += len(prepared.special_rules) - len(kept)
        prepared.special_rules = kept
    prepared.index = build_index(prepared.special_rules)
    prepared.stats = {
        "source": "规则动作",
        "loaded": total,
        "accu": len(prepared.accu_keys),
        "special": len(prepared.special_rules),
        "filtered_out": filtered_out,
    }
    return prepared
