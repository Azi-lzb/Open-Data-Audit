"""大集中统计系统：跨期/数值核对引擎。

对应 VBA ``D跨期数值核对.bas``：数据按 机构类代码+地区代码 组织，规则含
前提条件/left/right/校验公式；[...] 引用本期、{...} 引用上期（5 段：
指标,数据属性,币种,频度,批次）。表达式经受限解析器求值，0 值以 0.01 元
哨兵参与计算；错误写规则错误日志，绝不静默按 0 处理。
"""

from __future__ import annotations

import re
from pathlib import Path

from .comparison_engine import get_unit_factor
from .config import CentralConfig
from .csv_importer import read_central_csv
from .expression_parser import ExpressionError, evaluate_expression
from .models import CentralDataset

_TOKEN_RE = re.compile(r"\[[^\[\]]*\]|\{[^{}]*\}")

DEFAULT_TITLE = (
    "数据日期", "机构类代码", "机构类名称", "地区代码", "地区名称",
    "校验编码", "校验名称", "校验类型", "备注",
    "左值", "右值", "差异", "差异绝对值", "差异幅度(%)",
    "是否说明", "说明内容", "计算公式",
)


def _uni_code(*segments: str) -> str:
    return "".join(segments)


def _is_disabled(value: str) -> bool:
    return str(value or "").strip() in {"是", "1", "true", "True", "Y", "y"}


def _build_data_store(
    current: CentralDataset, previous: CentralDataset | None, *, target_unit: str
) -> tuple[dict, set[str], dict[str, str], dict[str, str]]:
    """VBA InitCurData/InitPreData：d_data[机构4+地区7][uni]["cur"/"pre"]。"""
    store: dict[str, dict] = {}
    fres: set[str] = set()
    org_names: dict[str, str] = {}
    region_names: dict[str, str] = {}
    sentinel = 0.01 * get_unit_factor("元") / get_unit_factor(target_unit)

    def number_of(value) -> float:
        if isinstance(value, (int, float)):
            number = float(value)
        else:
            number = 0.0
        return sentinel if number == 0 else number

    def feed(dataset: CentralDataset, side: str) -> None:
        for record in dataset.records:
            org_key = record.org_code + record.region_code
            fres.add(record.frequency + record.batch)
            org_names.setdefault(record.org_code, record.org_name)
            region_names.setdefault(record.region_code, record.region_name)
            bucket = store.setdefault(org_key, {})
            uni = _uni_code(
                record.indicator, record.data_attr, record.currency,
                record.frequency, record.batch,
            )
            entry = bucket.setdefault(uni, {})
            entry[side] = number_of(record.value)

    feed(current, "cur")
    if previous is not None:
        feed(previous, "pre")
    return store, fres, org_names, region_names


def _token_value(token_body: str, store_entry: dict, side: str) -> str:
    segments = [part.strip() for part in token_body.split(",")]
    if len(segments) == 5:
        uni = _uni_code(*segments)
    elif len(segments) == 1:
        uni = segments[0]
    else:
        return "0"
    value = store_entry.get(uni, {}).get(side)
    return repr(float(value)) if value is not None else "0"


def _substitute_tokens(expression: str, store_entry: dict) -> str:
    """VBA calu_rule：先 {...}（上期）再 [...]（本期），缺失取 0。"""
    result = re.compile(r"\{[^{}]*\}").sub(
        lambda m: _token_value(m.group(0)[1:-1], store_entry, "pre"), expression
    )
    result = re.compile(r"\[[^\[\]]*\]").sub(
        lambda m: _token_value(m.group(0)[1:-1], store_entry, "cur"), result
    )
    return result


def _rule_frequency_ok(expression: str, fres: set[str]) -> bool:
    """VBA CheckFres：每个 token 须 5 段且 频度&批次 已出现在数据中。"""
    for match in _TOKEN_RE.finditer(expression):
        segments = [part.strip() for part in match.group(0)[1:-1].split(",")]
        if len(segments) not in (1, 5):
            return False
        if len(segments) == 5 and (segments[3] + segments[4]) not in fres:
            return False
    return True


def _safe_regex_match(pattern: str, text: str) -> bool:
    if not pattern:
        return True
    try:
        return re.search(pattern, text) is not None
    except re.error:
        return False


def run_cross_period(
    *,
    current: CentralDataset,
    previous: CentralDataset | None,
    config: CentralConfig,
    target_unit: str = "亿元",
    source_unit: str = "元",
) -> tuple[list[dict], list[tuple]]:
    """求值全部跨期核对规则；返回 (结果行, 日志行)。"""
    store, fres, org_names, region_names = _build_data_store(
        current, previous, target_unit=target_unit
    )
    log_rows: list[tuple] = []
    results: list[dict] = []
    for row in config.cross_rules:
        if _is_disabled(row.get("禁用")):
            continue
        rule_id = row.get("校验编码") or row.get("校验名称") or ""
        left_expr = str(row.get("left") or "").strip()
        right_expr = str(row.get("right") or "").strip()
        formula = str(row.get("校验公式") or "").strip()
        if ("{" in left_expr or "{" in right_expr or "{" in (row.get("前提条件") or "")) and previous is None:
            continue
        if not left_expr or not right_expr:
            continue
        if not _rule_frequency_ok(left_expr + right_expr + (row.get("前提条件") or ""), fres):
            continue

        variables = {"Thd": 1.0}
        for org_key, store_entry in store.items():
            org_code = org_key[:4]
            region_code = org_key[-7:]
            if not _safe_regex_match(row.get("机构类代码", ""), org_code):
                continue
            if not _safe_regex_match(row.get("地区代码", ""), region_code):
                continue
            context = {"org_key": org_key, "store": store_entry}

            def calc(expression: str, mode: str) -> "float | None":
                substituted = _substitute_tokens(expression, store_entry)
                try:
                    return float(evaluate_expression(substituted, variables))
                except ExpressionError as exc:
                    log_rows.append((org_code, region_code, f"{mode}计算有误", rule_id, f"{expression} -> {exc}"))
                    return None

            premise = str(row.get("前提条件") or "").strip()
            if premise:
                premise_value = calc(premise, "前提条件")
                if premise_value is None:
                    continue
                if premise_value == 0:
                    continue
            left_value = calc(left_expr, "left_data")
            if left_value is None:
                continue
            right_value = calc(right_expr, "right_data")
            if right_value is None:
                continue

            formula_error = ""
            val_result = None
            if formula:
                try:
                    val_result = evaluate_expression(formula, {"left": left_value, "right": right_value, "Thd": 1.0})
                except ExpressionError as exc:
                    formula_error = f"校验公式计算有误：{exc}"
                    log_rows.append((org_code, region_code, "校验公式 计算有误", rule_id, formula))

            difference = left_value - right_value
            abs_difference = abs(difference)
            if left_value == 0 and right_value == 0:
                ratio = None
            else:
                ratio = abs_difference / max(abs(left_value), abs(right_value)) * 100

            inverted = bool(str(row.get("取反标识") or "").strip())
            suppress_ok = bool(str(row.get("无误是否提示") or "").strip())
            if formula and not formula_error and suppress_ok:
                truth = bool(val_result)
                if (inverted and truth) or (not inverted and not truth):
                    continue

            # 输出门槛（VBA 302-310）：left=0 且 right=0 → 整条不输出。
            if left_value == 0 and right_value == 0:
                continue

            check_type = "逻辑校验" if str(row.get("是否逻辑校验") or "").strip() not in {"", "0"} else "数值核对"
            messages: list[str] = []
            if check_type == "数值核对" and (left_value == 0) != (right_value == 0):
                messages.append("一边有，一边无;")
            if formula and not formula_error:
                truth = bool(val_result)
                hit = truth if not inverted else not truth
                if hit:
                    hint = str(row.get("校验提示") or "").strip()
                    if hint:
                        messages.append(hint)
            if formula_error:
                messages.append(formula_error)

            process = (
                f"left:{left_expr};{_substitute_tokens(left_expr, store_entry)};"
                f"right:{right_expr};{_substitute_tokens(right_expr, store_entry)};"
                f"校验公式:{formula};"
            )
            results.append({
                "数据日期": current.record_date,
                "机构类代码": org_code,
                "机构类名称": org_names.get(org_code, ""),
                "地区代码": region_code,
                "地区名称": region_names.get(region_code, ""),
                "校验编码": row.get("校验编码", ""),
                "校验名称": row.get("校验名称", ""),
                "校验类型": check_type,
                # 规则表“备注”说明比对口径（如 日月报核对/月报12批核对），逐行带出。
                "备注": str(row.get("备注") or "").strip(),
                "左值": left_value,
                "右值": right_value,
                "差异": difference,
                "差异绝对值": abs_difference,
                "差异幅度(%)": ratio,
                "是否说明": "".join(messages),
                "说明内容": "",
                "计算公式": process,
            })
    return results, log_rows


def _substitute_one(match: re.Match, store_entry: dict, side: str) -> str:
    token_body = match.group(0)[1:-1]
    segments = [part.strip() for part in token_body.split(",")]
    if len(segments) == 5:
        uni = _uni_code(*segments)
    elif len(segments) == 1:
        uni = segments[0]
    else:
        return "0"
    value = store_entry.get(uni, {}).get(side)
    return repr(float(value)) if value is not None else "0"
