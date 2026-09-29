from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Iterable

from .models import Issue


HISTORY_HEADERS = (
    "工作簿名",
    "工作表名",
    "定位单元格",
    "错误类型",
    "校验指标",
    "描述",
    "当前值",
    "对比值",
    "差值",
    "规则编号",
    "历史校验说明",
    "审核意见",
)

# 逐笔统计系统_配置.xlsx 的业务结果历史 sheet 名。追加规则在代码中内置
# （HISTORY_APPEND_RULES），不再使用可编辑的“历史追加配置”表。
HISTORY_AUDIT_SHEET = "核查表校验结果"
LOCAL_VALIDATION_HISTORY_SHEET = "本地校验结果"
# 本地校验结果历史 sheet 上要回写到汇总输出的三列。
LOCAL_VALIDATION_HISTORY_FIELDS = ("历史触发条数", "历史分类说明", "审核结果")


def _period_key(value: str) -> tuple[int, int, str]:
    text = value.strip()
    for fmt in ("%Y-%m", "%Y/%m", "%Y%m", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.year, dt.month, text
        except ValueError:
            pass
    if len(text) >= 6 and text[:4].isdigit() and "Q" in text.upper():
        try:
            quarter = int(text.upper().split("Q", 1)[1][:1])
            return int(text[:4]), quarter * 3, text
        except (ValueError, IndexError):
            pass
    return 0, 0, text


def classify_current_issues(
    current_period: str,
    current: list[Issue],
    history: Iterable[Issue],
    *,
    audited_org_codes: set[str] | None = None,
    batch_id: str = "",
    audit_time: str = "",
) -> tuple[list[Issue], list[Issue]]:
    old_rows = [row for row in history if row.period != current_period]
    old_triggered = [row for row in old_rows if row.triggered]
    periods = sorted({row.period for row in old_rows}, key=_period_key)
    previous_period = periods[-1] if periods else ""

    by_id: dict[str, list[Issue]] = defaultdict(list)
    for row in old_triggered:
        by_id[row.issue_id].append(row)
    for values in by_id.values():
        values.sort(key=lambda item: _period_key(item.period))

    current_ids = {item.issue_id for item in current}
    classified: list[Issue] = []
    for item in current:
        prior = by_id.get(item.issue_id, [])
        if not prior:
            item.status = "新增"
            item.first_seen_period = current_period
            item.previous_seen_period = ""
            item.consecutive_count = 1
        else:
            last = prior[-1]
            item.first_seen_period = prior[0].first_seen_period or prior[0].period
            item.previous_seen_period = last.period
            if previous_period and last.period == previous_period:
                item.status = "连续出现"
                item.consecutive_count = max(1, last.consecutive_count) + 1
            else:
                item.status = "再次出现"
                item.consecutive_count = 1
            item.institution_feedback = last.institution_feedback
            item.auditor_opinion = last.auditor_opinion
        classified.append(item)

    resolved: list[Issue] = []
    if previous_period:
        latest_by_id: dict[str, Issue] = {}
        for row in old_triggered:
            if row.period == previous_period:
                latest_by_id[row.issue_id] = row
        for issue_id, old in latest_by_id.items():
            in_scope = audited_org_codes is None or old.org_code in audited_org_codes
            if in_scope and issue_id not in current_ids:
                resolved.append(
                    old.clone(
                        period=current_period,
                        batch_id=batch_id or (current[0].batch_id if current else ""),
                        audit_time=audit_time or (current[0].audit_time if current else ""),
                        triggered=False,
                        status="已整改",
                        previous_seen_period=old.period,
                        consecutive_count=0,
                        target_value="",
                        comparison_value="",
                        reference_value="",
                        difference_value="",
                        formula_result="",
                    )
                )
    return classified, resolved


def merge_history(existing: Iterable[Issue], additions: Iterable[Issue]) -> list[Issue]:
    """Carry all manually written notes for the same readable rule number."""
    grouped: dict[str, list[Issue]] = defaultdict(list)
    for item in existing:
        grouped[item.issue_id].append(item)

    def combine(rows: Iterable[Issue], field: str) -> str:
        values: list[str] = []
        for row in rows:
            value = str(getattr(row, field, "") or "").strip()
            if value and value not in values:
                values.append(value)
        return "\n".join(values)

    for item in additions:
        previous = grouped.get(item.issue_id, [])
        if previous:
            item.institution_feedback = combine(previous, "institution_feedback")
            item.auditor_opinion = combine(previous, "auditor_opinion")
    merged: dict[str, Issue] = {item.issue_id: item for item in existing}
    for item in additions:
        merged[item.issue_id] = item
    return sorted(
        merged.values(),
        key=lambda item: (Path(item.source_file).name, item.sheet_name, item.formula_cell, item.issue_id),
    )


# 本地校验结果「历史追加」用的复合主键。规则编号只在单机构单表内唯一，
# 汇总跨机构跨表，故需这 8 个字段一起才能唯一确定一行；它们跨期稳定。
LOCAL_VALIDATION_IDENTITY_FIELDS = (
    "来源文件",
    "来源工作表",
    "批次",
    "表单名称",
    "规则编号",
    "规则类型",
    "规则描述",
    "校验字段",
)


def _split_columns(value: object) -> tuple[str, ...]:
    return tuple(
        item.strip() for item in str(value or "").replace("，", ",").replace("、", ",").split(",")
        if item.strip()
    )


# 追加规则（代码内置）：结果历史表 → (去重列, 追加的历史列)。
# 追加列为空元组时由读取方自动取“历史表中非去重列、且汇总表没有的列”。
HISTORY_APPEND_RULES = {
    LOCAL_VALIDATION_HISTORY_SHEET: (
        LOCAL_VALIDATION_IDENTITY_FIELDS, LOCAL_VALIDATION_HISTORY_FIELDS,
    ),
    "业务说明": (
        ("来源文件", "来源工作表", "业务表单名称", "业务表单英文名称"),
        ("历史说明",),
    ),
    "任务说明": (
        ("来源文件", "来源工作表", "社会统一信用代码", "金融机构编码", "任务名称"),
        ("历史说明",),
    ),
}

# 内置追加历史表的模板表头：新建空工作簿时按此预置空表，使用者照列填写即可；
# 表头一经存在不再改写（保留用户自定义列）。
HISTORY_TEMPLATE_HEADERS = {
    "业务说明": (
        "来源文件", "来源工作表", "机构名称", "统一社会信用代码", "报送报表",
        "数据日期", "序号", "业务表单名称", "业务表单英文名称", "说明内容", "历史说明",
    ),
    "任务说明": (
        "来源文件", "来源工作表", "机构名称", "统一社会信用代码", "报送报表",
        "数据日期", "序号", "社会统一信用代码", "金融机构编码", "任务名称",
        "说明内容", "历史说明",
    ),
}


def history_append_rule(result_sheet: str) -> tuple[tuple[str, ...], tuple[str, ...]]:
    """内置追加规则：返回 (去重列, 追加的历史列)；未登记的表返回空。"""
    keys, fields = HISTORY_APPEND_RULES.get(result_sheet, ((), ()))
    return keys, fields


def _norm_key_value(value: object) -> str:
    """复合键取值归一化：数字 1 / 1.0 / '1.0' / ' 1 ' 视为同一个键。

    汇总侧经 COM ``Value2`` 读到的整数是 float（1.0 → '1.0'），而历史表
    常为人工粘贴的文本 '1'；不归一化会导致明明相同的行配不上。
    """
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    try:
        number = float(text)
    except ValueError:
        return text
    if number.is_integer():
        return str(int(number))
    return repr(number)


def _composite_key(headers: list[str], row: list, key_fields: tuple[str, ...]) -> tuple[str, ...]:
    index = {name: i for i, name in enumerate(headers)}
    return tuple(
        _norm_key_value(row[index[name]])
        if name in index and index[name] < len(row)
        else ""
        for name in key_fields
    )


def join_local_validation_history(
    headers: list[str],
    history_rows: list[list],
    *,
    key_fields: tuple[str, ...] = LOCAL_VALIDATION_IDENTITY_FIELDS,
    history_fields: tuple[str, ...] = LOCAL_VALIDATION_HISTORY_FIELDS,
) -> dict[tuple[str, ...], list[str]]:
    """把历史行按复合键聚合，返回 {复合键: [每个 history_field 用 | 拼接后的值]}。

    同一复合键（同一规则编号组合）出现多行时，把每个历史字段的值按出现顺序用
    ``|`` 拼接（如 ``10|20``、``说明1|说明2``、``通过|不通过``），供汇总输出把
    历史信息回写进本地校验结果。
    """
    index = {name: i for i, name in enumerate(headers)}
    grouped: dict[tuple[str, ...], list[list]] = defaultdict(list)
    for row in history_rows:
        grouped[_composite_key(headers, row, key_fields)].append(row)
    result: dict[tuple[str, ...], list[str]] = {}
    for key, rows in grouped.items():
        joined: list[str] = []
        for field in history_fields:
            i = index.get(field)
            values: list[str] = []
            for row in rows:
                if i is not None and i < len(row) and row[i] not in (None, ""):
                    value = str(row[i]).strip()
                    if value:
                        values.append(value)
            joined.append("|".join(values))
        result[key] = joined
    return result
