"""V1/V2 并行运行后的只读差异报告。"""

from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import TYPE_CHECKING

from openpyxl import Workbook, load_workbook

if TYPE_CHECKING:
    from .service import V2RunResult


def _v1_rows(path: Path, sheet_name: str) -> list[dict[str, object]]:
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in book.sheetnames:
            return []
        iterator = book[sheet_name].iter_rows(values_only=True)
        headers = [str(value or "") for value in next(iterator, ())]
        return [{headers[index]: value for index, value in enumerate(row)} for row in iterator]
    finally:
        book.close()


def write_regression_report(v1_path: Path, v2: "V2RunResult", output_path: Path) -> Path:
    """比较数量和异常身份；不把不同结构强行判作业务错误。"""
    v1_period, v1_external = _v1_rows(v1_path, "两期对比"), _v1_rows(v1_path, "大集中对比")
    v1_rule = {(str(row.get("社会信用代码") or ""), str(row.get("指标编码") or ""), str(row.get("是否说明") or "")) for row in v1_period if row.get("是否说明")}
    v2_rule = {(item.social_credit_code, item.rule_id, item.description) for item in v2.findings if item.audit_type == "规则" and item.status == "异常"}
    v1_external_bad = {(str(row.get("社会信用代码") or ""), str(row.get("指标编码") or "")) for row in v1_external if row.get("是否说明")}
    v2_external_bad = {(item.social_credit_code, item.indicator_code) for item in v2.findings if item.audit_type == "外部核对" and item.status == "异常"}
    v2_by_type = Counter(item.audit_type for item in v2.findings if item.status == "异常")
    book = Workbook(); overview = book.active; overview.title = "汇总"
    rows = [
        ("V1 两期比较行", len(v1_period)), ("V2 环比结果行", len([i for i in v2.findings if i.audit_type == "环比"])),
        ("V1 大集中核对行", len(v1_external)), ("V2 大集中核对行", len([i for i in v2.findings if i.audit_type == "外部核对"])),
        ("V1 规则/说明命中行", len(v1_rule)), ("V2 规则命中", len(v2_rule)),
        ("V1 大集中异常", len(v1_external_bad)), ("V2 大集中异常", len(v2_external_bad)),
        ("V2 环比异常", v2_by_type.get("环比", 0)), ("V2 校验规则异常", v2_by_type.get("校验规则", 0)),
        ("说明", "V2 对规则按“机构+规则”输出一条独立结果，不能与 V1 挂载到多个指标行的数量直接等同。"),
        ("说明", "V2 对负数上期使用 (本期-上期)/abs(上期)，与 V1 原算法不同的记录应人工复核。"),
    ]
    overview.append(["项目", "数值"])
    for row in rows: overview.append(row)
    for title, values in (("V1有_V2无_外部", sorted(v1_external_bad - v2_external_bad)), ("V2有_V1无_外部", sorted(v2_external_bad - v1_external_bad))):
        sheet = book.create_sheet(title); sheet.append(["社会信用代码", "指标编码"])
        for row in values: sheet.append(row)
    rules = book.create_sheet("V2规则结果")
    rules.append(["机构", "规则编号", "规则描述", "涉及表单", "计算过程"])
    for item in v2.findings:
        if item.audit_type == "规则" and item.status == "异常":
            rules.append([item.institution_name, item.rule_id, item.description, "、".join(item.form_codes), item.calculation_trace])
    output_path.parent.mkdir(parents=True, exist_ok=True)
    book.save(output_path)
    return output_path
