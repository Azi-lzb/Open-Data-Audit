"""S2-F01 Excel 输出器。

业务引擎负责生成审核结论和取数语义；本模块只负责把三类 finding
按各自的字段契约写入三个独立工作表，不从空值猜测业务含义。
"""

from __future__ import annotations

from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from .models import AuditFinding, AuditSummary, IndicatorRecord, STATUS_ABNORMAL, STATUS_NORMAL, V2RunPaths, SEVERITIES


def build_summary(records: Iterable[IndicatorRecord], findings: Iterable[AuditFinding]) -> AuditSummary:
    record_list, finding_list = list(records), list(findings)
    abnormal = [item for item in finding_list if item.status == STATUS_ABNORMAL]
    return AuditSummary(
        institutions=len({item.institution_id for item in record_list}), forms=len({item.form_code for item in record_list if item.form_code}),
        indicators=len(record_list), total_findings=len(finding_list), abnormal_findings=len(abnormal),
        by_type=dict(Counter(item.audit_type for item in abnormal)), by_severity=dict(Counter(item.severity for item in abnormal if item.severity)),
    )


PERIOD_HEADERS = [
    "地区", "承接行", "机构类别", "机构名称", "社会信用代码", "表单", "指标代码", "指标名称",
    "规则编号", "审核级别", "审核说明", "本期值", "上期值", "差异值", "变动幅度", "取数说明", "计算过程",
]

EXTERNAL_HEADERS = [
    "地区", "承接行", "机构类别", "机构名称", "社会信用代码", "表单", "指标代码", "指标名称",
    "外部指标名称", "规则编号", "审核级别", "审核说明", "报表值", "外部值", "差异值", "取数说明", "计算过程",
]

VALIDATION_HEADERS = [
    "地区", "承接行", "机构类别", "机构名称", "社会信用代码", "涉及表单", "规则编号", "规则类型",
    "审核级别", "审核说明", "取值明细", "计算过程",
]

SHEET_ORDER = ("环比规则", "外部核对规则", "校验规则")

_AUDIT_TYPE_TO_SHEET = {
    "环比": "环比规则",
    "环比规则": "环比规则",
    "外部核对": "外部核对规则",
    "外部核对规则": "外部核对规则",
    "规则": "校验规则",
    "校验规则": "校验规则",
}

_LONG_TEXT_HEADERS = {"审核说明", "取数说明", "取值明细", "计算过程"}


def _validate_finding(item: AuditFinding) -> None:
    """Fail loudly on status/severity conflicts instead of correcting in export."""
    if item.status == STATUS_NORMAL:
        if item.severity not in (None, ""):
            raise ValueError(
                f"S2-F01 审核级别冲突：正常 finding 不应有审核级别；"
                f"finding_id={item.finding_id} 机构={item.institution_name} 表单={item.form_code}"
                f" 规则={item.rule_id} 指标={item.indicator_code} 状态={item.status} 级别={item.severity}"
            )
        return
    if item.status == STATUS_ABNORMAL and item.severity in SEVERITIES:
        return
    raise ValueError(
        f"S2-F01 审核级别冲突：异常 finding 必须是提示/关注/严重；"
        f"finding_id={item.finding_id} 机构={item.institution_name} 表单={item.form_code}"
        f" 规则={item.rule_id} 指标={item.indicator_code} 状态={item.status} 级别={item.severity}"
    )


def _value_or_blank(value: object) -> object:
    """Preserve real zero and numeric values; only missing values become blank cells."""
    return "" if value is None else value


def _period_rule_id(item: AuditFinding) -> str:
    return item.rule_id or item.band or "未命中"


def _period_description(item: AuditFinding) -> str:
    description = item.rule_description or item.description
    if description:
        return description
    rule_id = _period_rule_id(item)
    return "命中规则" if rule_id != "未命中" else "未命中规则"


def _period_values(item: AuditFinding) -> tuple[object, ...]:
    return (
        item.region, item.handling_branch, item.institution_type, item.institution_name, item.social_credit_code,
        "、".join(item.form_codes) or item.form_code, item.indicator_code, item.indicator_name,
        _period_rule_id(item), item.severity or "", _period_description(item),
        _value_or_blank(item.current_value), _value_or_blank(item.previous_value), _value_or_blank(item.difference_value),
        _value_or_blank(item.change_rate), item.retrieval_note or "", item.calculation_trace or "",
    )


def _external_values(item: AuditFinding) -> tuple[object, ...]:
    return (
        item.region, item.handling_branch, item.institution_type, item.institution_name, item.social_credit_code,
        "、".join(item.form_codes) or item.form_code, item.indicator_code, item.indicator_name,
        item.external_indicator_name, item.rule_id, item.severity or "", item.rule_description or item.description,
        _value_or_blank(item.current_value), _value_or_blank(item.external_value), _value_or_blank(item.difference_value),
        item.retrieval_note or "", item.calculation_trace or "",
    )


def _validation_form_scope(item: AuditFinding) -> str:
    related_forms = (record.form_code for record in item.related_records)
    forms = tuple(sorted({value for value in (*item.form_codes, item.form_code, *related_forms) if value}))
    if forms:
        return "、".join(forms)
    return "不适用" if item.rule_type in {"导入校验", "配置检查"} else "未识别"


def _validation_values(item: AuditFinding) -> tuple[object, ...]:
    rule_type = item.rule_type or "未识别"
    value_details = item.value_details or ("不适用" if rule_type in {"导入校验", "配置检查"} else "未识别")
    return (
        item.region or ("不适用" if not item.institution_name else ""),
        item.handling_branch or ("不适用" if not item.institution_name else ""),
        item.institution_type or ("不适用" if not item.institution_name else ""),
        item.institution_name or "全局", item.social_credit_code or ("不适用" if not item.institution_name else ""),
        _validation_form_scope(item), item.rule_id or "未配置", rule_type, item.severity or "",
        item.rule_description or item.description, value_details, item.calculation_trace or "",
    )


def _write_table(sheet, headers: list[str], rows: Iterable[Iterable[object]], *, percentage_rows: set[int] | None = None, point_rows: set[int] | None = None) -> None:
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in rows:
        sheet.append(list(row))

    percentage_rows = percentage_rows or set()
    point_rows = point_rows or set()
    if "变动幅度" in headers:
        column = headers.index("变动幅度") + 1
        for row_index in percentage_rows:
            sheet.cell(row=row_index, column=column).number_format = "0.00%"
        for row_index in point_rows:
            sheet.cell(row=row_index, column=column).number_format = "0.00"

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    header_index = {value: index + 1 for index, value in enumerate(headers)}
    for title in _LONG_TEXT_HEADERS.intersection(headers):
        column = header_index[title]
        for row in range(1, sheet.max_row + 1):
            sheet.cell(row=row, column=column).alignment = Alignment(vertical="top", wrap_text=True)
    for column in sheet.columns:
        letter = column[0].column_letter
        values = [str(cell.value or "") for cell in column]
        sheet.column_dimensions[letter].width = min(max(max((len(value) for value in values), default=0) + 2, 12), 42)


def _group_findings(findings: Iterable[AuditFinding]) -> dict[str, list[AuditFinding]]:
    grouped = {name: [] for name in SHEET_ORDER}
    for item in findings:
        _validate_finding(item)
        sheet_name = _AUDIT_TYPE_TO_SHEET.get(item.audit_type)
        if sheet_name is None:
            raise ValueError(f"S2-F01 遇到未知审核类型，未生成结果：{item.audit_type!r} finding_id={item.finding_id}")
        grouped[sheet_name].append(item)
    return grouped


def _hide_rows_with_description(sheet, description: str) -> None:
    """Keep findings in the workbook while hiding rows with an exact displayed description."""
    description_column = next(
        cell.column for cell in sheet[1] if cell.value == "审核说明"
    )
    for row_index in range(2, sheet.max_row + 1):
        if sheet.cell(row_index, description_column).value == description:
            sheet.row_dimensions[row_index].hidden = True


class ExcelExporter:
    def export(self, *, output_dir: Path, current_records: list[IndicatorRecord], previous_records: list[IndicatorRecord], findings: list[AuditFinding], run_info: dict[str, str], hide_unchanged_period_rows: bool = False, hide_matching_external_rows: bool = False) -> V2RunPaths:
        output_dir.mkdir(parents=True, exist_ok=True)
        result_path = output_dir / f"报表审核清单_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        grouped = _group_findings(findings)
        book = Workbook()
        first = book.active
        first.title = SHEET_ORDER[0]
        for sheet_name in SHEET_ORDER[1:]:
            book.create_sheet(sheet_name)

        period_items = grouped["环比规则"]
        period_rows = [_period_values(item) for item in period_items]
        percentage_rows = {index for index, item in enumerate(period_items, start=2) if item.value_type != "百分数"}
        point_rows = {index for index, item in enumerate(period_items, start=2) if item.value_type == "百分数"}
        _write_table(book["环比规则"], PERIOD_HEADERS, period_rows, percentage_rows=percentage_rows, point_rows=point_rows)
        _write_table(book["外部核对规则"], EXTERNAL_HEADERS, (_external_values(item) for item in grouped["外部核对规则"]))
        _write_table(book["校验规则"], VALIDATION_HEADERS, (_validation_values(item) for item in grouped["校验规则"]))
        if hide_unchanged_period_rows:
            _hide_rows_with_description(book["环比规则"], "无变动")
        if hide_matching_external_rows:
            _hide_rows_with_description(book["外部核对规则"], "与大集中系统数据一致")
        book.save(result_path)
        return V2RunPaths(result_path)
