"""V2 源数据与大集中数据导入器；所有 Excel 细节停留在本层。"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .models import IndicatorRecord, ReportAuditConfig, SourceRef
from .unit_conversion import AMOUNT_DATA_TYPES, convert_amount


class ImportErrorV2(ValueError):
    pass


NAME_CODE = "20201001"
ID_CODE = "20201002"


def _code(value: Any) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value or "").strip()
    # 兼容字符串形态的浮点指标代码（"20201028.0"），与 float 读出形态归一，
    # 避免与指标参照（"20201028"）lookup miss 后误按普通数值处理。
    if re.fullmatch(r"\d+\.0+", text):
        return str(int(float(text)))
    return text


def _number(value: Any) -> float | None:
    if isinstance(value, bool) or value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip().replace(",", ""))
    except ValueError:
        return None


def _decimal_number(value: Any) -> Decimal | None:
    if isinstance(value, bool) or value in (None, ""):
        return None
    try:
        number = value if isinstance(value, Decimal) else Decimal(str(value).strip().replace(",", ""))
    except (InvalidOperation, TypeError, ValueError):
        return None
    return number if number.is_finite() else None


def _period(date: str) -> str:
    text = str(date or "")
    digits = "".join(ch for ch in text if ch.isdigit())
    return f"{digits[:4]}-{digits[4:6]}" if len(digits) >= 6 else text


def _form(sheet_name: str, fallback: str) -> str:
    digits = "".join(ch for ch in str(sheet_name) if ch.isdigit())
    return digits[:5] if len(digits) >= 5 and digits.startswith("2020") else fallback


def _open(path: Path):
    if path.suffix.casefold() == ".xls":
        import xlrd
        book = xlrd.open_workbook(str(path))
        def rows(name: str):
            sheet = book.sheet_by_name(name)
            return [[sheet.cell_value(r, c) if c < sheet.ncols else None for c in range(3)] for r in range(sheet.nrows)]
        return list(book.sheet_names()), rows, lambda: None
    book = load_workbook(path, read_only=True, data_only=True)
    def rows(name: str):
        return [list(row[:3]) for row in book[name].iter_rows(values_only=True)]
    return list(book.sheetnames), rows, book.close


def _data_rows(rows: list[list[Any]]) -> list[tuple[int, str, str, Any]]:
    result = []
    for position, row in enumerate(rows[3:], start=4):
        if not row:
            continue
        code = _code(row[0] if len(row) > 0 else None)
        if not code:
            continue
        result.append((position, code, str(row[1] if len(row) > 1 and row[1] is not None else "").strip(), row[2] if len(row) > 2 else None))
    return result


@dataclass
class ImportResult:
    records: list[IndicatorRecord] = field(default_factory=list)
    duplicates: list[tuple[IndicatorRecord, IndicatorRecord]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def by_logical_key(self) -> dict[tuple[str, str, str], IndicatorRecord]:
        """重复项由 validation 处理；这里保留首条，绝不由后项静默覆盖。"""
        result: dict[tuple[str, str, str], IndicatorRecord] = {}
        for record in self.records:
            result.setdefault(record.logical_key, record)
        return result


def import_period_directory(directory: Path, *, config: ReportAuditConfig, label: str) -> ImportResult:
    if not directory.is_dir():
        raise ImportErrorV2(f"{label}目录不存在：{directory}")
    files = [p for p in sorted(directory.iterdir()) if p.is_file() and p.suffix.casefold() in {".xls", ".xlsx"} and not p.name.startswith("~$")]
    if not files:
        raise ImportErrorV2(f"{label}目录没有可读取的 .xls/.xlsx 文件：{directory}")
    result = ImportResult()
    seen: dict[tuple[str, str, str], IndicatorRecord] = {}
    # 按报表划分的部分子表没有名称指标，先缓存同批已识别的名称。
    names_by_id: dict[str, str] = {}
    pending: list[tuple[Path, str, int, str, str, Any, str, str, str]] = []
    for path in files:
        parts = path.stem.split("#")
        banks = len(parts) >= 4 and parts[0].casefold() == "banks"
        if banks:
            date, fallback_form, fallback_id, fallback_name = parts[1], parts[3], "", ""
        elif len(parts) >= 5 and parts[0].casefold() == "reports":
            fallback_id, date, fallback_form, fallback_name = parts[1], parts[2], "", parts[4]
        elif len(parts) >= 5:
            fallback_id, date, fallback_form, fallback_name = parts[0], parts[1], parts[3], parts[4]
        else:
            result.warnings.append(f"无法识别文件名，已跳过：{path.name}")
            continue
        sheets, read_rows, close = _open(path)
        try:
            for sheet_name in sheets:
                data = _data_rows(read_rows(sheet_name))
                lookup = {code: value for _row, code, _name, value in data}
                institution_id = _code(lookup.get(ID_CODE)) or (_code(sheet_name) if banks else _code(fallback_id))
                institution_name = str(lookup.get(NAME_CODE) or fallback_name or "").strip()
                if institution_id and institution_name:
                    names_by_id[institution_id] = institution_name
                form = _form(sheet_name, fallback_form)
                for source_row, code, name, value in data:
                    pending.append((path, sheet_name, source_row, institution_id, institution_name, value, code, name, form + "\x00" + str(date)))
        finally:
            close()
    for path, sheet, row, raw_id, raw_name, value, code, name, form_date in pending:
        form, date = form_date.split("\x00", 1)
        profile = config.institutions_by_id.get(raw_id) or config.institutions_by_name.get(raw_name)
        institution_id = profile.institution_id if profile else (raw_id or raw_name)
        institution_name = profile.institution_name if profile else (raw_name or names_by_id.get(raw_id, raw_id))
        indicator = config.indicators.get(code)
        expected_form = indicator.form_code if indicator else ""
        if expected_form and form and expected_form != form:
            result.warnings.append(f"表单归属不一致：{path.name}/{sheet} 第{row}行，声明表单 {form}，指标 {code} 配置归属 {expected_form}")
        is_text = bool(indicator and indicator.data_type == "文字")
        is_amount = bool(indicator and indicator.data_type in AMOUNT_DATA_TYPES)
        numeric = _decimal_number(value) if is_amount else _number(value)
        if is_text or numeric is None:
            normalized = str(value or "").strip() if is_text else value
            normalized_unit = ""
        elif is_amount:
            try:
                normalized = convert_amount(
                    numeric, config.unit_settings.source_unit,
                    config.unit_settings.output_file_unit,
                )
            except ValueError as exc:
                raise ImportErrorV2(
                    f"{path.name}/{sheet} 第{row}行指标 {code} 无法转换金额到输出单位：{exc}"
                ) from exc
            normalized_unit = config.unit_settings.output_file_unit
        else:
            # 百分数、个数及文字属性原值参与业务计算，不经过金额换算。
            normalized = numeric
            normalized_unit = ""
        record = IndicatorRecord(
            institution_id=institution_id, institution_name=institution_name,
            social_credit_code=profile.social_credit_code if profile else raw_id,
            institution_type=profile.institution_type if profile else "",
            handling_branch=profile.handling_branch if profile else "", region=profile.region if profile else "",
            form_code=form or expected_form, indicator_code=code, indicator_name=name or (indicator.indicator_name if indicator else ""),
            period=_period(date), source_date=date, value=value,
            # 保留配置的数据属性；百分数需在环比引擎按百分点差异处理，不能被统一降为“数值”。
            value_type=indicator.data_type if indicator and indicator.data_type else ("文字" if is_text else "数值"),
            source_unit=config.unit_settings.source_unit if is_amount else "原始单位",
            normalized_value=normalized, normalized_unit=normalized_unit,
            source=SourceRef(str(path.resolve()), sheet, row),
        )
        previous = seen.get(record.logical_key)
        if previous is not None:
            result.duplicates.append((previous, record))
            continue
        seen[record.logical_key] = record
        result.records.append(record)
    return result


@dataclass(frozen=True)
class ExternalValue:
    row_label: str
    indicator_name: str
    value: float | Decimal
    source: SourceRef


def _open_external(path: Path):
    """打开大集中数据文件，返回 (表名列表, rows(name)→矩阵, close)。

    仅支持 .xlsx：外部核对依赖「参照表」工作表做机构映射，.xls（xlrd 链路）
    与 .csv（无工作表概念）都会丢「参照表」，故明确不接受（DEFECT-4 定案）。
    """
    if path.suffix.casefold() != ".xlsx":
        raise ImportErrorV2(
            f"大集中数据仅支持 .xlsx 工作簿：{path.name}"
            "（.xls/.csv 会丢失「参照表」机构映射，不支持）")
    book = load_workbook(path, read_only=True, data_only=True)
    def rows(name: str):
        return [list(row) for row in book[name].iter_rows(values_only=True)]
    return list(book.sheetnames), rows, book.close


def import_external_workbook(path: Path) -> tuple[dict[tuple[str, str], ExternalValue], dict[str, str]]:
    """兼容“集中系统数据”与可选“参照表”布局；仅支持 .xlsx（.xls/.csv 会丢参照表）。"""
    if not path.is_file():
        raise ImportErrorV2(f"大集中文件不存在：{path}")
    sheets, read_rows, close = _open_external(path)
    try:
        if "集中系统数据" not in sheets:
            raise ImportErrorV2("大集中文件缺少“集中系统数据”工作表")
        matrix = read_rows("集中系统数据")
        header_row = next((i for i, row in enumerate(matrix) if any(isinstance(v, str) and v.strip() for v in row)), None)
        if header_row is None:
            raise ImportErrorV2("集中系统数据未找到表头")
        header = matrix[header_row]
        label_col = next((c for c in range(min(4, len(header))) if not str(header[c] or "").strip() and any(isinstance(r[c] if c < len(r) else None, str) and str(r[c]).strip() for r in matrix[header_row + 1:header_row + 5])), None)
        if label_col is None:
            raise ImportErrorV2("集中系统数据未找到机构行标签列")
        values: dict[tuple[str, str], ExternalValue] = {}
        for col, title in enumerate(header):
            title = str(title or "").strip()
            if not title or col == label_col:
                continue
            for row_number, row in enumerate(matrix[header_row + 1:], start=header_row + 2):
                if col >= len(row):
                    continue
                label, number = str(row[label_col] or "").strip(), _decimal_number(row[col])
                if label and number is not None:
                    values[(label, title)] = ExternalValue(
                        label, title, number, SourceRef(str(path.resolve()), "集中系统数据", row_number)
                    )
        org_map: dict[str, str] = {}
        if "参照表" in sheets:
            rows = read_rows("参照表")
            if rows:
                headers = [str(v or "").strip() for v in rows[0]]
                if "统一社会信用代码" in headers and "报表项目" in headers:
                    ci, ri = headers.index("统一社会信用代码"), headers.index("报表项目")
                    for row in rows[1:]:
                        if len(row) > max(ci, ri) and row[ci] and row[ri]:
                            org_map[str(row[ci]).strip()] = str(row[ri]).strip()
        return values, org_map
    finally:
        close()
