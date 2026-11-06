"""从大集中比较结果导出人工选定的指标说明文件。

这是比较完成后的人工复核工具：用户在比较结果的「输出说明」列标记“是”，
本模块仅整理这些行，不重算、不修改比较结果，也不依赖 Office/COM。
"""

from __future__ import annotations

from datetime import datetime
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
from shutil import copy2

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


_REQUIRED_HEADERS = (
    "机构类代码", "机构类名称", "地区代码", "地区名称", "指标代码", "指标名称", "数据值", "上期值",
    "输出说明", "说明内容",
)
_SELECTED_VALUES = {"是", "y", "yes", "true", "1", "√", "✓"}
_ORG_REFERENCE_SHEET = "机构地区参照"
_MATCH_ABS_TOLERANCE = Decimal("1e-12")
_MATCH_REL_TOLERANCE = Decimal("1e-12")


@dataclass(frozen=True)
class ExplanationImportResult:
    imported: int
    unmatched: int
    conflicts: int
    backup_path: Path
    conflict_report_path: Path | None = None


def _header_index(sheet) -> dict[str, int]:
    headers = {
        str(cell.value).strip(): cell.column
        for cell in sheet[1]
        if cell.value not in (None, "")
    }
    missing = [header for header in _REQUIRED_HEADERS if header not in headers]
    if missing:
        raise ValueError("比较结果缺少说明导出所需列：" + "、".join(missing))
    return headers


def _selected(value: object) -> bool:
    return str(value or "").strip().lower() in _SELECTED_VALUES


def _unit_from_change_header(header: str) -> str:
    match = re.search(r"[（(]([^（）()]+)[）)]\s*$", str(header or ""))
    return match.group(1).strip() if match else "亿元"


def _explanation_sheet_title(comparison_title: str) -> str:
    match = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月", str(comparison_title or ""))
    if not match:
        return "说明文件"
    return f"(说明){match.group(1)}.{int(match.group(2)):02d} 说明文件"


def _load_org_reference(config_path: Path | None) -> dict[tuple[str, str], str]:
    if config_path is None or not Path(config_path).is_file():
        return {}
    book = load_workbook(config_path, read_only=True, data_only=True)
    try:
        if _ORG_REFERENCE_SHEET not in book.sheetnames:
            return {}
        sheet = book[_ORG_REFERENCE_SHEET]
        headers = {str(cell.value).strip(): cell.column for cell in sheet[1] if cell.value not in (None, "")}
        required = ("机构类代码", "地区代码", "机构名称")
        if any(name not in headers for name in required):
            return {}
        mapping: dict[tuple[str, str], str] = {}
        names_by_org: dict[str, set[str]] = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if len(row) < max(headers[name] for name in required):
                continue
            org_code = str(row[headers["机构类代码"] - 1] or "").replace("'", "").strip()
            region_code = str(row[headers["地区代码"] - 1] or "").replace("'", "").strip()
            name = str(row[headers["机构名称"] - 1] or "").strip()
            if not org_code or not name:
                continue
            mapping[(org_code, region_code)] = name
            names_by_org.setdefault(org_code, set()).add(name)
        # 用户可在配置中显式填写地区代码“*”；只有唯一名称的机构也自动视为全地区适用。
        for org_code, names in names_by_org.items():
            if len(names) == 1:
                mapping.setdefault((org_code, "*"), next(iter(names)))
        return mapping
    finally:
        book.close()


def export_selected_explanations(
    comparison_file: Path,
    output_path: Path,
    *,
    config_path: Path | None = None,
    format_mode: str = "MONTHLY",
    trim_indicator_names: bool = False,
) -> int:
    """导出「输出说明」标记为“是”的行，返回导出条数。

    月报沿用旧版八列；日报在机构名称后插入地区代码、地区名称两列。
    两种格式的机构代码均保持原始代码，不拼接地区代码。
    """
    comparison_file, output_path = Path(comparison_file), Path(output_path)
    format_mode = str(format_mode or "MONTHLY").strip().upper()
    if format_mode not in {"MONTHLY", "DAILY"}:
        raise ValueError("说明文件格式只能选择 MONTHLY 或 DAILY")
    source = load_workbook(comparison_file, read_only=True, data_only=True)
    try:
        sheet = next(
            (item for item in source.worksheets if "指标代码" in {
                str(cell.value).strip() for cell in item[1] if cell.value not in (None, "")
            }),
            None,
        )
        if sheet is None:
            raise ValueError("比较结果中未找到包含“指标代码”的结果工作表")
        index = _header_index(sheet)
        change_header = next((name for name in index if name.startswith("增减额")), "增减额")
        unit = _unit_from_change_header(change_header)
        org_reference = _load_org_reference(config_path)
        selected_rows = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not _selected(row[index["输出说明"] - 1] if len(row) >= index["输出说明"] else None):
                continue
            org_code = str(row[index["机构类代码"] - 1] or "").replace("'", "").strip()
            region_code = str(row[index["地区代码"] - 1] or "").replace("'", "").strip()
            org_name = str(row[index["机构类名称"] - 1] or "").strip()
            region_name = str(row[index["地区名称"] - 1] or "").strip()
            base_name = (org_reference.get((org_code, region_code))
                         or org_reference.get((org_code, "*"))
                         or org_name)
            indicator_name = row[index["指标名称"] - 1]
            if trim_indicator_names and isinstance(indicator_name, str):
                indicator_name = indicator_name.strip()
            values = [org_code, base_name]
            if format_mode == "DAILY":
                values.extend((region_code, region_name))
            values.extend((
                row[index["指标代码"] - 1], indicator_name,
                row[index["数据值"] - 1], row[index["上期值"] - 1],
                row[index[change_header] - 1] if len(row) >= index[change_header] else None,
                row[index["说明内容"] - 1],
            ))
            selected_rows.append(tuple(values))
    finally:
        source.close()

    if format_mode == "DAILY":
        selected_rows.sort(key=lambda row: (
            str(row[0] or "").casefold(), str(row[1] or "").casefold(),
            str(row[2] or "").casefold(), str(row[4] or "").casefold(),
        ))
    else:
        selected_rows.sort(key=lambda row: (
            str(row[0] or "").casefold(), str(row[1] or "").casefold(),
            str(row[2] or "").casefold(),
        ))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    explanation = book.active
    explanation.title = _explanation_sheet_title(sheet.title)
    last_column = 10 if format_mode == "DAILY" else 8
    last_column_letter = get_column_letter(last_column)
    explanation.merge_cells(f"A1:{last_column_letter}1")
    explanation["A1"] = "xxx说明"
    source_font = Font(name="宋体", size=11)
    explanation["A1"].font = source_font
    explanation["A1"].alignment = Alignment(horizontal="center", vertical="center")
    headers = ["机构代码", "机构名称"]
    if format_mode == "DAILY":
        headers.extend(("地区代码", "地区名称"))
    headers.extend((
        "指标代码", "指标名称", f"本期金额（{unit}）", f"上期金额（{unit}）",
        f"增减额（{unit}）", "变动原因",
    ))
    for column, value in enumerate(headers, start=1):
        cell = explanation.cell(2, column, value)
        cell.font = source_font
        cell.alignment = Alignment(vertical="center")
        cell.number_format = "@" if column <= (6 if format_mode == "DAILY" else 4) else "General"
    for row in selected_rows:
        explanation.append(row)
    text_columns = 6 if format_mode == "DAILY" else 4
    for row in explanation.iter_rows(min_row=3, max_row=explanation.max_row,
                                     min_col=1, max_col=last_column):
        for cell in row:
            cell.font = source_font
            cell.alignment = Alignment(vertical="center")
            cell.number_format = "@" if cell.column <= text_columns else "0.00"

    if format_mode == "MONTHLY":
        start = 3
        while start <= explanation.max_row:
            code = explanation.cell(start, 1).value
            name = explanation.cell(start, 2).value
            end = start
            while (end + 1 <= explanation.max_row
                   and explanation.cell(end + 1, 1).value == code
                   and explanation.cell(end + 1, 2).value == name):
                end += 1
            if end > start:
                explanation.merge_cells(start_row=start, start_column=1, end_row=end, end_column=1)
                explanation.merge_cells(start_row=start, start_column=2, end_row=end, end_column=2)
            start = end + 1
        for letter, width in {"A": 9.0, "B": 24.875, "C": 13.875, "D": 37.625, "E": 9.0}.items():
            explanation.column_dimensions[letter].width = width
    else:
        # Daily format keeps institution/region values on every row so the added
        # region columns remain useful for filtering.
        widths = {"A": 9.0, "B": 24.875, "C": 12.0, "D": 18.0,
                  "E": 13.875, "F": 37.625, "G": 14.0, "H": 14.0,
                  "I": 14.0, "J": 48.0}
        for letter, width in widths.items():
            explanation.column_dimensions[letter].width = width
        explanation.auto_filter.ref = f"A2:{last_column_letter}{max(2, explanation.max_row)}"

    book.save(output_path)
    book.close()
    return len(selected_rows)


def _find_explanation_sheet(book):
    if "说明文件" in book.sheetnames:
        return book["说明文件"]
    for sheet in book.worksheets:
        for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 10), values_only=True):
            if "机构代码" in {str(value).strip() for value in row if value not in (None, "")}:
                return sheet
    raise ValueError("未找到说明文件工作表或“机构代码”表头")


def _explanation_header_index(sheet) -> tuple[int, dict[str, int]]:
    for row_number in range(1, min(sheet.max_row, 10) + 1):
        headers = {
            str(cell.value).strip(): cell.column
            for cell in sheet[row_number]
            if cell.value not in (None, "")
        }
        if all(name in headers for name in ("机构代码", "机构名称", "指标代码", "指标名称")):
            value_header = next((name for name in headers if name == "数据值" or name.startswith("本期金额")), None)
            change_header = next((name for name in headers if name.startswith("增减额")), None)
            if value_header and change_header and "变动原因" in headers:
                columns = {
                    "机构代码": headers["机构代码"],
                    "机构名称": headers["机构名称"],
                    **({"地区代码": headers["地区代码"]} if "地区代码" in headers else {}),
                    "指标代码": headers["指标代码"],
                    "指标名称": headers["指标名称"],
                    "数据值": headers[value_header],
                    "增减额": headers[change_header],
                    "变动原因": headers["变动原因"],
                }
                previous_header = next((name for name in headers if name == "上期值" or name.startswith("上期金额")), None)
                if previous_header:
                    columns["上期值"] = headers[previous_header]
                return row_number, columns
    raise ValueError("说明文件缺少机构代码、机构名称、指标代码、指标名称、数据值、增减额或变动原因列")


def _code_value(value: object) -> str:
    """机构/指标代码的匹配值：文本化、去 CSV 撇号、数字去尾随 .0、统一大写。

    代码是标识不是数量，禁止数值化——Decimal 归一会把 7020 变成 7.02E+3，
    与机构按文本填写的代码永不相等（真实批次实测教训）。大小写归一对齐
    Excel/VBA 时代的宽松比较（龙门 6K0u vs 6k0u 实测案例）。
    """
    if value is None:
        return ""
    text = str(value).strip().replace("'", "")
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text.upper()


def _amount_decimal(value: object) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        number = Decimal(text.replace(",", ""))
        return number if number.is_finite() else None
    except InvalidOperation:
        return None


def _amounts_match(left: object, right: object) -> bool:
    """Compare Excel numeric values with 1e-12 absolute and relative tolerance."""
    left_text = "" if left is None else str(left).strip()
    right_text = "" if right is None else str(right).strip()
    if not left_text or not right_text:
        return not left_text and not right_text

    left_number = _amount_decimal(left)
    right_number = _amount_decimal(right)
    if left_number is not None and right_number is not None:
        tolerance = max(
            _MATCH_ABS_TOLERANCE,
            _MATCH_REL_TOLERANCE * max(abs(left_number), abs(right_number)),
        )
        return abs(left_number - right_number) <= tolerance
    return left_text.casefold() == right_text.casefold()


def _identity(row: tuple[object, ...], columns: dict[str, int]) -> tuple[str, ...]:
    """Return the stable feedback key; institution name is part of identity."""
    org = _code_value(row[columns["机构代码"] - 1]) if len(row) >= columns["机构代码"] else ""
    org = org.replace(" ", "").replace("　", "")
    name = _institution_name_value(_row_value(row, columns, "机构名称"))
    indicator = _code_value(row[columns["指标代码"] - 1]) if len(row) >= columns["指标代码"] else ""
    indicator = indicator.replace(" ", "").replace("　", "")
    if "地区代码" in columns:
        region = _code_value(row[columns["地区代码"] - 1]) if len(row) >= columns["地区代码"] else ""
        region = region.replace(" ", "").replace("　", "")
        return org, name, indicator, region
    return org, name, indicator


def _identity_without_name(key: tuple[str, ...]) -> tuple[str, ...]:
    """Diagnostic index used only to report a name mismatch, never to match rows."""
    return (key[0], key[2], *key[3:])


def _institution_name_value(value: object) -> str:
    """Normalize institution names by removing whitespace and case-folding only."""
    return "".join(str(value or "").split()).casefold()


def _row_value(row: tuple[object, ...], columns: dict[str, int], name: str) -> object:
    position = columns[name] - 1
    return row[position] if len(row) > position else None


def _merged_identity_values(sheet, columns: dict[str, int]) -> dict[tuple[int, int], object]:
    """为纵向合并的机构/地区标识建立逐行取值，供 F3 反馈匹配使用。"""
    identity_columns = {
        columns[name]
        for name in ("机构代码", "机构名称", "地区代码")
        if name in columns
    }
    values: dict[tuple[int, int], object] = {}
    for merged in getattr(getattr(sheet, "merged_cells", None), "ranges", ()):
        if merged.min_col != merged.max_col or merged.min_col not in identity_columns:
            continue
        anchor = sheet.cell(merged.min_row, merged.min_col).value
        if anchor in (None, ""):
            continue
        for row_number in range(merged.min_row + 1, merged.max_row + 1):
            values[(row_number, merged.min_col)] = anchor
    return values


def _row_with_merged_identity(
    sheet, row_number: int, max_column: int, merged_values: dict[tuple[int, int], object]
) -> tuple[object, ...]:
    row = [sheet.cell(row_number, column).value for column in range(1, max_column + 1)]
    for column in range(1, max_column + 1):
        if (row_number, column) in merged_values:
            row[column - 1] = merged_values[(row_number, column)]
    return tuple(row)


def _last_explanation_row(sheet, header_row: int, columns: dict[str, int]) -> int:
    indicator_column = columns["指标代码"]
    for row_number in range(sheet.max_row, header_row, -1):
        if sheet.cell(row_number, indicator_column).value not in (None, ""):
            return row_number
    return header_row


def import_explanation_feedback(target_path: Path, feedback_paths: list[Path]) -> ExplanationImportResult:
    """回写机构反馈：名称参与身份键，数值按容差匹配，日报增加地区代码。"""
    target_path = Path(target_path)
    if not feedback_paths:
        raise ValueError("请至少选择一个机构反馈说明文件")
    target = load_workbook(target_path)
    try:
        target_sheet = _find_explanation_sheet(target)
        target_header_row, target_columns = _explanation_header_index(target_sheet)
        target_merged_values = _merged_identity_values(target_sheet, target_columns)
        target_last_row = _last_explanation_row(target_sheet, target_header_row, target_columns)
        target_last_column = max(target_columns.values())
        targets: dict[tuple[str, ...], list[tuple[int, tuple[object, ...]]]] = {}
        targets_without_name: dict[tuple[str, ...], list[tuple[int, tuple[object, ...]]]] = {}
        for row_number in range(target_header_row + 1, target_last_row + 1):
            row = _row_with_merged_identity(
                target_sheet, row_number, target_last_column, target_merged_values
            )
            key = _identity(row, target_columns)
            if len(key) >= 3 and key[0] and key[1] and key[2]:
                targets.setdefault(key, []).append((row_number, row))
                targets_without_name.setdefault(_identity_without_name(key), []).append((row_number, row))

        feedback_rows: list[tuple[tuple[str, ...], tuple[object, ...], dict[str, int], str, str, str]] = []
        for feedback_path in feedback_paths:
            source = load_workbook(feedback_path, data_only=True)
            try:
                sheet = _find_explanation_sheet(source)
                header_row, columns = _explanation_header_index(sheet)
                merged_values = _merged_identity_values(sheet, columns)
                last_row = _last_explanation_row(sheet, header_row, columns)
                last_column = max(columns.values())
                for row_number in range(header_row + 1, last_row + 1):
                    row = _row_with_merged_identity(sheet, row_number, last_column, merged_values)
                    key = _identity(row, columns)
                    reason = str(_row_value(row, columns, "变动原因") or "").strip()
                    if len(key) >= 3 and key[0] and key[1] and key[2] and reason:
                        feedback_rows.append((
                            key,
                            row,
                            columns,
                            str(_row_value(row, columns, "机构名称") or ""),
                            reason,
                            feedback_path.name,
                        ))
            finally:
                source.close()

        imported = unmatched = conflicts = 0
        conflict_rows: list[tuple[object, ...]] = []
        conflict_signatures: set[tuple[object, ...]] = set()
        blocked_targets: set[int] = set()
        matched_targets: dict[int, dict[str, object]] = {}

        def target_field(row: tuple[object, ...], name: str) -> object:
            return _row_value(row, target_columns, name)

        def add_conflict(
            candidates: list[tuple[int, tuple[object, ...]]],
            feedback_name: str,
            reason_sources: dict[str, set[str]],
            cause: str,
        ) -> None:
            nonlocal conflicts
            if not candidates:
                return
            candidate_numbers = tuple(sorted(number for number, _row in candidates))
            signature = (candidate_numbers, _institution_name_value(feedback_name), tuple(sorted(reason_sources)), cause)
            blocked_targets.update(candidate_numbers)
            if signature in conflict_signatures:
                return
            conflict_signatures.add(signature)
            conflicts += 1
            reasons_text = "\n".join(
                f"{reason}（来源：{'、'.join(sorted(sources))}）"
                for reason, sources in sorted(reason_sources.items())
            )
            sample_number, sample_row = candidates[0]
            conflict_rows.append((
                "冲突",
                str(target_field(sample_row, "机构代码") or ""),
                str(target_field(sample_row, "机构名称") or ""),
                str(target_field(sample_row, "指标代码") or ""),
                str(target_field(sample_row, "指标名称") or ""),
                str(target_field(sample_row, "数据值")) if target_field(sample_row, "数据值") is not None else "",
                str(target_field(sample_row, "增减额")) if target_field(sample_row, "增减额") is not None else "",
                "、".join(str(number) for number in candidate_numbers),
                reasons_text,
                f"{cause}；反馈机构名称：{feedback_name or '（空）'}",
            ))

        for key, feedback_row, feedback_columns, feedback_name, reason, source_name in feedback_rows:
            identity_candidates = targets.get(key, [])
            target_rows = [
                (row_number, row)
                for row_number, row in identity_candidates
                if _amounts_match(
                    _row_value(row, target_columns, "数据值"),
                    _row_value(feedback_row, feedback_columns, "数据值"),
                )
                and _amounts_match(
                    _row_value(row, target_columns, "增减额"),
                    _row_value(feedback_row, feedback_columns, "增减额"),
                )
            ]
            if not target_rows:
                # Exact identity includes institution name. If the same code,
                # indicator, optional region and numeric values exist under a
                # different name, report the mismatch instead of silently
                # treating it as an ordinary unmatched feedback row.
                diagnostic_candidates = targets_without_name.get(_identity_without_name(key), [])
                diagnostic_rows = [
                    (row_number, row)
                    for row_number, row in diagnostic_candidates
                    if _amounts_match(
                        _row_value(row, target_columns, "数据值"),
                        _row_value(feedback_row, feedback_columns, "数据值"),
                    )
                    and _amounts_match(
                        _row_value(row, target_columns, "增减额"),
                        _row_value(feedback_row, feedback_columns, "增减额"),
                    )
                ]
                if diagnostic_rows:
                    add_conflict(
                        diagnostic_rows,
                        feedback_name,
                        {reason: {source_name}},
                        "机构名称与目标说明文件不一致",
                    )
                else:
                    unmatched += 1
                continue

            if len(target_rows) > 1:
                add_conflict(
                    target_rows,
                    feedback_name,
                    {reason: {source_name}},
                    f"完整匹配键命中 {len(target_rows)} 个目标行",
                )
                continue
            target_row_number, target_row = target_rows[0]

            entry = matched_targets.setdefault(
                target_row_number,
                {"row": target_row, "reasons": {}, "names": set()},
            )
            reason_map = entry["reasons"]
            assert isinstance(reason_map, dict)
            sources = reason_map.setdefault(reason, set())
            assert isinstance(sources, set)
            sources.add(source_name)
            names = entry["names"]
            assert isinstance(names, set)
            names.add(feedback_name)

        for target_row_number, entry in matched_targets.items():
            if target_row_number in blocked_targets:
                continue
            target_row = entry["row"]
            reasons = entry["reasons"]
            names = entry["names"]
            assert isinstance(target_row, tuple) and isinstance(reasons, dict) and isinstance(names, set)
            if len(reasons) != 1:
                reason_sources = {
                    str(reason): set(sources)
                    for reason, sources in reasons.items()
                    if isinstance(sources, set)
                }
                add_conflict(
                    [(target_row_number, target_row)],
                    "、".join(sorted(names)),
                    reason_sources,
                    "多个反馈文件给出不同变动原因",
                )
                continue
            target_sheet.cell(target_row_number, target_columns["变动原因"]).value = next(iter(reasons))
            imported += 1

        backup_path = target_path.with_name(
            f"{target_path.stem}_导入前备份_{datetime.now():%Y%m%d_%H%M%S}{target_path.suffix}"
        )
        copy2(target_path, backup_path)
        target.save(target_path)
        conflict_report_path = None
        if conflict_rows:
            conflict_report_path = target_path.with_name(
                f"{target_path.stem}_导入冲突清单_{datetime.now():%Y%m%d_%H%M%S}{target_path.suffix}"
            )
            report = Workbook()
            sheet = report.active
            sheet.title = "冲突明细"
            sheet.append((
                "状态", "机构代码", "机构名称", "指标代码", "指标名称", "数据值", "增减额",
                "目标说明文件行号", "反馈文件及原因", "冲突原因",
            ))
            for row in conflict_rows:
                sheet.append(row)
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="C00000")
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = f"A1:J{sheet.max_row}"
            for column, width in enumerate((10, 14, 26, 14, 32, 14, 14, 20, 72, 38), start=1):
                sheet.column_dimensions[get_column_letter(column)].width = width
            report.save(conflict_report_path)
            report.close()
        return ExplanationImportResult(imported, unmatched, conflicts, backup_path, conflict_report_path)
    finally:
        target.close()
