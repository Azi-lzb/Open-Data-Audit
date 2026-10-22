"""Institution mapping and metric headers shared by V_EXT and its exporter."""
from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal

from .errors import RetrievalError


IDENTITY_LABELS = (
    "统一社会信用代码", "法人机构统一社会信用代码", "统一信用代码", "机构代码", "机构编码",
    "机构编号", "金融机构代码", "金融机构编码",
)
TARGET_LABELS = ("报表项目", "报表名称", "项目名称", "项目编码", "行项目")
PLACEHOLDER_HEADERS = {"null", "none", "nan", "n/a", "#n/a", "-", "--"}
EXTERNAL_METADATA_HEADERS = {
    "数据日期", "报表名称", "报表代码", "版本号", "锁定信息", "报表组", "行顺序号", "行名称",
}


def normalized(value) -> str:
    return re.sub(r"\s+", "", str(value or "")).rstrip(":：").lower()


def normalize_external_date(value) -> str:
    """Normalize a full date value from ordinary or time-series external sheets."""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, bool) or value in (None, ""):
        return ""
    if isinstance(value, (int, float, Decimal)):
        try:
            if value != int(value):
                return ""
            text = str(int(value))
        except (ValueError, OverflowError):
            return ""
    else:
        text = str(value).strip()
    if text.isdecimal() and len(text) != 8:
        return ""
    for pattern in ("%Y%m%d", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"):
        try:
            return datetime.strptime(text, pattern).date().isoformat()
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(text).date().isoformat()
    except ValueError:
        return ""


@dataclass(frozen=True)
class ExternalColumn:
    sheet: str
    header_row: int
    column: int
    key_column: int
    header: str
    rows: dict[str, tuple[int, ...]]
    data_date_column: int = 0
    row_dates: dict[int, str] = field(default_factory=dict)


@dataclass
class ExternalLayout:
    mapping_sheet: str
    mapping_header_row: int
    identity_column: int
    target_column: int
    mapping_rows: list[tuple[int, str, str]]
    columns: list[ExternalColumn]
    ignored_columns: list[ExternalColumn]
    catalog: object = None
    warnings: list[dict[str, str]] = field(default_factory=list)

    def institution_entry(self, identity_keys, names=()):
        entries, method = self.catalog.match(identity_keys, names)
        if not entries:
            raise RetrievalError("IDENTITY_NOT_MAPPED", "信用代码及机构名称均未在外部文件中匹配到报表项目")
        return next(entry for entry in entries if entry.report_item), method

    def institution(self, identity_keys: set[str]) -> tuple[int, str, str]:
        if self.catalog is not None:
            entry, _ = self.institution_entry(identity_keys)
            return entry.row, entry.code, entry.report_item
        keys = {normalized(key) for key in identity_keys if normalized(key)}
        matches = [entry for entry in self.mapping_rows if normalized(entry[1]) in keys]
        if not matches:
            raise RetrievalError("IDENTITY_NOT_MAPPED", "当前机构未在外部文件映射表中找到")
        if len(matches) != 1:
            raise RetrievalError("IDENTITY_MAPPING_AMBIGUOUS", "当前机构在外部文件映射表中匹配多行")
        return matches[0]

    def resolve(self, metric: str, report_item: str, sheet: str = "", *,
                data_date: str = "", row_hint: int | None = None) -> tuple[ExternalColumn, int]:
        columns = [column for column in self.columns if normalized(column.header) == normalized(metric)
                   and (not sheet or column.sheet == sheet)]
        if not columns:
            raise RetrievalError("METRIC_NOT_FOUND", "外部文件中不存在 V_EXT 指定的指标表头")
        if len(columns) != 1:
            raise RetrievalError("METRIC_AMBIGUOUS", "外部指标表头重复，不能唯一定位；可通过 sheet= 指定工作表")
        column = columns[0]
        rows = column.rows.get(normalized(report_item), ())
        if row_hint is not None:
            rows = tuple(row for row in rows if row == row_hint)
        elif data_date and column.data_date_column:
            selected_date = normalize_external_date(data_date)
            if not selected_date:
                raise RetrievalError("EXTERNAL_DATE_INVALID", "本期报告期不是有效的完整日期，无法匹配外部时序数据")
            rows = tuple(row for row in rows if column.row_dates.get(row) == selected_date)
            if not rows:
                raise RetrievalError("EXTERNAL_DATE_NOT_FOUND", f"外部时序数据中，报表项目“{report_item}”没有本期日期 {selected_date} 的记录")
        if not rows:
            raise RetrievalError("REPORT_ITEM_NOT_FOUND", "外部指标表中不存在机构映射得到的报表项目")
        if len(rows) != 1:
            raise RetrievalError("REPORT_ITEM_DUPLICATE", "外部指标表中机构映射得到的报表项目重复")
        return column, rows[0]


def discover_external_layout(book, *, support_time_series: bool = False) -> ExternalLayout:
    # Import here to share mapping semantics without a module-level dependency
    # cycle: institutions uses the header aliases defined in this module.
    from .institutions import InstitutionCatalog, INSTITUTION_SHEETS
    catalog = InstitutionCatalog([book.path])
    entries = [entry for entry in catalog.entries if entry.report_item]
    if not entries:
        raise RetrievalError("MAPPING_SHEET_MISSING", "外部文件中未找到包含机构代码和报表项目的映射表")
    first = entries[0]
    mapping_sheet, header_row = first.sheet, first.row - 1
    identity_col, target_col = first.code_column, first.target_column
    mapping_rows = [(entry.row, entry.code, entry.report_item) for entry in entries]
    mapping_sheets = INSTITUTION_SHEETS | {"参照表"}
    targets = {normalized(item[2]) for item in mapping_rows}
    columns, ignored_columns = [], []
    layout_warnings: list[dict[str, str]] = []
    warning_keys: set[tuple[str, str, str]] = set()

    def add_warning(sheet_name: str, row_no: int, message: str) -> None:
        location = f"第{row_no}行"
        key = (sheet_name, location, message)
        if key not in warning_keys:
            warning_keys.add(key)
            layout_warnings.append({"sheet": sheet_name, "location": location, "message": message})

    def has_value(value) -> bool:
        return value is not None and not (isinstance(value, str) and not value.strip())

    for sheet in book.book:
        if sheet.title in mapping_sheets:
            continue
        time_header = None
        if support_time_series:
            for row_no in range(1, min(sheet.max_row, 20) + 1):
                headers = [normalized(sheet.cell(row_no, col).value)
                           for col in range(1, sheet.max_column + 1)]
                date_columns = [col for col, value in enumerate(headers, 1) if value == normalized("数据日期")]
                label_columns = [col for col, value in enumerate(headers, 1) if value == normalized("行名称")]
                if len(date_columns) == 1 and len(label_columns) == 1:
                    time_header = (row_no, date_columns[0], label_columns[0])
                    break
        keys = defaultdict(lambda: defaultdict(list))
        explicit_header_rows: dict[int, int] = {}
        if time_header:
            header_row, _date_col, label_col = time_header
            explicit_header_rows[label_col] = header_row
            # Keep the declared time-series columns even when every row label is
            # missing, so malformed rows become visible warnings instead of a
            # workbook-level layout failure.
            keys[label_col]
            for row in range(header_row + 1, sheet.max_row + 1):
                value = sheet.cell(row, label_col).value
                if isinstance(value, str) and normalized(value) in targets:
                    keys[label_col][normalized(value)].append(row)
        else:
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and normalized(cell.value) in targets:
                        keys[cell.column][normalized(cell.value)].append(cell.row)
            # Explicit row-label headers let us still recognize a plain report
            # when all institution rows were accidentally cleared.
            for row_no in range(1, min(sheet.max_row, 30) + 1):
                row_headers = [normalized(sheet.cell(row_no, col).value)
                               for col in range(1, sheet.max_column + 1)]
                label_cols = [col for col, value in enumerate(row_headers, 1)
                              if value in {normalized("行名称"), normalized("报表项目")}]
                if len(label_cols) == 1:
                    key_col = label_cols[0]
                    if any(row_headers[col - 1] for col in range(key_col + 1, len(row_headers) + 1)):
                        explicit_header_rows.setdefault(key_col, row_no)
                        keys[key_col]
        for key_col, item_rows in keys.items():
            first_data_row = min(
                (row for rows in item_rows.values() for row in rows),
                default=explicit_header_rows.get(key_col, 1) + 1,
            )
            candidates = [row for row in range(1, min(first_data_row, 31))
                          if any(isinstance(sheet.cell(row, col).value, str) and sheet.cell(row, col).data_type != "f"
                                 and normalized(sheet.cell(row, col).value)
                                 for col in range(key_col + 1, sheet.max_column + 1))]
            if key_col in explicit_header_rows:
                data_header_row = explicit_header_rows[key_col]
            elif not candidates:
                continue
            else:
                data_header_row = max(candidates)
            headers = [book.merged_value(sheet.title, data_header_row, col)
                       for col in range(1, sheet.max_column + 1)]
            date_columns = [col for col, raw in enumerate(headers, 1)
                            if normalized(raw) == normalized("数据日期")]
            time_series = (support_time_series and len(date_columns) == 1
                           and normalized(headers[key_col - 1]) == normalized("行名称"))
            date_column = date_columns[0] if time_series else 0
            row_dates = ({row: normalize_external_date(book.cell_raw(sheet.title, row, date_column))
                          for rows in item_rows.values() for row in rows} if time_series else {})
            metric_columns: list[tuple[int, str]] = []
            for col in range(key_col + 1, sheet.max_column + 1):
                raw = book.merged_value(sheet.title, data_header_row, col)
                if isinstance(raw, str) and raw.strip() and not raw.startswith("="):
                    if support_time_series and time_series and normalized(raw) in EXTERNAL_METADATA_HEADERS:
                        continue
                    if normalized(raw) in PLACEHOLDER_HEADERS:
                        ignored_columns.append(ExternalColumn(
                            sheet.title, data_header_row, col, key_col, raw.strip(),
                            {key: tuple(rows) for key, rows in item_rows.items()}, date_column, row_dates,
                        ))
                        continue
                    metric_columns.append((col, raw.strip()))
                    columns.append(ExternalColumn(sheet.title, data_header_row, col, key_col, raw.strip(),
                                                  {key: tuple(rows) for key, rows in item_rows.items()},
                                                  date_column, row_dates))

            # Scan through the used data region instead of stopping at a blank
            # row. Report the same row-label/date anomalies for both layouts.
            data_start = data_header_row + 1
            row_has_values: dict[int, bool] = {}
            last_nonempty_row = data_start - 1
            for row in range(data_start, sheet.max_row + 1):
                # Metadata such as data date, report code and row sequence may be
                # present even when the row label and every metric are blank.
                # It must not be mistaken for a missing report item; only a row
                # empty across the full sheet is eligible for the blank-row note.
                present = any(has_value(sheet.cell(row, col).value)
                              for col in range(1, sheet.max_column + 1))
                row_has_values[row] = present
                if present:
                    last_nonempty_row = row
            for row in range(data_start, last_nonempty_row + 1):
                raw_label = sheet.cell(row, key_col).value
                label = normalized(raw_label)
                populated_metrics = [
                    title for col, title in metric_columns if has_value(sheet.cell(row, col).value)
                ]
                if not label:
                    if populated_metrics:
                        titles = "、".join(f"“{title}”" for title in populated_metrics)
                        add_warning(
                            sheet.title, row,
                            f"行名称为空，但指标列{titles}有数据；本行无法匹配报表项目，未参与外部取数。"
                            "请检查行名称是否被误删。",
                        )
                    elif not row_has_values[row] and row < last_nonempty_row:
                        add_warning(
                            sheet.title, row,
                            "行名称及指标数据均为空，已忽略该行；如属导出空行，可忽略。",
                        )
                    continue
                if time_series and label in targets:
                    raw_date = book.cell_raw(sheet.title, row, date_column)
                    if not normalize_external_date(raw_date):
                        date_display = (
                            "空白" if raw_date is None or (isinstance(raw_date, str) and not raw_date.strip())
                            else f"“{str(raw_date).strip()}”"
                        )
                        add_warning(
                            sheet.title, row,
                            f"行名称“{str(raw_label).strip()}”对应的数据日期{date_display}为空或无效，"
                            "本行未参与时序匹配；请检查数据日期是否误删或格式错误。",
                        )
    if not columns:
        raise RetrievalError("EXTERNAL_LAYOUT_MISSING", "未找到机构报表项目行及其右侧的外部指标表头")
    return ExternalLayout(mapping_sheet, header_row, identity_col, target_col, mapping_rows,
                          columns, ignored_columns, catalog, layout_warnings)
