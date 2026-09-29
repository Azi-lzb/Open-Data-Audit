"""大集中统计系统：金融表单模板读取与渲染。

模板结构（金融表单 xlsx）：「报表清单」表（报表代码|报表名称|频度|批次）+
各报表工作表（首行表头：指标代码|指标名称|行序号|<当期数据列...>，数据行
第 1 列为指标代码）。渲染按 VBA C系统导数转表 语义：表头扩为 本期/上期/
增减额/环比 四块，按 业务类+指标代码+数据属性+币种 匹配比较结果回填；
警戒色应用到每指标组的四个单元格；空行隐藏、空表按配置删除。

性能设计：模板元数据（表清单、清理列、每行回填键计划）只在
``load_template_meta`` 计算一次、跨机构复用；每个机构只写有数据命中的
单元格——未命中的指标行整行隐藏（等价 VBA 空行隐藏），无任何命中的报表
删除（IsDeleteWs=True）或整体隐藏，避免逐格全表重复扫描。
"""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from .comparison_exporter import VBA_COLOR_PALETTE

HEADER_PREFIXES = ("上期-", "增减额-", "环比-")
# 3.3 配置工作簿同时承载金融表单模板。它们用于程序维护，不应随每家机构
# 的转表结果一并输出。
EMBEDDED_CONFIG_SHEETS = ("使用说明", "转表设置")


class FormTemplateError(ValueError):
    pass


def _sheet_biz_class(sheet_name: str) -> str:
    if "A1" in sheet_name:
        return "人民币"
    if "A2" in sheet_name:
        return "外币"
    return "本外币"


def _column_currency(header: str, biz_class: str) -> str:
    if biz_class == "人民币":
        return "人民币"
    if biz_class == "外币":
        return "美元合计"
    if "人民币" in header:
        return "人民币"
    if "外币" in header:
        return "美元合计"
    return "人民币"


def _column_attr(header: str) -> str:
    return "发生额" if "发生额" in header else "余额"


_FILL_CACHE: dict[int, PatternFill] = {}


def _fill(sheet, row: int, column: int, color_index: int) -> None:
    fill = _FILL_CACHE.get(color_index)
    if fill is None:
        hex_value = VBA_COLOR_PALETTE.get(color_index)
        if not hex_value:
            return
        fill = PatternFill("solid", fgColor=hex_value)
        _FILL_CACHE[color_index] = fill
    sheet.cell(row, column).fill = fill


def _band_color(ratio: float, alerts: list[dict]) -> int:
    from .config import alert_band_for

    band = alert_band_for(ratio, alerts)
    if band is not None:
        try:
            return int(band.get("填充颜色") or 0)
        except (TypeError, ValueError):
            return 0
    return 0


class TemplateMeta:
    """模板级元数据：表清单、清理列与每行回填键计划（跨机构只算一次）。"""

    def __init__(self, template_path: Path) -> None:
        self.template_path = Path(template_path)
        self.template_bytes = self.template_path.read_bytes()
        book = load_workbook(BytesIO(self.template_bytes), read_only=True, data_only=True)
        try:
            if "报表清单" not in book.sheetnames:
                raise FormTemplateError("金融表单模板缺少「报表清单」工作表")
            self.forms: list[dict[str, str]] = []
            for row in book["报表清单"].iter_rows(min_row=2, max_col=4, values_only=True):
                code = str(row[0] or "").strip()
                if not code or "A" not in code:
                    continue
                self.forms.append({
                    "报表代码": code,
                    "报表名称": str(row[1] or "").strip(),
                    "频度": str(row[2] or "").strip(),
                    "批次": str(row[3] or "").strip(),
                })
            form_codes = {item["报表代码"] for item in self.forms}
            # 仅同时命中 3.3 专属配置页时才视为“内嵌配置工作簿”。
            # 外部旧模板偶有单独“使用说明”页，不得误删。
            self.internal_config_sheets = (
                EMBEDDED_CONFIG_SHEETS
                if all(name in book.sheetnames for name in EMBEDDED_CONFIG_SHEETS)
                else ()
            )

            self.sheets: dict[str, dict] = {}
            for sheet_name in book.sheetnames:
                if sheet_name not in form_codes:
                    continue
                sheet = book[sheet_name]
                rows_grid = [list(row) for row in sheet.iter_rows(values_only=True)]
                if not rows_grid:
                    continue
                header = [str(v or "").strip() for v in rows_grid[0]]
                max_col = len(header)
                if max_col <= 3:
                    continue
                drop = [c for c in range(5, max_col + 1)
                        if (not header[c - 1]) or any(
                            m in header[c - 1] for m in ("剥离", "退出", "上期", "增减额", "环比"))]
                keep = [c for c in range(4, max_col + 1) if c not in drop]
                base_headers = [header[c - 1] for c in keep]
                if not base_headers:
                    continue
                biz_class = _sheet_biz_class(sheet_name)
                fill_plan: list[tuple[int, list[tuple[int, str]]]] = []
                indicator_rows: list[int] = []
                for row_index, values in enumerate(rows_grid[1:], start=2):
                    indicator = str(values[0] or "").strip() if values else ""
                    if not indicator:
                        continue
                    indicator_rows.append(row_index)
                    keys = [
                        (4 + offset,
                         f"{biz_class}{indicator}{_column_attr(header)}"
                         f"{_column_currency(header, biz_class)}")
                        for offset, header in enumerate(base_headers)
                    ]
                    fill_plan.append((row_index, keys))
                self.sheets[sheet_name] = {
                    "drop": drop,
                    "base_headers": base_headers,
                    "biz_class": biz_class,
                    "fill_plan": fill_plan,
                    "indicator_rows": indicator_rows,
                }
        finally:
            book.close()


def has_embedded_form_template(path: Path | str) -> bool:
    """判断 3.3 配置是否已内嵌可用于转表的金融表单。"""
    target = Path(path)
    if not target.is_file():
        return False
    book = load_workbook(target, read_only=True, data_only=True)
    try:
        if "报表清单" not in book.sheetnames:
            return False
        for row in book["报表清单"].iter_rows(min_row=2, max_col=1, values_only=True):
            code = str(row[0] or "").strip()
            if code and "A" in code and code in book.sheetnames:
                return True
        return False
    finally:
        book.close()


def embed_financial_template_in_config(config_path: Path | str, template_path: Path | str) -> Path:
    """将金融表单完整嵌入 3.3 配置工作簿。

    以金融表单作为输出底稿，完整保留其工作表、公式、样式和关系；仅把 3.3 的
    配置页追加到末尾。这样配置文件自身即可作为 OPENPYXL/FAST_OOXML 的
    固定转表模板，避免运行时再选择一份外部金融表单。
    """
    from copy import copy

    from openpyxl.cell.cell import Cell, MergedCell

    config_target = Path(config_path)
    source_template = Path(template_path)
    if not config_target.is_file():
        raise FileNotFoundError(f"指标比较拆分配置不存在：{config_target}")
    if not source_template.is_file():
        raise FileNotFoundError(f"金融表单模板不存在：{source_template}")

    config_book = load_workbook(config_target, read_only=False, data_only=False)
    template_book = load_workbook(source_template, read_only=False, data_only=False)
    temporary = config_target.with_name(config_target.stem + "__embedding.xlsx")
    try:
        missing = [name for name in EMBEDDED_CONFIG_SHEETS if name not in config_book.sheetnames]
        if missing:
            raise FormTemplateError(f"指标比较拆分配置缺少工作表：{'、'.join(missing)}")

        style_cache: dict[tuple[int, tuple], object] = {}
        for name in EMBEDDED_CONFIG_SHEETS:
            source = config_book[name]
            target = template_book.create_sheet(name)
            for row in source.iter_rows():
                for cell in row:
                    if isinstance(cell, MergedCell):
                        continue
                    copied = target.cell(cell.row, cell.column, cell.value)
                    if cell.has_style:
                        key = (id(source.parent), tuple(cell._style))
                        target_style = style_cache.get(key)
                        if target_style is None:
                            prototype = Cell(target, row=1, column=1)
                            prototype.font = copy(cell.font)
                            prototype.fill = copy(cell.fill)
                            prototype.border = copy(cell.border)
                            prototype.alignment = copy(cell.alignment)
                            prototype.number_format = cell.number_format
                            prototype.protection = copy(cell.protection)
                            target_style = copy(prototype._style)
                            style_cache[key] = target_style
                        copied._style = copy(target_style)
                    if cell.comment is not None:
                        copied.comment = copy(cell.comment)
                    if cell.hyperlink is not None:
                        copied._hyperlink = copy(cell.hyperlink)
            for merged_range in source.merged_cells.ranges:
                target.merge_cells(str(merged_range))
            for key, dimension in source.column_dimensions.items():
                copied_dimension = target.column_dimensions[key]
                copied_dimension.width = dimension.width
                copied_dimension.hidden = dimension.hidden
                copied_dimension.outline_level = dimension.outline_level
            for index, dimension in source.row_dimensions.items():
                copied_dimension = target.row_dimensions[index]
                copied_dimension.height = dimension.height
                copied_dimension.hidden = dimension.hidden
                copied_dimension.outline_level = dimension.outline_level
            target.freeze_panes = source.freeze_panes
            target.sheet_view.showGridLines = source.sheet_view.showGridLines
            target.sheet_properties.tabColor = source.sheet_properties.tabColor
            target.sheet_state = source.sheet_state

        template_book.save(temporary)
        template_book.close()
        config_book.close()
        temporary.replace(config_target)
        return config_target
    finally:
        # save()/close() 前抛错时，释放文件句柄并清理临时文件。
        try:
            config_book.close()
        except Exception:
            pass
        try:
            template_book.close()
        except Exception:
            pass
        if temporary.exists():
            temporary.unlink()


def output_file_name(record_date: str, org_name: str, region_name: str) -> str:
    """VBA 命名：YYYY.M 机构名 地区名.xlsx（月份去前导零）。"""
    parts = str(record_date).split("-")
    if len(parts) >= 2:
        return f"{parts[0]}.{int(parts[1])} {org_name} {region_name}.xlsx"
    return f"{org_name} {region_name}.xlsx"


def render_org_workbook(
    meta: TemplateMeta,
    *,
    org_name: str,
    region_name: str,
    record_date: str,
    data: dict[str, tuple],        # key -> (本期, 上期)
    alerts: list[dict],
    hide_empty_rows: bool = True,
    delete_empty_sheets: bool = True,
) -> tuple[object, int]:
    """为单个机构渲染一份完整模板副本；返回 (工作簿, 回填单元格数)。

    只写有数据命中的行：未命中的指标行整行隐藏（等价 VBA 空行隐藏），
    无任何命中的报表删除（IsDeleteWs=True）或整体隐藏。
    """
    book = load_workbook(BytesIO(meta.template_bytes), data_only=False)
    filled = 0

    for sheet_name in meta.internal_config_sheets:
        if sheet_name in book.sheetnames:
            book.remove(book[sheet_name])

    for sheet_name, meta_sheet in meta.sheets.items():
        sheet = book[sheet_name]
        for column in sorted(meta_sheet["drop"], reverse=True):
            sheet.delete_cols(column)
        base_headers = meta_sheet["base_headers"]
        col_count = len(base_headers)
        if col_count <= 0:
            continue

        # 命中计划：只保留本机构有数据的 (行, 列, 键)。
        matched_rows: dict[int, list[tuple[int, str]]] = {}
        for row_index, keys in meta_sheet["fill_plan"]:
            hit_keys = [(column, key) for column, key in keys if key in data]
            if hit_keys:
                matched_rows[row_index] = hit_keys

        if not matched_rows and delete_empty_sheets:
            del book[sheet_name]      # 等价 VBA 空表删除
            continue

        # 扩表头：本期块复制 3 份（上期-/增减额-/环比-）。
        # 第一块新列紧跟最后一个保留列（max_col），后续块依次后移。
        max_col = 3 + col_count
        for offset, prefix in enumerate(HEADER_PREFIXES, start=1):
            for index, header in enumerate(base_headers):
                sheet.cell(
                    1, max_col + 1 + (offset - 1) * col_count + index,
                    f"{prefix}{header}",
                )

        matched_indicator_rows = set(meta_sheet["indicator_rows"])
        filled_rows: set[int] = set()
        for row_index, hit_keys in matched_rows.items():
            filled_rows.add(row_index)
            for column, key in hit_keys:
                current_value, previous_value = data[key]
                offset = column - 4
                cur_number = float(current_value) if isinstance(current_value, (int, float)) else 0.0
                prev_number = float(previous_value) if isinstance(previous_value, (int, float)) else 0.0
                change = cur_number - prev_number
                # 只创建/格式化有值的单元格（空值单元格显示无意义，创建会
                # 显著拖慢 openpyxl 保存）；显示格式对齐 VBA（本期 0.00）。
                color_cells = []
                if current_value is not None:
                    cur_cell = sheet.cell(row_index, column)
                    cur_cell.value = current_value
                    cur_cell.number_format = "0.00"
                    color_cells.append(column)
                if previous_value is not None:
                    prev_cell = sheet.cell(row_index, 4 + col_count + offset)
                    prev_cell.value = previous_value
                    prev_cell.number_format = "0.00"
                    color_cells.append(4 + col_count + offset)
                if change != 0:
                    change_cell = sheet.cell(row_index, 4 + 2 * col_count + offset)
                    change_cell.value = change
                    change_cell.number_format = "0.00"
                    color_cells.append(4 + 2 * col_count + offset)
                if prev_number != 0 and change != 0:
                    ratio = change / prev_number
                    ratio_cell = sheet.cell(row_index, 4 + 3 * col_count + offset)
                    ratio_cell.value = ratio
                    ratio_cell.number_format = "0.00%"
                    color_cells.append(4 + 3 * col_count + offset)
                    color = _band_color(ratio, alerts)
                else:
                    # 一边有一边无：红(3)=本期无上期有，黄(6)=本期有上期无。
                    color = 6 if current_value is None else (3 if previous_value is None else 0)
                for c in color_cells:
                    _fill(sheet, row_index, c, color)
                filled += 1

        # 空行隐藏：有指标代码但没有命中任何数据的行（D..末列全空，同 VBA）。
        if hide_empty_rows:
            for row_index in matched_indicator_rows - filled_rows:
                sheet.row_dimensions[row_index].hidden = True

    if not delete_empty_sheets:
        # 空表保留但整体隐藏（VBA：IsDeleteWs=False → xlSheetHidden）。
        for sheet_name in list(book.sheetnames):
            if sheet_name in meta.sheets and sheet_name in book.sheetnames:
                sheet = book[sheet_name]
                last_data_row = meta.sheets[sheet_name]["indicator_rows"][-1] if meta.sheets[sheet_name]["indicator_rows"] else 0
                col_count = len(meta.sheets[sheet_name]["base_headers"])
                has_content = any(
                    sheet.cell(r, c).value not in (None, "")
                    for r in range(2, min(last_data_row, 5000) + 1)
                    for c in range(4, min(3 + 4 * col_count, 20) + 1)
                )
                if not has_content:
                    sheet.sheet_state = "hidden"
    return book, filled
