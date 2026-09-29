"""大集中统计系统：比较结果导出（21 列 xlsx + 警戒色 + 运行日志）。"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .comparison_engine import get_round_digits
from .models import ComparisonRow

# 黄金输出 21 列表头（2026.08 比较文件实测顺序）。
RESULT_HEADERS = (
    "业务类", "数据日期", "机构类代码", "机构类名称", "地区代码", "地区名称",
    "指标顺序码", "指标代码", "指标名称", "数据属性", "币种", "频度", "批次",
    "数据值", "上期值", "增减额({unit})", "环比", "备注", "是否说明", "说明内容", "计算过程",
)
# 人工复核专用列：与规则引擎写入的「是否说明」严格分离。
OUTPUT_EXPLANATION_HEADER = "输出说明"

# 经典 64 色 VBA colorindex 调色板（导出着色用，与 Excel Interior.ColorIndex 对应）。
VBA_COLOR_PALETTE: dict[int, str] = {
    1: "FF000000", 2: "FFFFFFFF", 3: "FFFF0000", 4: "FF00FF00", 5: "FF0000FF",
    6: "FFFFFF00", 7: "FFFF00FF", 8: "FF00FFFF", 9: "FF800000", 10: "FF008000",
    11: "FF000080", 12: "FF808000", 13: "FF800080", 14: "FF008080", 15: "FFC0C0C0",
    16: "FF808080", 17: "FF9999FF", 18: "FF993366", 19: "FFFFFFCC",
    20: "FFCCFFFF", 21: "FF660066", 22: "FFFF8080", 23: "FF0066CC",
    24: "FFCCCCFF", 25: "FF000080", 26: "FFFF00FF", 27: "FFFFFF00",
    28: "FF00FFFF", 29: "FF800080", 30: "FF800080", 31: "FF008080",
    32: "FF0000FF", 33: "FF00CCFF", 34: "FFCCFFFF", 35: "FFCCFFCC",
    36: "FFFFFF99", 37: "FF99CCFF", 38: "FFFF99CC", 39: "FFCC99FF",
    40: "FFFFCC99", 41: "FF3366FF", 42: "FF33CCCC", 43: "FF99CC00",
    44: "FFFFCC00", 45: "FFFF9900", 46: "FFFF6600", 47: "FF808080",
    48: "FF969696", 49: "FF808000", 50: "FF00CC99", 51: "FF003399",
    52: "FF666699", 53: "FF808080", 54: "FF993366", 55: "FF333300",
}


def colorindex_to_hex(index: int | None) -> str | None:
    if not index:
        return None
    return VBA_COLOR_PALETTE.get(int(index))


def result_sheet_name(record_date: str) -> str:
    """表名「{年}年{月}月对比结果」（VBA 2993 行口径，去前导零月份）。"""
    if record_date:
        parts = record_date.split("-")
        if len(parts) >= 2:
            return f"{parts[0]}年{int(parts[1])}月对比结果"
    return "对比结果"


def write_comparison_workbook(
    rows: list[ComparisonRow],
    output_path: Path,
    *,
    target_unit: str = "亿元",
    sheet_date: str = "",
) -> Path:
    """写出 21 列比较结果工作簿（含警戒色与数字格式）。"""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.title = result_sheet_name(sheet_date)
    headers = [header.format(unit=target_unit) for header in RESULT_HEADERS] + [OUTPUT_EXPLANATION_HEADER]
    sheet.append(headers)
    digits = get_round_digits(target_unit)
    fill_cache: dict[int, PatternFill] = {}

    def fill_for(color_index: int) -> PatternFill:
        if color_index not in fill_cache:
            hex_value = colorindex_to_hex(color_index) or "FFFFFF00"
            fill_cache[color_index] = PatternFill("solid", fgColor=hex_value)
        return fill_cache[color_index]

    def display(value) -> object:
        if value is None:
            return None
        if isinstance(value, float):
            if value == int(value) and abs(value) < 1e15:
                return int(value)
            return value
        return value

    for row in rows:
        record = row.record
        sheet.append((
            record.biz_class,
            display(record.record_date),
            record.org_code,
            record.org_name,
            record.region_code,
            record.region_name,
            row.order_code_out or record.indicator,
            record.indicator,
            row.indicator_name_out or record.indicator_name,
            record.data_attr,
            record.currency,
            record.frequency,
            display(record.batch),
            display(record.value),
            display(row.prev_value),
            display(row.change),
            display(row.ratio),
            row.remark,
            row.need_explain,
            row.explain_content,
            row.process,
        ))
        if row.fill_color and row.fill_scope:
            hex_value = colorindex_to_hex(row.fill_color)
            if hex_value:
                fill = fill_for(row.fill_color)
                if row.fill_scope == "row":
                    for column in range(1, 19):
                        sheet.cell(sheet.max_row, column).fill = fill
                else:
                    sheet.cell(sheet.max_row, 18).fill = fill

    # 数字格式与表头样式（VBA 457-462/2984-2988）。
    for row_index in range(2, sheet.max_row + 1):
        for column in (14, 15, 16):
            sheet.cell(row_index, column).number_format = "0.00"
        sheet.cell(row_index, 17).number_format = "0.00%"
        sheet.cell(row_index, 2).number_format = "yyyy/m/d"
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(2, sheet.max_row)}"
    widths = (10, 12, 10, 24, 10, 10, 12, 10, 32, 9, 9, 6, 6, 14, 14, 14, 12, 24, 24, 24, 40, 12)
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    book.save(output_path)
    book.close()
    return output_path


def write_log_workbook(log_rows: list[tuple], output_path: Path, title: tuple[str, ...]) -> Path:
    """运行日志工作簿：一行一条（时间戳 + 消息列）。"""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.title = "运行日志"
    sheet.append(list(title))
    for row in log_rows:
        sheet.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), *row])
    book.save(output_path)
    book.close()
    return output_path
