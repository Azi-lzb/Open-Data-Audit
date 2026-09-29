"""黄金基线共享工具：输出快照、规范化比较、fixture 构建。

只被 tests/golden/ 下的黄金测试使用；不修改任何生产代码。
"""

from __future__ import annotations

import hashlib
import re
import sys
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.workbook.defined_name import DefinedName

from base_audit.history import HISTORY_AUDIT_SHEET
from base_audit.name_config import initialize_config

DATA_SHEET = "20202资产负债表"
EXTERNAL_SHEET = "参照表"
NAV_SHEET = "审核导航"

# 模板规则公式：结果文本遵循 级别|校验指标|描述 三段协议。
FORMULA_LOAN = '=IF(AND(C5>C4,参照表!$B$2>0),"软性|贷款大于存款|贷款高于存款","")'
FORMULA_RESERVE = '=IF(C6<C4*0.01,"错误|拨备不足|拨备低于存款百分之一","")'
# 条件格式触发阈值（元）：乙银行存款 500 万触发、甲银行全部不触发。
CF_THRESHOLD = 4500000
CF_COMMENT = "存款超过450万元，请核实"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def excel_available() -> bool:
    """探测本机能否启动独立 Excel/WPS 会话（黄金测试的 skip 条件）。"""
    if sys.platform != "win32":
        return False
    try:
        from base_audit.excel_com import ExcelSession

        with ExcelSession(engine_preference="自动"):
            return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# 快照与规范化
# ---------------------------------------------------------------------------

_TS_RE = re.compile(r"\d{8}_?\d{6}")
_DATETIME_RE = re.compile(r"20\d{2}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}")


def normalize_value(value: Any) -> Any:
    """把单元格值规范化为可比较形态：空串→None（真空单元格）、整值浮点转 int。"""
    if value == "":
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    if isinstance(value, str):
        value = _DATETIME_RE.sub("<DATETIME>", value)
        value = _TS_RE.sub("<TS>", value)
    return value


def snapshot_workbook(path: Path) -> dict[str, dict[str, Any]]:
    """读取工作簿的 sheet 集合、表头、数据行和超链接，供黄金断言使用。"""
    book = load_workbook(path, data_only=True)
    try:
        result: dict[str, dict[str, Any]] = {}
        for sheet in book.worksheets:
            rows = [list(row) for row in sheet.iter_rows(values_only=True)]
            # 去掉尾部全空行，行列值统一规范化。
            while rows and all(v in (None, "") for v in rows[-1]):
                rows.pop()
            normalized = [
                [normalize_value(v) for v in row] for row in rows
            ]
            links = []
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.hyperlink is not None:
                        links.append({
                            "cell": cell.coordinate,
                            "target": normalize_value(cell.hyperlink.target),
                            "location": cell.hyperlink.location or "",
                        })
            result[sheet.title] = {
                "headers": normalized[0] if normalized else [],
                "rows": normalized[1:],
                "hyperlinks": links,
                "freeze": sheet.freeze_panes,
            }
        return result
    finally:
        book.close()


def glob_one(directory: Path, pattern: str) -> Path:
    """按模式取唯一文件；不存在或多个都直接失败，避免黄金断言静默漂移。"""
    matches = sorted(directory.glob(pattern))
    if len(matches) != 1:
        raise AssertionError(f"期望 {pattern} 恰好一个文件，实际 {len(matches)} 个：{matches}")
    return matches[0]


# ---------------------------------------------------------------------------
# fixture 构建（模板 / 报送文件 / 外部文件 / 配置与历史）
# ---------------------------------------------------------------------------

def _add_name(book: Workbook, name: str, sheet: str, area: str) -> None:
    book.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!{area}"))


def make_config(tmp_path: Path) -> Path:
    """生成默认配置+历史工作簿（逐笔统计系统_历史审核配置.xlsx）。"""
    config_path = tmp_path / "逐笔统计系统_历史审核配置.xlsx"
    initialize_config(config_path)
    return config_path


def write_history_row(
    config_path: Path,
    *,
    workbook_name: str,
    rule_number: str,
    feedback: str,
    opinion: str,
) -> None:
    """向“历史核查表审核”写入一行历史记录（规则编号列=完整 issue_id）。"""
    book = load_workbook(config_path)
    try:
        sheet = book[HISTORY_AUDIT_SHEET]
        sheet.append([
            workbook_name,          # 工作簿名
            DATA_SHEET,             # 工作表名
            "C5",                   # 定位单元格
            "核实",                 # 错误类型
            "贷款大于存款",          # 校验指标
            "",                     # 描述
            "",                     # 当前值
            "",                     # 对比值
            "",                     # 差值
            rule_number,            # 规则编号（历史匹配键）
            feedback,               # 历史校验说明
            opinion,                # 审核意见
        ])
        book.save(config_path)
    finally:
        book.close()


def build_audit_template(path: Path) -> None:
    """结构化模板：审核规则表 + 命名区域（校验/表结构/条件格式/汇总/表头）。"""
    book = Workbook()
    sheet = book.active
    sheet.title = DATA_SHEET
    sheet["A1"] = "资产负债表"
    # 第 3 行固定表头 = 表结构区域期望值，也是汇总流程的表头区域。
    for column, text in zip("ABC", ("指标编号", "指标名称", "本期情况")):
        sheet[f"{column}3"] = text
    for row, (code, name) in zip((4, 5, 6), (
        (20202001, "各项存款"),
        (20202002, "各项贷款"),
        (20202003, "贷款减值准备"),
    )):
        sheet[f"A{row}"] = code
        sheet[f"B{row}"] = name
    # 审核规则公式单元格（结构化模板要求公式真实存在）。
    # D4 引用外部“参照表”的 A2（联合后由公式识别发现 → 只复制该外部表）。
    sheet["D4"] = FORMULA_LOAN
    sheet["D5"] = FORMULA_RESERVE

    rules = book.create_sheet("审核规则")
    rules.append(("规则编号", "启用", "报表代码", "工作表", "公式单元格", "定位单元格", "级别", "问题说明"))
    rules.append(("R001", "是", "20202", DATA_SHEET, "D4", "C5", "核实", "贷款大于存款"))
    rules.append(("R002", "是", "20202", DATA_SHEET, "D5", "C6", "错误", "拨备不足"))

    _add_name(book, "校验区域", DATA_SHEET, "$D$4:$D$5")
    _add_name(book, "表结构区域", DATA_SHEET, "$A$3:$C$3")
    _add_name(book, "任意行汇总区域", DATA_SHEET, "$A$4:$C$6")
    _add_name(book, "表头区域", DATA_SHEET, "$A$3:$C$3")
    # native 管线从模板解析条件格式区域（COM 管线从报送文件解析）。
    _add_name(book, "条件格式区域", DATA_SHEET, "$C$4:$C$6")
    book.save(path)
    book.close()


def build_source(
    path: Path,
    *,
    deposit: float,
    loan: float,
    reserve: float,
    with_conditional_format: bool = True,
    comment: str | None = None,
) -> None:
    """生成一份报送文件：固定三列布局（A 编号 / B 名称 / C 值），第 4 行起。"""
    book = Workbook()
    sheet = book.active
    sheet.title = DATA_SHEET
    sheet["A1"] = "资产负债表"
    for column, text in zip("ABC", ("指标编号", "指标名称", "本期情况")):
        sheet[f"{column}3"] = text
    sheet["A4"], sheet["B4"], sheet["C4"] = 20202001, "各项存款", deposit
    sheet["A5"], sheet["B5"], sheet["C5"] = 20202002, "各项贷款", loan
    sheet["A6"], sheet["B6"], sheet["C6"] = 20202003, "贷款减值准备", reserve
    if with_conditional_format:
        # 报送系统的标红方式：条件格式区域内的 cellIs 规则 + 单元格批注描述。
        _add_name(book, "条件格式区域", DATA_SHEET, "$C$4:$C$6")
        red = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
        sheet.conditional_formatting.add(
            "C4:C6",
            CellIsRule(operator="greaterThan", formula=[str(CF_THRESHOLD)], fill=red),
        )
        if comment:
            sheet["C4"].comment = Comment(comment, "报送系统")
    book.save(path)
    book.close()


def build_bad_source(path: Path) -> None:
    """表头被改过的文件：表结构比对必须把它筛掉。"""
    book = Workbook()
    sheet = book.active
    sheet.title = DATA_SHEET
    sheet["A1"] = "资产负债表"
    for column, text in zip("ABC", ("指标编号", "指标名称（改）", "本期情况")):
        sheet[f"{column}3"] = text
    sheet["A4"], sheet["B4"], sheet["C4"] = 20202001, "各项存款", 100
    book.save(path)
    book.close()


def build_external(path: Path) -> None:
    """外部辅助文件：参照表（被模板公式引用，会复制）+ 未引用表（不会复制）。"""
    book = Workbook()
    sheet = book.active
    sheet.title = EXTERNAL_SHEET
    sheet["A1"] = "参照项"
    sheet["A2"] = "基准值"
    sheet["B2"] = 123
    unused = book.create_sheet("未引用表")
    unused["A1"] = "模板公式没有引用这张表，不应出现在审核副本"
    book.save(path)
    book.close()
