"""大集中统计系统：导数导入与标准化（CSV / XLSX）。

VBA ``InitCurData/InitPreData``（B导数比较.bas 571-675、BJ文件读取.bas）的
Python 版：流式读取、前导单引号清理、日期与数值解析、按需单位换算。
"""

from __future__ import annotations

import csv
from pathlib import Path

from .models import (
    CentralDataset,
    CentralRecord,
    ImportIssue,
    parse_central_date,
    parse_number,
)

# 导数固定 14 列表头（业务类…数据值）；读取时按位置取列。
EXPECTED_HEADERS = (
    "业务类", "数据日期", "机构类代码", "机构类名称", "地区代码", "地区名称",
    "指标顺序码", "指标代码", "指标名称", "数据属性", "币种", "频度", "批次", "数据值",
)

SUPPORTED_SUFFIXES = {".csv", ".xlsx", ".xls"}


def _clean_code(value) -> str:
    """代码字段：去前导单引号与空白（VBA Replace("'","")）。"""
    return str(value or "").replace("'", "").strip()


def _is_abnormal_number(text: str) -> bool:
    """VBA IsAbnormalData 文本分支：空串/纯非数值文本为异常；纯数值文本正常。"""
    if text == "":
        return True
    number = parse_number(text)
    if number is None:
        return True
    # val(v)==0 的文本（如“无”）在 VBA 中判为异常；此处 0.0 且原文非 0 形态视为异常。
    if number == 0.0:
        stripped = text.lstrip("+−-")
        return not (stripped.startswith(("0", ".")) )
    return False


def read_central_csv(
    path: Path,
    *,
    unit_factor: float = 1.0,
    exempt_indicators: set[str] | None = None,
) -> CentralDataset:
    """读取一期大集中导数（CSV / XLSX / XLS，表头同为固定 14 列）。

    ``unit_factor`` = GetUnitVal(源单位)/GetUnitVal(目标单位)；「单位不需要转换
    指标参照表」中的指标代码不换算。异常数据值保持原样（不换算），并记入
    issues 供运行日志使用。
    """
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise ValueError(f"大集中原始导数只支持 csv、xlsx 或 xls：{path.name}")
    if suffix == ".xlsx":
        rows = _read_xlsx_grid(path)
    elif suffix == ".xls":
        rows = _read_xls_grid(path)
    else:
        rows = _read_csv_grid(path)
    return _build_dataset(
        rows, path.name, unit_factor=unit_factor, exempt_indicators=exempt_indicators,
    )


def _read_csv_grid(path: Path) -> list[list[str]]:
    """CSV：GBK 优先解码，UTF-8 BOM 兜底，返回文本网格。"""
    raw = path.read_bytes()
    for encoding in ("gbk", "utf-8-sig"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw.decode("gbk", errors="replace")
    return [list(fields) for fields in csv.reader(text.splitlines())]


def _read_xlsx_grid(path: Path) -> list[list[str]]:
    """XLSX：openpyxl 只读取第一个表头相符的工作表；单元格规范化为文本。"""
    from openpyxl import load_workbook

    def _cell_text(value) -> str:
        if value is None:
            return ""
        if isinstance(value, float) and value.is_integer():
            # Excel 会把 20260731、批次 2 之类的整数存成浮点；转回整数字面量
            # 才能与 CSV 的文本形态一致地进入日期/数值解析。
            return str(int(value))
        return str(value)

    book = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet in book.worksheets:
            grid = [
                [_cell_text(cell) for cell in row]
                for row in sheet.iter_rows(values_only=True)
            ]
            header = [cell.strip() for cell in (grid[0] if grid else [])]
            if header[: len(EXPECTED_HEADERS)] == list(EXPECTED_HEADERS):
                return grid
    finally:
        book.close()
    raise ValueError(f"XLSX 首行表头与大集中导数 14 列不符：{path.name}")


def _read_xls_grid(path: Path) -> list[list[str]]:
    """XLS：xlrd 只读取第一个表头相符的工作表；单元格规范化为文本。

    与报表采集的 .xls 口径一致（纯读取，不启动办公套件）；日期单元格按
    工作簿 datemode 转回 ``YYYY-MM-DD`` 文本，整数浮点转回整数字面量。
    """
    try:
        import xlrd
    except ModuleNotFoundError as exc:
        raise ValueError("XLS 无法读取：缺少 xlrd>=2.0 运行依赖，请安装后重试") from exc

    def _cell_text(cell) -> str:
        if cell.ctype == 0:          # 空
            return ""
        if cell.ctype == 1:          # 文本
            return str(cell.value)
        if cell.ctype == 2:          # 数值
            number = float(cell.value)
            return str(int(number)) if number.is_integer() else str(number)
        if cell.ctype == 3:          # 日期
            moment = xlrd.xldate.xldate_as_datetime(cell.value, book.datemode)
            return moment.strftime("%Y-%m-%d")
        if cell.ctype == 4:          # 布尔
            return "TRUE" if cell.value else "FALSE"
        if cell.ctype == 5:          # 错误
            return f"#ERR{int(cell.value)}"
        return str(cell.value)

    try:
        book = xlrd.open_workbook(str(path))
    except xlrd.XLRDError as exc:
        raise ValueError(f"XLS 无法读取（非有效 Excel 97-2003 工作簿）：{path.name}；{exc}") from exc
    try:
        for sheet in book.sheets():
            grid = [
                [_cell_text(sheet.cell(r, c)) if c < sheet.ncols else "" for c in range(
                    max(sheet.ncols, len(EXPECTED_HEADERS)))]
                for r in range(sheet.nrows)
            ]
            header = [cell.strip() for cell in (grid[0] if grid else [])]
            if header[: len(EXPECTED_HEADERS)] == list(EXPECTED_HEADERS):
                return grid
    finally:
        book.release_resources()
    raise ValueError(f"XLS 首行表头与大集中导数 14 列不符：{path.name}")


def _build_dataset(
    rows: list[list[str]],
    name: str,
    *,
    unit_factor: float,
    exempt_indicators: set[str] | None,
) -> CentralDataset:
    if not rows:
        raise ValueError(f"导数文件为空：{name}")
    header = [cell.strip() for cell in rows[0]]
    if header[: len(EXPECTED_HEADERS)] != list(EXPECTED_HEADERS):
        raise ValueError(f"表头与大集中导数 14 列不符：{name}")

    exempt = exempt_indicators or set()
    dataset = CentralDataset()
    for offset, fields in enumerate(rows[1:], start=2):
        if not any(cell.strip() for cell in fields):
            continue
        fields = list(fields) + [""] * (14 - len(fields))
        record_date = parse_central_date(fields[1])
        if not record_date:
            dataset.issues.append(ImportIssue(offset, f"数据日期无法解析：{fields[1]!r}"))
        raw_value = fields[13].strip()
        indicator = _clean_code(fields[7])
        if raw_value == "":
            value = None
        else:
            number = parse_number(raw_value)
            if number is None or (number == 0.0 and _is_abnormal_number(raw_value)):
                value = raw_value          # 异常文本原样保留，不换算
                dataset.issues.append(ImportIssue(offset, f"数据值非纯数值：{raw_value!r}"))
            else:
                value = number if indicator in exempt else number * unit_factor
        record = CentralRecord(
            biz_class=fields[0].strip(),
            record_date=record_date,
            org_code=_clean_code(fields[2]),
            org_name=fields[3].strip(),
            region_code=_clean_code(fields[4]),
            region_name=fields[5].strip(),
            order_code=_clean_code(fields[6]),
            indicator=indicator,
            indicator_name=fields[8].strip(),
            data_attr=fields[9].strip(),
            currency=fields[10].strip(),
            frequency=fields[11].strip(),
            batch=fields[12].strip(),
            value=value,
        )
        dataset.records.append(record)

    dates = {record.record_date for record in dataset.records if record.record_date}
    if len(dates) > 1:
        raise ValueError(f"{name}：数据日期不一致（{sorted(dates)}）")
    dataset.record_date = next(iter(dates), "")
    dataset.build_index()
    return dataset
