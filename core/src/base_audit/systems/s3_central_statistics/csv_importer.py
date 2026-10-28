"""大集中统计系统：导数导入与标准化（CSV / XLSX）。

VBA ``InitCurData/InitPreData``（B导数比较.bas 571-675、BJ文件读取.bas）的
Python 版：流式读取、前导单引号清理、日期与数值解析、按需单位换算。
"""

from __future__ import annotations

import csv
import time
from datetime import date, datetime
from importlib.metadata import PackageNotFoundError, version as package_version
from pathlib import Path
from typing import Callable

try:  # Calamine 是可选加速器；兼容性构建仍可自动回退 OpenPyXL。
    from python_calamine import load_workbook as _calamine_load
except ImportError:  # pragma: no cover - 由打包和目标环境决定
    _calamine_load = None

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
    xlsx_reader: str = "AUTO",
    on_read: Callable[[str], None] | None = None,
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
        rows = _read_xlsx_grid(path, reader=xlsx_reader, on_read=on_read)
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


def _read_xlsx_grid(
    path: Path,
    *,
    reader: str = "AUTO",
    on_read: Callable[[str], None] | None = None,
) -> list[list[str]]:
    """XLSX：按指定读取器返回第一个表头相符的工作表，并规范化为文本。"""
    reader = str(reader or "AUTO").strip().upper()
    if reader not in {"AUTO", "CALAMINE", "OPENPYXL"}:
        raise ValueError(f"未知 XLSX 读取方式：{reader}（应为 AUTO/CALAMINE/OPENPYXL）")

    started = time.perf_counter()
    fallback_reason = ""
    if reader in {"AUTO", "CALAMINE"}:
        try:
            if _calamine_load is None:
                raise ImportError("未安装 python-calamine")
            grid = _read_xlsx_grid_calamine(path)
            engine = "Calamine"
            try:
                engine_version = package_version("python-calamine")
            except PackageNotFoundError:
                engine_version = "未知版本"
        except Exception as exc:
            if reader == "CALAMINE":
                raise ValueError(f"XLSX Calamine 读取失败：{path.name}；{exc}") from exc
            fallback_reason = f"{type(exc).__name__}: {exc}"
            grid = _read_xlsx_grid_openpyxl(path)
            engine = "OpenPyXL"
            try:
                engine_version = package_version("openpyxl")
            except PackageNotFoundError:
                engine_version = "未知版本"
    else:
        grid = _read_xlsx_grid_openpyxl(path)
        engine = "OpenPyXL"
        try:
            engine_version = package_version("openpyxl")
        except PackageNotFoundError:
            engine_version = "未知版本"

    elapsed = time.perf_counter() - started
    if on_read is not None:
        message = (
            f"{path.name}：XLSX 读取引擎={engine} {engine_version}，"
            f"耗时 {elapsed:.3f} 秒"
        )
        if fallback_reason:
            message += f"；自动回退原因：{fallback_reason}"
        on_read(message)
    return grid


def _cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, date) and not isinstance(value, datetime):
        # OpenPyXL 把 Excel 日期序列读为零点 datetime；补时间部分以一致转成文本。
        value = datetime(value.year, value.month, value.day)
    if isinstance(value, float) and value.is_integer():
        # Excel 会把日期码、批次号等整数保存成浮点；与 OpenPyXL 口径相同。
        return str(int(value))
    return str(value)


def _read_xlsx_grid_openpyxl(path: Path) -> list[list[str]]:
    from openpyxl import load_workbook

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


def _read_xlsx_grid_calamine(path: Path) -> list[list[str]]:
    """Calamine 读取 XLSX，并补回其裁掉的空白边缘，保持 OpenPyXL 行列定位。"""
    assert _calamine_load is not None
    book = _calamine_load(path)
    try:
        for sheet_name in book.sheet_names:
            sheet = book.get_sheet_by_name(sheet_name)
            matrix = sheet.to_python()
            if isinstance(sheet.start, tuple):
                row_start, col_start = sheet.start
            else:
                row_start, col_start = 0, 0
            width = max(
                int(getattr(sheet, "total_width", 0) or 0),
                col_start + max((len(row) for row in matrix), default=0),
            )
            height = max(
                int(getattr(sheet, "total_height", 0) or 0),
                row_start + len(matrix),
            )
            grid = [[""] * width for _ in range(row_start)]
            for values in matrix:
                row = [""] * col_start + [_cell_text(value) for value in values]
                row.extend([""] * max(0, width - len(row)))
                grid.append(row)
            grid.extend([[""] * width for _ in range(max(0, height - len(grid)))])
            header = [cell.strip() for cell in (grid[0] if grid else [])]
            if header[: len(EXPECTED_HEADERS)] == list(EXPECTED_HEADERS):
                return grid
    finally:
        close = getattr(book, "close", None)
        if close is not None:
            close()
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
