"""报表导出文件的身份协议；机构和表单只从文件名、工作表名识别。"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from .discovery import read_xlsx_sheet_names


_FORM_CODE = re.compile(r"(?<!\d)(20\d{3})(?!\d)")
_DEFAULT_SHEETS = {"sheet", "sheet1", "sheet2", "sheet3"}


def _form_code(name: str) -> str:
    match = _FORM_CODE.search(name)
    return match.group(1) if match else ""


@dataclass(frozen=True)
class ReportSourceIdentity:
    institution_id: str
    institution_name: str
    form_key: str
    form_code: str
    data_date: str
    recognized: bool = True


def report_source_identity(path: Path, sheet_name: str) -> ReportSourceIdentity:
    """兼容逐机构逐表、reports、banks 三种导出，不访问正文单元格。

    数据日期仅用于取数和展示，不属于两期机构/表单配对键。
    非标准文件名仅供清单按文件名和表名展示，正式导入仍明确提示。
    """
    parts = [part.strip() for part in path.stem.split("#")]
    sheet_name = sheet_name.strip()
    if len(parts) >= 4 and parts[0].casefold() == "banks":
        form = _form_code(parts[3]) or parts[3]
        # banks 文件的工作表名是机构代码，不能从中推断表单号。
        return ReportSourceIdentity(sheet_name, "", form, _form_code(parts[3]), parts[1])
    if len(parts) >= 5:
        reports = parts[0].casefold() == "reports"
        institution_id = parts[1] if reports else parts[0]
        date = parts[2] if reports else parts[1]
        fallback_form = "" if reports else parts[3]
        sheet_form = _form_code(sheet_name)
        form_key = sheet_form or (
            fallback_form if sheet_name.casefold() in _DEFAULT_SHEETS and fallback_form else sheet_name
        )
        return ReportSourceIdentity(
            institution_id, "#".join(parts[4:]), form_key,
            sheet_form or _form_code(fallback_form), date,
        )
    return ReportSourceIdentity("", path.stem, sheet_name, _form_code(sheet_name), "", False)


def report_sheet_names(path: Path) -> tuple[str, ...]:
    """配对预览只读取工作表目录；不装载工作表正文或共享字符串。"""
    try:
        if path.suffix.casefold() == ".xls":
            import xlrd

            book = xlrd.open_workbook(str(path), on_demand=True)
            try:
                names = tuple(book.sheet_names())
            finally:
                book.release_resources()
        else:
            names = read_xlsx_sheet_names(path)
        if not names:
            raise ValueError("工作表列表为空或无法读取")
        return names
    except Exception as exc:
        raise ValueError(
            f"无法读取报送文件“{path.name}”的工作表列表；请确认文件完整且为 .xls/.xlsx 格式，必要时重新导出。"
        ) from exc


def report_source_warning(current_dir: Path | None, previous_dir: Path | None) -> str:
    """按已选目录的文件名生成导出格式提醒，不读取任何工作簿内容。"""
    found = []
    for label, directory in (("本期", current_dir), ("上期", previous_dir)):
        if directory is None or not directory.is_dir():
            continue
        count = sum(
            1 for path in directory.iterdir()
            if path.is_file() and path.suffix.casefold() in {".xls", ".xlsx"}
            and path.stem.split("#", 1)[0].strip().casefold() == "banks" and len(path.stem.split("#")) >= 4
        )
        if count:
            found.append(f"{label}{count}个")
    if not found:
        return ""
    return (
        f"发现 banks 导出文件（{'、'.join(found)}）。该格式不含机构名称，名称需依赖机构参照映射。"
        "建议在来源系统选择“默认单sheet”或“按报表单sheet”重新导出，"
        "使用“机构代码#日期#01#表单#机构名称”或“reports#机构代码#日期#01#机构名称”这两类文件名。"
    )
