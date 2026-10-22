"""Detect cached Excel errors in cells used by template validation formulas."""
from __future__ import annotations

from functools import lru_cache
from pathlib import Path
import re

from openpyxl import load_workbook

from .layout import named_regions
from .source_reader import read_book, shared_readers


_EXCEL_ERRORS = {
    "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
    "#SPILL!", "#CALC!", "#GETTING_DATA",
}
_EXCEL_ERROR_TOKEN_RE = re.compile(
    r"#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|SPILL!|CALC!|GETTING_DATA)", re.I)
_EXCEL_STRING_RE = re.compile(r'"(?:[^"]|"")*"')


def formula_error_tokens(formula: str) -> list[str]:
    """Find literal Excel error tokens without mistaking quoted text for errors."""
    expression = _EXCEL_STRING_RE.sub('""', str(formula or ""))
    return list(dict.fromkeys(match.group(0).upper()
                              for match in _EXCEL_ERROR_TOKEN_RE.finditer(expression)))


@lru_cache(maxsize=64)
def _template_validation_formulas(path_text: str, size: int, modified_ns: int):
    """Cache only formulas inside named validation areas for the current template version."""
    path = Path(path_text)
    workbook = read_book(path, data_only=False)
    try:
        formulas = {}
        for region in named_regions(workbook, ("校验区域",)):
            sheet = workbook[region.sheet]
            for row in range(region.min_row, region.max_row + 1):
                for column in range(region.min_col, region.max_col + 1):
                    cell = sheet.cell(row, column)
                    if cell.data_type == "f" and isinstance(cell.value, str):
                        key = (region.sheet, cell.coordinate)
                        previous = formulas.get(key)
                        if previous is not None and previous != cell.value:
                            raise ValueError(f"校验区域对同一单元格定义了冲突公式：{region.sheet}!{cell.coordinate}")
                        formulas[key] = cell.value
        return tuple((sheet, address, formula) for (sheet, address), formula in sorted(formulas.items()))
    finally:
        workbook.close()


@lru_cache(maxsize=64)
def _template_validation_errors(path_text: str, size: int, modified_ns: int):
    """Cache static template formula errors alongside the formula inventory."""
    path = Path(path_text)
    formulas = _template_validation_formulas(path_text, size, modified_ns)
    workbook = read_book(path, data_only=True)
    try:
        errors = []
        for sheet_name, address, formula in formulas:
            tokens = formula_error_tokens(formula)
            cached_value = None
            cached_error = False
            if sheet_name in workbook.sheetnames:
                cell = workbook[sheet_name][address]
                cached_value = cell.value
                cached_error = (cell.data_type == "e" or
                                (isinstance(cached_value, str) and
                                 cached_value.strip().upper() in _EXCEL_ERRORS))
            if not tokens and not cached_error:
                continue
            errors.append((sheet_name, address, tokens[0] if tokens else str(cached_value),
                           formula, tokens))
        return tuple(errors)
    finally:
        workbook.close()


@shared_readers
def scan_validation_formula_errors(source_path: str | Path,
                                   template_paths: list[str | Path] | tuple[str | Path, ...]) -> dict:
    """Return cached formula errors in source cells corresponding to template formula cells.

    The source workbook has no trusted named ranges, so the scan area comes only
    from each layout template's ``校验区域`` names. Numeric/data cells outside
    that region are deliberately ignored.
    """
    source_path = Path(source_path).resolve()
    templates = list(dict.fromkeys(Path(path).resolve() for path in template_paths))
    if not templates:
        return {
            "status": "SKIPPED",
            "reason": "本次规则没有读取布局模板；源工作簿也不提供可信的校验区域定义",
            "template_count": 0,
            "formula_cells_scanned": 0,
            "missing_sheets": [],
            "errors": [],
        }

    formulas = {}
    for template_path in templates:
        if not template_path.is_file():
            raise FileNotFoundError(f"布局模板不存在，无法扫描校验公式：{template_path}")
        stat = template_path.stat()
        for sheet, address, formula in _template_validation_formulas(
                str(template_path), stat.st_size, stat.st_mtime_ns):
            key = (sheet, address)
            existing = formulas.get(key)
            if existing is not None and existing != formula:
                raise ValueError(f"多个布局模板对校验公式定义不一致：{sheet}!{address}")
            formulas[key] = formula

    errors_by_location = {}
    for template_path in templates:
        stat = template_path.stat()
        for sheet_name, address, value, formula, literal_errors in _template_validation_errors(
                str(template_path), stat.st_size, stat.st_mtime_ns):
            key = (sheet_name, address, value)
            entry = errors_by_location.setdefault(key, {
                "sheet": sheet_name, "cell": address, "value": value,
                "formula": formula, "source": str(source_path),
                "origin": "template_formula", "template_paths": [],
            })
            if str(template_path) not in entry["template_paths"]:
                entry["template_paths"].append(str(template_path))
            if literal_errors:
                entry["formula_error_tokens"] = list(literal_errors)

    source = read_book(source_path, data_only=True)
    try:
        errors = []
        missing_sheets = set()
        for (sheet_name, address), formula in sorted(formulas.items()):
            if sheet_name not in source.sheetnames:
                missing_sheets.add(sheet_name)
                continue
            cell = source[sheet_name][address]
            value = cell.value
            if cell.data_type == "e" or (isinstance(value, str) and value.strip().upper() in _EXCEL_ERRORS):
                key = (sheet_name, address, str(value))
                entry = errors_by_location.get(key)
                if entry is None:
                    entry = {
                        "sheet": sheet_name, "cell": address, "value": str(value),
                        "formula": formula, "source": str(source_path),
                        "origin": "source_cached_value", "template_paths": [],
                    }
                    errors_by_location[key] = entry
                elif entry.get("origin") == "template_formula":
                    entry["origin"] = "template_formula_and_source_cached_value"
        errors = list(errors_by_location.values())
        return {
            "status": "COMPLETED",
            "template_count": len(templates),
            "formula_cells_scanned": len(formulas),
            "missing_sheets": sorted(missing_sheets),
            "errors": errors,
        }
    finally:
        source.close()
