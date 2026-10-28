"""Resolve numeric rectangles without treating copied validation tables as data."""
from __future__ import annotations

from functools import lru_cache
from pathlib import Path
from types import SimpleNamespace

from openpyxl import load_workbook

from .errors import RetrievalError
from .layout import NamedRegion, numeric_header_regions, numeric_regions
from .source_reader import file_hashing_enabled, file_sha256, read_book


from .runtime import default_data_dir

TEMPLATE_DIR = default_data_dir() / "templates"


def _label(value):
    return "" if value is None else str(value).strip()


def _merged_label(sheet, row, column):
    for area in sheet.merged_cells.ranges:
        if area.min_row <= row <= area.max_row and area.min_col <= column <= area.max_col:
            row, column = area.min_row, area.min_col
            break
    return _label(sheet.cell(row, column).value)


def _metadata(workbook):
    regions, _ = numeric_regions(workbook)
    layouts = []
    for region in regions:
        try:
            row_axis, column_axis = numeric_header_regions(SimpleNamespace(book=workbook), region)
        except RetrievalError:
            continue
        sheet = workbook[region.sheet]
        top, left = column_axis.min_row, min(row_axis.min_col, column_axis.min_col)
        height, width = region.min_row - top, region.max_col - left + 1
        if height * width > 10000:
            continue
        pattern = tuple((r, c, _merged_label(sheet, top + r, left + c))
                        for r in range(height) for c in range(width))
        if not any(text for _, _, text in pattern):
            continue
        last_path = tuple(_merged_label(sheet, region.max_row, c)
                          for c in range(row_axis.min_col, row_axis.max_col + 1))
        layouts.append((region, top, left, height, width, pattern,
                        row_axis.min_col, row_axis.max_col, last_path))
    return frozenset(region.sheet for region in regions), tuple(layouts)


@lru_cache(maxsize=128)
def _template_metadata(path_text, size, modified):
    path = Path(path_text)
    before = file_sha256(path) if file_hashing_enabled() else ""
    workbook = read_book(path, data_only=False)
    try:
        metadata = _metadata(workbook)
    finally:
        workbook.close()
    if file_hashing_enabled() and file_sha256(path) != before:
        raise RetrievalError("INPUT_CHANGED_DURING_RUN", "读取布局模板期间文件发生变化")
    return before, metadata


def project_numeric_regions(source, schema=None, *, metadata=None, sheet_name=None):
    """Relocate a template by its complete merged header block, never by C/I offsets."""
    _, layouts = metadata if metadata is not None else _metadata(schema.book)
    result = []
    for region, top, left, height, width, pattern, first_col, last_col, last_path in layouts:
        if sheet_name and region.sheet != sheet_name:
            continue
        if region.sheet not in source.book.sheetnames:
            continue
        sheet = source.book[region.sheet]
        # Start with a nonblank anchor, then verify every label and blank in the block.
        anchor_r, anchor_c, anchor_text = next(item for item in pattern if item[2])
        candidates = []
        for row in range(1, max(1, sheet.max_row - height + 2)):
            for col in range(1, max(1, sheet.max_column - width + 2)):
                if _label(source.merged_value(region.sheet, row + anchor_r, col + anchor_c)) != anchor_text:
                    continue
                if all(_label(source.merged_value(region.sheet, row + r, col + c)) == text
                       for r, c, text in pattern):
                    candidates.append((row - top, col - left))
        if len(candidates) != 1:
            code = "NUMERIC_REGION_AMBIGUOUS" if candidates else "NUMERIC_REGION_NOT_FOUND"
            raise RetrievalError(code, "模板数字区域的完整表头无法唯一匹配源文件，请在源文件定义数字区域或更新布局模板",
                                 {"sheet": region.sheet, "region": region.name})
        dr, dc = candidates[0]
        min_row, max_row = region.min_row + dr, region.max_row + dr
        # A unique final row label follows row insertions inside the data block.
        if any(last_path):
            endings = [row for row in range(min_row, sheet.max_row + 1)
                       if tuple(_label(source.merged_value(region.sheet, row, c + dc))
                                for c in range(first_col, last_col + 1)) == last_path]
            if len(endings) == 1:
                max_row = endings[0]
        min_col, max_col = region.min_col + dc, region.max_col + dc
        result.append(NamedRegion(region.name, region.sheet,
                                  f"{region.sheet}!R{min_row}C{min_col}:R{max_row}C{max_col}",
                                  min_row, min_col, max_row, max_col))
    return result


def resolve_numeric_regions(source, sheet_name):
    """Always take region definitions from the selected or registered template."""
    if sheet_name in getattr(source, "_numeric_region_overrides", {}):
        return source._numeric_region_overrides[sheet_name], None
    explicit = getattr(source, "template_path", None)
    if explicit is not None:
        path = Path(explicit).resolve()
        stat = path.stat()
        digest, metadata = _template_metadata(str(path), stat.st_size, stat.st_mtime_ns)
        projected = project_numeric_regions(source, metadata=metadata, sheet_name=sheet_name)
        if not projected:
            raise RetrievalError("NUMERIC_REGION_NOT_FOUND", "选定模板缺少该工作表可定位的数字区域", {"sheet": sheet_name, "template": path.name})
        return projected, (path, digest)
    candidates = []
    template_dir = getattr(source, "template_dir", None) or TEMPLATE_DIR
    for path in sorted(Path(template_dir).glob("*.xls*")):
        if path.name.startswith("~$"):
            continue
        stat = path.stat()
        digest, metadata = _template_metadata(str(path), stat.st_size, stat.st_mtime_ns)
        sheets, _ = metadata
        if sheet_name in sheets and sheets.issubset(source.book.sheetnames):
            candidates.append((len(sheets), path, digest, metadata))
    if not candidates:
        return [], None  # Compatibility for workbooks without any declared layout.
    best_score = max(item[0] for item in candidates)
    best = [item for item in candidates if item[0] == best_score]
    if len(best) != 1:
        raise RetrievalError("NUMERIC_REGION_AMBIGUOUS", "匹配到多个数字区域布局模板，请在源文件定义数字区域", {"sheet": sheet_name})
    _, path, digest, metadata = best[0]
    projected = project_numeric_regions(source, metadata=metadata, sheet_name=sheet_name)
    if not projected:
        raise RetrievalError("NUMERIC_REGION_NOT_FOUND", "布局模板缺少可定位的数字区域表头", {"sheet": sheet_name})
    return projected, (path, digest)
