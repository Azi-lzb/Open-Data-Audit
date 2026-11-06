"""Native workbook operations for the UOS/Linux edition.

This module deliberately has no COM, Wine, or Windows WPS dependency.  It
implements the part of the audit pipeline that is safe to do with OOXML:
named-range discovery, copying template formulas to an audit copy, importing
auxiliary sheets and reading LibreOffice-saved formula caches.
"""

from __future__ import annotations

from copy import copy, deepcopy
from datetime import datetime
import posixpath
from pathlib import Path
import re
from typing import Any, Dict, Iterable
import zipfile
from xml.etree import ElementTree as ET

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple, column_index_from_string, range_boundaries

from ..models import AuditRule, CopyRange, Issue, TemplateDefinition
from ..name_config import FeatureMapping, matches_named_range
from ..template import (
    REQUIRED_HEADERS,
    TemplateError,
    clean_rule_comment,
    normalize_template_name,
    parse_comment_rule_id,
    parse_formula_result,
    parse_rule_rows,
    is_negligible_zero_difference_result,
)


RESULT_KEYWORDS = ("错误", "硬性", "错", "软性")
FORMULA_ERRORS = (
    "#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!",
    "#SPILL!", "#CALC!", "#FIELD!", "#GETTING_DATA",
)

_OOXML_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_OOXML_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_OOXML_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _normalise_address(address: str) -> str:
    return address.replace("$", "").strip()


def _all_defined_names(workbook: Any) -> list[Any]:
    """Return workbook- and worksheet-scoped Excel names.

    Recent openpyxl versions keep local names on ``Worksheet.defined_names``
    instead of exposing all of them from ``Workbook.defined_names``.  Excel's
    name manager displays both, and the template protocol deliberately allows
    either scope, so resolving only the latter produces a false “missing named
    range” error for valid templates.
    """
    result = list(workbook.defined_names.values())
    for sheet in workbook.worksheets:
        result.extend(sheet.defined_names.values())
    return result


def _issue_id(source_file: Path | str, rule: AuditRule, indicator: str = "", check_type: str = "") -> str:
    """位置身份；校验类型非空时编入定位单元格与校验指标之间（与 COM 侧一致）。"""
    source_name = normalize_template_name(Path(str(source_file)).stem)
    parts = [
        source_name or "工作簿未识别", rule.sheet_name or "工作表未识别",
        rule.formula_cell or "公式单元格未识别",
    ]
    if check_type:
        parts.append(check_type)
    parts.append(indicator or "校验指标未识别")
    return "｜".join(parts)


def _formula_error(value: object) -> str:
    text = str(value or "").strip().upper()
    return next((item for item in FORMULA_ERRORS if text.startswith(item)), "")


def _has_result_marker(value: object) -> bool:
    return str(value or "").strip().startswith(RESULT_KEYWORDS)


def _in_range(cell: str, address: str) -> bool:
    try:
        row, column = coordinate_to_tuple(cell.replace("$", ""))
        min_col, min_row, max_col, max_row = range_boundaries(_normalise_address(address))
    except ValueError:
        return False
    return min_row <= row <= max_row and min_col <= column <= max_col


def template_formulas(template_path: Path, *, workbook: Any | None = None) -> list[object]:
    """收集模板全部公式文本；调用方可传入已打开工作簿避免重复解析。"""
    owns_workbook = workbook is None
    book = workbook or load_workbook(template_path, read_only=True, data_only=False, keep_links=False)
    try:
        return [
            cell.value for sheet in book.worksheets for row in sheet.iter_rows() for cell in row
            if isinstance(cell.value, str) and cell.value.startswith("=")
        ]
    finally:
        if owns_workbook:
            book.close()


def named_ranges(workbook: Any, mappings: Iterable[FeatureMapping]) -> list[CopyRange]:
    """Resolve configured Excel names to individual areas.

    Both workbook- and worksheet-scoped names are supported.  Non-range
    compatibility names such as ``_xlfn.SUMIFS`` have no destinations and are
    intentionally ignored.
    """
    mappings = tuple(mappings)
    result: list[CopyRange] = []
    seen: set[tuple[str, str]] = set()
    for defined_name in _all_defined_names(workbook):
        if not any(matches_named_range(mapping, defined_name.name) for mapping in mappings):
            continue
        try:
            destinations = tuple(defined_name.destinations)
        except (AttributeError, ValueError):
            continue
        for sheet_name, address in destinations:
            address = _normalise_address(address)
            if sheet_name not in workbook.sheetnames:
                continue
            try:
                range_boundaries(address)
            except ValueError:
                continue
            item = CopyRange(sheet_name, address)
            identity = (item.sheet_name, item.address.upper())
            if identity not in seen:
                result.append(item)
                seen.add(identity)
    return result


def require_named_ranges(workbook: Any, mapping: FeatureMapping) -> list[CopyRange]:
    missing: list[str] = []
    resolved: list[CopyRange] = []
    for range_name in mapping.range_names:
        expected = FeatureMapping(
            mapping.name, mapping.feature_type, (range_name,), mapping.workbook_limited,
            mapping.workbook_keyword, mapping.remark, mapping.output, mapping.input_note,
        )
        areas = named_ranges(workbook, (expected,))
        if areas:
            resolved.extend(areas)
        else:
            missing.append(range_name)
    if missing:
        raise TemplateError("模块“{}”缺少命名区域：{}".format(mapping.name, "、".join(missing)))
    return resolved


def named_single_cells_with_prefix(workbook: Any, prefix: str) -> list[tuple[str, CopyRange]]:
    """Return one-cell names such as ``全局单元格区域.数据日期``.

    This is used by the two row-summary modules to bring reporting metadata
    into every emitted record.  Names that do not resolve to exactly one cell
    are rejected rather than guessed.
    """
    result: list[tuple[str, CopyRange]] = []
    for defined_name in _all_defined_names(workbook):
        simple = defined_name.name.rsplit("!", 1)[-1]
        if not simple.casefold().startswith(prefix.casefold()):
            continue
        field_name = simple[len(prefix):].strip()
        if not field_name:
            continue
        try:
            destinations = tuple(defined_name.destinations)
        except (AttributeError, ValueError):
            continue
        for sheet_name, address in destinations:
            address = _normalise_address(address)
            try:
                min_col, min_row, max_col, max_row = range_boundaries(address)
            except ValueError:
                continue
            if min_col != max_col or min_row != max_row:
                raise ValueError("命名区域“{}”必须只定位一个单元格".format(simple))
            result.append((field_name, CopyRange(sheet_name, address)))
    return result


def template_structure_values(
    template_path: Path,
    definition: TemplateDefinition,
    *,
    workbook: Any | None = None,
) -> list[tuple[CopyRange, list[list[Any]]]]:
    """Read fixed labels; callers may reuse an already-open template workbook."""
    # 旧模板把稳定规则编号写在公式单元格批注中。ReadOnlyCell 不公开
    # comment，因此这里必须以普通模式读取模板；模板每个流程仅读取一次，
    # 批量报送文件仍使用流式/批量读取，不进入这条路径。
    owns_workbook = workbook is None
    book = workbook or load_workbook(template_path, read_only=False, data_only=False, keep_links=False)
    try:
        result: list[tuple[CopyRange, list[list[Any]]]] = []
        for item in definition.structure_ranges:
            if item.sheet_name not in book.sheetnames:
                raise TemplateError("模板缺少表结构工作表：{}".format(item.sheet_name))
            sheet = book[item.sheet_name]
            min_col, min_row, max_col, max_row = range_boundaries(item.address)
            result.append((item, [list(row) for row in sheet.iter_rows(
                min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col,
                values_only=True,
            )]))
        return result
    finally:
        if owns_workbook:
            book.close()


def read_template(
    template_path: Path,
    *,
    formula_mappings: Iterable[FeatureMapping] = (),
    structure_mappings: Iterable[FeatureMapping] = (),
    extraction_mappings: Iterable[FeatureMapping] = (),
    workbook: Any | None = None,
) -> TemplateDefinition:
    """Read a template; callers may reuse an already-open workbook."""
    formula_mappings = tuple(formula_mappings)
    structure_mappings = tuple(structure_mappings)
    extraction_mappings = tuple(extraction_mappings)
    # 旧模板把稳定规则编号写在公式单元格批注中。ReadOnlyCell 不公开
    # comment，因此这里必须以普通模式读取模板；模板每个流程仅读取一次，
    # 批量报送文件仍使用流式/批量读取，不进入这条路径。
    owns_workbook = workbook is None
    book = workbook or load_workbook(template_path, read_only=False, data_only=False, keep_links=False)
    try:
        structure_ranges = named_ranges(book, structure_mappings)
        extraction_ranges = named_ranges(book, extraction_mappings)
        if "审核规则" in book.sheetnames:
            sheet = book["审核规则"]
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                raise TemplateError("模板“审核规则”工作表为空")
            headers = [str(value).strip() if value is not None else "" for value in rows[0]]
            missing = [name for name in REQUIRED_HEADERS if name not in headers]
            if missing:
                raise TemplateError("审核规则表缺少字段：" + "、".join(missing))
            records = [dict(zip(headers, values)) for values in rows[1:]]
            rules = parse_rule_rows(records)
            errors = []
            for rule in rules:
                if rule.enabled and (
                    rule.sheet_name not in book.sheetnames
                    or not str(book[rule.sheet_name][rule.formula_cell].value or "").startswith("=")
                ):
                    errors.append("{}：{}!{}没有公式".format(rule.rule_id, rule.sheet_name, rule.formula_cell))
            if errors:
                raise TemplateError("；".join(errors))
            return TemplateDefinition(
                rules=rules,
                copy_ranges=[CopyRange(rule.sheet_name, rule.copy_range) for rule in rules if rule.enabled],
                structured=True,
                structure_ranges=structure_ranges,
                extraction_ranges=extraction_ranges,
            )

        copy_ranges = named_ranges(book, formula_mappings)
        # 条件格式提取、表结构比对等流程只需要模板定位区域，并不要求
        # 模板同时具备“校验区域”或任何公式。
        if formula_mappings and not copy_ranges:
            raise TemplateError("公式校验功能对应的命名区域中没有可用区域")
        rules: list[AuditRule] = []
        seen: set[tuple[str, str]] = set()
        seen_rule_ids: dict[str, tuple[str, str]] = {}
        for item in copy_ranges:
            sheet = book[item.sheet_name]
            min_col, min_row, max_col, max_row = range_boundaries(item.address)
            for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                for cell in row:
                    formula = str(cell.value or "")
                    if not formula.startswith("=") or (item.sheet_name, cell.coordinate) in seen:
                        continue
                    seen.add((item.sheet_name, cell.coordinate))
                    raw_comment = str(cell.comment.text or "") if cell.comment else ""
                    explicit_rule_id = parse_comment_rule_id(raw_comment)
                    rule_id = explicit_rule_id or "{}-{}".format(item.sheet_name, cell.coordinate)
                    identity = (item.sheet_name, cell.coordinate)
                    previous = seen_rule_ids.get(rule_id)
                    if previous and previous != identity:
                        raise TemplateError(
                            "规则编号“{}”重复：{}!{}、{}!{}".format(
                                rule_id, previous[0], previous[1], item.sheet_name, cell.coordinate,
                            )
                        )
                    seen_rule_ids[rule_id] = identity
                    rules.append(AuditRule(
                        rule_id=rule_id, enabled=True,
                        report_code=normalize_template_name(template_path.stem), sheet_name=item.sheet_name,
                        formula_cell=cell.coordinate, target_cell=cell.coordinate, severity="错误",
                        message=clean_rule_comment(raw_comment),
                        copy_range=item.address,
                        result_mode="keyword" if any(marker in formula for marker in RESULT_KEYWORDS) else "error_only",
                        value_from_result=True,
                    ))
        if formula_mappings and not rules:
            raise TemplateError("公式校验功能对应的命名区域中没有公式")
        return TemplateDefinition(rules, copy_ranges, False, structure_ranges, extraction_ranges)
    finally:
        if owns_workbook:
            book.close()


def copy_formula_ranges(template_path: Path, audit_path: Path, definition: TemplateDefinition) -> None:
    """Copy cell formulas and visible styles to a prepared audit copy."""
    template = load_workbook(template_path, data_only=False, keep_links=False)
    try:
        copy_formula_ranges_from_template(template, audit_path, definition)
    finally:
        template.close()


def copy_formula_ranges_from_template(template: Any, audit_path: Path, definition: TemplateDefinition) -> None:
    """Copy formulas from one already-open template into a single audit copy.

    批量审核可复用同一个模板对象，避免每份报送副本重复解析相同模板 XML；
    审核副本仍各自独立打开、保存，不共享任何可变状态。
    """
    audit = load_workbook(audit_path, data_only=False, keep_links=False)
    try:
        missing = sorted({rule.sheet_name for rule in definition.rules if rule.enabled} - set(audit.sheetnames))
        if missing:
            raise TemplateError("报送文件缺少工作表：" + "、".join(missing))
        copied: set[tuple[str, str]] = set()
        for item in definition.copy_ranges:
            identity = (item.sheet_name, item.address.upper())
            if identity in copied:
                continue
            copied.add(identity)
            if item.sheet_name not in audit.sheetnames:
                raise TemplateError("报送文件缺少工作表：" + item.sheet_name)
            source_sheet, target_sheet = template[item.sheet_name], audit[item.sheet_name]
            min_col, min_row, max_col, max_row = range_boundaries(item.address)
            for row in range(min_row, max_row + 1):
                for column in range(min_col, max_col + 1):
                    source = source_sheet.cell(row, column)
                    target = target_sheet.cell(row, column)
                    target.value = source.value
                    if source.has_style:
                        target._style = copy(source._style)
                    if source.number_format:
                        target.number_format = source.number_format
                    if source.alignment:
                        target.alignment = copy(source.alignment)
                    if source.protection:
                        target.protection = copy(source.protection)
                    if source.comment:
                        target.comment = copy(source.comment)
        if definition.structured:
            if "审核规则" in audit.sheetnames:
                del audit["审核规则"]
            _copy_sheet(template["审核规则"], audit, "审核规则")
        audit.calculation.fullCalcOnLoad = True
        audit.calculation.forceFullCalc = True
        audit.save(audit_path)
    finally:
        audit.close()


def _copy_sheet(source: Any, target_book: Any, target_name: str,
                style_cache: Dict[tuple, Any] | None = None) -> None:
    """跨簿复制工作表：值/公式/样式/批注/超链接/合并/行列/条件格式/数据验证。

    跨簿样式编号不同，必须按原型逐项登记到目标工作簿；直接搬 ``_style`` 会让
    目标簿引用不存在的字体/对齐索引，保存后读回报 ``IndexError``。
    """
    from openpyxl.cell.cell import Cell, MergedCell

    target = target_book.create_sheet(target_name)
    target.sheet_state = source.sheet_state
    target.freeze_panes = source.freeze_panes
    target.sheet_view.showGridLines = source.sheet_view.showGridLines
    target.sheet_properties.tabColor = source.sheet_properties.tabColor
    cache = {} if style_cache is None else style_cache
    for row in source.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            destination = target.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                key = (id(source.parent), tuple(cell._style))
                target_style = cache.get(key)
                if target_style is None:
                    prototype = Cell(target, row=1, column=1)
                    prototype.font = copy(cell.font)
                    prototype.fill = copy(cell.fill)
                    prototype.border = copy(cell.border)
                    prototype.alignment = copy(cell.alignment)
                    prototype.number_format = cell.number_format
                    prototype.protection = copy(cell.protection)
                    target_style = copy(prototype._style)
                    cache[key] = target_style
                destination._style = copy(target_style)
            if cell.comment:
                destination.comment = copy(cell.comment)
            if cell.hyperlink:
                destination._hyperlink = copy(cell.hyperlink)
    for key, dimension in source.row_dimensions.items():
        item = target.row_dimensions[key]
        item.height = dimension.height
        item.hidden = dimension.hidden
        item.outline_level = dimension.outline_level
    for key, dimension in source.column_dimensions.items():
        item = target.column_dimensions[key]
        item.width = dimension.width
        item.hidden = dimension.hidden
        item.outline_level = dimension.outline_level
    for merged in source.merged_cells.ranges:
        target.merge_cells(str(merged))
    for rng in source.conditional_formatting:
        for rule in rng.rules:
            target.conditional_formatting.add(str(rng.sqref), copy(rule))
    for validation in source.data_validations.dataValidation:
        target.add_data_validation(copy(validation))


def _workbook_worksheet_parts(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    by_id = {
        item.attrib.get("Id", ""): item
        for item in relationships.findall(f"{{{_OOXML_PACKAGE_REL_NS}}}Relationship")
    }
    result: dict[str, str] = {}
    sheets = workbook.find(f"{{{_OOXML_MAIN_NS}}}sheets")
    for sheet in sheets if sheets is not None else ():
        relation_id = sheet.attrib.get(f"{{{_OOXML_REL_NS}}}id", "")
        relation = by_id.get(relation_id)
        if relation is None or relation.attrib.get("TargetMode") == "External":
            continue
        if not relation.attrib.get("Type", "").endswith("/worksheet"):
            continue
        target = relation.attrib.get("Target", "")
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        result[sheet.attrib.get("name", "")] = part
    return result


def _sheet_data_fragment(raw: bytes) -> bytes | None:
    """Return the exact worksheet ``sheetData`` XML fragment, if present."""
    start = re.search(
        rb"<(?P<name>(?:[A-Za-z_][\w.-]*:)?sheetData)\b[^>]*(?:/>|>)",
        raw,
    )
    if start is None:
        return None
    if start.group(0).rstrip().endswith(b"/>"):
        return start.group(0)
    closing = b"</" + start.group("name") + b">"
    end = raw.find(closing, start.end())
    if end < 0:
        raise TemplateError("工作表 XML 的 sheetData 未闭合")
    return raw[start.start():end + len(closing)]


def _has_cell_payload(cell: ET.Element) -> bool:
    formula = cell.find(f"{{{_OOXML_MAIN_NS}}}f")
    value = cell.find(f"{{{_OOXML_MAIN_NS}}}v")
    inline = cell.find(f"{{{_OOXML_MAIN_NS}}}is")
    return (
        formula is not None
        or (value is not None and value.text is not None)
        or (inline is not None and "".join(inline.itertext()) != "")
    )


def _inline_shared_string_cell(cell: ET.Element, shared_strings: list[ET.Element]) -> ET.Element:
    """Convert a copied shared-string cell to inline text for the new package."""
    if cell.find(f"{{{_OOXML_MAIN_NS}}}f") is not None:
        raise TemplateError("合并区域内部公式使用共享字符串缓存，无法安全迁移：" + cell.attrib.get("r", ""))
    value = cell.find(f"{{{_OOXML_MAIN_NS}}}v")
    try:
        index = int(value.text or "") if value is not None else -1
        item = shared_strings[index]
    except (ValueError, IndexError, TypeError):
        raise TemplateError("合并区域内部共享字符串索引无效：" + cell.attrib.get("r", ""))
    if value is not None:
        cell.remove(value)
    inline = ET.Element(f"{{{_OOXML_MAIN_NS}}}is")
    inline.extend(deepcopy(list(item)))
    if not list(inline):
        ET.SubElement(inline, f"{{{_OOXML_MAIN_NS}}}t")
    cell.attrib["t"] = "inlineStr"
    cell.append(inline)
    return cell


def _restore_merged_cell_payloads(original_path: Path, generated_path: Path, output_path: Path) -> None:
    """Restore meaningful XML cells hidden inside merged ranges.

    openpyxl discards non-anchor cells inside merges. Preserve the rest of its
    normalized worksheet package, and overlay only original merged-range cell
    payloads that openpyxl cannot represent. Shared strings become inline
    strings so the output continues to use openpyxl's generated string table.
    """
    with zipfile.ZipFile(original_path, "r") as original, zipfile.ZipFile(generated_path, "r") as generated:
        original_sheets = _workbook_worksheet_parts(original)
        generated_sheets = _workbook_worksheet_parts(generated)
        generated_part_by_original_part: dict[str, str] = {}
        for sheet_name, original_part in original_sheets.items():
            generated_part = generated_sheets.get(sheet_name)
            if generated_part is None:
                raise TemplateError("openpyxl 外部表添加后，原工作表丢失：" + sheet_name)
            generated_part_by_original_part[generated_part] = original_part

        shared_strings: list[ET.Element] = []
        if "xl/sharedStrings.xml" in original.namelist():
            shared_root = ET.fromstring(original.read("xl/sharedStrings.xml"))
            shared_strings = shared_root.findall(f"{{{_OOXML_MAIN_NS}}}si")

        replacements: dict[str, bytes] = {}
        for generated_part, original_part in generated_part_by_original_part.items():
            original_raw = original.read(original_part)
            original_root = ET.fromstring(original_raw)
            original_data = original_root.find(f"{{{_OOXML_MAIN_NS}}}sheetData")
            merges = original_root.find(f"{{{_OOXML_MAIN_NS}}}mergeCells")
            if original_data is None or merges is None:
                continue

            merge_ranges = []
            for merged in merges:
                try:
                    min_col, min_row, max_col, max_row = range_boundaries(merged.attrib.get("ref", ""))
                except ValueError:
                    continue
                merge_ranges.append((min_col, min_row, max_col, max_row))
            if not merge_ranges:
                continue

            row_numbers = {int(row.attrib.get("r", "0") or 0) for row in original_data}
            merged_by_row = {
                row_number: [
                    (min_col, max_col, min_row, min_col)
                    for min_col, min_row, max_col, max_row in merge_ranges
                    if min_row <= row_number <= max_row
                ]
                for row_number in row_numbers
            }
            additions: dict[int, list[ET.Element]] = {}
            for source_row in original_data:
                row_number = int(source_row.attrib.get("r", "0") or 0)
                spans = merged_by_row.get(row_number, ())
                if not spans:
                    continue
                for source_cell in source_row.findall(f"{{{_OOXML_MAIN_NS}}}c"):
                    address = source_cell.attrib.get("r", "")
                    if not address or not _has_cell_payload(source_cell):
                        continue
                    cell_row, cell_column = coordinate_to_tuple(address)
                    if not any(
                        min_col <= cell_column <= max_col
                        and (cell_row, cell_column) != (anchor_row, anchor_column)
                        for min_col, max_col, anchor_row, anchor_column in spans
                    ):
                        continue
                    restored = deepcopy(source_cell)
                    # openpyxl may have renumbered styles and metadata. Cell
                    # values/formulas are the required payload; use default
                    # formatting for hidden non-anchor merged cells.
                    for attribute in ("s", "cm", "vm"):
                        restored.attrib.pop(attribute, None)
                    if restored.attrib.get("t") == "s":
                        restored = _inline_shared_string_cell(restored, shared_strings)
                    additions.setdefault(cell_row, []).append(restored)

            if not additions:
                continue
            generated_raw = generated.read(generated_part)
            generated_root = ET.fromstring(generated_raw)
            generated_data = generated_root.find(f"{{{_OOXML_MAIN_NS}}}sheetData")
            if generated_data is None:
                raise TemplateError("openpyxl 外部表添加后，原工作表数据区域丢失：" + Path(generated_part).name)
            generated_rows = {
                int(row.attrib.get("r", "0") or 0): row
                for row in generated_data.findall(f"{{{_OOXML_MAIN_NS}}}row")
            }
            for row_number, cells in additions.items():
                row = generated_rows.get(row_number)
                if row is None:
                    row = ET.Element(f"{{{_OOXML_MAIN_NS}}}row", {"r": str(row_number)})
                    generated_data.append(row)
                    generated_rows[row_number] = row
                for cell in cells:
                    address = cell.attrib.get("r", "")
                    existing = next((item for item in row.findall(f"{{{_OOXML_MAIN_NS}}}c") if item.attrib.get("r") == address), None)
                    if existing is not None:
                        row.remove(existing)
                    row.append(cell)
                row[:] = sorted(
                    row,
                    key=lambda item: column_index_from_string(
                        "".join(character for character in item.attrib.get("r", "A1") if character.isalpha()) or "A"
                    ) if item.tag == f"{{{_OOXML_MAIN_NS}}}c" else 0,
                )
            generated_data[:] = sorted(
                generated_data,
                key=lambda item: int(item.attrib.get("r", "0") or 0),
            )
            ET.register_namespace("", _OOXML_MAIN_NS)
            generated_fragment = ET.tostring(generated_data, encoding="utf-8")
            old_fragment = _sheet_data_fragment(generated_raw)
            if old_fragment is None:
                raise TemplateError("无法定位 openpyxl 输出的工作表数据区域：" + Path(generated_part).name)
            start = generated_raw.find(old_fragment)
            replacements[generated_part] = (
                generated_raw[:start] + generated_fragment + generated_raw[start + len(old_fragment):]
            )

        with zipfile.ZipFile(output_path, "w") as restored:
            written_parts: set[str] = set()
            for info in generated.infolist():
                payload = replacements.get(info.filename)
                if payload is None:
                    payload = generated.read(info.filename)
                restored.writestr(info, payload)
                written_parts.add(info.filename)
            for name, payload in replacements.items():
                if name not in written_parts:
                    restored.writestr(name, payload, compress_type=zipfile.ZIP_DEFLATED)


def add_external_sheets(audit_path: Path, external_path: Path, sheet_names: Iterable[str]) -> None:
    """Add external sheets while retaining values hidden inside target merges."""
    audit_path = Path(audit_path)
    openpyxl_temp = audit_path.with_name(audit_path.stem + ".external-openpyxl.tmp" + audit_path.suffix)
    restored_temp = audit_path.with_name(audit_path.stem + ".external-preserved.tmp" + audit_path.suffix)
    try:
        audit = load_workbook(audit_path, data_only=False, keep_links=False)
        external = None
        try:
            external = load_workbook(external_path, data_only=False, keep_links=False)
            names = tuple(sheet_names)
            missing = [name for name in names if name not in external.sheetnames]
            conflicts = [name for name in names if name in audit.sheetnames]
            if missing:
                raise TemplateError("外部文件缺少工作表：" + "、".join(missing))
            if conflicts:
                raise TemplateError("外部文件工作表与报送文件重名，不能覆盖原表：" + "、".join(conflicts))
            style_cache: Dict[tuple, Any] = {}
            for name in names:
                _copy_sheet(external[name], audit, name, style_cache)
            audit.save(openpyxl_temp)
        finally:
            audit.close()
            if external is not None:
                external.close()

        _restore_merged_cell_payloads(audit_path, openpyxl_temp, restored_temp)
        restored_temp.replace(audit_path)
    finally:
        openpyxl_temp.unlink(missing_ok=True)
        restored_temp.unlink(missing_ok=True)


def extract_issues(
    audit_path: Path,
    rules: Iterable[AuditRule],
    *, period: str, batch_id: str, org_code: str = "", org_name: str = "", source_file: Path | None = None,
    extraction_ranges: Iterable[CopyRange] = (),
    cached_values: dict[tuple[str, str], Any] | None = None,
) -> list[Issue]:
    """Build issue objects from saved caches or a live spreadsheet value map.

    ``cached_values`` is supplied by the Windows COM adapter during a batch
    run.  It avoids opening the sizeable audit XLSX again only to read a small
    set of configured formula/target cells.
    """
    source_file = source_file or audit_path
    allowed = tuple(extraction_ranges)
    enabled = [rule for rule in rules if rule.enabled and (
        not allowed or any(item.sheet_name == rule.sheet_name and _in_range(rule.formula_cell, item.address) for item in allowed)
    )]
    book = None
    try:
        audit_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        result: list[Issue] = []
        if cached_values is None:
            cached_values = _read_issue_value_blocks(audit_path, enabled)
        for rule in enabled:
            formula_result = cached_values.get((rule.sheet_name, rule.formula_cell))
            error = _formula_error(formula_result)
            if error:
                detail = (
                    "校验公式引用无效（#REF!）。请确认外部文件添加步骤已执行，且工作表名称与模板公式一致。"
                    if error == "#REF!" else "校验公式计算异常：{}，请检查模板公式及引用数据".format(error)
                )
                identity = _issue_id(source_file, rule)
                result.append(Issue(
                    issue_id=identity, period=period, batch_id=batch_id, audit_time=audit_time,
                    triggered=True, status="", first_seen_period="", previous_seen_period="", consecutive_count=1,
                    org_code=org_code, org_name=org_name, report_code=rule.report_code, sheet_name=rule.sheet_name,
                    rule_id=identity, severity="错误", formula_cell=rule.formula_cell, target_cell=rule.formula_cell,
                    target_value="", formula_result=error, message=detail, source_file=str(source_file.resolve()),
                    audit_file=str(audit_path.resolve()), detail=detail,
                ))
                continue
            if rule.result_mode == "error_only" or formula_result in (None, False, 0) or not str(formula_result).strip():
                continue
            if is_negligible_zero_difference_result(formula_result):
                continue
            if rule.result_mode == "keyword" and not _has_result_marker(formula_result):
                continue
            parsed = parse_formula_result(str(formula_result), default_severity=rule.severity, default_message=rule.message)
            target_value = parsed.value if rule.value_from_result else cached_values.get(
                (rule.sheet_name, rule.target_cell)
            )
            identity = _issue_id(source_file, rule, parsed.indicator, parsed.check_type)
            result.append(Issue(
                issue_id=identity, period=period, batch_id=batch_id, audit_time=audit_time,
                triggered=True, status="", first_seen_period="", previous_seen_period="", consecutive_count=1,
                org_code=org_code, org_name=org_name, report_code=rule.report_code, sheet_name=rule.sheet_name,
                rule_id=identity, severity=parsed.severity, formula_cell=rule.formula_cell, target_cell=rule.target_cell,
                target_value=target_value or "", formula_result=formula_result, message=parsed.message,
                source_file=str(source_file.resolve()), audit_file=str(audit_path.resolve()), check_field=parsed.indicator,
                check_type=parsed.check_type,
                comparison_value=parsed.comparison_value, reference_value=parsed.reference_value,
                difference_value=parsed.difference_value, detail=parsed.detail,
            ))
        return result
    finally:
        if book is not None:
            book.close()


def _read_issue_value_blocks(
    audit_path: Path, rules: Iterable[AuditRule],
) -> dict[tuple[str, str], Any]:
    """Read each sheet's required cells as one compact value matrix.

    ``ReadOnlyWorksheet.__getitem__`` is a streaming lookup: repeating it for
    hundreds of scattered rules repeatedly walks the worksheet XML.  Grouping
    rule cells into one rectangle per sheet keeps extraction linear in the
    actual cell area, then all rule parsing happens in normal Python memory.
    """
    cells_by_sheet: dict[str, set[str]] = {}
    for rule in rules:
        cells_by_sheet.setdefault(rule.sheet_name, set()).add(rule.formula_cell)
        if not rule.value_from_result:
            cells_by_sheet.setdefault(rule.sheet_name, set()).add(rule.target_cell)
    values: dict[tuple[str, str], Any] = {}
    book = load_workbook(audit_path, read_only=True, data_only=True, keep_links=False)
    try:
        for sheet_name, addresses in cells_by_sheet.items():
            if sheet_name not in book.sheetnames:
                raise TemplateError("审核副本缺少工作表：" + sheet_name)
            positions = {
                address: coordinate_to_tuple(address.replace("$", ""))
                for address in addresses
            }
            rows = [row for row, _ in positions.values()]
            columns = [column for _, column in positions.values()]
            first_row, last_row = min(rows), max(rows)
            first_column, last_column = min(columns), max(columns)
            matrix = [
                tuple(row) for row in book[sheet_name].iter_rows(
                    min_row=first_row, max_row=last_row,
                    min_col=first_column, max_col=last_column,
                    values_only=True,
                )
            ]
            for address, (row, column) in positions.items():
                row_offset, column_offset = row - first_row, column - first_column
                try:
                    value = matrix[row_offset][column_offset]
                except IndexError:
                    # A sparse or incorrect worksheet dimension can make a
                    # streamed rectangle shorter than the requested bounds.
                    # Retry that one formula/target cell instead of losing all
                    # otherwise valid findings from this workbook.
                    try:
                        value = book[sheet_name].cell(row=row, column=column).value
                    except Exception as exc:
                        raise TemplateError(
                            "读取审核副本缓存失败：{}!{}；工作表返回范围不完整".format(
                                sheet_name, address.replace("$", "")
                            )
                        ) from exc
                values[(sheet_name, address)] = value
    finally:
        book.close()
    return values
