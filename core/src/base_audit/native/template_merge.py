"""UOS/麒麟 联合模板制作（openpyxl 版，无 COM 依赖）。

与 Windows 的 ``merge_org.run_template_merge`` 保持同一业务语义：

1. **基准模板权威**：先把基准模板在文件系统层复制为产物，再只导入来源工作簿中
   基准里不存在的工作表。基准自带的集中系统数据、参照表等外部依赖工作表保持不动，
   因此跨表引用仍指向基准里的表（这正是必须选基准模板的原因）。
2. **命名区域作用域改写**：基准与来源里的单元格名称一律改写为工作簿级唯一名称
   ``<原名>_<工作表名>``，避免多个模板的同名区域（如 ``表结构区域_001``）冲突；
   跨工作表的多区域名称按工作表拆分为多条名称。
3. **跨簿公式重定向**：全部工作表就位后，把 ``[外部文件.xlsx]表名!`` 形式的引用
   改写为本地 ``'表名'!``；引用的表若未被并入，最终会以 ``#REF!`` 暴露并在检查
   报告里列出。
4. **检查报告**：与 Windows 版同结构（合并说明 / 工作表处理清单 / 命名区域处理清单
   / 公式检查）。

本模块只处理 OOXML；公式缓存不重算（模板用于“复制公式与区域”，由目标机器的办公
套件在审核时重算）。
"""

from __future__ import annotations

import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import range_boundaries

from ..merge_org import (
    _rebase_formula_to_local_sheets,
    _template_merge_output,
    _unique_global_name,
    _write_template_merge_report,
)
from ..template import TemplateError


# Excel 保留名称与公式兼容别名：不作为用户单元格区域迁移。
_RESERVED_PREFIXES = ("_xlnm.", "_xlfn.")
_RESERVED_NAMES = {"print_area", "print_titles"}


def _say(on_step: Optional[Callable[[str], None]], text: str) -> None:
    if on_step is not None:
        on_step(text)


def _iter_defined_names(owner):
    """兼容迭代 openpyxl 的 DefinedNameDict：值可能是 DefinedName 或纯字符串。

    工作簿级 ``book.defined_names`` 与工作表级 ``sheet.defined_names`` 都是
    ``{name: DefinedName}``；旧版本或手工构造时可能出现字符串值，这里统一成
    可读取 ``name``/``attr_text`` 的形式，避免 AttributeError。
    """
    container = getattr(owner, "defined_names", None)
    if container is None:
        return []
    result = []
    try:
        items = list(container.items())
    except AttributeError:
        items = [(None, value) for value in list(container)]
    for key, value in items:
        if hasattr(value, "attr_text"):
            result.append(value)
        else:
            result.append(_SimpleName(str(key), str(value)))
    return result


class _SimpleName:
    """把扁平的 name→attr_text 映射补成带属性的对象。"""

    __slots__ = ("name", "attr_text")

    def __init__(self, name: str, attr_text: str) -> None:
        self.name = name
        self.attr_text = attr_text


def _a1_for_bounds(sheet_title: str, min_col: int, min_row: int, max_col: int, max_row: int) -> str:
    """把区域边界还原为带工作表名的引用（绝对地址，Excel 习惯）。"""
    from openpyxl.utils import get_column_letter

    quoted = sheet_title.replace("'", "''")
    start = f"${get_column_letter(min_col)}${min_row}"
    end = f"${get_column_letter(max_col)}${max_row}"
    body = start if start == end else f"{start}:{end}"
    return f"'{quoted}'!{body}"


def _split_attr_text(attr_text: str) -> list[tuple[str, int, int, int, int]]:
    """把 DefinedName.attr_text 拆成 [(工作表, 左上行列, 右下行列)]。

    只接受指向本工作簿工作表的单元格区域；常量、公式别名返回空列表。
    """
    text = str(attr_text or "").strip().lstrip("=")
    if not text or text.startswith(("OFFSET(", "INDIRECT(", "{")):
        return []
    result: list[tuple[str, int, int, int, int]] = []
    # 逗号分隔的多区域（联合引用），逐段解析。
    for part in re.split(r",(?![^()]*\))", text):
        part = part.strip()
        if not part:
            continue
        match = re.match(r"^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Za-z]{1,3})?\$?(\d+)?(?::\$?([A-Za-z]{1,3})?\$?(\d+)?)?$", part)
        if not match:
            return []
        sheet = (match.group(1) or match.group(2) or "").replace("''", "'").strip()
        if not sheet:
            return []
        start_col, start_row = match.group(3), match.group(4)
        end_col, end_row = match.group(5), match.group(6)
        if not start_row:
            return []
        try:
            if start_col and end_row:
                min_col, min_row, max_col, max_row = range_boundaries(
                    f"{start_col}{start_row}:{end_col or start_col}{end_row}"
                )
            elif start_col:
                min_col = max_col = _column_index(start_col)
                min_row = max_row = int(start_row)
            else:
                # 整列引用（如 '表'!$A:$C）不参与迁移。
                return []
        except ValueError:
            return []
        result.append((sheet, min_col, min_row, max_col, max_row))
    return result


def _column_index(letters: str) -> int:
    index = 0
    for char in letters.upper():
        index = index * 26 + (ord(char) - 64)
    return index


def collect_names(book, sheet_title: str) -> list[tuple[str, list[tuple[str, int, int, int, int]]]]:
    """收集某工作表上可迁移的单元格名称：[(名称, 区域列表)]。

    同时覆盖工作簿级名称（指向本表的部分）与工作表局部名称，与 COM 版
    ``_names_for_source_sheet`` 的取舍一致：工作簿级名称若跨多表，只取落在本表
    的区域；工作表局部名称原样保留（不得因指向他表而被改写）。
    """
    entries: list[tuple[str, list[tuple[str, int, int, int, int]], bool]] = []
    seen: set[tuple[str, str]] = set()

    def add(raw_name: str, attr_text: str, is_local: bool) -> None:
        local_name = raw_name.rsplit("!", 1)[-1].strip().strip("'")
        folded = local_name.casefold()
        if folded.startswith(_RESERVED_PREFIXES) or folded in _RESERVED_NAMES:
            return
        key = (folded, str(attr_text))
        if key in seen:
            return
        seen.add(key)
        entries.append((local_name, _split_attr_text(attr_text), is_local))

    for defined in _iter_defined_names(book):
        add(str(defined.name), str(defined.attr_text or ""), False)
    sheet = book[sheet_title]
    for defined in _iter_defined_names(sheet):
        add(str(defined.name), str(defined.attr_text or ""), True)

    collected: list[tuple[str, list[tuple[str, int, int, int, int]]]] = []
    for name, areas, is_local in entries:
        if not areas:
            continue
        own = [item for item in areas if item[0].casefold() == sheet_title.casefold()]
        if not own:
            continue
        # 工作簿级名称跨表时只取本表部分；局部名称必须完整属于本表才迁移。
        if not is_local and len(own) != len(areas):
            pass  # 取本表部分即可（与 COM 版一致）
        collected.append((name, own))
    return collected


def localize_base_names(book) -> tuple[list[tuple[str, str, str, str, str, str]], list[str]]:
    """把基准模板里的单元格名称改写为 ``<原名>_<工作表名>``。

    返回 (记录, 失败名称列表)。
    """
    records: list[tuple[str, str, str, str, str, str]] = []
    failures: list[str] = []
    originals: list[tuple[str, str, bool]] = []
    for defined in _iter_defined_names(book):
        originals.append((str(defined.name), str(defined.attr_text or ""), False))
    for sheet in book.worksheets:
        for defined in _iter_defined_names(sheet):
            originals.append((str(defined.name), str(defined.attr_text or ""), True))

    used = {str(key).rsplit("!", 1)[-1].strip().strip("'").casefold() for key in book.defined_names}
    created: set[tuple[str, str, str]] = set()

    for raw_name, attr_text, is_local in originals:
        local_name = raw_name.rsplit("!", 1)[-1].strip().strip("'")
        folded = local_name.casefold()
        if folded.startswith(_RESERVED_PREFIXES) or folded in _RESERVED_NAMES:
            continue
        areas = _split_attr_text(attr_text)
        if not areas:
            continue  # 常量/公式别名保持原样
        grouped: dict[str, list[tuple[str, int, int, int, int]]] = {}
        for item in areas:
            grouped.setdefault(item[0], []).append(item)
        made: list[tuple[str, str, str]] = []
        try:
            for sheet_title, sheet_areas in grouped.items():
                if sheet_title not in book.sheetnames:
                    failures.append(f"{local_name}（工作表“{sheet_title}”不存在）")
                    continue
                refers_to = "=" + ",".join(
                    _a1_for_bounds(sheet_title, *area[1:]) for area in sheet_areas
                )
                semantic = (folded, sheet_title.casefold(), refers_to.casefold())
                if semantic in created:
                    continue
                created.add(semantic)
                target = _unique_global_name(local_name, sheet_title, used)
                book.defined_names.add(_new_defined_name(target, refers_to))
                made.append((target, sheet_title, refers_to))
            _drop_original(book, raw_name, local_name, is_local)
        except Exception as exc:  # noqa: BLE001 - 逐条记录，不中断整体
            failures.append(f"{local_name}（{exc}）")
            continue
        for target, sheet_title, refers_to in made:
            records.append(("基准模板", sheet_title, target, "工作簿", refers_to, "已复制"))
    return records, failures


def _new_defined_name(name: str, attr_text: str):
    from openpyxl.workbook.defined_name import DefinedName

    return DefinedName(name, attr_text=attr_text)


def _drop_original(book, raw_name: str, local_name: str, is_local: bool) -> None:
    """删除原名（工作簿级或工作表局部），失败不抛出。"""
    if is_local:
        for sheet in book.worksheets:
            if local_name in getattr(sheet, "defined_names", {}):
                try:
                    del sheet.defined_names[local_name]
                except Exception:  # noqa: BLE001
                    pass
        return
    for key in (raw_name, local_name):
        if key in book.defined_names:
            try:
                del book.defined_names[key]
            except Exception:  # noqa: BLE001
                pass


def copy_sheet_names(
    source_book, source_title: str, book, target_title: str, used: set[str],
    source_label: str = "",
) -> tuple[list[tuple[str, str, str, str, str, str]], list[str]]:
    """把来源工作表上的单元格名称复制为目标工作簿级唯一名称。

    ``source_label`` 用于检查报告的“来源工作簿”列（不用 openpyxl 的
    ``book.path``——那是字符串，取 ``.name`` 会抛 AttributeError）。
    """
    records: list[tuple[str, str, str, str, str, str]] = []
    failures: list[str] = []
    for local_name, areas in collect_names(source_book, source_title):
        try:
            refers_to = "=" + ",".join(
                _a1_for_bounds(target_title, *area[1:]) for area in areas
            )
            target = _unique_global_name(local_name, target_title, used)
            book.defined_names.add(_new_defined_name(target, refers_to))
            records.append((source_label, source_title, target, "工作簿", refers_to, "已复制"))
        except Exception as exc:  # noqa: BLE001
            failures.append(f"{local_name}（{exc}）")
    return records, failures


def copy_sheet(source_sheet, target_workbook, title: str, style_cache: dict) -> None:
    """跨簿复制工作表：值/公式/样式/批注/超链接/合并/行高列宽/条件格式/数据验证。

    样式按来源样式去重注册到目标工作簿，避免直接搬 ``_style`` 造成的
    ``alignmentId`` 越界（跨簿样式表编号不同）。
    """
    from copy import copy as _copy

    from openpyxl.cell.cell import Cell, MergedCell

    target = target_workbook.create_sheet(title)
    for row in source_sheet.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            copied = target.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                key = (id(source_sheet.parent), tuple(cell._style))
                target_style = style_cache.get(key)
                if target_style is None:
                    prototype = Cell(target, row=1, column=1)
                    prototype.font = _copy(cell.font)
                    prototype.fill = _copy(cell.fill)
                    prototype.border = _copy(cell.border)
                    prototype.alignment = _copy(cell.alignment)
                    prototype.number_format = cell.number_format
                    prototype.protection = _copy(cell.protection)
                    target_style = _copy(prototype._style)
                    style_cache[key] = target_style
                copied._style = _copy(target_style)
            if cell.comment is not None:
                copied.comment = _copy(cell.comment)
            if cell.hyperlink is not None:
                copied._hyperlink = _copy(cell.hyperlink)
    for merged_range in source_sheet.merged_cells.ranges:
        target.merge_cells(str(merged_range))
    for key, dimension in source_sheet.column_dimensions.items():
        item = target.column_dimensions[key]
        item.width = dimension.width
        item.hidden = dimension.hidden
        item.outline_level = dimension.outline_level
    for index, dimension in source_sheet.row_dimensions.items():
        item = target.row_dimensions[index]
        item.height = dimension.height
        item.hidden = dimension.hidden
        item.outline_level = dimension.outline_level
    target.freeze_panes = source_sheet.freeze_panes
    target.sheet_view.showGridLines = source_sheet.sheet_view.showGridLines
    target.sheet_properties.tabColor = source_sheet.sheet_properties.tabColor
    target.sheet_state = source_sheet.sheet_state
    # 条件格式与数据验证：模板要用于“定位与提取”，必须随表复制。
    for rng in source_sheet.conditional_formatting:
        for rule in rng.rules:
            target.conditional_formatting.add(str(rng.sqref), _copy(rule))
    for validation in source_sheet.data_validations.dataValidation:
        target.add_data_validation(_copy(validation))


def rebase_external_references(book) -> int:
    """把 ``[外部文件.xlsx]表名!`` 形式的引用改写为本地 ``'表名'!``。

    在全部工作表就位后执行；只改公式文本，不改数值。返回改写单元格数。
    """
    rewritten = 0
    for sheet in book.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                value = cell.value
                if not isinstance(value, str) or not value.startswith("="):
                    continue
                rebased = _rebase_formula_to_local_sheets(value)
                if rebased != value:
                    cell.value = rebased
                    rewritten += 1
    return rewritten


def scan_formula_findings(book) -> list[tuple[str, str, str, str, str]]:
    """扫描完成后的模板：``#REF!`` / ``#NAME?`` / 仍存在的外部工作簿引用。

    模板尚未重算，因此错误值来自公式文本本身或上次保存的缓存值。
    """
    findings: list[tuple[str, str, str, str, str]] = []
    for sheet in book.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                value = cell.value
                if not isinstance(value, str) or not value.startswith("="):
                    continue
                if "#REF!" in value:
                    findings.append((sheet.title, cell.coordinate, "#REF! 引用错误", value, ""))
                elif "#NAME?" in value:
                    findings.append((sheet.title, cell.coordinate, "#NAME? 名称错误", value, ""))
                elif re.search(r"\[[^\]]+\]", value):
                    findings.append((sheet.title, cell.coordinate, "仍引用外部工作簿", value, ""))
    return findings


def run_template_merge_native(
    *,
    base_template: Path,
    source_templates: list[Path],
    on_step: Optional[Callable[[str], None]] = None,
):
    """openpyxl 版联合模板制作；返回 ``TemplateMergeResult``。"""
    from ..merge_org import TemplateMergeResult

    base_template = Path(base_template).resolve()
    if not base_template.is_file():
        raise FileNotFoundError(f"底稿模板不存在：{base_template}")
    if base_template.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("openpyxl 版联合模板仅支持 .xlsx/.xlsm 底稿；.xls 请先另存为 .xlsx")

    unique_sources: list[Path] = []
    seen = {base_template}
    for raw_path in source_templates:
        path = Path(raw_path).resolve()
        if not path.is_file():
            raise FileNotFoundError(f"待合并工作簿不存在：{path}")
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError(f"openpyxl 版联合模板仅支持 .xlsx/.xlsm：{path.name}")
        if path not in seen:
            unique_sources.append(path)
            seen.add(path)
    if not unique_sources:
        raise ValueError("请至少选择一个除底稿模板外的来源工作簿")

    output_path, report_path = _template_merge_output(base_template, unique_sources)
    shutil.copy2(str(base_template), str(output_path))
    _say(on_step, "已创建底稿副本；原始模板不会修改")
    _say(on_step, "提示：请确认底稿模板已包含集中系统数据、参照表等外部依赖工作表")

    copied: list[tuple[str, str]] = []
    skipped: list[tuple[str, str]] = []
    named_ranges: list[tuple[str, str, str, str, str, str]] = []
    findings: list[tuple[str, str, str, str, str]] = []
    base_sheets: list[str] = []
    style_cache: dict = {}

    try:
        book = load_workbook(output_path)
        try:
            base_sheets = list(book.sheetnames)
            records, failures = localize_base_names(book)
            if failures:
                raise TemplateError(
                    f"底稿模板有 {len(failures)} 个命名区域处理失败：" + "；".join(failures[:5])
                )
            named_ranges.extend(records)
            used = {str(key).rsplit("!", 1)[-1].strip().strip("'").casefold() for key in book.defined_names}

            existing = {title.casefold() for title in book.sheetnames}
            for source_path in unique_sources:
                _say(on_step, f"正在复制工作簿：{source_path.name}")
                source_book = load_workbook(source_path)
                try:
                    for source_sheet in source_book.worksheets:
                        title = source_sheet.title
                        if title.casefold() in existing:
                            skipped.append((source_path.name, title))
                            continue
                        copy_sheet(source_sheet, book, title, style_cache)
                        records, failures = copy_sheet_names(
                            source_book, title, book, title, used, source_label=source_path.name
                        )
                        if failures:
                            raise TemplateError(
                                f"工作表“{title}”有 {len(failures)} 个命名区域迁移失败："
                                + "；".join(failures[:5])
                            )
                        named_ranges.extend(records)
                        existing.add(title.casefold())
                        copied.append((source_path.name, title))
                finally:
                    source_book.close()

            rewritten = rebase_external_references(book)
            if rewritten and on_step is not None:
                on_step(f"已把 {rewritten} 个跨簿引用改写为本地工作表引用")
            findings = scan_formula_findings(book)
            book.save(output_path)
        finally:
            book.close()
    except Exception:
        if output_path.exists():
            try:
                output_path.unlink()
            except OSError:
                pass
        raise

    _write_template_merge_report(
        report_path,
        base_template=base_template,
        base_sheets=base_sheets,
        sources=unique_sources,
        copied=copied,
        skipped=skipped,
        findings=findings,
        named_ranges=named_ranges,
    )
    _say(on_step, f"联合模板已生成：{output_path.name}")
    _say(on_step, f"检查报告已生成：{report_path.name}")
    return TemplateMergeResult(
        output_path=output_path,
        report_path=report_path,
        base_template=base_template,
        copied_sheets=tuple(copied),
        skipped_sheets=tuple(skipped),
        formula_findings=tuple(findings),
        copied_named_ranges=tuple(named_ranges),
    )
