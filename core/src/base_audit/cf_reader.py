"""条件格式：Direct OOXML 规则读取器与标准化规则模型。

职责边界（规则读取与规则求值彻底分离）：

- 本模块只回答「规则是什么」：ZIP + Direct OOXML/XML 读取
  ``xl/workbook.xml``（Sheet 名称 ↔ 部件映射）、每张工作表 XML 的
  ``<conditionalFormatting>/<cfRule>``（sqref、type、operator、formula、
  priority、stopIfTrue、dxfId）与 ``xl/styles.xml`` 的 ``<dxf>`` 填充色，
  并提供 ``CellValueProvider``（sheetData 数值缓存，供 PYTHON 求值器取值）。
- 产出 :class:`NormalizedConditionalRule`：三个求值器（PYTHON /
  COM_EVALUATE / DISPLAY_FORMAT_NATIVE）消费**完全相同**的规则输入，
  结果差异才可归因于求值而非读取。
- 不使用 openpyxl 作为规则读取主路径；不做求值。
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as StdET

SUPPORTED_RULE_TYPES = ("cellis", "expression")   # rule_type 已统一小写
UNSUPPORTED_RULE_TYPES = ("colorScale", "dataBar", "iconSet", "dataValidation", "top10", "aboveAverage", "uniqueValues", "duplicateValues")

_SHEET_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


class CfRuleReadError(RuntimeError):
    """Direct OOXML 规则读取失败（ZIP/核心 XML/部件损坏）。"""


@dataclass(frozen=True)
class NormalizedConditionalRule:
    """标准化条件格式规则：三个求值器的统一输入。"""

    rule_id: str              # 规则唯一标识：sheet|sqref|priority|序号
    sheet: str                # 工作表名称
    applies_to: str           # sqref（可含多个区域，空格分隔）
    rule_type: str            # cellIs / expression（已小写）；其他类型原样保留
    operator: str             # cellIs 运算符（expression 为空）
    formula1: str
    formula2: str
    priority: int
    stop_if_true: bool
    anchor_cell: str          # 所有 AppliesTo 区域的规范左上角锚点（如 C26）
    dxf_id: int | None
    bounds: tuple[tuple[int, int, int, int], ...] = ()  # 每个区域的 (min_col,min_row,max_col,max_row)
    raw_metadata: dict = field(default_factory=dict)    # 顺序、填充色等附加信息

    @property
    def supported(self) -> bool:
        return self.rule_type in SUPPORTED_RULE_TYPES


@dataclass
class CfRuleReadResult:
    rules: list[NormalizedConditionalRule]
    sheet_parts: dict[str, str]              # 表名 -> 部件路径
    cell_values: "CellValueProvider"
    read_seconds: float
    normalize_seconds: float
    sheets_with_rules: list[str]


def _col_number(letters: str) -> int:
    number = 0
    for ch in letters:
        number = number * 26 + (ord(ch.upper()) - 64)
    return number


def _col_letter(number: int) -> str:
    letters = ""
    while number > 0:
        number, rem = divmod(number - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _sqref_bounds(sqref: str) -> list[tuple[int, int, int, int]]:
    """把 sqref（可空格分隔多个区域）解析为 (min_col, min_row, max_col, max_row)。"""
    bounds = []
    for token in str(sqref or "").split():
        match = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?", token, re.I)
        if not match:
            continue
        col1, row1 = _col_number(match.group(1)), int(match.group(2))
        col2 = _col_number(match.group(3)) if match.group(3) else col1
        row2 = int(match.group(4)) if match.group(4) else row1
        bounds.append((min(col1, col2), min(row1, row2), max(col1, col2), max(row1, row2)))
    return bounds


def _anchor_of(bounds: list[tuple[int, int, int, int]]) -> str:
    if not bounds:
        return ""
    # Excel 的相对条件格式公式不能依赖 sqref token 的序列化顺序。
    # openpyxl 会对 MultiCellRange 排序，而 Direct OOXML 保留原始顺序；
    # 用整个 AppliesTo 的最小行/列建立稳定锚点，保证两条读取路径一致。
    lo_col = min(item[0] for item in bounds)
    lo_row = min(item[1] for item in bounds)
    return f"{_col_letter(lo_col)}{lo_row}"


class CellValueProvider:
    """Direct OOXML 的单元格值缓存（sheetData 的 <v> + inlineStr + sharedStrings）。

    供 PYTHON 求值器读取目标单元格当前值；COM_EVALUATE 由 Excel 自行取值，
    不消费本提供器。``sheet_parts`` 以工作表名称为键（调用方负责名称↔部件映射）。
    """

    def __init__(self, sheet_parts: dict[str, bytes], shared: list[str]) -> None:
        self._shared = shared
        self._cache: dict[str, dict[tuple[int, int], object]] = {}
        self._parts = sheet_parts   # 键 = 工作表名称

    def values(self, sheet: str) -> dict[tuple[int, int], object]:
        cached = self._cache.get(sheet)
        if cached is not None:
            return cached
        raw = self._parts.get(sheet)
        grid: dict[tuple[int, int], object] = {}
        if raw:
            try:
                root = StdET.fromstring(raw)
            except StdET.ParseError:
                root = None
            if root is not None:
                for cell in root.iter(f"{_SHEET_NS}c"):
                    ref = cell.get("r") or ""
                    match = re.fullmatch(r"([A-Z]+)(\d+)", ref, re.I)
                    if not match:
                        continue
                    value = self._cell_value(cell)
                    if value is not None:
                        grid[(_col_number(match.group(1)), int(match.group(2)))] = value
        self._cache[sheet] = grid
        return grid

    def _cell_value(self, cell) -> object:
        cell_type = cell.get("t")
        if cell_type == "inlineStr":
            texts = [t.text or "" for t in cell.iter(f"{_SHEET_NS}t")]
            text = "".join(texts).strip()
            return text or None
        v = cell.find(f"{_SHEET_NS}v")
        if v is None or v.text is None:
            return None
        text = v.text.strip()
        if cell_type == "s":
            try:
                index = int(text)
            except ValueError:
                return None
            return self._shared[index] if 0 <= index < len(self._shared) else None
        try:
            return float(text)
        except ValueError:
            return text or None

    def get(self, sheet: str, row: int, col: int):
        return self.values(sheet).get((col, row))


class OoxmlConditionalRuleReader:
    """ZIP + Direct OOXML 读取条件格式规则（不含求值）。"""

    def read(self, workbook_path: Path) -> CfRuleReadResult:
        import time as _time

        started = _time.monotonic()
        workbook_path = Path(workbook_path)
        try:
            with zipfile.ZipFile(workbook_path) as archive:
                part_order = list(archive.namelist())
                parts = {name: archive.read(name) for name in part_order}
        except zipfile.BadZipFile as exc:
            raise CfRuleReadError(f"报送文件无法作为 xlsx 读取：{exc}") from exc
        read_seconds = _time.monotonic() - started

        normalize_started = _time.monotonic()
        workbook_xml = parts.get("xl/workbook.xml", b"")
        rels_xml = parts.get("xl/_rels/workbook.xml.rels", b"")
        if not workbook_xml or not rels_xml:
            raise CfRuleReadError("报送文件缺少 workbook.xml 或其关系文件")

        # 表名 ↔ sheetN.xml 映射（属性顺序无关，openpyxl/Excel 两种写法都兼容）
        rel_targets: dict[str, str] = {}
        for rel in StdET.fromstring(rels_xml):
            rid = rel.get("Id") or ""
            target = rel.get("Target") or ""
            if rid and target:
                rel_targets[rid] = target
        sheet_parts: dict[str, str] = {}
        try:
            workbook_root = StdET.fromstring(workbook_xml)
        except StdET.ParseError as exc:
            raise CfRuleReadError(f"workbook.xml 无法解析：{exc}") from exc
        for sheet in workbook_root.iter():
            if not sheet.tag.endswith("}sheet") and sheet.tag != "sheet":
                continue
            name = sheet.get("name") or ""
            rid = ""
            for key, value in sheet.attrib.items():
                if key.endswith("}id") or key == "id":
                    rid = value
                    break
            target = rel_targets.get(rid, "")
            if name and target:
                part = target.lstrip("/") if target.startswith("/") else "xl/" + target
                sheet_parts[name] = part

        shared = self._read_shared_strings(parts.get("xl/sharedStrings.xml", b""))
        dxfs = self._read_dxfs(parts.get("xl/styles.xml", b""))
        values = CellValueProvider(
            {name: parts[part] for name, part in sheet_parts.items() if part in parts},
            shared,
        )

        rules: list[NormalizedConditionalRule] = []
        sheets_with_rules: list[str] = []
        for sheet_name, part in sorted(sheet_parts.items()):
            raw = parts.get(part)
            if raw is None:
                continue
            try:
                root = StdET.fromstring(raw)
            except StdET.ParseError:
                continue
            sheet_rules = []
            for formatting in root.iter(f"{_SHEET_NS}conditionalFormatting"):
                sqref = formatting.get("sqref") or ""
                bounds = _sqref_bounds(sqref)
                anchor = _anchor_of(bounds)
                for order, cf_rule in enumerate(formatting.findall(f"{_SHEET_NS}cfRule")):
                    rule = self._normalize_rule(
                        sheet_name, sqref, bounds, anchor, cf_rule, order, dxfs,
                    )
                    if rule is not None:
                        sheet_rules.append(rule)
            if sheet_rules:
                sheets_with_rules.append(sheet_name)
                rules.extend(sheet_rules)

        normalize_seconds = _time.monotonic() - normalize_started
        return CfRuleReadResult(
            rules=rules, sheet_parts=sheet_parts, cell_values=values,
            read_seconds=read_seconds, normalize_seconds=normalize_seconds,
            sheets_with_rules=sheets_with_rules,
        )

    def _normalize_rule(self, sheet: str, sqref: str,
                        bounds: list[tuple[int, int, int, int]], anchor: str,
                        cf_rule, order: int, dxfs: list[dict]) -> NormalizedConditionalRule | None:
        rule_type = (cf_rule.get("type") or "").strip()
        if not rule_type:
            return None
        priority = int(cf_rule.get("priority") or 0)
        stop_if_true = (cf_rule.get("stopIfTrue") or "").strip().lower() in ("1", "true")
        dxf_attr = cf_rule.get("dxfId")
        dxf_id = int(dxf_attr) if dxf_attr is not None and dxf_attr.isdigit() else None
        formulas = [
            "".join(formula.itertext()).strip()
            for formula in cf_rule.findall(f"{_SHEET_NS}formula")
        ]
        operator = (cf_rule.get("operator") or "").strip()
        fill_rgb = dxfs.get(dxf_id, {}).get("fill") if dxf_id is not None else None
        return NormalizedConditionalRule(
            rule_id=f"{sheet}|{sqref}|{priority}|{order}",
            sheet=sheet,
            applies_to=sqref,
            rule_type=rule_type.lower(),
            operator=operator,
            formula1=formulas[0] if formulas else "",
            formula2=formulas[1] if len(formulas) > 1 else "",
            priority=priority,
            stop_if_true=stop_if_true,
            anchor_cell=anchor,
            dxf_id=dxf_id,
            bounds=tuple(bounds),
            raw_metadata={"order": order, "fill_rgb": fill_rgb,
                          "raw_type": rule_type},
        )

    def _read_shared_strings(self, raw: bytes) -> list[str]:
        if not raw:
            return []
        try:
            root = StdET.fromstring(raw)
        except StdET.ParseError:
            return []
        values = []
        for si in root:
            values.append("".join(t.text or "" for t in si.iter() if t.tag.endswith("}t")))
        return values

    def _read_dxfs(self, raw: bytes) -> dict[int, dict]:
        """读 styles.xml 的 <dxfs>：dxfId -> {fill: ARGB 或 None}。"""
        if not raw:
            return {}
        try:
            root = StdET.fromstring(raw)
        except StdET.ParseError:
            return {}
        result: dict[int, dict] = {}
        dxfs_parent = None
        for element in root.iter():
            if element.tag.endswith("}dxfs") or element.tag == "dxfs":
                dxfs_parent = element
                break
        if dxfs_parent is None:
            return result
        for index, dxf in enumerate(dxfs_parent):
            fill_rgb = None
            for fill in dxf.iter():
                if not fill.tag.endswith("}patternFill") and fill.tag != "patternFill":
                    continue
                fg = None
                for child in fill:
                    if child.tag.endswith("}fgColor") or child.tag == "fgColor":
                        fg = child
                        break
                if fg is not None:
                    fill_rgb = fg.get("rgb") or fg.get("indexed") or fg.get("theme")
                    break
            result[index] = {"fill": fill_rgb}
        return result


# ---------------------------------------------------------------------------
# Reader 抽象的第二实现：openpyxl 对象模型读取（兼容/对照）
# ---------------------------------------------------------------------------

DirectOoxmlConditionalRuleReader = OoxmlConditionalRuleReader
"""Direct OOXML Reader 的正式名称别名（与 OpenPyxlConditionalRuleReader 并列）。"""


class OpenpyxlCellValueProvider:
    """openpyxl 对象模型的单元格值缓存（接口与 CellValueProvider 一致）。"""

    def __init__(self, book) -> None:
        self._book = book
        self._cache: dict[str, dict[tuple[int, int], object]] = {}

    def values(self, sheet: str) -> dict[tuple[int, int], object]:
        cached = self._cache.get(sheet)
        if cached is not None:
            return cached
        grid: dict[tuple[int, int], object] = {}
        if sheet in self._book.sheetnames:
            for row in self._book[sheet].iter_rows():
                for cell in row:
                    if cell.value is not None:
                        grid[(cell.column, cell.row)] = cell.value
        self._cache[sheet] = grid
        return grid

    def get(self, sheet: str, row: int, col: int):
        return self.values(sheet).get((col, row))


class OpenPyxlConditionalRuleReader:
    """openpyxl 对象模型读取条件格式规则（兼容/对照路径）。

    产出与 :class:`DirectOoxmlConditionalRuleReader` **完全相同**的
    :class:`NormalizedConditionalRule`；求值器不得感知读取方式差异。
    """

    def read(self, workbook_path: Path) -> CfRuleReadResult:
        import time as _time

        from openpyxl import load_workbook

        from .conditional_engine import _dxf_fill_rgb

        started = _time.monotonic()
        workbook_path = Path(workbook_path)
        book = load_workbook(workbook_path, data_only=True, keep_links=False)
        read_seconds = _time.monotonic() - started

        normalize_started = _time.monotonic()
        dxfs = getattr(book, "_differential_styles", None)
        rules: list[NormalizedConditionalRule] = []
        sheets_with_rules: list[str] = []
        for ws in book.worksheets:
            sheet_rules: list[NormalizedConditionalRule] = []
            for order, (cf,) in enumerate(
                    ((item,) for item in _all_conditional_formatting(ws))):
                sqref = str(cf.sqref)
                bounds = _sqref_bounds(sqref)
                anchor = _anchor_of(bounds)
                for rule_order, rule in enumerate(cf.rules):
                    rule_type = (rule.type or "").strip().lower()
                    if not rule_type:
                        continue
                    priority = int(rule.priority) if rule.priority is not None else 0
                    formulas = [str(f or "").strip() for f in (rule.formula or [])]
                    fill_rgb = _dxf_fill_rgb(dxfs, rule.dxfId)
                    normalized = NormalizedConditionalRule(
                        rule_id=f"{ws.title}|{sqref}|{priority}|{rule_order}",
                        sheet=ws.title,
                        applies_to=sqref,
                        rule_type=rule_type,
                        operator=(rule.operator or "").strip(),
                        formula1=formulas[0] if formulas else "",
                        formula2=formulas[1] if len(formulas) > 1 else "",
                        priority=priority,
                        stop_if_true=bool(rule.stopIfTrue),
                        anchor_cell=anchor,
                        dxf_id=rule.dxfId,
                        bounds=tuple(bounds),
                        raw_metadata={"order": rule_order, "fill_rgb": fill_rgb,
                                      "raw_type": rule.type or ""},
                    )
                    sheet_rules.append(normalized)
            if sheet_rules:
                sheets_with_rules.append(ws.title)
                rules.extend(sheet_rules)

        normalize_seconds = _time.monotonic() - normalize_started
        result = CfRuleReadResult(
            rules=rules,
            sheet_parts={},            # openpyxl 路径无部件概念
            cell_values=OpenpyxlCellValueProvider(book),
            read_seconds=read_seconds,
            normalize_seconds=normalize_seconds,
            sheets_with_rules=sheets_with_rules,
        )
        # book 由 provider 持有（惰性取值），服务层结束后随 provider 释放。
        result._openpyxl_book = book   # noqa: SLF001 — 内部传递，非序列化字段
        return result


def _all_conditional_formatting(ws):
    """openpyxl 条件格式清单（兼容旧版无 list 接口的情况）。"""
    try:
        return list(ws.conditional_formatting)
    except (AttributeError, TypeError):
        return []
