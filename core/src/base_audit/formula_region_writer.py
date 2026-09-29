"""公式校验复制：FormulaRegionWriter 抽象与两套写入后端。

职责边界：只负责「模板校验区域 → 审核副本」的公式与格式写入；
计算策略保持稳定基线不变（单实例、单副本 open→算→存→关）。

- ``ComRangeFormulaRegionWriter``（COM_RANGE）：现有稳定路径的包装
  （``ExcelSession.apply_rules`` 的 Range.Copy + Formula 矩阵重写）。
- ``DirectOoxmlFormulaRegionWriter``（DIRECT_OOXML）：ZIP + Direct OOXML
  高速写入——从模板直读校验区域每格公式/样式/合并，把用到的模板样式
  合并进副本 ``styles.xml``（索引映射，fills/fonts/borders/numFmts/cellXfs
  count 同步维护），ZIP 级重生成副本目标表 ``sheetData`` 的区域格与
  ``mergeCells``，丢弃 ``calcChain`` 并写 ``fullCalcOnLoad``；副本其余部件
  字节原样保留（报送数据、条件格式、数据校验不受影响）。

shared formula：模板区域大量使用 ``<f t="shared" si=..>``（主格有公式文本、
依赖格自闭合无文本）；写入前按主格+ref 平移展开为普通公式——依赖格直接
复制会产生空 ``<f>``（Excel 判定文件损坏），si 编号跨簿也可能冲突。

无 AUTO、无自动回退：DIRECT_OOXML 出错抛
:class:`UnsupportedFormulaRegionOperation` / :class:`FormulaRegionWriteError`。
"""

from __future__ import annotations

import html
import re
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as StdET

WRITER_COM_RANGE = "COM_RANGE"
WRITER_DIRECT_OOXML = "DIRECT_OOXML"
WRITER_MODES = (WRITER_COM_RANGE, WRITER_DIRECT_OOXML)

_CELL_RE = re.compile(r"<c\b[^>]*/>|<c\b[^>]*>.*?</c>", re.S)
_ROW_RE = re.compile(r"<row\b[^>]*(?:/>|>.*?</row>)", re.S)
_ROW_ATTRS_RE = re.compile(r"<row\b([^>]*?)(?:/>|>)")
_F_ANY_RE = re.compile(r"<f\b([^>]*)(?:/>|>(.*?)</f>)", re.S)


class FormulaRegionWriteError(RuntimeError):
    """公式区域写入失败（读取/合并/写出阶段），带结构化定位。"""

    def __init__(self, message: str, *, stage: str = "", part: str = "") -> None:
        super().__init__(message)
        self.stage = stage
        self.part = part


class UnsupportedFormulaRegionOperation(FormulaRegionWriteError):
    """DIRECT_OOXML 不支持的场景（如结构化模板的“审核规则”整表复制）。"""


def _attr(text: str, name: str) -> str:
    match = re.search(rf'{name}="([^"]*)"', text)
    return match.group(1) if match else ""


def _col_number(letters: str) -> int:
    number = 0
    for ch in letters.upper():
        number = number * 26 + (ord(ch) - 64)
    return number


def _col_letter(number: int) -> str:
    letters = ""
    while number > 0:
        number, rem = divmod(number - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _bounds_of(address: str) -> tuple[int, int, int, int] | None:
    match = re.fullmatch(
        r"\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?", str(address or "").strip(), re.I)
    if not match:
        return None
    c1, r1 = _col_number(match.group(1)), int(match.group(2))
    c2 = _col_number(match.group(3)) if match.group(3) else c1
    r2 = int(match.group(4)) if match.group(4) else r1
    return (min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2))


def _cell_ref(row: int, col: int) -> str:
    return f"{_col_letter(col)}{row}"


def _xml_escape(text: str) -> str:
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _f_attrs_xml(extra: str) -> str:
    """<f> 的附加属性（array 等）；shared/t/str 属于值语义，不回写。"""
    extra = extra.strip()
    if not extra:
        return ""
    cleaned = []
    for key, value in re.findall(r'(\w+)="([^"]*)"', extra):
        if key == "t" and value == "array":
            cleaned.append('t="array"')
        elif key in ("ref", "aca", "dt2D", "dtr", "del1", "del2", "r1", "r2"):
            cleaned.append(f'{key}="{value}"')
    return (" " + " ".join(cleaned)) if cleaned else ""


def _split_ref_cell(ref: str) -> tuple[int, int]:
    match = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", ref, re.I)
    return (int(match.group(2)), _col_number(match.group(1))) if match else (1, 1)


_FORMULA_STR_RE = re.compile(r'("(?:[^"]|"")*")')
_REF_TOKEN_RE = re.compile(
    r"(?<![A-Za-z0-9_!$\[])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![0-9A-Za-z_(])")


def _shift_refs_outside_strings(formula: str, dr: int, dc: int) -> str:
    """平移公式中的相对引用；字符串字面量内（含 CJK 工作表名）不平移。"""
    if dr == 0 and dc == 0:
        return formula

    def shift_segment(segment: str) -> str:
        def replace(match: re.Match) -> str:
            col_locked, letters, row_locked, digits = match.groups()
            if "!" in segment[: match.start()]:
                return match.group(0)
            col, row = _col_number(letters), int(digits)
            if not col_locked:
                col += dc
            if not row_locked:
                row += dr
            if col < 1 or row < 1:
                return "#REF!"
            return (f"{'$' if col_locked else ''}{_col_letter(col)}"
                    f"{'$' if row_locked else ''}{row}")

        return _REF_TOKEN_RE.sub(replace, segment)

    return "".join(
        part if index % 2 == 1 else shift_segment(part)
        for index, part in enumerate(_FORMULA_STR_RE.split(formula))
    )


_XML_ESCAPE_ORDER = (("&", "&amp;"), ("<", "&lt;"), (">", "&gt;"), ('"', "&quot;"))


def _xml_escape_formula(text: str) -> str:
    """公式文本写回 <f> 前的 XML 转义（& 先行，避免二次转义）。"""
    for ch, ent in _XML_ESCAPE_ORDER:
        text = text.replace(ch, ent)
    return text


# ---------------------------------------------------------------------------
# shared formula 展开（主格 + ref 平移）
# ---------------------------------------------------------------------------

_F_ANY_RE = re.compile(r"<f\b([^>]*)(?:/>|>(.*?)</f>)", re.S)


def _expand_shared_formulas(sheet_xml: str) -> str:
    """把 shared formula 依赖格按主格公式平移展开为普通公式。

    依赖格的 ``<f t="shared" si=..>`` 通常为自闭合（无公式文本）；直接复制
    会产生空 ``<f>``（Excel 判定文件损坏），si 编号跨簿也可能冲突。
    """
    masters: dict[str, tuple[str, str]] = {}   # si -> (主格公式文本, ref)

    def cell_f(cell: str):
        fm = _F_ANY_RE.search(cell)
        if fm is None:
            return None
        attrs = fm.group(1) or ""
        body = fm.group(2) or ""
        return attrs, body, fm.group(0)

    # 第一遍：收集 si 主格（有公式文本的 shared）
    for cell_match in _CELL_RE.finditer(sheet_xml):
        parsed = cell_f(cell_match.group(0))
        if parsed is None:
            continue
        attrs, body, _raw = parsed
        if 't="shared"' not in attrs or not body.strip():
            continue
        masters[_attr(attrs, "si")] = (body, _attr(attrs, "ref"))
    if not masters:
        return sheet_xml

    def translate(body: str, ref: str, cell_ref: str) -> str:
        master = _split_ref_cell(ref.split(":")[0])
        dest = _split_ref_cell(cell_ref)
        # 模板 XML 中公式为实体编码（&quot;/&amp;/&lt;/&gt;），必须先解码再平移，
        # 否则 CJK 表名字符串内的 "!" 会抑制后续相对引用的平移；写回前再编码。
        plain = html.unescape(body)
        shifted = _shift_refs_outside_strings(
            plain, dest[0] - master[0], dest[1] - master[1]
        )
        return _xml_escape_formula(shifted)

    # 第二遍：无公式文本的 shared 依赖格 → 展开为主格公式平移
    result = sheet_xml
    for cell_match in list(_CELL_RE.finditer(sheet_xml)):
        cell = cell_match.group(0)
        parsed = cell_f(cell)
        if parsed is None:
            continue
        attrs, body, raw_f = parsed
        if 't="shared"' not in attrs or body.strip():
            continue
        master = masters.get(_attr(attrs, "si"))
        if master is None:
            continue
        master_body, master_ref = master
        cell_ref = _attr(cell, "r")
        new_body = translate(master_body, master_ref, cell_ref)
        new_cell = cell.replace(raw_f, f"<f>{new_body}</f>", 1)
        result = result.replace(cell, new_cell, 1)
    return result


# ---------------------------------------------------------------------------
# 模板侧读取：校验区域格矩阵 + 合并 + 样式依赖
# ---------------------------------------------------------------------------


@dataclass
class _TemplateCell:
    row: int
    col: int
    style: str                    # 模板 s 索引（'' = 无）
    formula: str = ""             # 展开后的公式文本（无 = 号；'' = 常量格）
    formula_attrs: str = ""       # <f> 的附加属性（如 t="array" ref=".."）
    value_xml: str = ""           # 常量格的值 XML（已转 inlineStr/数值）
    is_blank: bool = False        # 模板中该格完全为空


@dataclass
class _TemplateRegion:
    sheet: str
    address: str
    bounds: tuple[int, int, int, int]
    cells: dict[tuple[int, int], _TemplateCell] = field(default_factory=dict)
    merges: list[str] = field(default_factory=list)


class _TemplateReader:
    """从模板 ZIP 读取校验区域格矩阵、合并与样式依赖。"""

    def __init__(self, template_path: Path, definition) -> None:
        self.definition = definition
        with zipfile.ZipFile(Path(template_path)) as archive:
            self.part_order = list(archive.namelist())
            self.parts = {name: archive.read(name) for name in self.part_order}
        self.shared = self._read_shared_strings(self.parts.get("xl/sharedStrings.xml", b""))

    def _read_shared_strings(self, raw: bytes) -> list[str]:
        if not raw:
            return []
        try:
            root = StdET.fromstring(raw)
        except StdET.ParseError:
            return []
        return ["".join(t.text or "" for t in si.iter() if t.tag.endswith("}t"))
                for si in root]

    def sheet_part(self, sheet_name: str) -> str | None:
        """表名 ↔ 部件映射（workbook.xml + rels，属性顺序无关）。"""
        workbook_xml = self.parts.get("xl/workbook.xml", b"").decode("utf-8")
        rels_xml = self.parts.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8")
        rel_targets = {}
        for rel in StdET.fromstring(rels_xml):
            rid, target = rel.get("Id") or "", rel.get("Target") or ""
            if rid and target:
                rel_targets[rid] = target
        for sheet in StdET.fromstring(workbook_xml).iter():
            if sheet.tag.split("}")[-1] != "sheet":
                continue
            if (sheet.get("name") or "") != sheet_name:
                continue
            rid = ""
            for key, value in sheet.attrib.items():
                if key.endswith("}id") or key == "id":
                    rid = value
                    break
            target = rel_targets.get(rid, "")
            if target:
                return target.lstrip("/") if target.startswith("/") else "xl/" + target
        return None

    def read_region(self, sheet_name: str, address: str) -> _TemplateRegion:
        part = self.sheet_part(sheet_name)
        if part is None or part not in self.parts:
            raise FormulaRegionWriteError(
                f"模板缺少工作表部件：{sheet_name}",
                stage="template_read", part="xl/worksheets",
            )
        xml = self.parts[part].decode("utf-8")
        # shared formula 展开：依赖格自闭合无文本，直接写入副本会产生空 <f>
        xml = _expand_shared_formulas(xml)
        bounds = _bounds_of(address)
        if bounds is None:
            raise FormulaRegionWriteError(
                f"校验区域地址无法解析：{address}",
                stage="template_read", sheet=sheet_name,
            )
        region = _TemplateRegion(sheet=sheet_name, address=address, bounds=bounds)
        lo_col, lo_row, hi_col, hi_row = bounds

        data_match = re.search(r"<sheetData\b[^>]*>(.*?)</sheetData>", xml, re.S)
        if data_match:
            for row_match in _ROW_RE.finditer(data_match.group(1)):
                row_xml = row_match.group(0)
                for cell_match in _CELL_RE.finditer(row_xml):
                    cell = cell_match.group(0)
                    ref = _attr(cell, "r")
                    if not ref:
                        continue
                    row = int(re.search(r"(\d+)$", ref).group(1))
                    col = _col_number(re.match(r"([A-Z]+)", ref).group(1))
                    if not (lo_col <= col <= hi_col and lo_row <= row <= hi_row):
                        continue
                    region.cells[(row, col)] = self._parse_cell(cell, self.shared)

        merges_match = re.search(r"<mergeCells\b[^>]*>(.*?)</mergeCells>", xml, re.S)
        if merges_match:
            for merge in re.findall(r'<mergeCell ref="([^"]+)"/>', merges_match.group(1)):
                merge_bounds = _bounds_of(merge)
                if merge_bounds and (
                        lo_col <= merge_bounds[0] and merge_bounds[2] <= hi_col
                        and lo_row <= merge_bounds[1] and merge_bounds[3] <= hi_row):
                    region.merges.append(merge)
        return region

    def _parse_cell(self, cell: str, shared: list[str]) -> _TemplateCell:
        ref = _attr(cell, "r")
        row = int(re.search(r"(\d+)$", ref).group(1))
        col = _col_number(re.match(r"([A-Z]+)", ref).group(1))
        style = _attr(cell, "s")
        cell_type = _attr(cell, "t")
        f_match = _F_ANY_RE.search(cell)
        v_match = re.search(r"<v>(.*?)</v>", cell, re.S)
        if f_match:
            attrs = f_match.group(1) or ""
            body = html.unescape(f_match.group(2) or "")
            extra = "".join(f' {k}="{v}"' for k, v in
                            (("t", _attr(attrs, "t")), ("ref", _attr(attrs, "ref")))
                            if v and _attr(attrs, "t") != "shared")
            return _TemplateCell(row=row, col=col, style=style,
                                 formula=body.strip(), formula_attrs=extra)
        if v_match:
            raw = html.unescape(v_match.group(1))
            if cell_type == "s":
                try:
                    text = shared[int(raw)] if int(raw) < len(shared) else ""
                except (ValueError, IndexError):
                    text = ""
                return _TemplateCell(row=row, col=col, style=style,
                                     value_xml=f'<is><t>{_xml_escape(text)}</t></is>')
            return _TemplateCell(row=row, col=col, style=style,
                                 value_xml=f"<v>{raw}</v>")
        if cell_type == "inlineStr":
            texts = html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", cell, re.S)))
            return _TemplateCell(row=row, col=col, style=style,
                                 value_xml=f"<is><t>{_xml_escape(texts)}</t></is>")
        return _TemplateCell(row=row, col=col, style=style, is_blank=True)


# ---------------------------------------------------------------------------
# 样式合并：模板 xf 依赖 → 副本 styles.xml 追加（索引映射）
# ---------------------------------------------------------------------------


class _StyleMerger:
    """把模板中用到的 xf（及其 font/fill/border/numFmt 依赖）追加进副本 styles。"""

    def __init__(self, template_styles: str, audit_styles: str) -> None:
        self.template_styles = template_styles
        self.audit_styles = audit_styles
        self._sections_t = self._split_sections(template_styles)
        self._map: dict[str, str] = {}        # 模板 s -> 副本 s
        self._dep_map: dict[str, str] = {}    # (类型, 原索引) -> 副本新索引
        self._pending: list[str] = []         # 追加的 xf XML
        self._extra: dict[str, list[str]] = {"fonts": [], "fills": [],
                                             "borders": [], "numFmts": []}
        self._base = {name: self._existing_count(audit_styles, name)
                      for name in ("fonts", "fills", "borders", "numFmts", "cellXfs")}

    @staticmethod
    def _split_sections(styles: str) -> dict[str, list[str]]:
        result = {}
        for name, item in (("numFmts", "numFmt"), ("fonts", "font"), ("fills", "fill"),
                           ("borders", "border"), ("cellXfs", "xf")):
            match = re.search(rf'<{name}\b[^>]*>(.*?)</{name}>', styles, re.S)
            if match:
                result[name] = re.findall(rf"<{item}\b[^>]*/>|<{item}\b[^>]*>.*?</{item}>",
                                          match.group(1), re.S)
        return result

    @staticmethod
    def _existing_count(styles: str, name: str) -> int:
        match = re.search(rf'<{name}\b[^>]*count="(\d+)"', styles)
        return int(match.group(1)) if match else 0

    def map_style(self, template_style: str) -> str:
        """模板 s 索引 → 副本 s 索引（依赖按需克隆追加；空样式映射 0）。"""
        if not template_style or template_style == "0":
            return "0"
        if template_style in self._map:
            return self._map[template_style]
        index = int(template_style)
        xfs = self._sections_t.get("cellXfs", [])
        if index >= len(xfs):
            self._map[template_style] = "0"   # 越界兜底：用默认样式
            return "0"
        merged = self._clone_dependencies(xfs[index])
        new_index = self._base["cellXfs"] + len(self._pending)
        self._pending.append(merged)
        self._map[template_style] = str(new_index)
        return self._map[template_style]

    def _clone_dependencies(self, xf: str) -> str:
        """克隆 xf 并把其 font/fill/border 依赖追加进副本对应 section。"""
        merged = xf
        for attr, section in (("fontId", "fonts"), ("fillId", "fills"),
                              ("borderId", "borders")):
            value = _attr(xf, attr)
            if not value:
                continue
            new_value = self._dep_map.get(f"{section}:{value}")
            if new_value is None:
                items = self._sections_t.get(section, [])
                index = int(value)
                if index >= len(items):
                    new_value = value
                else:
                    self._extra[section].append(items[index])
                    new_value = str(self._base[section] + len(self._extra[section]) - 1)
                self._dep_map[f"{section}:{value}"] = new_value
            merged = re.sub(rf'{attr}="[^"]*"', f'{attr}="{new_value}"', merged, count=1)
        return merged

    def merged_audit_styles(self) -> str:
        """返回追加依赖与新 xf 后的副本 styles.xml（count 同步维护）。"""
        styles = self.audit_styles
        for section in ("fonts", "fills", "borders"):
            items = self._extra[section]
            if not items:
                continue
            open_match = re.search(rf'<{section}\b[^>]*count="\d+"[^>]*>', styles)
            if open_match is None:
                continue
            new_count = self._existing_count(styles, section) + len(items)
            styles = styles.replace(
                open_match.group(0),
                re.sub(r'count="\d+"', f'count="{new_count}"', open_match.group(0)), 1)
            styles = styles.replace(f"</{section}>", "".join(items) + f"</{section}>", 1)
        if self._pending:
            open_match = re.search(r'<cellXfs\b[^>]*count="\d+"[^>]*>', styles)
            new_count = self._existing_count(styles, "cellXfs") + len(self._pending)
            styles = styles.replace(
                open_match.group(0),
                re.sub(r'count="\d+"', f'count="{new_count}"', open_match.group(0)), 1)
            styles = styles.replace("</cellXfs>", "".join(self._pending) + "</cellXfs>", 1)
        return styles


# ---------------------------------------------------------------------------
# Writer 协议与两个实现
# ---------------------------------------------------------------------------


@dataclass
class FormulaRegionWriteResult:
    written_cells: int
    merged_merges: int
    styles_appended: int
    seconds: float
    messages: list[str] = field(default_factory=list)


class FormulaRegionWriter:
    """写入器协议：把模板校验区域（公式+格式）写入审核副本。"""

    mode = ""

    def write(self, *, template_path: Path, audit_path: Path, definition) -> FormulaRegionWriteResult:
        raise NotImplementedError


class ComRangeFormulaRegionWriter:
    """COM_RANGE：现有稳定路径的包装（apply_rules 的复制段，calculate=False）。"""

    mode = WRITER_COM_RANGE

    def __init__(self, session) -> None:
        self.session = session

    def write(self, *, template_workbook, audit_workbook, definition,
              formula_overrides=None) -> None:
        self.session.apply_rules(
            template_workbook, audit_workbook, definition,
            calculate=False, formula_overrides=formula_overrides,
        )


class DirectOoxmlFormulaRegionWriter:
    """DIRECT_OOXML：ZIP 级把模板校验区域公式/样式写入审核副本。"""

    mode = WRITER_DIRECT_OOXML
    capabilities = {"formula_write": True, "style_write": True,
                    "merge_write": True, "structured_rule_sheet": False}

    def write(self, *, template_path: Path, audit_path: Path,
              definition, formula_overrides=None) -> FormulaRegionWriteResult:
        started = time.monotonic()
        if getattr(definition, "structured", False):
            # 结构化模板：ZIP 级写完校验区域后，用 openpyxl 把模板的“审核规则”
            # 工作表整表复制进副本（值/样式/批注，与 native _copy_sheet 同语义）。
            # 该表是纯静态规则文本，无需 COM；复制放在写出后的副本文件上。
            self._structured_rule_sheet_pending = True
        else:
            self._structured_rule_sheet_pending = False
        template_reader = _TemplateReader(template_path, definition)
        with zipfile.ZipFile(Path(audit_path)) as archive:
            part_order = list(archive.namelist())
            parts = {name: archive.read(name) for name in part_order}

        written_cells = 0
        merged_merges = 0
        all_template_styles: set[str] = set()
        regions: list[_TemplateRegion] = []
        for item in definition.copy_ranges:
            region = template_reader.read_region(item.sheet_name, item.address)
            regions.append(region)
            for cell in region.cells.values():
                if cell.style:
                    all_template_styles.add(cell.style)

        # ---- styles 合并 ----
        audit_styles = parts.get("xl/styles.xml", b"").decode("utf-8")
        template_styles = template_reader.parts.get("xl/styles.xml", b"").decode("utf-8")
        merger = _StyleMerger(template_styles, audit_styles)
        style_map = {s: merger.map_style(s) for s in sorted(all_template_styles)}
        parts["xl/styles.xml"] = merger.merged_audit_styles().encode("utf-8")
        styles_appended = len(merger._pending) + sum(len(v) for v in merger._extra.values())

        # ---- 逐目标表重生成 ----
        for region in regions:
            part = template_reader.sheet_part(region.sheet)
            if part is None or part not in parts:
                raise FormulaRegionWriteError(
                    f"审核副本缺少工作表部件：{region.sheet}",
                    stage="audit_patch", sheet=region.sheet,
                )
            xml = parts[part].decode("utf-8")
            xml, n_cells, n_merges = self._patch_sheet(
                xml, region, style_map)
            parts[part] = xml.encode("utf-8")
            written_cells += n_cells
            merged_merges += n_merges

        # ---- calcChain 丢弃（公式缓存失效，由重算重建） ----
        if "xl/calcChain.xml" in parts:
            parts.pop("xl/calcChain.xml", None)
            content_types = parts.get("[Content_Types].xml", b"").decode("utf-8")
            override = next((e for e in re.findall(r"<Override\b[^>]*/>", content_types)
                             if _attr(e, "PartName") == "/xl/calcChain.xml"), None)
            if override:
                content_types = content_types.replace(override, "", 1)
                parts["[Content_Types].xml"] = content_types.encode("utf-8")
            rels = parts.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8")
            calc_rel = next((e for e in re.findall(r"<Relationship\b[^>]*/>", rels)
                             if "calcChain.xml" in (e or "")), None)
            if calc_rel:
                rels = rels.replace(calc_rel, "", 1)
                parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

        # ---- fullCalcOnLoad（确保打开/重算时强制刷新公式缓存） ----
        workbook_xml = parts.get("xl/workbook.xml", b"").decode("utf-8")
        if "<calcPr" in workbook_xml:
            if "fullCalcOnLoad" not in workbook_xml:
                workbook_xml = re.sub(r"<calcPr\b([^>]*)/>",
                                      r'<calcPr\1 fullCalcOnLoad="1"/>', workbook_xml, count=1)
        else:
            workbook_xml = re.sub(r"(</sheets>)", r'\1<calcPr calcId="0" fullCalcOnLoad="1"/>',
                                  workbook_xml, count=1)
        parts["xl/workbook.xml"] = workbook_xml.encode("utf-8")

        # ---- 写出（其余部件字节原样） ----
        tmp = Path(audit_path).with_name(Path(audit_path).stem + ".tmp.xlsx")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as archive:
            for part in part_order:
                if part in parts:
                    archive.writestr(part, parts[part])
        tmp.replace(Path(audit_path))

        # ---- 结构化模板：整表复制“审核规则” ----
        structured_note = ""
        if getattr(self, "_structured_rule_sheet_pending", False):
            self._copy_structured_rule_sheet(template_path, Path(audit_path))
            structured_note = "，含“审核规则”表整表复制"

        seconds = time.monotonic() - started
        return FormulaRegionWriteResult(
            written_cells=written_cells, merged_merges=merged_merges,
            styles_appended=styles_appended, seconds=seconds,
            messages=[f"Direct OOXML 写入：{written_cells} 格（样式 {styles_appended} 项，"
                      f"合并 {merged_merges} 处{structured_note}，{seconds:.2f} 秒）"],
        )

    @staticmethod
    def _copy_structured_rule_sheet(template_path: Path, audit_path: Path) -> None:
        """把模板“审核规则”工作表整表复制进审核副本（openpyxl，静态文本表）。"""
        from copy import copy as _copy

        from openpyxl import load_workbook

        template = load_workbook(template_path, data_only=False, keep_links=False)
        try:
            if "审核规则" not in template.sheetnames:
                return
            source = template["审核规则"]
            audit = load_workbook(audit_path, data_only=False, keep_links=False)
            try:
                if "审核规则" in audit.sheetnames:
                    del audit["审核规则"]
                target = audit.create_sheet("审核规则")
                target.sheet_state = source.sheet_state
                target.freeze_panes = source.freeze_panes
                for row in source.iter_rows():
                    for cell in row:
                        destination = target[cell.coordinate]
                        destination.value = cell.value
                        if cell.has_style:
                            destination._style = _copy(cell._style)
                        if cell.comment:
                            destination.comment = _copy(cell.comment)
                for key, dimension in source.row_dimensions.items():
                    target.row_dimensions[key].height = dimension.height
                    target.row_dimensions[key].hidden = dimension.hidden
                for key, dimension in source.column_dimensions.items():
                    target.column_dimensions[key].width = dimension.width
                    target.column_dimensions[key].hidden = dimension.hidden
                for merged in source.merged_cells.ranges:
                    target.merge_cells(str(merged))
                audit.save(audit_path)
            finally:
                audit.close()
        finally:
            template.close()

    # ---- 单表重生成：区域内格按模板写入；区域外原样 ----

    def _patch_sheet(self, xml: str, region: _TemplateRegion,
                     style_map: dict[str, str]) -> tuple[str, int, int]:
        lo_col, lo_row, hi_col, hi_row = region.bounds
        written = 0
        out_rows: list[str] = []
        data_match = re.search(r"<sheetData\b[^>]*>(.*?)</sheetData>", xml, re.S)
        if data_match is None:
            raise FormulaRegionWriteError(
                "副本工作表缺少 sheetData", stage="audit_patch", sheet=region.sheet,
            )
        prefix, suffix = xml[: data_match.start()], xml[data_match.end():]
        for row_match in _ROW_RE.finditer(data_match.group(1)):
            row_xml = row_match.group(0)
            attrs_match = _ROW_ATTRS_RE.match(row_xml)
            attrs = attrs_match.group(1) if attrs_match else ""
            row_index = int(_attr(attrs, "r") or 0)
            in_region_rows = lo_row <= row_index <= hi_row
            cells: dict[int, str] = {}
            for cell_match in _CELL_RE.finditer(row_xml):
                cell = cell_match.group(0)
                ref = _attr(cell, "r")
                col = _col_number(re.match(r"([A-Z]+)", ref).group(1)) if ref else 0
                cells[col] = cell
            if in_region_rows:
                for col in range(lo_col, hi_col + 1):
                    template_cell = region.cells.get((row_index, col))
                    if template_cell is None:
                        # 模板区域内模板未定义的格：与 Range.Copy 整矩形覆盖
                        # 语义一致，清掉报送自带的公式/残值，避免模板外的
                        # 旧校验公式与模板规则重复触发。
                        existing = cells.get(col)
                        if existing is not None:
                            ref = _attr(existing, "r") or _cell_ref(row_index, col)
                            cells[col] = f'<c r="{ref}"/>'
                            written += 1
                        continue
                    style = style_map.get(template_cell.style, "")
                    style_attr = f' s="{style}"' if style else ""
                    ref = _cell_ref(row_index, col)
                    if template_cell.formula:
                        cells[col] = (
                            f'<c r="{ref}"{style_attr}>'
                            f"<f{_f_attrs_xml(template_cell.formula_attrs)}>"
                            f"{_xml_escape(template_cell.formula)}</f></c>"
                        )
                        written += 1
                    elif template_cell.value_xml:
                        # 内联字符串必须带 t="inlineStr"，否则 Excel 按数值格
                        # 解析非法子元素，整簿拒绝打开（真实副本曾因此损坏）。
                        type_attr = (' t="inlineStr"'
                                     if template_cell.value_xml.startswith("<is>") else "")
                        cells[col] = (
                            f'<c r="{ref}"{style_attr}{type_attr}>'
                            f"{template_cell.value_xml}</c>"
                        )
                        written += 1
                    elif template_cell.is_blank:
                        # 模板为空的格：清掉报送残留，保持与 Range.Copy 一致
                        cells[col] = f'<c r="{ref}"{style_attr}/>'
                        written += 1
            body = "".join(xml for _col, xml in sorted(cells.items()))
            if body:
                out_rows.append("<row" + attrs + ">" + body + "</row>")
            else:
                out_rows.append("<row" + attrs + "/>")

        xml = prefix + "<sheetData>" + "".join(out_rows) + "</sheetData>" + suffix

        # ---- 合并并入 ----
        merged_count = 0
        if region.merges:
            existing = set(re.findall(r'<mergeCell ref="([^"]+)"/>', xml))
            additions = [merge for merge in region.merges if merge not in existing]
            if additions:
                merged_count = len(additions)
                match = re.search(r'<mergeCells\b[^>]*count="(\d+)"[^>]*>', xml)
                add_xml = "".join(f'<mergeCell ref="{merge}"/>' for merge in additions)
                if match:
                    new_count = int(match.group(1)) + len(additions)
                    xml = xml.replace(match.group(0),
                                      re.sub(r'count="\d+"', f'count="{new_count}"',
                                             match.group(0)), 1)
                    xml = xml.replace("</mergeCells>", add_xml + "</mergeCells>", 1)
                else:
                    block = f'<mergeCells count="{len(additions)}">{add_xml}</mergeCells>'
                    xml = re.sub(r"(</sheetData>)", r"\1" + block, xml, count=1)
        return xml, written, merged_count
