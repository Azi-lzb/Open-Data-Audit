"""大集中统计系统：FAST_OOXML 高速渲染器（ZIP + OOXML/XML 直接处理）。

面向“固定 Excel 模板 × 多机构批量生成”：模板绝大部分内容与机构无关，
只有报表表的数据格、表头追加块、行隐藏与空表删除随机构变化。

原则：
- 不做 Workbook 全量对象化/全量序列化——模板 ZIP 的无关部件按字节原样拷贝；
- 工作表只重生成 ``<sheetData>`` 段：前缀（sheetViews/cols/…）与后缀
  （mergeCells/conditionalFormatting/pageMargins/extLst/…）原样保留；
- 样式在编译期一次性追加进 styles.xml（fills 与 cellXfs 的 count 同步维护，
  fill 索引与 xf 索引分开管理），运行时直接引用 ``s=`` 索引；
- 运行期新增文本用 inlineStr，不改动模板已有 sharedStrings；
- 删空表按「工作表身份」统一处理（SheetPlan，见下）：被删表的表级
  definedNames（_xlnm._FilterDatabase / Print_Area / Print_Titles 等）随表
  移除，保留表的 localSheetId 按删表后的新顺序经 old→new 映射一次性重算
  （支持一次删多表，禁止逐次 -1 的累计偏差写法）；无法证明安全的复杂
  definedName 保守处理：候选空表降级为隐藏，绝不生成可能损坏的文件；
- 输出后做结构校验，失败抛 FastOoxmlRenderError——绝不回退 openpyxl。
"""

from __future__ import annotations

import hashlib
import html
import re
import time
import zipfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as StdET

from .comparison_exporter import VBA_COLOR_PALETTE
from .ooxml_structure import (
    WORKBOOK_PART,
    DefinedName,
    OoxmlStructureError,
    extract_formula_texts,
    name_token_pattern,
    parse_defined_names,
    read_workbook_structure,
    resolve_worksheet_parts,
    validate_defined_names,
    validate_workbook_relationships,
)
from .form_template import EMBEDDED_CONFIG_SHEETS
from .renderers import (
    RENDERER_FAST_OOXML,
    FastOoxmlRenderError,
    OrgRenderResult,
    RendererCapabilities,
    UnsupportedFastRenderOperation,
    band_color_for,
)

RENDERER_SCHEMA_VERSION = 2

# 编译缓存：key = (sha256(模板字节), schema 版本)。同进程同模板只编译一次。
_COMPILED_CACHE: dict[tuple[str, int], "CompiledTemplate"] = {}

_HEADER_DROP_MARKERS = ("剥离", "退出", "上期", "增减额", "环比")

_ROW_RE = re.compile(r"<row\b[^>]*(?:/>|>.*?</row>)", re.S)
_ROW_ATTRS_RE = re.compile(r"<row\b([^>]*?)(?:/>|>)")
_CELL_RE = re.compile(r"<c\b[^>]*/>|<c\b[^>]*>.*?</c>", re.S)

# definedNames 块与条目的原文提取（namespace 前缀无关；决策用 ElementTree
# 解析，这里只负责保留原文重写——只允许 localSheetId 属性补丁）。
# 块匹配必须优先命中自闭合形态（openpyxl 对空块写 <definedNames/>），
# 否则贪婪回溯会把「空块 + 后续块」并成一个跨度，重写出畸形 XML。
_DEFINED_NAMES_BLOCK_RE = re.compile(
    r"<(?:[\w.]+:)?definedNames\b[^>]*?/>|"
    r"<(?:[\w.]+:)?definedNames\b[^>]*?>.*?</(?:[\w.]+:)?definedNames>", re.S)
_DEFINED_NAME_ITEM_RE = re.compile(
    r"<(?:[\w.]+:)?definedName\b[^>]*(?:/>|>.*?</(?:[\w.]+:)?definedName>)", re.S)


def _attr(text: str, name: str) -> str:
    match = re.search(rf'{name}="([^"]*)"', text)
    return match.group(1) if match else ""


def _col_number(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
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


def _num(value: float) -> str:
    """数值文本：整数值不带小数点（对齐 VBA CStr 显示口径）。"""
    if value == int(value) and abs(value) < 1e15:
        return str(int(value))
    return repr(value)


def _xml_escape(text: str) -> str:
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _sheet_biz_class(sheet_name: str) -> str:
    if "A1" in sheet_name:
        return "人民币"
    if "A2" in sheet_name:
        return "外币"
    return "本外币"


def _column_attr(header: str) -> str:
    return "发生额" if "发生额" in header else "余额"


def _column_currency(header: str, biz_class: str) -> str:
    if biz_class == "人民币":
        return "人民币"
    if biz_class == "外币":
        return "美元合计"
    if "人民币" in header:
        return "人民币"
    if "外币" in header:
        return "美元合计"
    return "人民币"


def _parse_rows(xml: str) -> list[tuple[int, str, dict[int, str], dict[int, str], bool]]:
    """解析 sheetData 内部为 (行号, row属性, 列号->单元格XML, 列号->s属性, hidden)。"""
    rows: list[tuple[int, str, dict[int, str], dict[int, str], bool]] = []
    for row_match in _ROW_RE.finditer(xml):
        row_xml = row_match.group(0)
        attrs_match = _ROW_ATTRS_RE.match(row_xml)
        attrs = attrs_match.group(1) if attrs_match else ""
        row_index = int(_attr(attrs, "r") or 0)
        cells: dict[int, str] = {}
        style_by_col: dict[int, str] = {}
        for cell_match in _CELL_RE.finditer(row_xml):
            cell_xml = cell_match.group(0)
            ref = _attr(cell_xml, "r")
            col = _col_number(ref) if ref else 0
            if col:
                cells[col] = cell_xml
                style_by_col[col] = _attr(cell_xml, "s")
        hidden = 'hidden="1"' in attrs or 'hidden="true"' in attrs
        rows.append((row_index, attrs, cells, style_by_col, hidden))
    return rows


def _cell_text(cell_xml: str | None, shared: list[str]) -> str:
    if not cell_xml:
        return ""
    cell_type = _attr(cell_xml, "t")
    value_match = re.search(r"<v>(.*?)</v>", cell_xml, re.S)
    if cell_type == "s" and value_match:
        index = int(value_match.group(1))
        text = shared[index] if 0 <= index < len(shared) else ""
    elif cell_type == "inlineStr":
        text = "".join(re.findall(r"<t[^>]*>(.*?)</t>", cell_xml, re.S))
    else:
        text = value_match.group(1) if value_match else ""
    # 正则提取不走 XML 解析器的实体解码，这里统一解码（&#20313;/&amp; 等），
    # 否则表头"发生额"会以实体形态参与匹配，导致列属性判定全部失败。
    return html.unescape(text)


def _inner_sheet_data(sheet_data: str) -> str:
    return re.sub(r"^<sheetData\b[^>]*>", "", re.sub(r"</sheetData>$", "", sheet_data))


# ---------------------------------------------------------------------------
# 编译产物
# ---------------------------------------------------------------------------


@dataclass
class _CompiledRow:
    attrs: str
    row_index: int
    cells: dict[int, str]           # 列号 -> 原始 <c> XML
    style_by_col: dict[int, str]    # 列号 -> s 属性


@dataclass
class _CompiledSheet:
    name: str
    part: str
    prefix: str
    suffix: str
    rows: list[_CompiledRow] = field(default_factory=list)
    base_headers: list[str] = field(default_factory=list)
    keep_cols: list[int] = field(default_factory=list)
    biz_class: str = ""
    fill_plan: list[tuple[int, list[tuple[int, str]]]] = field(default_factory=list)
    indicator_rows: list[int] = field(default_factory=list)
    header_style: str = ""
    base_style_by_col: dict[int, str] = field(default_factory=dict)
    new_dimension: str = ""


@dataclass
class CompiledTemplate:
    template_hash: str
    template_bytes: bytes
    renderer_version: int
    forms: list[dict[str, str]] = field(default_factory=list)
    sheets: dict[str, _CompiledSheet] = field(default_factory=dict)
    part_order: list[str] = field(default_factory=list)
    parts: dict[str, bytes] = field(default_factory=dict)
    styles_xml: str = ""
    style_map: dict[str, str] = field(default_factory=dict)
    workbook_xml: str = ""
    workbook_rels_xml: str = ""
    content_types_xml: str = ""
    sheet_registry: list[dict[str, str]] = field(default_factory=list)
    internal_config_sheets: tuple[str, ...] = ()
    cache_hit: bool = False


# ---------------------------------------------------------------------------
# TemplateCompiler
# ---------------------------------------------------------------------------


class TemplateCompiler:
    """解析模板 → CompiledTemplate；按 (sha256(模板字节), schema 版本) 缓存。"""

    def __init__(self, template_path: Path) -> None:
        self.template_path = Path(template_path)

    def compile(self) -> CompiledTemplate:
        template_bytes = self.template_path.read_bytes()
        digest = hashlib.sha256(template_bytes).hexdigest()
        key = (digest, RENDERER_SCHEMA_VERSION)
        cached = _COMPILED_CACHE.get(key)
        if cached is not None:
            cached.cache_hit = True
            return cached
        compiled = self._compile(template_bytes, digest)
        _COMPILED_CACHE[key] = compiled
        compiled.cache_hit = False
        return compiled

    def _compile(self, template_bytes: bytes, digest: str) -> CompiledTemplate:
        with zipfile.ZipFile(BytesIO(template_bytes)) as archive:
            part_order = list(archive.namelist())
            parts = {name: archive.read(name) for name in part_order}

        compiled = CompiledTemplate(
            template_hash=digest, template_bytes=template_bytes,
            renderer_version=RENDERER_SCHEMA_VERSION,
            parts=parts, part_order=part_order,
        )
        compiled.workbook_xml = parts.get("xl/workbook.xml", b"").decode("utf-8")
        compiled.workbook_rels_xml = parts.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8")
        compiled.content_types_xml = parts.get("[Content_Types].xml", b"").decode("utf-8")
        compiled.styles_xml = parts.get("xl/styles.xml", b"").decode("utf-8")

        # 工作表注册表：name -> (rId, part, state)。
        # 用 XML parser（namespace 无关）解析 workbook + rels（铁律：
        # namespace-aware 结构禁止用依赖前缀的正则）；同时把非 worksheet 的
        # relationship Target 也解析出来供后续删表/校验使用。
        structure = read_workbook_structure(
            compiled.workbook_xml, compiled.workbook_rels_xml)
        resolved_parts = resolve_worksheet_parts(
            structure, source_part=WORKBOOK_PART, parts=parts)
        rels = {
            relationship.relationship_id: relationship.target
            for relationship in structure.relationships.values()
            if not relationship.is_external
        }
        for position, entry in enumerate(structure.sheets):
            compiled.sheet_registry.append({
                "name": entry["name"], "rid": entry["rid"],
                "part": resolved_parts.get(entry["name"], ""), "state": entry["state"],
                # 工作表身份：模板原始顺序（old_index）与 sheetId。
                # 注意 localSheetId/bookViews 等顺序字段不是身份——它们是
                # 当前 workbook.xml 顺序里的位置，删表后必须按映射重算。
                "sheet_id": entry.get("sheet_id", ""),
                "old_index": position,
            })
        registry_names = {item["name"] for item in compiled.sheet_registry}
        compiled.internal_config_sheets = (
            EMBEDDED_CONFIG_SHEETS
            if all(name in registry_names for name in EMBEDDED_CONFIG_SHEETS)
            else ()
        )

        shared = self._read_shared_strings(parts.get("xl/sharedStrings.xml", b""))

        listing = next(
            (e for e in compiled.sheet_registry if e["name"] == "报表清单"), None)
        if listing is None:
            raise FastOoxmlRenderError(
                "金融表单模板缺少「报表清单」工作表",
                stage="template_compile", sheet="报表清单", part="xl/worksheets",
            )
        listing_part = listing["part"]
        if not listing_part or listing_part not in parts:
            raise FastOoxmlRenderError(
                f"无法读取工作表“报表清单”：当前 Excel 文件的内部工作表结构异常，"
                "程序无法正常读取该工作表。建议：使用 Excel 或 WPS 打开该配置文件，"
                "正常保存一次，关闭文件后重新执行。",
                stage="template_compile", sheet="报表清单", part=listing_part or "(未解析)",
            )
        compiled.forms = self._read_form_list(parts[listing_part], shared)
        form_codes = {item["报表代码"] for item in compiled.forms}

        for entry in compiled.sheet_registry:
            if entry["name"] not in form_codes:
                continue
            compiled.sheets[entry["name"]] = self._compile_sheet(
                entry["name"], entry["part"], parts, shared,
            )
        return compiled

    def _read_shared_strings(self, raw: bytes) -> list[str]:
        if not raw:
            return []
        values = []
        try:
            root = StdET.fromstring(raw)
        except StdET.ParseError:
            return values
        for si in root:
            text = "".join(
                t.text or "" for t in si.iter()
                if t.tag.endswith("}t") or t.tag == "t"
            )
            values.append(text)
        return values

    def _read_form_list(self, part_xml: bytes, shared: list[str]) -> list[dict[str, str]]:
        forms = []
        for row_index, _attrs, cells, _styles, _hidden in _parse_rows(part_xml.decode("utf-8")):
            if row_index < 2:
                continue
            code = _cell_text(cells.get(1), shared).strip()
            if not code or "A" not in code:
                continue
            forms.append({
                "报表代码": code,
                "报表名称": _cell_text(cells.get(2), shared).strip(),
                "频度": _cell_text(cells.get(3), shared).strip(),
                "批次": _cell_text(cells.get(4), shared).strip(),
            })
        return forms

    def _compile_sheet(self, name: str, part: str, parts: dict[str, bytes],
                       shared: list[str]) -> _CompiledSheet:
        raw = parts.get(part)
        if raw is None:
            raise FastOoxmlRenderError(
                f"模板缺少工作表部件：{name}", stage="template_compile", sheet=name, part=part,
            )
        xml = raw.decode("utf-8")
        data_match = re.search(r"<sheetData\b[^>]*(?:/>|>.*?</sheetData>)", xml, re.S)
        if data_match is None:
            raise FastOoxmlRenderError(
                "工作表缺少 sheetData 段", stage="template_compile", sheet=name, part=part,
            )
        compiled = _CompiledSheet(
            name=name, part=part,
            prefix=xml[: data_match.start()], suffix=xml[data_match.end():],
            biz_class=_sheet_biz_class(name),
        )

        base_headers: list[str] = []
        keep_cols: list[int] = []
        header_style = ""
        header_seen = False
        for row_index, attrs, cells, style_by_col, hidden in _parse_rows(
                _inner_sheet_data(data_match.group(0))):
            row = _CompiledRow(attrs=attrs, row_index=row_index,
                               cells=dict(cells), style_by_col=dict(style_by_col))
            if not header_seen:
                # 第 1 行 = 表头：确定保留列与被清理列
                max_col = max(cells) if cells else 0
                if max_col <= 3:
                    raise FastOoxmlRenderError(
                        f"工作表「{name}」缺少数据列",
                        stage="template_compile", sheet=name, part=part,
                    )
                for col in range(4, max_col + 1):
                    header = _cell_text(cells.get(col), shared)
                    if col >= 5 and ((not header) or any(
                            m in header for m in _HEADER_DROP_MARKERS)):
                        continue
                    base_headers.append(header)
                    keep_cols.append(col)
                    if not header_style:
                        header_style = style_by_col.get(col, "")
                if not base_headers:
                    raise FastOoxmlRenderError(
                        f"工作表「{name}」没有可用数据列",
                        stage="template_compile", sheet=name, part=part,
                    )
                header_seen = True
                compiled.rows.append(row)
                continue

            # 数据行：去掉被清理列的单元格；按保留列构建回填计划。
            for col in list(row.cells):
                if col >= 4 and col not in keep_cols:
                    del row.cells[col]
            indicator = _cell_text(row.cells.get(1), shared).replace("'", "").strip()
            if indicator:
                keys = [
                    (col, f"{compiled.biz_class}{indicator}{_column_attr(header)}"
                          f"{_column_currency(header, compiled.biz_class)}")
                    for col, header in zip(keep_cols, base_headers)
                ]
                compiled.fill_plan.append((row.row_index, keys))
                compiled.indicator_rows.append(row.row_index)
            compiled.rows.append(row)

        if not header_seen:
            raise FastOoxmlRenderError(
                f"工作表「{name}」缺少表头行", stage="template_compile", sheet=name, part=part,
            )
        compiled.base_headers = base_headers
        compiled.keep_cols = keep_cols
        col_count = len(base_headers)
        block_start = (max(keep_cols) + 1) if keep_cols else 4
        # 数据列基础样式：取首条数据行同列单元格的 s（新建格沿用列格式）
        first_data = next((r for r in compiled.rows if r.row_index > 1), None)
        for offset, col in enumerate(keep_cols):
            compiled.base_style_by_col[col] = (
                first_data.style_by_col.get(col, "") if first_data is not None else "")
            compiled.base_style_by_col[
                block_start + col_count + offset] = ""          # 上期块
            compiled.base_style_by_col[
                block_start + 2 * col_count + offset] = ""      # 增减额块
            compiled.base_style_by_col[
                block_start + 3 * col_count + offset] = ""      # 环比块
        compiled.header_style = header_style
        compiled.new_dimension = (
            f"A1:{_col_letter(3 + 4 * col_count)}"
            f"{max((r.row_index for r in compiled.rows), default=1)}"
        )
        return compiled


# ---------------------------------------------------------------------------
# SheetPlan：一次渲染的删表计划（工作表身份 × 顺序映射）
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class SheetInfo:
    """工作表身份（编译期确定，随 Registry 缓存，不含渲染期状态）。"""

    name: str
    sheet_id: str
    rid: str
    part: str
    old_index: int          # 模板 workbook.xml 中的原始位置


class SheetPlan:
    """一次渲染的删表/隐藏计划。

    身份（SheetName / rId / worksheet part）用于识别「删的是谁」；
    old_index → new_index 映射用于重算所有依赖工作表顺序的字段
    （localSheetId、bookViews）。顺序字段不是身份，禁止当身份用。
    """

    def __init__(self, registry: list[dict]) -> None:
        self.sheets: list[SheetInfo] = [
            SheetInfo(
                name=item["name"], sheet_id=item.get("sheet_id", ""),
                rid=item["rid"], part=item["part"], old_index=item["old_index"],
            )
            for item in registry
        ]
        self.by_name: dict[str, SheetInfo] = {s.name: s for s in self.sheets}
        self.by_old_index: dict[int, SheetInfo] = {s.old_index: s for s in self.sheets}
        self.by_rid: dict[str, SheetInfo] = {s.rid: s for s in self.sheets}
        self.by_part: dict[str, SheetInfo] = {s.part: s for s in self.sheets if s.part}
        self.deleted: set[str] = set()
        self.hidden: dict[str, str] = {}          # name → 降级原因（用户可读）
        # definedName 处置技术日志：defined_name/owner/old/new/action/reason
        self.defined_name_actions: list[dict] = []
        # 解析序号 → 处置决策（drop/remap/keep + 原因），由 _resolve 定义。
        self.defined_name_decisions: dict[int, dict] = {}

    def mark_deleted(self, name: str) -> None:
        self.deleted.add(name)
        self.hidden.pop(name, None)

    def mark_hidden(self, name: str, reason: str) -> None:
        if name not in self.deleted:
            self.hidden[name] = reason

    def survivors(self) -> list[SheetInfo]:
        return [s for s in self.sheets if s.name not in self.deleted]

    def old_to_new(self) -> dict[int, int]:
        """保留工作表 old_index → 新顺序下标（hidden 也保留，占位不变）。"""
        return {s.old_index: new for new, s in enumerate(self.survivors())}

    def is_empty(self) -> bool:
        return not self.deleted and not self.hidden


# ---------------------------------------------------------------------------
# FastOoxmlRenderer
# ---------------------------------------------------------------------------


class FastOoxmlRenderer:
    """高速模式渲染器：ZIP 字节拷贝 + 工作表 sheetData 定向重生成。"""

    mode = RENDERER_FAST_OOXML
    capabilities = RendererCapabilities()

    def __init__(self, template_path: Path, *, alerts: list[dict],
                 hide_empty_rows: bool, delete_empty_sheets: bool) -> None:
        started = time.monotonic()
        self.alerts = alerts
        self.hide_empty_rows = hide_empty_rows
        self.delete_empty_sheets = delete_empty_sheets
        self.compiler = TemplateCompiler(template_path)
        self.compiled = self.compiler.compile()
        self.cache_hit = self.compiled.cache_hit
        self.prepare_seconds = time.monotonic() - started
        self._style_map: dict[str, str] = {}
        self._styles_xml_patched = self._build_styles()

    # ---- 样式编译（每个渲染器实例一次） ----

    def _build_styles(self) -> str:
        """向 styles.xml 追加 0.00 / 0.00% 及警戒色变体；返回补丁后的 styles.xml。

        fill 索引与 xf 索引分别维护：先追加 fills 得到 fillId，再克隆基础 xf
        （替换 numFmtId / fillId）追加进 cellXfs，两处 count 同步 +N。
        """
        base_indexes: set[str] = {"0"}
        for sheet in self.compiled.sheets.values():
            for style in sheet.base_style_by_col.values():
                base_indexes.add(style or "0")
        colors: set[int] = {3, 6}    # 红=本期无上期有 / 黄=本期有上期无
        for band in self.alerts:
            try:
                color = int(band.get("填充颜色") or 0)
            except (TypeError, ValueError):
                continue
            if color:
                colors.add(color)

        styles = self.compiled.styles_xml
        xfs_match = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", styles, re.S)
        fills_match = re.search(r"<fills\b[^>]*>(.*?)</fills>", styles, re.S)
        if not xfs_match or not fills_match:
            raise FastOoxmlRenderError(
                "styles.xml 缺少 cellXfs 或 fills 定义",
                stage="style_compile", part="xl/styles.xml",
            )
        xf_list = re.findall(r"<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>", xfs_match.group(1), re.S)
        fill_count = len(re.findall(r"<fill>", fills_match.group(1)))

        new_fills: list[str] = []
        fill_index_by_color: dict[int, str] = {}
        new_xfs: list[str] = []
        style_map: dict[str, str] = {}

        def clone_xf(base_xml: str, *, num_fmt: str | None, fill_id: str | None) -> str:
            element = base_xml
            if num_fmt is not None:
                if 'numFmtId="' in element:
                    element = re.sub(r'numFmtId="[^"]*"', f'numFmtId="{num_fmt}"', element)
                else:
                    element = element.replace("<xf ", f'<xf numFmtId="{num_fmt}" ', 1)
                if "applyNumberFormat=" not in element:
                    element = element.replace("<xf ", '<xf applyNumberFormat="1" ', 1)
            if fill_id is not None:
                if 'fillId="' in element:
                    element = re.sub(r'fillId="[^"]*"', f'fillId="{fill_id}"', element)
                else:
                    element = element.replace("<xf ", f'<xf fillId="{fill_id}" ', 1)
                if "applyFill=" not in element:
                    element = element.replace("<xf ", '<xf applyFill="1" ', 1)
            return element

        def ensure_variant(base_idx: str, kind: str, color: int | None) -> str:
            key = f"{base_idx}|{kind}" + (f"|{color}" if color else "")
            if key in style_map:
                return style_map[key]
            idx = int(base_idx) if int(base_idx) < len(xf_list) else 0
            base_xml = xf_list[idx]
            num_fmt = "2" if kind == "value" else "10"   # 内置：2=0.00，10=0.00%
            fill_id = None
            if color:
                if color not in fill_index_by_color:
                    hex_value = VBA_COLOR_PALETTE.get(color, "FFFFFF00")
                    new_fills.append(
                        f'<fill><patternFill patternType="solid">'
                        f'<fgColor rgb="{hex_value}"/><bgColor indexed="64"/>'
                        f"</patternFill></fill>"
                    )
                    fill_index_by_color[color] = str(fill_count + len(new_fills) - 1)
                fill_id = fill_index_by_color[color]
            element = clone_xf(base_xml, num_fmt=num_fmt, fill_id=fill_id)
            new_index = str(len(xf_list) + len(new_xfs))
            new_xfs.append(element)
            style_map[key] = new_index
            return new_index

        for base_idx in sorted(base_indexes, key=lambda v: int(v or 0)):
            style_map[f"{base_idx}|value"] = ensure_variant(base_idx, "value", None)
            style_map[f"{base_idx}|percent"] = ensure_variant(base_idx, "percent", None)
        for base_idx in sorted(base_indexes, key=lambda v: int(v or 0)):
            for color in sorted(colors):
                style_map[f"{base_idx}|value|{color}"] = ensure_variant(base_idx, "value", color)
                style_map[f"{base_idx}|percent|{color}"] = ensure_variant(
                    base_idx, "percent", color)

        # 追加到关闭标签之前：新 fill/xf 的索引 = 原有数量起算，
        # 与 fill_index_by_color / style_map 的索引计算保持一致。
        if new_fills:
            styles = styles.replace(
                f'<fills count="{fill_count}">',
                f'<fills count="{fill_count + len(new_fills)}">',
                1,
            ).replace("</fills>", "".join(new_fills) + "</fills>", 1)
        if new_xfs:
            styles = styles.replace(
                f'<cellXfs count="{len(xf_list)}">',
                f'<cellXfs count="{len(xf_list) + len(new_xfs)}">',
                1,
            ).replace("</cellXfs>", "".join(new_xfs) + "</cellXfs>", 1)
        self._style_map = style_map
        return styles

    def _style(self, base: str, kind: str, color: int | None) -> str:
        if color:
            key = f"{base or '0'}|{kind}|{color}"
            if key in self._style_map:
                return self._style_map[key]
        return self._style_map.get(f"{base or '0'}|{kind}", "0")

    # ---- 渲染 ----

    def render_org(
        self, *, org_name: str, region_name: str, record_date: str,
        data: dict[str, tuple], output_path: Path,
    ) -> OrgRenderResult:
        stats = {
            "renderer": self.mode,
            "template_prepare_time": self.prepare_seconds,
            "template_cache_hit": self.cache_hit,
            "render_time": 0.0, "zip_write_time": 0.0,
            "validation_time": 0.0, "total_time": 0.0,
        }
        render_started = time.monotonic()
        parts = dict(self.compiled.parts)
        workbook_xml = self.compiled.workbook_xml
        rels_xml = self.compiled.workbook_rels_xml
        content_types = self.compiled.content_types_xml
        filled_total = 0

        # 3.3 配置可内嵌金融表单；配置页本身不是机构输出的一部分。
        # 删表统一走 SheetPlan：先收集决策（身份），渲染结束后一次性落盘
        # （顺序字段经 old→new 映射重算，支持一次删多表）。
        plan = SheetPlan(self.compiled.sheet_registry)
        for sheet_name in self.compiled.internal_config_sheets:
            plan.mark_deleted(sheet_name)

        candidates: set[str] = set()
        for sheet_name, meta_sheet in self.compiled.sheets.items():
            matched_rows: dict[int, list[tuple[int, str]]] = {}
            for row_index, keys in meta_sheet.fill_plan:
                hit_keys = [(col, key) for col, key in keys if key in data]
                if hit_keys:
                    matched_rows[row_index] = hit_keys
            if not matched_rows and self.delete_empty_sheets:
                candidates.add(sheet_name)
                continue
            generated, filled = self._render_sheet(
                sheet_name, meta_sheet, matched_rows, data, self.alerts,
                hide_empty_rows=self.hide_empty_rows,
            )
            parts[meta_sheet.part] = generated.encode("utf-8")
            filled_total += filled
        stats["render_time"] = time.monotonic() - render_started

        # 空表删除的两个保守闸门（决策完成前不落盘，保证多表删除原子正确）：
        # 1) 公式引用：被将保留的工作表公式引用的候选表 → 降级隐藏；
        # 2) definedNames：owner 随表删、保留表 localSheetId 重映射；
        #    复杂名称引用候选表且无法证明安全 → 候选表同样降级隐藏。
        deletable = self._resolve_formula_references(candidates, plan, parts)
        self._resolve_defined_names(workbook_xml, plan, deletable, parts)
        if not plan.is_empty():
            workbook_xml, rels_xml, content_types, parts = self._apply_plan(
                workbook_xml, rels_xml, content_types, parts, plan)
        else:
            # 无删表也无降级：保持既有兜底收敛（越界下标收敛到合法范围）。
            workbook_xml = self._clamp_book_views(workbook_xml)
        stats["defined_name_actions"] = list(plan.defined_name_actions)
        # 删除/隐藏统计：用户最终状态必须看得懂（不出现底层术语）。
        stats["deleted_sheet_count"] = len(plan.deleted)
        stats["hidden_sheet_count"] = len(plan.hidden)
        stats["kept_sheet_count"] = len(plan.sheets) - len(plan.deleted)
        extra_messages: list[str] = []
        if plan.hidden:
            stats["hidden_empty_sheets"] = sorted(plan.hidden)
            for hidden_name, reason in sorted(plan.hidden.items()):
                if reason == "defined_name_reference":
                    extra_messages.append(
                        f"工作表“{hidden_name}”包含需要保留的工作簿引用，"
                        "为避免损坏文件，已改为隐藏。"
                    )
                elif reason == "defined_name_formula_reference":
                    extra_messages.append(
                        f"工作表“{hidden_name}”仍被其他工作表引用，"
                        "为避免影响公式计算，已改为隐藏。"
                    )
        if plan.deleted or plan.hidden:
            extra_messages.append(
                f"空工作表处理完成：删除 {len(plan.deleted)} 张，"
                f"因存在工作簿引用而安全隐藏 {len(plan.hidden)} 张。"
            )

        # 删表后的 package 完整性检查：workbook↔rels↔part 不得有悬空；
        # definedNames 不得越界、不得指向已删除的工作表。
        try:
            validate_workbook_relationships(
                workbook_xml, rels_xml,
                set(parts) | set(self.compiled.part_order),
            )
            if plan.deleted:
                validate_defined_names(
                    workbook_xml, sheet_count=len(plan.survivors()),
                    deleted_sheet_names=plan.deleted,
                )
        except OoxmlStructureError as exc:
            raise FastOoxmlRenderError(
                exc.user_message, stage="validation",
                sheet=exc.sheet_name, part=str(exc.technical.get("resolved_part", "")),
            ) from exc

        zip_started = time.monotonic()
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as archive:
            for part in self.compiled.part_order:
                if part not in parts:
                    continue    # 已删除（空表 / calcChain）
                if part == "xl/styles.xml":
                    archive.writestr(part, self._styles_xml_patched)
                elif part == "xl/workbook.xml":
                    archive.writestr(part, workbook_xml)
                elif part == "xl/_rels/workbook.xml.rels":
                    archive.writestr(part, rels_xml)
                elif part == "[Content_Types].xml":
                    archive.writestr(part, content_types)
                else:
                    archive.writestr(part, parts[part])
        stats["zip_write_time"] = time.monotonic() - zip_started

        validation_started = time.monotonic()
        self._validate_output(output_path, parts)
        stats["validation_time"] = time.monotonic() - validation_started
        stats["total_time"] = time.monotonic() - render_started

        name = Path(output_path).name
        return OrgRenderResult(
            output_path=output_path, filled=filled_total, stats=stats,
            messages=[
                f"已生成：{name}（回填 {filled_total} 处，{stats['total_time']:.1f} 秒）",
                *extra_messages,
            ],
        )

    # ---- 单表重生成（只动 sheetData，保留前缀/后缀原样） ----

    def _render_sheet(
        self, sheet_name: str, meta_sheet, matched_rows: dict[int, list[tuple[int, str]]],
        data: dict[str, tuple], alerts: list[dict], *, hide_empty_rows: bool,
    ) -> tuple[str, int]:
        col_count = len(meta_sheet.base_headers)
        block_start = (max(meta_sheet.keep_cols) + 1) if meta_sheet.keep_cols else 4
        header_style = self._style(meta_sheet.header_style or "0", "value", None)
        filled = 0
        out_rows: list[str] = []

        for row in meta_sheet.rows:
            cells = dict(row.cells)
            for col in list(cells):
                if col >= 4 and col not in meta_sheet.keep_cols:
                    del cells[col]
            attrs = row.attrs
            hit_keys = matched_rows.get(row.row_index)

            if row.row_index == 1:
                # 追加 3 组表头（上期-/增减额-/环比-），inlineStr，样式沿用表头
                #（对齐 openpyxl 路径：表头格不带数字格式）。
                header_block_style = f' s="{header_style}"' if header_style else ""
                for block, prefix in enumerate(("上期-", "增减额-", "环比-")):
                    for offset, header in enumerate(meta_sheet.base_headers):
                        ref = f"{_col_letter(block_start + block * col_count + offset)}1"
                        cells[_col_number(ref)] = (
                            f'<c r="{ref}"{header_block_style} t="inlineStr">'
                            f"<is><t>{_xml_escape(prefix + header)}</t></is></c>"
                        )
            elif hit_keys:
                for column, key in hit_keys:
                    current_value, previous_value = data[key]
                    offset = column - 4
                    cur_number = float(current_value) if isinstance(current_value, (int, float)) else 0.0
                    prev_number = float(previous_value) if isinstance(previous_value, (int, float)) else 0.0
                    change = cur_number - prev_number
                    if prev_number != 0 and change != 0:
                        ratio = change / prev_number
                        color = band_color_for(ratio, self.alerts)
                    else:
                        ratio = None
                        # 一侧缺失：黄(6)=本期无上期有，红(3)=本期有上期无
                        color = 6 if current_value is None else (3 if previous_value is None else 0)
                    # 与 openpyxl 路径对齐：四个块都建格——空值格也带样式
                    #（警戒色/数字格式），空单元格写作 <c r= s=/>。
                    base_col_style = meta_sheet.base_style_by_col.get(column, "")
                    targets = [
                        (column, cur_number if isinstance(current_value, (int, float)) else None, "value"),
                        (4 + col_count + offset,
                         prev_number if previous_value is not None else None, "value"),
                        (4 + 2 * col_count + offset,
                         change if change != 0 else None, "value"),
                        (4 + 3 * col_count + offset,
                         (change / prev_number) if (prev_number != 0 and change != 0) else None,
                         "percent"),
                    ]
                    for col, value, kind in targets:
                        style = self._style(base_col_style, kind, color if color else None)
                        style_attr = f' s="{style}"' if style else ""
                        ref = f"{_col_letter(col)}{row.row_index}"
                        if value is None:
                            cells[col] = f'<c r="{ref}"{style_attr}/>'
                        else:
                            cells[col] = f'<c r="{ref}"{style_attr}><v>{_num(float(value))}</v></c>'
                    filled += 1

            if (hide_empty_rows and hit_keys is None and row.row_index > 1
                    and row.row_index in meta_sheet.indicator_rows
                    and 'hidden="' not in attrs):
                attrs = attrs + ' hidden="1"'

            cell_items = sorted(cells.items())
            if not cell_items:
                out_rows.append("<row" + attrs + "/>")
                continue
            out_rows.append(
                "<row" + attrs + ">" + "".join(xml for _col, xml in cell_items) + "</row>"
            )

        prefix = re.sub(
            r'(<dimension ref=")[^"]*(")',
            lambda m: m.group(1) + meta_sheet.new_dimension + m.group(2),
            meta_sheet.prefix, count=1,
        )
        return (
            prefix + "<sheetData>" + "".join(out_rows) + "</sheetData>" + meta_sheet.suffix,
            filled,
        )

    # ---- 空表删除 ----

    def _sheet_referenced_by_formulas(
        self, sheet_name: str, parts: dict[str, bytes], *, exclude: set[str] | None = None,
    ) -> bool:
        """被删表是否仍被其他工作表的公式引用（跨表引用保护，只读）。

        只做公式文本层判断（``'表名'!`` / ``表名!``），不重写任何公式。
        ``exclude`` 中的表不参与扫描——它们自身也将被删除，其公式会随之
        消失，不构成引用约束；扫描范围是注册表内全部工作表。
        """
        import re as _re

        escaped = _re.escape(sheet_name)
        pattern = _re.compile(
            rf"(?:'(?:{escaped})'|(?<![A-Za-z0-9_']){escaped})!")
        skip = set(exclude or ())
        skip.add(sheet_name)
        for info in self._plan.sheets:
            if info.name in skip:
                continue
            part = info.part
            if part not in parts:
                continue
            if pattern.search(parts[part].decode("utf-8", errors="ignore")):
                return True
        return False

    # ---- 删表决策与落盘（SheetPlan：工作表身份 × old→new 顺序映射） ----

    def _resolve_formula_references(
        self, candidates: set[str], plan: SheetPlan, parts: dict[str, bytes],
    ) -> set[str]:
        """公式引用闸门：返回仍可物理删除的候选表，其余降级隐藏。

        从「全部候选可删」出发迭代收敛：某候选被救回（隐藏）后，它的公式
        重新参与引用扫描，可能进一步救回其他候选——将被删除的表的公式
        不构成引用约束。
        """
        self._plan = plan
        doomed = set(candidates)
        changed = True
        while changed and doomed:
            changed = False
            for name in sorted(doomed, key=lambda n: plan.by_name[n].old_index):
                excluded = (doomed - {name}) | plan.deleted
                if self._sheet_referenced_by_formulas(name, parts, exclude=excluded):
                    doomed.discard(name)
                    # 被保留工作表的公式引用：物理删除会制造 #REF!，
                    # 降级为隐藏（不自动改写任何业务公式）。
                    plan.mark_hidden(name, "formula_reference")
                    changed = True
                    break
        return doomed

    def _resolve_defined_names(
        self, workbook_xml: str, plan: SheetPlan, deletable: set[str],
        parts: dict[str, bytes],
    ) -> None:
        """definedNames 闸门：完成候选降级、更新 plan 并生成逐名称处置决策。

        规则（按工作表身份判定；localSheetId 只是顺序位置，不是身份）：
        - owner（localSheetId → Registry 老下标 → 身份）属于被删表 → 名称随表移除；
        - owner 保留 → localSheetId 按 old→new 映射一次性重算（支持一次删多表）；
        - owner 保留但名称文本引用了待删候选表 → 交叉验证异常，无法安全重写：
          引用的候选空表降级为隐藏（名称保持有效），绝不改写公式文本；
        - 引用「明确业务删除」的内嵌配置页等无法挽留的目标 → 名称一并移除；
          但名称被最终保留工作表公式仍以名称引用时 → 阻止本次硬删（否则
          保留公式将变成 #NAME?——结构没坏、业务语义坏了）；
        - 候选空表自己的名称被保留公式以名称引用 → 同样救回该表（降级隐藏），
          绝不允许静默删名称制造 #NAME?；
        - localSheetId 越界等模板异常 → 丢弃该名称并记录，避免 Excel 拒开。
        """
        self._plan = plan
        deletable = set(deletable)
        if not deletable and not plan.deleted:
            return
        parsed = parse_defined_names(workbook_xml)
        if not parsed:
            for name in deletable:
                plan.mark_deleted(name)
            return
        raw_items = self._defined_name_raw_items(workbook_xml)
        if len(raw_items) != len(parsed):
            # 解析双通道不一致：结构不可信，保守终止（不产出可能损坏的文件）。
            raise UnsupportedFastRenderOperation(
                "工作簿名称定义（definedNames）结构无法可靠解析，"
                "为避免生成损坏的 Excel 文件，本次未执行删除。",
                stage="sheet_delete", sheet="",
            )

        def owner_of(item: DefinedName) -> SheetInfo | None:
            if item.local_sheet_id is None:
                return None
            return plan.by_old_index.get(item.local_sheet_id)

        # 公式文本缓存：part → 公式文本列表（名称 token 检测，含条件格式
        # 与数据验证公式；共享公式从属格由 host 公式覆盖）。
        formula_cache: dict[str, list[str]] = {}

        # 迭代收敛：降级会缩小被删集合，被救回表的引用约束随之消失。
        # vetoes 单调累积（救回后不再翻回），否则会在「触发/消失」间振荡。
        # 救回的唯一理由：待删名称（或指向待删表的名称）仍被最终保留工作表
        # 的公式以名称引用——静默删除会制造 #NAME?（结构没坏、语义坏了）。
        # 无人使用的名称允许随表删除，不产生悬空、不产生 #NAME?。
        vetoes: dict[str, str] = {}          # name → 降级原因
        rescue_logs: list[dict] = []
        while True:
            deleted_now = (plan.deleted | deletable) - set(vetoes)
            keep_names = {s.name for s in plan.sheets} - deleted_now
            rescue: dict[str, str] = {}
            for item in parsed:
                owner = owner_of(item)
                if owner is not None and owner.name in deleted_now:
                    # 名称将随 owner 删除：保留公式若仍以名称引用它，
                    # 救回候选 owner（降级隐藏），名称保持有效。
                    if owner.name in deletable and owner.name not in vetoes:
                        users = self._formula_name_users(
                            item.name, keep_names, parts, formula_cache)
                        if users:
                            rescue[owner.name] = "defined_name_formula_reference"
                            rescue_logs.append({
                                "sheet": owner.name,
                                "defined_name": item.name,
                                "referenced_by_sheet": "、".join(sorted(users)),
                                "action": "HIDE_INSTEAD_OF_DELETE",
                                "reason": "DEFINED_NAME_STILL_REFERENCED",
                            })
                    continue                    # 其余 owner 删除：名称随表移除
                soft = item.referenced_sheet_names() & deleted_now & deletable
                if not soft:
                    continue
                # 名称文本引用候选表：仅当名称本身仍被保留公式以名称引用时
                # 才必须救回；无人使用 → 名称随表安全删除。
                users = self._formula_name_users(
                    item.name, keep_names, parts, formula_cache)
                if users:
                    for sheet_name in soft:
                        if sheet_name not in vetoes:
                            rescue[sheet_name] = "defined_name_formula_reference"
                    if not any(
                            log.get("defined_name") == item.name
                            for log in rescue_logs):
                        rescue_logs.append({
                            "sheet": "、".join(sorted(soft)),
                            "defined_name": item.name,
                            "referenced_by_sheet": "、".join(sorted(users)),
                            "action": "HIDE_INSTEAD_OF_DELETE",
                            "reason": "DEFINED_NAME_STILL_REFERENCED",
                        })
            if set(rescue) <= set(vetoes):
                break
            vetoes.update(rescue)

        plan.defined_name_actions.extend(rescue_logs)
        for name, reason in vetoes.items():
            plan.mark_hidden(name, reason)
        for name in deletable - set(vetoes):
            plan.mark_deleted(name)

        # 基于最终存活集合生成逐名称处置决策。
        # 硬删表（内嵌配置页等）与候选空表分开：只有名称指向「无法挽留的
        # 硬删表」时才需要阻断检查（候选表的引用问题已在收敛循环处理）。
        hard_deleted = set(plan.deleted)
        mapping = plan.old_to_new()
        final_keep = {s.name for s in plan.sheets} - plan.deleted
        decisions: dict[int, dict] = {}
        for index, item in enumerate(parsed):
            owner = owner_of(item)
            refs = item.referenced_sheet_names()
            decision: dict = {}
            if owner is not None and owner.name in plan.deleted:
                decision = {"action": "drop", "reason": "owner_deleted",
                            "owner": owner.name}
            elif item.local_sheet_id is not None and owner is None:
                # 作用域指向不存在的表位置：模板本身异常，丢弃以免 Excel 拒开。
                decision = {"action": "drop", "reason": "scope_index_out_of_range"}
            elif refs and (refs & hard_deleted):
                # 名称引用了无法挽留的硬删表：丢弃名称前必须确认最终保留表
                # 公式没有仍以名称引用它——否则保留公式会变成 #NAME?
                # （结构没坏、业务语义坏了）。此时阻止本次硬删，
                # 由用户决定公式如何调整；绝不自动重写业务公式。
                users = self._formula_name_users(item.name, final_keep, parts, formula_cache)
                if users:
                    hard_refs = sorted(refs & hard_deleted)
                    plan.defined_name_actions.append({
                        "defined_name": item.name,
                        "referenced_by_sheet": "、".join(sorted(users)),
                        "blocked_hard_delete_sheet": "、".join(hard_refs),
                        "action": "BLOCK_HARD_DELETE",
                        "reason": "DEFINED_NAME_STILL_REFERENCED",
                        "old_local_sheet_id": item.local_sheet_id,
                        "new_local_sheet_id": None,
                    })
                    raise UnsupportedFastRenderOperation(
                        f"无法删除工作表“{hard_refs[0]}”：其他保留工作表"
                        "仍通过工作簿名称引用该工作表的数据。"
                        "请先确认相关公式是否需要调整。",
                        stage="sheet_delete", sheet=hard_refs[0],
                    )
                # 无人使用：丢弃名称（仅元数据，不改写业务公式），不留悬空。
                decision = {"action": "drop", "reason": "dangling_reference",
                            "deleted_ref": "、".join(sorted(refs & hard_deleted))}
            elif item.local_sheet_id is not None:
                decision = {"action": "remap", "old": item.local_sheet_id,
                            "new": mapping[item.local_sheet_id]}
            else:
                decision = {"action": "keep"}
            if owner is not None and refs and owner.name not in refs:
                # 交叉验证：作用域与引用不一致（异常，记录供排查）。
                decision["anomaly"] = (
                    f"scope={owner.name};refs={'、'.join(sorted(refs))}")
            if decision["action"] != "keep" or "anomaly" in decision:
                plan.defined_name_actions.append({
                    "defined_name": item.name,
                    **decision,
                    "old_local_sheet_id": item.local_sheet_id,
                    "new_local_sheet_id": decision.get("new"),
                })
            decisions[index] = decision
        plan.defined_name_decisions = decisions

    def _formula_name_users(
        self, name: str, keep_sheet_names: set[str], parts: dict[str, bytes],
        formula_cache: dict[str, list[str]],
    ) -> list[str]:
        """最终保留工作表的公式是否仍以名称引用 definedName；返回引用表清单。

        按 Excel 名称语义做 token 边界 + 大小写不敏感检测（``Rate`` 不会
        误判 ``InterestRate``；``myconfigvalue`` 命中 ``MyConfigValue``）。
        """
        pattern = name_token_pattern(name)
        users: list[str] = []
        for sheet_name in sorted(keep_sheet_names):
            info = self._plan.by_name.get(sheet_name)
            if info is None or info.part not in parts:
                continue
            if info.part not in formula_cache:
                formula_cache[info.part] = extract_formula_texts(parts[info.part])
            if any(pattern.search(formula) for formula in formula_cache[info.part]):
                users.append(sheet_name)
        return users

    @staticmethod
    def _defined_name_raw_items(workbook_xml: str) -> list[str]:
        """definedNames 块内的原始元素（保持原样，仅允许 localSheetId 补丁）。"""
        block = _DEFINED_NAMES_BLOCK_RE.search(workbook_xml)
        if block is None:
            return []
        return _DEFINED_NAME_ITEM_RE.findall(block.group(0))

    def _rewrite_defined_names(self, workbook_xml: str, plan: SheetPlan) -> str:
        """按处置决策重写 definedNames 块（保留原始前缀与属性，只补丁下标）。"""
        if not plan.defined_name_decisions:
            return workbook_xml
        block = _DEFINED_NAMES_BLOCK_RE.search(workbook_xml)
        if block is None:
            return workbook_xml
        block_text = block.group(0)
        raw_items = _DEFINED_NAME_ITEM_RE.findall(block_text)
        parsed = parse_defined_names(workbook_xml)
        if len(raw_items) != len(parsed):
            raise UnsupportedFastRenderOperation(
                "工作簿名称定义（definedNames）结构无法可靠解析，"
                "为避免生成损坏的 Excel 文件，本次未执行删除。",
                stage="sheet_delete", sheet="",
            )
        open_match = re.match(r"<(?:[\w.]+:)?definedNames\b[^>]*>", block_text)
        close_match = re.search(r"</(?:[\w.]+:)?definedNames>\s*$", block_text)
        open_tag = open_match.group(0) if open_match else "<definedNames>"
        close_tag = close_match.group(0).strip() if close_match else "</definedNames>"
        kept: list[str] = []
        for index, item in enumerate(raw_items):
            decision = plan.defined_name_decisions.get(index, {"action": "keep"})
            if decision["action"] == "drop":
                continue
            if decision["action"] == "remap":
                item = re.sub(
                    r'localSheetId="[^"]*"',
                    f'localSheetId="{decision["new"]}"', item, count=1)
            kept.append(item)
        inner = "".join(kept)
        replacement = (open_tag + inner + close_tag) if inner else ""
        return workbook_xml.replace(block_text, replacement, 1)

    def _apply_plan(
        self, workbook_xml: str, rels_xml: str, content_types: str,
        parts: dict[str, bytes], plan: SheetPlan,
    ) -> tuple[str, str, str, dict[str, bytes]]:
        """一次性落盘删表/隐藏计划；多表删除共享同一份 old→new 映射。"""
        # 1) definedNames：被删表的表级名称随表移除，保留表下标统一重算。
        workbook_xml = self._rewrite_defined_names(workbook_xml, plan)

        entries_by_name = {
            _attr(entry, "name"): entry
            for entry in re.findall(r"<(?:[\w.]+:)?sheet\b[^>]*/>", workbook_xml)
        }
        # 2) 被删表：sheet 条目 / relationship / Content-Types / part / sheet rels。
        for name in sorted(plan.deleted, key=lambda n: plan.by_name[n].old_index):
            entry = entries_by_name.get(name)
            info = plan.by_name[name]
            if entry is not None:
                workbook_xml = workbook_xml.replace(entry, "", 1)
            rel_entry = next(
                (e for e in re.findall(r"<(?:[\w.]+:)?Relationship\b[^>]*/>", rels_xml)
                 if _attr(e, "Id") == info.rid), None)
            if rel_entry is not None:
                rels_xml = rels_xml.replace(rel_entry, "", 1)
            override = next(
                (e for e in re.findall(r"<Override\b[^>]*/>", content_types)
                 if _attr(e, "PartName").lstrip("/") == info.part), None)
            if override is not None:
                content_types = content_types.replace(override, "", 1)
            parts.pop(info.part, None)
            parts.pop(f"xl/worksheets/_rels/{info.part.rsplit('/', 1)[-1]}.rels", None)
        # calcChain 引用旧 sheet 索引：丢弃后由 Excel 重建（安全的经典做法）。
        if plan.deleted and "xl/calcChain.xml" in parts:
            parts.pop("xl/calcChain.xml", None)
            calc_override = next(
                (e for e in re.findall(r"<Override\b[^>]*/>", content_types)
                 if _attr(e, "PartName") == "/xl/calcChain.xml"), None)
            if calc_override is not None:
                content_types = content_types.replace(calc_override, "", 1)
            calc_rel = next(
                (e for e in re.findall(r"<(?:[\w.]+:)?Relationship\b[^>]*/>", rels_xml)
                 if _attr(e, "Target").endswith("calcChain.xml")), None)
            if calc_rel is not None:
                rels_xml = rels_xml.replace(calc_rel, "", 1)
        # 3) 降级隐藏：OOXML 的 <sheet> 用 state 属性（不是 hidden 属性）。
        for name in plan.hidden:
            entry = entries_by_name.get(name)
            if entry is None:
                continue
            if 'state="' in entry:
                replacement = re.sub(r'state="[^"]*"', 'state="hidden"', entry)
            elif entry.endswith("/>"):
                replacement = entry[:-2] + ' state="hidden"/>'
            else:
                replacement = entry
            workbook_xml = workbook_xml.replace(entry, replacement, 1)
        # 4) bookViews 的 firstSheet/activeTab 按同一份 old→new 映射精确重算。
        workbook_xml = self._clamp_book_views(workbook_xml, plan.old_to_new())
        return workbook_xml, rels_xml, content_types, parts

    def _clamp_book_views(self, workbook_xml: str, old_to_new: dict[int, int] | None = None) -> str:
        """把 bookViews 的 firstSheet/activeTab 重映射到现存工作表下标。

        这两个属性是工作表序号。删表后序号会整体前移，若沿用原模板的序号，
        越界时 Excel 直接判定文件损坏并拒绝打开（症状为 COM 报“Workbooks 的
        Open 方法无效”）。有删表计划时直接用 plan 的 old→new 精确映射；
        否则按原始表序与当前条目集合兜底推导。原本被引用的表若自身已被
        删除，则收敛到范围内的最近一张。
        """
        sheet_count = len(re.findall(r"<(?:[\w.]+:)?sheet\b[^>]*/>", workbook_xml))
        if sheet_count == 0:
            return workbook_xml
        block = re.search(r"<bookViews>.*?</bookViews>", workbook_xml, re.S)
        if block is None:
            return workbook_xml
        if old_to_new is None:
            # 兜底：原始表序（编译期登记）→ 当前现存下标；未登记的表按原序。
            original_names = [item["name"] for item in self.compiled.sheet_registry]
            surviving = [name for name in original_names
                         if re.search(rf'<sheet\b[^>]*name="{re.escape(name)}"', workbook_xml)]
            old_to_new = {index: new_index for new_index, index in enumerate(
                position for position, name in enumerate(original_names) if name in surviving)}
        remap = old_to_new

        views = block.group(0)

        def _remap(attribute: str) -> None:
            nonlocal views

            def _replace(match: re.Match) -> str:
                raw = match.group(2)
                try:
                    value = int(raw)
                except (TypeError, ValueError):
                    value = 0
                if value in remap:
                    mapped = remap[value]
                else:
                    # 被引用的表已删除：收敛到范围内最近的下标，保证合法可打开。
                    mapped = min(max(value, 0), sheet_count - 1)
                return f'{match.group(1)}{attribute}="{mapped}"'

            views = re.sub(
                rf'(<workbookView\b[^>]*?){attribute}="(-?\d+)"', _replace, views)

        _remap("activeTab")
        _remap("firstSheet")
        return workbook_xml.replace(block.group(0), views, 1)

    # ---- 输出校验 ----

    def _validate_output(self, output_path: Path, parts: dict[str, bytes]) -> None:
        """结构校验：ZIP 完整性、核心 XML、目标 Sheet 与关系一致；失败不回退。"""
        stage = "validation"
        try:
            with zipfile.ZipFile(output_path) as archive:
                if archive.testzip() is not None:
                    raise FastOoxmlRenderError(
                        "输出 ZIP 完整性校验失败", stage=stage, part="(zip)",
                    )
                names = set(archive.namelist())
                for required in ("xl/workbook.xml", "xl/_rels/workbook.xml.rels",
                                 "[Content_Types].xml"):
                    if required not in names:
                        raise FastOoxmlRenderError(
                            f"输出缺少核心部件：{required}", stage=stage, part=required,
                        )
                workbook_out = archive.read("xl/workbook.xml")
                rels_out = archive.read("xl/_rels/workbook.xml.rels")
        except FastOoxmlRenderError:
            raise
        except Exception as exc:
            raise FastOoxmlRenderError(
                f"输出 ZIP 无法读取：{exc}", stage=stage,
            ) from exc

        try:
            StdET.fromstring(workbook_out)
            StdET.fromstring(rels_out)
            StdET.fromstring(self.compiled.content_types_xml.encode("utf-8"))
        except Exception as exc:
            raise FastOoxmlRenderError(
                f"核心 XML 无法解析：{exc}", stage=stage, part="xl/workbook.xml",
            ) from exc

        try:
            # namespace 无关的结构一致性校验（sheet↔relationship↔part 无悬空）。
            validate_workbook_relationships(workbook_out, rels_out, names)
        except OoxmlStructureError as exc:
            raise FastOoxmlRenderError(
                exc.user_message, stage=stage,
                sheet=exc.sheet_name, part=str(exc.technical.get("resolved_part", "")),
            ) from exc

        for part in self.compiled.part_order:
            if part not in parts and part in names:
                raise FastOoxmlRenderError(
                    "已删除的工作表部件仍存在于输出中", stage=stage, part=part,
                )
