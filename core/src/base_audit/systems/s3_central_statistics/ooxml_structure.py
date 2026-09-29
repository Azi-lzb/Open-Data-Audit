# -*- coding: utf-8 -*-
"""OOXML package 结构解析与完整性检查（namespace 无关、只读）。

为 Fast OOXML / Direct OOXML 提供 workbook.xml 与 workbook.xml.rels 的
**XML parser** 解析（ElementTree，不依赖 namespace 前缀写法），以及
relationship Target → ZIP part 的统一路径解析。

铁律（AGENTS「OOXML 结构与用户错误提示铁律」）：
- workbook.xml / rels 等 namespace-aware XML 禁止用依赖前缀的正则解析；
- 启用工作表的 rId 必须有 relationship、Target 必须落在 ZIP 内；
- 检查器只读；错误分「用户可见」与「技术日志」两个通道。
"""

from __future__ import annotations

import posixpath
from dataclasses import dataclass, field
from xml.etree import ElementTree as StdET

WORKBOOK_PART = "xl/workbook.xml"
WORKBOOK_RELS_PART = "xl/_rels/workbook.xml.rels"


class OoxmlStructureError(RuntimeError):
    """OOXML 包结构问题：用户提示与技术日志分离。"""

    def __init__(
        self, user_message: str, *, error_code: str, sheet_name: str = "",
        technical: dict | None = None,
    ) -> None:
        self.user_message = user_message
        self.error_code = error_code
        self.sheet_name = sheet_name
        self.technical = technical or {}
        detail = "；".join(f"{k}={v}" for k, v in self.technical.items())
        suffix = f" | {detail}" if detail else ""
        super().__init__(f"{user_message} [{error_code}{suffix}]")


@dataclass(frozen=True)
class Relationship:
    """一条 package relationship（与 namespace 前缀写法无关）。"""

    relationship_id: str
    target: str
    type: str = ""
    target_mode: str = "Internal"     # Internal / External

    @property
    def is_external(self) -> bool:
        return self.target_mode.strip().lower() == "external"


@dataclass
class WorkbookStructure:
    """workbook.xml 的 sheet 清单 + rels 的 relationship 映射（只读快照）。"""

    sheets: list[dict] = field(default_factory=list)          # name/rid/state
    relationships: dict[str, Relationship] = field(default_factory=dict)  # by Id


import re as _re

_PREFIXED_NAME = _re.compile(r"(<\/?|\s)(?:[\w.]+:)((?:[A-Za-z_][\w.-]*))")


def _strip_prefix_markers(xml_bytes: bytes) -> bytes:
    """剥掉标签/属性名上的 namespace 前缀（ns0:、pkg: 等）。

    只重写标签与属性名（前缀属于 XML 序列化细节，语义由 namespace URI 决定，
    而本模块只按本地名识别结构）；属性值（如公式、字符串）不受影响。
    同时清掉 xmlns 声明，避免保留未绑定前缀的 r:id 触发 unbound prefix。
    """
    text = xml_bytes.decode("utf-8", errors="replace")
    text = _PREFIXED_NAME.sub(lambda m: m.group(1) + m.group(2), text)
    text = text.replace("xmlns:ns0=", "xmlns_removed=").replace(
        "xmlns:pkg=", "xmlns_removed=")
    return text.encode("utf-8")


def _local_tag(element) -> str:
    """标签的本地名（剥离 namespace 前缀/URI），使 ns0:Relationship==Relationship。"""
    tag = element.tag
    if isinstance(tag, str) and "}" in tag:
        return tag.rsplit("}", 1)[-1]
    return str(tag)


def parse_relationships(rels_xml: bytes | str) -> dict[str, Relationship]:
    """解析 workbook.xml.rels → {Id: Relationship}；namespace 前缀无关。"""
    if isinstance(rels_xml, str):
        rels_xml = rels_xml.encode("utf-8")
    if not rels_xml.strip():
        return {}
    try:
        root = StdET.fromstring(_strip_prefix_markers(rels_xml))
    except StdET.ParseError as exc:
        raise OoxmlStructureError(
            "工作簿关系文件（workbook.xml.rels）不是有效的 XML，无法读取工作表结构。",
            error_code="OOXML_RELS_MALFORMED",
            technical={"parse_error": str(exc)},
        ) from exc
    relationships: dict[str, Relationship] = {}
    for element in root.iter():
        if _local_tag(element) != "Relationship":
            continue
        attributes = {
            _local_tag_key(key): value for key, value in element.attrib.items()
        }
        relationship_id = attributes.get("Id", "")
        target = attributes.get("Target", "")
        if not relationship_id or not target:
            continue
        relationships[relationship_id] = Relationship(
            relationship_id=relationship_id, target=target,
            type=attributes.get("Type", ""),
            target_mode=attributes.get("TargetMode") or "Internal",
        )
    return relationships


def _local_tag_key(key: str) -> str:
    return key.rsplit("}", 1)[-1]


def parse_workbook_sheets(workbook_xml: bytes | str) -> list[dict]:
    """解析 workbook.xml 的 <sheet> 清单：name / rid / state；namespace 无关。"""
    if isinstance(workbook_xml, str):
        workbook_xml = workbook_xml.encode("utf-8")
    if not workbook_xml.strip():
        raise OoxmlStructureError(
            "工作簿主文件（workbook.xml）为空，无法读取工作表结构。",
            error_code="OOXML_WORKBOOK_EMPTY",
        )
    try:
        root = StdET.fromstring(_strip_prefix_markers(workbook_xml))
    except StdET.ParseError as exc:
        raise OoxmlStructureError(
            "工作簿主文件（workbook.xml）不是有效的 XML，无法读取工作表结构。",
            error_code="OOXML_WORKBOOK_MALFORMED",
            technical={"parse_error": str(exc)},
        ) from exc
    sheets: list[dict] = []
    for element in root.iter():
        if _local_tag(element) != "sheet":
            continue
        attributes = {_local_tag_key(key): value for key, value in element.attrib.items()}
        name = attributes.get("name", "")
        relationship_id = attributes.get("r:id") or attributes.get("id") or ""
        if not name:
            continue
        sheets.append({
            "name": name, "rid": relationship_id,
            "state": attributes.get("state", ""),
            "sheet_id": attributes.get("sheetId", ""),
        })
    return sheets


def resolve_relationship_target(source_part: str, target: str) -> str:
    """OOXML relationship Target → ZIP 内部 part 路径（POSIX 规范化）。

    兼容三种写法：``worksheets/sheet1.xml``（相对源 part）、
    ``/xl/worksheets/sheet1.xml``（包根绝对）、``../theme/theme1.xml``（相对上跳）。
    """
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    base_dir = posixpath.dirname(source_part)
    if base_dir.endswith("/_rels"):
        # .rels 文件位于 <dir>/_rels/ 下，其 Target 相对于 <dir>（OOXML 规范）。
        base_dir = posixpath.dirname(base_dir)
    return posixpath.normpath(posixpath.join(base_dir, target))


def read_workbook_structure(
    workbook_xml: bytes | str, rels_xml: bytes | str,
) -> WorkbookStructure:
    """解析 workbook + rels 并做一致性检查；问题抛 OoxmlStructureError。"""
    structure = WorkbookStructure()
    structure.sheets = parse_workbook_sheets(workbook_xml)
    structure.relationships = parse_relationships(rels_xml)
    if not structure.relationships:
        raise OoxmlStructureError(
            "无法读取工作表结构：工作簿关系文件为空或不完整。"
            "请使用 Excel/WPS 打开该文件并正常保存一次后重试。",
            error_code="OOXML_RELS_EMPTY",
            technical={"rels_part": WORKBOOK_RELS_PART},
        )
    for sheet in structure.sheets:
        relationship_id = sheet["rid"]
        if not relationship_id:
            raise OoxmlStructureError(
                f"无法读取工作表“{sheet['name']}”：该工作表缺少内部引用编号。",
                error_code="OOXML_SHEET_RID_MISSING",
                sheet_name=sheet["name"],
                technical={"sheet": sheet["name"], "source": WORKBOOK_PART},
            )
        if relationship_id not in structure.relationships:
            raise OoxmlStructureError(
                f"无法读取工作表“{sheet['name']}”：当前 Excel 文件的内部工作表结构异常，"
                "程序无法正常读取该工作表。建议：使用 Excel 或 WPS 打开该配置文件，"
                "正常保存一次，关闭文件后重新执行。",
                error_code="OOXML_WORKSHEET_RELATION_MISSING",
                sheet_name=sheet["name"],
                technical={
                    "sheet": sheet["name"], "rid": relationship_id,
                    "source": WORKBOOK_PART, "rels": WORKBOOK_RELS_PART,
                },
            )
    return structure


def resolve_worksheet_parts(
    structure: WorkbookStructure, *,
    source_part: str = WORKBOOK_PART, parts: dict[str, bytes] | set[str] | None = None,
) -> dict[str, str]:
    """每个 sheet → 解析后的 worksheet part 路径；Target 缺失/悬空即报错。

    External relationship 不按 ZIP 内部 part 处理（跳过路径解析，
    由调用方决定语义）；``parts`` 提供时检查 Target 真实存在。
    """
    names = set(parts) if isinstance(parts, dict) else (parts or None)
    resolved: dict[str, str] = {}
    for sheet in structure.sheets:
        relationship = structure.relationships.get(sheet["rid"])
        if relationship is None or relationship.is_external:
            continue
        part = resolve_relationship_target(source_part, relationship.target)
        if names is not None and part not in names:
            raise OoxmlStructureError(
                f"无法读取工作表“{sheet['name']}”：当前 Excel 文件的内部工作表结构异常，"
                "程序无法正常读取该工作表。建议：使用 Excel 或 WPS 打开该配置文件，"
                "正常保存一次，关闭文件后重新执行。",
                error_code="OOXML_WORKSHEET_PART_MISSING",
                sheet_name=sheet["name"],
                technical={
                    "sheet": sheet["name"], "rid": sheet["rid"],
                    "target": relationship.target, "resolved_part": part,
                    "part_exists": False,
                },
            )
        resolved[sheet["name"]] = part
    return resolved


def validate_workbook_relationships(
    workbook_xml: bytes | str, rels_xml: bytes | str,
    parts: dict[str, bytes] | set[str],
) -> None:
    """删除/重写后的一致性检查：sheet↔relationship↔part 无悬空（只读）。"""
    structure = read_workbook_structure(workbook_xml, rels_xml)
    resolve_worksheet_parts(structure, parts=parts)


# ---------------------------------------------------------------------------
# 工作表公式文本提取（definedName 依赖保护用，namespace 无关、只读）
# ---------------------------------------------------------------------------

# worksheet part 内可能出现公式的元素：单元格公式 <f>，条件格式与
# 数据验证的 <formula>/<formula1>/<formula2>。
_FORMULA_LOCAL_TAGS = frozenset({"f", "formula", "formula1", "formula2"})


def extract_formula_texts(part_xml: bytes | str) -> list[str]:
    """提取 worksheet part 中的全部公式文本（namespace 前缀无关，只读）。

    共享公式从属格（``<f t="shared" si="n"/>``）没有文本，但其名称引用与
    host 公式完全一致（定义名称不受相对引用平移影响），由 host 覆盖。
    """
    if isinstance(part_xml, str):
        part_xml = part_xml.encode("utf-8")
    if not part_xml.strip():
        return []
    try:
        root = StdET.fromstring(_strip_prefix_markers(part_xml))
    except StdET.ParseError:
        return []
    texts: list[str] = []
    for element in root.iter():
        if _local_tag(element) in _FORMULA_LOCAL_TAGS:
            text = "".join(element.itertext()).strip()
            if text:
                texts.append(text)
    return texts


def name_token_pattern(name: str):
    """definedName 的公式引用检测正则：token 边界 + 大小写不敏感。

    Excel 名称语义下 ``myconfigvalue`` 与 ``MyConfigValue`` 是同一名称；
    边界两侧不得是名称字符（字母/数字/下划线/句点），因此 ``Rate`` 不会
    误判 ``InterestRate``，也不会吃掉更长名称的一部分。
    """
    return _re.compile(rf"(?<![\w.]){_re.escape(name)}(?![\w.])", _re.IGNORECASE)


# ---------------------------------------------------------------------------
# definedNames：解析与删表一致性校验（namespace 无关、只读）
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class DefinedName:
    """workbook.xml 中的一条 definedName（只读快照）。"""

    name: str
    text: str                      # 公式文本（已反转义）
    local_sheet_id: int | None     # 表级作用域的工作表序号；None=工作簿级
    hidden: bool = False

    def referenced_sheet_names(self) -> set[str]:
        """公式文本中出现的.sheet 名（尽力而为，用于交叉验证，不用于改写）。

        覆盖 ``'表C'!A1``（引号形式，含 3D ``'A:B'!``）与 ``表C!A1``（裸名）。
        识别不出的形态（外部引用、常量等）不产生名字——调用方按「无法确认」
        保守处理，绝不依据它改写公式。
        """
        refs: set[str] = set()
        for quoted in _QUOTED_SHEET_REF.findall(self.text):
            for part in quoted.replace("''", "'").split(":"):
                part = part.strip()
                if part:
                    refs.add(part)
        for bare in _BARE_SHEET_REF.findall(self.text):
            if bare:
                refs.add(bare)
        return refs


_QUOTED_SHEET_REF = _re.compile(r"'((?:[^']|'')*)'!")
# 裸名引用：! 前是连续标识符字符，且前面不是更长的名字/引用记号。
# #REF! 等错误值经 # 前瞻排除；$A$1 等不含 ! 不受影响。
_BARE_SHEET_REF = _re.compile(r"(?<![A-Za-z0-9_'.#\]])([\w.]*)!")


def parse_defined_names(workbook_xml: bytes | str) -> list[DefinedName]:
    """解析 workbook.xml 的 <definedNames> 块；namespace 前缀无关（只读）。"""
    if isinstance(workbook_xml, str):
        workbook_xml = workbook_xml.encode("utf-8")
    if not workbook_xml.strip():
        return []
    try:
        root = StdET.fromstring(_strip_prefix_markers(workbook_xml))
    except StdET.ParseError:
        return []
    names: list[DefinedName] = []
    for element in root.iter():
        if _local_tag(element) != "definedName":
            continue
        attributes = {_local_tag_key(key): value for key, value in element.attrib.items()}
        raw_local = attributes.get("localSheetId", "")
        local_id = int(raw_local) if raw_local.isdigit() else None
        text = "".join(element.itertext())
        names.append(DefinedName(
            name=attributes.get("name", ""),
            text=text,
            local_sheet_id=local_id,
            hidden=(attributes.get("hidden", "").strip().lower() == "1"),
        ))
    return names


def validate_defined_names(
    workbook_xml: bytes | str, *, sheet_count: int, deleted_sheet_names: set[str],
) -> None:
    """删表重写后的 definedNames 一致性检查（只读，违规抛 OoxmlStructureError）。

    覆盖：localSheetId 不越界；被删表的表级名称不再存在；
    剩余名称的公式文本不指向已删除的工作表。
    """
    for item in parse_defined_names(workbook_xml):
        if item.local_sheet_id is not None and item.local_sheet_id >= sheet_count:
            raise OoxmlStructureError(
                "工作簿内部名称定义指向了不存在的工作表位置，"
                "为避免生成损坏的 Excel 文件，本次操作已终止。"
                "建议：使用 Excel 或 WPS 打开该文件并正常保存一次后重试。",
                error_code="OOXML_DEFINED_NAME_INDEX_OUT_OF_RANGE",
                technical={
                    "defined_name": item.name,
                    "local_sheet_id": item.local_sheet_id,
                    "sheet_count": sheet_count,
                },
            )
        dangling = item.referenced_sheet_names() & set(deleted_sheet_names)
        if dangling:
            raise OoxmlStructureError(
                "工作簿内部名称定义仍引用已删除的工作表，"
                "为避免生成损坏的 Excel 文件，本次操作已终止。",
                error_code="OOXML_DEFINED_NAME_DANGLING_REFERENCE",
                technical={
                    "defined_name": item.name,
                    "deleted_sheet": "、".join(sorted(dangling)),
                    "text": item.text[:120],
                },
            )
