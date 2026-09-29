"""外部辅助表写入器。

``DIRECT_OOXML`` 的能力边界刻意很窄：它将外部工作簿的纯数值、布尔值和
文本单元格写成新的无样式工作表，不复制样式、批注、冻结窗格、合并、图片、
表格或数据验证。新增文本统一使用 inlineStr，因此不需迁移 sharedStrings。
遇到公式或会影响数据语义的对象明确失败，禁止悄悄生成不完整工作簿。
"""

from __future__ import annotations

from dataclasses import dataclass, field
import html
from pathlib import Path
import re
import time
import zipfile
from xml.etree import ElementTree as ET


EXTERNAL_SHEET_WRITER_OPENPYXL = "OPENPYXL"
EXTERNAL_SHEET_WRITER_DIRECT_OOXML = "DIRECT_OOXML"
EXTERNAL_SHEET_WRITERS = (
    EXTERNAL_SHEET_WRITER_OPENPYXL,
    EXTERNAL_SHEET_WRITER_DIRECT_OOXML,
)

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CONTENT_TYPE_SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
_REL_TYPE_SHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"


class ExternalSheetWriteError(RuntimeError):
    pass


class UnsupportedExternalSheetOperation(ExternalSheetWriteError):
    pass


@dataclass(frozen=True)
class _Cell:
    address: str
    kind: str
    value: str


@dataclass(frozen=True)
class _Sheet:
    name: str
    rows: tuple[tuple[int, tuple[_Cell, ...]], ...]
    ignored_features: tuple[str, ...] = ()


@dataclass
class ExternalSheetWriteResult:
    sheet_count: int
    cell_count: int
    prepare_seconds: float = 0.0
    write_seconds: float = 0.0
    messages: list[str] = field(default_factory=list)


def _part_from_target(target: str) -> str:
    clean = target.lstrip("/")
    return clean if clean.startswith("xl/") else "xl/" + clean


def _escape_text(value: str) -> str:
    escaped = html.escape(value, quote=False)
    if value[:1].isspace() or value[-1:].isspace():
        return f'<t xml:space="preserve">{escaped}</t>'
    return f"<t>{escaped}</t>"


class DirectOoxmlExternalSheetWriter:
    """将已验证的“纯数据”外部表快速附加到审核副本。"""

    mode = EXTERNAL_SHEET_WRITER_DIRECT_OOXML

    def __init__(self, external_path: Path, sheet_names: tuple[str, ...] | list[str]) -> None:
        self.external_path = Path(external_path)
        self.sheet_names = tuple(sheet_names)
        self._sheets: tuple[_Sheet, ...] = ()
        self.prepare_seconds = 0.0

    def prepare(self) -> ExternalSheetWriteResult:
        started = time.monotonic()
        try:
            with zipfile.ZipFile(self.external_path) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
        except (OSError, zipfile.BadZipFile) as exc:
            raise ExternalSheetWriteError(f"无法读取外部文件：{exc}") from exc
        try:
            workbook = ET.fromstring(parts["xl/workbook.xml"])
            relationships = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
        except (KeyError, ET.ParseError) as exc:
            raise ExternalSheetWriteError(f"外部文件工作簿结构无效：{exc}") from exc
        targets = {item.attrib.get("Id", ""): item.attrib.get("Target", "") for item in relationships}
        found: dict[str, str] = {}
        sheets_root = workbook.find(f"{{{_MAIN_NS}}}sheets")
        for item in sheets_root if sheets_root is not None else ():
            name = item.attrib.get("name", "")
            rid = item.attrib.get(f"{{{_REL_NS}}}id", "")
            if name and rid in targets:
                found[name] = _part_from_target(targets[rid])
        missing = [name for name in self.sheet_names if name not in found]
        if missing:
            raise ExternalSheetWriteError("外部文件缺少工作表：" + "、".join(missing))
        shared = _read_shared_strings(parts.get("xl/sharedStrings.xml", b""))
        self._sheets = tuple(
            _parse_data_sheet(name, parts.get(found[name]), shared)
            for name in self.sheet_names
        )
        self.prepare_seconds = time.monotonic() - started
        ignored = sorted({item for sheet in self._sheets for item in sheet.ignored_features})
        message = f"Direct OOXML 已准备 {len(self._sheets)} 张纯数据外部表"
        if ignored:
            message += "；未复制：" + "、".join(ignored)
        return ExternalSheetWriteResult(
            sheet_count=len(self._sheets),
            cell_count=sum(len(cells) for _row, cells in (row for sheet in self._sheets for row in sheet.rows)),
            prepare_seconds=self.prepare_seconds,
            messages=[message],
        )

    def write(self, audit_path: Path) -> ExternalSheetWriteResult:
        if not self._sheets:
            raise ExternalSheetWriteError("Direct OOXML 外部表尚未准备")
        started = time.monotonic()
        audit_path = Path(audit_path)
        try:
            with zipfile.ZipFile(audit_path) as archive:
                order = list(archive.namelist())
                parts = {name: archive.read(name) for name in order}
        except (OSError, zipfile.BadZipFile) as exc:
            raise ExternalSheetWriteError(f"无法读取审核副本：{exc}") from exc
        self._append_sheets(parts, order)
        temporary = audit_path.with_name(audit_path.stem + ".external.tmp.xlsx")
        try:
            with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
                for part in order:
                    archive.writestr(part, parts[part])
            _validate_output(temporary, self.sheet_names)
            temporary.replace(audit_path)
        finally:
            if temporary.exists():
                temporary.unlink(missing_ok=True)
        elapsed = time.monotonic() - started
        return ExternalSheetWriteResult(
            sheet_count=len(self._sheets),
            cell_count=sum(len(cells) for _row, cells in (row for sheet in self._sheets for row in sheet.rows)),
            prepare_seconds=self.prepare_seconds,
            write_seconds=elapsed,
            messages=[f"Direct OOXML 外部表写入：{len(self._sheets)} 张，{elapsed:.2f} 秒"],
        )

    def _append_sheets(self, parts: dict[str, bytes], order: list[str]) -> None:
        try:
            workbook_xml = parts["xl/workbook.xml"].decode("utf-8")
            rels_xml = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
            content_types_xml = parts["[Content_Types].xml"].decode("utf-8")
        except KeyError as exc:
            raise ExternalSheetWriteError(f"审核副本缺少 OOXML 部件：{exc}") from exc
        existing_names = set(re.findall(r'<sheet\b[^>]*\bname="([^"]+)"', workbook_xml))
        conflicts = [item.name for item in self._sheets if item.name in existing_names]
        if conflicts:
            raise ExternalSheetWriteError("外部文件工作表与报送文件重名，不能覆盖原表：" + "、".join(conflicts))
        sheet_numbers = [int(value) for value in re.findall(r"xl/worksheets/sheet(\d+)\.xml", "\n".join(parts))]
        relation_numbers = [int(value) for value in re.findall(r'\bId="rId(\d+)"', rels_xml)]
        sheet_ids = [int(value) for value in re.findall(r'\bsheetId="(\d+)"', workbook_xml)]
        # openpyxl may omit the relationship prefix when the original workbook
        # has no relationship-bearing workbook child.  We add r:id below, so
        # declare it explicitly instead of producing an invalid workbook.xml.
        root_start = workbook_xml.find("<workbook")
        root_end = workbook_xml.find(">", root_start)
        root_opening = workbook_xml[root_start : root_end + 1] if root_start >= 0 else ""
        if "xmlns:r=" not in root_opening:
            if root_start < 0 or root_end < 0:
                raise ExternalSheetWriteError("审核副本 workbook.xml 根节点无效")
            workbook_xml = (
                workbook_xml[:root_end]
                + f' xmlns:r="{_REL_NS}"'
                + workbook_xml[root_end:]
            )
        for offset, sheet in enumerate(self._sheets, start=1):
            part_number = (max(sheet_numbers) if sheet_numbers else 0) + offset
            relation_number = (max(relation_numbers) if relation_numbers else 0) + offset
            sheet_id = (max(sheet_ids) if sheet_ids else 0) + offset
            part = f"xl/worksheets/sheet{part_number}.xml"
            target = f"worksheets/sheet{part_number}.xml"
            relation_id = f"rId{relation_number}"
            parts[part] = _render_sheet(sheet).encode("utf-8")
            order.append(part)
            escaped_name = html.escape(sheet.name, quote=True)
            workbook_xml = workbook_xml.replace(
                "</sheets>",
                f'<sheet name="{escaped_name}" sheetId="{sheet_id}" r:id="{relation_id}"/></sheets>',
                1,
            )
            rels_xml = rels_xml.replace(
                "</Relationships>",
                f'<Relationship Type="{_REL_TYPE_SHEET}" Target="{target}" Id="{relation_id}"/></Relationships>',
                1,
            )
            content_types_xml = content_types_xml.replace(
                "</Types>",
                f'<Override PartName="/{part}" ContentType="{_CONTENT_TYPE_SHEET}"/></Types>',
                1,
            )
        parts["xl/workbook.xml"] = workbook_xml.encode("utf-8")
        parts["xl/_rels/workbook.xml.rels"] = rels_xml.encode("utf-8")
        parts["[Content_Types].xml"] = content_types_xml.encode("utf-8")


def _read_shared_strings(raw: bytes) -> list[str]:
    if not raw:
        return []
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as exc:
        raise ExternalSheetWriteError(f"外部文件 sharedStrings 无效：{exc}") from exc
    return ["".join(item.itertext()) for item in root.findall(f"{{{_MAIN_NS}}}si")]


def _parse_data_sheet(name: str, raw: bytes | None, shared: list[str]) -> _Sheet:
    if raw is None:
        raise ExternalSheetWriteError(f"外部文件缺少工作表部件：{name}")
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as exc:
        raise ExternalSheetWriteError(f"外部工作表“{name}”XML 无效：{exc}") from exc
    unsupported = {
        "mergeCells": "合并单元格",
        "conditionalFormatting": "条件格式",
        "dataValidations": "数据验证",
        "hyperlinks": "超链接",
        "tableParts": "表格对象",
        "drawing": "图片/绘图",
    }
    present = [label for tag, label in unsupported.items() if root.find(f"{{{_MAIN_NS}}}{tag}") is not None]
    if present:
        # 内层用单引号：f-string 嵌套同类引号是 PEP 701（3.12+）语法，
        # Win7 目标的 Python 3.8 无法编译本文件（真机 EXE 报 No module named）。
        joined = "、".join(present)
        raise UnsupportedExternalSheetOperation(
            f"外部工作表“{name}”含 {joined}；Direct OOXML 仅支持纯数据表，请切换 openpyxl"
        )
    ignored = []
    if root.find(f"{{{_MAIN_NS}}}legacyDrawing") is not None:
        ignored.append("批注")
    data = root.find(f"{{{_MAIN_NS}}}sheetData")
    rows: list[tuple[int, tuple[_Cell, ...]]] = []
    for row in (data if data is not None else ()):
        row_number = int(row.attrib.get("r", "0") or 0)
        cells: list[_Cell] = []
        for cell in row.findall(f"{{{_MAIN_NS}}}c"):
            address = cell.attrib.get("r", "")
            if not address:
                continue
            if cell.find(f"{{{_MAIN_NS}}}f") is not None:
                raise UnsupportedExternalSheetOperation(
                    f"外部工作表“{name}”含公式（{address}）；Direct OOXML 仅支持纯数据表，请切换 openpyxl"
                )
            kind = cell.attrib.get("t", "n")
            if kind == "s":
                value_node = cell.find(f"{{{_MAIN_NS}}}v")
                try:
                    value = shared[int(value_node.text or "")] if value_node is not None else ""
                except (ValueError, IndexError):
                    raise ExternalSheetWriteError(f"外部工作表“{name}”共享字符串索引无效：{address}")
                cells.append(_Cell(address, "text", value))
            elif kind == "inlineStr":
                inline = cell.find(f"{{{_MAIN_NS}}}is")
                cells.append(_Cell(address, "text", "".join(inline.itertext()) if inline is not None else ""))
            elif kind in {"n", "b", "e", "str", "d"}:
                value_node = cell.find(f"{{{_MAIN_NS}}}v")
                value = value_node.text if value_node is not None and value_node.text is not None else ""
                cells.append(_Cell(address, kind, value))
            else:
                raise UnsupportedExternalSheetOperation(
                    f"外部工作表“{name}”含不支持的单元格类型 {kind}（{address}），请切换 openpyxl"
                )
        if cells:
            rows.append((row_number, tuple(cells)))
    return _Sheet(name=name, rows=tuple(rows), ignored_features=tuple(ignored))


def _render_sheet(sheet: _Sheet) -> str:
    rows = []
    for row_number, cells in sheet.rows:
        payload = []
        for cell in cells:
            if cell.kind == "text":
                payload.append(f'<c r="{cell.address}" t="inlineStr"><is>{_escape_text(cell.value)}</is></c>')
            elif cell.kind == "n":
                payload.append(f'<c r="{cell.address}"><v>{html.escape(cell.value)}</v></c>')
            else:
                payload.append(f'<c r="{cell.address}" t="{cell.kind}"><v>{html.escape(cell.value)}</v></c>')
        rows.append(f'<row r="{row_number}">' + "".join(payload) + "</row>")
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{_MAIN_NS}"><sheetData>{"".join(rows)}</sheetData></worksheet>'


def _validate_output(path: Path, names: tuple[str, ...]) -> None:
    try:
        with zipfile.ZipFile(path) as archive:
            if archive.testzip() is not None:
                raise ExternalSheetWriteError("写出的 XLSX ZIP 校验失败")
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            ET.fromstring(archive.read("[Content_Types].xml"))
            rels = {item.attrib.get("Id", ""): item.attrib.get("Target", "") for item in relationships}
            sheets = workbook.find(f"{{{_MAIN_NS}}}sheets")
            actual = {
                item.attrib.get("name", "")
                for item in (sheets if sheets is not None else ())
            }
            missing = [name for name in names if name not in actual]
            if missing:
                raise ExternalSheetWriteError("写出后缺少工作表：" + "、".join(missing))
            for item in sheets if sheets is not None else ():
                if item.attrib.get("name") not in names:
                    continue
                rid = item.attrib.get(f"{{{_REL_NS}}}id", "")
                part = _part_from_target(rels.get(rid, ""))
                if not rid or part not in archive.namelist():
                    raise ExternalSheetWriteError(f"写出后工作表关系缺失：{item.attrib.get('name')}")
                ET.fromstring(archive.read(part))
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
        if isinstance(exc, ExternalSheetWriteError):
            raise
        raise ExternalSheetWriteError(f"写出的 XLSX 结构校验失败：{exc}") from exc
