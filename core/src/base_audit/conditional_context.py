"""条件格式统一结果构建层：Context / CommentReader / MessageResolver / Builder。

架构分工（Evaluator 只回答「是否触发、哪个格、命中哪些规则」）：

- :class:`ConditionalFormatContext`：每个工作簿构建一次，缓存表结构区域、
  合并格索引（经 :class:`~base_audit.indicator_resolver.IndicatorResolver`）、
  批注索引；静态信息一次读取、一次建索引、多次复用。
- :class:`OpenpyxlCommentReader` / :class:`OoxmlCommentReader`：批注读取的
  两个后端（openpyxl 对象模型 / Direct OOXML 直读 comments*.xml），产出
  相同的 :class:`CellComment`。传统批注（Traditional Note）为已支持口径；
  现代批注（Threaded Comments）明确记录为不支持，不假装支持。
- :class:`MessageResolver`：正式业务规则「批注优先、公式兜底」——批注存在
  且正文非空时 message=批注原文（保持既有「作者:
正文」格式），否则
  ``条件格式规则：<公式>``。这是业务规则，不是 Evaluator 策略。
- :func:`build_conditional_issues`：唯一的 Triggered Cell → Issue 组装器；
  三条求值路径（PYTHON / COM_EVALUATE / COM_DISPLAY）共用，除 source 等
  技术来源字段外业务输出一致。

指标定位语义以已验证的 COM ``_conditional_check_field`` 为金标准，经
:class:`IndicatorResolver` 纯静态复现（openpyxl 合并锚点 + 结构区域左扫/
上扫），不依赖 COM。
"""

from __future__ import annotations

import re
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as StdET

from .indicator_resolver import IndicatorResolver
from .models import CopyRange, Issue


# ---------------------------------------------------------------------------
# 批注模型与读取器
# ---------------------------------------------------------------------------


@dataclass
class CellComment:
    """标准化后的单元格批注（三后端统一产出）。"""

    cell: str            # "C26"
    author: str          # "l z b"
    text: str            # 批注正文（去掉作者行）："这是一个测试"
    formatted: str       # 既有展示格式原文（"作者:\n正文"），message 直接使用
    comment_type: str = "note"   # note / threaded（threaded 目前不支持，仅记录）


_AUTHOR_LINE_RE = re.compile(r"^([^\n:]{0,60}):\s*\n")


def _standardize_comment(cell: str, raw_text: str, author: str) -> CellComment:
    """拆分「作者行 + 正文」；既有业务格式就是 ``作者:\n正文``，保持不变。"""
    raw = str(raw_text or "").strip()
    matched = _AUTHOR_LINE_RE.match(raw)
    if matched and matched.group(1).strip():
        comment_author = matched.group(1).strip()
        body = raw[matched.end():].strip()
    else:
        comment_author = str(author or "").strip()
        body = raw
    return CellComment(cell=cell, author=comment_author, text=body,
                       formatted=raw, comment_type="note")


class OpenpyxlCommentReader:
    """openpyxl 对象模型批注读取（静态簿已在内存时的零 I/O 后端）。"""

    def __init__(self, static_book: Any) -> None:
        self._book = static_book
        self._cache: dict[tuple[str, str], CellComment | None] = {}

    def comment_for(self, sheet_name: str, cell: str) -> CellComment | None:
        key = (sheet_name, cell)
        if key in self._cache:
            return self._cache[key]
        result: CellComment | None = None
        try:
            if sheet_name in self._book.sheetnames:
                row = int("".join(ch for ch in cell if ch.isdigit()) or 0)
                letters = "".join(ch for ch in cell if ch.isalpha()).upper()
                column = 0
                for ch in letters:
                    column = column * 26 + (ord(ch) - 64)
                comment = self._book[sheet_name].cell(row=row, column=column).comment
                if comment is not None and str(comment.text or "").strip():
                    result = _standardize_comment(
                        cell, str(comment.text or ""), str(comment.author or ""))
        except Exception:
            result = None
        self._cache[key] = result
        return result


class OoxmlCommentReader:
    """Direct OOXML 传统批注读取：xl/comments*.xml + 工作表关系映射。

    一次读 ZIP、一次建索引（sheet 名 → cell → CellComment）。只支持传统
    批注（Traditional Note）；现代批注（Threaded Comments）出现时记录为
    不支持，不参与 message 生成。
    """

    def __init__(self, workbook_path: Path | str) -> None:
        self.workbook_path = Path(workbook_path)
        self.unsupported_threaded: list[str] = []

    def read(self) -> dict[str, dict[str, CellComment]]:
        result: dict[str, dict[str, CellComment]] = {}
        with zipfile.ZipFile(self.workbook_path) as archive:
            names = set(archive.namelist())
            sheet_parts = self._sheet_parts(archive)
            for sheet_name, part in sheet_parts.items():
                part_path = self._comments_part(archive, part, names)
                if part_path is None:
                    continue
                result[sheet_name] = self._parse_comments(part_path, archive)
            for name in sorted(names):
                if name.startswith("xl/threadedComments/"):
                    self.unsupported_threaded.append(name)
        return result

    @staticmethod
    def _sheet_parts(archive: zipfile.ZipFile) -> dict[str, str]:
        """workbook.xml + rels → {工作表名: 部件路径}。"""
        rels_xml = archive.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = {}
        for rel in StdET.fromstring(rels_xml):
            rid, target = rel.get("Id") or "", rel.get("Target") or ""
            if rid and target:
                targets[rid] = target
        result: dict[str, str] = {}
        for sheet in StdET.fromstring(archive.read("xl/workbook.xml")).iter():
            if sheet.tag.split("}")[-1] != "sheet":
                continue
            name = sheet.get("name") or ""
            rid = ""
            for key, value in sheet.attrib.items():
                if key.endswith("}id"):
                    rid = value
                    break
            target = targets.get(rid, "")
            if name and target:
                result[name] = target.lstrip("/") if target.startswith("/") else "xl/" + target
        return result

    @staticmethod
    def _comments_part(archive: zipfile.ZipFile, sheet_part: str,
                       names: set[str]) -> str | None:
        """工作表关系 → 传统批注部件路径（无则 None）。"""
        head, tail_name = sheet_part.rsplit("/", 1)
        rels_path = f"{head}/_rels/{tail_name}.rels"
        if rels_path not in names:
            return None
        for rel in StdET.fromstring(archive.read(rels_path)):
            if not (rel.get("Type") or "").endswith("/comments"):
                continue
            target = (rel.get("Target") or "").replace("\\", "/").lstrip("/")
            if target.startswith("xl/"):
                return target
            # 相对路径（如 ../comments1.xml）以工作表部件目录为基解析
            stack: list[str] = []
            for segment in f"{head}/{target}".split("/"):
                if segment == "..":
                    if stack:
                        stack.pop()
                elif segment in (".", ""):
                    continue
                else:
                    stack.append(segment)
            resolved = "/".join(stack)
            if resolved in names:
                return resolved
        return None

    def _parse_comments(self, part: str,
                        archive: zipfile.ZipFile) -> dict[str, CellComment]:
        root = StdET.fromstring(archive.read(part))
        authors: list[str] = []
        for element in root:
            if element.tag.split("}")[-1] == "authors":
                authors = [(item.text or "") for item in element]
        result: dict[str, CellComment] = {}
        for element in root:
            if element.tag.split("}")[-1] != "commentList":
                continue
            for comment in element:
                if comment.tag.split("}")[-1] != "comment":
                    continue
                cell = comment.get("ref") or ""
                if not cell:
                    continue
                try:
                    author = authors[int(comment.get("authorId") or 0)]
                except (ValueError, IndexError):
                    author = ""
                text = "".join(
                    node.text or "" for node in comment.iter()
                    if node.tag.split("}")[-1] == "t")
                if str(text or "").strip():
                    result[cell] = _standardize_comment(cell, text, author)
        return result


# ---------------------------------------------------------------------------
# 工作簿级 Context：一次构建，多次复用
# ---------------------------------------------------------------------------


class ConditionalFormatContext:
    """每工作簿一次的静态上下文：结构区域 + 合并格索引 + 批注索引。"""

    def __init__(self, static_book: Any, structure_ranges: Iterable[CopyRange],
                 resolver: IndicatorResolver, comment_backend: str,
                 comments: dict[str, dict[str, CellComment]],
                 unsupported_notes: list[str]) -> None:
        self.static_book = static_book
        self.structure_ranges = list(structure_ranges or ())
        self.resolver = resolver
        self.comment_backend = comment_backend
        self._comments = comments
        self.unsupported_notes = unsupported_notes
        self._comment_cache: dict[tuple[str, str], CellComment | None] = {}
        self.context_build_seconds = 0.0
        self.comment_read_seconds = 0.0

    @classmethod
    def build(cls, *, workbook_path: Path | None = None, static_book: Any,
              structure_ranges: Iterable[CopyRange]) -> "ConditionalFormatContext":
        started = time.perf_counter()
        comments: dict[str, dict[str, CellComment]] = {}
        unsupported: list[str] = []
        backend = "openpyxl"
        if workbook_path is not None:
            try:
                reader = OoxmlCommentReader(workbook_path)
                comments = reader.read()
                backend = "direct_ooxml"
                if reader.unsupported_threaded:
                    unsupported.append(
                        "现代批注（Threaded Comments）未支持，已忽略："
                        + "、".join(reader.unsupported_threaded))
            except Exception:
                comments, backend = {}, "openpyxl"
        resolver = IndicatorResolver(static_book, structure_ranges)
        context = cls(static_book, structure_ranges, resolver, backend,
                      comments, unsupported)
        context.context_build_seconds = time.perf_counter() - started
        return context

    def comment_for(self, sheet_name: str, cell: str) -> CellComment | None:
        started = time.perf_counter()
        try:
            key = (sheet_name, cell)
            if key in self._comment_cache:
                return self._comment_cache[key]
            result = self._comments.get(sheet_name, {}).get(cell)
            self._comment_cache[key] = result
            return result
        finally:
            self.comment_read_seconds += time.perf_counter() - started


# ---------------------------------------------------------------------------
# MessageResolver：批注优先、公式兜底（业务规则，非 Evaluator 策略）
# ---------------------------------------------------------------------------


class MessageResolver:
    """``有批注且正文非空 → 批注原文；否则 → 条件格式规则：<公式>``。"""

    @staticmethod
    def resolve(comment: CellComment | None, fallback_label: str) -> str:
        if comment is not None and comment.text.strip():
            return comment.formatted
        return fallback_label


# ---------------------------------------------------------------------------
# AuditResultBuilder：唯一的 Triggered Cell → Issue 组装器
# ---------------------------------------------------------------------------


def build_conditional_issues(
    results: list[Any], *, context: ConditionalFormatContext,
    ranges: list[CopyRange], workbook_path: Path, source_file: Path,
    period: str, batch_id: str, audit_time: str,
    org_code: str = "", org_name: str = "",
) -> tuple[list[Issue], dict]:
    """把触发结果组装为业务 Issue（三条求值路径共用同一实现）。

    ``results`` 的元素需具备 ``sheet/cell/triggered/message``；message 仅作
    为无批注时的兜底文案（``条件格式规则：<公式>`` 口径由各求值路径传入）。
    返回 ``(issues, build_stats)``。
    """
    from .template import normalize_template_name

    started = time.perf_counter()
    book = context.static_book
    by_cell = {}
    for hit in results:
        if not hit.triggered:
            continue
        row = int("".join(ch for ch in hit.cell if ch.isdigit()) or 0)
        letters = "".join(ch for ch in hit.cell if ch.isalpha()).upper()
        column = 0
        for ch in letters:
            column = column * 26 + (ord(ch) - 64)
        by_cell[(hit.sheet, row, column)] = hit

    issues: list[Issue] = []
    seen: set[tuple[str, int, int]] = set()
    workbook_key = normalize_template_name(source_file.stem) or source_file.stem
    message_seconds = 0.0
    from openpyxl.utils.cell import range_boundaries
    from openpyxl.utils import get_column_letter

    for area in ranges:
        if area.sheet_name not in book.sheetnames:
            continue
        baseline = book[area.sheet_name]
        min_col, min_row, max_col, max_row = range_boundaries(area.address)
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                if (area.sheet_name, row, col) in seen:
                    continue
                seen.add((area.sheet_name, row, col))
                hit = by_cell.get((area.sheet_name, row, col))
                if hit is None:
                    continue
                address = f"{get_column_letter(col)}{row}"
                indicator = context.resolver.resolve(area.sheet_name, row, col)
                comment = context.comment_for(area.sheet_name, address)
                message_started = time.perf_counter()
                try:
                    message = MessageResolver.resolve(comment, str(hit.message or ""))
                finally:
                    message_seconds += time.perf_counter() - message_started
                detail = message
                identity = "｜".join(
                    (workbook_key, area.sheet_name, address,
                     indicator or "校验指标未识别"))
                issues.append(Issue(
                    identity,
                    period, batch_id, audit_time, True, "", "", "", 1,
                    org_code, org_name, "",
                    area.sheet_name, "条件格式填充", "条件格式触发",
                    address, address,
                    baseline.cell(row, col).value, "", message,
                    str(source_file), str(workbook_path),
                    check_field=indicator, detail=detail,
                ))
    build_seconds = time.perf_counter() - started
    stats = {
        "indicator_resolve_time": context.resolver.resolve_seconds,
        "comment_read_time": context.comment_read_seconds,
        "message_build_time": message_seconds,
        "result_build_time": build_seconds,
        "comment_backend": context.comment_backend,
        "unsupported_notes": list(context.unsupported_notes),
    }
    return issues, stats
