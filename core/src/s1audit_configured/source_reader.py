"""Request-scoped XLSX values + OOXML metadata, with the rich reader as fallback.

The fast path reads the file once into immutable bytes. Calamine builds each
value matrix once; XML supplies formula/error semantics and structural metadata.
No source is ever saved. A scope is shared across identity, rule and CF consumers.
"""
from __future__ import annotations

from contextlib import contextmanager
from contextvars import ContextVar
from datetime import date, datetime
from functools import wraps
import hashlib
from io import BytesIO
import os
from pathlib import Path
import posixpath
from time import perf_counter
from types import SimpleNamespace
import warnings
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.reader.strings import read_string_table
from openpyxl.formatting.formatting import ConditionalFormatting
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format, is_timedelta_format
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900, from_excel, from_ISO8601
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

try:
    from python_calamine import CalamineWorkbook
except Exception:  # Missing module or a binary-loader failure must both fall back.
    CalamineWorkbook = None

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_POOL = ContextVar("audit3_source_readers", default=None)
_FILE_HASHING_ENABLED = ContextVar("audit3_file_hashing_enabled", default=True)


@contextmanager
def file_hashing(enabled: bool):
    """Temporarily enable or disable content hashing of workbook files."""
    token = _FILE_HASHING_ENABLED.set(bool(enabled))
    try:
        yield
    finally:
        _FILE_HASHING_ENABLED.reset(token)


def file_hashing_enabled() -> bool:
    return _FILE_HASHING_ENABLED.get()


def payload_sha256(payload: bytes) -> str:
    if not file_hashing_enabled():
        return ""
    digest = _new_file_digest()
    digest.update(payload)
    return digest.hexdigest()


def file_sha256(path) -> str:
    if not file_hashing_enabled():
        return ""
    digest = _new_file_digest()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _new_file_digest():
    return hashlib.sha256()


def _cast_numeric(raw):
    return float(raw) if any(char in raw for char in ".Ee") else int(raw)


def normalize_fast_value(value):
    if isinstance(value, float) and value.is_integer() and abs(value) <= 2**53:
        return int(value)
    if value == "":
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return datetime.combine(value, datetime.min.time())
    return value


class ReaderCell:
    __slots__ = ("row", "column", "value", "data_type", "number_format")

    def __init__(self, row, column, value, data_type, number_format="General"):
        self.row, self.column, self.value = row, column, value
        self.data_type, self.number_format = data_type, number_format

    @property
    def coordinate(self):
        return f"{get_column_letter(self.column)}{self.row}"


class ReaderSheet:
    def __init__(self, metadata, data_only):
        self._metadata, self._data_only = metadata, data_only
        self.title = metadata.title
        self.max_row, self.max_column = metadata.max_row, metadata.max_column
        self.merged_cells = SimpleNamespace(ranges=metadata.merges)
        self.defined_names = metadata.defined_names
        self.conditional_formatting = metadata.conditional_formats

    def cell(self, row, column):
        if row < 1 or column < 1:
            raise ValueError("Row and column values must be at least 1")
        meta, key = self._metadata, (row, column)
        matrix = meta.matrix
        value = matrix[row - 1][column - 1] if row <= len(matrix) and column <= len(matrix[row - 1]) else None
        value = meta.overrides.get(key, value)
        data_type = meta.types.get(key, "n")
        if key in meta.formulas:
            if not self._data_only:
                value, data_type = meta.formulas[key], "f"
            else:
                data_type = meta.cached_types.get(key, "n")
        # Excel ignores non-anchor payloads inside merged cells, as does openpyxl.
        if key in meta.merged_nonanchors:
            value, data_type = None, "n"
            number_format = "General"
        else:
            number_format = meta.formats.get(key, "General")
        return ReaderCell(row, column, value, data_type, number_format)

    def __getitem__(self, address):
        row, col = coordinate_to_tuple(address.replace("$", ""))
        return self.cell(row, col)

    def iter_rows(self, min_row=1, max_row=None, min_col=1, max_col=None, values_only=False):
        for row in range(min_row, (self.max_row if max_row is None else max_row) + 1):
            cells = tuple(self.cell(row, col) for col in range(min_col, (self.max_column if max_col is None else max_col) + 1))
            yield tuple(cell.value for cell in cells) if values_only else cells

    @property
    def values(self):
        return self.iter_rows(values_only=True)


class ReaderBook:
    def __init__(self, reader, data_only):
        self._reader = reader
        self.sheetnames = list(reader.metadata)
        self.defined_names = reader.defined_names
        self.worksheets = [ReaderSheet(meta, data_only) for meta in reader.metadata.values()]
        self._sheets = {sheet.title: sheet for sheet in self.worksheets}

    def __getitem__(self, name):
        return self._sheets[name]

    def __iter__(self):
        return iter(self.worksheets)

    def close(self):
        # The owning scope closes pooled readers after all four consumers finish.
        if not self._reader.pooled:
            self._reader.close()


class SourceReader:
    def __init__(self, path, *, engine=None):
        started = perf_counter()
        self.path = Path(path).resolve()
        self.payload = self.path.read_bytes()
        self.sha256 = payload_sha256(self.payload)
        self.pooled = False
        self._books = {}
        self.metadata = {}
        self.fallback_reason = ""
        self.rich_loads = 0
        self.engine = engine or os.environ.get("AUDIT3_READER_ENGINE", "mixed")
        if self.engine not in {"mixed", "openpyxl"}:
            raise ValueError("AUDIT3_READER_ENGINE must be mixed or openpyxl")
        if self.engine == "mixed":
            try:
                if CalamineWorkbook is None or os.environ.get("AUDIT3_DISABLE_CALAMINE") == "1":
                    raise ImportError("python-calamine 不可用")
                self._load_mixed()
            except Exception as exc:
                self.engine = "openpyxl"
                self.metadata.clear()
                self.fallback_reason = f"{type(exc).__name__}: {exc}"
                warnings.warn(f"{self.path.name}：快速读取失败，已回退 openpyxl（{self.fallback_reason}）", RuntimeWarning, stacklevel=2)
        if self.engine == "openpyxl":
            self._books[False] = load_workbook(BytesIO(self.payload), data_only=False, read_only=False)
            self.rich_loads += 1
        self.seconds = perf_counter() - started

    def _load_mixed(self):
        self.defined_names = {}
        with ZipFile(BytesIO(self.payload)) as archive:
            root = ET.fromstring(archive.read("xl/workbook.xml"))
            properties = root.find(NS + "workbookPr")
            epoch = CALENDAR_MAC_1904 if properties is not None and properties.get("date1904") in {"1", "true"} else CALENDAR_WINDOWS_1900
            relations = {node.get("Id"): node.get("Target") for node in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))}
            sheets = []
            for node in root.find(NS + "sheets"):
                target = relations[node.get(REL)]
                member = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
                if not member.startswith("xl/"):
                    raise ValueError("工作表关系超出 xl/ 目录")
                sheets.append((node.get("name"), member))
            styles, dxfs = self._styles(archive)
            shared_strings = []
            if "xl/sharedStrings.xml" in archive.namelist():
                # Match XML newline/OOXML escape and rich-text normalization;
                # calamine can retain CRLF that the old XML parser normalizes.
                shared_strings = read_string_table(BytesIO(archive.read("xl/sharedStrings.xml")))
            with CalamineWorkbook.from_filelike(BytesIO(self.payload)) as values:
                for title, member in sheets:
                    sheet = values.get_sheet_by_name(title)
                    # Both leading rows AND columns are trimmed by to_python().
                    top, left = sheet.start if isinstance(sheet.start, tuple) else (0, 0)
                    matrix = [[] for _ in range(top)] + [
                        [None] * left + [normalize_fast_value(value) for value in row]
                        for row in sheet.to_python()
                    ]
                    meta = SimpleNamespace(title=title, matrix=matrix, formulas={}, types={}, cached_types={},
                                           overrides={}, formats={}, merges=[], merged_nonanchors=set(),
                                           conditional_formats=[], defined_names={}, max_row=1, max_column=1)
                    self._sheet_metadata(archive, member, meta, styles, dxfs, shared_strings, epoch)
                    self.metadata[title] = meta
            definitions = root.find(NS + "definedNames")
            for node in definitions if definitions is not None else []:
                definition = DefinedName.from_tree(node)
                if definition.localSheetId is None:
                    self.defined_names[definition.name] = definition
                elif definition.localSheetId < len(sheets):
                    self.metadata[sheets[definition.localSheetId][0]].defined_names[definition.name] = definition

    @staticmethod
    def _styles(archive):
        if "xl/styles.xml" not in archive.namelist():
            return ["General"], []
        root = ET.fromstring(archive.read("xl/styles.xml"))
        formats = dict(BUILTIN_FORMATS)
        custom = root.find(NS + "numFmts")
        for item in custom if custom is not None else []:
            formats[int(item.get("numFmtId"))] = item.get("formatCode")
        xfs = root.find(NS + "cellXfs")
        styles = [formats.get(int(item.get("numFmtId", 0)), "General") for item in xfs] if xfs is not None else ["General"]
        dxfs = root.find(NS + "dxfs")
        return styles, [DifferentialStyle.from_tree(item) for item in dxfs] if dxfs is not None else []

    @staticmethod
    def _sheet_metadata(archive, member, meta, styles, dxfs, shared_strings, epoch):
        shared, followers = {}, []
        hyperlinks = []
        with archive.open(member) as stream:
            for _event, node in ET.iterparse(stream, events=("end",)):
                if node.tag == NS + "c":
                    address = node.get("r")
                    if not address:
                        raise ValueError("单元格缺少坐标")
                    key = coordinate_to_tuple(address)
                    meta.max_row, meta.max_column = max(meta.max_row, key[0]), max(meta.max_column, key[1])
                    kind = node.get("t", "n")
                    meta.types[key] = kind
                    fmt = styles[int(node.get("s", 0))]
                    if fmt != "General":
                        meta.formats[key] = fmt
                    value_node, formula = node.find(NS + "v"), node.find(NS + "f")
                    raw = value_node.text if value_node is not None else None
                    if kind == "e":
                        meta.overrides[key] = raw
                    elif kind == "b" and raw is not None:
                        meta.overrides[key] = bool(int(raw))
                    elif kind == "d" and raw is not None:
                        meta.overrides[key] = from_ISO8601(raw)
                    elif kind in {"n", ""} and raw is not None:
                        # Preserve openpyxl's lexical int/float distinction and exact
                        # large integers; calamine uses float for both in XLSX.
                        value = _cast_numeric(raw)
                        if is_date_format(fmt):
                            try:
                                value = from_excel(value, epoch, timedelta=is_timedelta_format(fmt))
                                kind = "d"
                                meta.types[key] = kind
                            except (ValueError, OverflowError):
                                value, kind = "#VALUE!", "e"
                                meta.types[key] = kind
                        meta.overrides[key] = value
                    elif kind == "inlineStr":
                        inline = node.find(NS + "is")
                        if inline is not None:
                            meta.overrides[key] = "".join(item.text or "" for item in inline.iter(NS + "t"))
                        meta.types[key] = "s" if inline is not None else "inlineStr"
                    elif kind == "s" and raw is not None:
                        expected = shared_strings[int(raw)]
                        matrix = meta.matrix
                        actual = matrix[key[0]-1][key[1]-1] if key[0] <= len(matrix) and key[1] <= len(matrix[key[0]-1]) else None
                        if expected != actual:
                            meta.overrides[key] = expected
                    elif kind == "str":
                        meta.overrides[key] = raw
                        meta.types[key] = "s" if raw is not None else "str"
                    if formula is not None:
                        formula_kind, index = formula.get("t"), formula.get("si")
                        text = "=" + (formula.text or "")
                        if formula_kind == "shared":
                            if formula.text:
                                shared[index] = (address, text)
                            else:
                                followers.append((key, address, index))
                        elif formula_kind == "array":
                            text = ArrayFormula(ref=formula.get("ref"), text=text)
                        elif formula_kind == "dataTable":
                            text = DataTableFormula(**formula.attrib)
                        meta.formulas[key] = text
                        meta.cached_types[key] = "s" if kind == "str" and raw is not None else kind
                        if raw is None:
                            meta.overrides[key] = None
                    node.clear()
                elif node.tag == NS + "row":
                    node.clear()
                elif node.tag == NS + "hyperlink":
                    hyperlinks.append(dict(node.attrib))
                elif node.tag == NS + "mergeCell":
                    area = CellRange(node.get("ref"))
                    meta.merges.append(area)
                    meta.max_row, meta.max_column = max(meta.max_row, area.max_row), max(meta.max_column, area.max_col)
                    meta.merged_nonanchors.update((r, c) for r in range(area.min_row, area.max_row + 1)
                                                 for c in range(area.min_col, area.max_col + 1)
                                                 if (r, c) != (area.min_row, area.min_col))
                    node.clear()
                elif node.tag == NS + "conditionalFormatting":
                    cf = ConditionalFormatting.from_tree(node)
                    for rule in cf.rules:
                        if rule.dxfId is not None:
                            rule.dxf = dxfs[rule.dxfId]
                    # openpyxl combines repeated elements with identical sqref.
                    previous = next((item for item in meta.conditional_formats if item.sqref == cf.sqref), None)
                    if previous is not None:
                        previous.rules.extend(cf.rules)
                    else:
                        meta.conditional_formats.append(cf)
                    node.clear()
        for key, address, index in followers:
            origin, formula = shared[index]
            meta.formulas[key] = Translator(formula, origin=origin).translate_formula(address)
        # Rich openpyxl creates cells for comments and hyperlinks even when no
        # <c> exists. Preserve those dimensions and hyperlink placeholder values.
        folder, filename = posixpath.split(member)
        rel_member = posixpath.join(folder, "_rels", filename + ".rels")
        relations = {}
        if rel_member in archive.namelist():
            relations = {node.get("Id"): node for node in ET.fromstring(archive.read(rel_member))}
        for relation in relations.values():
            if relation.get("Type", "").endswith("/comments"):
                target = relation.get("Target")
                target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
                for comment in ET.fromstring(archive.read(target)).iter(NS + "comment"):
                    row, col = coordinate_to_tuple(comment.get("ref"))
                    meta.max_row, meta.max_column = max(meta.max_row, row), max(meta.max_column, col)
        for link in hyperlinks:
            area = CellRange(link["ref"])
            meta.max_row, meta.max_column = max(meta.max_row, area.max_row), max(meta.max_column, area.max_col)
            relation = relations.get(link.get(REL))
            target = relation.get("Target") if relation is not None else link.get("location")
            for row in range(area.min_row, area.max_row + 1):
                for col in range(area.min_col, area.max_col + 1):
                    key = row, col
                    if key in meta.formulas or key in meta.merged_nonanchors:
                        continue
                    matrix = meta.matrix
                    value = matrix[row-1][col-1] if row <= len(matrix) and col <= len(matrix[row-1]) else None
                    if meta.overrides.get(key, value) is None:
                        meta.overrides[key], meta.types[key] = target, "s"

    def book(self, *, data_only=False):
        if data_only not in self._books:
            started = perf_counter()
            self._books[data_only] = (ReaderBook(self, data_only) if self.engine == "mixed" else
                                     load_workbook(BytesIO(self.payload), data_only=data_only, read_only=False, keep_links=False))
            if self.engine == "openpyxl":
                self.rich_loads += 1
                self.seconds += perf_counter() - started
        return self._books[data_only]

    def diagnostic(self):
        return {"path": str(self.path), "engine": self.engine, "file_reads": 1,
                "value_matrices": 1 if self.engine == "mixed" else 0,
                "rich_loads": self.rich_loads,
                "seconds": self.seconds, "fallback_reason": self.fallback_reason}

    def close(self):
        for book in self._books.values():
            if not isinstance(book, ReaderBook):
                book.close()
        self._books.clear()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()


@contextmanager
def reader_session():
    if _POOL.get() is not None:
        yield _POOL.get()
        return
    pool = {}
    token = _POOL.set(pool)
    try:
        yield pool
    finally:
        for reader in pool.values():
            reader.close()
        _POOL.reset(token)


def shared_readers(function):
    @wraps(function)
    def wrapped(*args, **kwargs):
        with reader_session():
            return function(*args, **kwargs)
    return wrapped


def get_reader(path):
    path = Path(path).resolve()
    pool = _POOL.get()
    stat = path.stat()
    key = (str(path), stat.st_size, stat.st_mtime_ns)
    if pool is None:
        return SourceReader(path)
    if key not in pool:
        pool[key] = SourceReader(path)
        pool[key].pooled = True
    return pool[key]


def read_book(path, *, data_only=False):
    return get_reader(path).book(data_only=data_only)


def reader_diagnostics():
    return [reader.diagnostic() for reader in (_POOL.get() or {}).values()]
