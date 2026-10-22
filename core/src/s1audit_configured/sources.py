from __future__ import annotations

import hashlib
import hmac
import json
import re
import secrets
from decimal import Decimal, InvalidOperation
from pathlib import Path
from collections import defaultdict
from dataclasses import dataclass
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .errors import ConfigurationError, RetrievalError
from .external_layout import discover_external_layout, normalized
from .layout import NamedRegion as _NamedRegion, in_numeric_regions, named_regions, row_path_candidates, scan_rule_paths, split_path
from .model import (
    Configuration, EvalValue, Evidence, ExternalMetricSpec, MetricSpec,
    RuleDefinition, RuntimeContext, SourceBinding, SourceContract,
)
from .source_reader import get_reader


_AMOUNT_FACTORS = {"元": Decimal(1), "万元": Decimal(10000), "亿元": Decimal(100000000)}
from .runtime import RuntimePaths, default_data_dir

DEFAULT_SCHEMA_SNAPSHOT_DIR = default_data_dir() / ".schema_snapshots"


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _norm(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "是" if value else "否"
    return str(value).strip()


def _path_digest(secret: bytes, path: tuple[str, ...]) -> str:
    """Opaque structural token; raw workbook labels never enter the snapshot JSON."""
    payload = json.dumps(list(path), ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return hmac.new(secret, payload, hashlib.sha256).hexdigest()


@dataclass
class _TemplatePathIndex:
    locations: dict[tuple[tuple[str, ...], tuple[str, ...]], list[tuple[int, int]]]
    row_locations: dict[tuple[str, ...], list[int]]
    column_locations: dict[tuple[str, ...], list[int]]
    tables: list[dict[str, Any]]


class _Workbook:
    def __init__(self, binding: SourceBinding, *, editable: bool = False, template_path=None, template_dir=None):
        self.template_path = template_path
        self.template_dir = template_dir
        self.binding = binding
        self.path = Path(binding.path).resolve()
        if not self.path.is_file():
            raise RetrievalError("SOURCE_READ_FAILED", "数据源文件不存在")
        try:
            if editable:
                # Annotated template/expression copies need the full editing
                # model. Audit inputs use the shared immutable reader below.
                self.sha256 = file_sha256(self.path)
                self.book = load_workbook(self.path, data_only=False, read_only=False)
            else:
                reader = get_reader(self.path)
                self.sha256 = reader.sha256
                self.book = reader.book(data_only=False)
        except Exception as exc:
            raise RetrievalError("SOURCE_READ_FAILED", f"数据源工作簿无法读取：{exc}") from exc
        self._merged_by_row: dict[str, dict[int, list[tuple[int, int, int, int]]]] = defaultdict(
            lambda: defaultdict(list)
        )
        self._structure_regions: dict[str, list[_NamedRegion]] = defaultdict(list)
        self._numeric_regions: dict[str, list[_NamedRegion]] = {}
        self.numeric_layout_dependencies: dict[Path, str] = {}
        for sheet in self.book.worksheets:
            for merged in sheet.merged_cells.ranges:
                min_col, min_row, max_col, max_row = merged.bounds
                for row in range(min_row, max_row + 1):
                    self._merged_by_row[sheet.title][row].append((min_col, max_col, min_row, min_col))
        for rows in self._merged_by_row.values():
            for ranges in rows.values():
                ranges.sort(key=lambda item: item[0])
        self._load_structure_regions()

    def close(self) -> None:
        self.book.close()

    def value(self, sheet_name: str, row: int, col: int) -> tuple[Any, str, int, int]:
        if sheet_name not in self.book.sheetnames:
            raise RetrievalError("SHEET_NOT_FOUND", "数据源缺少配置指定的工作表", {"sheet": sheet_name})
        sheet = self.book[sheet_name]
        anchor_row, anchor_col = self._anchor(sheet_name, row, col)
        cell = sheet.cell(anchor_row, anchor_col)
        if cell.data_type == "f":
            raise RetrievalError(
                "SOURCE_FORMULA_UNVERIFIED", "源数据目标单元格包含公式，未确认其重算缓存，不能作为审核输入",
                {"sheet": sheet_name, "cell": cell.coordinate},
            )
        if cell.data_type == "e":
            raise RetrievalError("VALUE_TYPE_INVALID", "源数据单元格包含 Excel 错误值", {"sheet": sheet_name, "cell": cell.coordinate})
        return cell.value, cell.coordinate, anchor_row, anchor_col

    def cell_raw(self, sheet_name: str, row: int, col: int) -> Any:
        if sheet_name not in self.book.sheetnames:
            raise RetrievalError("SHEET_NOT_FOUND", "数据源缺少配置指定的工作表", {"sheet": sheet_name})
        # Formula text is safe to inspect while locating labels. The actual referenced
        # value cell is checked by value() and rejected unless a trusted cache exists.
        return self.book[sheet_name].cell(row, col).value

    def merged_value(self, sheet_name: str, row: int, col: int) -> Any:
        anchor_row, anchor_col = self._anchor(sheet_name, row, col)
        return self.cell_raw(sheet_name, anchor_row, anchor_col)

    def header_path(self, sheet_name: str, coordinates: list[tuple[int, int]]) -> tuple[str, ...]:
        """Read a header path, collapsing repeats caused by one merged cell."""
        values: list[str] = []
        previous_anchor: tuple[int, int] | None = None
        for row, col in coordinates:
            anchor = self._anchor(sheet_name, row, col)
            value = _norm(self.cell_raw(sheet_name, *anchor))
            if anchor != previous_anchor and value:
                values.append(value)
            previous_anchor = anchor
        return tuple(values)

    def structure_regions(self, sheet_name: str) -> list[_NamedRegion]:
        return list(self._structure_regions.get(sheet_name, ()))

    def data_regions(self, sheet_name: str) -> list[_NamedRegion]:
        if sheet_name not in self._numeric_regions:
            from .numeric_scope import resolve_numeric_regions
            regions, dependency = resolve_numeric_regions(self, sheet_name)
            self._numeric_regions[sheet_name] = regions
            if dependency:
                self.numeric_layout_dependencies[dependency[0]] = dependency[1]
        return self._numeric_regions[sheet_name]

    def _load_structure_regions(self) -> None:
        for region in named_regions(self.book, ("表结构区域", "行索引区域", "列索引区域")):
            self._structure_regions[region.sheet].append(region)
        for regions in self._structure_regions.values():
            regions.sort(key=lambda item: (_region_number(item.name), item.min_col, item.min_row, item.name))

    def _anchor(self, sheet_name: str, row: int, col: int) -> tuple[int, int]:
        for min_col, max_col, anchor_row, anchor_col in self._merged_by_row.get(sheet_name, {}).get(row, ()):
            if min_col <= col <= max_col:
                return anchor_row, anchor_col
        return row, col


def _region_number(name: str) -> int:
    match = re.search(r"_(\d+)(?:#\d+)?$", name)
    return int(match.group(1)) if match else 1_000_000_000


def _merge_shapes(sheet: Any, region: _NamedRegion) -> list[str]:
    shapes: list[str] = []
    for merged in sheet.merged_cells.ranges:
        min_col, min_row, max_col, max_row = merged.bounds
        if (region.min_col <= min_col <= max_col <= region.max_col
                and region.min_row <= min_row <= max_row <= region.max_row):
            shapes.append(
                f"R{min_row - region.min_row}C{min_col - region.min_col}:"
                f"R{max_row - region.min_row}C{max_col - region.min_col}"
            )
    return sorted(shapes)


class SourceRepository:
    """Read-only source access, path retrieval and provenance collection."""

    def __init__(self, configuration: Configuration, context: RuntimeContext,
                 *, schema_snapshot_dir: str | Path | None = None, runtime_paths: RuntimePaths | None = None):
        self.runtime_paths = runtime_paths or RuntimePaths()
        self.configuration = configuration
        self.context = context
        self._books: dict[tuple[str, str], _Workbook] = {}
        self._retrieval_cache: dict[tuple[Any, ...], EvalValue | RetrievalError] = {}
        self._external_layouts = {}
        self.data_quality_warnings: list[dict[str, str]] = []
        self._data_quality_warning_keys: set[tuple[str, ...]] = set()
        self._template_indexes: dict[tuple[Any, ...], _TemplatePathIndex] = {}
        self._schema_snapshot_dir = Path(schema_snapshot_dir or self.runtime_paths.schema_snapshot_dir or DEFAULT_SCHEMA_SNAPSHOT_DIR).resolve()
        self._schema_salt: bytes | None = None
        self._schema_observations: dict[str, dict[str, Any]] = {}
        self._schema_events: list[dict[str, Any]] | None = None
        self._schema_snapshot_paths: dict[str, Path] = {}
        self.start_hashes: dict[str, str] = {}
        self.end_hashes: dict[str, str] = {}
        self._selected_paths: dict[str, Path] = {}
        self._layout_bindings: dict[str, SourceBinding] = {}
        self._contracts = {contract.source_id: contract for contract in configuration.source_contracts}

    def close(self) -> None:
        for book in self._books.values():
            book.close()
        self._books.clear()
        self._retrieval_cache.clear()
        self._template_indexes.clear()
        self._external_layouts.clear()

    def add_external_data_warning(self, *, source_path, sheet: str, location: str,
                                  message: str, institution_name: str = "",
                                  institution_code: str = "") -> None:
        warning = {
            "institution_name": institution_name,
            "institution_code": institution_code,
            "source_file": Path(source_path).name,
            "sheet": sheet,
            "location": location,
            "message": message,
        }
        key = tuple(warning.values())
        if key not in self._data_quality_warning_keys:
            self._data_quality_warning_keys.add(key)
            self.data_quality_warnings.append(warning)

    def hashes_unchanged(self) -> bool:
        unchanged = True
        for key, path in self._selected_paths.items():
            try:
                current = file_sha256(path)
            except OSError:
                current = "MISSING"
            self.end_hashes[key] = current
            if current != self.start_hashes.get(key):
                unchanged = False
        return unchanged

    def selected_bindings(self) -> list[tuple[str, SourceBinding]]:
        pairs: list[tuple[str, SourceBinding]] = []
        for role, source_map in (("current", self.context.current_sources),
                                 ("previous", self.context.previous_sources),
                                 ("external", self.context.external_sources)):
            pairs.extend((role, binding) for binding in source_map.values())
        pairs.extend(("layout", binding) for binding in self._layout_bindings.values())
        return pairs

    def snapshot_selected_hashes(self) -> None:
        for role, binding in self.selected_bindings():
            key = f"{role}:{binding.source_id}"
            path = Path(binding.path).resolve()
            self._selected_paths[key] = path
            try:
                self.start_hashes[key] = file_sha256(path)
            except OSError:
                self.start_hashes[key] = "MISSING"

    def _book(self, role: str, binding: SourceBinding) -> _Workbook:
        key = (role, binding.source_id)
        if key not in self._books:
            book = _Workbook(binding,
                             template_path=self.runtime_paths.template_path if role in {"current", "previous"} else None,
                             template_dir=self.runtime_paths.template_dir if role in {"current", "previous"} else None)
            self._books[key] = book
            self.start_hashes.setdefault(f"{role}:{binding.source_id}", book.sha256)
        return self._books[key]

    def _binding(self, role: str, rule: RuleDefinition, requested: str | None) -> SourceBinding:
        source_map = {
            "current": self.context.current_sources,
            "previous": self.context.previous_sources,
            "external": self.context.external_sources,
        }[role]
        selected = requested or (rule.default_source if role != "external" else "")
        if role == "external" and not selected:
            selected = self.context.default_external_source
        if selected:
            binding = source_map.get(selected)
            if binding is None:
                code = "SOURCE_NOT_SELECTED" if role == "previous" or role == "external" else "SOURCE_NOT_FOUND"
                raise RetrievalError(code, f"未选择或未登记逻辑数据源：{selected}")
            return binding
        candidates = [binding for binding in source_map.values() if binding.business == rule.business]
        if not candidates and len(source_map) == 1:
            candidates = list(source_map.values())
        if len(candidates) == 1:
            return candidates[0]
        if not candidates:
            code = "SOURCE_NOT_SELECTED" if role in {"previous", "external"} else "SOURCE_NOT_FOUND"
            raise RetrievalError(code, f"没有找到业务 {rule.business} 对应的{role}数据源")
        raise RetrievalError("SOURCE_AMBIGUOUS", f"业务 {rule.business} 对应多个{role}数据源，请显式选择逻辑来源")

    def _contract(self, binding: SourceBinding) -> SourceContract | None:
        return self._contracts.get(binding.source_id)

    def _identity_value(self, book: _Workbook, contract: SourceContract, label: str) -> str:
        if not contract.identity_sheet or not label:
            return ""
        if contract.identity_sheet not in book.book.sheetnames:
            raise RetrievalError("IDENTITY_NOT_FOUND", "身份契约指定的工作表不存在")
        sheet = book.book[contract.identity_sheet]
        matches: list[str] = []
        for row in range(1, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                value = book.merged_value(sheet.title, row, col)
                if _norm(value) == label:
                    right = book.merged_value(sheet.title, row, col + 1) if col < sheet.max_column else None
                    below = book.merged_value(sheet.title, row + 1, col) if row < sheet.max_row else None
                    candidate = _norm(right) or _norm(below)
                    if candidate:
                        matches.append(candidate)
        if len(set(matches)) > 1:
            raise RetrievalError("IDENTITY_AMBIGUOUS", "来源文件中的身份字段存在多项不同值")
        return matches[0] if matches else ""

    def _period_value(self, book: _Workbook, contract: SourceContract) -> str:
        if not contract.period_label:
            return ""
        return self._identity_value(book, contract, contract.period_label)

    def _validate_binding_context(self, role: str, binding: SourceBinding, book: _Workbook,
                                  contract: SourceContract | None) -> str:
        if contract is None:
            if not binding.period:
                raise RetrievalError("PERIOD_CONTRACT_MISSING", "来源未提供可验证的报告期")
            if role in {"current", "external"} and binding.period != self.context.period:
                raise RetrievalError("PERIOD_CONTRACT_MISMATCH", "本期来源期间与本次审核期间不一致")
            if role == "previous" and binding.period == self.context.period:
                raise RetrievalError("PERIOD_CONTRACT_MISMATCH", "上期来源期间与本期相同")
            if role in {"current", "previous"}:
                if not binding.institution_code:
                    raise RetrievalError("INSTITUTION_CONTRACT_MISSING", "本期/上期来源未提供可验证的机构身份")
                if self.context.institution_code and binding.institution_code != self.context.institution_code:
                    raise RetrievalError("INSTITUTION_MISMATCH", "来源文件机构身份与本次审核机构不一致")
            return binding.period
        identity = self._identity_value(book, contract, contract.institution_code_label)
        code = binding.institution_code or identity
        if binding.institution_code and identity and binding.institution_code != identity:
            raise RetrievalError("INSTITUTION_MISMATCH", "文件机构身份与运行选择的机构不一致")
        if code and self.context.institution_code and code != self.context.institution_code:
            raise RetrievalError("INSTITUTION_MISMATCH", "来源文件机构身份与本次审核机构不一致")
        if contract.institution_required and not code:
            raise RetrievalError("INSTITUTION_CONTRACT_MISSING", "来源契约要求机构身份，但文件或运行参数均未提供")
        if role in {"current", "previous"} and not code:
            raise RetrievalError("INSTITUTION_CONTRACT_MISSING", "本期/上期来源未提供可验证的机构身份")
        if contract.institution_name_label:
            name_from_file = self._identity_value(book, contract, contract.institution_name_label)
            if binding.institution_name and name_from_file and binding.institution_name != name_from_file:
                raise RetrievalError("INSTITUTION_MISMATCH", "文件机构名称与运行绑定机构名称不一致")
            if self.context.institution_name and name_from_file and self.context.institution_name != name_from_file:
                raise RetrievalError("INSTITUTION_MISMATCH", "来源文件机构名称与本次审核机构不一致")
        period = binding.period or self._period_value(book, contract)
        if binding.period and contract.period_label:
            extracted = self._period_value(book, contract)
            if extracted and extracted != binding.period:
                raise RetrievalError("PERIOD_CONTRACT_MISMATCH", "文件期间与运行绑定期间不一致")
        if contract.expected_period and period and contract.expected_period != period:
            raise RetrievalError("PERIOD_CONTRACT_MISMATCH", "来源期间与数据源契约要求不一致")
        if contract.period_required and not period:
            raise RetrievalError("PERIOD_CONTRACT_MISSING", "来源契约要求报告期，但文件未提供可验证期间")
        if not period:
            raise RetrievalError("PERIOD_CONTRACT_MISSING", "来源未提供可验证的报告期")
        if role == "current" and period and self.context.period and period != self.context.period:
            raise RetrievalError("PERIOD_CONTRACT_MISMATCH", "本期来源期间与本次审核期间不一致")
        if role == "previous" and period and self.context.period and period == self.context.period:
            raise RetrievalError("PERIOD_CONTRACT_MISMATCH", "上期来源期间与本期相同")
        if role == "external" and period and self.context.period and period != self.context.period:
            raise RetrievalError("PERIOD_CONTRACT_MISMATCH", "外部来源期间与本次审核期间不一致")
        return period

    def retrieve(self, function_name: str, args: tuple[Any, ...], keywords: dict[str, str],
                 rule: RuleDefinition) -> EvalValue:
        cache_key = (function_name, args, tuple(sorted(keywords.items())),
                     rule.business, rule.default_source, rule.default_sheet)
        cached = self._retrieval_cache.get(cache_key)
        if isinstance(cached, RetrievalError):
            raise cached
        if cached is not None:
            return cached
        role = "current" if function_name == "V_CUR" else "previous" if function_name == "V_PRE" else "external"
        try:
            if function_name in {"V_CUR", "V_PRE"}:
                value = self._retrieve_template(function_name, args, keywords, rule)
            elif function_name == "V_EXT":
                value = self._retrieve_external(args, keywords, rule)
            else:
                raise ConfigurationError(f"未知取值函数：{function_name}")
            self._retrieval_cache[cache_key] = value
            return value
        except RetrievalError as exc:
            detail = dict(exc.evidence)
            detail.setdefault("function", function_name)
            if function_name in {"V_CUR", "V_PRE"}:
                detail.setdefault("sheet", str(args[0]).strip() if len(args) == 3
                                  else keywords.get("sheet") or rule.default_sheet)
            try:
                binding = self._binding(role, rule, keywords.get("source"))
            except RetrievalError:
                binding = None
            if binding:
                detail.setdefault("source_id", binding.source_id)
                detail.setdefault("source_path", str(Path(binding.path).resolve()))
                detail.setdefault("source_sha256", self.start_hashes.get(f"{role}:{binding.source_id}", ""))
            enriched = RetrievalError(exc.code, exc.message, detail)
            self._retrieval_cache[cache_key] = enriched
            raise enriched from exc

    def _retrieve_template(self, function_name: str, args: tuple[Any, ...], keywords: dict[str, str],
                           rule: RuleDefinition) -> EvalValue:
        role = "current" if function_name == "V_CUR" else "previous"
        source_name = keywords.get("source")
        binding = self._binding(role, rule, source_name)
        book = self._book(role, binding)
        contract = self._contract(binding)
        if contract and contract.source_type in {"external", "外部"}:
            raise RetrievalError("SOURCE_CONTRACT_MISSING", "V_CUR/V_PRE 不能绑定外部来源契约")
        if role == "current" and contract and contract.source_type in {"previous", "上期"}:
            raise RetrievalError("SOURCE_CONTRACT_MISSING", "V_CUR 不能绑定仅供上期读取的数据源契约")
        source_period = self._validate_binding_context(role, binding, book, contract)
        business = binding.business or rule.business
        if len(args) == 3:
            positional_sheet = str(args[0]).strip()
            row_argument, column_argument = args[1], args[2]
        else:
            positional_sheet = ""
            row_argument, column_argument = args
        row_path = _as_path(row_argument, "行路径")
        col_path = _as_path(column_argument, "列路径")
        sheet_name = positional_sheet or keywords.get("sheet") or rule.default_sheet or (contract.default_sheet if contract else "")
        candidates = [metric for metric in self.configuration.metrics
                      if metric.business == business and metric.row_path == row_path
                      and metric.column_path == col_path and (not sheet_name or metric.sheet == sheet_name)]
        if not candidates and (not binding.business or binding.business == rule.business):
            # When the runtime business name differs from the source ID, permit the rule business
            # only if its complete metric key remains unique.
            candidates = [metric for metric in self.configuration.metrics
                          if metric.business == rule.business and metric.row_path == row_path
                          and metric.column_path == col_path and (not sheet_name or metric.sheet == sheet_name)]
        if not candidates:
            raise RetrievalError("CATALOG_METRIC_NOT_FOUND", "模板指标目录中不存在该业务、Sheet 和行列路径")
        if len(candidates) > 1:
            raise RetrievalError("METRIC_AMBIGUOUS", "模板指标路径匹配多条目录记录")
        metric = candidates[0]
        if not metric.confirmed:
            raise RetrievalError("METRIC_NOT_CONFIRMED", "目标指标口径尚未确认")
        if metric.metric_type != "text" and (not metric.source_unit or not metric.standard_unit):
            raise RetrievalError("UNIT_CONTRACT_MISSING", "目标指标缺少来源单位或标准单位")
        sheet_name = sheet_name or metric.sheet
        if sheet_name != metric.sheet:
            raise RetrievalError("METRIC_NOT_FOUND", "默认或显式工作表与指标目录不一致")
        if sheet_name not in book.book.sheetnames:
            raise RetrievalError("SHEET_NOT_FOUND", "源工作簿中不存在指标目录指定的工作表", {"sheet": sheet_name})
        value, address = self._locate_template_value(book, metric, role=role)
        return _value_with_unit(value, metric, binding, book, sheet_name, address,
                                source_period=source_period,
                                contract_version=contract.version if contract else "",
                                institution_code=self.context.institution_code)

    def _locate_template_value(self, book: _Workbook, metric: MetricSpec, *, role: str) -> tuple[Any, str]:
        index = self._template_index(book, metric, role=role)
        key = (metric.row_path, metric.column_path)
        matches = index.locations.get(key, [])
        if not matches:
            if not index.row_locations.get(metric.row_path):
                raise RetrievalError("ROW_PATH_NOT_FOUND", "工作表中未找到配置的完整行路径")
            if not index.column_locations.get(metric.column_path):
                raise RetrievalError("COLUMN_PATH_NOT_FOUND", "工作表中未找到配置的完整列路径")
            raise RetrievalError("STRUCTURE_ERROR", "完整行路径与列路径没有有效交点")
        if len(matches) > 1:
            sheet = book.book[metric.sheet]
            addresses = [sheet.cell(row, col).coordinate for row, col in matches]
            raise RetrievalError("METRIC_AMBIGUOUS", "完整行列路径匹配多个单元格", {"cells": addresses})
        row, col = matches[0]
        value, address, _, _ = book.value(metric.sheet, row, col)
        return value, address

    def _template_index(self, book: _Workbook, metric: MetricSpec, *, role: str) -> _TemplatePathIndex:
        data_regions = book.data_regions(metric.sheet)
        for path, digest in book.numeric_layout_dependencies.items():
            source_id = "LAYOUT_" + hashlib.sha256(str(path).encode("utf-8")).hexdigest()[:12]
            self._layout_bindings[source_id] = SourceBinding(source_id, path)
            self._selected_paths[f"layout:{source_id}"] = path
            self.start_hashes.setdefault(f"layout:{source_id}", digest)
        business = metric.business
        configured_metrics = [item for item in self.configuration.metrics
                              if item.business == business and item.sheet == metric.sheet
                              and item.row_header_titles == metric.row_header_titles]
        column_paths = tuple(sorted({item.column_path for item in configured_metrics}))
        key = (role, str(book.path), book.binding.source_id, business, metric.sheet,
               metric.row_header_titles, column_paths)
        cached = self._template_indexes.get(key)
        if cached is not None:
            return cached

        sheet = book.book[metric.sheet]
        titles = metric.row_header_titles
        if not titles:
            index = self._template_index_from_paths(book, metric, configured_metrics, role=role)
            self._template_indexes[key] = index
            self._record_schema_observations(book, index, role=role)
            return index
        regions = book.structure_regions(metric.sheet)
        row_regions: list[_NamedRegion] = []
        if regions:
            for region in regions:
                for row in range(region.min_row, region.max_row + 1):
                    for col in range(region.min_col, region.max_col - len(titles) + 2):
                        values = tuple(_norm(book.merged_value(metric.sheet, row, col + offset))
                                       for offset in range(len(titles)))
                        if values == titles:
                            row_regions.append(region)
                            break
                    if region in row_regions:
                        break
        else:
            row_regions = [_NamedRegion(
                "全工作表回退", metric.sheet, f"{metric.sheet}!A1:{get_column_letter(max(sheet.max_column, 1))}{max(sheet.max_row, 1)}",
                1, 1, max(sheet.max_row, 1), max(sheet.max_column, 1),
            )]

        if not row_regions:
            raise RetrievalError("STRUCTURE_ANCHOR_NOT_FOUND", "表结构命名区域中未找到配置的行表头路径标题")

        # A table's column-head area is the next numbered 表结构区域 to the right
        # of its row-head area. If a workbook only has one region, use the sheet's
        # remaining columns as a compatibility fallback.
        tables: list[dict[str, Any]] = []
        all_row_locations: dict[tuple[str, ...], list[int]] = defaultdict(list)
        all_column_locations: dict[tuple[str, ...], list[int]] = defaultdict(list)
        locations: dict[tuple[tuple[str, ...], tuple[str, ...]], list[tuple[int, int]]] = defaultdict(list)
        max_depth = max((len(path) for path in column_paths), default=1)
        for region in sorted(set(row_regions), key=lambda item: (_region_number(item.name), item.min_row, item.min_col)):
            anchors: list[tuple[int, int]] = []
            for row in range(region.min_row, region.max_row + 1):
                for col in range(region.min_col, region.max_col - len(titles) + 2):
                    values = tuple(_norm(book.merged_value(metric.sheet, row, col + offset))
                                   for offset in range(len(titles)))
                    if values == titles:
                        anchors.append((row, col))
            if not anchors:
                continue
            next_number = _region_number(region.name) + 1
            col_candidates = [item for item in regions if _region_number(item.name) == next_number
                              and item.min_col > region.max_col]
            column_region = min(col_candidates, key=lambda item: (item.min_col, item.min_row)) if col_candidates else None
            for anchor_index, (header_row, start_col) in enumerate(anchors, start=1):
                first_data_col = start_col + len(titles)
                first_col = max(first_data_col, column_region.min_col if column_region else first_data_col)
                last_col = min(sheet.max_column, column_region.max_col) if column_region else sheet.max_column
                if first_col > last_col:
                    continue
                rows_by_path: dict[tuple[str, ...], list[int]] = defaultdict(list)
                row_header_region = _NamedRegion(
                    region.name, metric.sheet,
                    f"{metric.sheet}!{get_column_letter(start_col)}{header_row + 1}:"
                    f"{get_column_letter(start_col + len(titles) - 1)}{region.max_row}",
                    header_row + 1, start_col, region.max_row, start_col + len(titles) - 1,
                )
                row_paths = row_path_candidates(
                    book, metric.sheet, row_header_region,
                    candidate_rows=range(header_row + 1, region.max_row + 1),
                )
                for row, path_info in row_paths.items():
                    for path in path_info["aliases"]:
                        rows_by_path[path].append(row)
                        all_row_locations[path].append(row)
                cols_by_path: dict[tuple[str, ...], list[int]] = defaultdict(list)
                for col in range(first_col, last_col + 1):
                    for depth in sorted({len(path) for path in column_paths}):
                        earliest = max(depth, header_row - depth + 1)
                        for end_row in range(earliest, header_row + 1):
                            path = book.header_path(metric.sheet, [(end_row - depth + 1 + offset, col)
                                                                   for offset in range(depth)])
                            if path in column_paths:
                                cols_by_path[path].append(col)
                                all_column_locations[path].append(col)
                for row_path in set(rows_by_path).intersection({item.row_path for item in configured_metrics}):
                    for col_path in set(cols_by_path).intersection(column_paths):
                        for row in rows_by_path[row_path]:
                            for col in sorted(set(cols_by_path[col_path])):
                                if not data_regions or in_numeric_regions(book, metric.sheet, row, col, data_regions):
                                    locations[(row_path, col_path)].append((row, col))

                expected_rows = sorted({item.row_path for item in configured_metrics})
                expected_columns = sorted(set(column_paths))
                row_counts = {"|".join(_path_digest(self._schema_secret(), path) for path in item):
                              min(2, len(rows_by_path.get(item, ()))) for item in expected_rows}
                column_counts = {"|".join(_path_digest(self._schema_secret(), path) for path in item):
                                 min(2, len(set(cols_by_path.get(item, ())))) for item in expected_columns}
                header_merges = [merged for merged in sheet.merged_cells.ranges
                                 if merged.min_col >= first_col and merged.max_col <= last_col
                                 and merged.min_row <= header_row
                                 and merged.max_row >= max(1, header_row - max_depth + 1)]
                column_merge_shapes = sorted(
                    f"R{item.min_row - header_row}C{item.min_col - first_col}:"
                    f"R{item.max_row - header_row}C{item.max_col - first_col}"
                    for item in header_merges
                )
                row_header_merges = [merged for merged in sheet.merged_cells.ranges
                                     if merged.min_col >= start_col
                                     and merged.max_col < start_col + len(titles)
                                     and merged.min_row <= header_row <= merged.max_row]
                row_merge_shapes = sorted(
                    f"R{item.min_row - header_row}C{item.min_col - start_col}:"
                    f"R{item.max_row - header_row}C{item.max_col - start_col}"
                    for item in row_header_merges
                )
                logical = {
                    "row_header_titles": _path_digest(self._schema_secret(), titles),
                    "row_header_width": len(titles),
                    "configured_row_presence": row_counts,
                    "configured_column_presence": column_counts,
                    "row_region_width": region.max_col - region.min_col + 1,
                    "column_region_width": last_col - first_col + 1,
                    "row_header_merge_shapes": row_merge_shapes,
                    "column_header_merge_shapes": column_merge_shapes,
                }
                positions = {
                    "row_header_anchor": sheet.cell(header_row, start_col).coordinate,
                    "row_region": region.reference,
                    "column_region": column_region.reference if column_region else "",
                    "configured_rows": {name: sorted(rows_by_path.get(path, ()))
                                        for name, path in zip(row_counts, expected_rows)},
                    "configured_columns": {name: sorted(set(cols_by_path.get(path, ())))
                                           for name, path in zip(column_counts, expected_columns)},
                    "header_merges": sorted(str(item) for item in header_merges + row_header_merges),
                }
                table = {
                    "business": business, "sheet": metric.sheet,
                    "table_id": f"{region.name}#{anchor_index}", "row_region": region.reference,
                    "column_region": column_region.reference if column_region else "",
                    "logical": logical, "positions": positions,
                }
                tables.append(table)

        for key_pair, coords in locations.items():
            locations[key_pair] = sorted(set(coords))
        index = _TemplatePathIndex(
            locations=dict(locations),
            row_locations={path: sorted(set(rows)) for path, rows in all_row_locations.items()},
            column_locations={path: sorted(set(cols)) for path, cols in all_column_locations.items()},
            tables=tables,
        )
        self._template_indexes[key] = index
        self._record_schema_observations(book, index, role=role)
        return index

    def _template_index_from_paths(self, book: _Workbook, metric: MetricSpec,
                                   configured_metrics: list[MetricSpec], *, role: str) -> _TemplatePathIndex:
        """Index rule paths directly when the workbook has no separate metric catalog."""
        sheet = book.book[metric.sheet]
        row_paths = {item.row_path for item in configured_metrics}
        key_pairs = {(item.row_path, item.column_path) for item in configured_metrics}
        locations, all_row_locations, all_column_locations, scanned_tables = scan_rule_paths(book, metric.sheet, key_pairs)
        tables: list[dict[str, Any]] = []
        for scanned in scanned_tables:
            region = scanned["region"]
            column_region = scanned["column_region"]
            search_region = scanned["search_region"]
            row_hits = scanned["row_hits"]
            columns_by_path = scanned["column_hits"]

            expected_rows = sorted(row_paths)
            expected_columns = sorted({column_path for _, column_path in key_pairs})
            row_tokens = {path: "|".join(_path_digest(self._schema_secret(), part) for part in path)
                          for path in expected_rows}
            column_tokens = {path: "|".join(_path_digest(self._schema_secret(), part) for part in path)
                             for path in expected_columns}
            row_merges = _merge_shapes(sheet, region)
            column_merges = _merge_shapes(sheet, search_region)
            logical = {
                "row_header_width": max((len(path) for path in expected_rows), default=0),
                "configured_row_presence": {
                    row_tokens[path]: min(2, len(set(row_hits.get(path, ())))) for path in expected_rows
                },
                "configured_column_presence": {
                    column_tokens[path]: min(2, len(columns_by_path.get(path, ()))) for path in expected_columns
                },
                "row_region_width": region.max_col - region.min_col + 1,
                "column_region_width": search_region.max_col - search_region.min_col + 1,
                "row_header_merge_shapes": row_merges,
                "column_header_merge_shapes": column_merges,
            }
            first_anchor = next((
                sheet.cell(row, last_col).coordinate
                for path in expected_rows for row, _, last_col in row_hits.get(path, ())
            ), "")
            positions = {
                "row_header_anchor": first_anchor,
                "row_region": region.reference,
                "column_region": search_region.reference if column_region else "",
                "configured_rows": {
                    row_tokens[path]: sorted({row for row, _, _ in row_hits.get(path, ())})
                    for path in expected_rows
                },
                "configured_columns": {
                    column_tokens[path]: sorted(columns_by_path.get(path, ()))
                    for path in expected_columns
                },
                "header_merges": sorted(set(row_merges + column_merges)),
            }
            tables.append({
                "business": metric.business, "sheet": metric.sheet,
                "table_id": f"{region.name}#1", "row_region": region.reference,
                "column_region": search_region.reference if column_region else "",
                "logical": logical, "positions": positions,
            })

        for key_pair, coords in locations.items():
            locations[key_pair] = sorted(set(coords))
        return _TemplatePathIndex(
            locations=dict(locations),
            row_locations={path: sorted(set(rows)) for path, rows in all_row_locations.items()},
            column_locations={path: sorted(set(cols)) for path, cols in all_column_locations.items()},
            tables=tables,
        )

    def _schema_secret(self) -> bytes:
        if self._schema_salt is None:
            salt_path = self._schema_snapshot_dir / ".schema-key"
            self._schema_snapshot_dir.mkdir(parents=True, exist_ok=True)
            try:
                with salt_path.open("xb") as stream:
                    stream.write(secrets.token_bytes(32))
            except FileExistsError:
                pass
            self._schema_salt = salt_path.read_bytes()
        return self._schema_salt

    def _record_schema_observations(self, book: _Workbook, index: _TemplatePathIndex, *, role: str) -> None:
        for table in index.tables:
            regions = book.data_regions(table["sheet"])
            table["logical"]["numeric_region_shapes"] = sorted(
                [region.max_row - region.min_row + 1, region.max_col - region.min_col + 1] for region in regions)
            table["positions"]["numeric_regions"] = [region.reference for region in regions]
            identity = {
                "role": role,
                "source_id": book.binding.source_id,
                "business": table["business"],
                "sheet": table["sheet"],
                "table_id": table["table_id"],
            }
            identity_text = json.dumps(identity, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
            snapshot_id = hashlib.sha256(identity_text.encode("utf-8")).hexdigest()
            self._schema_observations[snapshot_id] = {
                "snapshot_id": snapshot_id,
                "identity": identity,
                "logical": table["logical"],
                "positions": table["positions"],
                "summary": {
                    "row_header_width": table["logical"]["row_header_width"],
                    "configured_row_path_count": len(table["logical"]["configured_row_presence"]),
                    "configured_column_path_count": len(table["logical"]["configured_column_presence"]),
                    "matched_row_path_count": sum(value > 0 for value in table["logical"]["configured_row_presence"].values()),
                    "matched_column_path_count": sum(value > 0 for value in table["logical"]["configured_column_presence"].values()),
                    "row_region": table["row_region"], "column_region": table["column_region"],
                },
            }

    def compare_structure_snapshots(self) -> list[dict[str, Any]]:
        if self._schema_events is not None:
            return self._schema_events
        events: list[dict[str, Any]] = []
        for snapshot_id, observation in sorted(self._schema_observations.items()):
            path = self._snapshot_path(snapshot_id)
            self._schema_snapshot_paths[snapshot_id] = path
            event = {"snapshot_id": snapshot_id, **observation["identity"], **observation["summary"]}
            try:
                previous = json.loads(path.read_text(encoding="utf-8")) if path.exists() else None
                if previous is None:
                    event["status"] = "BASELINE_CREATED"
                elif previous.get("logical") != observation["logical"]:
                    event["status"] = "STRUCTURE_CHANGED"
                    event["detail"] = "配置路径或表头逻辑结构发生变化"
                elif previous.get("positions") != observation["positions"]:
                    event["status"] = "POSITION_CHANGED"
                    event["detail"] = "表结构逻辑未变，单元格位置已自动重定位"
                else:
                    event["status"] = "UNCHANGED"
                events.append(event)
            except Exception as exc:
                events.append({**event, "status": "SNAPSHOT_INVALID", "detail": f"历史结构快照无法读取：{type(exc).__name__}"})
        self._schema_events = events
        return events

    def persist_structure_snapshots(self) -> list[dict[str, Any]]:
        events = self.compare_structure_snapshots()
        try:
            self._schema_snapshot_dir.mkdir(parents=True, exist_ok=True)
            self._schema_secret()
        except Exception as exc:
            for event in events:
                if event.get("status") != "STRUCTURE_CHANGED":
                    event["status"] = "SNAPSHOT_WRITE_FAILED"
                    event["detail"] = f"结构快照未写入：{type(exc).__name__}"
            return events
        for snapshot_id, observation in self._schema_observations.items():
            event = next((item for item in events if item.get("snapshot_id") == snapshot_id), {})
            # A logical schema change needs review. Keep the accepted baseline
            # until the operator explicitly replaces it; position-only changes
            # are safe to reindex and accept automatically.
            if event.get("status") == "STRUCTURE_CHANGED":
                continue
            path = self._schema_snapshot_paths.get(snapshot_id) or self._snapshot_path(snapshot_id)
            payload = {
                "format_version": 1,
                "snapshot_id": snapshot_id,
                "identity": observation["identity"],
                "logical": observation["logical"],
                "positions": observation["positions"],
                "summary": observation["summary"],
            }
            temporary = path.with_name(f"{path.name}.{secrets.token_hex(6)}.tmp")
            try:
                temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
                temporary.replace(path)
            except Exception as exc:
                temporary.unlink(missing_ok=True)
                event["status"] = "SNAPSHOT_WRITE_FAILED"
                event["detail"] = f"结构快照未写入：{type(exc).__name__}"
        return events

    def _snapshot_path(self, snapshot_id: str) -> Path:
        return self._schema_snapshot_dir / f"{snapshot_id}.json"

    def _retrieve_external(self, args: tuple[Any, ...], keywords: dict[str, str],
                           rule: RuleDefinition) -> EvalValue:
        source_name = keywords.get("source")
        binding = self._binding("external", rule, source_name)
        book = self._book("external", binding)
        contract = self._contract(binding)
        metric_header = _norm(args[0])
        metric_matches = [metric for metric in self.configuration.external_metrics
                          if metric.source_id == binding.source_id and metric.metric_header == metric_header]
        if len(metric_matches) != 1:
            code = "METRIC_NOT_FOUND" if not metric_matches else "METRIC_AMBIGUOUS"
            raise RetrievalError(code, "外部指标目录未能唯一找到被引用的指标")
        metric = metric_matches[0]
        if not metric.confirmed:
            raise RetrievalError("METRIC_NOT_CONFIRMED", "外部指标口径尚未确认")
        if metric.metric_type == "amount" and (metric.source_unit not in _AMOUNT_FACTORS or metric.standard_unit not in _AMOUNT_FACTORS):
            raise RetrievalError("UNIT_CONTRACT_MISSING", "外部金额指标缺少有效的来源单位或标准单位")

        # S5 opts into layout detection so it can consume both the ordinary
        # report layout and the dated time-series layout. A saved external
        # source contract must not force S5 back onto the legacy fixed-row
        # reader, which cannot filter by report date or collect row warnings.
        if contract is None or self.context.allow_external_time_series:
            return self._retrieve_external_from_layout(book, binding, metric, metric_header, keywords.get("sheet", ""))
        if contract.mapping_sheet == "参照表":
            raise RetrievalError("MAPPING_SHEET_MISSING", "旧参照表已停用，请将机构映射统一维护在外部文件的机构参照表")
        if contract.mapping_sheet not in book.book.sheetnames:
            raise RetrievalError("MAPPING_SHEET_MISSING", "外部数据源缺少配置的机构映射工作表")
        source_period = self._validate_binding_context("external", binding, book, contract)

        mapping_sheet = book.book[contract.mapping_sheet]
        identity_col = _header_column(book, contract.mapping_sheet, contract.mapping_header_row, contract.mapping_identity_field)
        target_col = _header_column(book, contract.mapping_sheet, contract.mapping_header_row, contract.mapping_target_field)
        identity = self.context.institution_code
        if not identity:
            raise RetrievalError("INSTITUTION_CONTRACT_MISSING", "运行上下文没有当前机构代码，不能执行外部映射")
        map_rows: list[tuple[int, str]] = []
        for row in range(contract.mapping_data_start_row, mapping_sheet.max_row + 1):
            key = _norm(book.cell_raw(contract.mapping_sheet, row, identity_col))
            if key == identity:
                target = _norm(book.cell_raw(contract.mapping_sheet, row, target_col))
                map_rows.append((row, target))
        if not map_rows:
            raise RetrievalError("IDENTITY_NOT_MAPPED", "当前机构未在外部映射表中找到")
        if len(map_rows) != 1 or not map_rows[0][1]:
            raise RetrievalError("IDENTITY_MAPPING_AMBIGUOUS", "当前机构映射缺失或匹配多行，无法唯一确定报表项目")
        mapping_row, report_item = map_rows[0]

        data_sheet_name = keywords.get("sheet") or contract.default_sheet
        if not data_sheet_name or data_sheet_name not in book.book.sheetnames:
            raise RetrievalError("SHEET_NOT_FOUND", "外部数据源缺少配置指定的数据工作表")
        data_sheet = book.book[data_sheet_name]
        item_rows: list[int] = []
        for row in range(contract.data_start_row, data_sheet.max_row + 1):
            if _norm(book.cell_raw(data_sheet_name, row, contract.data_key_column)) == report_item:
                item_rows.append(row)
        if not item_rows:
            raise RetrievalError("REPORT_ITEM_NOT_FOUND", "外部数据表中不存在映射得到的报表项目")
        if len(item_rows) > 1:
            raise RetrievalError("REPORT_ITEM_DUPLICATE", "外部数据表中映射得到的报表项目重复")

        metric_cols: list[int] = []
        for col in range(contract.metric_first_column, data_sheet.max_column + 1):
            if _norm(book.cell_raw(data_sheet_name, contract.metric_header_row, col)) == metric_header:
                metric_cols.append(col)
        if not metric_cols:
            raise RetrievalError("METRIC_NOT_FOUND", "外部数据表中不存在配置指标标题")
        if len(metric_cols) > 1:
            raise RetrievalError("METRIC_AMBIGUOUS", "外部数据表的指标标题重复")
        data_row, data_col = item_rows[0], metric_cols[0]
        value, cell, _, _ = book.value(data_sheet_name, data_row, data_col)
        evidence = Evidence(
            source_id=binding.source_id, source_path=str(book.path), source_sha256=book.sha256,
            sheet=data_sheet_name, cell=cell, raw_value=value,
            raw_type=type(value).__name__, source_unit=metric.source_unit,
            standard_unit=metric.standard_unit,
            mapping_sheet=contract.mapping_sheet,
            mapping_cell=f"{get_column_letter(identity_col)}{mapping_row}:{get_column_letter(target_col)}{mapping_row}",
            mapping_key=identity, report_item=report_item, metric_header=metric_header,
            source_period=source_period, source_contract_version=contract.version,
            metric_definition_evidence=metric.evidence,
            institution_code=self.context.institution_code,
            precision=metric.precision,
        )
        return _convert_value(value, metric.data_type, metric.metric_type, metric.source_unit,
                              metric.standard_unit, metric.storage_scale, metric.allow_blank,
                              metric.precision, evidence)

    def _retrieve_external_from_layout(self, book: _Workbook, binding: SourceBinding,
                                       metric: ExternalMetricSpec, metric_header: str,
                                       sheet_filter: str = "") -> EvalValue:
        source_period = self._validate_binding_context("external", binding, book, None)
        layout = self._external_layouts.get(str(book.path))
        if layout is None:
            layout = discover_external_layout(
                book, support_time_series=self.context.allow_external_time_series,
            )
            self._external_layouts[str(book.path)] = layout
            for warning in layout.warnings:
                self.add_external_data_warning(
                    source_path=book.path, sheet=warning["sheet"], location=warning["location"],
                    message=warning["message"],
                )
        # The batch recognizer has already bound this file to exactly one
        # institution. Other identifiers in the workbook cannot change its row.
        entry, mapping_method = layout.institution_entry({self.context.institution_code}, [self.context.institution_name])
        mapping_row, report_item = entry.row, entry.report_item
        identity_column = entry.code_column if mapping_method == "统一社会信用代码" else entry.name_column
        mapping_key = entry.code if mapping_method == "统一社会信用代码" else entry.name
        column, data_row = layout.resolve(
            metric_header, report_item, sheet_filter,
            data_date=source_period if self.context.allow_external_time_series else "",
        )
        value, cell, _, _ = book.value(column.sheet, data_row, column.column)
        mapping_cell = (f"{get_column_letter(identity_column)}{mapping_row}:"
                        f"{get_column_letter(entry.target_column)}{mapping_row}")
        evidence = Evidence(
            source_id=binding.source_id, source_path=str(book.path), source_sha256=book.sha256,
            sheet=column.sheet, cell=cell, raw_value=value,
            raw_type=type(value).__name__, source_unit=metric.source_unit,
            standard_unit=metric.standard_unit, mapping_sheet=entry.sheet,
            mapping_cell=mapping_cell, mapping_key=mapping_key, report_item=report_item,
            metric_header=metric_header, source_period=source_period,
            institution_code=self.context.institution_code, precision=metric.precision,
            mapping_method=mapping_method,
        )
        if metric.data_type in {"number", "decimal", "integer", "numeric"}:
            external_number = _external_numeric_value(value)
            if external_number is None:
                if value is None or (isinstance(value, str) and not value.strip()):
                    issue = "指标单元格为空白"
                else:
                    issue = f"指标单元格内容“{_external_value_preview(value)}”不是数值"
                self.add_external_data_warning(
                    source_path=book.path, sheet=column.sheet, location=cell,
                    message=(f"报表项目“{report_item}”的指标“{metric_header}”{issue}，"
                             "已按 0 参与审核；请核查是否误删或改动了数据。"),
                    institution_name=self.context.institution_name,
                    institution_code=self.context.institution_code,
                )
                value = Decimal(0)
            else:
                value = external_number
        return _convert_value(value, metric.data_type, metric.metric_type, metric.source_unit,
                              metric.standard_unit, metric.storage_scale, metric.allow_blank,
                              metric.precision, evidence)


def _as_path(value: Any, label: str) -> tuple[str, ...]:
    parts = split_path(value)
    if not isinstance(value, (str, tuple, list)):
        raise RetrievalError("VALUE_TYPE_INVALID", f"{label}必须是字符串或字符串元组")
    if not parts or any(not part for part in parts):
        raise RetrievalError("VALUE_TYPE_INVALID", f"{label}不能为空")
    return parts


def _header_column(book: _Workbook, sheet_name: str, header_row: int, header: str) -> int:
    if not header:
        raise RetrievalError("SOURCE_CONTRACT_MISSING", "外部映射字段名未配置")
    sheet = book.book[sheet_name]
    matches = [col for col in range(1, sheet.max_column + 1)
               if _norm(book.cell_raw(sheet_name, header_row, col)) == header]
    if not matches:
        raise RetrievalError("MAPPING_FIELD_NOT_FOUND", "外部映射表缺少配置字段")
    if len(matches) > 1:
        raise RetrievalError("MAPPING_FIELD_AMBIGUOUS", "外部映射表配置字段重复")
    return matches[0]


def _value_with_unit(value: Any, metric: MetricSpec, binding: SourceBinding, book: _Workbook,
                     sheet: str, cell: str, *, source_period: str,
                     contract_version: str, institution_code: str) -> EvalValue:
    evidence = Evidence(
        source_id=binding.source_id, source_path=str(book.path), source_sha256=book.sha256,
        sheet=sheet, cell=cell, raw_value=value, raw_type=type(value).__name__,
        source_unit=metric.source_unit, standard_unit=metric.standard_unit,
        source_period=source_period, source_contract_version=contract_version,
        metric_layout_version=metric.layout_version,
        metric_definition_evidence=metric.evidence,
        institution_code=institution_code,
        precision=metric.precision,
    )
    return _convert_value(value, metric.data_type, metric.metric_type, metric.source_unit,
                          metric.standard_unit, metric.storage_scale, metric.allow_blank,
                          metric.precision, evidence)


def _external_numeric_value(value: Any) -> Decimal | None:
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        value = text
    try:
        number = value if isinstance(value, Decimal) else Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError):
        return None
    return number if number.is_finite() else None


def _external_value_preview(value: Any, *, limit: int = 80) -> str:
    text = str(value).replace("\r", " ").replace("\n", " ")
    return text if len(text) <= limit else text[:limit - 3] + "..."


def _convert_value(value: Any, data_type: str, metric_type: str, source_unit: str,
                   standard_unit: str, storage_scale: str, allow_blank: bool,
                   precision: int, evidence: Evidence) -> EvalValue:
    if value is None or (isinstance(value, str) and not value.strip()):
        if not allow_blank:
            raise RetrievalError("CELL_EMPTY", "目标单元格为空", {"sheet": evidence.sheet, "cell": evidence.cell})
        return EvalValue(None, "blank", standard_unit, metric_type, (evidence,))
    if data_type in {"text", "string"}:
        if not isinstance(value, str):
            raise RetrievalError("VALUE_TYPE_INVALID", "指标要求文本值，但源单元格不是文本", {"sheet": evidence.sheet, "cell": evidence.cell})
        return EvalValue(value, "string", None, "text", (evidence,), False)
    if data_type not in {"number", "decimal", "integer", "numeric"}:
        raise RetrievalError("VALUE_TYPE_INVALID", "指标目录中的数据类型不受支持")
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        raise RetrievalError("VALUE_TYPE_INVALID", "指标要求数值，但源单元格不是数值", {"sheet": evidence.sheet, "cell": evidence.cell})
    number = Decimal(str(value))
    if data_type == "integer" and number != number.to_integral_value():
        raise RetrievalError("VALUE_TYPE_INVALID", "指标要求整数，但源单元格包含小数", {"sheet": evidence.sheet, "cell": evidence.cell})
    conversion = ""
    if metric_type == "amount":
        if source_unit not in _AMOUNT_FACTORS or standard_unit not in _AMOUNT_FACTORS:
            raise RetrievalError("UNIT_CONTRACT_MISSING", "金额指标单位必须明确为元/万元/亿元")
        converted = number * _AMOUNT_FACTORS[source_unit] / _AMOUNT_FACTORS[standard_unit]
        conversion = f"{source_unit}→{standard_unit}:×{_AMOUNT_FACTORS[source_unit] / _AMOUNT_FACTORS[standard_unit]}"
        number = converted
    elif metric_type == "rate":
        if storage_scale == "fraction" and standard_unit in {"百分比", "%", "百分点"}:
            number *= 100
            conversion = "比例存储→百分比:×100"
        elif storage_scale == "percent" and standard_unit == "比例":
            number /= 100
            conversion = "百分比存储→比例:÷100"
        elif source_unit != standard_unit:
            if source_unit in {"比例", "fraction"} and standard_unit in {"百分比", "%", "百分点"}:
                number *= 100
                conversion = "比例→百分比:×100"
            elif source_unit in {"百分比", "%", "百分点"} and standard_unit == "比例":
                number /= 100
                conversion = "百分比→比例:÷100"
            elif source_unit in {"百分比", "%", "百分点"} and standard_unit in {"百分比", "%", "百分点"}:
                conversion = "比例单位别名规范化:数值不变"
            else:
                raise RetrievalError("UNIT_CONTRACT_MISSING", "比例指标来源单位与标准单位不兼容")
    elif metric_type == "count" and source_unit != standard_unit:
        raise RetrievalError("UNIT_CONTRACT_MISSING", "户数/笔数不得转换为其他单位")
    elif metric_type == "number" and source_unit != standard_unit:
        raise RetrievalError("UNIT_CONTRACT_MISSING", "普通数值指标不允许隐式换算")
    new_evidence = Evidence(**{**evidence.__dict__, "conversion": conversion})
    return EvalValue.number(number, unit=standard_unit, metric_type=metric_type,
                            evidence=(new_evidence,), precision=precision)
