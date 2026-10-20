"""Helpers for the optional S1-F02 unmatched-sheet summary template."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

from openpyxl.utils.cell import range_boundaries

from .models import CopyRange


FALLBACK_TEMPLATE_SHEET = "兜底模板"


@dataclass(frozen=True)
class SummaryFallbackTemplate:
    sheet_name: str
    data_range: CopyRange
    header_range: CopyRange


def split_summary_fallback_ranges(
    template_sheet_names: Iterable[str],
    data_ranges: Iterable[CopyRange],
    header_ranges: Iterable[CopyRange],
) -> tuple[list[CopyRange], list[CopyRange], SummaryFallbackTemplate | None]:
    """Separate the reserved fallback worksheet from ordinary source targets.

    When the reserved sheet is absent this is a no-op. When it exists, it must
    define exactly one data range and one header range on that worksheet; the
    ranges must share at least one column so they can describe one row table.
    """
    data_ranges = list(data_ranges)
    header_ranges = list(header_ranges)
    sheet_name = next(
        (name for name in template_sheet_names
         if str(name).casefold() == FALLBACK_TEMPLATE_SHEET.casefold()),
        None,
    )
    if sheet_name is None:
        return data_ranges, header_ranges, None

    fallback_data = [
        item for item in data_ranges
        if item.sheet_name.casefold() == sheet_name.casefold()
    ]
    fallback_headers = [
        item for item in header_ranges
        if item.sheet_name.casefold() == sheet_name.casefold()
    ]
    if len(fallback_data) != 1:
        raise ValueError(
            f"模板工作表“{sheet_name}”必须且只能定义一个“任意行汇总区域”"
        )
    if len(fallback_headers) != 1:
        raise ValueError(
            f"模板工作表“{sheet_name}”必须且只能定义一个“表头区域”"
        )

    data_bounds = range_boundaries(fallback_data[0].address)
    header_bounds = range_boundaries(fallback_headers[0].address)
    if max(data_bounds[0], header_bounds[0]) > min(data_bounds[2], header_bounds[2]):
        raise ValueError(
            f"模板工作表“{sheet_name}”的“表头区域”与“任意行汇总区域”没有共同列"
        )

    ordinary_data = [
        item for item in data_ranges
        if item.sheet_name.casefold() != sheet_name.casefold()
    ]
    ordinary_headers = [
        item for item in header_ranges
        if item.sheet_name.casefold() != sheet_name.casefold()
    ]
    return (
        ordinary_data,
        ordinary_headers,
        SummaryFallbackTemplate(sheet_name, fallback_data[0], fallback_headers[0]),
    )
