"""Conditional-format background extraction for the native edition.

实现已上移到 ``conditional_engine.OoxmlConditionalFormatEngine``（规则触发
检测引擎，priority/stopIfTrue 按 Excel 语义模拟、四态求值）；本模块只保留
原公开 API 作为兼容入口，native 管线节点与测试无需改动。

Design note
-----------
The Windows/Excel/WPS build reads a conditionally-applied background through
the COM ``DisplayFormat.Interior.Color`` API.  LibreOffice's UNO has **no**
equivalent: ``CellBackColor`` only reflects the cell's base fill, so a rule
that merely evaluates TRUE (e.g. ``cellIs > 0``) never surfaces through UNO.
Repeated checks on xlsx and .ods confirmed the rule object loads
(``ConditionalFormat.Count == 1``) but the rendered colour stays default.

On Linux this module therefore evaluates the submission's own conditional
format rules directly from the OOXML with openpyxl and reports the rule's
``dxf`` fill as the display colour.  The audit copy carries the source
submission's rules (reporting systems pre-flag suspect cells such as
``cellIs greaterThan 0``), and the already-calculated values are cached in
the file, so no separate formula render is required for the common ``cellIs``
form.  Rules that need a real formula renderer (``colorScale``, ``dataBar``,
``iconSet``, …) are reported as unsupported rather than guessed.

The public audit entry ``libreoffice_conditional_batch`` remains as a
compatibility context so the pipeline's step wiring and the step's
“失败后处理” policy are unchanged; it no longer starts a UNO listener because
none is needed for OOXML rule evaluation.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from ..conditional_engine import OoxmlConditionalFormatEngine
from ..models import CopyRange, Issue


def extract_conditional_format_issues(
    *, workbook_path: Path, ranges: list[CopyRange],
    structure_ranges: list[CopyRange], period: str, batch_id: str, source_file: Path,
) -> list[Issue]:
    """Return cells whose conditional-format rule triggers a background.

    Delegates to :class:`~base_audit.conditional_engine.OoxmlConditionalFormatEngine`
    (pure openpyxl OOXML rule evaluation; see that module for the priority /
    stopIfTrue semantics and the four-state evaluator).  Named ``条件格式区域``
    limits scanning; labels from ``表结构区域`` become the check indicator.
    """
    engine = OoxmlConditionalFormatEngine()
    issues, _extraction = engine.extract_issues(
        workbook_path, ranges, structure_ranges,
        period=period, batch_id=batch_id, source_file=source_file,
        audit_time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )
    return issues


def extract_conditional_format_issues_with_libreoffice(
    *, workbook_path: Path, ranges: list[CopyRange],
    structure_ranges: list[CopyRange], period: str, batch_id: str, source_file: Path,
) -> list[Issue]:
    """Compatibility entry: extraction is OOXML-native, LibreOffice not needed."""
    return extract_conditional_format_issues(
        workbook_path=workbook_path, ranges=ranges,
        structure_ranges=structure_ranges, period=period,
        batch_id=batch_id, source_file=source_file,
    )
