"""ConditionalFormatCompareReport：OOXML 与 Native 条件格式检测差异报告。

用途：同一个报送文件分别用两种模式提取，逐格比较触发结果——
这是确认「OOXML 模式可用」的人工验收依据（测试 A：Native；测试 B：OOXML）。

比较维度（两列都触发才算一致）：
  1. 触发异常数量
  2. 异常单元格位置（Sheet + Cell）
  3. 规则描述
  4/5. AuditResult 数量与核心字段（sheet/cell/severity/detail）

Native 路径需要 Windows + Excel/WPS；UOS 上不可用（无 Native 后端）。
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from .conditional_engine import (
    ConditionalFormatResult,
    NativeConditionalFormatEngine,
    OoxmlConditionalFormatEngine,
)


@dataclass
class ConditionalFormatDiffRow:
    """一行差异明细。"""

    sheet: str
    cell: str
    rule: str
    native_result: str      # "true" / "false" / "缺失"
    ooxml_result: str
    reason: str


@dataclass
class ConditionalFormatCompareReport:
    """一次 A/B 对比的完整报告。"""

    file: str
    mode_a: str = "NATIVE"
    mode_b: str = "OOXML"
    native_count: int = 0
    ooxml_count: int = 0
    match_count: int = 0
    diff_count: int = 0
    rows: list[ConditionalFormatDiffRow] = field(default_factory=list)
    native_unsupported: int = 0
    ooxml_unsupported: int = 0
    generated_at: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    @property
    def consistent(self) -> bool:
        return self.diff_count == 0

    def summary(self) -> str:
        verdict = "一致" if self.consistent else "存在差异"
        return (
            "条件格式 A/B 对比：{file}\n"
            "模式 A（Native）触发 {native} 格；模式 B（OOXML）触发 {ooxml} 格；"
            "一致 {match} 格；差异 {diff} 格 → {verdict}\n"
            "Native 暂不支持规则 {n_unsup} 条；OOXML 暂不支持规则 {o_unsup} 条"
        ).format(
            file=self.file, native=self.native_count, ooxml=self.ooxml_count,
            match=self.match_count, diff=self.diff_count, verdict=verdict,
            n_unsup=self.native_unsupported, o_unsup=self.ooxml_unsupported,
        )


def _result_text(result: ConditionalFormatResult | None) -> str:
    if result is None:
        return "false（缺失）"
    return "true" if result.triggered else "false"


def compare_workbook(*, workbook_path: Path, session, mapping,
                     structure_ranges, period: str, batch_id: str,
                     org_code: str = "", org_name: str = "",
                     audit_time: str = "") -> ConditionalFormatCompareReport:
    """对同一文件分别执行 Native（COM 渲染）与 OOXML（规则求值）并比较。

    ``mapping``：模板「条件格式结果提取」功能的 FeatureMapping；两条路径各自
    从同一个报送文件解析 ``条件格式区域`` 命名区域，保证比较口径同源。
    ``session``：已启动的 ExcelSession（Windows）。
    """
    from openpyxl import load_workbook as _load_workbook

    from .native.openpyxl_workbook import named_ranges

    workbook_path = Path(workbook_path)
    native_engine = NativeConditionalFormatEngine(session)
    ooxml_engine = OoxmlConditionalFormatEngine()

    # Native：需要 COM 工作簿对象（真实渲染）。
    native_issues, native_results, native_unsupported = ((), [], 0)
    workbook = session.open_workbook(workbook_path, read_only=True)
    try:
        native_issues, native_results, native_unsupported = native_engine.extract_issues(
            workbook, mapping=mapping, structure_ranges=structure_ranges,
            period=period, batch_id=batch_id,
            audit_time=audit_time, org_code=org_code, org_name=org_name,
            source_file=workbook_path, workbook_path=workbook_path,
        )
    finally:
        session.close_workbook(workbook)

    # OOXML：直接规则求值；命名区域与 Native 同源（同一报送文件）。
    book = _load_workbook(workbook_path, read_only=True, data_only=False, keep_links=False)
    try:
        mapping_ranges = named_ranges(book, (mapping,))
    finally:
        book.close()
    ooxml_issues, ooxml_extraction = ooxml_engine.extract_issues(
        workbook_path, mapping_ranges, structure_ranges,
        period=period, batch_id=batch_id, audit_time=audit_time,
        source_file=workbook_path,
    )

    def _key(result: ConditionalFormatResult):
        return (result.sheet, result.cell.upper())

    native_by_cell = {_key(r): r for r in native_results}
    ooxml_by_cell = {_key(r): r for r in ooxml_extraction.results}
    report = ConditionalFormatCompareReport(
        file=workbook_path.name,
        native_count=len(native_results), ooxml_count=len(ooxml_extraction.results),
        native_unsupported=native_unsupported,
        ooxml_unsupported=len(ooxml_extraction.unsupported),
    )
    for key in sorted(set(native_by_cell) | set(ooxml_by_cell)):
        native_hit = native_by_cell.get(key)
        ooxml_hit = ooxml_by_cell.get(key)
        if native_hit is not None and ooxml_hit is not None:
            if (native_hit.message or "") == (ooxml_hit.message or ""):
                report.match_count += 1
                continue
            report.rows.append(ConditionalFormatDiffRow(
                sheet=key[0], cell=key[1],
                rule=ooxml_hit.message or native_hit.message or "",
                native_result=_result_text(native_hit),
                ooxml_result=_result_text(ooxml_hit),
                reason="规则描述不一致",
            ))
            continue
        if native_hit is not None:
            report.rows.append(ConditionalFormatDiffRow(
                sheet=key[0], cell=key[1], rule=native_hit.message or "",
                native_result=_result_text(native_hit), ooxml_result="false（缺失）",
                reason="仅 Native 触发：OOXML 求值未命中或不支持该规则",
            ))
        else:
            report.rows.append(ConditionalFormatDiffRow(
                sheet=key[0], cell=key[1], rule=ooxml_hit.message or "",
                native_result="false（缺失）", ooxml_result=_result_text(ooxml_hit),
                reason="仅 OOXML 触发：COM 渲染未变色或规则在渲染时未成立",
            ))
    report.diff_count = len(report.rows)
    return report


def report_to_dict(report: ConditionalFormatCompareReport) -> dict:
    return {
        "file": report.file, "mode_a": report.mode_a, "mode_b": report.mode_b,
        "native_count": report.native_count, "ooxml_count": report.ooxml_count,
        "match_count": report.match_count, "diff_count": report.diff_count,
        "native_unsupported": report.native_unsupported,
        "ooxml_unsupported": report.ooxml_unsupported,
        "consistent": report.consistent,
        "generated_at": report.generated_at,
        "rows": [row.__dict__ for row in report.rows],
    }


def write_compare_report(report: ConditionalFormatCompareReport, output_path: Path) -> Path:
    """写出 Markdown 报告与同名 JSON（供人工验收与留档）。"""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "# 条件格式检测 A/B 对比报告", "",
        f"- 文件：{report.file}",
        f"- 模式 A：{report.mode_a}（Excel/WPS 真实渲染）",
        f"- 模式 B：{report.mode_b}（OOXML 规则求值）",
        f"- 生成时间：{report.generated_at}", "",
        "| 指标 | 数量 |", "| --- | --- |",
        f"| Native 触发 | {report.native_count} |",
        f"| OOXML 触发 | {report.ooxml_count} |",
        f"| 一致 | {report.match_count} |",
        f"| 差异 | {report.diff_count} |",
        f"| Native 暂不支持规则 | {report.native_unsupported} |",
        f"| OOXML 暂不支持规则 | {report.ooxml_unsupported} |", "",
    ]
    if report.rows:
        lines += ["## 差异明细", "",
                  "| Sheet | Cell | 规则 | Native | OOXML | 原因 |",
                  "| --- | --- | --- | --- | --- | --- |"]
        for row in report.rows:
            lines.append("| {sheet} | {cell} | {rule} | {native} | {ooxml} | {reason} |".format(
                sheet=row.sheet, cell=row.cell, rule=row.rule or "—",
                native=row.native_result, ooxml=row.ooxml_result, reason=row.reason,
            ))
    else:
        lines += ["## 差异明细", "", "无差异。"]
    output_path.write_text("\n".join(lines), encoding="utf-8")
    output_path.with_suffix(".json").write_text(
        json.dumps(report_to_dict(report), ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path
