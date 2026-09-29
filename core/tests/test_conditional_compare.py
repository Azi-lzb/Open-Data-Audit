"""ConditionalFormatCompareReport 测试：A/B 统计、差异明细与报告落盘。

Native 侧用桩会话模拟（真实 Excel/WPS 的 A/B 验收见文件末尾的 COM 用例）。
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "core" / "src"))

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

from base_audit.conditional_compare import (
    compare_workbook,
    report_to_dict,
    write_compare_report,
)
from base_audit.conditional_engine import ConditionalFormatResult

RED_FILL = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")


def _write_workbook(path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "数据"
    sheet["B2"] = 10
    sheet["B3"] = 0
    sheet.conditional_formatting.add(
        "B2:B3", CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL))
    book.save(path)
    book.close()


class _StubSession:
    """Native 侧桩：模拟 ExcelSession 的接口与可控的触发结果。"""

    engine_name = "Microsoft Excel"

    def __init__(self, results, unsupported=0):
        # 引擎从 session.last_conditional_results 读取同构结果旁路。
        self.last_conditional_results = list(results)
        self.last_conditional_unsupported_count = unsupported
        self.opened = 0

    def open_workbook(self, path, read_only=True):
        self.opened += 1
        return object()

    def close_workbook(self, workbook):
        pass

    def extract_conditional_format_issues(self, workbook, **kwargs):
        return []


def _result(sheet, cell, message, triggered=True):
    return ConditionalFormatResult(
        sheet=sheet, cell=cell, rule_type="cellis", formula="0",
        triggered=triggered, message=message,
    )


def _cf_mapping():
    from base_audit.name_config import FeatureMapping

    return FeatureMapping(
        name="条件格式结果提取", feature_type="conditional_format_issue_extract",
        range_names=("条件格式区域",), workbook_limited=False,
        workbook_keyword="", remark="",
    )


class TestCompareReport:
    def test_identical_results_report_consistent(self, tmp_path):
        path = tmp_path / "机构A.xlsx"
        session = _StubSession([_result("数据", "B2", "条件格式规则：B2>0")])
        from openpyxl.workbook.defined_name import DefinedName
        book = Workbook()
        sheet = book.active
        sheet.title = "数据"
        sheet["B2"] = 10
        book.defined_names.add(DefinedName("条件格式区域", attr_text="数据!$B$2:$B$3"))
        sheet.conditional_formatting.add(
            "B2:B3", CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL))
        book.save(path)
        book.close()
        report = compare_workbook(
            workbook_path=path, session=session, mapping=_cf_mapping(),
            structure_ranges=[], period="2026-08", batch_id="B1",
        )
        assert report.native_count == 1
        assert report.ooxml_count == 1
        assert report.match_count == 1
        assert report.diff_count == 0
        assert report.consistent is True

    def test_mismatch_produces_diff_rows(self, tmp_path):
        path = tmp_path / "机构B.xlsx"
        # Native 触发 B2/B3，OOXML 只触发 B2：应产生一行「仅 Native 触发」差异。
        session = _StubSession([
            _result("数据", "B2", "条件格式规则：B2>0"),
            _result("数据", "B3", "条件格式规则：B3>0"),
        ])
        from openpyxl.workbook.defined_name import DefinedName
        book = Workbook()
        sheet = book.active
        sheet.title = "数据"
        sheet["B2"] = 10
        sheet["B3"] = 0
        book.defined_names.add(DefinedName("条件格式区域", attr_text="数据!$B$2:$B$3"))
        sheet.conditional_formatting.add(
            "B2:B3", CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL))
        book.save(path)
        book.close()
        report = compare_workbook(
            workbook_path=path, session=session, mapping=_cf_mapping(),
            structure_ranges=[], period="2026-08", batch_id="B1",
        )
        assert report.ooxml_count == 1
        assert report.diff_count == 1
        assert report.consistent is False
        row = report.rows[0]
        assert (row.sheet, row.cell) == ("数据", "B3")
        assert row.native_result == "true"
        assert "仅 Native" in row.reason

    def test_write_report_outputs_markdown_and_json(self, tmp_path):
        path = tmp_path / "机构C.xlsx"
        session = _StubSession([], unsupported=2)
        from openpyxl.workbook.defined_name import DefinedName
        book = Workbook()
        sheet = book.active
        sheet.title = "数据"
        sheet["B2"] = 10
        book.defined_names.add(DefinedName("条件格式区域", attr_text="数据!$B$2:$B$3"))
        sheet.conditional_formatting.add(
            "B2:B3", CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL))
        book.save(path)
        book.close()
        report = compare_workbook(
            workbook_path=path, session=session, mapping=_cf_mapping(),
            structure_ranges=[], period="2026-08", batch_id="B1",
        )
        out = tmp_path / "报告" / "条件格式对比.md"
        written = write_compare_report(report, out)
        assert written.exists()
        text = written.read_text(encoding="utf-8")
        assert "条件格式检测 A/B 对比报告" in text
        payload = json.loads(out.with_suffix(".json").read_text(encoding="utf-8"))
        assert payload["consistent"] == report.consistent
        assert payload["native_unsupported"] == 2
