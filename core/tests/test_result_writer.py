from pathlib import Path

from openpyxl import Workbook, load_workbook

from base_audit.history import HISTORY_HEADERS
from base_audit.models import Issue
from base_audit.native.result_writer import write_audit_summary


def _issue(source: Path, audit: Path) -> Issue:
    return Issue(
        issue_id="甲银行｜数据表｜C6｜单位贷款", period="2026-09", batch_id="batch",
        audit_time="2026-09-11 12:00:00", triggered=True, status="新增",
        first_seen_period="2026-09", previous_seen_period="", consecutive_count=1,
        org_code="001", org_name="甲银行", report_code="单位贷款", sheet_name="数据表",
        rule_id="R001", severity="错误", formula_cell="C6", target_cell="C6",
        target_value=0, formula_result="错|单位贷款|不应该为0|0|0|0",
        message="不应该为0", source_file=str(source), audit_file=str(audit),
        check_field="单位贷款", comparison_value=0, difference_value=0,
        detail="不应该为0", institution_feedback="历史说明", auditor_opinion="复核",
    )


def test_openpyxl_result_writer_preserves_summary_contract(tmp_path):
    source = tmp_path / "甲银行.xlsx"
    audit = tmp_path / "甲银行_审核版.xlsx"
    for path in (source, audit):
        book = Workbook()
        book.active.title = "数据表"
        book.save(path)
        book.close()
    result = tmp_path / "结果.xlsx"

    write_audit_summary(result, [_issue(source, audit)])

    book = load_workbook(result)
    try:
        sheet = book["本期审核结果"]
        assert tuple(cell.value for cell in sheet[1]) == HISTORY_HEADERS
        assert tuple(cell.value for cell in sheet[2]) == (
            "甲银行.xlsx", "数据表", "C6", "错误", "单位贷款", "不应该为0",
            "0", "0", "0", "甲银行｜数据表｜C6｜单位贷款", "历史说明", "复核",
        )
        assert sheet["C2"].hyperlink.target == str(audit.resolve())
        assert sheet["C2"].hyperlink.location == "'数据表'!C6"
        assert sheet.freeze_panes == "A2"
    finally:
        book.close()
