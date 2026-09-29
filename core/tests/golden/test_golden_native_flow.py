"""黄金基线：DAG native（UOS）管线（mock LibreOffice 重算，任意平台可跑）。

冻结 DAG native 编排契约：结果工作簿命名与列序（描述在第 9 列，与 COM 版
不同）、条件格式 OOXML 求值、审核副本目录、历史库回写、源文件 SHA-256 不变。
真实 soffice 重算的缓存一致性按 docs/TEST_UOS.md / UOS_DAG验收记录 在实机验收。
"""

from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

import pytest

import golden_fixtures as gf

# 列序与 HISTORY_HEADERS（Windows 版）一致：描述在第 6 列，当前值第 7 列。
NATIVE_HEADERS = [
    "工作簿名", "工作表名", "定位单元格", "错误类型", "校验指标", "描述",
    "当前值", "对比值", "差值", "规则编号", "历史校验说明", "审核意见",
]


@pytest.fixture()
def native_env(tmp_path: Path):
    src = tmp_path / "源数据"
    src.mkdir()
    gf.build_source(src / "甲银行.xlsx", deposit=2_000_000, loan=3_000_000, reserve=30_000)
    gf.build_source(src / "乙银行.xlsx", deposit=5_000_000, loan=4_000_000, reserve=10_000,
                    comment=gf.CF_COMMENT)
    gf.build_bad_source(src / "丙银行.xlsx")
    template = tmp_path / "模板.xlsx"
    gf.build_audit_template(template)
    config = gf.make_config(tmp_path)
    hashes = {path.name: gf.sha256(path) for path in sorted(src.glob("*.xlsx"))}
    return tmp_path, src, template, config, hashes


def test_dag_native_flow_golden_contract(native_env):
    tmp_path, src, template, config, hashes = native_env
    from base_audit import engines
    from base_audit.engines.libreoffice import LibreOfficeCalculationResult
    from base_audit.service import AuditService
    from base_audit.workflow.runner import run_dag_flow

    def fake_recalculate(self, workbook_path):
        return LibreOfficeCalculationResult(
            workbook_path=workbook_path, engine_display="LibreOffice Calc (mock)", elapsed_seconds=0.0,
        )

    output_dir = tmp_path / "输出"
    service = AuditService(config_path=config)
    with patch.object(sys, "platform", "linux"), patch.object(
        engines.libreoffice.LibreOfficeCalculator, "recalculate", fake_recalculate
    ), patch.object(
        engines.libreoffice.LibreOfficeCalculator, "require_available", lambda self: None
    ):
        result = run_dag_flow(
            "dag:汇总核查表校验", service=service, template_path=template,
            input_dir=src, output_dir=output_dir, period="2026-06",
            history_path=config, recursive=False, write_flow_logs=False,
        )

    assert result.successful_files == 2
    assert result.failed_files == 1
    for name, before in hashes.items():
        assert gf.sha256(src / name) == before, name

    # ---- 结果工作簿：与 COM 版统一命名（输出名_时间戳），无期间段 ----
    assert result.summary_path is not None
    summary = gf.snapshot_workbook(result.summary_path)
    assert list(summary) == ["本期审核结果"]
    sheet = summary["本期审核结果"]
    # native 列序与 Windows 版不同：描述在第 9 列（平台差异在此冻结）
    assert sheet["headers"] == NATIVE_HEADERS
    rows = sheet["rows"]
    # mock 重算无公式缓存 → 公式规则不触发；条件格式走 OOXML 求值，乙 C4 触发。
    # 规则编号统一为“条件格式填充”（与 COM 版一致）；完整身份串在 issue_id。
    assert len(rows) == 1
    # check_field 走统一 IndicatorResolver（合并锚点+多级表头），含父级前缀；
    # 描述=规则兜底文案（该格无批注），当前值列与 Windows 列序一致。
    assert rows[0] == [
        "乙银行.xlsx", "20202资产负债表", "C4", "条件格式触发", "各项存款｜本期情况",
        "存款超过450万元，请核实", "5000000", None, None,
        "乙银行｜20202资产负债表｜C4｜各项存款｜本期情况", None, None,
    ]
    # 超链接指向审核副本目录（当前实现为绝对路径，Windows/UOS 一致）
    links = sheet["hyperlinks"]
    assert len(links) == 1 and links[0]["cell"] == "C2"
    assert links[0]["target"].replace("\\", "/").endswith("50_机构审核副本_<TS>/乙银行_审核版.xlsx")
    assert links[0]["location"] == "'20202资产负债表'!C4"
    # 审核副本目录默认保留
    copies_dir = result.copies["公式校验复制"]
    assert (copies_dir / "乙银行_审核版.xlsx").is_file()
    assert (copies_dir / "甲银行_审核版.xlsx").is_file()

    # ---- 历史库回写：命中项写回“核查表校验结果”表 ----
    from openpyxl import load_workbook

    book = load_workbook(config, read_only=True)
    try:
        history_rows = [
            list(row) for row in book["核查表校验结果"].iter_rows(values_only=True)
        ]
    finally:
        book.close()
    data_rows = [row for row in history_rows[1:] if any(v not in (None, "") for v in row)]
    assert len(data_rows) == 1
    assert data_rows[0][0] == "乙银行.xlsx" and data_rows[0][2] == "C4"
    # 历史库“规则编号”列存的是完整身份串（历史匹配键），与结果表的 rule_id 不同。
    assert data_rows[0][9] == "乙银行｜20202资产负债表｜C4｜各项存款｜本期情况"


def test_s1_business_outputs_are_identical_when_run_log_export_changes(native_env):
    """仅切换既有运行日志 Excel 输出时，S1 业务工作簿和结果数据保持一致。"""
    import re
    from openpyxl import load_workbook

    tmp_path, src, template, _config, hashes = native_env
    from base_audit import engines
    from base_audit.engines.libreoffice import LibreOfficeCalculationResult
    from base_audit.service import AuditService
    from base_audit.workflow.runner import run_dag_flow

    configs = [gf.make_config(tmp_path / "cfg_simple"), gf.make_config(tmp_path / "cfg_log")]
    output_dirs = [tmp_path / "out_simple", tmp_path / "out_log"]

    def fake_recalculate(self, workbook_path):
        return LibreOfficeCalculationResult(
            workbook_path=workbook_path, engine_display="LibreOffice Calc (mock)", elapsed_seconds=0.0,
        )

    results = []
    with patch.object(sys, "platform", "linux"), patch.object(
        engines.libreoffice.LibreOfficeCalculator, "recalculate", fake_recalculate
    ), patch.object(
        engines.libreoffice.LibreOfficeCalculator, "require_available", lambda self: None
    ):
        for config, output_dir, export_logs in zip(configs, output_dirs, (False, True)):
            results.append(run_dag_flow(
                "dag:汇总核查表校验", service=AuditService(config_path=config),
                template_path=template, input_dir=src, output_dir=output_dir,
                period="2026-06", history_path=config, recursive=False,
                write_flow_logs=export_logs,
            ))

    assert results[0].successful_files == results[1].successful_files == 2
    assert results[0].failed_files == results[1].failed_files == 1
    def business_issues(result):
        return [
            (issue.issue_id, issue.period, issue.triggered, issue.status,
             issue.first_seen_period, issue.previous_seen_period, issue.consecutive_count,
             issue.org_code, issue.org_name, issue.report_code, issue.sheet_name,
             issue.rule_id, issue.severity, issue.formula_cell, issue.target_cell,
             issue.target_value, issue.formula_result, issue.message, issue.source_file,
             issue.check_field, issue.comparison_value, issue.reference_value,
             issue.difference_value, issue.detail, issue.institution_feedback,
             issue.auditor_opinion)
            for issue in result.current_issues
        ]
    assert business_issues(results[0]) == business_issues(results[1])
    assert len(list(output_dirs[0].glob("*_运行日志_*.xlsx"))) == 0
    assert len(list(output_dirs[1].glob("*_运行日志_*.xlsx"))) == 1

    def business_snapshot(output_dir: Path):
        snapshots = {}
        for path in sorted(output_dir.rglob("*.xlsx")):
            if "_运行日志_" in path.name:
                continue
            key = re.sub(r"\d{8}_?\d{6}", "<TS>", path.relative_to(output_dir).as_posix())
            book = load_workbook(path, data_only=False)
            try:
                sheets = []
                for sheet in book.worksheets:
                    rows = []
                    for row in sheet.iter_rows():
                        rows.append(tuple(cell.value for cell in row))
                    links = tuple(sorted(
                        (cell.coordinate,
                         (cell.hyperlink.target or "").replace("\\", "/").rsplit("/", 1)[-1],
                         cell.hyperlink.location or "")
                        for row in sheet.iter_rows() for cell in row if cell.hyperlink
                    ))
                    sheets.append((sheet.title, tuple(rows), links))
                snapshots[key] = tuple(sheets)
            finally:
                book.close()
        return snapshots

    assert business_snapshot(output_dirs[0]) == business_snapshot(output_dirs[1])
    for name, before in hashes.items():
        assert gf.sha256(src / name) == before, name
