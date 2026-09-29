"""统信 UOS 原生管线（native 包）在 Windows 上可跑的 mock 测试。

覆盖：模板读取/公式复制（openpyxl_workbook）、条件格式 OOXML 求值
（conditional_scan，复用 unified 求值器）、soffice 重算引擎（mock
subprocess）、逐文件编排（audit_flow，mock LibreOfficeCalculator）、
汇总三列（summary）、历史读写（history_io）与平台分派（engines/service）。
"""

from __future__ import annotations

import shutil
import sys
import pytest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from urllib.parse import unquote, urlparse
from urllib.request import url2pathname

import pytest
from openpyxl import Workbook, load_workbook

from base_audit import engines, settings as settings_mod
from base_audit.engines.libreoffice import LibreOfficeCalculator
from base_audit.engines.libreoffice import (
    find_calc_engine,
)  # noqa: F401  (re-exported module surface)
from base_audit.native import summary as native_summary
from base_audit.native.conditional_scan import extract_conditional_format_issues
from base_audit.native.history_io import read_history_xlsx, write_history_xlsx
from base_audit.native.openpyxl_workbook import add_external_sheets, copy_formula_ranges, read_template


# ---------------------------------------------------------------------------
# LibreOffice 引擎（engines/libreoffice.py，移植自 flet/uos）
# ---------------------------------------------------------------------------

def test_direct_engine_respects_configured_executable(monkeypatch) -> None:
    monkeypatch.setenv("BASE_AUDIT_ENGINE", "direct")
    monkeypatch.setenv("BASE_AUDIT_SOFFICE", sys.executable)
    engine = engines.libreoffice.find_calc_engine()
    assert engine is not None
    assert engine.source == "环境变量 BASE_AUDIT_SOFFICE"


def test_linglong_engine_uses_verified_command_prefix(monkeypatch) -> None:
    monkeypatch.setenv("BASE_AUDIT_ENGINE", "ll-cli")
    monkeypatch.setattr(
        engines.libreoffice.shutil, "which",
        lambda name: "/usr/bin/ll-cli" if name == "ll-cli" else None,
    )
    monkeypatch.setattr(
        engines.libreoffice.subprocess, "run",
        lambda *args, **kwargs: SimpleNamespace(stdout="org.libreoffice.libreoffice 25.8", stderr="", returncode=0),
    )
    engine = engines.libreoffice.find_calc_engine()
    assert engine is not None
    assert engine.command_prefix == ("/usr/bin/ll-cli", "run", "org.libreoffice.libreoffice", "--", "soffice")


def test_recalculate_stages_and_replaces_only_after_success(tmp_path: Path, monkeypatch) -> None:
    workbook_path = tmp_path / "审核版.xlsx"
    book = Workbook()
    book.active["A1"] = "=SUM(2,3)"
    book.save(workbook_path)
    book.close()
    commands: list[list[str]] = []

    def fake_run(command, **kwargs):
        commands.append(command)
        output_dir = Path(command[command.index("--outdir") + 1])
        staged_arg = command[-1]
        staged = (
            Path(url2pathname(unquote(urlparse(staged_arg).path)))
            if staged_arg.startswith("file:") else Path(staged_arg)
        )
        shutil.copy2(staged, output_dir / staged.name)
        return SimpleNamespace(stdout="", stderr="", returncode=0)

    monkeypatch.setattr(engines.libreoffice.subprocess, "run", fake_run)
    result = LibreOfficeCalculator(sys.executable if False else Path(sys.executable)).recalculate(workbook_path)

    assert result.workbook_path == workbook_path
    assert "-env:UserInstallation=" in " ".join(commands[0])
    restored = load_workbook(workbook_path, data_only=False)
    assert restored.active["A1"].value == "=SUM(2,3)"
    restored.close()


# ---------------------------------------------------------------------------
# 模板读取与公式复制（native/openpyxl_workbook.py）
# ---------------------------------------------------------------------------

def _make_template(path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "报表"
    sheet["D1"] = "指标"
    sheet["D2"] = "=IF(C2>0,\"错误|C2|0|余额应为零\",\"\")"
    sheet["C2"] = 5
    book.create_sheet("审核规则").append(("规则编号", "启用", "报表代码", "工作表", "公式单元格", "定位单元格", "级别", "问题说明"))
    rules = book["审核规则"]
    rules.append(("R001", "是", "", "报表", "D2", "D2", "错误", "余额核验"))
    book.save(path)
    book.close()


def test_read_template_structured_rules(tmp_path: Path) -> None:
    template = tmp_path / "模板.xlsx"
    _make_template(template)
    definition = read_template(template)
    assert definition.structured
    assert [rule.rule_id for rule in definition.rules] == ["R001"]
    assert definition.rules[0].enabled


def test_copy_formula_ranges_writes_full_calc_flag(tmp_path: Path) -> None:
    template = tmp_path / "模板.xlsx"
    _make_template(template)
    audit = tmp_path / "机构A_审核版.xlsx"
    book = Workbook()
    book.active.title = "报表"
    book.active["C2"] = 5
    book.save(audit)
    book.close()
    definition = read_template(template)
    copy_formula_ranges(template, audit, definition)
    restored = load_workbook(audit, data_only=False)
    assert str(restored["报表"]["D2"].value).startswith("=IF")
    assert restored.calculation.fullCalcOnLoad
    restored.close()


def test_openpyxl_external_copy_then_formula_copy_keeps_local_sheet_reference(tmp_path: Path) -> None:
    """外部表先写入审核副本后，公式文本必须仍指向副本内的同名工作表。"""
    template = tmp_path / "模板.xlsx"
    _make_template(template)
    book = load_workbook(template)
    book["报表"]["D2"] = '=IF(参照表!$B$2>0,"错误|余额|不应为0|0|0|0","")'
    book.save(template)
    book.close()
    audit = tmp_path / "机构A_审核版.xlsx"
    book = Workbook()
    book.active.title = "报表"
    book.save(audit)
    book.close()
    external = tmp_path / "外部.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "参照表"
    sheet["B2"] = 1
    book.save(external)
    book.close()

    definition = read_template(template)
    add_external_sheets(audit, external, ("参照表",))
    copy_formula_ranges(template, audit, definition)

    restored = load_workbook(audit, data_only=False)
    try:
        assert restored["报表"]["D2"].value == '=IF(参照表!$B$2>0,"错误|余额|不应为0|0|0|0","")'
        assert "参照表" in restored.sheetnames
    finally:
        restored.close()


# ---------------------------------------------------------------------------
# 条件格式 OOXML 求值（native/conditional_scan.py）
# ---------------------------------------------------------------------------

def _cellis_workbook(path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "数据"
    sheet["B2"] = 10
    sheet["B3"] = 0
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill
    red = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
    sheet.conditional_formatting.add("B2:B3", CellIsRule(operator="greaterThan", formula=["0"], fill=red))
    book.save(path)
    book.close()


def test_extracts_cells_that_trigger_cellis(tmp_path: Path) -> None:
    workbook = tmp_path / "机构A.xlsx"
    _cellis_workbook(workbook)
    from base_audit.models import CopyRange
    issues = extract_conditional_format_issues(
        workbook_path=workbook, ranges=[CopyRange("数据", "B2:B3")],
        structure_ranges=[], period="2026-08", batch_id="B1", source_file=workbook,
    )
    assert [item.target_cell for item in issues] == ["B2"]


def test_expression_relative_translation_triggers(tmp_path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "数据"
    sheet["C2"] = 5
    sheet["D2"] = 3
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import PatternFill
    red = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
    sheet.conditional_formatting.add("C2:C2", FormulaRule(formula=["AND($D$2>0,$C2>$D$2)"], fill=red))
    path = tmp_path / "机构B.xlsx"
    book.save(path)
    book.close()
    from base_audit.models import CopyRange
    issues = extract_conditional_format_issues(
        workbook_path=path, ranges=[CopyRange("数据", "C2:C2")],
        structure_ranges=[], period="2026-08", batch_id="B1", source_file=path,
    )
    assert len(issues) == 1


# ---------------------------------------------------------------------------
# 平台分派（engines + settings + service 门面）
# ---------------------------------------------------------------------------

def test_platform_dispatch_windows_com() -> None:
    if sys.platform == "win32":
        assert engines.pipeline_kind("自动") == "com"
    else:
        assert engines.pipeline_kind("自动") == "native"


def test_platform_dispatch_simulated_linux() -> None:
    """UOS 引擎清单：LibreOffice 可用；Excel/WPS 展示但禁用（仅 Windows 支持）。"""
    with patch.object(sys, "platform", "linux"):
        assert engines.pipeline_kind("自动") == "native"
        items = engines.available_engines()
        assert [item["value"] for item in items if not item.get("disabled")] == [
            "自动", "LibreOffice Calc"]
        disabled = {item["value"] for item in items if item.get("disabled")}
        assert disabled == {"Microsoft Excel", "WPS 表格"}
        assert engines.valid_engine_values() == {"自动", "LibreOffice Calc"}


def test_settings_accept_libreoffice_engine() -> None:
    stored = settings_mod.UserSettings(calculation_engine="LibreOffice Calc")
    assert stored.calculation_engine == "LibreOffice Calc"


def test_excel_session_rejects_non_windows() -> None:
    from base_audit.excel_com import ExcelSession, ExcelUnavailableError
    with patch.object(sys, "platform", "linux"):
        with pytest.raises(ExcelUnavailableError):
            ExcelSession("自动").__enter__()


# ---------------------------------------------------------------------------
# 汇总三列（native/summary.py）
# ---------------------------------------------------------------------------

def test_region_summary_appends_history_columns_only_for_identity_sheets(tmp_path: Path) -> None:
    from base_audit.history import (
        HISTORY_AUDIT_SHEET, LOCAL_VALIDATION_HISTORY_FIELDS, LOCAL_VALIDATION_HISTORY_SHEET,
    )
    from base_audit.name_config import FeatureMapping, USED_RANGE_SUMMARY_FUNCTION
    from openpyxl.workbook.defined_name import DefinedName

    template = tmp_path / "汇总模板.xlsx"
    book = Workbook()
    local = book.active
    local.title = "本地校验结果"
    for offset, name in enumerate(["批次", "表单名称", "规则编号", "规则类型", "规则描述", "校验字段"], start=1):
        local.cell(1, offset, name)
    book.save(template)
    book.close()

    source = tmp_path / "机构A_2026-08.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "本地校验结果"
    sheet.append(("批次", "表单名称", "规则编号", "规则类型", "规则描述", "校验字段"))
    sheet.append(("", "", "R1", "", "余额核验", "余额"))
    book.save(source)
    book.close()

    # 历史配置：一条 R1 的历史记录 → 三列应出现 1/说明/待整改
    config = tmp_path / "历史审核配置.xlsx"
    book = Workbook()
    book.remove(book.active)
    lv = book.create_sheet(LOCAL_VALIDATION_HISTORY_SHEET)
    lv.append(["来源文件", "来源工作表", "批次", "表单名称", "规则编号", "规则类型", "规则描述", "校验字段", "当前值", "对比值", "差值", *LOCAL_VALIDATION_HISTORY_FIELDS])
    lv.append(["机构A", "本地校验结果", "", "", "R1", "", "余额核验", "余额", "", "", "", 1, "利率超限", "待整改"])
    book.create_sheet(HISTORY_AUDIT_SHEET).append(("占位",))
    book.save(config)
    book.close()

    # 命名区域：汇总区域=数据区（含表头），表头区域=表头行
    book = load_workbook(template)
    book.defined_names.add(DefinedName("汇总区域", attr_text="'本地校验结果'!$A$2:$F$2"))
    book.defined_names.add(DefinedName("表头区域", attr_text="'本地校验结果'!$A$1:$F$1"))
    book.save(template)
    book.close()

    feature = FeatureMapping("本地校验汇总", USED_RANGE_SUMMARY_FUNCTION, ("汇总区域",), False, "", "")
    path = native_summary.run_region_summaries(
        template_path=template, input_dir=tmp_path, output_dir=tmp_path / "汇总输出",
        features=[feature], recursive=False, output_name="汇总测试",
        history_config_path=config,
    )
    out = load_workbook(path)
    local_sheet = out["本地校验结果"]
    headers = [cell.value for cell in local_sheet[1]]
    assert headers[-3:] == list(LOCAL_VALIDATION_HISTORY_FIELDS)
    last_row = [cell.value for cell in local_sheet[2]]
    assert last_row[0] == "机构A"  # 来源文件去日期
    assert last_row[4] == "R1"    # 0来源文件 1来源工作表 2批次 3表单名称 4规则编号
    assert last_row[-3:] == ["1", "利率超限", "待整改"]  # join 后为字符串
    out.close()


# ---------------------------------------------------------------------------
# service 入口分派（preflight / summarize_regions / run / merge_template_files）
# ---------------------------------------------------------------------------


def test_native_enrich_history_columns_appends_by_rule(tmp_path: Path) -> None:
    """历史说明富化节点在 native 侧：对汇总表按内置规则补人工列。"""
    from base_audit.native.summary import enrich_history_columns
    from base_audit.name_config import initialize_config

    config = initialize_config(tmp_path / "逐笔统计系统_配置.xlsx")
    book = load_workbook(config)
    # initialize_config 已按模板表头预建“业务说明”；此处填入一行历史。
    sheet = book["业务说明"]
    headers = [cell.value for cell in sheet[1]]
    positions = {name: index for index, name in enumerate(headers)}
    for name, value in {
        "来源文件": "甲银行", "来源工作表": "业务说明",
        "业务表单名称": "表A", "业务表单英文名称": "TAB_A", "历史说明": "往期人工说明",
    }.items():
        sheet.cell(2, positions[name] + 1, value)
    book.save(config)
    book.close()

    summary = tmp_path / "汇总.xlsx"
    book = Workbook()
    book.active.title = "业务说明"
    book["业务说明"].append(["来源文件", "来源工作表", "业务表单名称", "业务表单英文名称", "本期情况"])
    book["业务说明"].append(["甲银行", "业务说明", "表A", "TAB_A", 123])
    book.save(summary)
    book.close()

    enrich_history_columns(summary, config)
    out = load_workbook(summary, read_only=True)
    rows = list(out["业务说明"].iter_rows(values_only=True))
    out.close()
    assert rows[0][-1] == "历史说明"
    assert rows[1][-1] == "往期人工说明"


def _service(tmp_path: Path):
    from base_audit.name_config import initialize_config
    from base_audit.service import AuditService
    history = tmp_path / "历史审核配置.xlsx"
    initialize_config(history)
    return AuditService(config_path=history)


def test_preflight_dispatches_to_native_on_linux(tmp_path: Path) -> None:
    from unittest.mock import patch
    from base_audit import engines
    template = tmp_path / "模板.xlsx"
    _make_template(template)
    source = tmp_path / "源" / "机构A.xlsx"
    source.parent.mkdir()
    book = Workbook(); book.active.title = "报表"; book.active["C2"] = 5
    book.save(source); book.close()

    def fake_find(*args, **kwargs):
        return None

    with patch.object(sys, "platform", "linux"), patch.object(
        engines.libreoffice, "find_calc_engine", fake_find
    ):
        service = _service(tmp_path)
        result = service.preflight(
            template_path=template, input_dir=source.parent, output_dir=tmp_path / "输出",
            recursive=False, write_report=False,
        )
    assert result.total_files == 1 and result.matched_files == 1
    assert "审核前检查完成" in result.summary_text()


def test_merge_template_files_runs_with_openpyxl_on_linux(tmp_path: Path) -> None:
    """native 下用 openpyxl 制作联合模板：基准权威、同名跳过、名称加表后缀。"""
    from unittest.mock import patch

    from openpyxl.workbook.defined_name import DefinedName

    base = tmp_path / "！10.基础数据-单位贷款202609.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "单位贷款信息"
    sheet["A1"], sheet["B1"] = "指标", "数值"
    sheet["A2"], sheet["B2"] = "余额", 1
    book.defined_names.add(DefinedName("校验区域_001", attr_text="'单位贷款信息'!$A$2:$B$2"))
    book.create_sheet("集中系统数据")["A1"] = "机构"
    book.save(base)
    book.close()

    source = tmp_path / "！70.债券业务202608.xlsx"
    book = Workbook()
    other = book.active
    other.title = "存量债券投资"
    other["A1"] = "债券"
    book.defined_names.add(DefinedName("校验区域_001", attr_text="'存量债券投资'!$A$1:$A$1"))
    book.create_sheet("单位贷款信息")["A1"] = "同名表不应覆盖基准"
    book.save(source)
    book.close()

    with patch.object(sys, "platform", "linux"):
        service = _service(tmp_path)
        result = service.merge_template_files(base_template=base, source_templates=[source])

    assert result.output_path.is_file()
    merged = load_workbook(result.output_path)
    try:
        # 基准权威：依赖表保留，同名表未被覆盖
        assert "集中系统数据" in merged.sheetnames
        assert merged["单位贷款信息"]["A1"].value == "指标"
        # 来源新表已并入，命名区域带工作表后缀且原名清除
        assert "存量债券投资" in merged.sheetnames
        names = {str(key) for key in merged.defined_names}
        assert "校验区域_001_单位贷款信息" in names
        assert "校验区域_001_存量债券投资" in names
        assert "校验区域_001" not in names
    finally:
        merged.close()


def test_dag_combine_sheets_runs_with_openpyxl_on_linux(tmp_path: Path) -> None:
    """组合工作表不需要计算引擎；native DAG 不得导入或回退 COM。"""
    from base_audit.workflow.runner import run_dag_flow

    source_dir = tmp_path / "源"
    source_dir.mkdir()
    first = source_dir / "杭州分行_金融基础数据-单位贷款_202609.xlsx"
    second = source_dir / "杭州分行_金融基础数据-个人贷款_202609.xlsx"
    for path, sheet_name in ((first, "单位贷款"), (second, "个人贷款")):
        book = Workbook()
        book.active.title = sheet_name
        book.active["A1"] = sheet_name
        book.save(path)
        book.close()

    # DAG 的组合规则来自 Excel 配置，不再依赖旧流程 JSON。测试使用
    # 最小真实配置，覆盖“按文件名关键字”这一执行路径。
    config_dir = tmp_path / "config"
    config_dir.mkdir()
    config_book = Workbook()
    config_book.active.title = "组合方案"
    config_book.active.append(["方案编号", "方案名称", "分组方式", "方案状态", "输出标识", "文件名正则"])
    config_book.active.append(["default", "测试组合", "文件名正则", "当前使用", "测试组合", "(?P<组合>.+?)_金融基础数据-"])
    groups = config_book.create_sheet("组合分组")
    groups.append(["方案编号", "组合名称", "匹配关键字", "启用"])
    config_book.save(config_dir / "config.xlsx")
    config_book.close()

    with patch.object(sys, "platform", "linux"):
        service = _service(tmp_path)
        result = run_dag_flow(
            "dag:组合工作表",
            service=service,
            template_path=None,
            input_dir=source_dir,
            output_dir=tmp_path / "输出",
            period="",
            history_path=tmp_path / "历史审核配置.xlsx",
            recursive=False,
            write_flow_logs=False,
        )

    assert result.successful_items == 1
    output = result.items[0].output_path
    assert output is not None and output.is_file()
    merged = load_workbook(output, read_only=True)
    try:
        assert merged.sheetnames == ["个人贷款", "单位贷款"]
    finally:
        merged.close()


def test_custom_combine_workflow_persistence_is_disabled(tmp_path: Path) -> None:
    """普通编排创建的 custom DAG 要能真实执行，而非只能保存和查看。"""
    from base_audit.workflow.catalog import save_custom_workflow
    from base_audit.workflow.graph import WorkflowBinding, WorkflowDefinition, WorkflowNode
    from base_audit.workflow.runner import run_dag_flow

    source_dir = tmp_path / "源"
    source_dir.mkdir()
    for suffix, sheet_name in (("甲", "表甲"), ("乙", "表乙")):
        path = source_dir / f"宁波分行_金融基础数据-{suffix}.xlsx"
        book = Workbook(); book.active.title = sheet_name; book.save(path); book.close()
    history = tmp_path / "历史审核配置.xlsx"
    service = _service(tmp_path)
    workflow = WorkflowDefinition(
        workflow_id="custom:native-combine",
        name="宁波组合",
        settings={"flow_name": "宁波组合", "template": "combine_sheets"},
        nodes=[WorkflowNode(
            "combine", "excel.combine_sheets", display_name="组合工作表",
            parameters={"flow_name": "宁波组合", "plan_id": "default"},
        )],
        output_bindings=[WorkflowBinding("result", "combine", "result")],
    )
    with pytest.raises(ValueError, match="不保存自定义 DAG"):
        save_custom_workflow(history, workflow)


def test_calc_pr_flags_restored_without_touching_cached_values(tmp_path: Path) -> None:
    """soffice 回写会丢 calcPr 标记；补丁须置位且不破坏公式缓存（UOS 验收 #4）。"""
    import re
    import zipfile

    from base_audit.engines.libreoffice import _restore_recalc_flags

    def make_book(path: Path) -> None:
        book = Workbook()
        sheet = book.active
        sheet["A1"], sheet["B1"] = 2, 3
        sheet["C1"] = "=A1*B1"
        book.save(path)
        book.close()

    def cached_formula_cells(path: Path) -> int:
        count = 0
        with zipfile.ZipFile(path) as z:
            for name in z.namelist():
                if re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                    xml = z.read(name).decode("utf-8", "ignore")
                    count += sum(
                        1 for m in re.finditer(r"<c[^>]*>(.*?)</c>", xml, re.S)
                        if "<f>" in m.group(1)
                    )
        return count

    target = tmp_path / "副本.xlsx"
    make_book(target)
    before = cached_formula_cells(target)

    _restore_recalc_flags(target)

    with zipfile.ZipFile(target) as z:
        workbook_xml = z.read("xl/workbook.xml").decode("utf-8")
    assert 'fullCalcOnLoad="1"' in workbook_xml
    assert 'forceFullCalc="1"' in workbook_xml
    # 工作表 XML 逐字节未变（缓存与公式都不受影响）
    assert cached_formula_cells(target) == before
    load_workbook(target, read_only=True).close()  # 文件仍可正常打开


def test_dag_audit_dispatches_to_native(tmp_path: Path, monkeypatch) -> None:
    """DAG 审核入口在 native 平台走 run_dag_native_audit（mock LibreOffice 重算）。"""
    from unittest.mock import patch
    from base_audit import engines
    from base_audit.workflow.runner import run_dag_flow

    template = tmp_path / "模板.xlsx"
    _make_template(template)
    from openpyxl import load_workbook
    from openpyxl.workbook.defined_name import DefinedName
    book = load_workbook(template)
    book.defined_names.add(DefinedName("校验区域", attr_text="报表!$D$2:$D$2"))
    book.defined_names.add(DefinedName("表结构区域", attr_text="报表!$D$1:$D$1"))
    book.save(template); book.close()
    source = tmp_path / "源" / "机构A.xlsx"
    source.parent.mkdir()
    book = Workbook(); book.active.title = "报表"; book.active["C2"] = 5; book.active["D1"] = "指标"
    book.save(source); book.close()

    def fake_recalculate(self, workbook_path):
        from base_audit.engines.libreoffice import LibreOfficeCalculationResult
        return LibreOfficeCalculationResult(
            workbook_path=workbook_path, engine_display="LibreOffice Calc (mock)", elapsed_seconds=0.0,
        )

    service = _service(tmp_path)
    history = tmp_path / "历史审核配置.xlsx"
    with patch.object(sys, "platform", "linux"), patch.object(
        engines.libreoffice.LibreOfficeCalculator, "recalculate", fake_recalculate
    ), patch.object(engines.libreoffice.LibreOfficeCalculator, "require_available", lambda self: None):
        result = run_dag_flow(
            "dag:汇总核查表校验", service=service, template_path=template,
            input_dir=source.parent, output_dir=tmp_path / "输出",
            period="2026-08", history_path=history, recursive=False,
        )
    assert result.successful_files == 1 and result.failed_files == 0


# ---------------------------------------------------------------------------
# 报表采集系统：跨期比较配置绑定（pcConfig）
# ---------------------------------------------------------------------------

def test_period_config_file_round_trip(tmp_path: Path) -> None:
    stored = settings_mod.UserSettings(period_config_file="D:/配置/报表采集系统_比较配置.xlsx")
    assert stored.period_config_file == "D:/配置/报表采集系统_比较配置.xlsx"


def test_choose_period_compare_binds_config(tmp_path: Path) -> None:
    from types import SimpleNamespace
    from base_audit.web_app import WebApi

    target = tmp_path / "我的跨期比较配置.xlsx"
    window = SimpleNamespace(create_file_dialog=lambda *args, **kwargs: (str(target),))
    webview = SimpleNamespace(FileDialog=SimpleNamespace(OPEN=10, FOLDER=20), windows=[window])
    with patch.dict(sys.modules, {"webview": webview}):
        api = WebApi(tmp_path)
        value = api.choose_period_compare("pcConfig")
    assert value == str(target.resolve())
    assert api.state["pcConfig"] == str(target.resolve())
    assert api.settings.period_config_file == str(target.resolve())
    # 重新加载后保持绑定
    reloaded = WebApi(tmp_path)
    assert reloaded.state["pcConfig"] == str(target.resolve())


def test_state_has_default_period_config(tmp_path: Path) -> None:
    from base_audit.web_app import WebApi
    api = WebApi(tmp_path)
    assert api.state["pcConfig"] == str(tmp_path / "config" / "2.报表采集系统_配置.xlsx")


def test_central_tolerance_default_and_round_trip(tmp_path: Path) -> None:
    from base_audit.web_app import WebApi

    api = WebApi(tmp_path)
    assert api.state["centralTolerance"] == 100.0
    assert api.settings.central_diff_tolerance_yuan == 100.0
    # 页面推送新容差 → 状态与设置持久化；重载后仍保留
    api.update({"centralTolerance": 250})
    assert api.state["centralTolerance"] == 250.0
    reloaded = WebApi(tmp_path)
    assert reloaded.state["centralTolerance"] == 250.0


def test_central_tolerance_invalid_falls_back_to_default(tmp_path: Path) -> None:
    from base_audit.settings import SettingsStore
    from base_audit.web_app import WebApi

    api = WebApi(tmp_path)
    api.update({"centralTolerance": "abc"})
    assert api.state["centralTolerance"] == 100.0
    api.update({"centralTolerance": -5})
    assert api.state["centralTolerance"] == 100.0
    # 设置文件中的非法/越界值在加载时同样回落默认 100
    store = tmp_path / "用户设置.json"
    store.write_text('{"central_diff_tolerance_yuan": "abc"}', encoding="utf-8")
    assert SettingsStore(store).load().central_diff_tolerance_yuan == 100.0
    store.write_text('{"central_diff_tolerance_yuan": 99999999}', encoding="utf-8")
    assert SettingsStore(store).load().central_diff_tolerance_yuan == 100.0


# ---------------------------------------------------------------------------
# 配置文件改名迁移（历史：→逐笔统计系统_历史审核配置；比较：→报表采集系统_比较配置）
# ---------------------------------------------------------------------------

def test_legacy_history_workbook_migrates_to_prefixed_name(tmp_path: Path) -> None:
    from base_audit.name_config import (
        HISTORY_WORKBOOK_NAME, initialize_config,
    )
    legacy = tmp_path / "历史审核配置.xlsx"
    from openpyxl import Workbook
    book = Workbook(); book.active["A1"] = "旧历史数据"
    book.save(legacy); book.close()
    target = tmp_path / HISTORY_WORKBOOK_NAME
    initialize_config(target)
    assert target.is_file() and not legacy.exists()
    restored = load_workbook(target, read_only=True)
    # 迁移保留旧数据；程序随后会在最前插入“使用说明”表，因此按内容查找。
    assert any(
        restored[name].cell(1, 1).value == "旧历史数据" for name in restored.sheetnames
    )
    restored.close()


def test_legacy_period_config_migrates_to_prefixed_name(tmp_path: Path) -> None:
    from base_audit.period_compare import (
        CONFIG_WORKBOOK_NAME, LEGACY_CONFIG_WORKBOOK_NAME, ensure_default_config,
    )
    legacy = tmp_path / LEGACY_CONFIG_WORKBOOK_NAME
    book = Workbook(); book.active["A1"] = "用户自维护指标"
    book.save(legacy); book.close()
    target = tmp_path / CONFIG_WORKBOOK_NAME
    created = ensure_default_config(target)
    assert created is False  # 走的是迁移，不是新建
    assert target.is_file() and not legacy.exists()
    restored = load_workbook(target, read_only=True)
    assert restored.active["A1"].value == "用户自维护指标"
    restored.close()
