"""黄金基线：汇总核查表校验（Windows COM 管线，端到端输出契约）。

冻结现状：副本目录与命名、审核副本内 sheet 顺序、结果工作簿的
sheet/表头/行序/超链接、历史富化、源文件 SHA-256 不变、运行日志。
阶段 0 只冻结行为，不改变任何执行路径；报表采集系统不受影响。
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

import golden_fixtures as gf

pytestmark = [
    pytest.mark.skipif(
        not gf.excel_available(), reason="需要 Windows + 可用的 Excel/WPS COM 引擎"
    ),
    pytest.mark.golden_com,
]

HISTORY_HEADERS = [
    "工作簿名", "工作表名", "定位单元格", "错误类型", "校验指标", "描述",
    "当前值", "对比值", "差值", "规则编号", "历史校验说明", "审核意见",
]
ISSUE_JIA = "甲银行｜20202资产负债表｜D4｜贷款大于存款"


@pytest.fixture()
def audit_env(tmp_path: Path):
    """两家正常机构 + 一个结构损坏文件 + 外部辅助文件 + 预写历史。"""
    src = tmp_path / "源数据"
    src.mkdir()
    gf.build_source(src / "甲银行.xlsx", deposit=2_000_000, loan=3_000_000, reserve=30_000)
    gf.build_source(src / "乙银行.xlsx", deposit=5_000_000, loan=4_000_000, reserve=10_000,
                    comment=gf.CF_COMMENT)
    gf.build_bad_source(src / "丙银行.xlsx")
    template = tmp_path / "模板.xlsx"
    gf.build_audit_template(template)
    external = tmp_path / "外部参照.xlsx"
    gf.build_external(external)
    config = gf.make_config(tmp_path)
    gf.write_history_row(
        config,
        workbook_name="甲银行.xlsx",
        rule_number=ISSUE_JIA,
        feedback="去年已说明",
        opinion="继续关注",
    )
    hashes = {
        path.name: gf.sha256(path)
        for path in sorted(src.glob("*.xlsx"))
    }
    return tmp_path, src, template, external, config, hashes


def test_audit_flow_golden_contract(audit_env):
    tmp_path, src, template, external, config, hashes = audit_env
    from base_audit.service import AuditService
    from base_audit.workflow.runner import run_dag_flow

    output_dir = tmp_path / "输出"
    service = AuditService(config_path=config)
    result = run_dag_flow(
        "dag:汇总核查表校验",
        service=service,
        template_path=template,
        input_dir=src,
        output_dir=output_dir,
        period="2026-06",
        history_path=config,
        external_path=external,
        recursive=False,
        write_flow_logs=True,
    )

    # ---- 运行结果计数：两家成功，结构损坏的丙被筛掉 ----
    assert result.successful_files == 2
    assert result.failed_files == 1
    failed_names = {Path(item.source_path).name for item in result.files if item.error}
    assert failed_names == {"丙银行.xlsx"}

    # ---- 原始报送文件只读：SHA-256 前后一致 ----
    for name, before in hashes.items():
        assert gf.sha256(src / name) == before, name

    # ---- 审核副本目录：默认流程 50 步输出“机构审核副本” ----
    audit_folder = gf.glob_one(output_dir, "机构审核副本_*")
    copy_jia = gf.glob_one(audit_folder, "甲银行_审核版*.xlsx")
    copy_yi = gf.glob_one(audit_folder, "乙银行_审核版*.xlsx")
    assert not list(audit_folder.glob("丙银行*"))
    assert set(result.copies) == {"公式校验复制"}

    # ---- 审核副本：不再写审核导航；外部表追加在报送表之后 ----
    snap_jia = gf.snapshot_workbook(copy_jia)
    # 结构化模板（含“审核规则”表）在写入时整表复制进副本——Windows 与 UOS 一致。
    assert list(snap_jia) == [gf.DATA_SHEET, gf.EXTERNAL_SHEET, "审核规则"]
    # 外部表内容完整复制
    assert snap_jia[gf.EXTERNAL_SHEET]["rows"] == [["基准值", 123]]

    # ---- 结果工作簿：单 sheet、12 列表头、行序固定、超链接指向审核副本 ----
    summary_path = gf.glob_one(output_dir, "汇总核查表校验_[0-9]*.xlsx")
    summary = gf.snapshot_workbook(summary_path)
    assert list(summary) == ["本期审核结果"]
    sheet = summary["本期审核结果"]
    assert sheet["headers"] == HISTORY_HEADERS
    # 行序 = 文件名字典序（Unicode 码点：乙 U+4E59 < 甲 U+7532，故乙先处理），
    # 同一文件内条件格式结果追加在公式结果之后。
    rows = sheet["rows"]
    assert len(rows) == 3
    assert rows[0][0] == "乙银行.xlsx" and rows[0][3] == "错误"
    assert rows[0][4] == "拨备不足" and rows[0][9] == "乙银行｜20202资产负债表｜D5｜拨备不足"
    assert rows[1][0] == "乙银行.xlsx" and rows[1][3] == "条件格式触发"
    # COM 条件格式的规则编号=完整位置身份（含行标签）；native 版不带行标签。
    assert rows[1][4] == "各项存款｜本期情况"
    assert rows[1][9] == "乙银行｜20202资产负债表｜C4｜各项存款｜本期情况"
    assert rows[1][2] == "C4"
    assert rows[2][0] == "甲银行.xlsx" and rows[2][3] == "核实"
    assert rows[2][4] == "贷款大于存款" and rows[2][9] == ISSUE_JIA
    assert rows[2][10] == "去年已说明" and rows[2][11] == "继续关注"
    # 超链接：定位单元格列指向对应工作簿（公式问题→审核副本；条件格式→源文件）
    links = {link["cell"]: link for link in sheet["hyperlinks"]}
    assert Path(links["C2"]["target"]).name == Path(copy_yi).name
    assert links["C2"]["location"] == "'20202资产负债表'!C6"
    assert Path(links["C3"]["target"]).name == Path(copy_yi).name   # 条件格式行链接审核副本
    assert links["C3"]["location"] == "'20202资产负债表'!C4"
    assert Path(links["C4"]["target"]).name == Path(copy_jia).name
    assert links["C4"]["location"] == "'20202资产负债表'!C5"

    # ---- 运行日志：每功能一个 sheet ----
    log_path = gf.glob_one(output_dir, "汇总核查表校验_运行日志_*.xlsx")
    log = gf.snapshot_workbook(log_path)
    assert {"检查校验区域", "检查表结构区域", "表结构比对", "公式校验复制", "校验结果提取"} <= set(log)


def test_audit_flow_rerun_repeats_contract(audit_env):
    """同输入重跑：契约必须可重复（产出数量/命名模式稳定），旧产物不被复用。"""
    tmp_path, src, template, external, config, _hashes = audit_env
    from base_audit.service import AuditService

    output_dir = tmp_path / "输出"
    service = AuditService(config_path=config)
    from base_audit.workflow.runner import run_dag_flow

    for _ in range(2):
        result = run_dag_flow(
            "dag:汇总核查表校验",
            service=service,
            template_path=template,
            input_dir=src,
            output_dir=output_dir,
            period="2026-06",
            history_path=config,
            recursive=False,
            write_flow_logs=False,
        )
        assert result.successful_files == 2
    # 两次运行各生成一套带时间戳的产物
    assert len(list(output_dir.glob("汇总核查表校验_[0-9]*.xlsx"))) == 2
    assert len(list(output_dir.glob("机构审核副本_*"))) == 2
