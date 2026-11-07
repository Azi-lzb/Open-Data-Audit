"""Windows Excel COM 管线的副本修复：全角括号表名裸引用 + 冻结窗格矛盾。

Excel 拒绝打开这两类非规范报送文件（WPS 能打开）；审核流程在复制审核副本时
对副本做无害规范化，原始文件不动。测试用合成工作簿复现两类缺陷并验证修复，
另用 handler 集成用例验证引擎门控（Excel/自动修、WPS 不修）。
"""

from __future__ import annotations

import re
import shutil
import sys
import zipfile
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from openpyxl import Workbook, load_workbook

from base_audit.excel_open_repair import repair_audit_copy_for_excel
from base_audit.workflow.adapters.audit_handlers import handle_files_copy
from base_audit.workflow.artifacts import FileSetArtifact, FileRecord
from base_audit.workflow.types import ArtifactSubtype

HIDDEN_SHEET = "个人贷款发生额信息（隐藏发放）"


def _defective_workbook(path: Path) -> None:
    """生成含两类缺陷的工作簿：裸引用 + 冻结窗格行矛盾。"""
    book = Workbook()
    main = book.active
    main.title = "发生额"
    main["A1"] = f"={HIDDEN_SHEET}!B2+1"
    main["A2"] = 5
    book.create_sheet(HIDDEN_SHEET)["B2"] = 3
    main.freeze_panes = "C6"  # openpyxl 写出 ySplit=5 + topLeftCell=C6（一致）
    book.save(path)
    book.close()
    # 注入缺陷：把 ySplit 改成 6，使 topLeftCell 的行 6 落进冻结区（Excel 拒绝）；
    # 同时把 openpyxl 写出的数字字符引用（&#20010;…）解码为字面 UTF-8，
    # 与真实缺陷文件（WPS 生成，字面存储）的形态一致。
    def _decode_ncr(text: str) -> str:
        return re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), text)

    _patch_zip(path, {"xl/worksheets/sheet1.xml": [('ySplit="5"', 'ySplit="6"')]})
    # 单独处理：先按 NCR 解码再写回（_patch_zip 只做字符串替换，这里直接重打包）
    import zipfile as _zf
    with _zf.ZipFile(path) as zin:
        parts = {name: zin.read(name) for name in zin.namelist()}
    parts["xl/worksheets/sheet1.xml"] = _decode_ncr(
        parts["xl/worksheets/sheet1.xml"].decode("utf-8")).encode("utf-8")
    with _zf.ZipFile(path, "w", _zf.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(name, data)


def _patch_zip(path: Path, replacements: dict[str, list[tuple[str, str]]]) -> None:
    with zipfile.ZipFile(path) as zin:
        parts = {name: zin.read(name) for name in zin.namelist()}
    for name, pairs in replacements.items():
        text = parts[name].decode("utf-8")
        for old, new in pairs:
            assert old in text, f"{name} 缺少待替换片段：{old}"
            text = text.replace(old, new)
        parts[name] = text.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(name, data)


def _context(engine: str):
    return SimpleNamespace(
        log=lambda message: None,
        run=SimpleNamespace(services={"service": SimpleNamespace(engine_preference=engine)},
                            resources={}),
    )


def _params(tmp_path: Path) -> dict:
    return {"stage_dir": str(tmp_path / "stage"), "flow_name": "S1-F01"}


def _inputs(source: Path) -> dict:
    record = FileRecord(source=str(source), org_code="机构", org_name="机构",
                        version=0, stage_paths=(str(source),))
    return {"accepted_files": FileSetArtifact(
        records=(record,), subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET)}


def test_repair_quotes_bare_fullwidth_paren_refs(tmp_path: Path) -> None:
    src = tmp_path / "原始.xlsx"
    _defective_workbook(src)
    before = src.read_bytes()

    result = repair_audit_copy_for_excel(src)

    assert result["changed"] is True
    assert result["quoted"] >= 1, result
    assert result["pane_fixed"] == 1, result
    repaired = load_workbook(src, data_only=False)
    try:
        formula = str(repaired["发生额"]["A1"].value)
        assert f"'{HIDDEN_SHEET}'!B2" in formula, formula
        pane = repaired["发生额"].freeze_panes
        assert str(pane) == "C7", f"冻结窗格 topLeftCell 应修正到 ySplit+1 行：{pane}"
    finally:
        repaired.close()
    assert src.read_bytes() != before


def test_repair_leaves_clean_workbook_untouched(tmp_path: Path) -> None:
    src = tmp_path / "正常.xlsx"
    book = Workbook()
    book.active["A1"] = 1
    book.save(src)
    book.close()
    before = src.read_bytes()

    result = repair_audit_copy_for_excel(src)

    assert result["changed"] is False
    assert src.read_bytes() == before, "无缺陷文件必须保持字节不变"


def test_repair_skips_non_zip_gracefully(tmp_path: Path) -> None:
    src = tmp_path / "旧格式.xlsx"
    src.write_bytes(b"\xd0\xcf\x11\xe0 not a zip")
    before = src.read_bytes()

    result = repair_audit_copy_for_excel(src)

    assert result["changed"] is False
    assert result.get("reason")
    assert src.read_bytes() == before


def test_files_copy_repairs_for_excel_engine(tmp_path: Path) -> None:
    src = tmp_path / "报送.xlsx"
    _defective_workbook(src)
    original = src.read_bytes()

    result = handle_files_copy(_context("Microsoft Excel"), _params(tmp_path), _inputs(src))

    record = result.outputs["workbooks"].records[0]
    assert result.metrics["excel_open_repair_files"] == 1
    assert result.metrics["excel_open_repair_quoted"] >= 1
    repaired = load_workbook(Path(record.audit), data_only=False)
    try:
        assert f"'{HIDDEN_SHEET}'!" in str(repaired["发生额"]["A1"].value)
    finally:
        repaired.close()
    assert src.read_bytes() == original, "原始报送文件不得改动"


def test_files_copy_skips_repair_for_wps_engine(tmp_path: Path) -> None:
    src = tmp_path / "报送.xlsx"
    _defective_workbook(src)
    before = src.read_bytes()

    result = handle_files_copy(_context("WPS 表格"), _params(tmp_path), _inputs(src))

    record = result.outputs["workbooks"].records[0]
    assert result.metrics["excel_open_repair_files"] == 0
    assert Path(record.audit).read_bytes() == before, "WPS 引擎不做修复"


def test_files_copy_repairs_for_auto_engine(tmp_path: Path) -> None:
    src = tmp_path / "报送.xlsx"
    _defective_workbook(src)
    original = src.read_bytes()

    result = handle_files_copy(_context("自动"), _params(tmp_path), _inputs(src))

    assert result.metrics["excel_open_repair_files"] == 1
    assert src.read_bytes() == original


def test_real_world_defect_files_repaired(tmp_path: Path) -> None:
    """真实缺陷样本（隔离目录）修复计数与根因分析一致；目录缺失时跳过。"""
    inputs = (Path(__file__).resolve().parents[2] / "test-output"
              / "s1f1_com_isolation_20261107" / "source_bundle" / "inputs_129")
    expected = {
        "TCL科技集团财务有限公司_金融基础数据-票据业务_B00_2026-08-31_在线核查表.xlsx": 2928,
        "TCL科技集团财务有限公司_金融基础数据-个人贷款_B00_2026-08-31_在线核查表.xlsx": 456,
        "TCL科技集团财务有限公司_金融基础数据采集（单位贷款置换旧债发生额）_2026-08-31_数据核对表.xlsx": 0,
    }
    if not inputs.is_dir():
        pytest_skip = __import__("pytest")
        pytest_skip.skip("本机无隔离测试样本目录")
    for name, quoted in expected.items():
        src = inputs / name
        if not src.is_file():
            continue
        work = tmp_path / name
        shutil.copy2(src, work)
        result = repair_audit_copy_for_excel(work)
        assert result["quoted"] == quoted, (name, result)
        assert result["changed"] is True
        assert work.read_bytes() != src.read_bytes()
