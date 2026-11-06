"""外部文件添加的降级链：Direct OOXML → openpyxl → 记录失败但链路继续。

S1-F01 的 ``excel.external_sheet_copy`` 节点在 UOS/native 与 Windows/COM 两条
适配器上都必须能在 Direct OOXML 拒绝复杂外部表时继续添加工作表；两个写入器
都失败时也只能标记该文件失败，不能中止整条审核链。
"""

from __future__ import annotations

import hashlib
from pathlib import Path
import posixpath
import shutil
from types import SimpleNamespace
import zipfile
from xml.etree import ElementTree as ET

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import column_index_from_string, coordinate_to_tuple

from base_audit.external import ExternalSheetPlan
from base_audit.workflow.adapters.native_audit_handlers import handle_external_copy
from base_audit.workflow.artifacts import FileSetArtifact, FileRecord
from base_audit.workflow.types import ArtifactSubtype, NodeRunStatus


_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _audit_workbook(path: Path) -> None:
    book = Workbook()
    book.active.title = "报表"
    book.active["A1"] = "数据"
    book.save(path)
    book.close()


def _audit_with_merged_payload(path: Path) -> None:
    """Create a source workbook with real shared-string cells inside merges."""
    book = Workbook()
    sheet = book.active
    sheet.title = "报表"
    sheet["A1"] = "原始数据"
    sheet.merge_cells("B5:B69")
    sheet.merge_cells("C5:C25")
    sheet.merge_cells("E5:E10")
    sheet["J1"] = "=1+1"
    book.save(path)
    book.close()

    with zipfile.ZipFile(path, "r") as source:
        parts = {name: source.read(name) for name in source.namelist()}
    workbook = ET.fromstring(parts["xl/workbook.xml"])
    rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    relation_by_id = {
        item.attrib["Id"]: item
        for item in rels.findall(f"{{{_PKG_REL_NS}}}Relationship")
    }
    report = next(sheet for sheet in workbook.find(f"{{{_MAIN_NS}}}sheets") if sheet.attrib["name"] == "报表")
    target = relation_by_id[report.attrib[f"{{{_DOC_REL_NS}}}id"]].attrib["Target"]
    part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
    root = ET.fromstring(parts[part])
    sheet_data = root.find(f"{{{_MAIN_NS}}}sheetData")
    assert sheet_data is not None

    strings = ["原始B8值", "原始C8值", "另一个合并区域值"]
    for address, index in (("B8", 0), ("C8", 1), ("E8", 2)):
        row_number, _column = coordinate_to_tuple(address)
        row = next((item for item in sheet_data if int(item.attrib["r"]) == row_number), None)
        if row is None:
            row = ET.Element(f"{{{_MAIN_NS}}}row", {"r": str(row_number)})
            sheet_data.append(row)
        cell = ET.Element(f"{{{_MAIN_NS}}}c", {"r": address, "t": "s"})
        ET.SubElement(cell, f"{{{_MAIN_NS}}}v").text = str(index)
        row.append(cell)
        row[:] = sorted(
            row,
            key=lambda item: column_index_from_string(
                "".join(ch for ch in item.attrib.get("r", "A1") if ch.isalpha()) or "A"
            ),
        )
    sheet_data[:] = sorted(sheet_data, key=lambda item: int(item.attrib["r"]))
    ET.register_namespace("", _MAIN_NS)
    parts[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    shared_root = ET.Element(
        f"{{{_MAIN_NS}}}sst",
        {"count": str(len(strings)), "uniqueCount": str(len(strings))},
    )
    for value in strings:
        item = ET.SubElement(shared_root, f"{{{_MAIN_NS}}}si")
        ET.SubElement(item, f"{{{_MAIN_NS}}}t").text = value
    parts["xl/sharedStrings.xml"] = ET.tostring(shared_root, encoding="utf-8", xml_declaration=True)

    existing_ids = [
        int(item.attrib["Id"][3:])
        for item in rels.findall(f"{{{_PKG_REL_NS}}}Relationship")
        if item.attrib.get("Id", "").startswith("rId") and item.attrib["Id"][3:].isdigit()
    ]
    ET.SubElement(
        rels,
        f"{{{_PKG_REL_NS}}}Relationship",
        {
            "Id": "rId{}".format(max(existing_ids, default=0) + 1),
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
            "Target": "sharedStrings.xml",
        },
    )
    ET.register_namespace("", _PKG_REL_NS)
    parts["xl/_rels/workbook.xml.rels"] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)

    content_types = ET.fromstring(parts["[Content_Types].xml"])
    ET.SubElement(
        content_types,
        f"{{{_CONTENT_NS}}}Override",
        {
            "PartName": "/xl/sharedStrings.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        },
    )
    ET.register_namespace("", _CONTENT_NS)
    parts["[Content_Types].xml"] = ET.tostring(content_types, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        for name, payload in parts.items():
            target.writestr(name, payload)


def _sheet_parts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path, "r") as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        by_id = {item.attrib["Id"]: item for item in rels.findall(f"{{{_PKG_REL_NS}}}Relationship")}
        result = {}
        for sheet in workbook.find(f"{{{_MAIN_NS}}}sheets"):
            target = by_id[sheet.attrib[f"{{{_DOC_REL_NS}}}id"]].attrib["Target"]
            result[sheet.attrib["name"]] = (
                target.lstrip("/") if target.startswith("/")
                else posixpath.normpath(posixpath.join("xl", target))
            )
        return result


def _cell_text_values(path: Path, sheet_name: str, addresses: tuple[str, ...]) -> dict[str, str]:
    with zipfile.ZipFile(path, "r") as archive:
        root = ET.fromstring(archive.read(_sheet_parts(path)[sheet_name]))
        shared_strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared_strings = ["".join(item.itertext()) for item in shared_root.findall(f"{{{_MAIN_NS}}}si")]
        cells = {
            cell.attrib.get("r", ""): cell
            for cell in root.iter(f"{{{_MAIN_NS}}}c")
        }
        result = {}
        for address in addresses:
            cell = cells[address]
            if cell.attrib.get("t") == "s":
                index = int(cell.find(f"{{{_MAIN_NS}}}v").text)
                result[address] = shared_strings[index]
            else:
                inline = cell.find(f"{{{_MAIN_NS}}}is")
                assert inline is not None
                result[address] = "".join(inline.itertext())
        return result


def _plain_external(path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "集中系统数据"
    sheet["A1"] = "机构"
    sheet["B2"] = 123.0
    sheet["C2"] = "文本"
    book.save(path)
    book.close()


def _complex_external(path: Path) -> None:
    """模拟真实业务外部表：合并单元格 + 公式，Direct OOXML 明确不支持。"""
    book = Workbook()
    sheet = book.active
    sheet.title = "参照表"
    sheet["A1"] = "机构代码"
    sheet["B1"] = "机构名称"
    sheet.merge_cells("A1:B1")
    sheet["A2"] = "A001"
    sheet["B2"] = '=CONCAT("机构",A2)'
    book.save(path)
    book.close()


def _formula_external(path: Path) -> None:
    """A formula-bearing table without features requiring openpyxl fallback."""
    book = Workbook()
    sheet = book.active
    sheet.title = "参照表"
    sheet["A1"] = "代码"
    sheet["D2"] = "=1+1"
    book.save(path)
    book.close()


def _mixed_external(path: Path) -> None:
    """纯数据表与公式表共存，贴近真实外部文件结构。"""
    book = Workbook()
    plain = book.active
    plain.title = "集中系统数据"
    plain["A1"] = "机构"
    plain["B2"] = 456.0
    formula = book.create_sheet("参照表")
    formula["A1"] = "代码"
    formula["A2"] = "A001"
    formula["D2"] = "=VLOOKUP(A2,参照表!A:B,2,0)"
    book.save(path)
    book.close()


def _context(tmp_path: Path) -> SimpleNamespace:
    return SimpleNamespace(
        log=lambda message: None,
        run=SimpleNamespace(services={"service": SimpleNamespace(external_sheet_writer="DIRECT_OOXML")}),
    )


def _params(tmp_path: Path, external: Path) -> dict:
    return {
        "stage_dir": str(tmp_path / "stage"),
        "external_path": str(external),
        "feature_name": "外部文件添加",
        "flow_name": "S1-F01",
    }


def _inputs(audit: Path, sheets: tuple[str, ...], source: Path | None = None) -> dict:
    record = FileRecord(
        source=str(source or audit),
        audit=str(audit),
        org_code="机构",
        org_name="机构",
        version=0,
        stage_paths=(str(audit),),
    )
    return {
        "template": SimpleNamespace(payload={"external_plan": ExternalSheetPlan(sheets, "公式识别")}),
        "workbooks": FileSetArtifact(records=(record,), subtype=ArtifactSubtype.EXCEL_WORKBOOK_SET),
    }


def test_plain_external_uses_direct_ooxml_fast_path(tmp_path: Path) -> None:
    external = tmp_path / "外部纯数据.xlsx"
    _plain_external(external)
    audit = tmp_path / "审核.xlsx"
    _audit_workbook(audit)

    result = handle_external_copy(
        _context(tmp_path), _params(tmp_path, external), _inputs(audit, ("集中系统数据",)),
    )

    assert result.status is NodeRunStatus.SUCCEEDED
    assert result.metrics["external_openpyxl_fallback_writes"] == 0
    assert result.metrics["external_failed_files"] == 0
    written = Path(result.outputs["workbooks"].records[0].audit)
    book = load_workbook(written, data_only=False)
    try:
        assert "集中系统数据" in book.sheetnames
    finally:
        book.close()


def test_complex_external_falls_back_to_openpyxl_instead_of_aborting(tmp_path: Path) -> None:
    """合并单元格/公式曾让 native 节点直接抛错，整条 S1-F01 链被中止。"""
    external = tmp_path / "外部复杂.xlsx"
    _complex_external(external)
    audit = tmp_path / "审核.xlsx"
    _audit_workbook(audit)

    result = handle_external_copy(
        _context(tmp_path), _params(tmp_path, external), _inputs(audit, ("参照表",)),
    )

    assert result.status is NodeRunStatus.SUCCEEDED
    assert result.metrics["external_failed_files"] == 0
    written = Path(result.outputs["workbooks"].records[0].audit)
    book = load_workbook(written, data_only=False)
    try:
        assert book.sheetnames == ["报表", "参照表"]
        sheet = book["参照表"]
        assert [str(item) for item in sheet.merged_cells.ranges] == ["A1:B1"]
        assert sheet["B2"].value == '=CONCAT("机构",A2)'
    finally:
        book.close()


def test_formula_external_uses_direct_path_and_keeps_target_merged_values(tmp_path: Path) -> None:
    external = tmp_path / "外部公式.xlsx"
    _formula_external(external)
    source = tmp_path / "源文件.xlsx"
    _audit_with_merged_payload(source)
    source_hash_before = hashlib.sha256(source.read_bytes()).hexdigest()
    audit = tmp_path / "审核副本.xlsx"
    shutil.copy2(source, audit)
    expected_values = {"B8": "原始B8值", "C8": "原始C8值", "E8": "另一个合并区域值"}

    result = handle_external_copy(
        _context(tmp_path), _params(tmp_path, external), _inputs(audit, ("参照表",), source=source),
    )

    assert result.status is NodeRunStatus.SUCCEEDED
    assert result.metrics["external_openpyxl_fallback_writes"] == 0
    assert result.metrics["external_failed_files"] == 0
    assert hashlib.sha256(source.read_bytes()).hexdigest() == source_hash_before
    written = Path(result.outputs["workbooks"].records[0].audit)
    assert _cell_text_values(written, "报表", tuple(expected_values)) == expected_values
    restored = load_workbook(written, data_only=False)
    try:
        assert restored.sheetnames == ["报表", "参照表"]
        assert restored["参照表"]["D2"].value == "=1+1"
    finally:
        restored.close()


def test_openpyxl_fallback_preserves_values_inside_merged_ranges(tmp_path: Path) -> None:
    external = tmp_path / "外部复杂.xlsx"
    _complex_external(external)
    source = tmp_path / "源文件.xlsx"
    _audit_with_merged_payload(source)
    source_hash_before = hashlib.sha256(source.read_bytes()).hexdigest()
    audit = tmp_path / "审核副本.xlsx"
    shutil.copy2(source, audit)
    expected_values = {
        "B8": "原始B8值",
        "C8": "原始C8值",
        "E8": "另一个合并区域值",
    }
    assert _cell_text_values(audit, "报表", tuple(expected_values)) == expected_values

    result = handle_external_copy(
        _context(tmp_path), _params(tmp_path, external), _inputs(audit, ("参照表",), source=source),
    )

    assert result.status is NodeRunStatus.SUCCEEDED
    assert result.metrics["external_failed_files"] == 0
    assert hashlib.sha256(source.read_bytes()).hexdigest() == source_hash_before
    written = Path(result.outputs["workbooks"].records[0].audit)
    assert _cell_text_values(written, "报表", tuple(expected_values)) == expected_values
    with zipfile.ZipFile(written, "r") as archive:
        assert archive.testzip() is None
        report_xml = ET.fromstring(archive.read(_sheet_parts(written)["报表"]))
        restored_cells = {
            cell.attrib.get("r"): cell
            for cell in report_xml.iter(f"{{{_MAIN_NS}}}c")
        }
        assert all(restored_cells[address].attrib.get("t") == "inlineStr" for address in expected_values)
        assert "xl/worksheets/sheet2.xml" in archive.namelist()
    restored = load_workbook(written, data_only=False)
    try:
        assert restored.sheetnames == ["报表", "参照表"]
        assert {str(item) for item in restored["报表"].merged_cells.ranges} == {"B5:B69", "C5:C25", "E5:E10"}
        assert restored["报表"]["J1"].value == "=1+1"
        assert restored["参照表"]["B2"].value == '=CONCAT("机构",A2)'
    finally:
        restored.close()


def test_mixed_plan_falls_back_for_whole_batch_and_keeps_formula(tmp_path: Path) -> None:
    """计划同时含纯数据表与公式表时，快路径整批失败，转 openpyxl 后两张表都要在。"""
    external = tmp_path / "外部混合.xlsx"
    _mixed_external(external)
    audit = tmp_path / "审核.xlsx"
    _audit_workbook(audit)

    result = handle_external_copy(
        _context(tmp_path), _params(tmp_path, external),
        _inputs(audit, ("集中系统数据", "参照表")),
    )

    assert result.status is NodeRunStatus.SUCCEEDED
    assert result.metrics["external_failed_files"] == 0
    written = Path(result.outputs["workbooks"].records[0].audit)
    book = load_workbook(written, data_only=False)
    try:
        assert book.sheetnames == ["报表", "集中系统数据", "参照表"]
        assert book["集中系统数据"]["B2"].value == 456.0
        assert book["参照表"]["D2"].value == "=VLOOKUP(A2,参照表!A:B,2,0)"
    finally:
        book.close()


def test_external_copy_keeps_chain_when_both_writers_fail(tmp_path: Path) -> None:
    """openpyxl 也失败时只标记该文件失败，审核副本保留，节点降级为 PARTIAL。"""
    missing = tmp_path / "不存在的外部.xlsx"
    audit = tmp_path / "审核.xlsx"
    _audit_workbook(audit)

    result = handle_external_copy(
        _context(tmp_path), _params(tmp_path, missing), _inputs(audit, ("参照表",)),
    )

    assert result.status is NodeRunStatus.PARTIAL
    assert result.metrics["external_failed_files"] == 1
    assert len(result.partial_failures) == 1
    record = result.outputs["workbooks"].records[0]
    assert record.stage_errors, "失败必须进入文件级结果，不能静默吞掉"
    assert "外部文件添加" in record.stage_errors[0]
    assert Path(record.audit).exists(), "审核副本必须保留以便链路继续"
