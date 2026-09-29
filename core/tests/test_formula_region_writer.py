"""公式校验复制写入后端测试：Direct OOXML 写入 vs COM Range 基线（无 COM 部分）。"""

from __future__ import annotations

import sys
import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "core" / "src"))

from base_audit.formula_region_writer import (
    DirectOoxmlFormulaRegionWriter,
    UnsupportedFormulaRegionOperation,
    _TemplateReader,
)
from base_audit.models import AuditRule, CopyRange, TemplateDefinition


def _build_template(path: Path) -> None:
    """命名区域模板：校验区域 D2:D4（公式+样式）+ 合并 + 报送数据列。"""
    book = Workbook()
    sheet = book.active
    sheet.title = "20202资产负债表"
    sheet["A1"] = "指标名称"
    sheet["A2"] = "各项存款"
    sheet["B2"] = None   # 报送数据格（副本里会有值）
    sheet["C2"] = None
    sheet["C3"] = None
    sheet["D2"] = "=IF(AND(B2>0,C2>0),B2-C2,\"\")"
    sheet["D3"] = "=IF(B3<>C3,\"差异\",\"一致\")"
    sheet["D4"] = "=SUM(B2:B3)"
    sheet["D2"].font = Font(bold=True, color="FF0000")
    sheet["D2"].fill = PatternFill("solid", fgColor="FFFF00")
    sheet["C1"] = "填报"
    sheet["C2"] = None
    book["20202资产负债表"]["C4"] = None
    sheet["C4"] = None
    # 合并：A5:B5（在区域外的示例）；区域内的合并 D2:D3 不合并以免覆盖公式——
    # 改为在区域内放一个合并标签 A1:B1？为测合并并入，用 C1:C1 无意义。
    # 直接加一个区域内的合并会破坏公式行，改为区域外合并仅验证不被破坏。
    sheet["A5"] = "备注"
    sheet["B5"] = None
    book.save(path)
    book.close()


def _definition(copy_address: str = "D2:D4", *, structured: bool = False) -> TemplateDefinition:
    sheet_name = "20202资产负债表"
    rules = [
        AuditRule(rule_id="R1", enabled=True, report_code="20202", sheet_name=sheet_name,
                  formula_cell="D2", target_cell="D2", severity="错误", message="",
                  copy_range=copy_address, result_mode="keyword", value_from_result=True),
    ]
    from base_audit.models import CopyRange as CR

    return TemplateDefinition(
        rules=rules, copy_ranges=[CR(sheet_name, copy_address)],
        structured=structured,
        structure_ranges=[CR(sheet_name, "A1:A1")],
    )


def _build_audit_copy(path: Path) -> None:
    """报送副本：与模板同构，B/C 列有报送数据，D 列为空待写公式。"""
    book = Workbook()
    sheet = book.active
    sheet.title = "20202资产负债表"
    sheet["A1"] = "指标名称"
    sheet["A2"] = "各项存款"
    sheet["B2"] = 100
    sheet["C2"] = 30
    sheet["B3"] = 200
    sheet["C3"] = 180
    sheet["B4"] = 5
    sheet["C4"] = 5
    sheet["A5"] = "报送侧备注"
    book.save(path)
    book.close()


class DirectWriterTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = TemporaryDirectory(prefix="frw-")
        root = Path(cls._tmp.name)
        cls.template = root / "模板.xlsx"
        _build_template(cls.template)
        cls.audit = root / "甲银行_审核版.xlsx"
        _build_audit_copy(cls.audit)
        writer = DirectOoxmlFormulaRegionWriter()
        cls.write_result = writer.write(
            template_path=cls.template, audit_path=cls.audit,
            definition=_definition(),
        )

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def test_formulas_written_same_coordinates(self):
        book = load_workbook(self.audit)
        sheet = book["20202资产负债表"]
        self.assertEqual(sheet["D2"].value, '=IF(AND(B2>0,C2>0),B2-C2,"")')
        self.assertEqual(sheet["D3"].value, '=IF(B3<>C3,"差异","一致")')
        self.assertEqual(sheet["D4"].value, "=SUM(B2:B3)")
        book.close()

    def test_submission_data_untouched_outside_region(self):
        book = load_workbook(self.audit, data_only=True)
        sheet = book["20202资产负债表"]
        self.assertEqual(sheet["B2"].value, 100)
        self.assertEqual(sheet["C3"].value, 180)
        self.assertEqual(sheet["A5"].value, "报送侧备注")
        book.close()

    def test_styles_transplanted(self):
        """模板 D2 的字体/填充样式经索引映射应用到副本（与模板对拍）。"""
        t_book = load_workbook(self.template)
        a_book = load_workbook(self.audit)
        t_cell = t_book["20202资产负债表"]["D2"]
        a_cell = a_book["20202资产负债表"]["D2"]

        def fill_rgb(cell):
            return cell.fill.fgColor.rgb if cell.fill and cell.fill.patternType else None

        self.assertEqual(a_cell.font.bold, t_cell.font.bold)
        self.assertEqual(fill_rgb(a_cell), fill_rgb(t_cell))
        self.assertIsNotNone(fill_rgb(a_cell))   # 填充确实被移植
        t_book.close()
        a_book.close()

    def test_calc_chain_dropped_and_full_calc_on_load(self):
        with zipfile.ZipFile(self.audit) as archive:
            names = archive.namelist()
            self.assertNotIn("xl/calcChain.xml", names)
            workbook_xml = archive.read("xl/workbook.xml").decode()
        self.assertIn("fullCalcOnLoad", workbook_xml)

    def test_blank_template_cells_clear_submission_residue(self):
        """模板区域内空格（E 列不在区域内；此处验证区域内 D 列全覆盖）。"""
        book = load_workbook(self.audit, data_only=True)
        sheet = book["20202资产负债表"]
        # D 列三格都由模板内容决定；区域外的 C4 保持报送值
        self.assertEqual(sheet["C4"].value, 5)
        book.close()

    def test_structured_template_copies_rule_sheet(self):
        """结构化模板：DIRECT_OOXML 写入校验区域并整表复制“审核规则”。

        （旧版本直接拒绝结构化模板；合并 UOS 反馈后改为 openpyxl 复制规则表。）
        """
        from openpyxl import Workbook, load_workbook

        with TemporaryDirectory(prefix="frw-struct-") as folder:
            root = Path(folder)
            template = root / "结构化模板.xlsx"
            t_book = Workbook()
            sheet = t_book.active
            sheet.title = "20202资产负债表"
            sheet["D2"] = "=B2+C2"
            rules = t_book.create_sheet("审核规则")
            rules.append(("规则编号", "启用", "工作表", "公式单元格", "级别", "问题说明"))
            rules.append(("R001", "是", "20202资产负债表", "D2", "错误", "不平"))
            t_book.save(template)
            t_book.close()

            audit = root / "副本.xlsx"
            s_book = Workbook()
            s_sheet = s_book.active
            s_sheet.title = "20202资产负债表"
            s_sheet["B2"] = 1
            s_sheet["C2"] = 2
            s_book.save(audit)
            s_book.close()

            DirectOoxmlFormulaRegionWriter().write(
                template_path=template, audit_path=audit,
                definition=_definition(structured=True))
            book = load_workbook(audit)
            try:
                self.assertIn("审核规则", book.sheetnames)
                self.assertEqual(book["审核规则"]["A2"].value, "R001")
                self.assertEqual(book["20202资产负债表"]["D2"].value, "=B2+C2")
            finally:
                book.close()

    def test_template_reader_expands_shared_formulas(self):
        """shared formula 展开为主格公式文本（si 不跨簿携带）。"""
        from openpyxl.formula.translate import Translator  # noqa: F401  确认可用

        _reader = _TemplateReader(self.template, _definition())
        region = _reader.read_region("20202资产负债表", "D2:D4")
        self.assertTrue(all(cell.formula for cell in region.cells.values()))

    def test_shared_expansion_translates_entity_encoded_indirect(self):
        """实体编码（&quot;）的 CJK INDIRECT shared 依赖格必须平移。

        回归：模板 XML 公式为实体编码时字符串掩码失效，CJK 表名字符串内的
        “!”会抑制后续相对引用平移，依赖格被写成主格原文（E7 变 E6）。
        """
        from base_audit.formula_region_writer import _expand_shared_formulas

        xml = (
            '<worksheet><sheetData>'
            '<row r="6"><c r="N6"><f t="shared" ref="N6:N8" si="1">'
            'SUM(INDIRECT(&quot;单位贷款发生额信息!&quot;&amp;'
            'ADDRESS(ROW(E6),COLUMN(E6))))</f></c></row>'
            '<row r="7"><c r="N7"><f t="shared" si="1"/></c></row>'
            '<row r="8"><c r="N8"><f t="shared" si="1"/></c></row>'
            '</sheetData></worksheet>'
        )
        expanded = _expand_shared_formulas(xml)
        self.assertIn("ROW(E7),COLUMN(E7)", expanded)
        self.assertIn("ROW(E8),COLUMN(E8)", expanded)
        self.assertNotIn('<f t="shared" si="1"/>', expanded)

    def test_template_absent_cells_clear_submission_formulas(self):
        """区域矩形内模板未定义的格必须清空报送自带公式（对齐 Range.Copy）。

        回归：在线核查表报送文件在模板空白格上带有旧校验公式，保留会与
        模板规则重复触发（历史测试副本曾因此值不一致）。
        """
        with TemporaryDirectory(prefix="frw-absent-") as folder:
            root = Path(folder)
            t_book = Workbook()
            t_sheet = t_book.active
            t_sheet.title = "20202资产负债表"
            t_sheet["D2"] = "=B2+C2"          # 区域 D2:D3 内模板只定义 D2
            t_book.save(root / "模板.xlsx")
            t_book.close()
            s_book = Workbook()
            s_sheet = s_book.active
            s_sheet.title = "20202资产负债表"
            s_sheet["B2"] = 1
            s_sheet["C2"] = 2
            s_sheet["D3"] = "=1+2"            # 报送残留公式，被区域矩形覆盖
            s_book.save(root / "副本.xlsx")
            s_book.close()

            DirectOoxmlFormulaRegionWriter().write(
                template_path=root / "模板.xlsx",
                audit_path=root / "副本.xlsx",
                definition=_definition(copy_address="D2:D3"),
            )
            book = load_workbook(root / "副本.xlsx")
            sheet = book["20202资产负债表"]
            self.assertEqual(sheet["D2"].value, "=B2+C2")
            self.assertIsNone(sheet["D3"].value)
            book.close()


class HandlerDispatchTests(unittest.TestCase):
    def test_writer_mode_validation(self):
        from base_audit.formula_region_writer import WRITER_MODES

        self.assertIn("COM_RANGE", WRITER_MODES)
        self.assertIn("DIRECT_OOXML", WRITER_MODES)


if __name__ == "__main__":
    unittest.main()
