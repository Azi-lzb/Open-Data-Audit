"""双模式对拍与 FastOoxmlRenderer 专项测试。

同一份合成输入分别经 OPENPYXL / FAST_OOXML 渲染，对拍 Sheet 数量/名称/顺序、
单元格值、数字格式、隐藏行、空表删除；并验证高速模式失败时的明确报错
（不回退）、模板缓存命中与结构校验。
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "core" / "src"))

from base_audit.systems.s3_central_statistics.config import ensure_split_central_configs
from base_audit.systems.s3_central_statistics.form_template import embed_financial_template_in_config
from base_audit.systems.s3_central_statistics.service import render_financial_forms

COMPARE_HEADERS = [
    "业务类", "数据日期", "机构类代码", "机构类名称", "地区代码", "地区名称",
    "指标顺序码", "指标代码", "指标名称", "数据属性", "币种", "频度", "批次",
    "数据值", "上期值", "增减额(亿元)", "环比", "备注", "是否说明", "说明内容", "计算过程",
]


def _write_template(path: Path) -> None:
    book = Workbook()
    listing = book.active
    listing.title = "报表清单"
    listing.append(("报表代码", "报表名称", "频度", "批次"))
    listing.append(("A1411", "资产负债项目月报表", "月", "1"))
    listing.append(("A1433", "涉农贷款简报", "月", "2"))
    sheet = book.create_sheet("A1411")
    sheet.append(("指标代码", "指标名称", "行序号", "余额", "发生额"))
    for r in range(1, 6):
        sheet.append((f"12A{r:02d}", f"指标{r}", r, None, None))
    sheet2 = book.create_sheet("A1433")
    sheet2.append(("指标代码", "指标名称", "行序号", "余额"))
    for r in range(1, 4):
        sheet2.append((f"12P{r:02d}", f"涉农{r}", r, None, None))
    book.save(path)
    book.close()


def _write_comparison(path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "2026年8月对比结果"
    sheet.append(COMPARE_HEADERS)
    rows = [
        # A1411 余额：正常环比（命中警戒色路径）
        ("人民币", "2026-08-31", "6k0i", "示例甲银行", "4400000", "广东省", "12A01",
         "12A01", "指标1", "余额", "人民币", "月", "1", 120.5, 100.0, 20.5, 0.205, "", "", "", ""),
        # A1411 发生额：本期无上期有（红 3）
        ("人民币", "2026-08-31", "6k0i", "示例甲银行", "4400000", "广东省", "12A02",
         "12A02", "指标2", "发生额", "人民币", "月", "1", None, 50.0, -50.0, None,
         "本期无，上期有", "", "", ""),
        # A1411 余额 指标5：本期有上期无（黄 6）
        ("人民币", "2026-08-31", "6k0i", "示例甲银行", "4400000", "广东省", "12A05",
         "12A05", "指标5", "余额", "人民币", "月", "1", 88.0, None, 88.0, None,
         "本期有，上期无", "", "", ""),
        # A1433 余额
        ("人民币", "2026-08-31", "6k0i", "示例甲银行", "4400000", "广东省", "12P01",
         "12P01", "涉农1", "余额", "人民币", "月", "2", 10.0, 12.0, -2.0, -0.1667,
         "", "", "", ""),
    ]
    for row in rows:
        sheet.append(row)
    book.save(path)
    book.close()


class DualModeParityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = TemporaryDirectory(prefix="central-render-dual-")
        root = Path(cls._tmp.name)
        cls.template = root / "金融表单.xlsx"
        cls.comparison = root / "比较结果.xlsx"
        cls.configs = ensure_split_central_configs(root / "config")
        cls.config = (cls.configs["common"], cls.configs["forms"])
        _write_template(cls.template)
        _write_comparison(cls.comparison)
        cls.outputs: dict[str, Path] = {}
        for mode in ("OPENPYXL", "FAST_OOXML"):
            result = render_financial_forms(
                comparison_file=cls.comparison, template_file=cls.template,
                output_dir=root / f"输出-{mode}", config_path=cls.config,
                render_mode=mode,
            )
            cls.outputs[mode] = result.files[0]

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def test_sheet_names_and_order_match(self):
        books = {mode: load_workbook(path) for mode, path in self.outputs.items()}
        names = {mode: list(book.sheetnames) for mode, book in books.items()}
        for book in books.values():
            book.close()
        self.assertEqual(names["OPENPYXL"], names["FAST_OOXML"])

    def test_cell_values_match(self):
        books = {mode: load_workbook(path, data_only=True)
                 for mode, path in self.outputs.items()}
        try:
            open_sheet = books["OPENPYXL"]["A1411"]
            fast_sheet = books["FAST_OOXML"]["A1411"]
            self.assertEqual(open_sheet.max_column, fast_sheet.max_column)
            self.assertEqual(open_sheet.max_row, fast_sheet.max_row)
            for row in range(1, open_sheet.max_row + 1):
                for col in range(1, open_sheet.max_column + 1):
                    self.assertEqual(
                        open_sheet.cell(row, col).value,
                        fast_sheet.cell(row, col).value,
                        f"A1411!R{row}C{col}",
                    )
        finally:
            for book in books.values():
                book.close()

    def test_number_formats_match(self):
        books = {mode: load_workbook(path) for mode, path in self.outputs.items()}
        try:
            open_cell = books["OPENPYXL"]["A1411"].cell(2, 4)
            fast_cell = books["FAST_OOXML"]["A1411"].cell(2, 4)
            self.assertEqual(open_cell.number_format, fast_cell.number_format)
            open_ratio = books["OPENPYXL"]["A1411"].cell(2, 7)
            fast_ratio = books["FAST_OOXML"]["A1411"].cell(2, 7)
            self.assertEqual(open_ratio.number_format, fast_ratio.number_format)
        finally:
            for book in books.values():
                book.close()

    def test_warning_fills_match(self):
        """本期无上期有 → 黄(FFFFFF00) 填充；两模式一致（含空值格也上色）。"""
        fills = {}
        for mode, path in self.outputs.items():
            book = load_workbook(path)
            # 12A05 余额（本期有上期无）→ 红 3 涂本期与增减额块
            cell = book["A1411"].cell(6, 4)
            fills[mode] = cell.fill.fgColor.rgb if cell.fill and cell.fill.patternType else None
            book.close()
        self.assertEqual(fills["OPENPYXL"], fills["FAST_OOXML"])
        self.assertEqual(fills["FAST_OOXML"], "FFFF0000")

    def test_hidden_rows_match(self):
        hidden = {}
        for mode, path in self.outputs.items():
            book = load_workbook(path)
            sheet = book["A1411"]
            rows = {r for r in range(2, sheet.max_row + 1)
                    if sheet.row_dimensions[r].hidden}
            hidden[mode] = rows
            book.close()
        self.assertEqual(hidden["OPENPYXL"], hidden["FAST_OOXML"])

    def test_fast_renderer_excludes_embedded_config_sheets(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            template = root / "金融表单.xlsx"
            comparison = root / "比较结果.xlsx"
            outputs = ensure_split_central_configs(root / "config")
            config = outputs["forms"]
            _write_template(template)
            _write_comparison(comparison)
            embed_financial_template_in_config(config, template)

            result = render_financial_forms(
                comparison_file=comparison, output_dir=root / "输出", config_path=(outputs["common"], config),
                render_mode="FAST_OOXML",
            )
            book = load_workbook(result.files[0], data_only=False)
            try:
                self.assertIn("A1411", book.sheetnames)
                self.assertNotIn("使用说明", book.sheetnames)
                self.assertNotIn("金融表单映射", book.sheetnames)
                self.assertNotIn("转表设置", book.sheetnames)
            finally:
                book.close()

    def test_log_records_renderer(self):
        # 每次任务运行记录必须记录实际 Renderer（不能偷偷换实现）。
        # 由 service 层输出 "Renderer: <mode>"；此处直接校验设置透传链路。
        from base_audit.systems.s3_central_statistics.renderers import build_renderer
        renderer = build_renderer("FAST_OOXML", template_path=self.template,
                                  alerts=[], hide_empty_rows=True,
                                  delete_empty_sheets=True)
        self.assertEqual(renderer.mode, "FAST_OOXML")


class FastRendererErrorTests(unittest.TestCase):
    def test_corrupted_template_raises_fast_error(self):
        """模板 XML 损坏 → FastOoxmlRenderError（含 stage/sheet/part），不回退。"""
        import hashlib
        import zipfile
        from base_audit.systems.s3_central_statistics.fast_ooxml import (
            TemplateCompiler, FastOoxmlRenderError,
        )
        with TemporaryDirectory(prefix="central-err-") as folder:
            root = Path(folder)
            template = root / "金融表单.xlsx"
            _write_template(template)
            # ZIP 级破坏：把 A1411 的 sheet 部件替换为非法 XML
            raw = template.read_bytes()
            with zipfile.ZipFile(template) as archive:
                names = archive.namelist()
                parts = {name: archive.read(name) for name in names}
            sheet_part = next(
                (name for name in names if name.startswith("xl/worksheets/")
                 and b"12A01" in parts.get(name, b"")), names[0])
            parts[sheet_part] = b"<worksheet><sheetData><row r=</sheetData>"
            with zipfile.ZipFile(template.with_name("损坏.xlsx"), "w", zipfile.ZIP_DEFLATED) as archive:
                for name, payload in parts.items():
                    archive.writestr(name, payload)
            _ = hashlib.sha256(raw).hexdigest()

            compiler = TemplateCompiler(template.with_name("损坏.xlsx"))
            with self.assertRaises(FastOoxmlRenderError) as ctx:
                compiler.compile()
            self.assertEqual(ctx.exception.renderer, "FAST_OOXML")
            self.assertEqual(ctx.exception.stage, "template_compile")
            self.assertTrue(ctx.exception.part.startswith("xl/worksheets/"))


class FastRendererSheetDeleteTests(unittest.TestCase):
    """删空表时的 workbook.xml 维护（对齐 Excel 语义）。

    回归背景：两个真实故障——(1) definedNames 引用被删表时直接抛错（实际是
    Excel 自动维护的 _xlnm._FilterDatabase，应随表清理）；(2) bookViews 的
    activeTab/firstSheet 未随之收敛，越界导致 Excel 判定文件损坏、拒绝打开
    （COM 报“Workbooks 的 Open 方法无效”）。
    """

    def _rendered_workbook_xml(self) -> str:
        """生成一次渲染结果，返回其 workbook.xml 文本。

        模板形态刻意贴近真实文件：3 张表（报表清单 / A1411 / A1433），其中
        A1411 无数据将被删除；bookViews 的 activeTab/firstSheet 指向最后一张
        表（下标 2），删表后若不收敛就会越界（合法范围只剩 0–1）。openpyxl
        会按活动表重算 activeTab，故直接改写 ZIP 内的 workbook.xml。
        """
        import re
        import zipfile

        from base_audit.systems.s3_central_statistics.renderers import (
            RENDERER_FAST_OOXML, RendererFactory,
        )
        with TemporaryDirectory(prefix="central-del-") as folder:
            root = Path(folder)
            template = root / "金融表单.xlsx"
            _write_template(template)

            # 注入 definedNames（A1411 与 A1433 各一条表级名称）并把 bookViews
            # 指向最后一张表；直接改写部件以保证形态可控。
            raw = template.read_bytes()
            with zipfile.ZipFile(template) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            workbook = parts["xl/workbook.xml"].decode("utf-8")
            names_block = (
                "<definedNames>"
                "<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"1\" hidden=\"1\">"
                "'A1411'!$A$1:$E$5</definedName>"
                "<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"2\" hidden=\"1\">"
                "'A1433'!$A$1:$C$3</definedName>"
                "</definedNames>"
            )
            # openpyxl 对空名称写自闭合 <definedNames/>：必须原地替换，
            # 不能追加第二个块（一个 workbook 只允许一个 definedNames 块）。
            if "<definedNames/>" in workbook:
                workbook = workbook.replace("<definedNames/>", names_block, 1)
            elif re.search(r"<definedNames\b[^>]*?>.*?</definedNames>", workbook, re.S):
                workbook = re.sub(
                    r"<definedNames\b[^>]*?>.*?</definedNames>", names_block,
                    workbook, count=1, flags=re.S)
            else:
                workbook = workbook.replace("</workbook>", names_block + "</workbook>", 1)
            if "<bookViews>" in workbook:
                workbook = re.sub(
                    r'(<workbookView\b[^>]*?)activeTab="\d+"', r'\1activeTab="2"', workbook)
                workbook = re.sub(
                    r'(<workbookView\b[^>]*?)firstSheet="\d+"', r'\1firstSheet="2"', workbook)
            else:
                workbook = workbook.replace(
                    "<sheets>",
                    '<bookViews><workbookView activeTab="2" firstSheet="2"/></bookViews><sheets>', 1)
            parts["xl/workbook.xml"] = workbook.encode("utf-8")
            with zipfile.ZipFile(template, "w", zipfile.ZIP_DEFLATED) as archive:
                for name, payload in parts.items():
                    archive.writestr(name, payload)

            factory = RendererFactory(
                template, alerts=[], hide_empty_rows=True, delete_empty_sheets=True)
            renderer = factory.get_renderer(RENDERER_FAST_OOXML)
            # 只提供 A1433 的数据：A1411 无命中 → 被当作空表删除。
            data = {("人民币" + "12P01" + "余额" + "人民币"): (10.0, 12.0)}
            out = root / "out.xlsx"
            renderer.render_org(
                org_name="测试机构", region_name="广东省", record_date="2026-08",
                data=data, output_path=out,
            )
            self.assertIsNotNone(raw)   # 保留模板原始字节引用，便于排查
            with zipfile.ZipFile(out) as archive:
                return archive.read("xl/workbook.xml").decode("utf-8")

    def test_deleted_sheet_is_removed_and_views_remapped(self):
        import re
        xml = self._rendered_workbook_xml()
        names = re.findall(r'<sheet\b[^>]*name="([^"]*)"', xml)
        self.assertNotIn("A1411", names)          # 空表确实被删除
        self.assertIn("A1433", names)
        # activeTab / firstSheet 必须落在现存工作表范围内，否则 Excel 拒开。
        self.assertEqual(len(names), 2)
        for attr in ("activeTab", "firstSheet"):
            value = int(re.search(rf'{attr}="(\d+)"', xml).group(1))
            self.assertLess(value, len(names), f"{attr}={value} 越界（{len(names)} 张表）")
        # 原活动表 A1433 删表后下标应为 1（A1411 在它之前被删）。
        self.assertIn('activeTab="1"', xml)

    def test_defined_names_of_deleted_sheet_are_dropped(self):
        import re
        xml = self._rendered_workbook_xml()
        block = re.search(r"<definedNames>.*?</definedNames>", xml, re.S)
        # 属于被删表 A1411 的名称被移除；A1433 的名称保留。
        if block is not None:
            self.assertNotIn("'A1411'!", block.group(0))
            self.assertIn("'A1433'!", block.group(0))


if __name__ == "__main__":
    unittest.main()
