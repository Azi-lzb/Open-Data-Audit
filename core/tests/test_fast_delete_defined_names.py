# -*- coding: utf-8 -*-
"""Fast/Direct OOXML 删表 × Sheet-local definedNames 安全处理回归。

设计基线（AGENTS「OOXML 结构与用户错误提示铁律」+ V3 删空表任务）：
- 工作表身份 = SheetName / rId / worksheet part；localSheetId 只是顺序位置；
- 删多表共享同一份 old_index → new_index 映射，禁止逐次 -1 的累计偏差；
- 被删表自己的表级名称（_xlnm._FilterDatabase / Print_Area / Print_Titles）
  随表移除；保留表 localSheetId 按新顺序重算，公式文本不改写；
- 无法证明安全的复杂 definedName 引用候选空表 → 该表降级隐藏（保守），
  引用「明确业务删除」的内嵌配置页 → 名称一并移除；
- 报表清单仍是 3.3 必需表；金融表单映射不是依赖；
- 输出必须通过结构校验、openpyxl 重开、FastOoxmlCompiler 再编译。
"""

from __future__ import annotations

import re
import sys
import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from base_audit.systems.s3_central_statistics.fast_ooxml import (  # noqa: E402
    FastOoxmlRenderer,
    TemplateCompiler,
)
from base_audit.systems.s3_central_statistics.ooxml_structure import (  # noqa: E402
    parse_defined_names,
    parse_workbook_sheets,
    validate_workbook_relationships,
)

CODES = ("A1411", "A1412", "A1413", "A1414", "A1415")


def _build_template(path: Path, *, codes: tuple[str, ...] = CODES,
                    filters: set[str] = frozenset(),
                    print_areas: set[str] = frozenset(),
                    print_titles: set[str] = frozenset()) -> None:
    """构造协议最小可用的金融表单模板：报表清单 + 每表一行唯一指标。"""
    book = Workbook()
    listing = book.active
    listing.title = "报表清单"
    listing.append(("报表代码", "报表名称", "频度", "批次"))
    for code in codes:
        listing.append((code, f"表{code}", "月", "1"))
    for code in codes:
        sheet = book.create_sheet(code)
        sheet.append(("指标代码", "指标名称", "行序号", "余额", "发生额"))
        sheet.append((f"{code}X01", "指标", 1, None, None))
        if code in filters:
            sheet.auto_filter.ref = "A1:E2"
        if code in print_areas:
            sheet.print_area = "A1:E5"
        if code in print_titles:
            sheet.print_title_rows = "1:1"
    book.save(path)
    book.close()


def _data_for(codes) -> dict[str, tuple]:
    """让给定报表命中的回填数据（只有这些表非空，其余为空表候选）。"""
    return {
        f"人民币{code}X01余额人民币": (10.0, 12.0) for code in codes
    }


def _rewrite_workbook_xml(path: Path, transform) -> None:
    """对模板 ZIP 内的 workbook.xml 做原文变换（测试注入用）。"""
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["xl/workbook.xml"] = transform(parts["xl/workbook.xml"].decode("utf-8")).encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)


def _render(root: Path, template: Path, codes, *, delete_empty_sheets: bool = True):
    """渲染一次并返回 (结果, 输出路径, 输出 workbook.xml 文本)。"""
    renderer = FastOoxmlRenderer(
        template, alerts=[], hide_empty_rows=True,
        delete_empty_sheets=delete_empty_sheets)
    output = root / "out.xlsx"
    result = renderer.render_org(
        org_name="测试机构", region_name="广东省", record_date="2026-08",
        data=_data_for(codes), output_path=output)
    with zipfile.ZipFile(output) as archive:
        xml = archive.read("xl/workbook.xml").decode("utf-8")
        names = set(archive.namelist())
        rels = archive.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    # 结构完整性（每次渲染后都检查，覆盖 §12.1/2/3）。
    validate_workbook_relationships(xml, rels, names)
    return result, output, xml


def _sheet_names(xml: str) -> list[str]:
    return [s["name"] for s in parse_workbook_sheets(xml)]


def _name_by_text(xml: str, sheet_name: str):
    return next((n for n in parse_defined_names(xml) if sheet_name in n.text), None)


class DefinedNameDeleteTests(unittest.TestCase):
    """测试 1–5：删表与 localSheetId 重映射（含多表同时删除）。"""

    def test_1_deleted_sheet_own_filterdatabase_is_dropped(self) -> None:
        """删除的 Sheet 自己有 _FilterDatabase → 名称随表一起删除。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            _build_template(template, filters={"A1411"})
            _result, output, xml = _render(root, template, codes=set(CODES) - {"A1411"})
            self.assertNotIn("A1411", _sheet_names(xml))
            self.assertIsNone(_name_by_text(xml, "A1411"))
            self.assertEqual(len(parse_defined_names(xml)), 0)
            book = load_workbook(output)
            self.assertNotIn("A1411", book.sheetnames)
            book.close()

    def test_2_survivor_local_sheet_id_remapped_after_front_delete(self) -> None:
        """前面的表删除 → 保留表 _FilterDatabase localSheetId 重映射，文本不变。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            _build_template(template, filters={"A1413"})
            _result, output, xml = _render(root, template, codes=set(CODES) - {"A1411"})
            # 老下标：报表清单=0，A1411=1 … A1413=3；删 A1411 后 A1413 → 2。
            item = _name_by_text(xml, "A1413")
            self.assertIsNotNone(item)
            self.assertEqual(item.local_sheet_id, 2)
            self.assertEqual(item.text, "'A1413'!$A$1:$E$2")
            book = load_workbook(output)
            self.assertEqual(book.sheetnames, ["报表清单", "A1412", "A1413", "A1414", "A1415"])
            book.close()

    def test_3_multi_delete_all_local_sheet_ids_follow_new_order(self) -> None:
        """一次删除多张表 → 所有保留表的 localSheetId 与新顺序一致。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            _build_template(template, filters={"A1413", "A1415"})
            _result, output, xml = _render(
                root, template, codes=set(CODES) - {"A1412", "A1414"})
            self.assertEqual(_sheet_names(xml), ["报表清单", "A1411", "A1413", "A1415"])
            # A1413：老 localSheetId=3（老序 0清单,1A,2B,3C,4D,5E）→ 新序 2。
            # A1415：老 5 → 新 3。
            self.assertEqual(_name_by_text(xml, "A1413").local_sheet_id, 2)
            self.assertEqual(_name_by_text(xml, "A1415").local_sheet_id, 3)
            # 旧序 1..5 全部不应残留。
            for stale in (1, 4, 5):
                self.assertNotIn(
                    stale, [n.local_sheet_id for n in parse_defined_names(xml)])

    def test_4_multiple_deleted_sheets_filters_all_removed(self) -> None:
        """多张被删表各自的 _FilterDatabase 全部删除，不得残留。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            _build_template(template, filters={"A1412", "A1414"})
            _result, output, xml = _render(
                root, template, codes=set(CODES) - {"A1412", "A1414"})
            self.assertNotIn("A1412", _sheet_names(xml))
            self.assertNotIn("A1414", _sheet_names(xml))
            self.assertEqual(len(parse_defined_names(xml)), 0)
            book = load_workbook(output)
            self.assertNotIn("A1412", book.sheetnames)
            self.assertNotIn("A1414", book.sheetnames)
            book.close()

    def test_5_multiple_survivor_filters_all_remapped(self) -> None:
        """删除前方表后，所有保留表的 _FilterDatabase 全部正确 remap。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            _build_template(template, filters={"A1411", "A1413", "A1415"})
            _result, output, xml = _render(
                root, template, codes=set(CODES) - {"A1412", "A1414"})
            self.assertEqual(_name_by_text(xml, "A1411").local_sheet_id, 1)
            self.assertEqual(_name_by_text(xml, "A1413").local_sheet_id, 2)
            self.assertEqual(_name_by_text(xml, "A1415").local_sheet_id, 3)
            # 输出可被 FastOoxmlCompiler 再次编译（测试 12）。
            compiled = TemplateCompiler(output).compile()
            self.assertTrue(any(e["name"] == "报表清单" for e in compiled.sheet_registry))


class NamespaceAndPrintNamesTests(unittest.TestCase):
    """测试 6/7：namespace 形态与 Print_Area / Print_Titles。"""

    def test_6_ns_prefixed_workbook_defined_names_handled(self) -> None:
        """workbook.xml 带 ns0: 前缀（openpyxl 重存形态）时同样正确处理。"""

        def add_prefix(xml: str) -> str:
            uri = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            # 默认 xmlns 改绑为 ns0 前缀声明，再把所有标签加上 ns0: 前缀。
            xml = xml.replace(f'xmlns="{uri}"', f'xmlns:ns0="{uri}"', 1)
            xml = re.sub(r"<(?!/)(?![?!])", "<ns0:", xml)
            return xml.replace("</", "</ns0:")

        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            _build_template(template, filters={"A1413"})
            _rewrite_workbook_xml(template, add_prefix)
            with zipfile.ZipFile(template) as archive:
                self.assertIn("ns0:definedName",
                              archive.read("xl/workbook.xml").decode("utf-8"))
            _result, output, xml = _render(root, template, codes=set(CODES) - {"A1411"})
            item = _name_by_text(xml, "A1413")
            self.assertIsNotNone(item)
            self.assertEqual(item.local_sheet_id, 2)
            book = load_workbook(output)
            self.assertNotIn("A1411", book.sheetnames)
            book.close()

    def test_7_print_area_and_titles_follow_owner_sheet(self) -> None:
        """Print_Area/Print_Titles：owner 被删→随之删除；owner 保留→remap。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            # (a) owner 被删：名称一并消失。
            template = root / "a.xlsx"
            _build_template(template, print_areas={"A1412"}, print_titles={"A1412"})
            _result, output, xml = _render(root, template, codes=set(CODES) - {"A1412"})
            self.assertNotIn("A1412", _sheet_names(xml))
            self.assertEqual(len(parse_defined_names(xml)), 0)
            # (b) owner 保留：localSheetId remap，文本不变。
            template = root / "b.xlsx"
            _build_template(template, print_areas={"A1413"}, print_titles={"A1413"})
            _result, output, xml = _render(root, template, codes=set(CODES) - {"A1411"})
            area = next(n for n in parse_defined_names(xml)
                        if n.name.endswith("Print_Area"))
            titles = next(n for n in parse_defined_names(xml)
                          if n.name.endswith("Print_Titles"))
            self.assertEqual(area.local_sheet_id, 2)
            self.assertEqual(titles.local_sheet_id, 2)
            self.assertEqual(area.text, "'A1413'!$A$1:$E$5")
            book = load_workbook(output)
            book.close()


class DefinedNameFormulaDependencyTests(unittest.TestCase):
    """删 definedName 前必须确认保留公式没有仍以名称引用（防 #NAME?）。

    任务语义：无人使用的名称随表安全删除；仍被保留公式以名称引用的名称
    → 候选表降级隐藏（自动场景）/ 阻止硬删（明确业务删除场景）；
    绝不自动重写业务公式。
    """

    def _template_with_names(
            self, path: Path, names: dict[str, str], formulas: dict[str, str] | None = None,
            **build_kwargs) -> None:
        """构造模板 + 注入 definedNames + 在指定表 B2 写公式。"""
        _build_template(path, **build_kwargs)
        if formulas:
            book = load_workbook(path)
            for sheet_name, formula in formulas.items():
                book[sheet_name]["B2"] = formula
            book.save(path)
            book.close()

        def inject(xml: str) -> str:
            entries = "".join(
                f"<definedName name=\"{name}\">{ref}</definedName>"
                for name, ref in names.items())
            block = f"<definedNames>{entries}</definedNames>"
            if "<definedNames/>" in xml:
                return xml.replace("<definedNames/>", block, 1)
            return xml.replace("</workbook>", block + "</workbook>", 1)

        _rewrite_workbook_xml(path, inject)

    def test_1_unused_defined_name_is_dropped_with_sheet(self) -> None:
        """名称无人使用：删除旧配置页 + 该名称，文件正常，无 #NAME? 风险。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            self._template_with_names(
                template, names={"MyOldConfig": "'A1412'!$A$1"})
            _result, output, xml = _render(root, template, codes=set(CODES) - {"A1412"})
            self.assertNotIn("A1412", _sheet_names(xml))
            self.assertEqual(len(parse_defined_names(xml)), 0)   # 名称一并删除
            book = load_workbook(output)
            self.assertNotIn("A1412", book.sheetnames)
            book.close()

    def test_2_referenced_defined_name_degrades_to_hidden(self) -> None:
        """保留表公式仍以名称引用：候选空表降级隐藏，名称与公式原样保留。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            self._template_with_names(
                template, names={"MyOldConfig": "'A1412'!$A$1"},
                formulas={"A1411": "=MyOldConfig*100"})
            result, output, xml = _render(
                root, template, codes=set(CODES) - {"A1412"})
            # 旧配置页未物理删除，而是隐藏。
            self.assertIn("A1412", _sheet_names(xml))
            book = load_workbook(output)
            self.assertEqual(book["A1412"].sheet_state, "hidden")
            self.assertEqual(
                book["A1411"]["B2"].value, "=MyOldConfig*100")   # 公式保持不变
            book.close()
            # 名称保留且指向未变。
            item = _name_by_text(xml, "A1412")
            self.assertIsNotNone(item)
            self.assertEqual(item.name, "MyOldConfig")
            # 用户提示说人话（无 definedName/localSheetId/token/rId 术语）。
            joined = "".join(result.messages)
            self.assertIn("已改为隐藏", joined)
            for term in ("definedName", "localSheetId", "token", "rId"):
                self.assertNotIn(term, joined)
            # 技术日志记录 HIDE_INSTEAD_OF_DELETE。
            actions = result.stats.get("defined_name_actions", [])
            self.assertTrue(any(
                a.get("action") == "HIDE_INSTEAD_OF_DELETE"
                and a.get("reason") == "DEFINED_NAME_STILL_REFERENCED"
                and a.get("defined_name") == "MyOldConfig"
                for a in actions))

    def test_3_hard_delete_blocked_when_defined_name_referenced(self) -> None:
        """明确业务硬删除遇名称依赖：阻止硬删、明确报错、不产 #NAME? 文件。"""
        from base_audit.systems.s3_central_statistics.renderers import (
            FastOoxmlRenderError, UnsupportedFastRenderOperation)

        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            self._template_with_names(
                template, codes=("A1411", "A1412"),
                names={"MyConfig": "'转表设置'!$A$1"},
                formulas={"A1411": "=MyConfig*2"})
            book = load_workbook(template)
            book.create_sheet("使用说明")
            book.create_sheet("转表设置")
            book.save(template)
            book.close()
            output = root / "out.xlsx"
            renderer = FastOoxmlRenderer(
                template, alerts=[], hide_empty_rows=True, delete_empty_sheets=True)
            with self.assertRaises(UnsupportedFastRenderOperation) as ctx:
                renderer.render_org(
                    org_name="测试机构", region_name="广东省", record_date="2026-08",
                    data=_data_for({"A1411"}), output_path=output)
            message = str(ctx.exception)
            self.assertIn("无法删除工作表", message)
            self.assertIn("仍通过工作簿名称引用", message)
            # 用户消息不含底层术语；不落一个 #NAME? 文件。
            for term in ("definedName", "localSheetId", "token", "rId"):
                self.assertNotIn(term, ctx.exception.message)
            self.assertFalse(output.exists())

    def test_4_name_reference_matching_is_case_insensitive(self) -> None:
        """=myconfigvalue*2 必须识别为对 MyConfigValue 的引用。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            _build_template(template)
            book = load_workbook(template)
            book["A1411"]["B2"] = "=myconfigvalue*2"
            book.save(template)
            book.close()

            # owner-local 名称：只有公式名称引用检查会命中（文本引用仅覆盖
            # workbook-level 名称指向候选表的分支），专测 token 大小写语义。
            def inject_local(xml: str) -> str:
                block = (
                    "<definedNames>"
                    "<definedName name=\"MyConfigValue\" localSheetId=\"2\">"
                    "'A1412'!$A$1</definedName>"
                    "</definedNames>")
                if "<definedNames/>" in xml:
                    return xml.replace("<definedNames/>", block, 1)
                return xml.replace("</workbook>", block + "</workbook>", 1)

            _rewrite_workbook_xml(template, inject_local)
            result, output, xml = _render(
                root, template, codes=set(CODES) - {"A1412"})
            # 名称被引用（大小写不敏感）→ A1412 救回隐藏，名称保留。
            self.assertIn("A1412", _sheet_names(xml))
            book = load_workbook(output)
            self.assertEqual(book["A1412"].sheet_state, "hidden")
            book.close()

    def test_5_substring_is_not_misdetected_as_reference(self) -> None:
        """=InterestRate*2 不得误判为引用 Rate：无人使用 → 正常删除。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            self._template_with_names(
                template, names={"Rate": "'A1412'!$A$1"},
                formulas={"A1411": "=InterestRate*2"})
            _result, output, xml = _render(
                root, template, codes=set(CODES) - {"A1412"})
            self.assertNotIn("A1412", _sheet_names(xml))
            self.assertEqual(len(parse_defined_names(xml)), 0)
            book = load_workbook(output)
            self.assertNotIn("A1412", book.sheetnames)
            book.close()

    def test_6_same_local_name_across_sheets_does_not_break_keepers(self) -> None:
        """同名 local 名称作用域不串：保留表引用自己作用域的名称不受影响。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            _build_template(template)
            book = load_workbook(template)
            # 保留表 A1411 自己作用域内也有同名 MyRange 并在公式中使用。
            book["A1411"]["B2"] = "=MyRange*100"
            book.save(template)
            book.close()

            def inject(xml: str) -> str:
                block = (
                    "<definedNames>"
                    "<definedName name=\"MyRange\" localSheetId=\"2\">"
                    "'A1412'!$A$1</definedName>"
                    "<definedName name=\"MyRange\" localSheetId=\"1\">"
                    "'A1411'!$A$1</definedName>"
                    "</definedNames>")
                if "<definedNames/>" in xml:
                    return xml.replace("<definedNames/>", block, 1)
                return xml.replace("</workbook>", block + "</workbook>", 1)

            _rewrite_workbook_xml(template, inject)
            # 保守策略：保留表公式出现同名 token，无法可靠区分作用域时
            # 不硬删 → A1412 降级隐藏（不错误硬删），两个同名名称都保留。
            _result, output, xml = _render(
                root, template, codes=set(CODES) - {"A1412"})
            self.assertIn("A1412", _sheet_names(xml))
            book = load_workbook(output)
            self.assertEqual(book["A1412"].sheet_state, "hidden")
            self.assertEqual(book["A1411"]["B2"].value, "=MyRange*100")
            book.close()
            names = parse_defined_names(xml)
            self.assertEqual(
                sorted((n.name, n.local_sheet_id) for n in names),
                [("MyRange", 1), ("MyRange", 2)])

    def test_7_partial_reference_rescues_whole_sheet(self) -> None:
        """多名称只有部分被引用：整表降级隐藏，不得只删未引用名称留半残。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            self._template_with_names(
                template,
                names={
                    "NameA": "'A1412'!$A$1",
                    "NameB": "'A1412'!$B$2",
                    "NameC": "'A1412'!$C$3",
                },
                formulas={"A1411": "=NameB*2"})
            _result, output, xml = _render(
                root, template, codes=set(CODES) - {"A1412"})
            # A1412 整体隐藏，三个名称全部保留（无半残状态）。
            self.assertIn("A1412", _sheet_names(xml))
            kept_names = {n.name for n in parse_defined_names(xml)}
            self.assertEqual(kept_names, {"NameA", "NameB", "NameC"})
            book = load_workbook(output)
            self.assertEqual(book["A1412"].sheet_state, "hidden")
            book.close()

    def test_unused_name_referencing_internal_config_sheet_is_dropped(self) -> None:
        """无人使用的名称引用内嵌配置页 → 名称一并移除，页面仍物理删除。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            _build_template(template, codes=("A1411", "A1412"))
            # 用 openpyxl 追加两张内嵌配置页（存在即触发 internal 删除路径）。
            book = load_workbook(template)
            book.create_sheet("使用说明")
            book.create_sheet("转表设置")
            book.save(template)
            book.close()

            def inject(xml: str) -> str:
                block = (
                    "<definedNames>"
                    "<definedName name=\"引用配置页\">'转表设置'!$A$1</definedName>"
                    "</definedNames>")
                if "<definedNames/>" in xml:
                    return xml.replace("<definedNames/>", block, 1)
                return xml.replace("</workbook>", block + "</workbook>", 1)

            _rewrite_workbook_xml(template, inject)
            # 配置页不是报表清单表单，编译即可见 registry 中存在。
            compiled = TemplateCompiler(template).compile()
            self.assertEqual(
                set(compiled.internal_config_sheets), {"使用说明", "转表设置"})
            _result, output, xml = _render(root, template, codes={"A1411", "A1412"})
            self.assertNotIn("转表设置", _sheet_names(xml))
            self.assertNotIn("使用说明", _sheet_names(xml))
            self.assertEqual(len(parse_defined_names(xml)), 0)   # 悬空名称已移除
            book = load_workbook(output)
            self.assertNotIn("转表设置", book.sheetnames)
            book.close()

    def test_8_filterdatabase_regression_unchanged(self) -> None:
        """_FilterDatabase/Print_Area/Print_Titles 原有行为不回归。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            _build_template(
                template, filters={"A1412"}, print_areas={"A1412"},
                print_titles={"A1412"})
            # 无公式引用这些自动名称 → 随表安全删除。
            _result, output, xml = _render(
                root, template, codes=set(CODES) - {"A1412"})
            self.assertNotIn("A1412", _sheet_names(xml))
            self.assertEqual(len(parse_defined_names(xml)), 0)
            book = load_workbook(output)
            self.assertNotIn("A1412", book.sheetnames)
            book.close()

    def test_9_real_33_contract_still_holds(self) -> None:
        """3.3 真实配置：报表清单必需、金融表单映射非依赖（不回归）。"""
        template = ROOT.parent / "config" / "3.3大集中指标比较拆分_配置.xlsx"
        if not template.is_file():
            self.skipTest("缺少 3.3 正式配置")
        compiled = TemplateCompiler(template).compile()
        self.assertTrue(any(e["name"] == "报表清单" for e in compiled.sheet_registry))
        self.assertNotIn(
            "金融表单映射", [e["name"] for e in compiled.sheet_registry])


class BusinessContractTests(unittest.TestCase):
    """测试 9/10：报表清单必需、金融表单映射非依赖——本轮修复不得改变。"""

    def test_9_missing_listing_sheet_fails_compile(self) -> None:
        """删除报表清单后 3.3 编译必须明确报缺，不得用其他表替代。"""
        from base_audit.systems.s3_central_statistics.renderers import FastOoxmlRenderError

        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            _build_template(template)
            trimmed = root / "no-listing.xlsx"
            book = load_workbook(template)
            del book["报表清单"]
            book.save(trimmed)
            book.close()
            with self.assertRaises(FastOoxmlRenderError) as ctx:
                TemplateCompiler(trimmed).compile()
            self.assertIn("报表清单", str(ctx.exception))

    def test_10_no_form_mapping_sheet_still_compiles_and_renders(self) -> None:
        """没有金融表单映射：编译与渲染正常（本轮未重新引入旧依赖）。"""
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            _build_template(template, filters={"A1411"})
            compiled = TemplateCompiler(template).compile()
            self.assertNotIn("金融表单映射", [e["name"] for e in compiled.sheet_registry])
            _result, output, xml = _render(root, template, codes=set(CODES) - {"A1411"})
            self.assertNotIn("A1411", _sheet_names(xml))
            book = load_workbook(output)
            self.assertTrue(book.sheetnames)
            book.close()


class ReopenAndRecompileTests(unittest.TestCase):
    """测试 11/12：输出可 openpyxl 重开、可再次进入 FastOoxmlCompiler。"""

    def test_output_reopens_and_recompiles_with_filters(self) -> None:
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            _build_template(template, filters={"A1411", "A1414"})
            _result, output, xml = _render(
                root, template, codes={"A1411", "A1413", "A1415"})
            # A1412/A1414 删除；A1414 的 filter 随删，A1411 的保留。
            self.assertEqual(_sheet_names(xml), ["报表清单", "A1411", "A1413", "A1415"])
            self.assertIsNone(_name_by_text(xml, "A1414"))
            self.assertEqual(_name_by_text(xml, "A1411").local_sheet_id, 1)
            # openpyxl 重开（测试 11）。
            book = load_workbook(output)
            self.assertEqual(book.sheetnames, ["报表清单", "A1411", "A1413", "A1415"])
            book.close()
            # 再次 FastOoxmlCompiler（测试 12）。
            compiled = TemplateCompiler(output).compile()
            self.assertIn("A1411", compiled.sheets)
            # 再渲染一轮仍正常（编译产物二次消费）。
            renderer2 = FastOoxmlRenderer(
                output, alerts=[], hide_empty_rows=True, delete_empty_sheets=True)
            output2 = root / "out2.xlsx"
            renderer2.render_org(
                org_name="测试机构", region_name="广东省", record_date="2026-08",
                data=_data_for({"A1411"}), output_path=output2)
            with zipfile.ZipFile(output2) as archive:
                xml2 = archive.read("xl/workbook.xml").decode("utf-8")
            self.assertEqual(_sheet_names(xml2), ["报表清单", "A1411"])
            self.assertEqual(_name_by_text(xml2, "A1411").local_sheet_id, 1)


class DeleteDisabledPathTests(unittest.TestCase):
    """delete_empty_sheets=False：普通 FAST 路径不触发删表与名称重写。"""

    def test_no_delete_keeps_sheets_and_names_untouched(self) -> None:
        with TemporaryDirectory(prefix="dn-") as folder:
            root = Path(folder)
            template = root / "t.xlsx"
            _build_template(template, filters={"A1411"})
            renderer = FastOoxmlRenderer(
                template, alerts=[], hide_empty_rows=True,
                delete_empty_sheets=False)
            output = root / "out.xlsx"
            renderer.render_org(
                org_name="测试机构", region_name="广东省", record_date="2026-08",
                data=_data_for({"A1413"}), output_path=output)
            with zipfile.ZipFile(output) as archive:
                xml = archive.read("xl/workbook.xml").decode("utf-8")
            self.assertEqual(len(_sheet_names(xml)), 6)
            item = _name_by_text(xml, "A1411")
            self.assertIsNotNone(item)
            self.assertEqual(item.local_sheet_id, 1)   # 原值不动


if __name__ == "__main__":
    unittest.main()
