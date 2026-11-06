from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "core" / "src"))

from base_audit.systems.s3_central_statistics.explanation_exporter import (
    export_selected_explanations,
    import_explanation_feedback,
)


class ExplanationExporterTests(unittest.TestCase):
    def test_exports_only_rows_marked_for_explanation(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "比较结果.xlsx"
            output = Path(folder) / "指标说明.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "2026年8月对比结果"
            sheet.append(("机构类代码", "机构类名称", "地区代码", "地区名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "输出说明", "说明内容"))
            sheet.append(("6k0i", "测试机构", "4400000", "广东省", "A001", "指标一", 2, 1, 1, "是", "本期增加"))
            sheet.append(("6k0i", "测试机构", "4400000", "广东省", "A002", "指标二", 3, 2, 1, "否", "不导出"))
            book.save(source)
            book.close()

            self.assertEqual(export_selected_explanations(source, output), 1)
            result = load_workbook(output, data_only=True)
            explanation = result["(说明)2026.08 说明文件"]
            self.assertEqual(explanation.max_row, 3)
            self.assertEqual(explanation["A1"].value, "xxx说明")
            self.assertIn("A1:H1", {str(area) for area in explanation.merged_cells.ranges})
            self.assertEqual(explanation.max_column, 8)
            self.assertEqual(explanation.cell(3, 3).value, "A001")
            self.assertEqual(explanation.cell(3, 8).value, "本期增加")
            # E/F/G 显示两位小数（显示格式，单元格数值不四舍五入）。
            self.assertEqual(explanation.cell(3, 5).number_format, "0.00")
            self.assertEqual(explanation.cell(3, 7).number_format, "0.00")
            self.assertEqual(result.sheetnames, ["(说明)2026.08 说明文件"])
            result.close()

    def test_uses_precise_then_wildcard_org_reference_without_adding_region_suffix(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "比较结果.xlsx"
            config = Path(folder) / "3.0大集中通用配置.xlsx"
            output = Path(folder) / "指标说明.xlsx"

            book = Workbook()
            sheet = book.active
            sheet.title = "比较结果"
            sheet.append(("机构类代码", "机构类名称", "地区代码", "地区名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "输出说明", "说明内容"))
            sheet.append(("7020", "村镇银行", "33445", "广东省", "A001", "指标一", 2, 1, 1, "是", ""))
            sheet.append(("6k0i", "示例甲银行", "4400000", "广东省", "A002", "指标二", 3, 2, 1, "是", ""))
            book.save(source)
            book.close()

            book = Workbook()
            sheet = book.active
            sheet.title = "机构地区参照"
            sheet.append(("机构类代码", "地区代码", "机构名称"))
            sheet.append(("7020", "33445", "示例乙银行"))
            # 唯一名称会自动视为该机构类通用映射，无需人为复制全部地区代码。
            sheet.append(("6k0i", "*", "示例甲银行"))
            book.save(config)
            book.close()

            self.assertEqual(export_selected_explanations(source, output, config_path=config), 2)
            result = load_workbook(output, data_only=True)
            explanation = result["说明文件"]
            names = {
                explanation.cell(row, 1).value: explanation.cell(row, 2).value
                for row in range(3, explanation.max_row + 1)
                if explanation.cell(row, 1).value is not None
            }
            self.assertEqual(names, {"7020": "示例乙银行", "6k0i": "示例甲银行"})
            result.close()

    def test_imports_feedback_by_org_indicator_current_and_change_with_name_check(self):
        with tempfile.TemporaryDirectory() as folder:
            target_path = Path(folder) / "指标说明.xlsx"
            feedback_path = Path(folder) / "机构反馈.xlsx"
            for path, reason in ((target_path, "待机构填写"), (feedback_path, "本期业务正常增加")):
                book = Workbook()
                sheet = book.active
                sheet.title = "说明文件"
                sheet.merge_cells("A1:H1")
                sheet["A1"] = "xx说明"
                sheet.append(("机构代码", "机构名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "变动原因"))
                previous = 1.0 if path == target_path else 999.0
                current = 2.0 if path == target_path else 2.0000000000005
                change = 1.0 if path == target_path else 1.0000000000005
                name = "测试机构" if path == target_path else "测 试机构"
                sheet.append(("6k0i", name, "A001", "指标一", current, previous, change, reason))
                book.save(path)
                book.close()

            result = import_explanation_feedback(target_path, [feedback_path])
            self.assertEqual((result.imported, result.unmatched, result.conflicts), (1, 0, 0))
            self.assertTrue(result.backup_path.is_file())
            target = load_workbook(target_path, data_only=True)
            self.assertEqual(target["说明文件"].cell(3, 8).value, "本期业务正常增加")
            target.close()

    def test_keeps_same_named_indicators_separate_by_indicator_code(self):
        with tempfile.TemporaryDirectory() as folder:
            target_path = Path(folder) / "指标说明.xlsx"
            feedback_path = Path(folder) / "机构反馈.xlsx"
            for path, reasons in (
                (target_path, ("待填写一", "待填写二")),
                (feedback_path, ("原因一", "原因二")),
            ):
                book = Workbook()
                sheet = book.active
                sheet.title = "说明文件"
                sheet.merge_cells("A1:H1")
                sheet["A1"] = "xx说明"
                sheet.append(("机构代码", "机构名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "变动原因"))
                sheet.append(("6k0p", "示例乙银行（广东省）", "12P1Z", "农产品加工贷款", 0.05024, 0.05046, -0.00022, reasons[0]))
                sheet.append(("6k0p", "示例乙银行（广东省）", "12P28", "农产品加工贷款", 0.05024, 0.05046, -0.00022, reasons[1]))
                book.save(path)
                book.close()

            result = import_explanation_feedback(target_path, [feedback_path])
            self.assertEqual((result.imported, result.unmatched, result.conflicts), (2, 0, 0))
            target = load_workbook(target_path, data_only=True)
            self.assertEqual(target["说明文件"].cell(3, 8).value, "原因一")
            self.assertEqual(target["说明文件"].cell(4, 8).value, "原因二")
            target.close()

    def test_writes_conflict_report_with_feedback_source_and_reason(self):
        with tempfile.TemporaryDirectory() as folder:
            target_path = Path(folder) / "指标说明.xlsx"
            feedback_a = Path(folder) / "机构反馈A.xlsx"
            feedback_b = Path(folder) / "机构反馈B.xlsx"
            for path, reason in (
                (target_path, "待机构填写"),
                (feedback_a, "业务原因甲"),
                (feedback_b, "业务原因乙"),
            ):
                book = Workbook()
                sheet = book.active
                sheet.title = "说明文件"
                sheet.merge_cells("A1:H1")
                sheet["A1"] = "xx说明"
                sheet.append(("机构代码", "机构名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "变动原因"))
                sheet.append(("6k0i", "测试机构", "A001", "指标一", 2.0, 1.0, 1.0, reason))
                book.save(path)
                book.close()

            result = import_explanation_feedback(target_path, [feedback_a, feedback_b])
            self.assertEqual(result.conflicts, 1)
            self.assertIsNotNone(result.conflict_report_path)
            self.assertTrue(result.conflict_report_path.is_file())
            report = load_workbook(result.conflict_report_path, data_only=True)
            row = list(report["冲突明细"].iter_rows(min_row=2, values_only=True))[0]
            self.assertEqual(row[0], "冲突")
            self.assertIn("多个反馈", row[-1])
            self.assertIn("机构反馈A.xlsx", "\n".join(str(value or "") for value in row))
            self.assertIn("业务原因甲", "\n".join(str(value or "") for value in row))
            report.close()

    def test_monthly_export_keeps_original_institution_code_and_name(self):
        """月报按 VBA 的八列表头输出，不拼地区后缀或地区码。"""
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "比较结果.xlsx"
            output = Path(folder) / "指标说明.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "比较结果"
            sheet.append(("机构类代码", "机构类名称", "地区代码", "地区名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "输出说明", "说明内容"))
            sheet.append(("7020", "村镇银行", "4413231", "惠东县乡1", "12A38", "指标甲", 2, 1, 1, "是", ""))
            sheet.append(("6k0i", "广东惠州农商行", "4400000", "广东省", "A002", "指标乙", 3, 2, 1, "是", ""))
            book.save(source)
            book.close()

            self.assertEqual(export_selected_explanations(source, output), 2)
            result = load_workbook(output, data_only=True)
            explanation = result["说明文件"]
            self.assertEqual(
                tuple(explanation.cell(2, column).value for column in range(1, 9)),
                ("机构代码", "机构名称", "指标代码", "指标名称", "本期金额（亿元）", "上期金额（亿元）", "增减额（亿元）", "变动原因"),
            )
            self.assertEqual(explanation["A1"].value, "xxx说明")
            self.assertIn("A1:H1", {str(item) for item in explanation.merged_cells.ranges})
            self.assertIsNone(explanation.auto_filter.ref)
            self.assertIsNone(explanation.freeze_panes)
            self.assertEqual(explanation["A2"].number_format, "@")
            self.assertEqual(explanation["E3"].number_format, "0.00")
            self.assertEqual(explanation.column_dimensions["A"].width, 9.0)
            self.assertEqual(explanation.column_dimensions["B"].width, 24.875)
            values = {
                explanation.cell(row, 1).value: explanation.cell(row, 2).value
                for row in range(3, explanation.max_row + 1)
                if explanation.cell(row, 1).value is not None
            }
            self.assertEqual(values, {"7020": "村镇银行", "6k0i": "广东惠州农商行"})
            result.close()

    def test_daily_export_adds_region_columns_without_changing_institution_code(self):
        """日报在机构名称后新增地区列，地区不同的机构行保持可区分。"""
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "比较结果.xlsx"
            output = Path(folder) / "指标说明.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "比较结果"
            sheet.append(("机构类代码", "机构类名称", "地区代码", "地区名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "输出说明", "说明内容"))
            sheet.append(("6k0i", "广东惠州农商行", "4400000", "广东省", "A001", "指标甲", 5, 2, 3, "是", ""))
            sheet.append(("6k0i", "广东惠州农商行", "4413000", "惠州市", "A001", "指标甲", 5, 2, 3, "是", ""))
            sheet.append(("6k0u", "龙门农商行", "4400000", "广东省", "A001", "指标甲", 7, 2, 5, "是", ""))
            book.save(source)
            book.close()

            self.assertEqual(export_selected_explanations(source, output, format_mode="DAILY"), 3)
            result = load_workbook(output, data_only=True)
            explanation = result["说明文件"]
            self.assertEqual(
                tuple(explanation.cell(2, column).value for column in range(1, 11)),
                ("机构代码", "机构名称", "地区代码", "地区名称", "指标代码", "指标名称", "本期金额（亿元）", "上期金额（亿元）", "增减额（亿元）", "变动原因"),
            )
            self.assertEqual(
                [tuple(explanation.cell(row, column).value for column in range(1, 5)) for row in range(3, 6)],
                [
                    ("6k0i", "广东惠州农商行", "4400000", "广东省"),
                    ("6k0i", "广东惠州农商行", "4413000", "惠州市"),
                    ("6k0u", "龙门农商行", "4400000", "广东省"),
                ],
            )
            self.assertEqual(explanation.auto_filter.ref, "A2:J5")
            self.assertEqual({str(r) for r in explanation.merged_cells.ranges}, {"A1:J1"})
            result.close()

    def test_feedback_import_fills_down_merged_monthly_institution_cells(self):
        """VBA 月报纵向合并机构代码/名称后，F3 仍可逐指标回写说明。"""
        with tempfile.TemporaryDirectory() as folder:
            target_path = Path(folder) / "指标说明.xlsx"
            feedback_path = Path(folder) / "机构反馈.xlsx"
            for path, reasons in (
                (target_path, ("待机构填写", "待机构填写")),
                (feedback_path, ("指标一原因", "指标二原因")),
            ):
                book = Workbook()
                sheet = book.active
                sheet.title = "(说明)2026.10 说明文件"
                sheet.merge_cells("A1:H1")
                sheet["A1"] = "xxx说明"
                sheet.append(("机构代码", "机构名称", "指标代码", "指标名称", "本期金额（亿元）", "上期金额（亿元）", "增减额（亿元）", "变动原因"))
                sheet.append(("6k0i", "广东惠州农商行", "A001", "指标一", 2.0, 1.0, 1.0, reasons[0]))
                sheet.append((None, None, "A002", "指标二", 3.0, 2.0, 1.0, reasons[1]))
                sheet.merge_cells("A3:A4")
                sheet.merge_cells("B3:B4")
                book.save(path)
                book.close()

            result = import_explanation_feedback(target_path, [feedback_path])
            self.assertEqual((result.imported, result.unmatched, result.conflicts), (2, 0, 0))
            target = load_workbook(target_path, data_only=True)
            sheet = target["(说明)2026.10 说明文件"]
            self.assertEqual(sheet["H3"].value, "指标一原因")
            self.assertEqual(sheet["H4"].value, "指标二原因")
            target.close()

    def test_import_matches_change_with_excel_tolerance_and_ignores_previous_value(self):
        """机构代码大小写与名称空白规范化，金额按 1e-12 相对/绝对容差。"""
        with tempfile.TemporaryDirectory() as folder:
            target_path = Path(folder) / "指标说明.xlsx"
            feedback_path = Path(folder) / "机构反馈.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "说明文件"
            sheet.append(("机构代码", "机构名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "变动原因"))
            sheet.append(("6k0u", "龙门农村商业银行（广东省）", "12A18", "指标甲", 100.0, 50.0, 50.0, "待机构填写"))
            sheet.append((7020, "村镇银行（惠州市）", "12A22", "指标乙", 0.0, -1.0, 0.0, "待机构填写"))
            book.save(target_path)
            book.close()
            book = Workbook()
            sheet = book.active
            sheet.title = "说明文件"
            sheet.append(("机构代码", "机构名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "变动原因"))
            # 前一行金额刻意不同以验证上期值不参与；本期值和增减额仅有 1e-12 量级差异。
            sheet.append(("6K0u", "龙门 农村商业银行（广东省）", "12a18", "前缀_指标甲", "100.00000000005", "999", "50.000000000002", "精度容差原因"))
            sheet.append(("7020", "村镇银行（惠州市）", "12A22", "指标乙", 0.0, 888, "0.0000000000009", "绝对容差原因"))
            book.save(feedback_path)
            book.close()

            result = import_explanation_feedback(target_path, [feedback_path])
            self.assertEqual((result.imported, result.unmatched, result.conflicts), (2, 0, 0))
            target = load_workbook(target_path, data_only=True)
            # 本表头在第 1 行，数据从第 2 行起。
            self.assertEqual(target["说明文件"].cell(2, 8).value, "精度容差原因")
            self.assertEqual(target["说明文件"].cell(3, 8).value, "绝对容差原因")
            target.close()

    def test_institution_name_mismatch_blocks_feedback_and_reports_conflict(self):
        with tempfile.TemporaryDirectory() as folder:
            target_path = Path(folder) / "指标说明.xlsx"
            feedback_path = Path(folder) / "机构反馈.xlsx"
            for path, name, reason in (
                (target_path, "惠州农商银行", "待机构填写"),
                (feedback_path, "龙门农商银行", "反馈原因"),
            ):
                book = Workbook()
                sheet = book.active
                sheet.title = "说明文件"
                sheet.append(("机构代码", "机构名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "变动原因"))
                sheet.append(("6k0i", name, "A001", "指标一", 10.0, 8.0, 2.0, reason))
                book.save(path)
                book.close()

            result = import_explanation_feedback(target_path, [feedback_path])
            self.assertEqual((result.imported, result.unmatched, result.conflicts), (0, 0, 1))
            self.assertIsNotNone(result.conflict_report_path)
            report = load_workbook(result.conflict_report_path, data_only=True)
            row = list(report["冲突明细"].iter_rows(min_row=2, values_only=True))[0]
            self.assertIn("机构名称与目标说明文件不一致", row[-1])
            self.assertIn("龙门农商银行", row[-1])
            report.close()
            target = load_workbook(target_path, data_only=True)
            self.assertEqual(target["说明文件"]["H2"].value, "待机构填写")
            target.close()

    def test_institution_name_in_key_disambiguates_fuzzy_numeric_candidates(self):
        with tempfile.TemporaryDirectory() as folder:
            target_path = Path(folder) / "指标说明.xlsx"
            feedback_path = Path(folder) / "机构反馈.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "说明文件"
            sheet.append(("机构代码", "机构名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "变动原因"))
            sheet.append(("6k0i", "候选甲", "A001", "指标一", 100.0, 90.0, 10.0, "待填写甲"))
            sheet.append(("6k0i", "候选乙", "A001", "指标一", 100.00000000005, 90.0, 10.0, "待填写乙"))
            book.save(target_path)
            book.close()
            book = Workbook()
            sheet = book.active
            sheet.title = "说明文件"
            sheet.append(("机构代码", "机构名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "变动原因"))
            sheet.append(("6k0i", "候选乙", "A001", "指标一", 100.000000000025, 1.0, 10.0, "应回写到乙"))
            book.save(feedback_path)
            book.close()

            result = import_explanation_feedback(target_path, [feedback_path])
            self.assertEqual((result.imported, result.unmatched, result.conflicts), (1, 0, 0))
            target = load_workbook(target_path, data_only=True)
            self.assertEqual(target["说明文件"]["H2"].value, "待填写甲")
            self.assertEqual(target["说明文件"]["H3"].value, "应回写到乙")
            target.close()

    def test_village_banks_sharing_code_7020_match_by_institution_name(self):
        """同一机构代码下多个村镇银行名称不同，反馈必须只回写到名称相同的机构。"""
        with tempfile.TemporaryDirectory() as folder:
            target_path = Path(folder) / "指标说明.xlsx"
            feedback_path = Path(folder) / "机构反馈.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "说明文件"
            sheet.append(("机构代码", "机构名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "变动原因"))
            sheet.append(("7020", "甲村镇银行", "A001", "指标一", 100.0, 90.0, 10.0, "甲待填写"))
            sheet.append(("7020", "乙村镇银行", "A001", "指标一", 100.0, 90.0, 10.0, "乙待填写"))
            book.save(target_path)
            book.close()

            book = Workbook()
            sheet = book.active
            sheet.title = "说明文件"
            sheet.append(("机构代码", "机构名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "变动原因"))
            sheet.append(("7020", "乙村镇银行", "A001", "指标一", 100.0, 80.0, 10.0, "乙银行反馈原因"))
            book.save(feedback_path)
            book.close()

            result = import_explanation_feedback(target_path, [feedback_path])
            self.assertEqual((result.imported, result.unmatched, result.conflicts), (1, 0, 0))
            target = load_workbook(target_path, data_only=True)
            self.assertEqual(target["说明文件"]["H2"].value, "甲待填写")
            self.assertEqual(target["说明文件"]["H3"].value, "乙银行反馈原因")
            target.close()

    def test_duplicate_feedback_candidates_are_reported_as_ambiguous(self):
        with tempfile.TemporaryDirectory() as folder:
            target_path = Path(folder) / "指标说明.xlsx"
            feedback_path = Path(folder) / "机构反馈.xlsx"
            for path, reason in ((target_path, "待填写"), (feedback_path, "不应静默回写")):
                book = Workbook()
                sheet = book.active
                sheet.title = "说明文件"
                sheet.append(("机构代码", "机构名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "变动原因"))
                sheet.append(("6k0i", "测试机构", "A001", "指标一", 20.0, 10.0, 10.0, reason))
                if path == target_path:
                    sheet.append(("6k0i", "测试机构", "A001", "指标一", 20.0, 10.0, 10.0, "仍待填写"))
                book.save(path)
                book.close()

            result = import_explanation_feedback(target_path, [feedback_path])
            self.assertEqual((result.imported, result.unmatched, result.conflicts), (0, 0, 1))
            report = load_workbook(result.conflict_report_path, data_only=True)
            row = list(report["冲突明细"].iter_rows(min_row=2, values_only=True))[0]
            self.assertIn("2 个目标", row[-1])
            self.assertEqual(row[7], "2、3")
            report.close()
            target = load_workbook(target_path, data_only=True)
            self.assertEqual((target["说明文件"]["H2"].value, target["说明文件"]["H3"].value), ("待填写", "仍待填写"))
            target.close()

    def test_feedback_without_previous_value_column_uses_current_and_change(self):
        with tempfile.TemporaryDirectory() as folder:
            target_path = Path(folder) / "指标说明.xlsx"
            feedback_path = Path(folder) / "机构反馈.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "说明文件"
            sheet.append(("机构代码", "机构名称", "指标代码", "指标名称", "数据值", "上期值", "增减额(亿元)", "变动原因"))
            sheet.append(("6k0i", "测试机构", "A001", "指标一", 12.0, 10.0, 2.0, "待填写"))
            book.save(target_path)
            book.close()
            book = Workbook()
            sheet = book.active
            sheet.title = "说明文件"
            sheet.append(("机构代码", "机构名称", "指标代码", "指标名称", "数据值", "增减额(亿元)", "变动原因"))
            sheet.append(("6k0i", "测试机构", "A001", "指标一", 12.0, 2.0, "本期净增加"))
            book.save(feedback_path)
            book.close()

            result = import_explanation_feedback(target_path, [feedback_path])
            self.assertEqual((result.imported, result.unmatched, result.conflicts), (1, 0, 0))
            target = load_workbook(target_path, data_only=True)
            self.assertEqual(target["说明文件"]["H2"].value, "本期净增加")
            target.close()
