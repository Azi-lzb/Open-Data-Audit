"""条件格式重构测试：Direct OOXML 规则读取、标准化模型、三求值器分离。"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "core" / "src"))

from base_audit.cf_reader import (
    NormalizedConditionalRule,
    OoxmlConditionalRuleReader,
)
from base_audit.conditional_evaluators import (
    EVALUATOR_COM,
    EVALUATOR_COM_DISPLAY,
    EVALUATOR_PYTHON,
    ComFormulaEvaluator,
    ConditionalFormatService,
    PythonConditionalRuleEvaluator,
    effective_evaluator_mode,
)


def _build_rule_workbook(path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "资产负债表"
    sheet["A1"] = "指标"
    sheet["B1"] = "数值"
    sheet["A2"] = "贷款"
    sheet["B2"] = 5
    sheet["B3"] = -1
    sheet["B4"] = 0
    sheet["B5"] = 3
    # cellIs：B2:B4 大于 0（B5 不在区域内）
    sheet.conditional_formatting.add(
        "B2:B4", CellIsRule(operator="greaterThan", formula=["0"], stopIfTrue=False))
    # expression：C 列引用 B 列（锚点 C2）
    sheet["C2"] = None
    sheet.conditional_formatting.add(
        "C2:C5", FormulaRule(formula=["$B2<0"], stopIfTrue=True))
    # colorScale：不支持类型
    from openpyxl.formatting.rule import ColorScaleRule

    sheet.conditional_formatting.add("D2:D4", ColorScaleRule(
        start_type="num", start_value=0, start_color="FF0000",
        end_type="num", end_value=10, end_color="00FF00"))
    book.save(path)
    book.close()


class ReaderTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = TemporaryDirectory(prefix="cf-reader-")
        cls.workbook = Path(cls._tmp.name) / "报送.xlsx"
        _build_rule_workbook(cls.workbook)
        cls.result = OoxmlConditionalRuleReader().read(cls.workbook)

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def test_rules_normalized(self):
        rules = self.result.rules
        cellis = next(r for r in rules if r.rule_type == "cellis")
        self.assertEqual(cellis.sheet, "资产负债表")
        self.assertEqual(cellis.applies_to, "B2:B4")
        self.assertEqual(cellis.operator, "greaterThan")
        self.assertEqual(cellis.formula1, "0")
        self.assertEqual(cellis.anchor_cell, "B2")
        self.assertTrue(cellis.supported)
        expression = next(r for r in rules if r.rule_type == "expression")
        self.assertEqual(expression.applies_to, "C2:C5")
        self.assertIn("B2", expression.formula1.replace("$", ""))
        self.assertTrue(expression.stop_if_true)
        # colorScale 读取但不支持
        colorscale = next(r for r in rules if r.rule_type == "colorscale")
        self.assertFalse(colorscale.supported)

    def test_sheet_part_mapping(self):
        self.assertIn("资产负债表", self.result.sheet_parts)
        self.assertTrue(self.result.sheet_parts["资产负债表"].startswith("xl/worksheets/"))

    def test_cell_value_provider(self):
        provider = self.result.cell_values
        self.assertEqual(provider.get("资产负债表", 2, 2), 5)
        self.assertEqual(provider.get("资产负债表", 3, 2), -1)
        self.assertIsNone(provider.get("资产负债表", 99, 99))


class PythonEvaluatorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = TemporaryDirectory(prefix="cf-py-")
        cls.workbook = Path(cls._tmp.name) / "报送.xlsx"
        _build_rule_workbook(cls.workbook)
        cls.reader_result = OoxmlConditionalRuleReader().read(cls.workbook)
        cls.evaluator = PythonConditionalRuleEvaluator()

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def test_cellis_true_false_and_scope(self):
        rules = [r for r in self.reader_result.rules if r.rule_type == "cellis"]
        areas = [(2, 2, 4, 4)]   # B2:D4（含无关列，测试 bounds 过滤）
        results, stats = self.evaluator.evaluate_sheet(
            "资产负债表", rules, areas, self.reader_result)
        cells = {r.cell for r in results}
        # B2=5、B4=0？0 不触发 >0；B3=-1 不触发；B4=0 不触发。
        # 注意：B 列之外的同区域格（如 C2/D2）值空，cellIs 空值不触发。
        self.assertEqual(cells, {"B2"})
        self.assertEqual(stats["evaluated_cells"], 9)

    def test_expression_rule_with_stop_if_true(self):
        rules = [r for r in self.reader_result.rules if r.rule_type == "expression"]
        results, _stats = self.evaluator.evaluate_sheet(
            "资产负债表", rules, [(3, 3, 3, 3)], self.reader_result)
        # C3 对应 $B3=-1 < 0 → 触发（相对引用按锚点平移）
        self.assertEqual([r.cell for r in results], ["C3"])
        self.assertTrue(results[0].triggered)

    def test_unsupported_type_not_treated_as_false(self):
        rules = [r for r in self.reader_result.rules if r.rule_type == "colorscale"]
        results, stats = self.evaluator.evaluate_sheet(
            "资产负债表", rules, [(2, 2, 4, 4)], self.reader_result)
        self.assertEqual(results, [])
        self.assertTrue(stats["unsupported_cells"])


class ComExpressionBuilderTests(unittest.TestCase):
    def test_expression_shift(self):
        from base_audit.cf_reader import _sqref_bounds

        rule = NormalizedConditionalRule(
            rule_id="t", sheet="S", applies_to="C2:C5", rule_type="expression",
            operator="", formula1="$B2<0", formula2="", priority=1,
            stop_if_true=True, anchor_cell="C2", dxf_id=None,
            bounds=tuple(_sqref_bounds("C2:C5")),
        )
        evaluator = ComFormulaEvaluator.__new__(ComFormulaEvaluator)   # 不需 COM
        expression, error = evaluator._build_expression(rule, 4, 3, None)   # C4
        self.assertEqual(error, "")
        # $B2 锁列（$B）不锁行：C4 时行 +2 → $B4
        self.assertEqual(expression.replace(" ", ""), "$B4<0")

    def test_cellis_between_template(self):
        rule = NormalizedConditionalRule(
            rule_id="t", sheet="S", applies_to="B2:B4", rule_type="cellis",
            operator="between", formula1="1", formula2="9", priority=1,
            stop_if_true=False, anchor_cell="B2", dxf_id=None, bounds=(),
        )
        evaluator = ComFormulaEvaluator.__new__(ComFormulaEvaluator)
        expression, error = evaluator._build_expression(rule, 3, 2, None)   # B3
        self.assertEqual(error, "")
        self.assertEqual(expression, "AND(B3>=1,B3<=9)")


class ModeMappingTests(unittest.TestCase):
    def test_legacy_mapping(self):
        self.assertEqual(effective_evaluator_mode("OOXML"), EVALUATOR_PYTHON)
        self.assertEqual(effective_evaluator_mode("NATIVE"), EVALUATOR_COM_DISPLAY)
        self.assertEqual(effective_evaluator_mode("DISPLAY_FORMAT_NATIVE"), EVALUATOR_COM_DISPLAY)
        self.assertEqual(effective_evaluator_mode("COM_EVALUATE"), EVALUATOR_COM)
        self.assertEqual(effective_evaluator_mode(""), EVALUATOR_COM_DISPLAY)

    def test_service_rejects_unknown_mode(self):
        service = ConditionalFormatService("BOGUS")
        with self.assertRaises(RuntimeError):
            service.extract_issues(
                workbook_path=Path("x.xlsx"), ranges=[], structure_ranges=[],
                period="", batch_id="", audit_time="", source_file=Path("x.xlsx"))


class ServicePythonTests(unittest.TestCase):
    def test_end_to_end_python_mode(self):
        with TemporaryDirectory(prefix="cf-svc-") as folder:
            workbook = Path(folder) / "报送.xlsx"
            _build_rule_workbook(workbook)
            service = ConditionalFormatService(EVALUATOR_PYTHON)
            issues, stats = service.extract_issues(
                workbook_path=workbook, ranges=[], structure_ranges=[],
                period="2026-08", batch_id="t1",
                audit_time="2026-09-12 00:00:00", source_file=workbook,
            )
            self.assertEqual(stats["mode"], EVALUATOR_PYTHON)
            self.assertGreater(stats["rules_total"], 0)
            self.assertGreater(stats["triggered"], 0)
            self.assertGreater(len(issues), 0)
            issue = issues[0]
            self.assertEqual(issue.sheet_name, "资产负债表")
            self.assertTrue(str(issue.message))


if __name__ == "__main__":
    unittest.main()


class ReaderParityTests(unittest.TestCase):
    """两个 RuleReader 产出完全相同的 NormalizedConditionalRule。"""

    def test_direct_ooxml_and_openpyxl_rules_identical(self):
        with TemporaryDirectory(prefix="cf-reader-parity-") as folder:
            workbook = Path(folder) / "报送.xlsx"
            _build_rule_workbook(workbook)
            direct = OoxmlConditionalRuleReader().read(workbook)
            from base_audit.cf_reader import OpenPyxlConditionalRuleReader

            openpyxl_result = OpenPyxlConditionalRuleReader().read(workbook)

            def model(rule: NormalizedConditionalRule) -> tuple:
                return (
                    rule.sheet, rule.applies_to, rule.rule_type, rule.operator,
                    rule.formula1, rule.formula2, rule.priority, rule.stop_if_true,
                    rule.anchor_cell,
                )

            direct_models = sorted(model(r) for r in direct.rules)
            openpyxl_models = sorted(model(r) for r in openpyxl_result.rules)
            self.assertEqual(direct_models, openpyxl_models)


class RuleMatcherTests(unittest.TestCase):
    def test_exact_multiple_unknown(self):
        from base_audit.conditional_evaluators import (
            MATCH_EXACT, MATCH_MULTIPLE, MATCH_UNKNOWN,
            match_rules_for_cell,
        )
        from base_audit.cf_reader import _sqref_bounds

        def rule(applies_to: str) -> NormalizedConditionalRule:
            return NormalizedConditionalRule(
                rule_id=f"t|{applies_to}", sheet="S", applies_to=applies_to,
                rule_type="expression", operator="", formula1="A1>0", formula2="",
                priority=1, stop_if_true=False, anchor_cell="A1", dxf_id=None,
                bounds=tuple(_sqref_bounds(applies_to)),
            )

        rules = [rule("A1:A10"), rule("A5:B10")]
        # A1 只被第一条覆盖 → EXACT
        self.assertEqual(match_rules_for_cell(1, 1, rules).status, MATCH_EXACT)
        # A5 被两条覆盖 → MULTIPLE_CANDIDATES（不伪造唯一结论）
        match = match_rules_for_cell(5, 1, rules)
        self.assertEqual(match.status, MATCH_MULTIPLE)
        self.assertEqual(len(match.rules), 2)
        # C1 无覆盖 → UNKNOWN
        self.assertEqual(match_rules_for_cell(1, 3, rules).status, MATCH_UNKNOWN)
