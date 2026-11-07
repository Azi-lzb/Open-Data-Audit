"""Excel 15 位有效数字语义回归测试（conditional_format）。

事故：示例 D35 `$D35<>$D27+$D28+$D34`、示例 E35 同类规则——两侧数学相等
（31925.83 vs 31894.55+31.28+0），双精度尾差 3.6e-12；Excel 判不触发，
本求值器曾用精确比较误报。本测试钉住四处收敛点。
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from base_audit.conditional_format import (  # noqa: E402
    _excel_round15,
    evaluate_cellis_rule,
    evaluate_expression_condition,
)


class _Sheet:
    """最小工作表桩：坐标 → 值。"""

    def __init__(self, values: dict[str, object]):
        self._values = values
        self.title = "汇总表"

    def cell(self, row: int, column: int):
        letters = ""
        number = column
        while number > 0:
            number, rem = divmod(number - 1, 26)
            letters = chr(65 + rem) + letters
        key = f"{letters}{row}"
        return type("C", (), {"value": self._values.get(key)})()


def _eval(formula: str, values: dict[str, object]):
    ws = _Sheet(values)
    return evaluate_expression_condition(formula, ws, 35, 4, 35, 4)


class TestExcelRound15:
    def test_absorbs_tail_diff(self):
        assert _excel_round15(31925.829999999998) == 31925.83
        assert _excel_round15(10967.800000000001) == 10967.8

    def test_keeps_real_difference(self):
        assert _excel_round15(31925.84) == 31925.84
        assert _excel_round15(0.01) == 0.01

    def test_identity_cases(self):
        assert _excel_round15(0.0) == 0.0
        assert _excel_round15(True) is True          # 布尔原样
        assert _excel_round15("x") == "x"            # 非数值原样
        assert _excel_round15(float("inf")) == float("inf")


class TestRealIncident:
    """真机两组：数学相等 + 浮点尾差 → 必须判不触发。"""

    def test_huizhou_d35_no_trigger(self):
        result = _eval("$D35<>$D27+$D28+$D34", {
            "D27": 31894.55, "D28": 31.28, "D34": None, "D35": 31925.83})
        assert result.status == "FALSE", result

    def test_huidong_e35_no_trigger(self):
        result = _eval("$E35<>$E27+$E28+$E34", {
            "E27": 10960.28, "E28": 7.52, "E34": None, "E35": 10967.8})
        assert result.status == "FALSE", result

    def test_same_total_rule_no_trigger(self):
        result = _eval("$D35<>$D21", {"D21": 31925.83, "D35": 31925.83})
        assert result.status == "FALSE", result


class TestNoRegressionOnRealDiffs:
    def test_actual_difference_still_triggers(self):
        result = _eval("$D35<>$D27+$D28+$D34", {
            "D27": 31894.55, "D28": 31.28, "D34": 0, "D35": 32000.00})
        assert result.status == "TRUE", result

    def test_small_but_real_difference_triggers(self):
        """0.01 级别差异必须保留（15 位收敛不吞真差异）。"""
        result = _eval("$D35<>$D21", {"D21": 100.0, "D35": 100.01})
        assert result.status == "TRUE", result

    def test_threshold_rule_unaffected(self):
        """百分比阈值规则（J 列型式）不受影响。"""
        in_range = _eval('AND($J10<>"",OR($J10<-30,$J10>30))', {"J10": -3.17})
        assert in_range.status == "FALSE", in_range
        out_range = _eval('AND($J10<>"",OR($J10<-30,$J10>30))', {"J10": -50.0})
        assert out_range.status == "TRUE", out_range


class TestCellIsRounding:
    def test_cellis_equal_absorbs_tail(self):
        result = evaluate_cellis_rule(
            "equal", ["31925.83"], 31925.829999999998,
            resolve_operand=lambda formulas, index: float(formulas[index]))
        assert result.status == "TRUE", result

    def test_cellis_equal_keeps_real_diff(self):
        result = evaluate_cellis_rule(
            "equal", ["100.01"], 100.0,
            resolve_operand=lambda formulas, index: float(formulas[index]))
        assert result.status == "FALSE", result


class TestExcelBlankAndErrorSemantics:
    """Excel 空白/除零/类型比较语义（依据：129 份 COM 对拍，见
    test-output/s1f1_mode_parity_20261107/parity_report.md）。

    真实 Office 行为（Evaluate 实测，Excel 与 WPS 一致）：
    - 空/空除法 → #DIV/0! 错误，错误传播 → 条件格式不成立；
    - 空白单元格比较按对方类型折算（数值折 0、文本折空串）；
    - 字面量 "" 是文本：0<>"" 为 TRUE；
    - t="e" 错误缓存（#DIV/0! 等）参与比较 → 不成立。
    """

    def test_blank_over_blank_division_does_not_trigger(self):
        result = _eval("C8/C7<=0.9", {})
        assert result.status == "FALSE", result

    def test_blank_numerator_is_zero(self):
        result = _eval("C8/C7<=0.9", {"C7": 5.0})
        assert result.status == "TRUE", result

    def test_blank_denominator_is_division_by_zero(self):
        result = _eval("C8/C7<=0.9", {"C8": 5.0})
        assert result.status == "FALSE", result

    def test_normal_division_still_triggers(self):
        result = _eval("C8/C7<=0.9", {"C8": 4.0, "C7": 5.0})
        assert result.status == "TRUE", result

    def test_guarded_rule_blank_target_stays_silent(self):
        """AND($C7<>"",OR($C7<=1,$C7>=10))：空白目标格不触发（_eval 锚点无平移，$C7 即取值格）。"""
        result = _eval('AND($C7<>"",OR($C7<=1,$C7>=10))', {"C7": None})
        assert result.status == "FALSE", result

    def test_guarded_rule_small_value_triggers(self):
        result = _eval('AND($C7<>"",OR($C7<=1,$C7>=10))', {"C7": 0.5})
        assert result.status == "TRUE", result

    def test_guarded_rule_normal_value_silent(self):
        result = _eval('AND($C7<>"",OR($C7<=1,$C7>=10))', {"C7": 3})
        assert result.status == "FALSE", result

    def test_blank_equals_number_and_empty_text(self):
        assert _eval("C8=0", {}).status == "TRUE"
        assert _eval('C8=""', {}).status == "TRUE"
        assert _eval("C8<>0", {}).status == "FALSE"
        assert _eval('C8<>""', {}).status == "FALSE"

    def test_blank_against_text_compares_as_empty_text(self):
        assert _eval('C8="x"', {}).status == "FALSE"
        assert _eval('C8<>"x"', {}).status == "TRUE"

    def test_zero_literal_is_not_empty_text(self):
        """数字与文本永不相等：0="" 为 FALSE，0<>"" 为 TRUE。"""
        assert _eval('0=""', {}).status == "FALSE"
        assert _eval('0<>""', {}).status == "TRUE"

    def test_error_cached_value_never_triggers(self):
        """t="e" 缓存的 #DIV/0! 参与比较 → 错误传播 → 不成立。"""
        assert _eval("F38>=30", {"F38": "#DIV/0!"}).status == "FALSE"
        assert _eval('F38<>""', {"F38": "#DIV/0!"}).status == "FALSE"

    def test_blank_arithmetic_coerces_to_zero(self):
        assert _eval("C8*2<=4", {}).status == "TRUE"
        assert _eval("C8*2<=4", {"C8": 3}).status == "FALSE"

    def test_sum_skips_blank_cells(self):
        values = {"C1": 2, "C2": 2, "C3": None}
        assert _eval("SUM(C1:C3)>5", values).status == "FALSE"
        assert _eval("SUM(C1:C3)>3", values).status == "TRUE"

    def test_cellis_blank_target_boundary_unchanged(self):
        """已知边界：cellIs 空白目标格按不触发（当前语料 0 条 cellIs，保持现行为）。"""
        result = evaluate_cellis_rule(
            "lessThan", ["0.9"], None,
            resolve_operand=lambda formulas, index: float(formulas[index]))
        assert result.status == "FALSE", result
