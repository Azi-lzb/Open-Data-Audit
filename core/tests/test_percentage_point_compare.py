# -*- coding: utf-8 -*-
"""百分数指标按百分点差异计算（数据属性驱动分流）。

铁律：数据属性=百分数 的指标跨期变化 = 本期 - 上期（百分点），
不计算相对变化率、不计算倍率、不参与普通 Bxxx 增降幅警戒；
普通数值/文字指标逻辑不变；分流按 数据属性，禁止指标编号硬编码。
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

from config_paths import resolve_test_config

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from base_audit.period_compare import (  # noqa: E402
    IndicatorDef,
    IndicatorValue,
    PeriodConfig,
    compare_periods,
    load_period_config,
)


def _real_config() -> PeriodConfig:
    return load_period_config(resolve_test_config("2.报表采集系统_配置.xlsx"))


def _cfg_with(code: str, data_type: str, *, no_unit_convert: bool = True,
              alerts: bool = True) -> PeriodConfig:
    from base_audit.period_compare import AlertRange

    cfg = PeriodConfig(indicators={
        code: IndicatorDef(code=code, name=f"指标{code}", data_type=data_type,
                           no_unit_convert=no_unit_convert),
    })
    if alerts:
        # 与真实配置一致的警戒档（截取），供普通数值指标回归断言。
        cfg.alerts = [
            AlertRange(lower_bound=-0.9, remark="降幅(-90%,-80%]", fill_row=False),
            AlertRange(lower_bound=-0.8, remark="降幅(-80%,-50%]", fill_row=False),
            AlertRange(lower_bound=-0.5, remark="降幅(-50%,-30%]", fill_row=False),
            AlertRange(lower_bound=0.3, remark="增幅[30%,50%)", fill_row=False),
            AlertRange(lower_bound=0.5, remark="增幅[50%,1倍)", fill_row=False),
            AlertRange(lower_bound=1.0, remark="增幅[1倍,5倍)", fill_row=False),
        ]
    return cfg


def _iv(code: str, value) -> IndicatorValue:
    return IndicatorValue(org_name="测试机构", org_code="6k0i", code=code,
                          name=f"指标{code}", value=value, date="2026-06-30", form="20202")


def _indices(cur_vals, pre_vals):
    cur = {("测试机构", code): _iv(code, value) for code, value in cur_vals.items()}
    pre = {("测试机构", code): _iv(code, value) for code, value in pre_vals.items()}
    return cur, pre


class PercentagePointTests(unittest.TestCase):
    """spec 十六 测试 2/3/4/5/6：五组真实数值（用真实配置的警戒档与指标参照）。"""

    def setUp(self):
        self.cfg = _real_config()

    def _assert_percentage(self, code: str, cur_value, pre_value, expected: float):
        cur, pre = _indices({code: cur_value}, {code: pre_value})
        rows = compare_periods(cur, pre, self.cfg)
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertAlmostEqual(row["当期数"], float(cur_value))
        self.assertAlmostEqual(row["上期数"], float(pre_value))
        self.assertAlmostEqual(row["环比变动"], expected, places=6)
        self.assertEqual(row["备注"], "")          # 不命中任何 Bxxx 警戒
        self.assertEqual(row["级别"], "")
        self.assertIn("百分点", row["计算过程"])     # 计算过程说明可见

    def test_20201028_1_vs_9_92(self) -> None:
        self._assert_percentage("20201028", 1, 9.92, -8.92)

    def test_20201029_2_vs_4_67(self) -> None:
        self._assert_percentage("20201029", 2, 4.67, -2.67)

    def test_20201031_4_vs_2_72(self) -> None:
        self._assert_percentage("20201031", 4, 2.72, 1.28)

    def test_20201032_5_vs_2_16(self) -> None:
        self._assert_percentage("20201032", 5, 2.16, 2.84)

    def test_20201037_10_vs_1_79(self) -> None:
        self._assert_percentage("20201037", 10, 1.79, 8.21)


class ZeroValueTests(unittest.TestCase):
    """spec 十一：0 值场景无除零、无异常倍率。"""

    def test_pre_zero(self) -> None:
        cfg = _cfg_with("Z001", "百分数")
        cur, pre = _indices({"Z001": 5}, {"Z001": 0})
        rows = compare_periods(cur, pre, cfg)
        self.assertAlmostEqual(rows[0]["环比变动"], 5.0)
        self.assertEqual(rows[0]["备注"], "")

    def test_cur_zero(self) -> None:
        cfg = _cfg_with("Z001", "百分数")
        cur, pre = _indices({"Z001": 0}, {"Z001": 5})
        rows = compare_periods(cur, pre, cfg)
        self.assertAlmostEqual(rows[0]["环比变动"], -5.0)

    def test_both_zero(self) -> None:
        """VBA 口径：两期都为 0 的行不进入结果（零值不读），无异常倍率产生。"""
        cfg = _cfg_with("Z001", "百分数")
        cur, pre = _indices({"Z001": 0}, {"Z001": 0})
        rows = compare_periods(cur, pre, cfg)
        self.assertEqual(rows, [])


class NotHardcodedTests(unittest.TestCase):
    """spec 十六 测试 11：新指标代码（不在 20201028..37）同样走百分点逻辑。"""

    def test_new_percentage_code_uses_percentage_point(self) -> None:
        cfg = _cfg_with("99990001", "百分数")
        cur, pre = _indices({"99990001": 1}, {"99990001": 9.92})
        rows = compare_periods(cur, pre, cfg)
        self.assertAlmostEqual(rows[0]["环比变动"], -8.92)
        self.assertEqual(rows[0]["备注"], "")


class DataQualityTests(unittest.TestCase):
    """spec 十二：缺失/文字类沿用既有数据质量规则。"""

    def test_missing_previous_reports_presence(self) -> None:
        """单边缺失沿用既有数据质量规则：只标存在性，不算变化率。"""
        cfg = _cfg_with("Z002", "百分数")
        cur, pre = _indices({"Z002": 5}, {})
        rows = compare_periods(cur, pre, cfg)
        self.assertEqual(rows[0]["备注"], "本期有，上期无")
        self.assertIsNone(rows[0]["环比变动"])       # 存在性行不产生变化率

    def test_missing_current_reports_presence(self) -> None:
        cfg = _cfg_with("Z002", "百分数")
        cur, pre = _indices({}, {"Z002": 5})
        rows = compare_periods(cur, pre, cfg)
        self.assertEqual(rows[0]["备注"], "本期无，上期有")
        self.assertIsNone(rows[0]["环比变动"])


class NormalIndicatorRegressionTests(unittest.TestCase):
    """spec 十六 测试 1/10：普通数值与文字指标逻辑不变。"""

    def test_normal_value_still_uses_relative_change(self) -> None:
        cfg = _cfg_with("N001", "余额", no_unit_convert=False)
        from base_audit.period_compare import AlertRange

        cfg.alerts = [AlertRange(lower_bound=0.3, remark="增幅[30%,50%)", fill_row=False)]
        cur, pre = _indices({"N001": 150}, {"N001": 100})
        rows = compare_periods(cur, pre, cfg)
        self.assertAlmostEqual(rows[0]["环比变动"], 50.0)
        self.assertIn("增幅", rows[0]["备注"])   # 仍进入 Bxxx 警戒
        self.assertEqual(rows[0]["计算过程"], "")   # 非百分数无百分点说明

    def test_text_indicator_not_numeric_compared(self) -> None:
        cfg = _cfg_with("T001", "文字")
        cur, pre = _indices({"T001": "甲银行"}, {"T001": "甲银行"})
        rows = compare_periods(cur, pre, cfg)
        # 相同文字：不触发“文字变动”，也不进入数值环比。
        self.assertEqual(rows[0]["备注"], "")
        self.assertIsNone(rows[0]["环比变动"])
        # 文字变动场景：备注=文字变动，仍不产生数值环比。
        cur2, pre2 = _indices({"T001": "甲银行"}, {"T001": "乙银行"})
        rows2 = compare_periods(cur2, pre2, cfg)
        self.assertEqual(rows2[0]["备注"], "文字变动")
        self.assertIsNone(rows2[0]["环比变动"])


if __name__ == "__main__":
    unittest.main()
