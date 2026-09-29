# -*- coding: utf-8 -*-
"""百分数指标的环比语义：差异值就是变动率（百分点），警戒档校验不适用。"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from base_audit.period_compare import (  # noqa: E402
    IndicatorDef,
    IndicatorValue,
    PeriodConfig,
    compare_periods,
)

_ORG = "测试机构"


def _cfg() -> PeriodConfig:
    indicators = {
        code: IndicatorDef(code=code, name=f"股东持股比例{code}", data_type="百分数", no_unit_convert=True)
        for code in ("20201028", "20201029")
    }
    return PeriodConfig(indicators=indicators)


def _iv(code: str, value) -> IndicatorValue:
    return IndicatorValue(
        org_name="测试机构", org_code="6k0i", code=code, name=code,
        value=value, date="2026-08-31", form="20202",
    )


def _indices(values):
    return {(v.org_name, v.code): v for v in values}


class PercentageIndicatorTests(unittest.TestCase):
    def test_relative_drop_no_longer_flags(self) -> None:
        """修复前：4.67 vs 7.34 相对降幅 -57% 命中 [-80,-50) 档并标异常。

        百分数指标差异即变动率（-0.83 个百分点），不应触发警戒档。
        """
        cfg = _cfg()
        cur = _indices([_iv("20201029", 4.67)])
        pre = _indices([_iv("20201029", 7.34)])
        rows = compare_periods(cur, pre, cfg)
        self.assertEqual(rows[0]["环比变动"], -2.67)
        self.assertEqual(rows[0]["备注"], "")
        self.assertEqual(rows[0]["级别"], "")

    def test_percentage_small_change_no_trigger(self) -> None:
        cfg = _cfg()
        cur = _indices([_iv("20201028", 9.92)])
        pre = _indices([_iv("20201028", 10.01)])
        rows = compare_periods(cur, pre, cfg)
        self.assertAlmostEqual(rows[0]["环比变动"], -0.09)
        self.assertEqual(rows[0]["备注"], "")

    def test_percentage_missing_previous_still_reports_presence(self) -> None:
        """上期缺失的「本期有，上期无」属于数据存在性异常，仍保留。"""
        cfg = _cfg()
        cur = _indices([_iv("20201028", 9.92)])
        pre = _indices([])
        rows = compare_periods(cur, pre, cfg)
        self.assertEqual(rows[0]["备注"], "本期有，上期无")


if __name__ == "__main__":
    unittest.main()
