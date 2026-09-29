"""当前表达式引擎测试：求值语义与跨年累计口径。

迁移期的 V2/旧三表对拍已退役；这里仅保留正式 V3 运行仍依赖的表达式语义回归。
合成数据生成层（core/tools/expression_lae_synthetic.py）未随公开仓发布，
其自检用例（TestSyntheticData）仅在私有仓保留。
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
for _p in (str(ROOT / "src"), str(ROOT)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from base_audit.systems.s3_central_statistics.expression_lae import (  # noqa: E402
    STATUS_ERROR, STATUS_OK, STATUS_UNSUPPORTED,
    ExpressionContext, evaluate_expression_lae, parse_expression,
)



class _Rec:
    def __init__(self, value=None, indicator="12N0D", attr="余额",
                 currency="人民币", frequency="月", batch="1"):
        self.biz_class = "人民币"
        self.org_code = "6k0i"
        self.region_code = "4400000"
        self.indicator = indicator
        self.data_attr = attr
        self.currency = currency
        self.frequency = frequency
        self.batch = batch
        self.value = value


class _Row:
    def __init__(self, value=None, **kw):
        self.record = _Rec(value, **kw)


class _Cell:
    def __init__(self, value):
        self.value = value


KEY = "人民币6k0i440000012N0D余额人民币月1"


def _ctx(*, cur=None, pre=None, cur_date="2026-07-31", pre_date="2026-06-30"):
    return ExpressionContext(
        current_date=cur_date, previous_date=pre_date,
        current_index={KEY: _Cell(cur)} if cur is not None else {},
        previous_index={KEY: _Cell(pre)} if pre is not None else {},
        row=_Row(cur), unit_factor=1.0, thd=1.0,
    )


class TestEngine:
    """引擎语义（§十 / §十一）。"""

    def test_basic_reference_and_ops(self):
        ctx = _ctx(cur=100.0, pre=80.0)
        assert evaluate_expression_lae("[,,,12N0D,余额,,,] > 50", ctx).triggered
        assert evaluate_expression_lae("[,,,12N0D,余额,,,] > {,,,12N0D,余额,,,}", ctx).triggered

    def test_trigger_mode_inverted(self):
        ctx = _ctx(cur=100.0, pre=80.0)
        out = evaluate_expression_lae("[,,,12N0D,余额,,,] >= {,,,12N0D,余额,,,}",
                                     ctx, "表达式不成立")
        assert out.status == STATUS_OK and out.triggered is False

    def test_missing_record_is_zero(self):
        ctx = _ctx(cur=None, pre=None)
        out = evaluate_expression_lae("[,,,12N0D,余额,,,] = 0", ctx)
        assert out.status == STATUS_OK and out.triggered

    def test_zero_value_uses_sentinel(self):
        """存在且为 0 → 0.01 哨兵（与 Legacy complex_rule_engine 一致）。"""
        ctx = _ctx(cur=0.0, pre=100.0)
        assert evaluate_expression_lae("[,,,12N0D,余额,,,] = 0", ctx).triggered is False
        assert evaluate_expression_lae("[,,,12N0D,余额,,,] <> 0", ctx).triggered is True

    def test_not_have_data_uses_presence_not_numeric_sentinel(self):
        """Legacy「指标不应有数」命中真 0，但不命中空值或缺记录。"""
        expression = "AND(EXISTS([,,,12N0D,余额,,,]), NOT(ISBLANK([,,,12N0D,余额,,,])))"

        assert evaluate_expression_lae(expression, _ctx(cur=0.0)).triggered is True
        assert evaluate_expression_lae(expression, _ctx(cur="")).triggered is False
        assert evaluate_expression_lae(expression, _ctx(cur=None)).triggered is False

    def test_if_is_lazy_no_division_by_zero(self):
        """{上期}=0 时不得求值除法分支（§十一）。"""
        ctx = _ctx(cur=100.0, pre=None)     # 上期缺失 → 0
        guarded = evaluate_expression_lae(
            "IF({,,,12N0D,余额,,,}=0, TRUE, [,,,12N0D,余额,,,]/{,,,12N0D,余额,,,} < 1.3)",
            ctx)
        assert guarded.status == STATUS_OK and guarded.triggered
        # 对照：无保护地求值应报 ERROR（不得静默 FALSE）
        raw = evaluate_expression_lae("[,,,12N0D,余额,,,]/{,,,12N0D,余额,,,} < 1.3", ctx)
        assert raw.status == STATUS_ERROR and raw.triggered is False

    def test_and_or_short_circuit(self):
        ctx = _ctx(cur=100.0, pre=1.0)
        assert evaluate_expression_lae("AND(FALSE, 1/0=1)", ctx).triggered is False
        assert evaluate_expression_lae("OR(TRUE, 1/0=1)", ctx).triggered is True

    def test_ifs_first_match_only(self):
        ctx = _ctx(cur=100.0, pre=1.0)
        out = evaluate_expression_lae("IFS(TRUE, 1=1, TRUE, 1/0=1)", ctx)
        assert out.status == STATUS_OK and out.triggered

    def test_unknown_function_is_unsupported_not_false(self):
        ctx = _ctx(cur=1.0, pre=1.0)
        out = evaluate_expression_lae("VLOOKUP(1,2,3)", ctx)
        assert out.status == STATUS_UNSUPPORTED and out.triggered is False

    def test_parse_error_is_error(self):
        ctx = _ctx(cur=1.0, pre=1.0)
        assert evaluate_expression_lae("[,,,12N0D,余额,,,] >", ctx).status == STATUS_ERROR
        assert evaluate_expression_lae("", ctx).status == STATUS_ERROR

    def test_date_functions(self):
        ctx = _ctx(cur=1.0, pre=1.0, cur_date="2026-01-01", pre_date="2025-12-31")
        assert evaluate_expression_lae("MONTH(本期日期)=1", ctx).triggered
        assert evaluate_expression_lae("DAY(本期日期)=1", ctx).triggered
        assert evaluate_expression_lae("YEAR(本期日期)=2026", ctx).triggered
        assert evaluate_expression_lae("YEAR(本期日期)=YEAR(上期日期)", ctx).triggered is False

    def test_math_functions(self):
        ctx = _ctx(cur=1.0, pre=1.0)
        for expr in ("ABS(-3)=3", "ROUND(1.2345,2)=1.23", "INT(3.9)=3",
                     "MOD(7,3)=1", "MIN(1,2)=1", "MAX(1,2)=2", "SUM(1,2,3)=6"):
            assert evaluate_expression_lae(expr, ctx).triggered, expr

    def test_no_eval_exec_in_source(self):
        """§十：表达式引擎源码中不得出现 eval(/exec(。"""
        source = (ROOT / "src" / "base_audit" / "systems" / "s3_central_statistics"
                  / "expression_lae.py").read_text(encoding="utf-8")
        assert "eval(" not in source.replace("evaluate_", "")
        assert "exec(" not in source
        assert "__import__" not in source

    def test_no_office_dependency(self):
        """§九：不得依赖 COM/UNO/soffice（检查 import，而非文档字符串）。"""
        import base_audit.systems.s3_central_statistics.expression_lae as mod

        source = Path(mod.__file__).read_text(encoding="utf-8")
        assert "import win32com" not in source
        assert "import uno" not in source
        assert "import subprocess" not in source



class TestCrossYearBusinessRule:
    """累计跨年分支保持 VBA/Legacy 既有语义，尚未作业务口径改写。"""

    def test_legacy_preserves_cross_year_comparison(self):
        from base_audit.systems.s3_central_statistics.indicator_rule_engine import _accu_message
        from base_audit.systems.s3_central_statistics.models import CentralRecord, ComparisonRow

        def row(value, change):
            rec = CentralRecord(
                biz_class="人民币", record_date="2026-06-30", org_code="6k0i",
                org_name="累计", region_code="4400000", region_name="", order_code="1",
                indicator="12N0D", indicator_name="累计", data_attr="余额",
                currency="人民币", frequency="月", batch="1", value=value)
            return ComparisonRow(record=rec, prev_value=100.0, change=change)

        # 跨年：保持已恢复的 VBA 分支，迁移侧只能作为待业务终裁项。
        assert _accu_message(row(90.0, -10.0), "2026-06-30", "2025-06-30")
        assert _accu_message(row(110.0, 10.0), "2026-03-31", "2025-10-31")
        # 同年下降仍触发
        assert "不应减少" in _accu_message(row(90.0, -10.0), "2026-06-30", "2026-05-31")
        # 负数任何日期都触发
        assert "不应为负数" in _accu_message(row(-1.0, -101.0), "2026-06-30", "2025-06-30")
        # 1月1日结转不判断
        assert _accu_message(row(50.0, -50.0), "2026-01-01", "2025-12-31") == ""
