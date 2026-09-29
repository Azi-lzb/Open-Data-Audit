from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from base_audit.systems.s3_central_statistics.config import CentralConfig
from base_audit.systems.s3_central_statistics.indicator_rule_engine import enrich_indicator_rules
from base_audit.systems.s3_central_statistics.models import CentralDataset, CentralRecord, ComparisonRow


def _record(
    indicator: str,
    value: float,
    *,
    currency: str = "人民币",
    data_attr: str = "余额",
    record_date: str = "2026-08-31",
) -> CentralRecord:
    return CentralRecord(
        biz_class="人民币", record_date=record_date, org_code="6k0i", org_name="测试机构",
        region_code="4400000", region_name="广东省", order_code="1", indicator=indicator,
        indicator_name=indicator, data_attr=data_attr, currency=currency, frequency="月", batch="2",
        value=value,
    )


def _run(
    action: str,
    current: float,
    opponent: float,
    threshold: float,
    *,
    target_unit: str = "亿元",
) -> ComparisonRow:
    current_record = _record("A", current)
    opponent_record = _record("B", opponent)
    dataset = CentralDataset(records=[current_record, opponent_record], record_date="2026-08-31")
    dataset.build_index()
    row = ComparisonRow(record=current_record, prev_value=current, change=0.0, ratio=0.0)
    config = CentralConfig(path=Path("config.xlsx"), rules=[{
        "规则编号": "PAIR-1", "类型": "特殊阈值", "指标代码": "A",
        "指标名称": "A", "说明/比较指标": "B", "备注": action,
        "人民币阀值(亿元)": str(threshold), "美元阀值(亿美元)": "",
        "绝对值变幅(%)": "", "禁用": "",
    }])
    enrich_indicator_rules(
        [row], config, current_index=dataset.key_index,
        current_date="2026-08-31", previous_date="2026-07-31", target_unit=target_unit,
    )
    return row


def test_corresponding_indicator_rule_reads_current_period_index() -> None:
    row = _run("指标需要对应存在。", current=10.0, opponent=0.0, threshold=0.0)
    assert "对应指标无数" in row.need_explain


def test_equal_indicator_rule_reads_current_period_index() -> None:
    row = _run("指标应相等核查。", current=10.0, opponent=11.0, threshold=0.0)
    assert "不等于B偏差-1.0000" in row.need_explain


def test_ratio_rule_uses_only_its_selected_action() -> None:
    row = _run("指标相除值核查。", current=10.0, opponent=2.0, threshold=4.0)
    assert "超过4.0" in row.need_explain
    assert "对应指标" not in row.need_explain


def test_small_ratio_rule_uses_only_its_selected_action() -> None:
    row = _run("指标相除值太小核查。", current=1.0, opponent=10.0, threshold=0.2)
    assert "小于0.2" in row.need_explain


def test_ratio_converts_both_indicators_to_rule_unit() -> None:
    row = _run(
        "指标相除值核查。", current=10.0, opponent=2.0, threshold=4.0,
        target_unit="万元",
    )
    assert "超过4.0" in row.need_explain


def test_subtraction_uses_its_own_action_and_converts_both_sides() -> None:
    row = _run(
        "指标相减应大于核查。", current=100000.0, opponent=20000.0, threshold=7.0,
        target_unit="万元",
    )
    assert row.need_explain == ""  # 10亿元 - 2亿元 = 8亿元，不小于7亿元


def test_amplitude_action_reads_only_amplitude_threshold() -> None:
    current = _record("A", 10.0)
    row = ComparisonRow(record=current, prev_value=5.0, change=5.0, ratio=1.0)
    config = CentralConfig(path=Path("config.xlsx"), rules=[{
        "规则编号": "AMPLITUDE-1", "类型": "特殊阈值", "指标代码": "A",
        "指标名称": "A", "备注": "指标变幅异常需说明。",
        "人民币阀值(亿元)": "999", "美元阀值(亿美元)": "",
        "绝对值变幅(%)": "0.3", "禁用": "",
    }])
    enrich_indicator_rules(
        [row], config, current_index={}, current_date="2026-08-31",
        previous_date="2026-07-31",
    )
    assert "100.00%" in row.need_explain
    assert "变动幅度30%" in row.need_explain
    assert "阈值999亿元" not in row.need_explain


def test_no_data_action_still_triggers_for_real_zero() -> None:
    record = _record("A", 0.0)
    row = ComparisonRow(record=record, prev_value=0.0, change=0.0)
    enrich_indicator_rules(
        [row], _single_special_config("指标不应有数，需说明。"),
        current_index={record.key(): record}, current_date="2026-08-31",
        previous_date="2026-07-31",
    )
    assert row.need_explain == "VBA-1-指标不应有数，需说明。-阈值不适用"


def test_integer_rule_only_marks_actual_fractional_value() -> None:
    config = _single_special_config("指标应为整数。")
    integer = _record("A", 1.0)
    integer_row = ComparisonRow(record=integer, prev_value=1.0, change=0.0, ratio=0.0)
    enrich_indicator_rules(
        [integer_row], config, current_index={integer.key(): integer},
        current_date="2026-08-31", previous_date="2026-07-31",
    )
    fractional = _record("A", 1.2)
    fractional_row = ComparisonRow(record=fractional, prev_value=1.2, change=0.0, ratio=0.0)
    enrich_indicator_rules(
        [fractional_row], config, current_index={fractional.key(): fractional},
        current_date="2026-08-31", previous_date="2026-07-31",
    )
    assert integer_row.need_explain == ""
    assert fractional_row.need_explain == "VBA-1-指标应为整数。-阈值不适用"


def test_change_rule_with_amount_and_amplitude_requires_both_conditions() -> None:
    current = _record("A", 0.01)
    row = ComparisonRow(record=current, prev_value=0.009, change=0.001, ratio=1.0)
    config = CentralConfig(path=Path("config.xlsx"), rules=[{
        "规则编号": "CHANGE-1", "类型": "特殊阈值", "指标代码": "A",
        "指标名称": "A", "备注": "指标有变动需说明。",
        "人民币阀值(亿元)": "1", "美元阀值(亿美元)": "",
        "绝对值变幅(%)": "0.3", "禁用": "",
    }])
    enrich_indicator_rules(
        [row], config, current_index={}, current_date="2026-08-31",
        previous_date="2026-07-31",
    )
    assert row.need_explain == ""  # 变幅达标，但 0.01 亿元未达到 1 亿元。


def test_change_action_requires_actual_change_even_without_threshold() -> None:
    record = _record("A", 25.090579)
    row = ComparisonRow(record=record, prev_value=25.090579, change=0.0, ratio=0.0)
    enrich_indicator_rules(
        [row], _single_special_config("指标有变动需说明。"),
        current_index={record.key(): record}, current_date="2026-08-31",
        previous_date="2026-07-31",
    )
    assert row.need_explain == ""
    record.value = 0.0
    row.prev_value = 1.0
    row.change = -1.0
    enrich_indicator_rules(
        [row], _single_special_config("指标有变动需说明。"),
        current_index={record.key(): record}, current_date="2026-08-31",
        previous_date="2026-07-31",
    )
    assert row.need_explain == "VBA-1-指标有变动需说明。-阈值0亿元"  # 本期真实零仍有下降


def test_change_action_uses_change_amount_not_current_balance() -> None:
    record = _record("A", 3.128064)
    row = ComparisonRow(record=record, prev_value=3.13757, change=-0.009506, ratio=-0.0030297332)
    enrich_indicator_rules(
        [row], _single_special_config("指标有变动需说明。", rmb="2"),
        current_index={record.key(): record}, current_date="2026-08-31",
        previous_date="2026-07-31",
    )
    assert row.need_explain == ""
    row.change = -2.0
    enrich_indicator_rules(
        [row], _single_special_config("指标有变动需说明。", rmb="2"),
        current_index={record.key(): record}, current_date="2026-08-31",
        previous_date="2026-07-31",
    )
    assert row.need_explain == "VBA-1-指标有变动需说明。-阈值2亿元"
    row.need_explain = ""
    row.change = -2.000001
    enrich_indicator_rules(
        [row], _single_special_config("指标有变动需说明。", rmb="2"),
        current_index={record.key(): record}, current_date="2026-08-31",
        previous_date="2026-07-31",
    )
    assert row.need_explain == "VBA-1-指标有变动需说明。-阈值2亿元"


def test_corresponding_indicator_missing_primary_attaches_to_present_counterpart() -> None:
    counterpart = _record("12MUX", 0.096658)
    row = ComparisonRow(record=counterpart, prev_value=0.096658, change=0.0, ratio=0.0)
    config = CentralConfig(path=Path("config.xlsx"), rules=[{
        "规则编号": "SP0640", "类型": "特殊阈值", "指标代码": "12MUQ",
        "说明/比较指标": "12MUX", "备注": "指标需要对应存在。",
        "人民币阀值(亿元)": "", "美元阀值(亿美元)": "",
        "绝对值变幅(%)": "", "禁用": "",
    }])
    enrich_indicator_rules(
        [row], config, current_index={counterpart.key(): counterpart},
        current_date="2026-08-31", previous_date="2026-07-31",
    )
    assert row.need_explain == (
        "SP0640-指标需要对应存在。-阈值不适用-"
        "判定：12MUX有数,12MUQ无数,请核实。"
    )

    row.need_explain = ""
    primary = _record("12MUQ", 0.0)
    enrich_indicator_rules(
        [row], config,
        current_index={counterpart.key(): counterpart, primary.key(): primary},
        current_date="2026-08-31", previous_date="2026-07-31",
    )
    assert row.need_explain == ""  # 已有主指标行时由该行原有动作路径处理

    row.need_explain = ""
    counterpart.value = 0.0
    enrich_indicator_rules(
        [row], config, current_index={counterpart.key(): counterpart},
        current_date="2026-08-31", previous_date="2026-07-31",
    )
    assert row.need_explain == ""  # 真实零不是“有数”


def _single_special_config(action: str, *, rmb: str = "") -> CentralConfig:
    return CentralConfig(path=Path("config.xlsx"), rules=[{
        "规则编号": "VBA-1", "类型": "特殊阈值", "指标代码": "A",
        "指标名称": "A", "备注": action,
        "人民币阀值(亿元)": rmb, "美元阀值(亿美元)": "",
        "绝对值变幅(%)": "", "禁用": "",
    }])


def test_divisible_rule_uses_rmb_threshold_as_divisor() -> None:
    record = _record("A", 0.031)
    row = ComparisonRow(record=record, prev_value=0.031, change=0.0, ratio=0.0)
    enrich_indicator_rules(
        [row], _single_special_config("指标应能被某数整除。", rmb="1000000"),
        current_index={record.key(): record}, current_date="2026-08-31",
        previous_date="2026-07-31",
    )
    assert "应能被1000000整除" in row.need_explain


def test_balance_flow_rule_reads_counterpart_records() -> None:
    balance = _record("A", 110.0, data_attr="余额")
    flow = _record("A", 5.0, data_attr="发生额")
    previous_balance = _record("A", 100.0, data_attr="余额", record_date="2026-07-31")
    row = ComparisonRow(record=balance, prev_value=100.0, change=10.0, ratio=0.1)
    current_index = {balance.key(): balance, flow.key(): flow}
    previous_index = {previous_balance.key(): previous_balance}
    enrich_indicator_rules(
        [row], _single_special_config("指标本期余额应小于（上期余额+本期发生额）。"),
        current_index=current_index, previous_index=previous_index,
        current_date="2026-08-31", previous_date="2026-07-31", target_unit="元",
    )
    assert "上期余额+本期发生额" in row.need_explain


def test_accumulated_balance_flow_rule_reads_current_and_previous_flows() -> None:
    flow = _record("A", 5.0, data_attr="发生额")
    balance = _record("A", 110.0, data_attr="余额")
    previous_flow = _record("A", 1.0, data_attr="发生额", record_date="2026-07-31")
    previous_balance = _record("A", 100.0, data_attr="余额", record_date="2026-07-31")
    row = ComparisonRow(record=flow, prev_value=1.0, change=4.0, ratio=4.0)
    current_index = {flow.key(): flow, balance.key(): balance}
    previous_index = {previous_flow.key(): previous_flow, previous_balance.key(): previous_balance}
    enrich_indicator_rules(
        [row], _single_special_config("指标本期余额应小于（上期余额+本期当年发生额-上期当年累计发生额）。"),
        current_index=current_index, previous_index=previous_index,
        current_date="2026-08-31", previous_date="2026-07-31", target_unit="元",
    )
    assert "上期余额+本期发生额" in row.need_explain
