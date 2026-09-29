from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from base_audit.systems.s3_central_statistics.complex_rule_engine import (
    SCHEMA_FIVE_SEGMENT_V1,
    SCHEMA_LEGACY_8,
    _filter_rules_by_date,
    compile_complex_rules,
    compile_five_segment_rules,
    enrich_complex_rules,
)
from base_audit.systems.s3_central_statistics.config import CentralConfig, load_central_config
from base_audit.systems.s3_central_statistics.models import CentralRecord, ComparisonRow


@pytest.fixture(params=[SCHEMA_LEGACY_8, SCHEMA_FIVE_SEGMENT_V1])
def expression_schema(request):
    return request.param


def _config(source_group: str) -> CentralConfig:
    config = CentralConfig(path=ROOT)
    config.complex_rules = [{
        "规则编号": "CX结0054",
        "来源": source_group,
        "校验描述": "结转规则命中",
        "校验规则": "[人民币,,,12MRK,余额,人民币,月,1] > 0",
        "禁用": "",
    }]
    config.five_segment_rules = [{
        "规则编号": "CX结0054",
        "来源分组": source_group,
        "规则说明": "结转规则命中",
        "校验表达式": "[12MRK,余额,人民币,月,1] > 0",
        "是否取反": "否",
        "启用": "是",
        "容差值(万元)": "",
    }]
    return config


def _execute(
    config: CentralConfig, expression_schema: str, current_date: str,
    *, fres: set[str] | None = None,
) -> ComparisonRow:
    record = CentralRecord(
        biz_class="人民币",
        record_date=current_date,
        org_code="440000",
        org_name="测试机构",
        region_code="44",
        region_name="测试地区",
        order_code="1",
        indicator="12MRK",
        indicator_name="测试指标",
        data_attr="余额",
        currency="人民币",
        frequency="月",
        batch="1",
        value=1.0,
    )
    row = ComparisonRow(record=record)
    enrich_complex_rules(
        [row], config,
        current_index={record.key(): record},
        previous_index={},
        fres=fres,
        expression_dates=(current_date, "2026-07-31"),
        expression_schema=expression_schema,
    )
    return row


def test_rollover_expression_runs_only_on_january_first(expression_schema):
    config = _config("结转")

    august_row = _execute(config, expression_schema, "2026-08-31")
    assert august_row.need_explain == ""

    january_row = _execute(config, expression_schema, "2027-01-01")
    assert january_row.need_explain == "结转规则命中"


def test_non_rollover_expression_remains_active_on_august_31(expression_schema):
    row = _execute(_config("自定义"), expression_schema, "2026-08-31")
    assert row.need_explain == "结转规则命中"


@pytest.mark.parametrize("source_group", ["单频", "跨期", "年报", "自定义", "跨期核对"])
def test_non_rollover_expression_is_not_loaded_on_january_first(
    expression_schema, source_group: str,
) -> None:
    row = _execute(_config(source_group), expression_schema, "2027-01-01")
    assert row.need_explain == ""


def test_expression_frequency_comes_from_tokens_not_source_group(expression_schema) -> None:
    config = _config("自定义")  # token 明确写“月,1”
    missing_frequency = _execute(
        config, expression_schema, "2026-08-31", fres={"", "季1"},
    )
    matching_frequency = _execute(
        config, expression_schema, "2026-08-31", fres={"", "月1"},
    )
    assert missing_frequency.need_explain == ""
    assert matching_frequency.need_explain == "结转规则命中"


def test_expression_requires_previous_token_frequency_too(expression_schema) -> None:
    config = _config("跨期")
    config.complex_rules[0]["校验规则"] = (
        "[人民币,,,12MRK,余额,人民币,月,1] > "
        "{人民币,,,12PRE,余额,人民币,季,1}"
    )
    config.five_segment_rules[0]["校验表达式"] = (
        "[12MRK,余额,人民币,月,1] > {12PRE,余额,人民币,季,1}"
    )
    # 只有本期月1，没有上期季1：VBA 检查两个 token 后整条跳过。
    row = _execute(config, expression_schema, "2026-08-31", fres={"", "月1"})
    assert row.need_explain == ""


def test_real_config_expression_groups_follow_date_gate(expression_schema) -> None:
    config = load_central_config(ROOT.parent / "config" / "3.1大集中执行比较_配置.xlsx")
    compiler = (compile_five_segment_rules if expression_schema == SCHEMA_FIVE_SEGMENT_V1
                else compile_complex_rules)
    compiled = compiler(config)
    def unique_rules(index):
        return {id(rule): rule for rules in index.values() for rule in rules}.values()

    assert any(rule.source_group == "结转" for rule in unique_rules(compiled))
    assert any(rule.source_group != "结转" for rule in unique_rules(compiled))
    january = _filter_rules_by_date(compiled, "2027-01-01")
    august = _filter_rules_by_date(compiled, "2026-08-31")
    assert all(rule.source_group == "结转" for rule in unique_rules(january))
    assert all(rule.source_group != "结转" for rule in unique_rules(august))
