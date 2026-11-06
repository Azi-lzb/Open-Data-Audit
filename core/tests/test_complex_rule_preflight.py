from pathlib import Path

import pytest

from base_audit.systems.s3_central_statistics.complex_rule_engine import (
    SCHEMA_FIVE_SEGMENT_V1,
    ComplexRuleCompileError,
    compile_complex_rules,
    inspect_complex_rules,
    enrich_complex_rules,
)
from base_audit.systems.s3_central_statistics.config import CentralConfig, validate_central_config
from base_audit.systems.s3_central_statistics.models import CentralRecord, ComparisonRow


def _rule(rule_id: str, expression: str, **overrides) -> dict[str, str]:
    row = {
        "规则编号": rule_id,
        "校验描述": "测试规则",
        "校验规则": expression,
        "取反标识": "",
        "Thd值(万元)": "",
        "禁用": "",
        "来源": "单频",
        "__工作表__": "表达式校验",
        "__行号__": "12",
    }
    row.update(overrides)
    return row


def _config(*rows: dict[str, str]) -> CentralConfig:
    return CentralConfig(path=Path("3.1.xlsx"), complex_rules=list(rows))


def test_valid_rule_survives_and_enters_index() -> None:
    cfg = _config(_rule("CX001", "[人民币,,,12P1F,余额,人民币,月,2] > 0"))
    result = inspect_complex_rules(cfg)
    assert (result.enabled_rows, result.compiled_rules, result.failed_rules) == (1, 1, 0)
    assert "12P1F" in result.index
    assert compile_complex_rules(cfg)["12P1F"][0].rule_id == "CX001"


def test_seven_segment_token_reports_candidate_but_never_mutates() -> None:
    expression = "[人民币,,,12P1F余额,人民币,月,2] > 0"
    row = _rule("CX002", expression)
    cfg = _config(row)
    result = inspect_complex_rules(cfg)
    assert result.compiled_rules == 0
    assert result.failed_rules == 1
    issue = next(issue for issue in result.issues if issue.error_code == "TOKEN_SEGMENT_COUNT")
    assert issue.rule_id == "CX002"
    assert issue.excel_row == "12"
    assert "7 个逗号" in issue.message and "大概率取不到数据" in issue.message
    assert "12P1F,余额" in issue.suggestion
    assert row["校验规则"] == expression
    with pytest.raises(ComplexRuleCompileError) as exc_info:
        compile_complex_rules(cfg)
    assert "CX002" in str(exc_info.value)
    assert "程序未自动修改配置" in str(exc_info.value)


def test_empty_indicator_is_reported() -> None:
    cfg = _config(_rule("CX003", "[人民币,,,,余额,人民币,月,2] > 0"))
    result = inspect_complex_rules(cfg)
    assert result.failed_rules == 1
    assert any(issue.error_code == "EMPTY_INDICATOR" for issue in result.issues)


def test_rule_without_current_target_is_reported() -> None:
    cfg = _config(_rule("CX004", "{人民币,,,12P1F,余额,人民币,月,2] > 0".replace("]", "}")))
    result = inspect_complex_rules(cfg)
    assert result.failed_rules == 1
    assert any(issue.error_code == "NO_CURRENT_TARGET" for issue in result.issues)


def test_invalid_org_regex_is_reported_before_runtime() -> None:
    cfg = _config(_rule("CX005", "[人民币,(,地区,12P1F,余额,人民币,月,2] > 0"))
    result = inspect_complex_rules(cfg)
    assert result.failed_rules == 1
    assert any(issue.error_code == "INVALID_REGEX" for issue in result.issues)


def test_invalid_expression_uses_formal_parser() -> None:
    cfg = _config(_rule("CX006", "[人民币,,,12P1F,余额,人民币,月,2] +"))
    result = inspect_complex_rules(cfg)
    assert result.failed_rules == 1
    assert any(issue.error_code == "EXPRESSION_PARSE_ERROR" for issue in result.issues)


def test_invalid_nonempty_thd_is_not_silently_defaulted() -> None:
    cfg = _config(_rule("CX007", "[人民币,,,12P1F,余额,人民币,月,2] > Thd", **{"Thd值(万元)": "abc"}))
    result = inspect_complex_rules(cfg)
    assert result.failed_rules == 1
    assert any(issue.error_code == "INVALID_THD" for issue in result.issues)


def test_multi_target_rule_counts_once_but_indexes_each_target() -> None:
    cfg = _config(_rule(
        "CX008",
        "[人民币,,,12P1F,余额,人民币,月,2] > [人民币,,,12P2F,余额,人民币,月,2]",
    ))
    result = inspect_complex_rules(cfg)
    assert result.compiled_rules == 1
    assert result.failed_rules == 0
    assert set(result.index) == {"12P1F", "12P2F"}


def test_disabled_broken_rule_does_not_block_execution() -> None:
    cfg = _config(_rule("CX009", "[坏token] > 0", **{"禁用": "是"}))
    result = inspect_complex_rules(cfg)
    assert (result.disabled_rows, result.enabled_rows, result.failed_rules) == (1, 0, 0)
    assert compile_complex_rules(cfg) == {}


def test_validate_config_surfaces_compile_issue_with_location_and_suggestion() -> None:
    cfg = _config(_rule("CX010", "[人民币,,,12P1F余额,人民币,月,2] > 0"))
    errors = validate_central_config(cfg)
    message = "\n".join(errors)
    assert "CX010" in message
    assert "第12行" in message
    assert "原始 token" in message
    assert "建议人工复核" in message
    assert "程序未自动修改配置" in message


def test_formal_enrichment_blocks_when_enabled_rule_cannot_compile() -> None:
    cfg = _config(_rule("CX011", "[人民币,,,12P1F余额,人民币,月,2] > 0"))
    with pytest.raises(ComplexRuleCompileError):
        enrich_complex_rules(
            [], cfg, current_index={}, previous_index={},
            target_unit="亿元", exempt=set(),
        )


def test_enrich_reports_not_ready_rules_by_frequency_key() -> None:
    """数据批次未就绪的规则整条跳过（VBA notfreNum），编号进统计供运行日志展示。"""
    cfg = _config(
        _rule("CX跨A", "[人民币,,,12P1F,余额,人民币,月,2] > [人民币,,,12P1E,余额,人民币,月,1]"),
        _rule("CX跨B", "[人民币,,,12P1F,余额,人民币,月,1] > 0"),
    )
    stats: dict = {}
    # fres 只含 月1（模拟仅导入月报1批）：需要 月2 的 CX跨A 跳过，CX跨B 就绪。
    enrich_complex_rules(
        [], cfg, current_index={}, previous_index={},
        target_unit="亿元", exempt=set(), fres={"", "月", "月1"},
        stats_out=stats,
    )
    assert stats["expression_rules_ready"] == ["CX跨B"]
    assert stats["expression_rules_skipped_not_ready"] == ["CX跨A"]


def _scope_test_row(org_code: str = "6k0i") -> ComparisonRow:
    record = CentralRecord(
        biz_class="人民币", record_date="2026-10-31", org_code=org_code,
        org_name="示例机构", region_code="4400000", region_name="广东省",
        order_code="001", indicator="12P1F", indicator_name="测试指标",
        data_attr="余额", currency="人民币", frequency="月", batch="1",
        value=1.0,
    )
    return ComparisonRow(record=record)


def test_legacy_cross_period_rule_without_vba_org_scope_is_reported_and_skipped() -> None:
    rule = _rule(
        "CX_SCOPE_EMPTY",
        "[人民币,,,12P1F,余额,人民币,月,1] > 0",
        **{"来源": "跨期核对"},
    )
    cfg = _config(rule)
    cfg.five_segment_rules = [{
        "规则编号": "CX_SCOPE_EMPTY", "机构类代码": "", "地区代码": "",
    }]
    row = _scope_test_row()
    stats: dict = {}

    enrich_complex_rules(
        [row], cfg, current_index={row.record.key(): row.record},
        previous_index={}, target_unit="亿元", exempt=set(),
        fres={"", "月", "月1"}, stats_out=stats,
    )

    assert row.need_explain == ""
    assert stats["expression_rules_skipped_vba_scope_missing"] == ["CX_SCOPE_EMPTY"]
    assert stats["expression_rules_ready"] == []


@pytest.mark.parametrize(
    ("org_pattern", "org_code", "expected"),
    [("^6", "6k0i", True), ("^7", "6k0i", False)],
)
def test_legacy_cross_period_rule_respects_paired_vba_org_scope(
    org_pattern: str, org_code: str, expected: bool,
) -> None:
    rule = _rule(
        "CX_SCOPE_SET",
        "[人民币,,,12P1F,余额,人民币,月,1] > 0",
        **{"来源": "跨期核对"},
    )
    cfg = _config(rule)
    cfg.five_segment_rules = [{
        "规则编号": "CX_SCOPE_SET", "机构类代码": org_pattern, "地区代码": "",
    }]
    row = _scope_test_row(org_code)

    enrich_complex_rules(
        [row], cfg, current_index={row.record.key(): row.record},
        previous_index={}, target_unit="亿元", exempt=set(),
        fres={"", "月", "月1"},
    )

    assert bool(row.need_explain) is expected


def _five_segment_rule(rule_id: str, *, org_pattern: str = "", region_pattern: str = "") -> dict[str, str]:
    return {
        "规则编号": rule_id,
        "来源分组": "跨期核对",
        "来源表单": "测试表",
        "机构类代码": org_pattern,
        "地区代码": region_pattern,
        "规则说明": "五段式测试规则",
        "校验表达式": "[12P1F,余额,人民币,月,1] > 0",
        "触发方式": "表达式成立",
        "容差值(万元)": "",
        "启用": "是",
        "停用原因": "",
        "备注": "",
    }


def test_five_segment_cross_period_rule_without_org_scope_is_skipped() -> None:
    rule_id = "CX_FIVE_EMPTY_SCOPE"
    cfg = _config(_rule(
        rule_id, "[人民币,,,12P1F,余额,人民币,月,1] > 0",
        **{"来源": "跨期核对"},
    ))
    cfg.five_segment_rules = [_five_segment_rule(rule_id)]
    row = _scope_test_row()
    stats: dict = {}

    enrich_complex_rules(
        [row], cfg, current_index={row.record.key(): row.record},
        previous_index={}, target_unit="亿元", exempt=set(),
        fres={"", "月", "月1"}, stats_out=stats,
        expression_schema=SCHEMA_FIVE_SEGMENT_V1,
    )

    assert row.need_explain == ""
    assert stats["expression_rules_skipped_vba_scope_missing"] == [rule_id]
    assert stats["expression_rules_ready"] == []


@pytest.mark.parametrize(
    ("org_pattern", "source_group", "expected"),
    [("^6", "跨期核对", True), ("^7", "跨期核对", False), ("", "自定义", True)],
)
def test_five_segment_scope_gate_matches_only_vba_cross_period_group(
    org_pattern: str, source_group: str, expected: bool,
) -> None:
    rule_id = "CX_FIVE_SCOPE"
    cfg = _config(_rule(
        rule_id, "[人民币,,,12P1F,余额,人民币,月,1] > 0",
        **{"来源": source_group},
    ))
    five_rule = _five_segment_rule(rule_id, org_pattern=org_pattern)
    five_rule["来源分组"] = source_group
    cfg.five_segment_rules = [five_rule]
    row = _scope_test_row()

    enrich_complex_rules(
        [row], cfg, current_index={row.record.key(): row.record},
        previous_index={}, target_unit="亿元", exempt=set(),
        fres={"", "月", "月1"}, expression_schema=SCHEMA_FIVE_SEGMENT_V1,
    )

    assert bool(row.need_explain) is expected
