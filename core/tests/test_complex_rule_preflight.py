from pathlib import Path

import pytest

from base_audit.systems.s3_central_statistics.complex_rule_engine import (
    ComplexRuleCompileError,
    compile_complex_rules,
    inspect_complex_rules,
    enrich_complex_rules,
)
from base_audit.systems.s3_central_statistics.config import CentralConfig, validate_central_config


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
