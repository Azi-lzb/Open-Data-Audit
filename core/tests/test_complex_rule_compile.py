# -*- coding: utf-8 -*-
"""复杂规则编译存活检查（inspect/strict）的 11 个验收场景。

对应 AGENTS.md「VBA 一比一复刻与配置复核铁律」§4–§7：
启用规则要么成功进入执行索引，要么生成带完整定位的 issue；
inspect 只读；compile strict 默认阻断；禁用规则不参与也不阻断。
"""

from __future__ import annotations

import copy
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from base_audit.systems.s3_central_statistics.complex_rule_engine import (  # noqa: E402
    ComplexRuleCompileError,
    compile_complex_rules,
    inspect_complex_rules,
)
from base_audit.systems.s3_central_statistics.config import CentralConfig  # noqa: E402

VALID_EXPR = "and([人民币,,,12M01,余额,人民币,月,1]>0, {人民币,,,12M01,余额,人民币,月,1}=0)"


def _row(rule_id: str, expr: str, *, row: str = "2", disabled: str = "", thd: str = "") -> dict:
    return {
        "规则编号": rule_id, "校验描述": "测试规则", "校验规则": expr,
        "取反标识": "", "Thd值(万元)": thd, "禁用": disabled,
        "__行号__": row, "__工作表__": "表达式校验",
    }


def _config(rows: list[dict]) -> CentralConfig:
    return CentralConfig(path=Path("3.1大集中执行比较_配置.xlsx"), complex_rules=list(rows))


def _ids(result) -> set[str]:
    return {issue.rule_id for issue in result.issues}


def _codes(result, rule_id: str) -> set[str]:
    return {issue.error_code for issue in result.issues if issue.rule_id == rule_id}


# 1. 合法规则
def test_valid_rule_compiles() -> None:
    result = inspect_complex_rules(_config([_row("CX001", VALID_EXPR, row="5")]))
    assert result.total_rows == 1 and result.enabled_rows == 1
    assert result.compiled_rules == 1 and result.failed_rules == 0
    assert not result.issues
    assert set(result.index) == {"12M01"}


# 2. 非 8 段 token
def test_token_with_seven_segments_fails_with_full_localization() -> None:
    expr = "[人民币,,,12P1F余额,人民币,月,2] >= [人民币,,,12P08,余额,人民币,月,2]"
    result = inspect_complex_rules(_config([_row("CX001", expr, row="126")]))
    assert result.failed_rules == 1
    issue = next(i for i in result.issues if i.error_code == "TOKEN_SEGMENT_COUNT")
    assert issue.rule_id == "CX001" and issue.excel_row == "126"
    assert issue.sheet == "表达式校验" and issue.field == "校验表达式"
    assert "12P1F余额" in issue.offending_token
    assert "7 个逗号" in issue.message and "8 段式" in issue.message
    # 高置信候选：指标与数据属性之间疑似缺逗号
    assert "12P1F,余额" in issue.suggestion


# 3. 指标槽为空
def test_empty_indicator_slot_fails() -> None:
    result = inspect_complex_rules(_config([_row("CX001", "[人民币,,,,余额,人民币,月,2] > 0")]))
    assert "EMPTY_INDICATOR" in _codes(result, "CX001")


# 4. 没有可注册本期目标（幽灵规则）
def test_previous_only_rule_has_no_current_target() -> None:
    result = inspect_complex_rules(_config([_row("CX001", "{人民币,,,12M01,余额,人民币,月,1} = 0")]))
    assert "NO_CURRENT_TARGET" in _codes(result, "CX001")
    assert result.compiled_rules == 0 and result.failed_rules == 1


# 5. 非法机构/地区正则
def test_invalid_org_regex_fails() -> None:
    result = inspect_complex_rules(
        _config([_row("CX001", "[人民币,^(6|9,,12M01,余额,人民币,月,1] > 0")]))
    assert "INVALID_REGEX" in _codes(result, "CX001")


# 6. 表达式语法非法（正式 parser 判定）
def test_invalid_expression_syntax_fails() -> None:
    result = inspect_complex_rules(_config([_row("CX001", "[,,,12M01,余额,人民币,月,1] > ")])
                                   if False else
                                   _config([_row("CX001", "[,,,12M01,余额,人民币,月,1] > ")]))
    assert "EXPRESSION_PARSE_ERROR" in _codes(result, "CX001")


# 7. 非空非法 Thd
def test_invalid_nonempty_thd_fails() -> None:
    result = inspect_complex_rules(
        _config([_row("CX001", VALID_EXPR, thd="按业务口径")]))
    assert "INVALID_THD" in _codes(result, "CX001")
    assert result.failed_rules == 1


# 8. 一条规则多指标索引，compiled_rules 按规则身份去重
def test_multi_index_rule_counts_once() -> None:
    expr = "and([,,,12M01,余额,人民币,月,1] > 0, [,,,12M02,余额,人民币,月,1] > 0)"
    result = inspect_complex_rules(_config([_row("CX001", expr)]))
    assert set(result.index) == {"12M01", "12M02"}
    assert result.compiled_rules == 1


# 9. strict 正式编译：存在启用失败规则即阻断，消息含编号与人工修正指引
def test_strict_compile_raises_with_rule_ids() -> None:
    with pytest.raises(ComplexRuleCompileError) as exc_info:
        compile_complex_rules(_config([
            _row("CX001", "[,,,12M01,余额,人民币,月,1] > 0", row="5"),
            _row("CX002", "[人民币,,,12P1F余额,人民币,月,2] > 0", row="6"),
        ]))
    message = str(exc_info.value)
    assert "配置存在 1 条启用但无法执行的表达式规则" in message
    assert "CX002" in message and "CX001" not in message.split("。")[0]
    # 合法规则仍进入索引（strict 报错对象里保留完整结果供调用方使用）。
    assert "12M01" in exc_info.value.result.index


# 10. 禁用规则：不编译、不报 issue、不阻断
def test_disabled_broken_rule_is_ignored() -> None:
    result = inspect_complex_rules(_config([
        _row("CX001", "[人民币,,,12P1F余额,人民币,月,2] > 0", disabled="是"),
        _row("CX002", VALID_EXPR),
    ]))
    assert result.disabled_rows == 1 and result.failed_rules == 0
    assert result.compiled_rules == 1
    assert not result.issues


# 11. inspect 只读：配置对象不被检查器修改
def test_inspect_does_not_mutate_config() -> None:
    rows = [
        _row("CX001", "[人民币,,,12P1F余额,人民币,月,2] > 0", row="9"),
        _row("CX002", VALID_EXPR, row="10", thd="0.5"),
    ]
    config = _config(rows)
    before = copy.deepcopy(config.complex_rules)
    inspect_complex_rules(config)
    assert config.complex_rules == before


# 补充：strict=False 允许拿回索引（供工具/对拍只读使用，不抛错）
def test_compile_non_strict_returns_index_with_failures() -> None:
    result_or_index = compile_complex_rules(
        _config([_row("CX001", "[人民币,,,12P1F余额,人民币,月,2] > 0")]), strict=False)
    assert result_or_index == {}
