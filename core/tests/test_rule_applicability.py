"""Expression V2 rule applicability seam tests.

The legacy engines decide whether a rule applies before evaluating its
business assertion.  V2 must preserve that decision independently from the
expression itself.
"""

from __future__ import annotations

from dataclasses import dataclass

from base_audit.systems.s3_central_statistics.rule_applicability import match, scope_from_rule


@dataclass
class Record:
    biz_class: str = "人民币"
    org_code: str = "6k0i"
    region_code: str = "4400000"
    indicator: str = "12N0D"
    data_attr: str = "余额"
    currency: str = "人民币"
    frequency: str = "月"
    batch: str = "2"


def test_special_scope_matches_legacy_filters() -> None:
    scope = scope_from_rule({
        "适用指标代码": "12N0D",
        "适用数据属性": "余额/发生额",
        "适用币种": "人民币",
        "适用频度": "月",
        "适用批次": "2",
        "适用场景": "",
    })

    result = match(scope, Record(), current_date="2026-07-31")

    assert result.applicable is True
    assert result.reasons == ()


def test_scope_reports_all_non_matching_dimensions() -> None:
    scope = scope_from_rule({
        "适用指标代码": "12N0D",
        "适用数据属性": "余额",
        "适用币种": "人民币",
        "适用频度": "月",
        "适用批次": "2",
    })

    result = match(
        scope,
        Record(indicator="22MU6", data_attr="发生额", currency="美元合计", frequency="季", batch="1"),
        current_date="2026-07-31",
    )

    assert result.applicable is False
    assert result.reasons == ("指标代码", "数据属性", "币种", "频度", "批次")


def test_transfer_scene_only_applies_on_january_first() -> None:
    scope = scope_from_rule({"适用场景": "结转"})

    assert match(scope, Record(), current_date="2026-01-01").applicable is True
    assert match(scope, Record(), current_date="2026-01-02").reasons == ("适用场景",)


def test_expression_scope_preserves_org_region_regular_expressions() -> None:
    scope = scope_from_rule({
        "适用业务类": "人民币",
        "适用机构类代码": "^6k0[ij]$",
        "适用地区代码": "^4400",
        "适用指标代码": "12N0D",
        "适用数据属性": "余额",
        "适用币种": "人民币",
        "适用频度": "月",
        "适用批次": "2",
    })

    assert match(scope, Record(), current_date="2026-07-31").applicable is True
    assert match(scope, Record(org_code="6k0p"), current_date="2026-07-31").reasons == ("机构类代码",)
