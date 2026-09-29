"""Rule applicability for Expression V2.

Legacy comparison rules have two independent concerns: whether a rule applies
to a comparison row, and whether its business assertion is triggered.  This
module owns the first concern so migration, shadow comparison and future V2
execution cannot silently implement different filters.
"""

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Mapping, Protocol


SCOPE_COLUMNS = (
    "适用业务类", "适用机构类代码", "适用地区代码", "适用指标代码",
    "适用数据属性", "适用币种", "适用频度", "适用批次", "适用场景",
)


class RecordLike(Protocol):
    biz_class: str
    org_code: str
    region_code: str
    indicator: str
    data_attr: str
    currency: str
    frequency: str
    batch: str


@dataclass(frozen=True)
class RuleScope:
    biz_class: str = ""
    org_code: str = ""
    region_code: str = ""
    indicator: str = ""
    data_attr: str = ""
    currency: str = ""
    frequency: str = ""
    batch: str = ""
    scene: str = ""


@dataclass(frozen=True)
class MatchResult:
    applicable: bool
    reasons: tuple[str, ...] = ()


def _text(value: object) -> str:
    return str(value or "").replace("'", "").strip()


def _empty_or_wildcard(value: str) -> bool:
    return not value or value in {"*", "0"}


def scope_from_rule(rule: Mapping[str, object]) -> RuleScope:
    """Read V2's explicit scope columns without applying any business rule.

    Empty and ``*`` mean unrestricted.  The migration preserves Legacy's
    special-rule attribute behavior: a slash-delimited field means any listed
    attribute, while an unsplit legacy value is still accepted as a substring.
    """
    return RuleScope(
        biz_class=_text(rule.get("适用业务类")),
        org_code=_text(rule.get("适用机构类代码")),
        region_code=_text(rule.get("适用地区代码")),
        indicator=_text(rule.get("适用指标代码")),
        data_attr=_text(rule.get("适用数据属性")),
        currency=_text(rule.get("适用币种")),
        frequency=_text(rule.get("适用频度")),
        batch=_text(rule.get("适用批次")),
        scene=_text(rule.get("适用场景")),
    )


def _matches_exact(expected: str, actual: str) -> bool:
    if _empty_or_wildcard(expected):
        return True
    return actual in tuple(item.strip() for item in expected.replace("，", ",").split(",") if item.strip())


def _matches_attr(expected: str, actual: str) -> bool:
    if _empty_or_wildcard(expected):
        return True
    values = tuple(item.strip() for item in re.split(r"[,，/]", expected) if item.strip())
    if values:
        return actual in values
    # Compatibility with VBA's InStr(dataColor, attribute) behavior.
    return actual in expected


def _matches_regex(expected: str, actual: str) -> bool:
    if _empty_or_wildcard(expected):
        return True
    try:
        return re.search(expected.replace("左中", "[").replace("右中", "]"), actual) is not None
    except re.error:
        return False


def match(scope: RuleScope, record: RecordLike, *, current_date: str) -> MatchResult:
    """Return whether one V2 rule applies to one comparison record.

    Reasons use user-facing V2 column labels and are deterministic, enabling
    a shadow report to distinguish an inapplicable rule from a failed rule.
    """
    reasons: list[str] = []
    if not _matches_exact(scope.biz_class, str(record.biz_class)):
        reasons.append("业务类")
    if not _matches_regex(scope.org_code, str(record.org_code)):
        reasons.append("机构类代码")
    if not _matches_regex(scope.region_code, str(record.region_code)):
        reasons.append("地区代码")
    if not _matches_exact(scope.indicator, str(record.indicator)):
        reasons.append("指标代码")
    if not _matches_attr(scope.data_attr, str(record.data_attr)):
        reasons.append("数据属性")
    if not _matches_exact(scope.currency, str(record.currency)):
        reasons.append("币种")
    if not _matches_exact(scope.frequency, str(record.frequency)):
        reasons.append("频度")
    if not _matches_exact(scope.batch, str(record.batch)):
        reasons.append("批次")
    # Legacy only gives special meaning to the literal “结转”.  Other scene
    # labels are descriptive and must not turn into an invented filter.
    if scope.scene == "结转" and not str(current_date).endswith("-01-01"):
        reasons.append("适用场景")
    return MatchResult(not reasons, tuple(reasons))
