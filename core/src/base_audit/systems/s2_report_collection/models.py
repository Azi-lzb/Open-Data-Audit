"""V2 报表采集审核的显式领域对象。"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


STATUS_NORMAL = "正常"
STATUS_ABNORMAL = "异常"
SEVERITIES = ("提示", "关注", "严重")


@dataclass(frozen=True)
class InstitutionProfile:
    institution_id: str
    institution_name: str
    social_credit_code: str = ""
    institution_type: str = ""
    handling_branch: str = ""
    region: str = ""
    report_item: str = ""


@dataclass(frozen=True)
class IndicatorConfig:
    indicator_code: str
    indicator_name: str = ""
    form_code: str = ""
    data_type: str = ""
    disabled: bool = False
    note: str = ""
    # S2-F01 环比策略字段追加在旧字段之后，保留旧配置/测试的构造兼容性。
    period_enabled: bool = True
    period_strategy_group: str = ""
    min_change_value: float | None = None


@dataclass(frozen=True)
class ExternalCheckRule:
    """一个报表指标与大集中指标之间的比较约定；系统标识仅供内部兼容。"""
    rule_id: str
    indicator_code: str
    external_system: str
    external_indicator_name: str
    comparison_method: str = "等于"
    tolerance: float = 0.0
    severity: str = "严重"
    enabled: bool = True
    source: SourceRef | None = None


@dataclass(frozen=True)
class PeriodBandConfig:
    rule_id: str
    lower: float | None
    upper: float | None
    description: str = ""
    severity: str = ""
    disabled: bool = False
    strategy_group: str = ""
    data_type: str = ""

    def includes(self, value: float) -> bool:
        """V2 统一采用 [lower, upper)；lower 空=无穷小，upper 空=无穷大。"""
        return (self.lower is None or value >= self.lower) and (
            self.upper is None or value < self.upper)


@dataclass(frozen=True)
class RuleConfig:
    rule_id: str
    rule_type: str
    description: str
    expression: str
    severity: str = "提示"
    threshold: float | None = None
    invert: bool = False
    disabled: bool = False
    note: str = ""
    source: SourceRef | None = None


@dataclass
class UnitSettings:
    source_unit: str = "元"
    external_file_unit: str = "元"
    output_file_unit: str = "元"


@dataclass
class ReportAuditConfig:
    indicators: dict[str, IndicatorConfig] = field(default_factory=dict)
    institutions_by_id: dict[str, InstitutionProfile] = field(default_factory=dict)
    institutions_by_name: dict[str, InstitutionProfile] = field(default_factory=dict)
    bands: list[PeriodBandConfig] = field(default_factory=list)
    external_rules: list[ExternalCheckRule] = field(default_factory=list)
    rules: list[RuleConfig] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    unit_settings: UnitSettings = field(default_factory=UnitSettings)
    hide_unchanged_period_rows: bool = False
    hide_matching_external_rows: bool = False


@dataclass(frozen=True)
class SourceRef:
    source_file: str
    source_sheet: str
    source_row: int


@dataclass(frozen=True)
class IndicatorRecord:
    institution_id: str
    institution_name: str
    social_credit_code: str
    institution_type: str
    handling_branch: str
    region: str
    form_code: str
    indicator_code: str
    indicator_name: str
    period: str
    source_date: str
    value: Any
    value_type: str
    source_unit: str
    normalized_value: Any
    normalized_unit: str
    source: SourceRef

    @property
    def logical_key(self) -> tuple[str, str, str]:
        """当前真实数据的逻辑唯一键；form_code 是显式业务属性而非隐式主键。"""
        return (self.institution_id, self.indicator_code, self.period)


@dataclass(frozen=True)
class RelatedRecord:
    indicator_code: str
    indicator_name: str
    form_code: str
    period: str
    value: Any
    source: SourceRef


@dataclass(frozen=True)
class AuditFinding:
    finding_id: str
    audit_type: str
    status: str
    severity: str = ""
    region: str = ""
    handling_branch: str = ""
    institution_type: str = ""
    institution_name: str = ""
    social_credit_code: str = ""
    form_code: str = ""
    form_codes: tuple[str, ...] = ()
    period: str = ""
    indicator_code: str = ""
    indicator_name: str = ""
    rule_id: str = ""
    rule_description: str = ""
    external_indicator_name: str = ""
    current_value: Any = ""
    previous_value: Any = ""
    external_value: Any = ""
    difference_value: Any = ""
    change_rate: float | None = None
    value_type: str = ""
    band: str = ""
    description: str = ""
    calculation_trace: str = ""
    source_refs: tuple[SourceRef, ...] = ()
    related_records: tuple[RelatedRecord, ...] = ()
    # S2-F01 输出层使用的明确语义字段。它们由读取/匹配/规则引擎产生，
    # exporter 只负责展示，不能从空值反向猜测缺失原因。
    retrieval_note: str = ""
    rule_type: str = ""
    value_details: str = ""


@dataclass(frozen=True)
class AuditSummary:
    institutions: int
    forms: int
    indicators: int
    total_findings: int
    abnormal_findings: int
    by_type: dict[str, int]
    by_severity: dict[str, int]


@dataclass(frozen=True)
class V2RunPaths:
    result: Path
