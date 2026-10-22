from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Any, Mapping

from .excel_precision import excel_round15


@dataclass(frozen=True)
class MetricSpec:
    business: str
    sheet: str
    row_path: tuple[str, ...]
    column_path: tuple[str, ...]
    row_header_titles: tuple[str, ...]
    data_type: str = "number"
    metric_type: str = "amount"
    source_unit: str = "万元"
    standard_unit: str = "万元"
    storage_scale: str = "native"
    allow_blank: bool = False
    precision: int = 2
    evidence: str = ""
    confirmed: bool = True
    layout_version: str = ""


@dataclass(frozen=True)
class ExternalMetricSpec:
    source_id: str
    metric_header: str
    data_type: str = "number"
    metric_type: str = "amount"
    source_unit: str = "万元"
    standard_unit: str = "万元"
    storage_scale: str = "native"
    allow_blank: bool = False
    precision: int = 2
    evidence: str = ""
    confirmed: bool = True


@dataclass(frozen=True)
class SourceContract:
    source_id: str
    source_type: str
    default_sheet: str = ""
    identity_sheet: str = ""
    institution_code_label: str = ""
    institution_name_label: str = ""
    period_label: str = ""
    mapping_sheet: str = ""
    mapping_header_row: int = 1
    mapping_data_start_row: int = 2
    mapping_identity_field: str = ""
    mapping_target_field: str = ""
    data_header_row: int = 1
    data_start_row: int = 2
    data_key_column: int = 1
    metric_header_row: int = 1
    metric_first_column: int = 2
    report_item_label: str = ""
    expected_period: str = ""
    version: str = ""
    row_header_anchor: str = ""
    period_required: bool = False
    institution_required: bool = False


@dataclass(frozen=True)
class ApplicabilityScope:
    scope_id: str
    business: str
    sheet: str
    row_path_prefix: tuple[str, ...]
    description: str = ""
    confirmed: bool = True


@dataclass(frozen=True)
class RuleDefinition:
    rule_id: str
    revision: int
    enabled: bool
    business: str
    default_source: str
    default_sheet: str
    name: str
    severity: str
    check_type: str
    metric: str
    audit_expression: str
    description: str
    audit_expression_inverted: bool = False
    current_expression: str = ""
    comparison_expression: str = ""
    difference_expression: str = ""
    value_dictionary: str = ""
    effective_from: str = ""
    effective_to: str = ""
    applicable_expression: str = ""
    applicability_scope: str = ""
    numeric_mode: str = "decimal"
    formula_cell: str = ""


@dataclass(frozen=True)
class Configuration:
    rules: tuple[RuleDefinition, ...]
    metrics: tuple[MetricSpec, ...] = ()
    external_metrics: tuple[ExternalMetricSpec, ...] = ()
    source_contracts: tuple[SourceContract, ...] = ()
    applicability_scopes: tuple[ApplicabilityScope, ...] = ()
    version: str = ""
    version_status: str = ""
    effective_from: str = ""
    effective_to: str = ""
    release_date: str = ""
    release_note: str = ""
    sha256: str = ""
    path: Path | None = None


@dataclass(frozen=True)
class SourceBinding:
    source_id: str
    path: Path
    business: str = ""
    period: str = ""
    institution_code: str = ""
    institution_name: str = ""


@dataclass
class RuntimeContext:
    period: str
    institution_code: str
    institution_name: str
    current_sources: Mapping[str, SourceBinding]
    previous_sources: Mapping[str, SourceBinding] = field(default_factory=dict)
    external_sources: Mapping[str, SourceBinding] = field(default_factory=dict)
    default_external_source: str = ""
    allow_external_time_series: bool = False


@dataclass(frozen=True)
class Evidence:
    source_id: str
    source_path: str
    source_sha256: str
    sheet: str
    cell: str
    raw_value: Any
    raw_type: str
    source_unit: str = ""
    standard_unit: str = ""
    conversion: str = ""
    mapping_sheet: str = ""
    mapping_cell: str = ""
    mapping_key: str = ""
    report_item: str = ""
    metric_header: str = ""
    source_period: str = ""
    source_contract_version: str = ""
    metric_layout_version: str = ""
    metric_definition_evidence: str = ""
    mapping_method: str = ""
    institution_code: str = ""
    precision: int | None = None


@dataclass(frozen=True)
class EvalValue:
    value: Any
    kind: str
    unit: str | None = None
    metric_type: str = ""
    evidence: tuple[Evidence, ...] = ()
    is_literal: bool = False
    precision: int | None = None
    numeric_mode: str = "decimal"

    @classmethod
    def number(cls, value: Decimal | int | float, *, unit: str | None = None,
               metric_type: str = "", evidence: tuple[Evidence, ...] = (),
               is_literal: bool = False, precision: int | None = None,
               numeric_mode: str = "decimal") -> "EvalValue":
        stored = excel_round15(float(value)) if numeric_mode == "excel_binary64" else Decimal(str(value))
        return cls(stored, "number", unit, metric_type, evidence, is_literal, precision, numeric_mode)

    @classmethod
    def string(cls, value: str, *, is_literal: bool = False) -> "EvalValue":
        return cls(value, "string", None, "text", (), is_literal)

    @classmethod
    def boolean(cls, value: bool, *, evidence: tuple[Evidence, ...] = (),
                is_literal: bool = False) -> "EvalValue":
        return cls(value, "boolean", None, "boolean", evidence, is_literal)


@dataclass
class RuleResult:
    rule_id: str
    revision: int
    business: str
    period: str
    institution_code: str
    institution_name: str
    status: str
    triggered: bool | None
    severity: str
    check_type: str
    metric: str
    description: str = ""
    current_value: Any = None
    comparison_value: Any = None
    difference_value: Any = None
    error_code: str = ""
    error_message: str = ""
    variables: dict[str, Any] = field(default_factory=dict)
    evidence: list[Evidence] = field(default_factory=list)
    retrieval_trace: list[dict[str, Any]] = field(default_factory=list)
    configuration_version: str = ""
    run_id: str = ""
    rule_name: str = ""
    scope_path: str = ""
    current_unit: str = ""
    comparison_unit: str = ""
    difference_unit: str = ""
    current_precision: int | None = None
    comparison_precision: int | None = None
    difference_precision: int | None = None
    unused_variables: list[str] = field(default_factory=list)
    error_evidence: dict[str, Any] = field(default_factory=dict)
    sheet_name: str = ""
    formula_cell: str = ""


@dataclass
class AuditRunResult:
    run_id: str
    period: str
    configuration_version: str
    configuration_sha256: str
    status: str
    results: list[RuleResult]
    ledger: dict[str, Any]
    output_files: dict[str, str] = field(default_factory=dict)
    conditional_format_results: list[dict[str, Any]] = field(default_factory=list)
