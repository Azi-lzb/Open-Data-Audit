"""外部核对引擎：只按已加载的外部核对规则比较标准化记录。"""

from __future__ import annotations

import hashlib
from collections import defaultdict
from decimal import Decimal

from .importer import (
    MISSING_INSTITUTION_LABEL_FALLBACK_KEY,
    MISSING_INSTITUTION_LABEL_FALLBACK_REASON,
    ExternalValue,
    ExternalValueKey,
    ExternalValues,
    normalize_external_date,
)
from .models import AuditFinding, ExternalCheckRule, IndicatorRecord, ReportAuditConfig, STATUS_ABNORMAL, STATUS_NORMAL
from .unit_conversion import AMOUNT_DATA_TYPES, as_decimal, convert_amount


def _id(*parts: str) -> str:
    return hashlib.sha1("|".join(parts).encode()).hexdigest()[:16]


class ExternalComparisonEngine:
    """不读取 Excel、不接触 UI，只执行 ``IndicatorRecord × ExternalCheckRule`` 比较。"""

    def run(
        self,
        current: dict[tuple[str, str, str], IndicatorRecord],
        external: dict[ExternalValueKey, ExternalValue],
        external_org_map: dict[str, str],
        config: ReportAuditConfig,
    ) -> list[AuditFinding]:
        rules_by_indicator: dict[str, list[ExternalCheckRule]] = defaultdict(list)
        for rule in config.external_rules:
            if rule.enabled:
                rules_by_indicator[rule.indicator_code].append(rule)

        findings: list[AuditFinding] = []
        time_series = isinstance(external, ExternalValues) and external.time_series
        missing_label_fallback = external.get(MISSING_INSTITUTION_LABEL_FALLBACK_KEY)
        if (
            time_series
            or (
                missing_label_fallback is not None
                and missing_label_fallback.fallback_reason != MISSING_INSTITUTION_LABEL_FALLBACK_REASON
            )
        ):
            missing_label_fallback = None
        for record in sorted(current.values(), key=lambda r: (r.institution_id, r.indicator_code)):
            if record.value_type == "文字" or not isinstance(record.normalized_value, (int, float, Decimal)):
                continue
            indicator = config.indicators.get(record.indicator_code)
            is_amount = bool(indicator and indicator.data_type in AMOUNT_DATA_TYPES)
            profile = config.institutions_by_id.get(record.institution_id) or config.institutions_by_name.get(record.institution_name)
            row_label = external_org_map.get(record.social_credit_code or record.institution_id) or (profile.report_item if profile else "")
            data_date = normalize_external_date(record.source_date) if time_series else ""
            for rule in rules_by_indicator.get(record.indicator_code, ()):
                if row_label:
                    missing_description = "大集中系统数据未找到该报表项目，请先完善大集中数据的报表定义"
                    missing_reason = f"报表项目“{row_label}”、指标“{rule.external_indicator_name}”无可读取数值"
                else:
                    missing_description = "请完善报表采集系统配置或外部文件中的机构参照映射"
                    missing_reason = "缺少机构映射或报表项目"
                if time_series:
                    external_item = external.get((data_date, row_label, rule.external_indicator_name)) if data_date and row_label else None
                    if not data_date:
                        missing_description = "本期报送日期无效，无法匹配大集中时序数据"
                        missing_reason = f"本期报送日期“{record.source_date}”不是完整日期"
                    elif data_date not in external.available_dates:
                        missing_description = f"未找到本期日期（{data_date}）的大集中系统数据"
                        missing_reason = f"外部时序数据没有数据日期 {data_date}，不借用其他日期"
                    elif not row_label:
                        missing_reason = "缺少机构映射或报表项目"
                    elif (data_date, row_label) not in external.date_labels:
                        missing_reason = f"数据日期 {data_date} 未找到报表项目“{row_label}”"
                    else:
                        missing_reason = f"数据日期 {data_date}、报表项目“{row_label}”、指标“{rule.external_indicator_name}”无可读取数值"
                else:
                    external_item = external.get((row_label, rule.external_indicator_name)) if row_label else None
                used_missing_label_fallback = external_item is None and missing_label_fallback is not None
                if used_missing_label_fallback:
                    external_item = missing_label_fallback
                current_value = as_decimal(record.normalized_value)
                if external_item is None:
                    external_value = None
                elif is_amount:
                    try:
                        external_value = convert_amount(
                            external_item.value,
                            config.unit_settings.external_file_unit,
                            config.unit_settings.output_file_unit,
                        )
                    except ValueError as exc:
                        raise ValueError(
                            f"外部核对规则 {rule.rule_id} 的金额单位换算失败：{exc}"
                        ) from exc
                else:
                    # 百分数、计数等不换算单位，但仍用 Decimal 做数值比较，
                    # 避免导入时先经过二进制浮点后再计算差异。
                    external_value = as_decimal(external_item.value)
                if external_value is None:
                    status, severity = STATUS_ABNORMAL, rule.severity
                    description = missing_description
                    difference, trace = "", f"未比较：{missing_reason}"
                    retrieval_note = "外部未匹配"
                else:
                    difference = abs(current_value - external_value)
                    # 左值始终是报表指标，右值始终是大集中指标。
                    # 仅“等于”使用容差；余额/累发及其容差都已处于输出文件单位。
                    method = rule.comparison_method
                    tolerance = as_decimal(rule.tolerance) if method == "等于" else Decimal(0)
                    if method == "等于":
                        matched = difference <= tolerance
                    elif method == "不等于":
                        matched = current_value != external_value
                    elif method == "大于":
                        matched = current_value > external_value
                    elif method == "大于等于":
                        matched = current_value >= external_value
                    elif method == "小于":
                        matched = current_value < external_value
                    elif method == "小于等于":
                        matched = current_value <= external_value
                    else:
                        raise ValueError(f"外部核对规则 {rule.rule_id} 的比较方式无效：{method}")
                    status = STATUS_NORMAL if matched else STATUS_ABNORMAL
                    severity = rule.severity if status == STATUS_ABNORMAL else ""
                    if method == "等于":
                        description = "与大集中系统数据一致" if matched else "与大集中系统数据不一致"
                    else:
                        description = "满足外部核对规则" if matched else "不满足外部核对规则"
                    if used_missing_label_fallback:
                        description = f"{MISSING_INSTITUTION_LABEL_FALLBACK_REASON}；{description}"
                    unit_text = f" {config.unit_settings.output_file_unit}" if is_amount else ""
                    trace = (
                        f"差异绝对值= abs({current_value:.2f} - {external_value:.2f}) = "
                        f"{difference:.2f}{unit_text}"
                    )
                    retrieval_note = MISSING_INSTITUTION_LABEL_FALLBACK_REASON if used_missing_label_fallback else ""
                findings.append(AuditFinding(
                    finding_id=_id("外部核对", rule.rule_id, record.institution_id, record.indicator_code, record.period),
                    audit_type="外部核对", status=status, severity=severity,
                    region=record.region, handling_branch=record.handling_branch, institution_type=record.institution_type,
                    institution_name=record.institution_name, social_credit_code=record.social_credit_code,
                    report_item=row_label,
                    form_code=record.form_code, form_codes=(record.form_code,), period=record.period,
                    indicator_code=record.indicator_code, indicator_name=record.indicator_name,
                    rule_id=rule.rule_id, rule_description=description,
                    external_indicator_name=rule.external_indicator_name,
                    current_value=current_value,
                    external_value=external_value, difference_value=difference,
                    description=description, calculation_trace=trace,
                    retrieval_note=retrieval_note,
                    source_refs=(record.source, external_item.source) if external_item else (record.source,),
                ))
        return findings
