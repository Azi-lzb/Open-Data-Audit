"""外部核对引擎：只按已加载的外部核对规则比较标准化记录。"""

from __future__ import annotations

import hashlib
from collections import defaultdict
from decimal import Decimal

from .importer import ExternalValue
from .models import AuditFinding, ExternalCheckRule, IndicatorRecord, ReportAuditConfig, STATUS_ABNORMAL, STATUS_NORMAL
from .unit_conversion import AMOUNT_DATA_TYPES, as_decimal, convert_amount


def _id(*parts: str) -> str:
    return hashlib.sha1("|".join(parts).encode()).hexdigest()[:16]


class ExternalComparisonEngine:
    """不读取 Excel、不接触 UI，只执行 ``IndicatorRecord × ExternalCheckRule`` 比较。"""

    def run(
        self,
        current: dict[tuple[str, str, str], IndicatorRecord],
        external: dict[tuple[str, str], ExternalValue],
        external_org_map: dict[str, str],
        config: ReportAuditConfig,
    ) -> list[AuditFinding]:
        rules_by_indicator: dict[str, list[ExternalCheckRule]] = defaultdict(list)
        for rule in config.external_rules:
            if rule.enabled:
                rules_by_indicator[rule.indicator_code].append(rule)

        findings: list[AuditFinding] = []
        for record in sorted(current.values(), key=lambda r: (r.institution_id, r.indicator_code)):
            if record.value_type == "文字" or not isinstance(record.normalized_value, (int, float, Decimal)):
                continue
            indicator = config.indicators.get(record.indicator_code)
            is_amount = bool(indicator and indicator.data_type in AMOUNT_DATA_TYPES)
            profile = config.institutions_by_id.get(record.institution_id) or config.institutions_by_name.get(record.institution_name)
            row_label = external_org_map.get(record.social_credit_code or record.institution_id) or (profile.report_item if profile else "")
            for rule in rules_by_indicator.get(record.indicator_code, ()):
                external_item = external.get((row_label, rule.external_indicator_name)) if row_label else None
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
                    description = "未找到大集中系统数据"
                    difference, trace = "", "未比较：缺少机构映射或外部系统数值"
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
                    unit_text = f" {config.unit_settings.output_file_unit}" if is_amount else ""
                    symbol = {"等于": "=", "不等于": "≠", "大于": ">", "大于等于": "≥", "小于": "<", "小于等于": "≤"}[method]
                    if method == "等于":
                        trace = (
                            f"abs({current_value} - {external_value}) = {difference}{unit_text} "
                            f"<= 容差 {tolerance}{unit_text}：{'成立' if matched else '不成立'}"
                        )
                    else:
                        trace = (
                            f"{current_value} {symbol} {external_value}：{'成立' if matched else '不成立'}；"
                            f"绝对差异 = {difference}{unit_text}"
                        )
                    retrieval_note = ""
                findings.append(AuditFinding(
                    finding_id=_id("外部核对", rule.rule_id, record.institution_id, record.indicator_code, record.period),
                    audit_type="外部核对", status=status, severity=severity,
                    region=record.region, handling_branch=record.handling_branch, institution_type=record.institution_type,
                    institution_name=record.institution_name, social_credit_code=record.social_credit_code,
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
