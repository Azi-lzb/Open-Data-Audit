"""两期环比审核引擎。该模块不读取或写入 Excel。"""

from __future__ import annotations

import hashlib
from decimal import Decimal
from typing import Any

from .models import AuditFinding, IndicatorRecord, PeriodBandConfig, ReportAuditConfig, STATUS_ABNORMAL, STATUS_NORMAL
from .unit_conversion import AMOUNT_DATA_TYPES, as_decimal


class PeriodRuleExecutionError(ValueError):
    """环比策略无法唯一确定时的显式运行错误。"""


def _id(*parts: str) -> str:
    return hashlib.sha1("|".join(parts).encode()).hexdigest()[:16]


def _numeric(value: Any) -> float | Decimal | None:
    if isinstance(value, bool) or value in (None, ""):
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return None


def _base(record: IndicatorRecord, *, status: str, severity: str = "", description: str = "", **values: Any) -> AuditFinding:
    source_refs = values.pop("source_refs", (record.source,))
    return AuditFinding(
        finding_id=_id("环比", record.institution_id, record.indicator_code, record.period),
        audit_type="环比", status=status, severity=severity,
        region=record.region, handling_branch=record.handling_branch, institution_type=record.institution_type,
        institution_name=record.institution_name, social_credit_code=record.social_credit_code,
        form_code=record.form_code, form_codes=(record.form_code,) if record.form_code else (),
        period=record.period, indicator_code=record.indicator_code, indicator_name=record.indicator_name,
        description=description, source_refs=source_refs, **values,
    )


def _retrieval_note(record: IndicatorRecord | None, label: str) -> str:
    """Return an explicit source-value state for a period side."""
    if record is None:
        return f"{label}缺少指标"
    if record.normalized_value in (None, ""):
        return f"{label}无数据"
    if record.value_type != "文字" and _numeric(record.normalized_value) is None:
        return f"{label}取值非数值"
    return ""


def _join_notes(*notes: str) -> str:
    return "；".join(dict.fromkeys(note for note in notes if note))


def _indicator_config(config: ReportAuditConfig, record: IndicatorRecord):
    return config.indicators.get(record.indicator_code)


def _strategy_bands(config: ReportAuditConfig, record: IndicatorRecord) -> tuple[str, list[PeriodBandConfig]]:
    indicator = _indicator_config(config, record)
    group = indicator.period_strategy_group.strip() if indicator else ""
    if not group or group == "不适用":
        # 保持直接构造旧 ReportAuditConfig 的测试/调用兼容性；正式新配置的
        # 数值指标必须有策略组，缺失会在配置检查阶段阻断。
        return group, [band for band in config.bands if not band.disabled and not band.strategy_group]
    data_type = indicator.data_type or record.value_type
    return group, [
        band for band in config.bands
        if not band.disabled and band.strategy_group == group and (not band.data_type or band.data_type == data_type)
    ]


def _match_strategy(
    config: ReportAuditConfig,
    record: IndicatorRecord,
    percent_value: float,
    *,
    legacy_decimal_value: float | None = None,
) -> tuple[str, PeriodBandConfig | None]:
    """按配置口径匹配策略。

    新“环比策略”统一使用百分数数值：30 表示 30%，单元格不使用百分比
    显示格式。旧“环比规则”及直接构造的无策略组配置继续使用小数口径，
    保留历史兼容性。
    """
    group, candidates = _strategy_bands(config, record)
    value = percent_value if group else (
        legacy_decimal_value if legacy_decimal_value is not None else percent_value
    )
    matches = [band for band in candidates if band.includes(value)]
    if len(matches) > 1:
        raise PeriodRuleExecutionError(
            f"指标 {record.indicator_code} 环比策略组 {group or '旧配置'} 在变动值 {value} 上匹配多个规则："
            f"{'、'.join(band.rule_id for band in matches)}"
        )
    if group and len(matches) == 0:
        raise PeriodRuleExecutionError(
            f"指标 {record.indicator_code} 环比策略组 {group} 在变动值 {value} 上没有唯一匹配规则"
        )
    return group, matches[0] if matches else None


def _min_change(config: ReportAuditConfig, record: IndicatorRecord) -> float | Decimal | None:
    indicator = _indicator_config(config, record)
    if indicator is None or indicator.data_type == "文字" or indicator.min_change_value is None:
        return None
    if indicator.data_type in AMOUNT_DATA_TYPES:
        # 金额门槛与源数据一样直接按输出文件单位解释，不再套另一套单位。
        return as_decimal(indicator.min_change_value)
    return indicator.min_change_value


def _gate_passed(abs_change: float, minimum: float | None) -> bool:
    # 0 的正式语义是关闭第二道绝对变动门槛，而不是重新执行 >= 0。
    return minimum is None or minimum == 0 or abs_change >= minimum


def _trace_number(value: float | Decimal) -> str:
    if isinstance(value, Decimal):
        return format(value, "f")
    return f"{value:.2f}"


def _strategy_rule_label(group: str, band: PeriodBandConfig | None) -> str:
    """返回结果表中可唯一定位策略的规则编号。"""
    if band is None:
        return ""
    return f"{group}-{band.rule_id}" if group else band.rule_id


def _strategy_description(
    band: PeriodBandConfig | None,
    *,
    abnormal: bool,
    gate_passed: bool,
) -> str:
    """生成适合筛选的稳定审核说明，不拼接每行实际变动值。"""
    if band and band.description:
        return band.description
    if abnormal and band:
        return "命中规则"
    if band and not gate_passed and band.severity:
        return "未达到最小变动值"
    return "未命中规则"


class PeriodComparisonEngine:
    """以 ``机构代码 + 指标编码`` 对齐两期；form_code 用于一致性校验与呈现。"""

    def run(self, current: dict[tuple[str, str, str], IndicatorRecord], previous: dict[tuple[str, str, str], IndicatorRecord], config: ReportAuditConfig) -> list[AuditFinding]:
        current_by = {(r.institution_id, r.indicator_code): r for r in current.values()}
        previous_by = {(r.institution_id, r.indicator_code): r for r in previous.values()}
        findings: list[AuditFinding] = []
        for key in sorted(set(current_by) | set(previous_by)):
            cur, pre = current_by.get(key), previous_by.get(key)
            anchor = cur or pre
            assert anchor is not None
            indicator_config = _indicator_config(config, anchor)
            if indicator_config is not None and (indicator_config.disabled or not indicator_config.period_enabled):
                continue
            if cur is None:
                findings.append(_base(anchor, status=STATUS_ABNORMAL, severity="提示", description="本期无，上期有", previous_value=pre.normalized_value,
                                      band="", retrieval_note=_join_notes(_retrieval_note(cur, "本期"), _retrieval_note(pre, "上期"))))
                continue
            if pre is None:
                findings.append(_base(anchor, status=STATUS_ABNORMAL, severity="提示", description="本期有，上期无", current_value=cur.normalized_value,
                                      band="", retrieval_note=_join_notes(_retrieval_note(cur, "本期"), _retrieval_note(pre, "上期"))))
                continue
            if cur.form_code != pre.form_code:
                findings.append(_base(cur, status=STATUS_ABNORMAL, severity="严重", description=f"两期表单不一致：本期 {cur.form_code}，上期 {pre.form_code}", current_value=cur.normalized_value, previous_value=pre.normalized_value,
                                      band="", retrieval_note=_join_notes(_retrieval_note(cur, "本期"), _retrieval_note(pre, "上期")), source_refs=(cur.source, pre.source)))
                continue
            if cur.value_type == "文字" or pre.value_type == "文字":
                changed = str(cur.normalized_value or "") != str(pre.normalized_value or "")
                findings.append(_base(cur, status=STATUS_ABNORMAL if changed else STATUS_NORMAL, severity="提示" if changed else "", description="文字指标变化" if changed else "无变动", current_value=cur.normalized_value, previous_value=pre.normalized_value,
                                      band="", retrieval_note=_join_notes(_retrieval_note(cur, "本期"), _retrieval_note(pre, "上期"), "变动幅度未计算"), source_refs=(cur.source, pre.source)))
                continue
            cn, pn = _numeric(cur.normalized_value), _numeric(pre.normalized_value)
            if cn is None and pn is None:
                # 两侧都无数值：与“完全一致”同口径，状态=正常。
                findings.append(_base(cur, status=STATUS_NORMAL, description="无变动", current_value=cur.normalized_value, previous_value=pre.normalized_value,
                                      band="", retrieval_note=_join_notes(_retrieval_note(cur, "本期"), _retrieval_note(pre, "上期")), source_refs=(cur.source, pre.source)))
                continue
            if cn == pn:
                # 本期与上期完全一致（含两侧均为 0）：统一“无变动”，
                # 不区分数值/百分数，不再进入环比档位。
                findings.append(_base(cur, status=STATUS_NORMAL, description="无变动", current_value=cn, previous_value=pn, difference_value=0.0, change_rate=0.0, source_refs=(cur.source, pre.source)))
                continue
            if cn is None:
                # 单侧存在性异常：明确标识；环比不计算。
                findings.append(_base(cur, status=STATUS_ABNORMAL, severity="提示", description="本期无，上期有", previous_value=pn, band="",
                                      retrieval_note=_join_notes(_retrieval_note(cur, "本期"), _retrieval_note(pre, "上期")), source_refs=(cur.source, pre.source)))
                continue
            if pn is None:
                findings.append(_base(cur, status=STATUS_ABNORMAL, severity="提示", description="本期有，上期无", current_value=cn, band="",
                                      retrieval_note=_join_notes(_retrieval_note(cur, "本期"), _retrieval_note(pre, "上期")), source_refs=(cur.source, pre.source)))
                continue
            configured_type = indicator_config.data_type if indicator_config else ""
            if configured_type == "百分数" or cur.value_type == "百分数" or pre.value_type == "百分数":
                # 百分数指标（值本身即百分数，9.92 表示 9.92%）：差异值就是
                # 变动率（单位：百分点），不计算相对变动 (本期-上期)/上期、
                # 不计算倍率。新环比策略直接用百分点数值匹配（-8.92 个百分点
                # = 配置值 -8.92）；旧无策略组配置仍按 -0.0892 兼容。
                # 中间提示档（新配置 B005(-30,0)/B006(0,30)）
                # 照常命中 → 状态=异常、级别=提示；相对降幅/增幅档
                # （-89.92% → B002 之类）不再产生。「变动幅度」列按百分点
                # 数值输出（exporter 对该行用非百分比格式）。
                difference = cn - pn
                group, band = _match_strategy(
                    config, cur, difference,
                    legacy_decimal_value=difference / (Decimal(100) if isinstance(difference, Decimal) else 100.0),
                )
                minimum = _min_change(config, cur)
                gate_passed = _gate_passed(abs(difference), minimum)
                abnormal = bool(band and (band.description or band.severity) and gate_passed)
                severity = (band.severity or "提示") if abnormal and band else ""
                description = _strategy_description(
                    band, abnormal=abnormal, gate_passed=gate_passed
                )
                rule_label = _strategy_rule_label(group, band)
                trace = (f"{cn:.2f}% - {pn:.2f}% = {difference:.2f} 个百分点"
                         "（百分数按百分点差异匹配环比规则，不适用相对增降幅）")
                if group:
                    min_text = "不适用" if minimum is None else f"{minimum:g}"
                    trace += f"；策略组={group}；命中规则={rule_label or '未命中'}；绝对变动值={abs(difference):.2f}；最小变动值={min_text}；{'门槛通过' if gate_passed else '未达到最小变动值'}"
                findings.append(_base(cur, status=STATUS_ABNORMAL if abnormal else STATUS_NORMAL, severity=severity,
                                      description=description, band=rule_label,
                                      current_value=cn, previous_value=pn,
                                      difference_value=difference, change_rate=difference,
                                      value_type="百分数",
                                      calculation_trace=trace,
                                      source_refs=(cur.source, pre.source)))
                continue
            if pn == 0:
                findings.append(_base(cur, status=STATUS_ABNORMAL, severity="提示", description="本期有，上期无", current_value=cn, previous_value=pn, difference_value=cn,
                                      band="未命中", retrieval_note="变动幅度未计算", source_refs=(cur.source, pre.source)))
                continue
            if cn == 0:
                findings.append(_base(cur, status=STATUS_ABNORMAL, severity="提示", description="本期无，上期有", current_value=cn, previous_value=pn, difference_value=-pn,
                                      band="未命中", retrieval_note="变动幅度未计算", source_refs=(cur.source, pre.source)))
                continue
            difference = cn - pn
            # V2 用 abs(previous) 保留“增减方向”；负数上期不再反向解释趋势。
            change_rate = difference / abs(pn)
            percent_value = change_rate * (Decimal(100) if isinstance(change_rate, Decimal) else 100.0)
            group, band = _match_strategy(
                config, cur, percent_value, legacy_decimal_value=change_rate
            )
            minimum = _min_change(config, cur)
            gate_passed = _gate_passed(abs(difference), minimum)
            abnormal = bool(band and (band.description or band.severity) and gate_passed)
            description = _strategy_description(
                band, abnormal=abnormal, gate_passed=gate_passed
            )
            severity = (band.severity or "提示") if abnormal and band else ""
            rule_label = _strategy_rule_label(group, band)
            trace = (f"({_trace_number(cn)} - {_trace_number(pn)}) / "
                     f"abs({_trace_number(pn)}) = {change_rate:.2%}")
            if group:
                min_text = "不适用" if minimum is None else f"{minimum:g}"
                trace += (f"；策略组={group}；命中规则={rule_label or '未命中'}；"
                          f"绝对变动值={_trace_number(abs(difference))}；最小变动值={min_text}；"
                          f"{'门槛通过' if gate_passed else '未达到最小变动值'}")
            findings.append(_base(cur, status=STATUS_ABNORMAL if abnormal else STATUS_NORMAL, severity=severity, description=description, band=rule_label, current_value=cn, previous_value=pn, difference_value=difference, change_rate=change_rate, value_type=cur.value_type or "数值", calculation_trace=trace, source_refs=(cur.source, pre.source)))
        return findings
