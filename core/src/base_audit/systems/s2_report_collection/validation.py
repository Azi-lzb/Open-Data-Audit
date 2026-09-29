"""导入后校验规则检查（重复指标、机构身份、表单识别等）。"""

from __future__ import annotations

import hashlib
from collections import Counter

from .importer import ImportResult
from .models import AuditFinding, IndicatorRecord, STATUS_ABNORMAL


def _finding_id(*parts: str) -> str:
    return hashlib.sha1("|".join(parts).encode("utf-8")).hexdigest()[:16]


def _quality(record: IndicatorRecord, description: str, *, severity: str = "严重", trace: str = "") -> AuditFinding:
    return AuditFinding(
        finding_id=_finding_id("质量", record.institution_id, record.indicator_code, record.period, description),
        audit_type="校验规则", status=STATUS_ABNORMAL, severity=severity,
        region=record.region, handling_branch=record.handling_branch, institution_type=record.institution_type,
        institution_name=record.institution_name, social_credit_code=record.social_credit_code,
        form_code=record.form_code, form_codes=(record.form_code,) if record.form_code else (), period=record.period,
        indicator_code=record.indicator_code, indicator_name=record.indicator_name,
        rule_id="Q001", rule_description=description, description=description,
        calculation_trace=trace, rule_type="导入校验", source_refs=(record.source,),
    )


def validate_import(result: ImportResult) -> list[AuditFinding]:
    findings: list[AuditFinding] = []
    for left, right in result.duplicates:
        suffix = "；同一指标编码跨表单冲突" if left.form_code != right.form_code else ""
        findings.append(_quality(
            right, "重复指标：同一机构、指标编码、数据期出现多条记录" + suffix,
            trace=f"首条：{left.source.source_file}/{left.source.source_sheet}!第{left.source.source_row}行；"
                  f"重复：{right.source.source_file}/{right.source.source_sheet}!第{right.source.source_row}行",
        ))
    for record in result.records:
        if not record.institution_id:
            findings.append(_quality(record, "缺少稳定机构身份", severity="严重"))
        if not record.form_code:
            findings.append(_quality(record, "未识别表单代码", severity="关注"))
    for text in result.warnings:
        # 文件级导入告警没有可靠单一记录时，保留为系统级质量发现。
        findings.append(AuditFinding(
            finding_id=_finding_id("质量", text), audit_type="校验规则", status=STATUS_ABNORMAL,
            severity="关注", rule_id="Q001", rule_description=text, description=text, rule_type="导入校验",
        ))
    return findings


def validate_config_warnings(warnings: list[str]) -> list[AuditFinding]:
    return [AuditFinding(
        finding_id=_finding_id("配置", warning), audit_type="校验规则", status=STATUS_ABNORMAL,
        severity="关注", rule_id="Q001", rule_description="配置检查：" + warning, description="配置检查：" + warning, rule_type="配置检查",
    ) for warning in warnings]
