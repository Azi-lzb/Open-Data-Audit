"""V2 编排层：仅连接导入、校验、引擎和导出器。"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .config import load_config
from .config_check import check_s2_config
from .exporter import ExcelExporter, build_summary
from .external_engine import ExternalComparisonEngine
from .importer import import_external_workbook, import_period_directory
from .models import AuditFinding, AuditSummary, V2RunPaths
from .period_engine import PeriodComparisonEngine
from .rule_engine import RuleEngine
from .validation import validate_config_warnings, validate_import


@dataclass(frozen=True)
class V2RunResult:
    findings: tuple[AuditFinding, ...]
    summary: AuditSummary
    paths: V2RunPaths


class PeriodAuditV2Service:
    """报表采集 V2 的唯一入口；V1 ``period_compare.py`` 保持不变。"""

    def run(self, *, current_dir: Path, previous_dir: Path, config_path: Path, output_dir: Path, central_path: Path | None = None) -> V2RunResult:
        config_report = check_s2_config(config_path)
        if not config_report["passed"]:
            details = "；".join(item["message"] for item in config_report["issues"] if item["level"] == "blocking")
            raise ValueError(f"审核配置存在阻断错误，不能执行报表规则：{details}")
        config = load_config(config_path)
        current_import = import_period_directory(current_dir, config=config, label="本期")
        previous_import = import_period_directory(previous_dir, config=config, label="上期")
        current, previous = current_import.by_logical_key(), previous_import.by_logical_key()
        quality = [*validate_config_warnings(config.warnings), *validate_import(current_import), *validate_import(previous_import)]
        period_findings = PeriodComparisonEngine().run(current, previous, config)
        rule_findings, rule_quality = RuleEngine().run(current, previous, config)
        external_findings: list[AuditFinding] = []
        if central_path is not None:
            values, org_map = import_external_workbook(central_path)
            external_findings = ExternalComparisonEngine().run(current, values, org_map, config)
        findings = [*period_findings, *external_findings, *rule_findings, *quality, *rule_quality]
        summary = build_summary(current.values(), findings)
        paths = ExcelExporter().export(
            output_dir=output_dir, current_records=list(current.values()), previous_records=list(previous.values()), findings=findings,
            hide_unchanged_period_rows=config.hide_unchanged_period_rows,
            hide_matching_external_rows=config.hide_matching_external_rows,
            run_info={
                "运行时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "本期目录": str(current_dir), "上期目录": str(previous_dir),
                "比较配置": str(config_path), "大集中文件": str(central_path or "未选择"), "本期指标数": str(len(current)), "上期指标数": str(len(previous)),
                "外部核对规则数": str(sum(1 for rule in config.external_rules if rule.enabled)),
            },
        )
        return V2RunResult(tuple(findings), summary, paths)

    def check_config(self, config_path: Path) -> tuple[str, ...]:
        """供设置页/未来 UI 调用的只读配置检查入口。"""
        report = check_s2_config(config_path)
        return tuple(item["message"] for item in report["issues"])
