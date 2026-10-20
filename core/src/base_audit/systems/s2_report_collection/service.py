"""V2 编排层：仅连接导入、校验、引擎和导出器。"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable

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

    def run(
        self, *, current_dir: Path, previous_dir: Path, config_path: Path,
        output_dir: Path, central_path: Path | None = None,
        on_progress: Callable[[dict[str, object]], None] | None = None,
    ) -> V2RunResult:
        def progress(stage: str, label: str, percent: float, current: int = 0, total: int = 0, message: str = "") -> None:
            if on_progress:
                on_progress({"stage": stage, "label": label, "percent": percent, "current": current, "total": total, "message": message})

        progress("config", "检查审核配置", 0, message=f"读取配置文件：{config_path.name}")
        config_report = check_s2_config(config_path)
        if not config_report["passed"]:
            details = "；".join(item["message"] for item in config_report["issues"] if item["level"] == "blocking")
            raise ValueError(f"审核配置存在阻断错误，不能执行报表规则：{details}")
        config = load_config(config_path)
        progress("current", "读取本期报表", 5, message=f"本期目录：{current_dir}")
        current_import = import_period_directory(
            current_dir, config=config, label="本期",
            on_progress=lambda current, total, message: progress("current", "读取本期报表", 5 + 25 * current / total, current, total, message),
        )
        progress("previous", "读取上期报表", 30, message=f"上期目录：{previous_dir}")
        previous_import = import_period_directory(
            previous_dir, config=config, label="上期",
            on_progress=lambda current, total, message: progress("previous", "读取上期报表", 30 + 25 * current / total, current, total, message),
        )
        progress("validate", "检查导入数据", 55, message="正在整理本期与上期指标")
        current, previous = current_import.by_logical_key(), previous_import.by_logical_key()
        quality = [*validate_config_warnings(config.warnings), *validate_import(current_import), *validate_import(previous_import)]
        progress(
            "validate", "检查导入数据", 59,
            message=f"完成：本期 {len(current)} 项指标，上期 {len(previous)} 项指标，{len(quality)} 条导入/配置提示",
        )
        progress(
            "period", "执行环比规则", 60,
            message=f"正在对比本期 {len(current)} 项与上期 {len(previous)} 项指标",
        )
        period_findings = PeriodComparisonEngine().run(current, previous, config)
        progress("period", "执行环比规则", 66, message=f"完成：生成 {len(period_findings)} 条环比结果")
        progress("rules", "执行校验规则", 67, message=f"正在按配置检查本期 {len(current)} 项指标")
        rule_findings, rule_quality = RuleEngine().run(current, previous, config)
        progress(
            "rules", "执行校验规则", 72,
            message=f"完成：生成 {len(rule_findings)} 条校验结果，另有 {len(rule_quality)} 条规则数据提示",
        )
        external_findings: list[AuditFinding] = []
        if central_path is not None:
            progress("external_read", "读取大集中数据", 73, message=f"正在读取：{central_path.name}")
            values, org_map = import_external_workbook(central_path)
            progress(
                "external_read", "读取大集中数据", 76,
                message=f"完成：读取 {len(values)} 条指标数据，机构参照 {len(org_map)} 条",
            )
            progress(
                "external", "执行外部核对", 77,
                message=f"正在用 {len(values)} 条大集中指标数据核对报表",
            )
            external_findings = ExternalComparisonEngine().run(current, values, org_map, config)
            progress("external", "执行外部核对", 79, message=f"完成：生成 {len(external_findings)} 条外部核对结果")
        findings = [*period_findings, *external_findings, *rule_findings, *quality, *rule_quality]
        summary = build_summary(current.values(), findings)
        progress(
            "export", "生成审核结果", 80,
            message=f"共 {len(findings)} 条审核结果，按环比、外部核对、校验规则分类写入",
        )
        paths = ExcelExporter().export(
            output_dir=output_dir, current_records=list(current.values()), previous_records=list(previous.values()), findings=findings,
            hide_unchanged_period_rows=config.hide_unchanged_period_rows,
            hide_matching_external_rows=config.hide_matching_external_rows,
            on_progress=lambda stage, current, total: progress(
                "save" if stage == "save" else "export",
                "保存审核工作簿" if stage == "save" else f"写入{stage}",
                95 if stage == "save" else 80 + (14 * current / total if total else 0),
                current, total,
            ),
            run_info={
                "运行时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "本期目录": str(current_dir), "上期目录": str(previous_dir),
                "比较配置": str(config_path), "大集中文件": str(central_path or "未选择"), "本期指标数": str(len(current)), "上期指标数": str(len(previous)),
                "外部核对规则数": str(sum(1 for rule in config.external_rules if rule.enabled)),
            },
        )
        progress("complete", "审核完成", 100)
        return V2RunResult(tuple(findings), summary, paths)

    def check_config(self, config_path: Path) -> tuple[str, ...]:
        """供设置页/未来 UI 调用的只读配置检查入口。"""
        report = check_s2_config(config_path)
        return tuple(item["message"] for item in report["issues"])
