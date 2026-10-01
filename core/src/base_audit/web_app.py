from __future__ import annotations

import os
import threading
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

from .app_identity import app_title, icon_path, load_product_info, render_app_page

from .discovery import (
    classify_source_files,
    detect_period,
    explanation_files,
    recommend_template,
)
from .history import HISTORY_AUDIT_SHEET, HISTORY_HEADERS
from .name_config import (
    HISTORY_WORKBOOK_NAME,
    check_summary_config,
    ensure_summary_config_guide,
    initialize_history_workbook,
)
from .node_flow_config import release_config_dir
from .combine_settings import normalize_combine_settings, validate_combine_plans
from .period_compare import (
    CONFIG_WORKBOOK_NAME,
    check_period_config,
    ensure_period_config_guide,
    reset_period_config_workbook,
)
from .systems.s3_central_statistics.config import (
    CENTRAL_CONFIG_ERROR,
    COMMON_CONFIG_NAME,
    COMPARISON_CONFIG_NAME,
    CROSS_CONFIG_NAME,
    FORM_CONFIG_NAME,
    check_central_config_bundle,
    ensure_split_central_configs,
)
from .engines import (
    available_engines,
    available_summary_readers,
    engine_hint,
    pipeline_kind,
    valid_engine_values,
    valid_summary_reader_values,
)
from .service import AuditService
from .systems.s2_report_collection.config_check import check_s2_config, format_config_report
from .path_browser import browse_directory
from .settings import (
    CONDITIONAL_FORMAT_ENGINE_MODES,
    EXPRESSION_EVALUATION_BACKENDS,
    EXPRESSION_EVALUATION_MODES,
    CONDITIONAL_FORMAT_EVALUATORS,
    CONDITIONAL_FORMAT_RULE_READERS,
    EXTERNAL_SHEET_WRITERS,
    FILE_PICKER_MODES,
    FORMULA_REGION_WRITERS,
    DEFAULT_CENTRAL_RULE_ENGINE,
    DEFAULT_EXPRESSION_EVALUATION_BACKEND,
    DEFAULT_EXPRESSION_EVALUATION_MODE,
    DEFAULT_XLSX_RENDER_MODE,
    XLSX_RENDER_MODES,
    SettingsStore,
    UI_THEMES,
    UI_ACCENTS,
    DEFAULT_UI_ACCENT,
    hide_application_data_directory,
)


def _dag_module_payload(definition: object) -> dict[str, object]:
    """Serialize registry-only module metadata for the visual DAG editor."""
    def port(item: object) -> dict[str, object]:
        artifact_type = getattr(item, "artifact_type")
        subtype = getattr(item, "subtype", None)
        accepted = getattr(item, "accepted_subtypes", ())
        return {
            "name": getattr(item, "name"),
            "type": artifact_type.value,
            "subtype": subtype.value if subtype else "",
            "acceptedSubtypes": [value.value for value in accepted],
            "required": bool(getattr(item, "required", False)),
            "cardinality": getattr(item, "cardinality", "ONE").value if hasattr(getattr(item, "cardinality", None), "value") else "ONE",
            "description": getattr(item, "description", ""),
        }
    return {
        "moduleId": getattr(definition, "module_id"),
        "version": getattr(definition, "version"),
        "name": getattr(definition, "name"),
        "category": getattr(definition, "category").value,
        "description": getattr(definition, "description"),
        "capabilities": list(getattr(definition, "platform_capabilities", ())),
        "inputs": [port(item) for item in getattr(definition, "input_ports", ())],
        "outputs": [port(item) for item in getattr(definition, "output_ports", ())],
        "parameters": getattr(definition, "parameter_schema", {}),
    }


def pywebview_file_types(file_types: tuple[tuple[str, str], ...]) -> tuple[str, ...]:
    """Tk 风格 (label, "*.a *.b") 过滤器 → pywebview "label (*.a;*.b)"。

    pywebview 的过滤器校验（webview.util.parse_file_type）只接受分号分隔的
    ``*.ext`` 列表，空格分隔会被判 ``is not a valid file filter``。
    """
    return tuple(f"{label} ({';'.join(patterns.split())})" for label, patterns in file_types)


class WebApi:
    def __init__(self, project_root: Path, *, file_picker_default: str = "系统原生") -> None:
        self.project_root = project_root
        self._file_picker_default = file_picker_default if file_picker_default in FILE_PICKER_MODES else "系统原生"
        # 本机环境摘要（glibc/Python/LibreOffice）：高级设置展示与报障定位用。
        from .system_info import system_environment_summary, system_environment_text
        self._system_environment_text = system_environment_text()
        self._system_environment_info = system_environment_summary()
        self._start_environment_probe()
        hide_application_data_directory(project_root / "data")
        self.settings_store = SettingsStore(project_root / "data" / "用户设置.json")
        self.settings = self.settings_store.load(file_picker_default=self._file_picker_default)
        # 旧布局（如已取消的 unified/config/）被删除后，失效路径会在加载时清空；
        # 回写一次，避免每次启动都重新清理。
        if self.settings_store.sanitized:
            self.settings_store.save(self.settings)
        config_dir = release_config_dir(project_root)
        # 用户给发行目录中的默认配置加了序号前缀。旧名称文件即使仍残留，
        # 也不能让历史绑定继续指向旧副本；清空后统一落到当前默认名称。
        renamed_defaults = {
            "history_config": (("1.笔统计系统_配置.xlsx", "逐笔统计系统_配置.xlsx"), HISTORY_WORKBOOK_NAME),
            "period_config_file": ("报表采集系统_配置.xlsx", CONFIG_WORKBOOK_NAME),
            "central_common_config": ("大集中通用配置.xlsx", COMMON_CONFIG_NAME),
            "central_comparison_config": ("大集中执行比较_配置.xlsx", COMPARISON_CONFIG_NAME),
            "central_cross_config": ("大集中本期数值核对_配置.xlsx", CROSS_CONFIG_NAME),
            "central_forms_config": ("大集中指标比较拆分_配置.xlsx", FORM_CONFIG_NAME),
        }
        rebound = False
        for field, (old_name, new_name) in renamed_defaults.items():
            saved = str(getattr(self.settings, field, "") or "").strip()
            if not saved:
                continue
            saved_path = Path(saved)
            try:
                is_release_default = saved_path.parent.resolve() == config_dir.resolve()
            except OSError:
                is_release_default = False
            old_names = old_name if isinstance(old_name, tuple) else (old_name,)
            if is_release_default and saved_path.name in old_names and (config_dir / new_name).is_file():
                setattr(self.settings, field, str(config_dir / new_name))
                rebound = True
        if rebound:
            self.settings_store.save(self.settings)
        try:
            central_defaults = ensure_split_central_configs(config_dir)
        except (RuntimeError, OSError) as exc:
            # 让界面仍可启动并显示四个预期路径，执行时再给出明确缺失提示。
            central_defaults = {
                "common": config_dir / COMMON_CONFIG_NAME,
                "comparison": config_dir / COMPARISON_CONFIG_NAME,
                "cross": config_dir / CROSS_CONFIG_NAME,
                "forms": config_dir / FORM_CONFIG_NAME,
            }
            self.settings_store.sanitized = True
        # 逐笔统计系统配置：优先用上次选择的路径，否则回填发行目录默认文件，
        # 界面从启动起就显示默认路径，用户不必再手动搜索绑定。
        self.history_path = (
            Path(self.settings.history_config)
            if self.settings.history_config
            else release_config_dir(project_root) / HISTORY_WORKBOOK_NAME
        )
        bundled_templates = project_root / "templates"
        formal_templates = project_root / "2026-07-31" / "模板文件"
        # 模板目录仅在本机真实存在时才预填；否则留空，避免发行包/新环境
        # 显示不存在的死路径。六册配置的默认绑定不受影响。
        default_template_dir = next(
            (str(p) for p in (bundled_templates, formal_templates) if p.is_dir()),
            "",
        )
        saved_input = self.settings.last_input_dir
        saved_output = self.settings.last_output_dir
        self.state: dict[str, Any] = {
            "appInfo": load_product_info(),
            "busy": False,
            "cancelRequested": False,
            "status": "就绪",
            "log": [],
            "input": saved_input,
            "templateDir": self.settings.last_template_dir or default_template_dir,
            "template": "",
            "templateManual": False,
            "external": self.settings.last_external_file,
            # 未绑定自定义配置时回填默认路径；运行语义不变（此前运行时同样回退该文件）。
            "historyConfig": self.settings.history_config or str(self.history_path),
            "output": saved_output,
            "outputAuto": not self.settings.output_pinned,
            "outputPinned": self.settings.output_pinned,
            "recursiveDepth": self.settings.recursive_depth,
            "extraFiles": list(self.settings.extra_files),
            # 运行明细、运行日志 Excel、已有性能信息展示彼此独立。
            "showRunDetailLogs": self.settings.show_run_detail_logs,
            "exportRunLogs": self.settings.export_run_logs,
            "showPerformanceDiagnostics": self.settings.show_performance_diagnostics,
            # 旧前端/旧外壳兼容别名：只表示是否导出运行日志 Excel。
            "writeFlowLogs": self.settings.export_run_logs,
            "filePickerMode": self.settings.file_picker_mode,
            # 本机环境摘要（glibc/Python/LibreOffice）：高级设置展示、报障定位。
            "systemEnvironment": self._system_environment_text,
            "systemEnvironmentInfo": self._system_environment_info,
            # 桌面外壳提供系统原生对话框；Flask 可覆盖为 false。
            "filePickerSystemNativeAvailable": True,
            "xlsxRenderMode": self.settings.xlsx_render_mode,
            # V3 已转正：规则引擎不再暴露选择，历史存档值一律按 v3 执行。
            "centralRuleEngine": "v3",
            "expressionEvaluationMode": (
                self.settings.expression_evaluation_mode
                if self.settings.expression_evaluation_mode in EXPRESSION_EVALUATION_MODES
                else "SBE"
            ),
            "expressionEvaluationBackend": (
                self.settings.expression_evaluation_backend
                if self.settings.expression_evaluation_backend in EXPRESSION_EVALUATION_BACKENDS
                else "PYTHON"
            ),
            "centralExpressionSchema": self.settings.central_expression_schema,
            "confirmBeforeRun": self.settings.confirm_before_run,
            # 输出“生效模式”：UOS（native 管线）没有 Excel/WPS COM，固定 OOXML，
            # 设置中心也不提供 Native 选项。
            # UOS（native 管线）没有 Excel/WPS COM，仅提供 PYTHON 求值器。
            "conditionalFormatEvaluator": (
                self.settings.conditional_format_evaluator
                if self.settings.conditional_format_evaluator in CONDITIONAL_FORMAT_EVALUATORS
                and (pipeline_kind("自动") == "com"
                     or self.settings.conditional_format_evaluator == "PYTHON")
                else "PYTHON"
            ),
            "conditionalFormatRuleReader": self.settings.conditional_format_rule_reader,
            "formulaRegionWriter": self.settings.formula_region_writer,
            "externalSheetWriter": self.settings.external_sheet_writer,
            "conditionalFormatNativeAvailable": pipeline_kind("自动") == "com",
            "uiTheme": (
                self.settings.ui_theme
                if self.settings.ui_theme in UI_THEMES
                else "日间"
            ),
            "uiAccent": (
                self.settings.ui_accent
                if self.settings.ui_accent in UI_ACCENTS
                else DEFAULT_UI_ACCENT
            ),
            "uiAccentOptions": list(UI_ACCENTS),
            "calculationEngine": (
                self.settings.calculation_engine
                if self.settings.calculation_engine in valid_engine_values()
                else "自动"
            ),
            "summaryReadEngine": (
                self.settings.summary_read_engine
                if self.settings.summary_read_engine in valid_summary_reader_values()
                else "纯 Python"
            ),
            "engines": available_engines(),
            "summaryReadEngines": available_summary_readers(),
            "engineHint": engine_hint(),
            "sourceFiles": [],
            "selectedFiles": [],
            # 区分首次识别（默认全选）与用户明确取消全部勾选。不能用
            # selectedFiles 是否为空来推断，否则执行前刷新会把空选择还原成全选。
            "selectedFilesInitialized": False,
            "mixedTemplates": [],
            "explanationFiles": [],
            # 报表采集系统路径：独立保存，不和逐笔统计系统的输入/输出混用。
            "pcCurDir": self.settings.period_current_dir,
            "pcPreDir": self.settings.period_previous_dir,
            "pcCentral": self.settings.period_central_file,
            "pcConfig": self.settings.period_config_file or str(release_config_dir(project_root) / CONFIG_WORKBOOK_NAME),
            "pcOutput": self.settings.period_output_dir,
            "pcOutputAuto": self.settings.period_output_auto,
            # 大集中统计系统：一册通用配置 + 三册功能配置。
            "centralCur": self.settings.central_current_csv,
            "centralPre": self.settings.central_previous_csv,
            "centralOutput": (str(Path(self.settings.central_current_csv).parent / "执行结果")
                              if self.settings.central_output_auto and self.settings.central_current_csv
                              else self.settings.central_output_dir),
            "centralOutputAuto": self.settings.central_output_auto,
            "centralCrossCur": self.settings.central_cross_current_csv,
            "centralCrossOutput": self.settings.central_cross_output_dir,
            "centralCrossOutputAuto": self.settings.central_cross_output_auto,
            "centralCompareFile": self.settings.central_compare_file,
            "centralExplainFile": self.settings.central_explanation_file,
            "centralTemplate": self.settings.central_template_file,
            "centralFormOutput": (str(Path(self.settings.central_compare_file).parent / "执行结果")
                                  if self.settings.central_form_output_auto and self.settings.central_compare_file
                                  else self.settings.central_form_output_dir),
            "centralFormOutputAuto": self.settings.central_form_output_auto,
            "centralCommonConfig": self.settings.central_common_config or str(central_defaults["common"]),
            "centralComparisonConfig": self.settings.central_comparison_config or str(central_defaults["comparison"]),
            "centralCrossConfig": self.settings.central_cross_config or str(central_defaults["cross"]),
            "centralFormsConfig": self.settings.central_forms_config or str(central_defaults["forms"]),
            "centralTolerance": self.settings.central_diff_tolerance_yuan,
            "centralExpressionSchema": self.settings.central_expression_schema,
            "periodPairs": [],
            "periodSummary": {},
            "periodConfigCheck": {},
        # 业务系统配置的健康状况：非空表示缺失/损坏，前端提示去设置中心重置。
        "configIssues": {},
        }
        self._cancel_event = threading.Event()
        initialize_history_workbook(self.history_path)
        ensure_summary_config_guide(self.history_path)
        # 报表采集配置缺失时按默认模板生成，保证首次使用即可打开维护。
        period_config = Path(self.state["pcConfig"]) if self.state["pcConfig"] else None
        try:
            if check_period_config(self.history_path, period_config):
                reset_period_config_workbook(self.history_path, period_config)
            ensure_period_config_guide(self.history_path, period_config)
        except (RuntimeError, OSError) as exc:
            self._log(f"报表采集配置初始化提示：{exc}")
        self._refresh_config_issues()
        self._refresh_period_config_check()
        if self.state["pcCurDir"] or self.state["pcPreDir"]:
            self._refresh_period_pairs()

    def _refresh_config_issues(self) -> None:
        """刷新业务配置的缺失/损坏提示，供设置中心显示与一键重置。"""
        summary_config = str(self.state.get("historyConfig") or self.history_path)
        period_config = str(self.state.get("pcConfig") or "")
        common_config = str(self.state.get("centralCommonConfig") or "")
        central_issues = list(dict.fromkeys(filter(None, (
            check_central_config_bundle(common_config, self.state.get("centralComparisonConfig"), "comparison"),
            check_central_config_bundle(common_config, self.state.get("centralCrossConfig"), "cross"),
            check_central_config_bundle(common_config, self.state.get("centralFormsConfig"), "forms"),
        ))))
        period_report = check_s2_config(Path(period_config) if period_config else release_config_dir(self.project_root) / CONFIG_WORKBOOK_NAME)
        self.state["configIssues"] = {
            "summary": check_summary_config(Path(summary_config)),
            "period": check_period_config(self.history_path, Path(period_config) if period_config else None) or (
                "；".join(item["message"] for item in period_report["issues"] if item["level"] == "blocking")
                if not period_report["passed"] else ""
            ),
            "central": "；".join(central_issues),
        }
        self.state["periodConfigCheck"] = period_report
        self.state["configDefaults"] = {
            key: str(path) if path.is_file() else ""
            for key, path in self._config_default_paths().items()
        }

    def _refresh_period_config_check(self) -> dict[str, Any]:
        """刷新 S2 配置只读检查结果；不修改配置文件。"""
        config_path = Path(self.state.get("pcConfig") or release_config_dir(self.project_root) / CONFIG_WORKBOOK_NAME)
        report = check_s2_config(config_path)
        self.state["periodConfigCheck"] = report
        return report

    def check_period_audit_config(self) -> dict[str, Any]:
        """公开给 Flask/pywebview 的 S2 配置检查桥接方法。"""
        return self._refresh_period_config_check()

    def get_state(self) -> dict[str, Any]:
        return self.state

    def refresh_system_environment(self) -> dict[str, Any]:
        """刷新本机环境摘要（glibc/Python/LibreOffice）；高级设置展示用。"""
        self.state["systemEnvironment"] = self._system_environment_text
        self.state["systemEnvironmentInfo"] = self._system_environment_info
        return self.state

    def _start_environment_probe(self) -> None:
        """后台补全 LibreOffice 版本；仅 Linux 且已找到引擎时启动。"""
        if sys.platform == "win32":
            return
        if self._system_environment_info.get("libreoffice") == "未找到":
            return

        def probe() -> None:
            from .system_info import system_environment_summary, system_environment_text
            system_environment_summary(probe_version=True)
            self._system_environment_text = system_environment_text()
            self._system_environment_info = system_environment_summary()
            self.state["systemEnvironment"] = self._system_environment_text
            self.state["systemEnvironmentInfo"] = self._system_environment_info

        threading.Thread(target=probe, name="base-audit-env-probe", daemon=True).start()

    def initialize_config(self) -> dict[str, Any]:
        """兼容旧界面调用；通用 DAG 配置已取消。"""
        self._log("通用 DAG 配置已取消：流程使用已验收的代码默认值")
        self.state["status"] = "无需初始化通用配置"
        return self.state

    def reset_config(self) -> dict[str, Any]:
        """兼容旧界面调用；不再创建或重置 config/config.xlsx。"""
        self._log("通用 DAG 配置已取消，未创建 config.xlsx")
        self.state["status"] = "无需重置通用配置"
        return self.state

    def _backup_config_file(self, target: Path, *, marker: str = "备份") -> Path:
        """Copy a configuration workbook to a timestamped sibling backup folder."""
        import shutil

        backup_dir = target.parent / "备份文件夹"
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = f"{target.stem}_{marker}_{stamp}"
        backup = backup_dir / f"{base}{target.suffix}"
        index = 1
        while backup.exists():
            backup = backup_dir / f"{base}_{index:02d}{target.suffix}"
            index += 1
        shutil.copy2(target, backup)
        return backup

    def _config_default_paths(self) -> dict[str, Path]:
        """Return the independent, user-approved full-workbook defaults."""
        default_dir = release_config_dir(self.project_root) / "默认配置"
        return {
            "summary": default_dir / HISTORY_WORKBOOK_NAME,
            "period": default_dir / CONFIG_WORKBOOK_NAME,
            "common": default_dir / COMMON_CONFIG_NAME,
            "comparison": default_dir / COMPARISON_CONFIG_NAME,
            "cross": default_dir / CROSS_CONFIG_NAME,
            "forms": default_dir / FORM_CONFIG_NAME,
        }

    def _save_config_default(self, source: Path, *, key: str, label: str) -> dict[str, Any]:
        """Save one complete active workbook as the future reset source."""
        import shutil

        source = Path(source)
        if not source.is_file():
            raise RuntimeError(f"无法设为默认：当前绑定的“{source.name}”不存在")
        snapshot = self._config_default_paths()[key]
        snapshot.parent.mkdir(parents=True, exist_ok=True)
        previous_backup = None
        if snapshot.is_file():
            previous_backup = self._backup_config_file(snapshot, marker="旧默认备份")
        shutil.copy2(source, snapshot)
        self._refresh_config_issues()
        self.state["status"] = f"{label}默认配置已保存"
        suffix = f"；原默认已备份至：{previous_backup}" if previous_backup else ""
        self._log(f"已将{label}设为默认：{snapshot}{suffix}")
        return self.state

    def set_summary_config_default(self) -> dict[str, Any]:
        return self._save_config_default(
            Path(self.state.get("historyConfig") or self.history_path),
            key="summary", label="逐笔统计系统配置",
        )

    def set_period_config_default(self) -> dict[str, Any]:
        source = Path(self.state.get("pcConfig") or release_config_dir(self.project_root) / CONFIG_WORKBOOK_NAME)
        return self._save_config_default(source, key="period", label="报表采集系统配置")

    def set_central_config_default(self, profile: str) -> dict[str, Any]:
        _, label, source = self._central_config_profile(profile)
        return self._save_config_default(source, key=profile, label=label)

    def reset_summary_config(self) -> dict[str, Any]:
        """重建/补全逐笔统计系统_配置.xlsx（只补结构，不清空已有人工记录）。

        文件缺失或损坏时另存备份并重建空模板；正常时仅补齐缺失的表头与
        “使用说明”，绝不动已有历史与人工列。
        """
        import shutil

        path = Path(self.state.get("historyConfig") or self.history_path)
        snapshot = self._config_default_paths()["summary"]
        if snapshot.is_file():
            backup = self._backup_config_file(path) if path.is_file() else None
            path.parent.mkdir(parents=True, exist_ok=True)
            try:
                shutil.copy2(snapshot, path)
            except OSError as exc:
                raise RuntimeError(f"无法从默认配置恢复“{path.name}”：{exc}") from exc
            self.history_path = path
            self.state["status"] = "逐笔统计系统配置已从默认快照恢复"
            self._refresh_config_issues()
            self._log(f"已从默认快照恢复逐笔统计系统配置：{path}（当前文件备份：{backup or '无'}）")
            return self.state
        issue = check_summary_config(path)
        if issue and path.is_file():
            try:
                backup = self._backup_config_file(path, marker="损坏备份")
                path.unlink()
                self._log(f"已备份损坏的逐笔统计配置：{backup}")
            except OSError as exc:
                self._log(f"备份损坏的逐笔统计配置失败（继续重建）：{exc}")
            except Exception as exc:
                self._log(f"备份损坏的逐笔统计配置失败（继续重建）：{exc}")
        self.history_path = path
        initialize_history_workbook(path)
        ensure_summary_config_guide(path)
        self._log(f"已重置逐笔统计系统配置结构：{path}（已有历史与人工列保持不变）")
        self.state["status"] = "逐笔统计系统配置已重置"
        self._refresh_config_issues()
        return self.state

    def reset_period_config(self) -> dict[str, Any]:
        """重建报表采集系统_配置.xlsx：重置前自动备份当前文件。"""
        bound = str(self.state.get("pcConfig") or "")
        target = Path(bound) if bound else None
        actual = target if target is not None else (
            release_config_dir(self.project_root) / CONFIG_WORKBOOK_NAME
        )
        snapshot = self._config_default_paths()["period"]
        backup = None
        if actual.is_file():
            try:
                backup = self._backup_config_file(actual)
                self._log(f"已备份原报表采集配置：{backup}")
            except OSError as exc:
                self._log(f"备份报表采集配置失败（继续重置）：{exc}")
        if snapshot.is_file():
            actual.parent.mkdir(parents=True, exist_ok=True)
            try:
                import shutil
                shutil.copy2(snapshot, actual)
            except OSError as exc:
                raise RuntimeError(f"无法从默认配置恢复报表采集配置：{exc}") from exc
            self._log(f"已从默认快照恢复报表采集系统配置：{actual}（当前文件备份：{backup or '无'}）")
        else:
            if actual.is_file():
                try:
                    actual.unlink()
                except OSError as exc:
                    raise RuntimeError(f"无法重置报表采集配置：请先关闭“{actual.name}”后重试") from exc
            reset_period_config_workbook(self.history_path, target)
            self._log(f"已重置报表采集系统配置（出厂结构）：{actual}")
        self.state["status"] = "报表采集系统配置已重置"
        self._refresh_config_issues()
        return self.state

    def reset_general_settings(self) -> dict[str, Any]:
        """Restore only the general settings shown in the settings page.

        Path history, period-comparison and record-level system settings stay
        intact because they share the same user-settings JSON file.
        """
        self.state["calculationEngine"] = "自动"
        self.state["summaryReadEngine"] = "纯 Python"
        self.state["recursiveDepth"] = -1
        self.state["showRunDetailLogs"] = False
        self.state["exportRunLogs"] = False
        self.state["showPerformanceDiagnostics"] = False
        self.state["writeFlowLogs"] = False
        self.state["confirmBeforeRun"] = True
        self.state["uiTheme"] = "日间"
        self.state["filePickerMode"] = self._file_picker_default
        self.state["systemEnvironment"] = self._system_environment_text
        self.state["conditionalFormatEvaluator"] = "PYTHON"
        self.state["conditionalFormatRuleReader"] = "DIRECT_OOXML"
        self.state["formulaRegionWriter"] = "DIRECT_OOXML"
        self.state["externalSheetWriter"] = "DIRECT_OOXML"
        self.state["xlsxRenderMode"] = DEFAULT_XLSX_RENDER_MODE
        self.state["centralRuleEngine"] = DEFAULT_CENTRAL_RULE_ENGINE
        self.state["expressionEvaluationMode"] = DEFAULT_EXPRESSION_EVALUATION_MODE
        self.state["expressionEvaluationBackend"] = DEFAULT_EXPRESSION_EVALUATION_BACKEND
        self.state["centralExpressionSchema"] = "FIVE_SEGMENT_V1"
        self._save_settings()
        self._log("已恢复常规设置默认值（未修改报表采集系统、逐笔统计系统或历史审核说明）")
        self.state["status"] = "常规设置已恢复默认"
        return self.state

    def get_dag_editor_data(self) -> dict[str, object]:
        """普通编辑器数据：功能目录 + 流程模板 + 已验收 DAG 的业务草稿视图。"""
        from .workflow.catalog import list_workflows
        from .workflow.compiler import business_draft_from_workflow
        from .workflow.functions import build_business_flow_templates, build_function_catalog

        functions = build_function_catalog()
        templates = build_business_flow_templates()
        return {
            "functions": [item.to_dict() for item in functions.values()],
            "flowTemplates": [item.to_dict(functions) for item in templates.values()],
            "workflows": [
                {**item.to_dict(), "businessDraft": business_draft_from_workflow(item)}
                for item in list_workflows(self.history_path)
            ],
            "combineSheetsPlans": self.settings.combine_sheets_plans,
            "activeCombineSheetsPlanId": self.settings.active_combine_sheets_plan_id,
        }

    def validate_config_editor_draft(self, draft: dict[str, Any]) -> dict[str, object]:
        """隐藏设置菜单只允许维护组合工作表方案。"""
        plans = draft.get("combineSheetsPlans") if isinstance(draft, dict) else None
        active = draft.get("activeCombineSheetsPlanId") if isinstance(draft, dict) else None
        errors = validate_combine_plans(plans, active)
        return {"valid": not errors, "errors": errors}

    def save_config_editor_draft(self, draft: dict[str, Any]) -> dict[str, object]:
        if self.state["busy"]:
            return {"ok": False, "errors": ["任务正在运行，暂不能修改组合方案"]}
        checked = self.validate_config_editor_draft(draft)
        if not checked["valid"]:
            return {"ok": False, "errors": checked["errors"]}
        plans, active = normalize_combine_settings(
            draft.get("combineSheetsPlans"), draft.get("activeCombineSheetsPlanId"),
        )
        self.settings.combine_sheets_plans = plans
        self.settings.active_combine_sheets_plan_id = active
        self._save_settings()
        self._log("设置已保存")
        return {
            "ok": True,
            "errors": [],
            "data": {
                "mappingHeaders": [], "modules": [], "defaultFlows": [], "customFlows": [], "flowDisplay": {},
                "combineSheetsPlans": plans, "activeCombineSheetsPlanId": active,
            },
        }

    def save_business_workflow(self, payload: dict[str, Any]) -> dict[str, object]:
        """旧入口保留为回归兼容，但 DAG 自定义图不再落盘。"""
        if self.state["busy"]:
            return {"ok": False, "errors": ["任务正在运行，暂不能修改流程"]}
        if str(payload.get("workflowId") or "").startswith("custom:"):
            return {"ok": False, "errors": ["当前版本不保存自定义 DAG；固定流程由程序维护"]}
        from .workflow.catalog import save_custom_workflow
        from .workflow.compiler import BusinessWorkflowDraft, compile_business_workflow
        from .workflow.runner import build_registry
        from .workflow.validation import validate_workflow

        try:
            definition = compile_business_workflow(BusinessWorkflowDraft.from_dict(payload))
            capabilities = ("com",) if sys.platform.startswith("win") else ("native",)
            errors = validate_workflow(definition, build_registry(), capabilities)
            if errors:
                return {"ok": False, "errors": errors}
            path = save_custom_workflow(self.history_path, definition)
        except Exception as exc:
            return {"ok": False, "errors": [str(exc)]}
        self._log("已保存普通业务流程：{}".format(definition.name))
        return {
            "ok": True,
            "errors": [],
            "path": str(path),
            "workflow": definition.to_dict(),
        }

    def validate_dag_workflow(self, payload: dict[str, Any]) -> dict[str, object]:
        from .workflow.graph import WorkflowDefinition
        from .workflow.runner import build_registry
        from .workflow.validation import validate_workflow
        try:
            definition = WorkflowDefinition.from_dict(payload)
            capabilities = ("com",) if sys.platform.startswith("win") else ("native",)
            errors = validate_workflow(definition, build_registry(), capabilities)
        except Exception as exc:
            errors = [str(exc)]
        return {"valid": not errors, "errors": errors}

    def save_dag_workflow(self, payload: dict[str, Any]) -> dict[str, object]:
        """保留接口形状，但不再保存用户自定义 DAG 图或连线 JSON。"""
        if self.state["busy"]:
            return {"ok": False, "errors": ["任务正在运行，暂不能修改 DAG"]}
        if str(payload.get("workflowId") or "").startswith("custom:"):
            return {"ok": False, "errors": ["当前版本不保存自定义 DAG；固定流程由程序维护"]}
        from .workflow.catalog import save_custom_workflow
        from .workflow.graph import WorkflowDefinition

        checked = self.validate_dag_workflow(payload)
        if not checked["valid"]:
            return {"ok": False, "errors": checked["errors"]}
        try:
            path = save_custom_workflow(self.history_path, WorkflowDefinition.from_dict(payload))
        except Exception as exc:
            return {"ok": False, "errors": [str(exc)]}
        self._log("已保存 DAG 流程：{}".format(path.name))
        return {"ok": True, "errors": [], "path": str(path)}

    def delete_dag_workflow(self, workflow_id: str) -> dict[str, object]:
        """Delete one user-created DAG; built-in workflows are immutable."""
        if self.state["busy"]:
            return {"ok": False, "errors": ["任务正在运行，暂不能删除 DAG"]}
        try:
            from .workflow.catalog import delete_custom_workflow

            delete_custom_workflow(self.history_path, str(workflow_id))
        except Exception as exc:
            return {"ok": False, "errors": [str(exc)]}
        self._log("已删除自定义 DAG 流程：{}".format(workflow_id))
        return {"ok": True, "errors": []}

    def open_config(self) -> dict[str, Any]:
        """Backward-compatible alias for :meth:`open_history_explanation`."""
        return self.open_history_explanation()

    def open_config_file(self) -> dict[str, Any]:
        """兼容旧界面调用；通用 config.xlsx 已退役。"""
        self._log("通用 config.xlsx 已退役；组合方案请在设置中心隐藏菜单维护")
        self.state["status"] = "通用配置文件已取消"
        return self.state

    def open_history_explanation(self) -> dict[str, Any]:
        """Open the Excel workbook that stores historical audit explanations."""
        import os
        path = self.history_path
        if not path.is_file():
            self.state["status"] = "逐笔统计系统配置不存在"
            self._log(f"未找到逐笔统计系统配置：{path}")
            return self.state
        try:
            os.startfile(path)
        except Exception as exc:
            self.state["status"] = "无法打开逐笔统计系统配置"
            self._log(f"打开逐笔统计系统配置失败：{exc}")
            return self.state
        self._log(f"已打开逐笔统计系统配置：{path}")
        return self.state

    def get_history_page(self, query: dict[str, Any] | None = None) -> dict[str, object]:
        """Return one read-only, filtered page of historical audit explanations.

        The worksheet is streamed with openpyxl: large history files never
        become a giant WebView payload.  This endpoint deliberately provides
        no editing operation; formal maintenance remains in Excel.
        """
        from openpyxl import load_workbook

        query = query if isinstance(query, dict) else {}

        def number(name: str, default: int, minimum: int, maximum: int) -> int:
            try:
                value = int(query.get(name, default))
            except (TypeError, ValueError):
                value = default
            return max(minimum, min(maximum, value))

        page = number("page", 1, 1, 1000000)
        page_size = number("pageSize", 50, 10, 100)
        keyword = str(query.get("keyword") or "").strip().casefold()
        error_type = str(query.get("errorType") or "").strip()
        opinion = str(query.get("opinion") or "").strip()
        path = self.history_path
        if not path.is_file():
            return {
                "columns": list(HISTORY_HEADERS), "items": [], "total": 0,
                "page": page, "pageSize": page_size, "totalPages": 0,
                "errorTypes": [], "opinions": [], "error": "逐笔统计系统配置不存在",
            }
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            if HISTORY_AUDIT_SHEET not in workbook.sheetnames:
                raise ValueError(f"逐笔统计系统配置中缺少“{HISTORY_AUDIT_SHEET}”工作表")
            rows = workbook[HISTORY_AUDIT_SHEET].iter_rows(values_only=True)
            headers = [str(value or "").strip() for value in next(rows, ())]
            missing = [name for name in HISTORY_HEADERS if name not in headers]
            if missing:
                raise ValueError("逐笔统计系统配置缺少列：" + "、".join(missing))
            positions = {name: headers.index(name) for name in HISTORY_HEADERS}
            total = 0
            error_types: set[str] = set()
            opinions: set[str] = set()
            items: list[dict[str, str]] = []
            start = (page - 1) * page_size
            end = start + page_size
            for raw in rows:
                item = {
                    name: str(raw[index] if index < len(raw) and raw[index] is not None else "").strip()
                    for name, index in positions.items()
                }
                if item["错误类型"]:
                    error_types.add(item["错误类型"])
                if item["审核意见"]:
                    opinions.add(item["审核意见"])
                haystack = " ".join(item.values()).casefold()
                if keyword and keyword not in haystack:
                    continue
                if error_type and item["错误类型"] != error_type:
                    continue
                if opinion and item["审核意见"] != opinion:
                    continue
                if start <= total < end:
                    items.append(item)
                total += 1
            total_pages = (total + page_size - 1) // page_size
            # A filter can make the requested page out of range.  Reply with
            # the real last page marker; the front end can request it once.
            return {
                "columns": list(HISTORY_HEADERS), "items": items, "total": total,
                "page": page, "pageSize": page_size, "totalPages": total_pages,
                "errorTypes": sorted(error_types), "opinions": sorted(opinions),
            }
        finally:
            workbook.close()

    def open_user_guide(self) -> dict[str, Any]:
        """Open the adjacent Word user guide with the default application."""
        import os
        path = self.project_root / "基础数据审核工具使用说明.docx"
        if not path.is_file():
            self.state["status"] = "使用说明不存在"
            self._log(f"未找到使用说明：{path}")
            return self.state
        try:
            os.startfile(path)
        except Exception as exc:
            self.state["status"] = "无法打开使用说明"
            self._log(f"打开使用说明失败：{exc}")
            return self.state
        self._log(f"已打开使用说明：{path}")
        return self.state

    def open_path(self, path: str) -> dict[str, Any]:
        """Open a local directory or file with the system default application."""
        import os
        target = Path(path)
        if not target.exists():
            self.state["status"] = "路径不存在"
            self._log(f"未找到路径：{path}")
            return self.state
        try:
            os.startfile(target)
        except Exception as exc:
            self.state["status"] = "无法打开路径"
            self._log(f"打开路径失败：{exc}")
            return self.state
        self._log(f"已打开：{path}")
        return self.state

    def browse_dir(self, path: str = "", mode: str = "") -> dict[str, Any]:
        """提供浏览器内置选择器的受控目录清单。"""
        return browse_directory(self.project_root, path, mode)

    def _tk_dialog(self, *, folder: bool, current: str, multiple: bool = False,
                   file_types: tuple[tuple[str, str], ...] = ()) -> tuple[str, ...]:
        """按用户显式设置调用 Tk；缺少 tkinter 时给出可操作的错误。"""
        try:
            import tkinter as tk
            from tkinter import filedialog
        except ImportError as exc:
            raise RuntimeError("未安装 Tk 图形组件；请安装 python3-tk，或在高级设置改用“浏览器内置”") from exc
        root = tk.Tk()
        root.withdraw()
        try:
            root.attributes("-topmost", True)
            if folder:
                value = filedialog.askdirectory(parent=root, initialdir=current, mustexist=True)
                return (str(value),) if value else ()
            if multiple:
                return tuple(str(item) for item in filedialog.askopenfilenames(
                    parent=root, initialdir=current, filetypes=file_types,
                ))
            value = filedialog.askopenfilename(parent=root, initialdir=current, filetypes=file_types)
            return (str(value),) if value else ()
        finally:
            root.destroy()

    def _select_path(self, *, folder: bool, current: str,
                     file_types: tuple[tuple[str, str], ...] = (), multiple: bool = False) -> tuple[str, ...]:
        """按高级设置选择系统原生或 Tk 对话框；浏览器内置模式由前端提交路径。"""
        mode = str(self.state.get("filePickerMode") or "系统原生")
        if mode == "浏览器内置":
            return ()
        if mode == "Tk 对话框":
            return self._tk_dialog(folder=folder, current=current, file_types=file_types, multiple=multiple)
        import webview
        file_dialog = getattr(webview, "FileDialog", None)
        dialog = (
            getattr(file_dialog, "FOLDER" if folder else "OPEN", None)
            if file_dialog is not None
            else getattr(webview, "FOLDER_DIALOG" if folder else "OPEN_DIALOG")
        )
        kwargs: dict[str, Any] = {"directory": current}
        if multiple:
            kwargs["allow_multiple"] = True
        if not folder and file_types:
            kwargs["file_types"] = pywebview_file_types(file_types)
        selected = webview.windows[0].create_file_dialog(dialog, **kwargs)
        return tuple(str(item) for item in (selected or ()))

    def choose_folder(self, field: str, path: str = "") -> str:
        current = str(self.state.get(field) or self.project_root)
        selected = (str(path),) if path else self._select_path(folder=True, current=current)
        value = str(selected[0]) if selected else ""
        if value:
            previous = self.state.get(field, "")
            self.state[field] = value
            if field == "templateDir" and value != previous:
                self.state["templateManual"] = False
            if field in {"input", "templateDir"}:
                if field == "input" and not self.state["outputPinned"]:
                    self.state["output"] = str(Path(value) / "执行结果")
                    self.state["outputAuto"] = True
                if field == "input" and value != previous:
                    # 切换源目录时才对新目录采用首次默认全选；同一目录的
                    # 执行前刷新必须保留用户已取消的选择。
                    self.state["selectedFilesInitialized"] = False
                # 仅通过选择源数据目录触发一次模板推荐；选择模板目录只刷新文件清单。
                # 用户手动选定模板后，仍以手动选择为准。
                self._recognize(allow_template_auto=(field == "input"))
            elif field == "output":
                self.state["outputAuto"] = False
                self.state["outputPinned"] = True
        return value

    def choose_file(self, field: str, path: str = "") -> str:
        current = str(self.state.get(field) or self.project_root)
        selected = (str(path),) if path else self._select_path(
            folder=False, current=str(Path(current).parent if Path(current).is_file() else current),
            file_types=(("Excel 文件", "*.xlsx *.xlsm"),),
        )
        value = str(selected[0]) if selected else ""
        if value:
            self.state[field] = value
            if field == "template":
                self.state["templateManual"] = True
        return value

    def clear_path(self, field: str) -> dict[str, Any]:
        """Clear one selected path and persist that choice to 用户设置.json.

        This only forgets an application setting.  It never removes the
        corresponding directory or workbook from disk.
        """
        fields = {
            "input", "templateDir", "template", "external", "output",
            "historyConfig", "pcCurDir", "pcPreDir", "pcCentral",
            "pcConfig", "pcOutput",
            # 大集中统计系统页的路径与配置绑定，与 CENTRAL_KINDS 一一对应。
            "centralCur", "centralPre", "centralOutput", "centralCrossCur", "centralCrossOutput", "centralCompareFile", "centralExplainFile",
            "centralTemplate", "centralFormOutput", "centralCommonConfig",
            "centralComparisonConfig", "centralCrossConfig", "centralFormsConfig",
        }
        if self.state["busy"] or field not in fields:
            return self.state

        if field == "historyConfig":
            # An empty persisted value means “use the default history workbook”
            # on the next launch, while the UI correctly shows no custom binding.
            self.history_path = release_config_dir(self.project_root) / HISTORY_WORKBOOK_NAME
            self.state[field] = ""
            self.settings.history_config = ""
        else:
            self.state[field] = ""

        if field == "input":
            self.state["sourceFiles"] = []
            self.state["selectedFiles"] = []
            self.state["selectedFilesInitialized"] = False
            self.state["mixedTemplates"] = []
            self.state["explanationFiles"] = []
            self.state["detectedPeriod"] = ""
        elif field == "template":
            self.state["templateManual"] = False
        elif field == "pcOutput":
            # Do not silently restore a path the user explicitly cleared.
            self.state["pcOutputAuto"] = False

        if field in {"pcCurDir", "pcPreDir"}:
            self._refresh_period_pairs()
        if field in {"centralCommonConfig", "centralComparisonConfig", "centralCrossConfig", "centralFormsConfig"}:
            self._refresh_config_issues()
        self._save_settings()
        self._log(f"已清空路径设置：{field}")
        return self.state

    def choose_history_config(self, path: str = "") -> str:
        """选择逐笔统计系统_配置.xlsx 并记住路径（历史数据由人工维护）。"""
        if self.state["busy"]:
            return str(self.history_path)
        current = self.history_path if self.history_path.is_file() else self.project_root
        selected = (str(path),) if path else self._select_path(
            folder=False, current=str(Path(current).parent if Path(current).is_file() else current),
            file_types=(("Excel 文件", "*.xlsx *.xlsm"),),
        )
        value = str(selected[0]) if selected else ""
        if value:
            self.history_path = Path(value)
            self.state["historyConfig"] = value
            self.settings.history_config = value
            self.settings_store.save(self.settings)
            self._log(f"已选择逐笔统计系统配置：{value}")
            self._refresh_config_issues()
        return value

    def create_template_merge(self, base: str = "", picked: list[str] | None = None) -> bool:
        """Interactive, manual-only creation of one combined audit template.

        It deliberately bypasses the workbench form and template recommendation:
        the user selects both the base and every input workbook in dialogs.
        """
        if self.state["busy"]:
            return False
        base_dir = self.state.get("templateDir") or self.project_root
        self.state["status"] = "第 1 步：请选择基准模板"
        self._log(
            "第 1 步/2：请选择基准模板（它会另存为联合模板；"
            "应包含集中系统数据、参照表等公共依赖工作表）"
        )
        selected_base = (str(base),) if base else self._select_path(
            folder=False, current=str(base_dir), file_types=(("Excel 文件", "*.xlsx *.xlsm"),),
        )
        if not selected_base:
            self._log("已取消制作联合模板：未选择底稿模板")
            return False
        base_template = Path(selected_base[0]).resolve()
        self.state["status"] = "第 2 步：请选择要并入的模板"
        self._log(
            f"已选择基准模板：{base_template.name}。第 2 步/2："
            "请批量选择要并入的其他模板（不要重复选择基准模板）"
        )
        selected_sources = tuple(str(item) for item in picked) if picked is not None else self._select_path(
            folder=False, current=str(base_template.parent), multiple=True,
            file_types=(("Excel 文件", "*.xlsx *.xlsm"),),
        )
        if not selected_sources:
            self._log("已取消制作联合模板：未选择待复制工作簿")
            return False
        source_templates = [Path(path).resolve() for path in selected_sources]
        if not [path for path in source_templates if path != base_template]:
            self._log("已取消制作联合模板：待复制工作簿不能只有底稿模板本身")
            return False
        self.state["busy"] = True
        self.state["status"] = "正在制作联合模板，请勿关闭窗口……"
        self._log(f"开始制作联合模板：底稿“{base_template.name}”")
        self._log("提示：请将含外部依赖工作表的模板选作底稿；原始文件不会修改")
        threading.Thread(
            target=self._template_merge_worker,
            args=(base_template, source_templates),
            daemon=True,
        ).start()
        return True

    def _template_merge_worker(
        self, base_template: Path, source_templates: list[Path]
    ) -> None:
        started = time.monotonic()
        try:
            service = AuditService(
                config_path=self.history_path,
                engine_preference=str(self.state["calculationEngine"]),
                combine_sheets_plans=self.settings.combine_sheets_plans,
                active_combine_sheets_plan_id=self.settings.active_combine_sheets_plan_id,
            )
            result = service.merge_template_files(
                base_template=base_template,
                source_templates=source_templates,
                on_step=self._log_detail,
            )
            self.state["status"] = result.summary_text().splitlines()[0]
            self._log(result.summary_text())
        except Exception as exc:
            self.state["status"] = "制作联合模板失败"
            self._log("制作联合模板失败：" + str(exc))
        finally:
            elapsed = time.monotonic() - started
            self._log(f"任务结束，本次耗时：{elapsed:.1f} 秒")
            self.state["busy"] = False

    def append_extra_files(self, paths: list[str] | None = None) -> list[str]:
        """待处理清单：用当前选择器批量追加源数据目录外的工作簿。"""
        if self.state["busy"]:
            return list(self.state.get("extraFiles", []))
        current = str(self.state.get("input") or self.project_root)
        selected = tuple(str(item) for item in paths) if paths is not None else self._select_path(
            folder=False, current=current, multiple=True,
            file_types=(("Excel 文件", "*.xlsx *.xlsm"),),
        )
        if not selected:
            return list(self.state.get("extraFiles", []))
        added = []
        existing = set(self.state.get("extraFiles", []))
        for item in selected:
            path = str(Path(item).resolve())
            if path not in existing and Path(path).is_file():
                existing.add(path)
                added.append(path)
        if added:
            self.state["extraFiles"] = sorted(existing)
            self.settings.extra_files = list(self.state["extraFiles"])
            self.settings_store.save(self.settings)
            self._log(f"已追加 {len(added)} 个待处理文件")
            self._recognize(allow_template_auto=False)
        return list(self.state.get("extraFiles", []))

    def remove_extra_file(self, path: str) -> list[str]:
        """从追加清单移除一个手动追加的文件（不影响源数据目录扫描结果）。"""
        files = [item for item in self.state.get("extraFiles", []) if item != path]
        if len(files) != len(self.state.get("extraFiles", [])):
            self.state["extraFiles"] = files
            self.settings.extra_files = files
            self.settings_store.save(self.settings)
            self._log(f"已移除追加文件：{Path(path).name}")
            self._recognize(allow_template_auto=False)
            selected = [item for item in (self.state.get("selectedFiles") or []) if item != path]
            self.state["selectedFiles"] = selected
        return list(self.state.get("extraFiles", []))

    def _refresh_period_pairs(self) -> None:
        """按机构+表单配对两期已选目录中的文件，供报表采集页展示。"""
        from .period_compare import list_period_pairs

        cur = self.state.get("pcCurDir")
        pre = self.state.get("pcPreDir")
        if not cur and not pre:
            self.state["periodPairs"] = []
            return
        try:
            self.state["periodPairs"] = list_period_pairs(
                Path(cur) if cur else None, Path(pre) if pre else None
            )
        except Exception as exc:
            self.state["periodPairs"] = []
            self._log(f"两期文件配对失败：{exc}")

    # ---- 大集中统计系统 ----

    CENTRAL_FILE_KINDS = {
        "centralCur", "centralPre", "centralCrossCur", "centralCompareFile", "centralExplainFile", "centralTemplate",
        "centralCommonConfig", "centralComparisonConfig", "centralCrossConfig", "centralFormsConfig",
    }
    CENTRAL_KINDS = {
        "centralCur", "centralPre", "centralOutput", "centralCrossCur", "centralCrossOutput", "centralCompareFile", "centralExplainFile",
        "centralTemplate", "centralFormOutput", "centralCommonConfig",
        "centralComparisonConfig", "centralCrossConfig", "centralFormsConfig",
    }

    def choose_central(self, kind: str, path: str = "") -> str:
        """大集中统计系统页：选择导数文件（CSV/XLSX）/比较结果/模板/配置/输出目录；只记录不执行。"""
        if self.state["busy"] or kind not in self.CENTRAL_KINDS:
            return ""
        titles = {
            "centralCur": "选择本期大集中导数文件（CSV/XLSX）",
            "centralPre": "选择上期大集中导数文件（CSV/XLSX）",
            "centralOutput": "选择比较结果输出目录",
            "centralCrossCur": "选择本期数值核对的数据文件（CSV/XLSX）",
            "centralCrossOutput": "选择本期数值核对输出目录",
            "centralCompareFile": "选择比较结果.xlsx",
            "centralExplainFile": "选择说明文件.xlsx",
            "centralTemplate": "选择金融表单模板.xlsx",
            "centralFormOutput": "选择转表输出目录",
            "centralCommonConfig": f"选择{COMMON_CONFIG_NAME}",
            "centralComparisonConfig": f"选择{COMPARISON_CONFIG_NAME}",
            "centralCrossConfig": f"选择{CROSS_CONFIG_NAME}",
            "centralFormsConfig": f"选择{FORM_CONFIG_NAME}",
        }
        current = self.state.get(kind) or self.state.get("input") or str(self.project_root)
        if path:
            selected = (str(path),)
        elif kind in self.CENTRAL_FILE_KINDS:
            selected = self._select_path(
                folder=False, current=str(Path(current).parent) if Path(current).is_file() else str(current),
                file_types=(("数据文件", "*.csv *.xlsx *.xls"),) if kind in {"centralCur", "centralPre", "centralCrossCur"}
                            else (("Excel 文件", "*.xlsx"),),
            )
        else:
            selected = self._select_path(folder=True, current=str(current))
        value = str(selected[0]) if selected else ""
        if value:
            self.state[kind] = str(Path(value).resolve())
            if kind == "centralCrossCur" and self.state.get("centralCrossOutputAuto", True):
                self.state["centralCrossOutput"] = str(Path(value).resolve().parent / "执行结果")
            elif kind == "centralCur" and self.state.get("centralOutputAuto", True):
                self.state["centralOutput"] = str(Path(value).resolve().parent / "执行结果")
            elif kind == "centralCompareFile":
                # 合并后的大集中操作卡共用一个输出目录；手动选择比较结果时也随之带入。
                if self.state.get("centralOutputAuto", True):
                    self.state["centralOutput"] = str(Path(value).resolve().parent)
                if self.state.get("centralFormOutputAuto", True):
                    self.state["centralFormOutput"] = str(Path(value).resolve().parent)
            elif kind == "centralCrossOutput":
                self.state["centralCrossOutputAuto"] = False
            elif kind == "centralOutput":
                self.state["centralOutputAuto"] = False
            elif kind == "centralFormOutput":
                self.state["centralFormOutputAuto"] = False
            label = {
                "centralCur": "本期导数文件", "centralPre": "上期导数文件",
                "centralOutput": "比较输出目录", "centralCompareFile": "比较结果", "centralExplainFile": "说明文件",
                "centralCrossCur": "本期数值核对数据", "centralCrossOutput": "本期数值核对输出目录",
                "centralTemplate": "金融表单模板", "centralFormOutput": "转表输出目录",
                "centralCommonConfig": "大集中通用配置",
                "centralComparisonConfig": "执行比较配置",
                "centralCrossConfig": "本期数值核对配置",
                "centralFormsConfig": "指标比较拆分配置",
            }[kind]
            self._log(f"大集中统计系统：{label}已选择")
            self._save_settings()
            if kind in {"centralCommonConfig", "centralComparisonConfig", "centralCrossConfig", "centralFormsConfig"}:
                self._refresh_config_issues()
        return self.state.get(kind, "")

    def toggle_central_cross_output_pin(self) -> bool:
        if self.state["busy"]:
            return not self.state.get("centralCrossOutputAuto", True)
        auto = not bool(self.state.get("centralCrossOutputAuto", True))
        self.state["centralCrossOutputAuto"] = auto
        if auto and self.state.get("centralCrossCur"):
            self.state["centralCrossOutput"] = str(Path(self.state["centralCrossCur"]).parent / "执行结果")
        self._save_settings()
        return not auto

    def toggle_central_output_pin(self, kind: str) -> bool:
        mapping = {
            "centralOutput": ("centralOutputAuto", "centralCur"),
            "centralFormOutput": ("centralFormOutputAuto", "centralCompareFile"),
        }
        if self.state["busy"] or kind not in mapping:
            return False
        auto_key, source_key = mapping[kind]
        auto = not bool(self.state.get(auto_key, True))
        self.state[auto_key] = auto
        if auto and self.state.get(source_key):
            self.state[kind] = str(Path(self.state[source_key]).parent / "执行结果")
        self._save_settings()
        return not auto

    @staticmethod
    def _central_check_report(label: str, issues: list[dict], stats: dict | None = None) -> str:
        """组装配置检查弹框全文（3.1/3.2 统一格式）；同一文本写入运行记录。

        结构：总数（启用/停用）→【问题】→【取数词表提示】（无提示不显示）。
        每条问题独占一行（序号 + 缩进），保证弹框与运行记录可读、可复制。
        """
        stats = stats or {}
        total = stats.get("total", len(issues))
        enabled = stats.get("enabled")
        header = f"检查完成：{label} 共 {total} 条规则"
        if enabled is not None:
            header += f"（启用 {enabled}，停用 {total - enabled}）"
        lines = [header + "。", ""]
        problem_rows = [item for item in issues if item.get("问题")]
        if problem_rows:
            count = sum(len(item["问题"]) for item in problem_rows)
            lines.append(f"【问题】{count} 处：")
            index = 0
            for item in problem_rows[:50]:
                for entry in item["问题"]:
                    index += 1
                    lines.append(f"  {index}. {item['规则编号']}（{item['启用']}）：{entry}")
        else:
            lines.append("【问题】无")
        vocabulary = stats.get("vocabulary") or []
        if vocabulary:
            lines.append("")
            lines.append(f"【取数词表提示】共 {len(vocabulary)} 处：")
            for index, note in enumerate(vocabulary[:50], 1):
                lines.append(f"  {index}. {note}")
            if len(vocabulary) > 50:
                lines.append(f"  ……另有 {len(vocabulary) - 50} 处略。")
        return chr(10).join(lines)

    @staticmethod
    def _central_preflight_text(label: str, reports: list[tuple[str, dict]], scope: str) -> str:
        """Render read-only S3 workbook checks without conflating warnings with errors."""
        def message(issue: object) -> str:
            if isinstance(issue, dict):
                location = "".join(
                    str(issue[key]) if key == "sheet" else f"第{issue[key]}行"
                    for key in ("sheet", "row") if issue.get(key) is not None
                )
                return f"{location}：{issue.get('message', '')}" if location else str(issue.get("message", ""))
            return str(issue)

        errors = [message(issue) for _, report in reports for issue in report.get("errors", [])]
        warnings = [message(issue) for _, report in reports for issue in report.get("warnings", [])]
        lines = [f"检查完成：{label}，错误 {len(errors)} 处，提示 {len(warnings)} 处。", "", "【本次读取的配置工作簿】"]
        for book_label, report in reports:
            lines.append(f"  {book_label}：{Path(report['path']).name}")
        lines.extend(("", "【实际检查项目】"))
        for book_label, report in reports:
            for check in report.get("checks", []):
                raw_status = check.get("status")
                status = {"passed": "通过", "warning": "提示", "failed": "未通过",
                          "error": "未通过", "not_run": "未检查"}.get(raw_status, "未检查")
                lines.append(f"  [{status}] {book_label}／{check['name']}：{check['detail']}")
            for sheet in report.get("sheet_stats", []):
                sheet_name = sheet.get("sheet") or sheet.get("name")
                if "present" in sheet:
                    status = "存在" if sheet["present"] else "缺失"
                    lines.append(f"  工作表 {sheet_name}：{status}；{sheet.get('scope', '仅检查存在性')}。")
        lines.extend(("", f"【错误】{len(errors)} 处："))
        lines.extend(f"  {index}. {entry}" for index, entry in enumerate(errors[:50], 1))
        if not errors:
            lines.append("  无")
        if len(errors) > 50:
            lines.append(f"  ……另有 {len(errors) - 50} 处略。")
        lines.extend(("", f"【提示】{len(warnings)} 处："))
        lines.extend(f"  {index}. {entry}" for index, entry in enumerate(warnings[:50], 1))
        if not warnings:
            lines.append("  无")
        if len(warnings) > 50:
            lines.append(f"  ……另有 {len(warnings) - 50} 处略。")
        lines.extend(("", "【检查范围】", f"  {scope}"))
        return "\n".join(lines)

    def check_central_common_config(self) -> dict[str, Any]:
        """Read-only preflight for the active 3.0 common workbook."""
        if self.state["busy"]:
            raise RuntimeError("任务正在运行，请完成后再检查配置")
        from .systems.s3_central_statistics.common_config_check import check_common_config

        _, _, path = self._central_config_profile("common")
        report = check_common_config(path)
        report_text = self._central_preflight_text(
            "大集中通用配置（3.0）", [("通用配置", report)],
            "只读检查工作簿、必需工作表、运行参数单位及环比警戒区间；"
            "单位换算例外与机构地区参照只检查工作表是否存在，不检查行内容；不运行真实业务数据。",
        )
        for line in report_text.splitlines():
            self._log_detail(line)
        self.state["status"] = f"配置检查完成：大集中通用配置（3.0），错误 {len(report['errors'])} 处，提示 {len(report['warnings'])} 处"
        self._refresh_config_issues()
        return {**report, "total_issues": len(report["errors"]),
                "source_workbooks": [str(path)], "report_text": report_text}

    def check_central_forms_config(self) -> dict[str, Any]:
        """Read-only preflight for the active 3.3 forms and its common dependency."""
        if self.state["busy"]:
            raise RuntimeError("任务正在运行，请完成后再检查配置")
        from .systems.s3_central_statistics.common_config_check import check_common_config
        from .systems.s3_central_statistics.forms_config_check import check_forms_config

        common_path, forms_path = self._central_config_paths("forms")
        common_report = check_common_config(common_path)
        forms_report = check_forms_config(forms_path)
        reports = [("通用配置", common_report), ("指标比较拆分配置", forms_report)]
        report_text = self._central_preflight_text(
            "指标比较拆分（3.3）", reports,
            "只读检查两册配置的可前置校验项、转表设置和内嵌表单引用；"
            "不逐格检查灵活表单模板，也不运行真实业务数据。",
        )
        for line in report_text.splitlines():
            self._log_detail(line)
        errors = len(common_report["errors"]) + len(forms_report["errors"])
        warnings = len(common_report["warnings"]) + len(forms_report["warnings"])
        self.state["status"] = f"配置检查完成：指标比较拆分（3.3），错误 {errors} 处，提示 {warnings} 处"
        self._refresh_config_issues()
        return {"passed": errors == 0, "total_issues": errors,
                "common_report": common_report, "forms_report": forms_report,
                "source_workbooks": [str(common_path), str(forms_path)],
                "report_text": report_text}

    def check_central_expression_config(self) -> dict[str, Any]:
        """只读检查本次 3.1 所选表、表达式及动作规则的运行时可用性。"""
        if self.state["busy"]:
            raise RuntimeError("任务正在运行，请完成后再检查配置")
        from .systems.s3_central_statistics.complex_rule_engine import SCHEMA_FIVE_SEGMENT_V1
        from .systems.s3_central_statistics.config import (
            ACTION_RULE_SHEET, EXPRESSION_RULE_SHEET, EXPRESSION_RULE_SHEET_OLD,
            FIVE_SEGMENT_RULE_SHEET, FIVE_SEGMENT_RULE_SHEET_MIGRATED,
            INDICATOR_SHEET, check_action_rules, check_expression_rules,
            load_central_config,
        )
        from openpyxl import load_workbook

        schema = (
            SCHEMA_FIVE_SEGMENT_V1
            if str(self.state.get("centralExpressionSchema")) == "FIVE_SEGMENT_V1"
            else "LEGACY_8"
        )
        common_path, feature_path = self._central_config_paths("comparison")
        for path in (common_path, feature_path):
            if not path.is_file():
                raise ValueError(f"本次执行比较配置工作簿不存在：{path}")
        book = load_workbook(feature_path, read_only=True, data_only=True)
        try:
            sheet_names = set(book.sheetnames)
        finally:
            book.close()
        if schema == SCHEMA_FIVE_SEGMENT_V1:
            expression_sheet = (
                FIVE_SEGMENT_RULE_SHEET_MIGRATED
                if FIVE_SEGMENT_RULE_SHEET_MIGRATED in sheet_names else
                FIVE_SEGMENT_RULE_SHEET
                if FIVE_SEGMENT_RULE_SHEET in sheet_names and EXPRESSION_RULE_SHEET in sheet_names
                else ""
            )
            expression_label = "五段式"
        else:
            expression_sheet = (
                EXPRESSION_RULE_SHEET if EXPRESSION_RULE_SHEET in sheet_names else
                EXPRESSION_RULE_SHEET_OLD if EXPRESSION_RULE_SHEET_OLD in sheet_names else ""
            )
            expression_label = "八段式（兼容）"
        required = (INDICATOR_SHEET, ACTION_RULE_SHEET)
        missing = [name for name in required if name not in sheet_names]
        if not expression_sheet:
            missing.append(
                FIVE_SEGMENT_RULE_SHEET if schema == SCHEMA_FIVE_SEGMENT_V1
                else EXPRESSION_RULE_SHEET
            )
        if missing:
            raise ValueError(
                f"执行比较配置“{feature_path.name}”在{expression_label}模式下缺少本次需读取的工作表："
                f"{'、'.join(missing)}；请检查工作簿及设置中心的 3.1 表达式规则语法。"
            )
        config = load_central_config((common_path, feature_path))
        schema_stats: dict = {}
        issues = check_expression_rules(config, schema=schema, stats_out=schema_stats)
        bad = [item for item in issues if item["问题"]]
        action_stats: dict = {}
        action_errors = check_action_rules(config, stats_out=action_stats)
        if not action_stats["total"]:
            action_errors.append(
                f"「{ACTION_RULE_SHEET}」没有可读取规则；本次执行比较无法加载动作规则。"
            )
        expression_errors = sum(len(item["问题"]) for item in bad)
        error_count = expression_errors + len(action_errors)
        lines = [
            f"检查完成：执行比较 3.1（{expression_label}），共 {error_count} 处错误。",
            "",
            "【本次读取的配置工作簿】",
            f"  通用配置：{common_path.name}",
            f"  执行比较配置：{feature_path.name}",
            "",
            f"【本模式参与执行的 3.1 工作表】{INDICATOR_SHEET}、{ACTION_RULE_SHEET}、{expression_sheet}",
            f"  {INDICATOR_SHEET}：已读取 {len(config.indicators)} 条。",
            f"  {ACTION_RULE_SHEET}：共 {action_stats['total']} 条（启用 {action_stats['enabled']}，"
            f"停用 {action_stats['disabled']}）；启用的累计规则 {action_stats['cumulative']} 条、"
            f"特殊动作 {action_stats['special']} 条。",
            f"  {expression_sheet}：共 {schema_stats.get('total', len(issues))} 条"
            f"（启用 {schema_stats.get('enabled', 0)}，停用 "
            f"{schema_stats.get('disabled', schema_stats.get('total', len(issues)) - schema_stats.get('enabled', 0))}）。",
            "",
            "【实际检查项目】",
            "  [通过] 两册配置文件可读取，且本模式所需的 3.1 工作表存在。",
            f"  [{'未通过' if action_errors else '通过'}] 规则动作：运行时识别 "
            f"{action_stats['recognized']}/{action_stats['enabled']} 条启用规则；"
            "停用规则不参与动作执行检查。",
            f"  [{'未通过' if expression_errors else '通过'}] {expression_label}表达式："
            f"检查 {schema_stats.get('total', len(issues))} 条，错误 {expression_errors} 处。",
            "",
            f"【错误】{error_count} 处：",
        ]
        error_lines = [
            f"{item['规则编号']}（{item['启用']}）：{problem}"
            for item in bad for problem in item["问题"]
        ] + action_errors
        if error_lines:
            for index, message in enumerate(error_lines[:50], 1):
                lines.append(f"  {index}. {message}")
            if len(error_lines) > 50:
                lines.append(f"  ……另有 {len(error_lines) - 50} 处略。")
        else:
            lines.append("  无")
        vocabulary = schema_stats.get("vocabulary") or []
        if vocabulary:
            lines.extend(("", f"【提示】词表复核 {len(vocabulary)} 处（不计入错误）："))
            lines.extend(f"  {index}. {note}" for index, note in enumerate(vocabulary[:50], 1))
            if len(vocabulary) > 50:
                lines.append(f"  ……另有 {len(vocabulary) - 50} 处略。")
        lines.extend((
            "", "【检查范围】",
            "  本次只读检查配置及规则可解析性；不读取本期/上期业务数据，不判断规则是否命中。",
            "  未选中的另一张表达式表不参与本次规则判断。",
        ))
        report_text = "\n".join(lines)
        for line in report_text.splitlines():
            self._log_detail(line)
        self.state["status"] = f"配置检查完成：执行比较 3.1（{expression_label}），错误 {error_count} 处"
        # 配置修复并检查通过后，同步刷新右上角「配置需检查」标记。
        self._refresh_config_issues()
        return {"checked": schema_stats.get("total", len(issues)), "issues": bad,
                "total_issues": error_count, "action_errors": action_errors,
                "action_stats": action_stats, "schema": schema,
                "source_workbooks": [str(common_path), str(feature_path)],
                "active_sheets": [*required, expression_sheet], "report_text": report_text}

    def check_central_cross_formulas(self) -> dict[str, Any]:
        """只读检查 3.2 跨期数值核对：token 结构/词表/作用域正则/公式语法。"""
        if self.state["busy"]:
            raise RuntimeError("任务正在运行，请完成后再检查配置")
        from openpyxl import load_workbook
        from .systems.s3_central_statistics.config import CROSS_SHEET, check_cross_formulas, load_central_config
        from .systems.s3_central_statistics.common_config_check import check_common_config

        common_path, cross_path = self._central_config_paths("cross")
        for path in (common_path, cross_path):
            if not path.is_file():
                raise ValueError(f"本次数值核对配置工作簿不存在：{path}")
        book = load_workbook(cross_path, read_only=True, data_only=True)
        try:
            if CROSS_SHEET not in book.sheetnames:
                raise ValueError(
                    f"本期数值核对配置“{cross_path.name}”缺少必需工作表“{CROSS_SHEET}”；"
                    "请检查所选工作簿。"
                )
        finally:
            book.close()
        config = load_central_config((common_path, cross_path))
        common_report = check_common_config(common_path)
        cross_stats: dict = {}
        issues = check_cross_formulas(config, stats_out=cross_stats)
        bad = [item for item in issues if item["问题"]]
        report_text = "\n".join((
            "【本次读取的配置工作簿】",
            f"  通用配置：{common_path.name}",
            f"  本期数值核对配置：{cross_path.name}",
            f"【本模式参与执行的 3.2 工作表】{CROSS_SHEET}",
            "【实际检查项目】",
            "  [通过] 两册配置文件可读取，且本模式所需的 3.2 工作表存在。",
            f"  [{'未通过' if bad else '通过'}] 跨期数值核对："
            f"已检查 {cross_stats.get('total', len(issues))} 条规则的"
            "前提条件、left、right、校验公式中的五段 token 结构、机构/地区作用域正则与公式语法；"
            f"错误 {sum(len(item['问题']) for item in bad)} 处。",
            f"  [{'未通过' if common_report['errors'] else '通过'}] 通用配置 3.0："
            f"四张必需工作表、运行参数单位及环比警戒区间；"
            f"错误 {len(common_report['errors'])} 处，提示 {len(common_report['warnings'])} 处。",
            "",
            self._central_check_report("本期数值核对配置（3.2）", issues, cross_stats),
            "",
            self._central_preflight_text(
                "依赖的通用配置（3.0）", [("通用配置", common_report)],
                "单位换算例外、机构地区参照只检查工作表是否存在。",
            ),
            "",
            "【检查范围】只读检查配置和规则语法，不读取本期业务数据，不判断规则是否命中。",
        ))
        for line in report_text.splitlines():
            self._log_detail(line)
        total_errors = sum(len(i["问题"]) for i in bad) + len(common_report["errors"])
        self.state["status"] = (
            f"配置检查完成：本期数值核对配置（3.2）共 {cross_stats.get('total', len(issues))} 条，"
            f"错误 {total_errors} 处，提示 {len(common_report['warnings'])} 处"
        )
        # 同上：检查后刷新「配置需检查」标记。
        self._refresh_config_issues()
        return {"checked": cross_stats.get("total", len(issues)), "issues": bad,
                "total_issues": total_errors,
                "common_report": common_report,
                "source_workbooks": [str(common_path), str(cross_path)],
                "active_sheets": [CROSS_SHEET],
                "report_text": report_text}

    def _central_config_paths(self, profile: str) -> tuple[Path, Path]:
        mapping = {
            "comparison": ("centralComparisonConfig", COMPARISON_CONFIG_NAME),
            "cross": ("centralCrossConfig", CROSS_CONFIG_NAME),
            "forms": ("centralFormsConfig", FORM_CONFIG_NAME),
        }
        if profile not in mapping:
            raise ValueError(f"未知大集中配置类型：{profile}")
        config_dir = release_config_dir(self.project_root)
        common = Path(str(self.state.get("centralCommonConfig") or config_dir / COMMON_CONFIG_NAME))
        state_key, default_name = mapping[profile]
        feature = Path(str(self.state.get(state_key) or config_dir / default_name))
        return common, feature

    def _central_config_profile(self, profile: str) -> tuple[str, str, Path]:
        """Return the state key, readable label and active file for one config book.

        The active file is the one chosen on the 大集中 page.  When no custom
        path has been selected it resolves to the release-level config folder.
        This keeps “打开” and “重置” aligned with the file that will actually
        be used by the next run.
        """
        mapping = {
            "common": ("centralCommonConfig", "大集中通用配置", COMMON_CONFIG_NAME),
            "comparison": ("centralComparisonConfig", "执行比较配置", COMPARISON_CONFIG_NAME),
            "cross": ("centralCrossConfig", "本期数值核对配置", CROSS_CONFIG_NAME),
            "forms": ("centralFormsConfig", "指标比较拆分配置", FORM_CONFIG_NAME),
        }
        if profile not in mapping:
            raise ValueError(f"未知大集中配置类型：{profile}")
        state_key, label, default_name = mapping[profile]
        config_dir = release_config_dir(self.project_root)
        active = Path(str(self.state.get(state_key) or config_dir / default_name))
        return state_key, label, active

    def reset_central_config_profile(self, profile: str) -> dict[str, Any]:
        """Reset one active central-statistics config book with a sibling backup.

        A custom path remains custom: the reset writes the clean profile to the
        selected location rather than silently switching a user's workflow back
        to the release config folder.
        """
        import shutil

        state_key, label, target = self._central_config_profile(profile)
        snapshot = self._config_default_paths()[profile]
        target.parent.mkdir(parents=True, exist_ok=True)
        backup_path = None
        if target.is_file():
            try:
                backup_path = self._backup_config_file(target)
            except OSError as exc:
                raise RuntimeError(f"无法备份“{target.name}”：{exc}") from exc
        try:
            source = snapshot
            if not source.is_file():
                _, _, release_name = {
                    "common": ("", "", COMMON_CONFIG_NAME),
                    "comparison": ("", "", COMPARISON_CONFIG_NAME),
                    "cross": ("", "", CROSS_CONFIG_NAME),
                    "forms": ("", "", FORM_CONFIG_NAME),
                }[profile]
                source = release_config_dir(self.project_root) / release_name
            if not source.is_file():
                raise RuntimeError(f"默认配置快照不存在：{snapshot}")
            shutil.copy2(source, target)
        except (OSError, RuntimeError) as exc:
            raise RuntimeError(f"无法重置“{target.name}”：{exc}") from exc
        self.state[state_key] = str(target.resolve())
        self._save_settings()
        self._refresh_config_issues()
        self.state["status"] = f"{label}已重置"
        backup_text = str(backup_path) if backup_path else "原文件不存在，未生成备份"
        source_text = f"默认快照：{snapshot}" if snapshot.is_file() else "出厂结构"
        self._log(f"已重置{label}：{target}（{source_text}；{backup_text}）")
        return self.state

    def reset_central_config(self) -> dict[str, Any]:
        """将 3.0/3.1/3.2/3.3 一次性恢复为“默认配置”目录中的快照。"""
        import shutil

        mapping = {
            "common": "centralCommonConfig",
            "comparison": "centralComparisonConfig",
            "cross": "centralCrossConfig",
            "forms": "centralFormsConfig",
        }
        for profile, state_key in mapping.items():
            _, _, target = self._central_config_profile(profile)
            snapshot = self._config_default_paths()[profile]
            if not snapshot.is_file():
                raise RuntimeError(f"默认配置快照不存在：{snapshot}")
            target.parent.mkdir(parents=True, exist_ok=True)
            if target.is_file():
                self._backup_config_file(target)
            shutil.copy2(snapshot, target)
            self.state[state_key] = str(target.resolve())
        self._save_settings()
        self._log("已从默认配置恢复 3.0/3.1/3.2/3.3")
        self.state["status"] = "大集中配置已重置"
        self._refresh_config_issues()
        return self.state

    def _central_output_dir(self, kind_key: str) -> Path:
        value = str(self.state.get(kind_key) or "")
        return Path(value) if value else None

    def start_central_comparison(self) -> bool:
        """执行比较：两期大集中导数（CSV/XLSX）→ 比较结果.xlsx。"""
        if self.state["busy"]:
            return False
        cur, pre = self.state.get("centralCur"), self.state.get("centralPre")
        if not cur or not pre:
            self._log("执行比较失败：请先选择本期与上期导数文件")
            return False
        output_dir = self._central_output_dir("centralOutput") or (Path(cur).parent / "执行结果")
        output_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = output_dir / f"比较结果_{stamp}.xlsx"
        self.state["busy"] = True
        self.state["status"] = "正在执行大集中比较，请勿关闭窗口……"
        threading.Thread(
            target=self._central_comparison_worker,
            args=(Path(cur), Path(pre), output_path),
            daemon=True,
        ).start()
        return True

    def _central_comparison_worker(self, current_csv: Path, previous_csv: Path, output_path: Path) -> None:
        from .systems.s3_central_statistics.service import run_comparison

        started = time.monotonic()
        try:
            expression_mode = str(self.state.get("expressionEvaluationMode") or "SBE").upper()
            expression_backend = str(self.state.get("expressionEvaluationBackend") or "PYTHON").upper()
            office_adapter = None
            if expression_backend == "OFFICE":
                from .systems.s3_central_statistics.office_eval import OfficeEvaluationAdapter

                office_adapter = OfficeEvaluationAdapter("自动")
            try:
                if office_adapter is not None:
                    office_adapter.__enter__()
                result = run_comparison(
                    current_csv=current_csv, previous_csv=previous_csv,
                    output_path=output_path, config_path=self._central_config_paths("comparison"),
                    on_step=self._log_detail, write_flow_logs=bool(self.state["exportRunLogs"]),
                    rule_engine="v3",
                    office_evaluator=office_adapter,
                    expression_mode=expression_mode,
                    expression_backend=expression_backend,
                    expression_schema=str(self.state.get("centralExpressionSchema") or "LEGACY_8"),
                )
            finally:
                if office_adapter is not None:
                    office_adapter.__exit__(None, None, None)
            self.state["centralCompareFile"] = str(result.output_path)
            if self.state.get("centralFormOutputAuto", True):
                self.state["centralFormOutput"] = str(result.output_path.parent)
            # 比较结果自动带入指标拆分/说明导出卡片，须落盘才能在重启后仍显示。
            self._save_settings()
            self.state["status"] = f"执行比较完成：{len(result.rows)} 行"
            self._log(f"执行比较完成：{result.output_path}")
        except Exception as exc:
            self.state["status"] = "执行比较失败"
            self._log("执行比较失败：" + str(exc) + "\n" + traceback.format_exc())
        finally:
            self._log(f"任务结束，本次耗时：{time.monotonic() - started:.1f} 秒")
            self.state["busy"] = False

    def start_central_cross_period(self) -> bool:
        """本期数值核对：专用单期数据文件（CSV/XLSX）→ 对比结果工作簿。"""
        if self.state["busy"]:
            return False
        # 本期数值核对与两期执行比较使用不同来源，不能静默借用两期的数据文件。
        cur = self.state.get("centralCrossCur")
        if not cur:
            self._log("本期数值核对失败：请先选择数据文件（CSV/XLSX）")
            return False
        output_dir = (self._central_output_dir("centralOutput")
                      or self._central_output_dir("centralCrossOutput")
                      or (Path(cur).parent / "执行结果"))
        output_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = output_dir / f"数值核对结果_{stamp}.xlsx"
        self.state["busy"] = True
        self.state["status"] = "正在执行本期数值核对，请勿关闭窗口……"
        threading.Thread(
            target=self._central_cross_period_worker,
            args=(Path(cur), "", output_path),
            daemon=True,
        ).start()
        return True

    def _central_cross_period_worker(self, current_csv: Path, previous_csv: str, output_path: Path) -> None:
        from .systems.s3_central_statistics.service import run_cross_period_check

        started = time.monotonic()
        try:
            result = run_cross_period_check(
                current_csv=current_csv, previous_csv=previous_csv,
                output_path=output_path, config_path=self._central_config_paths("cross"),
                on_step=self._log_detail, write_flow_logs=bool(self.state["exportRunLogs"]),
            )
            self.state["status"] = f"本期数值核对完成：{len(result.rows)} 行"
            self._log(f"本期数值核对完成：{result.output_path}")
        except Exception as exc:
            self.state["status"] = "本期数值核对失败"
            self._log("本期数值核对失败：" + str(exc) + "\n" + traceback.format_exc())
        finally:
            self._log(f"任务结束，本次耗时：{time.monotonic() - started:.1f} 秒")
            self.state["busy"] = False

    def start_central_forms(self) -> bool:
        """指标比较拆分：比较结果 + 3.3 内嵌金融表单 → 每机构一个 xlsx。"""
        if self.state["busy"]:
            return False
        compare_file = self.state.get("centralCompareFile")
        if not compare_file or not Path(compare_file).is_file():
            self._log("指标比较拆分失败：请先选择比较结果.xlsx（可先执行比较，结果会自动带入）")
            return False
        output_root = (self._central_output_dir("centralOutput")
                       or self._central_output_dir("centralFormOutput")
                       or (Path(compare_file).parent / "执行结果"))
        output_dir = output_root / f"指标比较拆分_{datetime.now():%Y%m%d_%H%M%S}"
        output_dir.mkdir(parents=True, exist_ok=True)
        self.state["busy"] = True
        self.state["status"] = "正在执行指标比较拆分，请勿关闭窗口……"
        threading.Thread(
            target=self._central_forms_worker,
            args=(Path(compare_file), output_dir),
            daemon=True,
        ).start()
        return True

    def _central_forms_worker(self, compare_file: Path, output_dir: Path) -> None:
        from .systems.s3_central_statistics.service import render_financial_forms

        started = time.monotonic()
        try:
            render_mode = str(self.state.get("xlsxRenderMode") or DEFAULT_XLSX_RENDER_MODE)
            self._log(f"Renderer: {render_mode}")
            result = render_financial_forms(
                comparison_file=compare_file,
                output_dir=output_dir, config_path=self._central_config_paths("forms"),
                hide_empty_rows=True, delete_empty_sheets=True,
                render_mode=render_mode, on_step=self._log_detail,
            )
            self.state["status"] = f"金融表单生成完成：{len(result.files)} 个文件"
            self._log(f"金融表单生成完成：{len(result.files)} 个文件，输出目录 {output_dir}")
        except Exception as exc:
            self.state["status"] = "指标比较拆分失败"
            self._log("指标比较拆分失败：" + str(exc) + "\n" + traceback.format_exc())
        finally:
            self._log(f"任务结束，本次耗时：{time.monotonic() - started:.1f} 秒")
            self.state["busy"] = False

    def start_central_explanation_export(self) -> bool:
        """将比较结果中人工标记“输出说明=是”的行导出为简化说明文件。"""
        if self.state["busy"]:
            return False
        compare_file = self.state.get("centralCompareFile")
        if not compare_file or not Path(compare_file).is_file():
            self._log("导出指标说明失败：请先选择比较结果.xlsx")
            return False
        output_root = (self._central_output_dir("centralOutput")
                       or self._central_output_dir("centralFormOutput")
                       or (Path(compare_file).parent / "执行结果"))
        output_root.mkdir(parents=True, exist_ok=True)
        output_path = output_root / f"指标说明_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        self.state["busy"] = True
        self.state["status"] = "正在导出选定指标说明，请勿关闭窗口……"
        threading.Thread(
            target=self._central_explanation_export_worker,
            args=(Path(compare_file), output_path),
            daemon=True,
        ).start()
        return True

    def _central_explanation_export_worker(self, compare_file: Path, output_path: Path) -> None:
        from .systems.s3_central_statistics.explanation_exporter import export_selected_explanations

        started = time.monotonic()
        try:
            count = export_selected_explanations(
                compare_file, output_path, config_path=self._central_config_paths("comparison")[0]
            )
            self.state["centralExplainFile"] = str(output_path)
            self._save_settings()
            self.state["status"] = f"指标说明导出完成：{count} 条"
            self._log(f"指标说明导出完成：{count} 条，文件 {output_path}")
        except Exception as exc:
            self.state["status"] = "指标说明导出失败"
            self._log("指标说明导出失败：" + str(exc) + "\n" + traceback.format_exc())
        finally:
            self._log(f"任务结束，本次耗时：{time.monotonic() - started:.1f} 秒")
            self.state["busy"] = False

    def import_central_explanation_feedback(self, paths: list[str] | None = None) -> bool:
        """批量选择机构反馈说明，并按业务标识回写当前说明文件。"""
        if self.state["busy"]:
            return False
        target = self.state.get("centralExplainFile")
        if not target or not Path(target).is_file():
            self._log("批量导入说明失败：请先选择或导出说明文件.xlsx")
            return False
        selected = tuple(str(item) for item in paths) if paths is not None else self._select_path(
            folder=False, current=str(Path(target).parent), multiple=True,
            file_types=(("Excel 文件", "*.xlsx"),),
        )
        if not selected:
            return False
        feedback_paths = [Path(item) for item in selected if Path(item).is_file()]
        if not feedback_paths:
            self._log("批量导入说明失败：未选择有效的机构反馈文件")
            return False
        self.state["busy"] = True
        self.state["status"] = "正在批量导入机构说明，请勿关闭窗口……"
        threading.Thread(
            target=self._central_explanation_import_worker,
            args=(Path(target), feedback_paths), daemon=True,
        ).start()
        return True

    def _central_explanation_import_worker(self, target: Path, feedback_paths: list[Path]) -> None:
        from .systems.s3_central_statistics.explanation_exporter import import_explanation_feedback

        started = time.monotonic()
        try:
            result = import_explanation_feedback(target, feedback_paths)
            self.state["status"] = (
                f"机构说明导入完成：{result.imported} 条"
                + (f"；冲突 {result.conflicts} 条，已生成冲突清单" if result.conflict_report_path else "")
            )
            self._log(
                f"机构说明导入完成：写入 {result.imported} 条；未匹配 {result.unmatched} 条；"
                f"冲突 {result.conflicts} 条；导入前备份 {result.backup_path}"
                + (f"；冲突清单 {result.conflict_report_path}" if result.conflict_report_path else "")
            )
        except Exception as exc:
            self.state["status"] = "机构说明导入失败"
            self._log("机构说明导入失败：" + str(exc) + "\n" + traceback.format_exc())
        finally:
            self._log(f"任务结束，本次耗时：{time.monotonic() - started:.1f} 秒")
            self.state["busy"] = False

    def choose_period_compare(self, kind: str, path: str = "") -> str:
        """报表采集系统页：选择跨期比较的输入或输出路径。

        只弹窗并把结果记入 state，供页面回显；不触发执行。
        """
        if self.state["busy"] or kind not in {"pcCurDir", "pcPreDir", "pcCentral", "pcConfig", "pcOutput"}:
            return ""
        titles = {
            "pcCurDir": "选择当期（本期）数据目录",
            "pcPreDir": "选择上期数据目录",
            "pcCentral": "选择大集中数据文件",
            "pcConfig": f"选择{CONFIG_WORKBOOK_NAME}",
            "pcOutput": "选择输出目录",
        }
        current = self.state.get(kind) or self.state.get("input") or str(self.project_root)
        if path:
            selected = (str(path),)
        elif kind in {"pcCentral", "pcConfig"}:
            selected = self._select_path(
                folder=False, current=str(Path(current).parent) if Path(current).is_file() else str(current),
                # 大集中数据仅 .xlsx（.xls/.csv 会丢「参照表」）；配置册为 OOXML 工作簿。
                file_types=(("Excel 工作簿", "*.xlsx"),) if kind == "pcCentral"
                            else (("Excel 文件", "*.xlsx *.xlsm"),),
            )
        else:
            selected = self._select_path(folder=True, current=str(current))
        value = str(selected[0]) if selected else ""
        if value:
            self.state[kind] = str(Path(value).resolve())
            if kind == "pcCurDir" and self.state.get("pcOutputAuto", True):
                self.state["pcOutput"] = str(Path(value).resolve() / "执行结果")
            elif kind == "pcOutput":
                self.state["pcOutputAuto"] = False
            label = {
                "pcCurDir": "当期目录", "pcPreDir": "上期目录",
                "pcCentral": "大集中数据", "pcConfig": "报表采集系统配置", "pcOutput": "输出目录",
            }[kind]
            self._log(f"跨期比较：{label}已选择")
            self._save_settings()
            if kind == "pcConfig":
                self._refresh_config_issues()
                self._refresh_period_config_check()
        self._refresh_period_pairs()
        return self.state.get(kind, "")

    def toggle_period_output_pin(self) -> bool:
        """切换报表采集输出目录的固定状态；未固定时跟随本期目录。"""
        if self.state["busy"]:
            return not self.state.get("pcOutputAuto", True)
        auto = not bool(self.state.get("pcOutputAuto", True))
        self.state["pcOutputAuto"] = auto
        if auto and self.state.get("pcCurDir"):
            self.state["pcOutput"] = str(Path(self.state["pcCurDir"]) / "执行结果")
        self._save_settings()
        self._log("跨期比较：输出目录" + ("跟随本期目录" if auto else "已固定"))
        return not auto

    def start_period_compare(self) -> bool:
        """报表采集系统页：用页面已选择的三路径启动跨期比较。"""
        if self.state["busy"]:
            return False
        cur = self.state.get("pcCurDir")
        pre = self.state.get("pcPreDir")
        if not cur or not pre:
            self._log("跨期比较失败：请先选择本期待处理与上期待处理目录")
            return False
        current_dir = Path(cur)
        previous_dir = Path(pre)
        central_raw = self.state.get("pcCentral")
        central_path = Path(central_raw) if central_raw else None
        if central_path is not None and not central_path.is_file():
            self._log(f"跨期比较失败：大集中数据文件不存在：{central_path}")
            return False
        output_dir = Path(self.state.get("pcOutput") or (current_dir / "执行结果"))
        self.state["busy"] = True
        self.state["status"] = "正在执行跨期比较，请勿关闭窗口……"
        self._log(f"开始跨期比较：当期“{current_dir.name}” vs 上期“{previous_dir.name}”")
        threading.Thread(
            target=self._period_compare_worker,
            args=(current_dir, previous_dir, central_path, output_dir),
            daemon=True,
        ).start()
        return True

    def _period_compare_worker(
        self,
        current_dir: Path,
        previous_dir: Path,
        central_path: Path | None,
        output_dir: Path,
    ) -> None:
        from .period_compare import (
            PeriodCompareError,
            ensure_default_config,
            run_period_compare,
        )

        started = time.monotonic()
        config_path = Path(self.state.get("pcConfig") or release_config_dir(self.project_root) / CONFIG_WORKBOOK_NAME)
        try:
            if ensure_default_config(config_path):
                self._log(f"首次使用：已生成默认配置模板 {config_path.name}，可按需维护外部文件对比规则/机构参照/环比策略/校验规则")
            output_path = run_period_compare(
                current_dir=current_dir,
                previous_dir=previous_dir,
                central_path=central_path,
                output_dir=output_dir,
                config_path=config_path,
                central_tolerance_yuan=self.state.get("centralTolerance"),
                on_step=self._log_detail,
            )
            self.state["status"] = "跨期比较完成"
            self._log(f"跨期比较完成：{output_path}")
        except PeriodCompareError as exc:
            self.state["status"] = "跨期比较失败"
            self._log("跨期比较失败：" + str(exc))
        except Exception as exc:
            self.state["status"] = "跨期比较失败"
            self._log("跨期比较失败：" + str(exc) + "\n" + traceback.format_exc())
        finally:
            elapsed = time.monotonic() - started
            self._log(f"任务结束，本次耗时：{elapsed:.1f} 秒")
            self.state["busy"] = False

    def start_period_compare_v2(self) -> bool:
        """启动当前报表采集审核主流程，输出带时间戳的报表审核清单。"""
        if self.state["busy"]:
            return False
        config_report = self._refresh_period_config_check()
        if not config_report["passed"]:
            self._log(format_config_report(config_report))
            self.state["status"] = "跨期比较失败：报表采集系统配置存在阻断错误"
            return False
        cur, pre = self.state.get("pcCurDir"), self.state.get("pcPreDir")
        if not cur or not pre:
            self._log("跨期比较失败：请先选择本期与上期目录")
            return False
        central = Path(self.state["pcCentral"]) if self.state.get("pcCentral") else None
        if central is not None and not central.is_file():
            self._log(f"跨期比较失败：大集中数据文件不存在：{central}")
            return False
        self.state["busy"] = True
        self.state["status"] = "正在执行跨期比较，请勿关闭窗口……"
        threading.Thread(
            target=self._period_compare_v2_worker,
            args=(Path(cur), Path(pre), central, Path(self.state.get("pcOutput") or (Path(cur) / "执行结果"))),
            daemon=True,
        ).start()
        return True

    def _period_compare_v2_worker(self, current_dir: Path, previous_dir: Path, central_path: Path | None, output_dir: Path) -> None:
        started = time.monotonic()
        try:
            from .systems.s2_report_collection import PeriodAuditV2Service
            result = PeriodAuditV2Service().run(
                current_dir=current_dir, previous_dir=previous_dir, central_path=central_path,
                output_dir=output_dir,
                config_path=Path(self.state.get("pcConfig") or release_config_dir(self.project_root) / CONFIG_WORKBOOK_NAME),
            )
            self.state["periodSummary"] = {
                "institutions": result.summary.institutions, "forms": result.summary.forms,
                "indicators": result.summary.indicators, "abnormal": result.summary.abnormal_findings,
                "byType": result.summary.by_type, "result": str(result.paths.result),
            }
            self.state["status"] = f"跨期比较完成：异常 {result.summary.abnormal_findings} 条"
            self._log(f"跨期比较完成：审核清单 {result.paths.result}")
        except Exception as exc:
            self.state["status"] = "跨期比较失败"
            self._log("跨期比较失败：" + str(exc) + "\n" + traceback.format_exc())
        finally:
            self._log(f"跨期比较任务结束，本次耗时：{time.monotonic() - started:.1f} 秒")
            # 配对清单只在选择目录/启动时计算；执行后刷新一次，
            # 让早前一次性配对失败的空列表自愈（例如文件被占用）。
            self._refresh_period_pairs()
            self.state["busy"] = False

    def update(self, values: dict[str, Any]) -> dict[str, Any]:
        for key in ("input", "templateDir", "template", "external", "output"):
            if key in values:
                new_value = str(values[key]).strip()
                previous = self.state.get(key, "")
                if key == "output" and new_value != self.state.get("output", ""):
                    self.state["outputAuto"] = False
                    self.state["outputPinned"] = True
                self.state[key] = new_value
                if key == "input" and new_value != previous:
                    self.state["selectedFilesInitialized"] = False
                if key == "template" and new_value and new_value != previous:
                    self.state["templateManual"] = True
                elif key == "templateDir" and new_value != previous:
                    self.state["templateManual"] = False
        if "outputPinned" in values:
            value = values["outputPinned"]
            pinned = (
                value if isinstance(value, bool)
                else str(value).strip().casefold() in {"1", "true", "yes", "y", "是"}
            )
            self.state["outputPinned"] = pinned
            self.state["outputAuto"] = not pinned
            if not pinned and self.state["input"]:
                self.state["output"] = str(Path(self.state["input"]) / "执行结果")
        if "recursiveDepth" in values:
            raw = values["recursiveDepth"]
            try:
                depth = int(raw)
            except (TypeError, ValueError):
                depth = -1 if str(raw).strip() in {"最深处", "true", "1", "是"} else 0
            if depth != self.state["recursiveDepth"]:
                self.state["recursiveDepth"] = depth
        if "recursive" in values:
            value = values["recursive"]
            recursive = (
                value if isinstance(value, bool)
                else str(value).strip().casefold() in {"1", "true", "yes", "y", "是"}
            )
            depth = -1 if recursive else 0
            changed = depth != self.state["recursiveDepth"]
            self.state["recursiveDepth"] = depth
            if changed:
                # 复选框直接决定待审核清单的扫描范围；不触发模板自动改写。
                self._recognize(allow_template_auto=False)
        if "xlsxRenderMode" in values:
            mode = str(values["xlsxRenderMode"]).strip().upper()
            if mode not in XLSX_RENDER_MODES:
                mode = DEFAULT_XLSX_RENDER_MODE
            self.state["xlsxRenderMode"] = mode
        legacy_log_setting = "writeFlowLogs" in values
        if "showRunDetailLogs" in values:
            value = values["showRunDetailLogs"]
            self.state["showRunDetailLogs"] = (
                value if isinstance(value, bool)
                else str(value).strip().casefold() in {"1", "true", "yes", "y", "是"}
            )
        if "exportRunLogs" in values:
            value = values["exportRunLogs"]
            self.state["exportRunLogs"] = (
                value if isinstance(value, bool)
                else str(value).strip().casefold() in {"1", "true", "yes", "y", "是"}
            )
        if "showPerformanceDiagnostics" in values:
            value = values["showPerformanceDiagnostics"]
            self.state["showPerformanceDiagnostics"] = (
                value if isinstance(value, bool)
                else str(value).strip().casefold() in {"1", "true", "yes", "y", "是"}
            )
        if legacy_log_setting:
            # 旧 API 仍表示原来的“两个能力同时开启/关闭”；新 API 则完全独立。
            value = values["writeFlowLogs"]
            legacy_value = (
                value if isinstance(value, bool)
                else str(value).strip().casefold() in {"1", "true", "yes", "y", "是"}
            )
            if "showRunDetailLogs" not in values:
                self.state["showRunDetailLogs"] = legacy_value
            if "exportRunLogs" not in values:
                self.state["exportRunLogs"] = legacy_value
        if legacy_log_setting or any(
            key in values for key in (
                "showRunDetailLogs", "exportRunLogs", "showPerformanceDiagnostics"
            )
        ):
            # 旧状态字段只保留为导出开关的兼容别名。
            self.state["writeFlowLogs"] = bool(self.state["exportRunLogs"])
            self._save_settings()
        if "confirmBeforeRun" in values:
            value = values["confirmBeforeRun"]
            self.state["confirmBeforeRun"] = (
                value if isinstance(value, bool)
                else str(value).strip().casefold() in {"1", "true", "yes", "y", "是"}
            )
            self._save_settings()
        if "uiTheme" in values:
            candidate = str(values["uiTheme"]).strip()
            if candidate not in UI_THEMES:
                self._log("背景颜色只能选择：" + "、".join(UI_THEMES))
            else:
                self.state["uiTheme"] = candidate
                self._save_settings()
        if "uiAccent" in values:
            candidate = str(values["uiAccent"]).strip()
            if candidate not in UI_ACCENTS:
                self._log("主题色只能选择：" + "、".join(UI_ACCENTS))
            else:
                self.state["uiAccent"] = candidate
                self._save_settings()
        if "filePickerMode" in values:
            candidate = str(values["filePickerMode"]).strip()
            if candidate not in FILE_PICKER_MODES:
                self._log("文件/目录选择方式只能选择：" + "、".join(FILE_PICKER_MODES))
            else:
                self.state["filePickerMode"] = candidate
                self._save_settings()
        if "expressionEvaluationMode" in values:
            candidate = str(values["expressionEvaluationMode"]).strip().upper()
            if candidate not in EXPRESSION_EVALUATION_MODES:
                self._log("表达式处理方式只能选择：" + "、".join(EXPRESSION_EVALUATION_MODES))
            else:
                self.state["expressionEvaluationMode"] = candidate
                self._save_settings()
        if "expressionEvaluationBackend" in values:
            candidate = str(values["expressionEvaluationBackend"]).strip().upper()
            if candidate not in EXPRESSION_EVALUATION_BACKENDS:
                self._log("表达式求值方式只能选择：" + "、".join(EXPRESSION_EVALUATION_BACKENDS))
            else:
                self.state["expressionEvaluationBackend"] = candidate
                self._save_settings()
        if "centralExpressionSchema" in values:
            candidate = str(values["centralExpressionSchema"]).strip().upper()
            if candidate not in ("LEGACY_8", "FIVE_SEGMENT_V1"):
                self._log("3.1 表达式规则语法只能选择：FIVE_SEGMENT_V1、LEGACY_8")
            else:
                self.state["centralExpressionSchema"] = candidate
                self._save_settings()
        if "conditionalFormatEvaluator" in values:
            candidate = str(values["conditionalFormatEvaluator"]).strip().upper()
            if candidate not in CONDITIONAL_FORMAT_EVALUATORS:
                self._log("条件格式判定方式只能选择：" + "、".join(CONDITIONAL_FORMAT_EVALUATORS))
            elif candidate in {"COM_EVALUATE", "COM_DISPLAY"} and pipeline_kind("自动") != "com":
                # UOS 没有 Excel/WPS COM：两个 COM 相关判定方式不可用。
                self._log("当前平台没有 Excel/WPS 引擎，条件格式判定固定使用 Python 规则求值")
                self.state["conditionalFormatEvaluator"] = "PYTHON"
                self._save_settings()
            else:
                self.state["conditionalFormatEvaluator"] = candidate
                self._save_settings()
        if "conditionalFormatRuleReader" in values:
            candidate = str(values["conditionalFormatRuleReader"]).strip().upper()
            if candidate not in CONDITIONAL_FORMAT_RULE_READERS:
                self._log("条件格式规则读取方式只能选择：" + "、".join(CONDITIONAL_FORMAT_RULE_READERS))
            else:
                self.state["conditionalFormatRuleReader"] = candidate
                self._save_settings()
        if "formulaRegionWriter" in values:
            candidate = str(values["formulaRegionWriter"]).strip().upper()
            if candidate not in FORMULA_REGION_WRITERS:
                self._log("公式校验复制方式只能选择：" + "、".join(FORMULA_REGION_WRITERS))
            else:
                self.state["formulaRegionWriter"] = candidate
                self._save_settings()
        if "externalSheetWriter" in values:
            candidate = str(values["externalSheetWriter"]).strip().upper()
            if candidate not in EXTERNAL_SHEET_WRITERS:
                self._log("外部文件添加方式只能选择：" + "、".join(EXTERNAL_SHEET_WRITERS))
            else:
                self.state["externalSheetWriter"] = candidate
                self._save_settings()
        if "calculationEngine" in values:
            candidate = str(values["calculationEngine"]).strip()
            if candidate not in valid_engine_values():
                self._log("计算引擎只能选择：" + "、".join(sorted(valid_engine_values())))
            else:
                self.state["calculationEngine"] = candidate
                self._save_settings()
        if "summaryReadEngine" in values:
            candidate = str(values["summaryReadEngine"]).strip()
            if candidate not in valid_summary_reader_values():
                self._log("汇总读取方式只能选择：" + "、".join(
                    item["label"] for item in available_summary_readers()
                ))
            else:
                self.state["summaryReadEngine"] = candidate
                self._save_settings()
        if "centralTolerance" in values:
            raw = values["centralTolerance"]
            try:
                tolerance = round(float(raw), 2)
            except (TypeError, ValueError):
                tolerance = 100.0
            if tolerance < 0 or tolerance > 1_000_000:
                self._log("大集中核对容差需在 0~1000000 元之间，已恢复默认 100 元")
                tolerance = 100.0
            if tolerance != self.state.get("centralTolerance"):
                self.state["centralTolerance"] = tolerance
                self._save_settings()
        return self.state

    def recognize(self) -> dict[str, Any]:
        self._recognize(force_template=True)
        return self.state

    def inspect_template_formulas(self) -> dict[str, Any]:
        """只读检查模板：公式明确错误 + 业务命名区域盘点（只提示不阻断）。"""
        if self.state["busy"]:
            raise RuntimeError("任务正在运行，请完成后再检查模板公式")
        template = Path(str(self.state.get("template") or ""))
        if not template.is_file():
            raise ValueError("请先选择或自动识别匹配模板")
        from .formula_inspection import inspect_template_formulas

        report = inspect_template_formulas(template)
        ranges = "、".join(
            "{} {}个".format(group["family"], group["count"])
            for group in report.get("namedRangeGroups", []))
        others = len(report.get("otherNames") or [])
        internals = len(report.get("internalNames") or [])
        self._log(
            "检查模板：共 {} 个公式；明确错误 {} 处；命名区域 {}；其他 {} 个；内部名称 {} 个".format(
                report["formulaCount"], len(report["errors"]),
                ranges or "无", others, internals,
            )
        )
        return report

    def set_selected_files(self, paths: list[str]) -> dict[str, Any]:
        allowed = {item["path"] for item in self.state.get("sourceFiles", [])}
        self.state["selectedFiles"] = [path for path in paths if path in allowed]
        # 空列表同样是用户的有效选择，后续刷新不得把它解释成首次加载。
        self.state["selectedFilesInitialized"] = True
        return self.state

    def start(self, action: str, strict: bool = True) -> bool:
        if self.state["busy"]:
            return False
        self.state["busy"] = True
        self.state["status"] = "正在处理，请勿关闭窗口……"
        # Python 3.7 is used by the Win7 package; str.removeprefix arrived in
        # Python 3.9, so use slicing here instead.
        display = action[len("flow:"):] if action.startswith("flow:") else action
        action_name = {"check": "审核前检查", "audit": "汇总核查表校验", "summary": "汇总校验结果说明"}.get(display, display)
        self._log(f"开始执行：{action_name}")
        threading.Thread(target=self._worker, args=(action, strict), daemon=True).start()
        return True

    def start_dag_flow(
        self,
        workflow_id: str,
        values: dict[str, Any] | None = None,
        selected_files: list[str] | None = None,
        strict: bool = True,
    ) -> bool:
        """DAG 并行通道：与既有流程同口径的显式 DAG 工作流（用于新旧对比）。"""
        if values:
            self.update(values)
        if selected_files is not None:
            self.set_selected_files(selected_files)
        name = str(workflow_id).strip()
        if not name:
            self._log("DAG 流程编号不能为空")
            return False
        if not name.startswith("dag:"):
            name = "dag:" + name
        if self.state["busy"]:
            return False
        self._cancel_event.clear()
        self.state["cancelRequested"] = False
        self.state["busy"] = True
        self.state["status"] = "正在处理，请勿关闭窗口……"
        self._log(f"开始执行（DAG 编排）：{name[len('dag:'):]}")
        threading.Thread(target=self._worker, args=(name, strict), daemon=True).start()
        return True

    def cancel_execution(self) -> bool:
        """请求停止当前任务；不会强杀正在执行的 Office 调用。"""
        if not self.state.get("busy"):
            return False
        if self._cancel_event.is_set():
            return True
        self._cancel_event.set()
        self.state["cancelRequested"] = True
        self.state["status"] = "正在终止：将于当前文件/节点结束后停止"
        self._log("已请求终止：当前 Excel/WPS 操作完成后，将停止后续文件和节点；已生成文件会保留")
        return True

    def _recognize(
        self,
        *,
        force_template: bool = False,
        allow_template_auto: bool = False,
    ) -> None:
        input_dir = Path(self.state["input"]) if self.state["input"] else None
        template_dir = Path(self.state["templateDir"]) if self.state["templateDir"] else None
        if input_dir and input_dir.is_dir():
            period = detect_period(input_dir, recursive=self.state["recursiveDepth"])
            self.state["detectedPeriod"] = period.period
            self.state["explanationFiles"] = [
                {"path": str(path), "name": path.name}
                for path in explanation_files(
                    input_dir, recursive=self.state["recursiveDepth"]
                )
            ]
            if not self.state["outputPinned"]:
                self.state["output"] = str(input_dir / "执行结果")
                self.state["outputAuto"] = True
        elif self.state.get("explanationFiles"):
            self.state["explanationFiles"] = []
        if input_dir and template_dir and input_dir.is_dir() and template_dir.is_dir():
            files = classify_source_files(
                template_dir, input_dir, self.project_root / "data" / "模板索引.json",
                recursive=self.state["recursiveDepth"],
            )
            self.state["sourceFiles"] = files
            current = set(self.state.get("selectedFiles") or [])
            available = {item["path"] for item in files}
            if not self.state.get("selectedFilesInitialized", False):
                self.state["selectedFiles"] = [item["path"] for item in files]
                self.state["selectedFilesInitialized"] = True
            else:
                self.state["selectedFiles"] = [path for path in current if path in available]
            templates = sorted({item["template"] for item in files if item["template"] != "未识别"})
            self.state["mixedTemplates"] = templates
            # 自动匹配只能由两处触发：用户点击“自动识别模板”，或选择源数据目录。
            # 执行流程、刷新目录、选择模板目录等场景只更新清单，不得改写模板选择。
            should_recommend = force_template or (
                allow_template_auto and not self.state.get("templateManual")
            )
            if should_recommend:
                result = recommend_template(
                    template_dir, input_dir, self.project_root / "data" / "模板索引.json",
                    recursive=self.state["recursiveDepth"],
                )
                if result.template_path:
                    self.state["template"] = str(result.template_path)
                    self.state["templateManual"] = False
                self._log(result.details)
            if len(templates) > 1:
                self._log("发现多种报表：" + "、".join(templates) + "。建议取消勾选不属于本次审核类型的文件。")
            # 用户手动追加的文件（源数据目录之外）并入清单，同样参与模板匹配。
            extra = [Path(path) for path in self.state.get("extraFiles", []) if Path(path).is_file()]
            known = {item["path"] for item in files}
            fresh = [path for path in extra if str(path.resolve()) not in known and path.parent != input_dir]
            for path in fresh:
                rows = classify_source_files(
                    template_dir, path.parent, self.project_root / "data" / "模板索引.json",
                    recursive=False,
                )
                for row in rows:
                    if Path(row["path"]).resolve() == path.resolve() and row["path"] not in known:
                        files.append(row)
                        known.add(row["path"])

    def _worker(self, action: str, strict: bool = True) -> None:
        started = time.monotonic()
        try:
            # 执行前仅刷新待处理文件、数据期等状态，不重新匹配或改写模板。
            self._log("正在刷新待审核文件清单……")
            self._recognize(allow_template_auto=False)
            self._log(
                f"待处理文件：{len(self.state.get('selectedFiles', []))} 个；"
                "正在读取执行流程……"
            )
            self._save_settings()
            service = AuditService(
                config_path=self.history_path,
                engine_preference=str(self.state["calculationEngine"]),
                summary_read_engine=str(self.state["summaryReadEngine"]),
                conditional_format_evaluator=str(self.state.get("conditionalFormatEvaluator") or "COM_DISPLAY"),
                conditional_format_rule_reader=str(self.state.get("conditionalFormatRuleReader") or "DIRECT_OOXML"),
                formula_region_writer=str(self.state.get("formulaRegionWriter") or "COM_RANGE"),
                external_sheet_writer=str(self.state.get("externalSheetWriter") or "DIRECT_OOXML"),
                combine_sheets_plans=self.settings.combine_sheets_plans,
                active_combine_sheets_plan_id=self.settings.active_combine_sheets_plan_id,
            )
            output = Path(self.state["output"])
            selected = [Path(path) for path in self.state.get("selectedFiles", [])]
            if not selected:
                raise ValueError("请至少勾选一个待审核文件")
            if action == "check":
                result = service.preflight(template_path=Path(self.state["template"]), input_dir=Path(self.state["input"]), output_dir=output, selected_files=selected, external_path=Path(self.state["external"]) if self.state["external"] else None, extra_files=[Path(p) for p in self.state.get("extraFiles", [])], recursive=self.state["recursiveDepth"], on_step=self._log_detail)
            else:
                if not action.startswith("dag:"):
                    raise ValueError("旧版流程已移除，请使用 DAG 流程")
                workflow_id = action
                workflow_label = workflow_id[len("dag:"):]
                standalone = workflow_id in {"dag:组合联合核查表", "dag:组合工作表"}
                if standalone:
                    self._log(f"当前流程（DAG）：{workflow_label}；仅使用源数据目录")
                else:
                    self._log(f"当前流程（DAG）：{workflow_label}；模板：{Path(self.state['template']).name}")
                self._log("正在启动 DAG 工作流引擎……", detail=True)
                from .workflow.runner import run_dag_flow

                result = run_dag_flow(
                    workflow_id, service=service,
                    template_path=None if standalone else Path(self.state["template"]),
                    input_dir=Path(self.state["input"]), output_dir=output,
                    period="" if standalone else (self.state.get("detectedPeriod") or Path(self.state["input"]).name),
                    history_path=self.history_path, selected_files=selected,
                    external_path=None if standalone else (Path(self.state["external"]) if self.state["external"] else None),
                    extra_files=[] if standalone else [Path(p) for p in self.state.get("extraFiles", [])],
                    # DAG 节点可耗时数秒到数十秒；运行记录必须显示当前节点，
                    # 不能只作为“详细”日志被简单视图过滤掉。
                    recursive=self.state["recursiveDepth"], on_step=self._log, strict=strict,
                    write_flow_logs=bool(self.state["exportRunLogs"]),
                    cancel_event=self._cancel_event,
                )
            self.state["status"] = result.summary_text().splitlines()[0]
            self._log(result.summary_text())
        except Exception as exc:
            if self._cancel_event.is_set() and "用户请求终止" in str(exc):
                self.state["status"] = "任务已终止，已生成文件已保留"
                self._log("任务已终止，未执行的文件/节点未处理；已生成文件已保留")
            else:
                self.state["status"] = "执行失败"
                self._log("执行失败：" + str(exc))
        finally:
            elapsed = time.monotonic() - started
            self._log(f"任务结束，本次耗时：{elapsed:.1f} 秒")
            self.state["busy"] = False
            self.state["cancelRequested"] = False

    def win_minimize(self) -> None:
        """自绘标题栏的窗口控制：最小化。"""
        import webview
        if webview.windows:
            webview.windows[0].minimize()

    def win_maximize(self, restore: bool = False) -> None:
        """自绘标题栏的窗口控制：最大化 / 还原。restore=True 时还原。"""
        import webview
        if not webview.windows:
            return
        window = webview.windows[0]
        if restore:
            window.restore()
        else:
            window.maximize()

    def win_close(self) -> None:
        """自绘标题栏的窗口控制：关闭窗口。"""
        import webview
        if webview.windows:
            webview.windows[0].destroy()

    def _log(self, text: str, detail: bool = False) -> None:
        """记录一条运行日志。detail=False 为面向用户的简明结果，detail=True
        为逐功能/逐文件的调试细节，前端“简单”模式会过滤掉 detail 条目。"""
        stamp = datetime.now().strftime("%H:%M:%S")
        entry = {"text": f"[{stamp}] {text}", "detail": bool(detail)}
        self.state["log"] = (self.state["log"] + [entry])[-100:]

    def _log_detail(self, text: str) -> None:
        self._log(text, detail=True)

    def _save_settings(self) -> None:
        self.settings.last_input_dir = self.state["input"]
        self.settings.last_template_dir = self.state["templateDir"]
        self.settings.last_external_file = self.state["external"]
        self.settings.last_output_dir = self.state["output"]
        self.settings.output_pinned = bool(self.state["outputPinned"])
        self.settings.recursive_depth = int(self.state["recursiveDepth"])
        self.settings.recursive_folders = self.settings.recursive_depth != 0
        self.settings.show_run_detail_logs = bool(self.state.get("showRunDetailLogs", False))
        self.settings.export_run_logs = bool(self.state.get("exportRunLogs", False))
        self.settings.show_performance_diagnostics = bool(
            self.state.get("showPerformanceDiagnostics", False)
        )
        # 旧字段只作为兼容别名保存，规范语义为“导出运行日志 Excel”。
        self.settings.write_flow_logs = self.settings.export_run_logs
        self.settings.xlsx_render_mode = (
            str(self.state.get("xlsxRenderMode"))
            if self.state.get("xlsxRenderMode") in XLSX_RENDER_MODES
            else DEFAULT_XLSX_RENDER_MODE
        )
        self.settings.central_rule_engine = "v3"
        self.settings.expression_evaluation_mode = (
            str(self.state.get("expressionEvaluationMode"))
            if self.state.get("expressionEvaluationMode") in EXPRESSION_EVALUATION_MODES
            else "SBE"
        )
        self.settings.expression_evaluation_backend = (
            str(self.state.get("expressionEvaluationBackend"))
            if self.state.get("expressionEvaluationBackend") in EXPRESSION_EVALUATION_BACKENDS
            else "PYTHON"
        )
        self.settings.central_expression_schema = (
            "LEGACY_8"
            if str(self.state.get("centralExpressionSchema")) == "LEGACY_8"
            else "FIVE_SEGMENT_V1"
        )
        self.settings.confirm_before_run = bool(self.state["confirmBeforeRun"])
        self.settings.ui_theme = (
            str(self.state.get("uiTheme"))
            if self.state.get("uiTheme") in UI_THEMES
            else "日间"
        )
        self.settings.ui_accent = (
            str(self.state.get("uiAccent"))
            if self.state.get("uiAccent") in UI_ACCENTS
            else DEFAULT_UI_ACCENT
        )
        self.settings.file_picker_mode = (
            str(self.state.get("filePickerMode"))
            if self.state.get("filePickerMode") in FILE_PICKER_MODES
            else "系统原生"
        )
        self.settings.conditional_format_evaluator = (
            str(self.state.get("conditionalFormatEvaluator"))
            if self.state.get("conditionalFormatEvaluator") in CONDITIONAL_FORMAT_EVALUATORS
            else "PYTHON"
        )
        self.settings.conditional_format_rule_reader = (
            str(self.state.get("conditionalFormatRuleReader"))
            if self.state.get("conditionalFormatRuleReader") in CONDITIONAL_FORMAT_RULE_READERS
            else "DIRECT_OOXML"
        )
        self.settings.formula_region_writer = (
            str(self.state.get("formulaRegionWriter"))
            if self.state.get("formulaRegionWriter") in FORMULA_REGION_WRITERS
            else "DIRECT_OOXML"
        )
        self.settings.external_sheet_writer = (
            str(self.state.get("externalSheetWriter"))
            if self.state.get("externalSheetWriter") in EXTERNAL_SHEET_WRITERS
            else "DIRECT_OOXML"
        )
        self.settings.calculation_engine = str(self.state["calculationEngine"])
        self.settings.summary_read_engine = str(self.state["summaryReadEngine"])
        self.settings.period_current_dir = str(self.state.get("pcCurDir") or "")
        self.settings.period_previous_dir = str(self.state.get("pcPreDir") or "")
        self.settings.period_central_file = str(self.state.get("pcCentral") or "")
        self.settings.period_config_file = str(self.state.get("pcConfig") or "")
        self.settings.period_output_dir = str(self.state.get("pcOutput") or "")
        self.settings.period_output_auto = bool(self.state.get("pcOutputAuto", True))
        self.settings.central_diff_tolerance_yuan = float(self.state.get("centralTolerance") or 100.0)
        self.settings.central_current_csv = str(self.state.get("centralCur") or "")
        self.settings.central_previous_csv = str(self.state.get("centralPre") or "")
        self.settings.central_output_dir = str(self.state.get("centralOutput") or "")
        self.settings.central_output_auto = bool(self.state.get("centralOutputAuto", True))
        self.settings.central_cross_current_csv = str(self.state.get("centralCrossCur") or "")
        self.settings.central_cross_output_dir = str(self.state.get("centralCrossOutput") or "")
        self.settings.central_cross_output_auto = bool(self.state.get("centralCrossOutputAuto", True))
        self.settings.central_compare_file = str(self.state.get("centralCompareFile") or "")
        self.settings.central_explanation_file = str(self.state.get("centralExplainFile") or "")
        self.settings.central_template_file = str(self.state.get("centralTemplate") or "")
        self.settings.central_form_output_dir = str(self.state.get("centralFormOutput") or "")
        self.settings.central_form_output_auto = bool(self.state.get("centralFormOutputAuto", True))
        self.settings.central_common_config = str(self.state.get("centralCommonConfig") or "")
        self.settings.central_comparison_config = str(self.state.get("centralComparisonConfig") or "")
        self.settings.central_cross_config = str(self.state.get("centralCrossConfig") or "")
        self.settings.central_forms_config = str(self.state.get("centralFormsConfig") or "")
        self.settings_store.save(self.settings)


def launch_web(project_root: Path) -> None:
    bundle_root = Path(getattr(sys, "_MEIPASS", project_root))
    # Python.NET 3 normally discovers the interpreter DLL from a regular
    # Python installation.  In a PyInstaller one-file build (in particular
    # the Python 3.7 Win7 package) that DLL lives in the temporary _MEI
    # directory instead.  Tell the CLR bridge its exact location before
    # pywebview imports ``clr``; otherwise pywebview hides the real loader
    # error behind the misleading "pythonnet is not installed" message.
    if getattr(sys, "frozen", False) and sys.platform == "win32":
        python_dll = bundle_root / ("python{}{}.dll".format(sys.version_info[0], sys.version_info[1]))
        if python_dll.is_file():
            os.environ.setdefault("PYTHONNET_PYDLL", str(python_dll))
        os.environ.setdefault("PYTHONNET_RUNTIME", "netfx")

        # pywebview turns every WinForms import failure into the generic
        # "pythonnet is not installed" message.  Verify the bridge first and
        # leave a diagnosable report beside the EXE when an old Windows system
        # cannot load one of its native/CLR dependencies.
        try:
            import clr

            clr.AddReference("System.Windows.Forms")
        except Exception as exc:
            report = project_root / "Win7界面启动诊断.txt"
            details = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
            try:
                report.write_text(
                    "基础数据审核工具界面启动诊断\n"
                    "Python.NET / WinForms 初始化失败。\n\n"
                    + details,
                    encoding="utf-8",
                )
            except OSError:
                pass
            raise RuntimeError(
                "Windows 界面组件初始化失败。请将 EXE 同目录的“Win7界面启动诊断.txt”发给维护人员。"
            ) from exc

    import webview

    initialize_history_workbook(release_config_dir(project_root) / HISTORY_WORKBOOK_NAME)
    hide_application_data_directory(project_root / "data")
    html = bundle_root / "web" / "index.html" if getattr(sys, "frozen", False) else project_root / "frontend" / "web" / "index.html"
    if not html.is_file():
        raise RuntimeError("本地界面文件缺失")
    page = _bridge_alias_page(html)
    # 使用操作系统原生标题栏（最小化/最大化/关闭由系统提供），界面内不再自绘。
    webview.create_window(app_title(), page.as_uri(), js_api=WebApi(project_root), width=1180, height=820, min_size=(900, 650), frameless=False)
    webview.start(gui="edgechromium", icon=str(icon_path(".ico")))


def _bridge_alias_page(html: Path) -> Path:
    """为真 pywebview 窗口生成带别名垫片的临时页面。

    前端统一调用 ``bridge.api.方法名(...)`` 并等待 ``bridgeready``；pywebview
    原生只提供 ``window.pywebview`` 和 ``pywebviewready``。这里在 <head> 注入
    别名脚本把两者桥接起来，共用页面文件本身保持与 shell-flask 完全一致。
    """
    import tempfile

    source = render_app_page(html.read_text(encoding="utf-8"))
    if "</head>" not in source:
        raise RuntimeError("本地界面文件缺失 <head>，无法注入桥接别名")
    alias = (
        "<script>window.addEventListener('pywebviewready',function(){"
        "window.bridge=window.pywebview;"
        "window.dispatchEvent(new Event('bridgeready'));"
        "});</script>"
    )
    handle, name = tempfile.mkstemp(prefix="audit_bridge_", suffix=".html")
    import os as _os

    with _os.fdopen(handle, "w", encoding="utf-8") as stream:
        stream.write(source.replace("</head>", alias + "</head>", 1))
    return Path(name)
