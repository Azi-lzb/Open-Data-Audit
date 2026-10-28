"""Additive WebApi mixin. No S1 form fields or worker methods are reused."""
from __future__ import annotations

import copy
import json
import threading
import time
import traceback
import uuid
from datetime import datetime
from pathlib import Path

from base_audit.node_flow_config import release_config_dir
from s1audit_configured.errors import ConfigurationError
from s1audit_configured.runtime import AuditCancelled

from .contracts import S5Request
from .expression_export import export_expressions
from .formula_candidates import export_formula_candidates
from .service import run_s5_f1, validate_s5_config
from .recognition import classify_s5_files, file_period_hint, recognize_s5


CONFIG_NAME = "5.逐笔统计系统_审核配置.xlsx"
LEGACY_CONFIG_NAMES = ("S5_逐笔统计系统_审核配置.xlsx",)
LEGACY_TEMPLATE_DIR_NAME = "S5_模板文件"
_PATH_FIELDS = {"input_dir", "template_dir", "template_path", "config_path", "external_path", "output_dir"}
_BOOL_FIELDS = {"template_manual", "period_manual", "business_manual", "output_pinned"}
_FORM_FIELDS = _PATH_FIELDS | _BOOL_FIELDS | {"business", "period", "external_period", "previous_files", "previous_period", "identity_sheet", "identity_cell", "selected_files"}
_S5_TOOL_KEYS = {"vcur", "vext", "candidates"}
_S5_TOOL_FIELDS = {
    "vcur": {"outputDir"},
    "vext": {"outputDir", "unit", "sourceId"},
    "candidates": {"outputDir", "business", "unit"},
}
_S5_TOOL_PICKERS = {
    "vcur_output": ("vcur", "outputDir", True, False),
    "vext_output": ("vext", "outputDir", True, False),
    "candidate_output": ("candidates", "outputDir", True, False),
}


class S5WebApiMixin:
    def _init_s5(self):
        config_dir = release_config_dir(self.project_root)
        self._s5_settings_path = self.project_root / "data" / "s5" / "界面设置.json"
        saved = {}
        if self._s5_settings_path.is_file():
            try:
                saved = json.loads(self._s5_settings_path.read_text(encoding="utf-8"))
                if not isinstance(saved, dict):
                    saved = {}
            except (OSError, ValueError):
                self._log("S5 界面设置无法读取，本次使用默认路径；原文件未修改")
        default_config_path = next(
            (config_dir / name for name in (CONFIG_NAME, *LEGACY_CONFIG_NAMES)
             if (config_dir / name).is_file()),
            config_dir / CONFIG_NAME,
        )
        named_range_templates = config_dir / "模板文件_命名区域"
        legacy_templates = config_dir / LEGACY_TEMPLATE_DIR_NAME
        default_template_dir = (
            named_range_templates if named_range_templates.is_dir()
            else legacy_templates if legacy_templates.is_dir()
            else None
        )
        external_default = config_dir / "外部文件-大集中相关数据.xlsx"
        resolver = getattr(self, "_external_reference_config_path", None)
        if callable(resolver):
            try:
                external_default = resolver()
            except RuntimeError:
                pass
        elif not external_default.is_file():
            candidates = sorted(
                path for path in config_dir.glob("外部文件-大集中相关数据*.xlsx")
                if path.is_file()
            )
            if candidates:
                external_default = candidates[0]

        form = {key: "" for key in _FORM_FIELDS}
        form.update(config_path=str(default_config_path),
                    template_dir=str(default_template_dir) if default_template_dir else "",
                    external_path=str(external_default) if external_default.is_file() else "",
                    previous_files=[], selected_files=[])
        form.update({key: False for key in _BOOL_FIELDS})
        form.update({key: value for key, value in saved.items() if key in _FORM_FIELDS})
        # Detected values are session results. Keep only values the user
        # explicitly fixed; do not silently recognize templates on startup.
        for field, flag in (("template_path", "template_manual"),
                            ("period", "period_manual"),
                            ("business", "business_manual")):
            if not form.get(flag):
                form[field] = ""
        saved_config_path = str(form.get("config_path") or "")
        if saved_config_path and default_config_path.is_file():
            try:
                saved_config = Path(saved_config_path).expanduser()
                is_release_config = saved_config.resolve().parent == config_dir.resolve()
            except (OSError, RuntimeError):
                is_release_config = False
            if is_release_config and saved_config.name in LEGACY_CONFIG_NAMES:
                form["config_path"] = str(default_config_path)
        saved_template_dir = str(form.get("template_dir") or "")
        if default_template_dir and saved_template_dir:
            try:
                saved_template = Path(saved_template_dir).expanduser()
                is_legacy_release_template = (
                    saved_template.name == LEGACY_TEMPLATE_DIR_NAME
                    and saved_template.resolve().parent == config_dir.resolve()
                )
            except (OSError, RuntimeError):
                is_legacy_release_template = False
            if is_legacy_release_template:
                form["template_dir"] = str(default_template_dir)
        saved_template_dir = str(form.get("template_dir") or "")
        if not saved_template_dir or not Path(saved_template_dir).is_dir():
            form["template_dir"] = (
                str(default_template_dir) if default_template_dir else saved_template_dir
            )
        if form.get("input_dir") and not form.get("output_pinned"):
            input_root = Path(form["input_dir"]).expanduser()
            output_path = Path(form.get("output_dir") or "").expanduser()
            try:
                is_legacy_default = (
                    output_path.name == "S5_执行结果"
                    and output_path.resolve().parent == input_root.resolve()
                )
            except (OSError, RuntimeError):
                is_legacy_default = False
            if is_legacy_default or not str(form.get("output_dir") or "").strip():
                form["output_dir"] = str(input_root / "执行结果")
        saved_external = str(form.get("external_path") or "")
        if not saved_external or not Path(saved_external).is_file():
            form["external_path"] = (
                str(external_default) if external_default.is_file() else saved_external
            )
        form.update(files=[], configuration={}, task={}, progress={}, configurationCheckTask={},
                    recognition={"errors": {}})
        self.state["s5"] = form
        self._s5_config_check_lock = threading.RLock()
        self._s5_tool_settings_path = self.project_root / "data" / "s5" / "维护工具设置.json"
        defaults = {
            "vcur": {"outputDir": str(self.project_root / "data" / "s5" / "维护输出" / "表达式参考")},
            "vext": {"outputDir": str(self.project_root / "data" / "s5" / "维护输出" / "表达式参考"), "unit": "万元", "sourceId": "EXT"},
            "candidates": {"outputDir": str(self.project_root / "data" / "s5" / "维护输出" / "配置候选"), "business": "", "unit": "万元"},
        }
        if self._s5_tool_settings_path.is_file():
            try:
                stored = json.loads(self._s5_tool_settings_path.read_text(encoding="utf-8"))
                if isinstance(stored, dict):
                    for name in _S5_TOOL_KEYS:
                        if isinstance(stored.get(name), dict):
                            for key, value in stored[name].items():
                                if key not in _S5_TOOL_FIELDS[name]:
                                    continue
                                if key == "unit":
                                    defaults[name][key] = value if isinstance(value, str) and value in {"元", "万元", "亿元"} else "万元"
                                else:
                                    defaults[name][key] = str(value or "")
            except (OSError, ValueError):
                self._log("S5 配置维护工具设置无法读取，本次使用默认路径；原文件未修改")
        form["maintenanceTools"] = defaults
        form["maintenanceTask"] = {}
        self._s5_scan_files()

    def _s5_recognize(self):
        form = self.state["s5"]
        form["recognition"] = recognize_s5(form, self.project_root / "data")
        return form["recognition"]

    def recognize_s5_inputs(self, preserve_template=False):
        form = self.state["s5"]
        if preserve_template:
            if not form.get("template_path"):
                raise ValueError("请先手动选择匹配模板")
            form["template_manual"] = True
        else:
            form["template_manual"] = False
            form["template_path"] = ""
        self._s5_recognize()
        self._s5_annotate_files()
        return form

    def _s5_annotate_files(self):
        form = self.state["s5"]
        form["files"] = classify_s5_files(
            [item["path"] for item in form.get("files", [])],
            template_dir=form.get("template_dir") or "",
            template_path=form.get("template_path") or "",
            config_path=form.get("config_path") or "",
            input_dir=form.get("input_dir") or "",
            index_path=Path(self.project_root) / "data" / "s5" / "模板索引.json",
            selected_files=form.get("selected_files", []),
            business=form.get("business") or "",
            period=form.get("period") or "",
        )

    def _s5_scan_files(self):
        form = self.state["s5"]
        cached = {item.get("path"): item for item in form.get("files", []) if item.get("path")}
        root = Path(form["input_dir"]) if form.get("input_dir") else None
        paths = sorted((p for p in root.iterdir() if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm"}
                        and not p.name.startswith("~$")), key=lambda p: p.name) if root and root.is_dir() else []
        rows = []
        for path in paths:
            resolved = str(path.resolve())
            row = dict(cached.get(resolved) or {})
            row.update(path=resolved, name=path.name)
            row.setdefault("template", "未识别")
            row.setdefault("match_status", "pending")
            row.setdefault("business", "")
            row["period"] = file_period_hint(path, form.get("input_dir") or "")
            rows.append(row)
        form["files"] = rows
        allowed = {item["path"] for item in form["files"]}
        form["selected_files"] = [p for p in form.get("selected_files", []) if p in allowed]

    def update_s5(self, values):
        form = self.state["s5"]
        previous_input = form.get("input_dir")
        previous_template_dir = form.get("template_dir")
        previous_template_path = form.get("template_path")
        previous_selected = list(form.get("selected_files") or [])
        for key, value in values.items():
            if key not in _FORM_FIELDS:
                continue
            if key in _BOOL_FIELDS:
                form[key] = value if isinstance(value, bool) else str(value).lower() in {"true", "1", "是"}
            elif key in {"previous_files", "selected_files"}:
                if not isinstance(value, list):
                    raise ValueError("文件清单必须为路径列表")
                form[key] = list(dict.fromkeys(str(p) for p in value))
            else:
                form[key] = str(value or "").strip()
        self._s5_scan_files()
        if form.get("input_dir") != previous_input:
            form["selected_files"] = [item["path"] for item in form["files"]]
            if not form.get("output_pinned") and "output_dir" not in values:
                form["output_dir"] = str(Path(form["input_dir"]) / "执行结果") if form["input_dir"] else ""
            for field, flag in [("template_path", "template_manual"), ("period", "period_manual"), ("business", "business_manual")]:
                if field not in values:
                    form[field], form[flag] = "", False
            form["recognition"] = {"errors": {}}
        if form.get("template_dir") != previous_template_dir and "template_path" not in values:
            form["template_path"], form["template_manual"] = "", False
            for item in form["files"]:
                item.update(template="未识别", match_status="pending", business="", match_detail="")
        for field, flag in [("template_path", "template_manual"), ("period", "period_manual"), ("business", "business_manual")]:
            if field in values:
                form[flag] = bool(form[field])
        if form.get("template_path") != previous_template_path:
            form["recognition"] = {"errors": {}}
            for item in form["files"]:
                item.update(template="未识别", match_status="pending", business="", match_detail="")
        if "selected_files" in values and form.get("selected_files") != previous_selected:
            for field, flag in (("period", "period_manual"), ("business", "business_manual")):
                if not form.get(flag):
                    form[field] = ""
            form["recognition"] = {"errors": {}}
        if "output_dir" in values and "output_pinned" not in values:
            form["output_pinned"] = bool(form["output_dir"])
        if values.get("output_pinned") is False and form.get("input_dir"):
            form["output_dir"] = str(Path(form["input_dir"]) / "执行结果")
        if "config_path" in values:
            form["configuration"] = {}
            if "business" not in values:
                form["business"] = ""
                form["business_manual"] = False
            form["recognition"] = {"errors": {}}
            for item in form["files"]:
                item["business"] = ""
        self._s5_settings_path.parent.mkdir(parents=True, exist_ok=True)
        self._s5_settings_path.write_text(json.dumps({key: form[key] for key in _FORM_FIELDS}, ensure_ascii=False, indent=2), encoding="utf-8")
        return form

    def choose_s5(self, field, path=""):
        if field not in _PATH_FIELDS:
            raise ValueError("未知的 S5 文件选择项")
        current = self.state["s5"].get(field) or str(self.project_root)
        folder = field in {"input_dir", "template_dir", "output_dir"}
        chosen = (path,) if path else self._select_path(folder=folder, current=current,
                                                     file_types=(("Excel 文件", "*.xlsx *.xlsm"),))
        if chosen:
            form = self.update_s5({field: chosen[0]})
            if field == "template_path":
                self._s5_recognize()
                self._s5_annotate_files()
            return form
        return self.state["s5"]

    def check_s5_config(self, scope="current_business"):
        scope = str(scope or "current_business").strip()
        if scope not in {"current_business", "all_enabled"}:
            return {"accepted": False, "message": f"未知的 S5 配置检查范围：{scope}"}
        with self._execution_start_lock:
            form = self.state["s5"]
            path = form.get("config_path")
            if not path or not Path(path).is_file():
                raise ValueError("请选择 S5 审核配置文件")
            if scope == "current_business":
                template = str(form.get("template_path") or "").strip()
                if not template or not Path(template).is_file():
                    return {"accepted": False, "message": "当前业务检查需要先点击“识别模板”或手动选择匹配模板"}
            elif not form.get("template_dir") or not Path(form["template_dir"]).is_dir():
                return {"accepted": False, "message": "全量检查需要有效的模板文件目录"}
            with self._s5_config_check_lock:
                current = form.get("configurationCheckTask") or {}
            if current.get("task_state") in {"queued", "running"}:
                if current.get("scope") == scope:
                    return {"accepted": True, "task_id": current.get("task_id"), "message": "S5 配置检查正在运行"}
                running_scope = "当前业务" if current.get("scope") == "current_business" else "全量启用规则"
                return {"accepted": False, "message": f"S5 {running_scope}检查正在运行，请完成后再启动另一种检查"}
            if self._execution_busy():
                return {"accepted": False, "message": self.state.get("status") or "其他任务正在运行，请完成后再检查 S5 配置"}
            task_id = uuid.uuid4().hex
            snapshot = copy.deepcopy(form)
            snapshot["configurationCheckScope"] = scope
            with self._s5_config_check_lock:
                form["configurationCheckTask"] = {
                    "task_id": task_id, "task_state": "queued", "scope": scope,
                    "message": "正在启动 S5 配置检查…",
                    "progress": {"phase": "queued", "percent": 0, "current": 0, "total": 0,
                                 "message": "准备读取审核配置"},
                    "started_at_epoch": time.time(),
                }
            self.state["status"] = "S5 配置检查正在运行"
            scope_label = "当前业务" if scope == "current_business" else "全量启用规则"
            self._log(f"S5 检查配置开始（{scope_label}）：{Path(path).name}；只检查启用规则")
            thread = threading.Thread(
                target=self._s5_config_check_worker, args=(task_id, snapshot), daemon=True,
            )
            thread.start()
            return {"accepted": True, "task_id": task_id, "message": "S5 配置检查已启动"}

    def get_s5_config_check(self):
        with self._s5_config_check_lock:
            task = (self.state.get("s5") or {}).get("configurationCheckTask") or {}
            return copy.deepcopy(task)

    @staticmethod
    def _s5_config_check_output_dir(form):
        if form.get("output_dir"):
            base = Path(form["output_dir"]).expanduser().resolve()
            if base.name == "S5_执行结果" and not form.get("output_pinned"):
                return base.parent / "执行结果"
            return base
        if form.get("input_dir"):
            return Path(form["input_dir"]).expanduser().resolve() / "执行结果"
        if form.get("config_path"):
            return Path(form["config_path"]).expanduser().resolve().parent / "执行结果"
        return Path(".").resolve() / "执行结果"

    @staticmethod
    def _format_s5_config_check_report(report):
        lines = [
            "S5 配置检查报告",
            f"检查范围：{report.get('scope_label', '当前业务启用配置')}",
            f"检查业务：{report.get('business') or ('全部业务' if report.get('scope') in {'all', 'all_enabled'} else '（未识别）')}",
            f"配置文件：{report.get('path', '')}",
            f"配置版本：{report.get('version', '')}（{report.get('version_status', '')}）",
            f"规则总数：{report.get('rules', 0)}；启用：{report.get('enabled_rules', 0)}；禁用：{report.get('disabled_rules', 0)}",
            f"检查结论：{'有效' if report.get('valid') else '发现问题'}；{'完整' if report.get('complete') else '有项目未检查'}",
            "",
            str(report.get("message") or ""),
            "",
            "分项检查：",
        ]
        for key, title in (("expressions", "审核表达式"), ("dictionary_keys", "字典键引用"),
                           ("template_addresses", "模板地址"), ("external_addresses", "外部文件地址")):
            check = (report.get("checks") or {}).get(key) or {}
            extra = ""
            if "checked" in check:
                extra = f"；定位 {check.get('resolved', 0)}/{check.get('checked', 0)}"
            elif "rules" in check:
                extra = f"；规则 {check.get('rules', 0)} 条"
            lines.append(f"- {title}：{check.get('status', 'not_checked')}{extra}")
            if check.get("message"):
                lines.append(f"  {check['message']}")
        business_checks = report.get("business_checks") or []
        if business_checks:
            lines.extend(["", "逐业务模板地址检查："])
            for item in business_checks:
                check = item.get("template_addresses") or {}
                lines.append(
                    f"- {item.get('name', '')}：启用 {item.get('enabled_rules', 0)} 条；"
                    f"模板 {item.get('template') or '未匹配'}；地址 {check.get('status', 'not_checked')} "
                    f"（定位 {check.get('resolved', 0)}/{check.get('checked', 0)}）"
                )
                if check.get("message") and check.get("status") in {"not_checked", "partial", "failed"}:
                    lines.append(f"  {check['message']}")
        warnings = report.get("warnings") or []
        if warnings:
            lines.extend(["", "提示：", *[f"- {item}" for item in warnings]])
        errors = report.get("configuration_errors") or []
        if errors:
            lines.extend(["", f"问题明细（{len(errors)} 项）："])
            for index, issue in enumerate(errors, 1):
                locations = issue.get("locations") or []
                where = "；".join(
                    " / ".join(str(location.get(key) or "") for key in ("sheet", "cell", "rule_id") if location.get(key))
                    for location in locations
                )
                if issue.get("business"):
                    where = "业务“" + str(issue["business"]) + "”" + (" / " + where if where else "")
                lines.append(f"{index}. {(where + '：') if where else ''}{issue.get('message', '')}")
        return "\n".join(lines) + "\n"

    @staticmethod
    def _s5_config_check_report_summary(report):
        checks = {}
        for key, value in (report.get("checks") or {}).items():
            summary = {name: item for name, item in value.items() if name != "errors"}
            if isinstance(summary.get("businesses"), list):
                summary["businesses"] = [
                    {name: item for name, item in business.items() if name != "errors"}
                    for business in summary["businesses"]
                ]
            checks[key] = summary
        errors = report.get("configuration_errors") or []
        return {
            "valid": bool(report.get("valid")), "complete": bool(report.get("complete")),
            "scope": report.get("scope", ""), "scope_label": report.get("scope_label", ""),
            "business": report.get("business", ""),
            "message": report.get("message", ""), "rules": report.get("rules", 0),
            "enabled_rules": report.get("enabled_rules", 0), "disabled_rules": report.get("disabled_rules", 0),
            "checks": checks, "warnings": list(report.get("warnings") or []),
            "business_checks": [
                {key: value for key, value in item.items() if key != "errors"}
                for item in report.get("business_checks", [])
            ],
            "configuration_errors": errors[:30], "error_count": len(errors),
        }

    def _s5_config_check_worker(self, task_id, form):
        started = time.monotonic()
        live_form = self.state["s5"]
        task = live_form["configurationCheckTask"]
        scope = form.get("configurationCheckScope") or "current_business"
        scope_label = "当前业务" if scope == "current_business" else "全量启用规则"
        with self._s5_config_check_lock:
            task.update(task_state="running", scope=scope, message=f"正在检查 S5 {scope_label}…")
        progress_state = {"phase": "", "bucket": -1}

        def report_progress(event):
            event = dict(event or {})
            percent = max(0, min(99, float(event.get("percent", 0) or 0)))
            with self._s5_config_check_lock:
                task["progress"] = {**event, "percent": percent,
                                    "elapsed_seconds": round(time.monotonic() - started, 1)}
                task["message"] = event.get("message") or "正在检查 S5 配置…"
                current_message = task["message"]
            bucket = int(percent // 10)
            if event.get("phase") != progress_state["phase"] or bucket > progress_state["bucket"]:
                self._log(f"S5 检查配置：{current_message}")
                progress_state.update(phase=event.get("phase", ""), bucket=bucket)

        report = None
        report_path = ""
        try:
            report = validate_s5_config(
                form.get("config_path") or "",
                template_path=form.get("template_path") or "",
                template_dir=form.get("template_dir") or "",
                template_index_path=str(Path(self.project_root) / "data" / "s5" / "模板索引.json"),
                external_path=form.get("external_path") or "",
                business=form.get("business") or "",
                scope=scope,
                on_progress=report_progress,
            )
            task_state = "failed" if not report.get("valid") else "succeeded" if report.get("complete") else "incomplete"
            try:
                output_dir = self._s5_config_check_output_dir(form)
                output_dir.mkdir(parents=True, exist_ok=True)
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                report_file = output_dir / f"S5配置检查_{scope_label}_{stamp}_{task_id[:8]}.txt"
                report_file.write_text(self._format_s5_config_check_report(report), encoding="utf-8")
                report_path = str(report_file)
            except OSError as exc:
                report.setdefault("warnings", []).append(f"检查结果已显示，但报告文件保存失败：{exc}")
                self._log(f"S5 配置检查报告保存失败：{exc}")
            live_form["configuration"] = self._s5_config_check_report_summary(report)
            live_form["recognition"] = form.get("recognition") or live_form.get("recognition") or {}
            with self._s5_config_check_lock:
                task.update(
                    task_id=task_id, task_state=task_state, message=report.get("message", "S5 配置检查完成"),
                    report=self._s5_config_check_report_summary(report), report_path=report_path,
                    elapsed_seconds=round(time.monotonic() - started, 2),
                )
            self.state["status"] = {
                "succeeded": "S5 配置检查完成", "incomplete": "S5 配置检查完成（有项目未检查）",
                "failed": "S5 配置检查发现问题",
            }[task_state]
            self._log(f"{self.state['status']}；报告：{report_path or '未能保存'}")
        except Exception as exc:
            report = {
                "valid": False, "complete": False, "message": f"S5 配置检查失败：{exc}",
                "configuration_errors": [{"message": str(exc), "locations": []}], "warnings": [],
                "checks": {}, "rules": 0, "enabled_rules": 0, "disabled_rules": 0,
            }
            try:
                output_dir = self._s5_config_check_output_dir(form)
                output_dir.mkdir(parents=True, exist_ok=True)
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                report_file = output_dir / f"S5配置检查_{stamp}_{task_id[:8]}.txt"
                report_file.write_text(self._format_s5_config_check_report(report), encoding="utf-8")
                report_path = str(report_file)
            except OSError as save_exc:
                self._log(f"S5 配置检查失败报告保存失败：{save_exc}")
            live_form["configuration"] = self._s5_config_check_report_summary(report)
            with self._s5_config_check_lock:
                task.update(task_id=task_id, task_state="failed", message=report["message"],
                            report=self._s5_config_check_report_summary(report), report_path=report_path,
                            elapsed_seconds=round(time.monotonic() - started, 2))
            self.state["status"] = "S5 配置检查失败"
            self._log(report["message"])
            self._log(traceback.format_exc(), detail=True)
        finally:
            with self._s5_config_check_lock:
                current_progress = task.get("progress") or {}
                task["progress"] = {**current_progress, "percent": 100,
                                     "elapsed_seconds": round(time.monotonic() - started, 1),
                                     "message": self.state.get("status", "S5 配置检查完成")}
            self._log(f"S5 配置检查任务结束，耗时 {time.monotonic() - started:.1f} 秒")

    def update_s5_maintenance(self, tool, values):
        if tool not in _S5_TOOL_KEYS or not isinstance(values, dict):
            raise ValueError("未知的 S5 配置维护工具或参数格式")
        settings = self.state["s5"]["maintenanceTools"][tool]
        for key, value in values.items():
            if key not in _S5_TOOL_FIELDS[tool]:
                continue
            if key == "unit":
                value = str(value or "万元").strip()
                if value not in {"元", "万元", "亿元"}:
                    raise ValueError("金额单位仅支持元、万元或亿元")
                settings[key] = value
            else:
                settings[key] = str(value or "").strip()
        self._s5_tool_settings_path.parent.mkdir(parents=True, exist_ok=True)
        self._s5_tool_settings_path.write_text(
            json.dumps(self.state["s5"]["maintenanceTools"], ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return self.state["s5"]["maintenanceTools"][tool]

    def choose_s5_maintenance(self, field, paths=None):
        spec = _S5_TOOL_PICKERS.get(field)
        if spec is None:
            raise ValueError("未知的 S5 配置维护路径")
        tool, setting, folder, multiple = spec
        current = self.state["s5"]["maintenanceTools"][tool].get(setting)
        current_path = (current[0] if isinstance(current, list) and current else current) or str(self.project_root)
        if not folder and Path(str(current_path)).is_file():
            current_path = str(Path(str(current_path)).parent)
        if folder:
            probe = Path(str(current_path)).expanduser()
            while not probe.is_dir() and probe.parent != probe:
                probe = probe.parent
            current_path = str(probe if probe.is_dir() else self.project_root)
        if paths is not None:
            chosen = tuple(str(item) for item in paths) if isinstance(paths, (list, tuple)) else (str(paths),)
        else:
            chosen = self._select_path(
                folder=folder,
                current=str(current_path),
                multiple=multiple,
                file_types=(
                    ("Excel 工作簿", "*.xlsx *.xlsm"),
                ),
            )
        if not chosen:
            return self.state["s5"]["maintenanceTools"][tool]
        value = list(chosen) if multiple else str(chosen[0])
        return self.update_s5_maintenance(tool, {setting: value})

    def start_s5_maintenance(self, tool, values=None):
        if tool not in _S5_TOOL_KEYS:
            raise ValueError("未知的 S5 配置维护工具")
        with self._execution_start_lock:
            if self._execution_busy():
                return {"accepted": False, "message": "已有任务正在运行，请等待完成"}
            if values:
                self.update_s5_maintenance(tool, values)
            settings = copy.deepcopy(self.state["s5"]["maintenanceTools"][tool])
            form = self.state["s5"]
            if tool in {"vcur", "candidates"}:
                template_text = str(form.get("template_path") or "").strip()
                if not template_text:
                    raise ValueError("请先在 S5 面板选择源数据并匹配模板")
                template_path = Path(template_text).expanduser().resolve()
                if not template_path.is_file():
                    raise ValueError("S5 面板当前匹配模板不存在，请重新选择或识别模板")
                if template_path.suffix.lower() not in {".xlsx", ".xlsm"}:
                    raise ValueError("S5 面板匹配模板仅支持 XLSX/XLSM 工作簿")
                paths = [template_path]
                if tool == "vcur" and template_path.suffix.lower() != ".xlsx":
                    raise ValueError("V_CUR / V_PRE 表达式参考目前仅支持 XLSX 模板")
            else:
                external_text = str(form.get("external_path") or "").strip()
                if not external_text:
                    raise ValueError("请先在 S5 面板选择大集中系统数据")
                external_path = Path(external_text).expanduser().resolve()
                if not external_path.is_file():
                    raise ValueError("S5 面板当前选择的大集中系统数据不存在，请重新选择")
                if external_path.suffix.lower() != ".xlsx":
                    raise ValueError("V_EXT 表达式参考目前仅支持 XLSX 外部文件")
                paths = [external_path]
            output_dir = Path(settings.get("outputDir") or "").expanduser().resolve()
            if not str(settings.get("outputDir") or "").strip():
                raise ValueError("请选择输出目录")
            external_path = None
            if tool == "candidates" and str(form.get("external_path") or "").strip():
                external_path = Path(form["external_path"]).expanduser().resolve()
                if not external_path.is_file() or external_path.suffix.lower() not in {".xlsx", ".xlsm"}:
                    raise ValueError("S5 面板当前大集中系统数据不存在或格式不支持，请重新选择")
            if tool == "vext" and settings.get("unit") not in {"元", "万元", "亿元"}:
                raise ValueError("外部金额单位仅支持元、万元或亿元")
            if tool == "candidates" and settings.get("unit") not in {"元", "万元", "亿元"}:
                raise ValueError("外部金额单位仅支持元、万元或亿元")
            if tool == "vcur":
                settings.update(files=[str(path) for path in paths], outputDir=str(output_dir),
                                template_dir=str(self.state["s5"].get("template_dir") or ""),
                                layoutTemplate=str(paths[0]))
            elif tool == "vext":
                settings.update(files=[str(path) for path in paths], outputDir=str(output_dir))
            else:
                settings.update(files=[str(path) for path in paths], outputDir=str(output_dir),
                                externalFile=str(external_path or ""))
            task_id = uuid.uuid4().hex
            self.state["status"] = f"S5 配置维护工具 {tool} 正在运行"
            self.state["s5"]["maintenanceTask"] = {
                "task_id": task_id, "tool": tool, "task_state": "running",
                "message": "正在生成配置维护副本；输入文件保持只读。", "started_at_epoch": time.time(),
            }
            threading.Thread(target=self._s5_maintenance_worker, args=(task_id, tool, settings), daemon=True).start()
            return {"accepted": True, "task_id": task_id}

    def _s5_maintenance_worker(self, task_id, tool, settings):
        started = time.monotonic()
        labels = {"vcur": "V_CUR / V_PRE 表达式参考", "vext": "V_EXT 外部表达式", "candidates": "校验公式转配置候选"}
        files = settings.get("files", [])
        self._log(f"S5 配置维护：{labels[tool]} 开始；输入 {len(files)} 个工作簿")
        task = self.state["s5"]["maintenanceTask"]
        try:
            if tool == "vcur":
                result = export_expressions(
                    files, mode="vcur", output_root=settings["outputDir"],
                    template=settings.get("layoutTemplate") or None,
                    template_dir=settings.get("template_dir") or None,
                )
                state = "failed" if result.get("status") == "FAILED" else "partial_failed" if result.get("status") == "COMPLETED_WITH_ISSUES" else "succeeded"
                summary = result.get("counts", {})
                message = f"参考副本 {len(result.get('files', []))} 份；可用表达式 {summary.get('ready', 0)} 项；待处理或跳过 {summary.get('failed', 0) + summary.get('skipped', 0)} 项。"
            elif tool == "vext":
                result = export_expressions(
                    files, mode="vext", output_root=settings["outputDir"],
                    external_unit=settings.get("unit", "万元"), source_id=settings.get("sourceId", "EXT"),
                )
                state = "failed" if result.get("status") == "FAILED" else "partial_failed" if result.get("status") == "COMPLETED_WITH_ISSUES" else "succeeded"
                summary = result.get("counts", {})
                message = f"参考副本 {len(result.get('files', []))} 份；可用表达式 {summary.get('ready', 0)} 项；待处理或跳过 {summary.get('failed', 0) + summary.get('skipped', 0)} 项。"
            else:
                result = export_formula_candidates(
                    files, output_root=settings["outputDir"], business=settings.get("business", ""),
                    external=settings.get("externalFile") or None, external_unit=settings.get("unit", "万元"),
                )
                state = "partial_failed" if result.get("counts", {}).get("pending", 0) else "succeeded"
                summary = result.get("counts", {})
                message = f"候选规则 {summary.get('candidates', 0)} 条，待人工处理 {summary.get('pending', 0)} 项；候选规则默认禁用。"
            task.update(task_id=task_id, tool=tool, task_state=state, message=message, result=result,
                        elapsed_seconds=round(time.monotonic() - started, 2))
            self.state["status"] = f"S5 配置维护：{labels[tool]}生成完成"
            self._log(f"S5 配置维护：{message} 输出：{result.get('output_dir', '')}")
        except Exception as exc:
            task.update(task_id=task_id, tool=tool, task_state="failed", message=str(exc),
                        elapsed_seconds=round(time.monotonic() - started, 2))
            self.state["status"] = f"S5 配置维护：{labels[tool]}执行失败"
            self._log(f"S5 配置维护失败：{exc}")
            self._log(traceback.format_exc(), detail=True)
        finally:
            self._log(f"S5 配置维护任务结束，耗时 {time.monotonic() - started:.1f} 秒")

    def start_s5(self, function, values=None):
        if function != "S5-F01":
            raise ValueError("S5 目前只支持 F1 汇总核查表校验")
        with self._execution_start_lock:
            if self._execution_busy():
                return {"accepted": False, "message": "已有任务正在运行，请等待完成"}
            if values:
                self.update_s5(values)
            form = copy.deepcopy(self.state["s5"])
            if not form.get("template_path"):
                raise ValueError("请先点击“识别模板”或手动选择匹配模板")
            if not form.get("period"):
                raise ValueError("尚未识别报告期，请先点击“识别模板”并核实源文件日期")
            if not form.get("business"):
                raise ValueError("尚未识别审核业务，请先点击“识别模板”并核实业务配置")
            form["files"] = form["selected_files"]
            request = S5Request.from_dict(form)
            task_id = uuid.uuid4().hex
            self._cancel_event.clear()
            self.state.update(busy=True, cancelRequested=False, status=f"{function} 正在运行")
            self.state["s5"]["task"] = {"task_id": task_id, "system": "S5", "function": function,
                                          "task_state": "queued", "files_total": len(request.files),
                                          "started_at_epoch": time.time()}
            self.state["s5"]["progress"] = {"phase": "queued", "message": "准备审核", "percent": 0}
            threading.Thread(target=self._s5_worker, args=(task_id, function, request), daemon=True).start()
            return {"accepted": True, "task_id": task_id}

    def _s5_worker(self, task_id, function, request):
        started = time.monotonic()
        form = self.state["s5"]
        form["task"]["task_state"] = "running"
        self._log(f"{function} 开始：{len(request.files)} 份文件；本期 {request.period}")

        def progress(event):
            phase = event.get("phase", "audit")
            if phase == "timing":
                current_progress = form.get("progress", {})
                form["progress"] = {
                    **current_progress,
                    "elapsedSeconds": event.get("total_elapsed_seconds", time.monotonic() - started),
                    "timingStage": event.get("stage", ""),
                    "timingSeconds": event.get("elapsed_seconds", 0),
                }
                self._log(
                    f"{function} 耗时：{event.get('message', event.get('stage', '阶段'))}"
                    f"；累计 {event.get('total_elapsed_seconds', 0):.2f} 秒"
                )
                return
            current, total = event.get("file_index", 0), event.get("files_total", 0)
            percent = (5 + 15 * max(0, current - 1) / total if phase == "identity" else
                       20 + 70 * max(0, current - 1) / total) if total else 0
            if phase == "export":
                percent = 95
            form["progress"] = {**event, "percent": percent, "elapsedSeconds": time.monotonic() - started}
            detail = f" {current}/{total}" if total else ""
            name = Path(event.get("source", "")).name
            self._log(f"{function}：{event.get('message', '')}{detail}" + (f"；{name}" if name else ""))

        result = {}
        try:
            result = run_s5_f1(request, data_dir=self.project_root / "data", on_progress=progress, cancel_event=self._cancel_event)
            self.state["status"] = f"{function} " + {"succeeded": "审核完成", "partial_failed": "审核完成，部分项目需处理", "failed": "执行失败"}[result["task_state"]]
            if result.get("summary"):
                self._log(result["summary"])
            stage_timings = result.get("stage_timings", {})
            if stage_timings:
                labels = (
                    ("configuration_load_seconds", "配置加载"),
                    ("institution_reference_load_seconds", "机构参照加载"),
                    ("file_recognition_seconds", "文件识别"),
                    ("rule_audit_seconds", "规则审核"),
                    ("formula_scan_seconds", "模板公式扫描"),
                    ("conditional_format_scan_seconds", "条件格式扫描"),
                    ("result_export_seconds", "结果导出"),
                    ("batch_seconds", "审核与结果生成合计"),
                    ("total_seconds", "总计"),
                )
                details = "；".join(
                    f"{label} {stage_timings[key]:.2f} 秒" for key, label in labels
                    if key in stage_timings
                )
                if details:
                    self._log(f"{function} 耗时明细：{details}")
            self._log(f"{function}：{self.state['status']}；结果：{result.get('result_path') or '未生成'}")
            for record in result.get("institutions", []):
                if record.get("status") not in {"COMPLETED", "COMPLETED_WITH_ISSUES"}:
                    self._log(f"{function} 文件未完成：{Path(record['source']).name}；{record.get('message') or record.get('error_code') or record.get('status')}")
            for warning in result.get("warnings", []):
                self._log(f"{function} 提示：{warning}")
        except Exception as exc:
            cancelled = isinstance(exc, AuditCancelled) or (self._cancel_event.is_set() and "用户请求终止" in str(exc))
            result = {"task_state": "cancelled" if cancelled else "failed", "message": str(exc)}
            if isinstance(exc, ConfigurationError):
                result["configuration_errors"] = [exc.as_dict()]
            self.state["status"] = f"{function} 已终止" if cancelled else f"{function} 执行失败"
            self._log(f"{self.state['status']}：{exc}")
            if not cancelled:
                self._log(traceback.format_exc(), detail=True)
        finally:
            form["task"] = {"task_id": task_id, "system": "S5", "function": function,
                            "files_total": len(request.files), **result,
                            "elapsed_seconds": round(time.monotonic() - started, 2)}
            form["progress"].update(message=self.state["status"],
                                     percent=100 if result.get("task_state") in {"succeeded", "partial_failed"} else form["progress"].get("percent", 0),
                                     elapsedSeconds=time.monotonic() - started)
            try:
                target = self._s5_settings_path.parent / "任务记录"
                target.mkdir(parents=True, exist_ok=True)
                (target / f"{task_id}.json").write_text(json.dumps(form["task"], ensure_ascii=False, indent=2), encoding="utf-8")
            except OSError as exc:
                self._log(f"S5 任务记录未保存：{exc}")
            self._log(f"{function} 任务结束，耗时 {time.monotonic() - started:.1f} 秒")
            self.state.update(busy=False, cancelRequested=False)
