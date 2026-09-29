"""DAG 流程运行器：把四条默认 DAG 接到现有服务与引擎上。

本模块是基础数据审核的唯一流程运行通道。
"""

from __future__ import annotations

from dataclasses import replace
from datetime import datetime
from pathlib import Path
import tempfile
import threading
import time
from typing import Callable, Dict, Optional

from .context import RunContext
from .artifact_store import ArtifactStore
from .artifacts import MetadataArtifact
from .defaults import default_workflows
from .registry import ModuleRegistry
from .scheduler import run_workflow

from ..engines import pipeline_kind
from ..feature_log import FeatureLog
from ..service import AuditService
from ..service import _output_prefix


class DagFlowError(RuntimeError):
    pass


def _describe_node_problems(node_runs) -> str:
    """汇总失败/被跳过节点的原因，用于无 issues 产出时的可诊断报错。"""
    parts = []
    for run in node_runs or ():
        status = getattr(run.status, "value", str(run.status))
        if status in ("FAILED", "SKIPPED"):
            reason = getattr(run, "error", None) or getattr(run, "reason", None)
            if reason:
                node_id = getattr(run, "node_id", "?")
                parts.append(f"节点「{node_id}」{status}：{reason}")
    return "；".join(parts) or "未记录节点失败原因"


def run_dag_native_audit(
    *, service: AuditService, template_path: Path, input_dir: Path, output_dir: Path,
    period: str, history_path: Path, selected_files=None, external_path: Path | None = None,
    extra_files=None, recursive: bool = False, on_step=None, write_flow_logs: bool = True,
    definition=None, cancel_event=None,
):
    """UOS/native 审核 DAG。

    此入口只构造运行上下文并调度 DAG 节点，不读取旧流程运行时资源。
    """
    from ..history import merge_history
    from ..models import AuditRunResult
    from ..native.history_io import read_history_xlsx, write_history_xlsx

    log = on_step or (lambda _text: None)
    batch_id = datetime.now().strftime("%Y%m%d%H%M%S")
    definition = definition or default_workflows()["dag:汇总核查表校验"]
    if definition.workflow_id == "dag:汇总核查表校验":
        # S1F1 公式计算强依赖本机计算引擎；先实测可启动再开跑，避免跑到
        # “公式校验复制”才发现 LibreOffice 缺失或与系统 glibc 不匹配。
        from ..engines import calculation_engine_preflight
        from ..system_info import system_environment_text
        log(f"[运行环境] {system_environment_text()}")
        ok, message = calculation_engine_preflight()
        log(f"[计算引擎预检] {message}")
        if not ok:
            raise DagFlowError(f"汇总核查表校验未开始：{message}")
    flow_name = (
        "汇总核查表校验" if definition.workflow_id == "dag:汇总核查表校验"
        else str(definition.settings.get("flow_name") or definition.name)
    )
    feature_log = FeatureLog(flow_name, output_dir) if write_flow_logs else None
    workspace = tempfile.TemporaryDirectory(prefix="base-audit-native-dag-")
    runtime = _runtime_parameters({
        "input_dir": input_dir, "output_dir": output_dir, "template_path": template_path,
        "external_path": external_path, "selected_files": selected_files, "extra_files": extra_files,
        "recursive": recursive, "period": period, "batch_id": batch_id,
        "audit_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "flow_name": flow_name,
    })
    runtime.update({
        "stage_dir": workspace.name,
        "audit_output_dir": str(Path(output_dir) / f"50_机构审核副本_{batch_id}"),
    })
    # default_workflows() 是内存定义，但仍复制节点，防止本次运行污染下一次参数。
    definition = replace(definition, nodes=[replace(node, parameters={**runtime, **node.parameters}) for node in definition.nodes])
    context = RunContext(workflow_id=definition.workflow_id, store=ArtifactStore(),
                         platform_capabilities=("native",), log=log)
    context.resources["feature_log"] = feature_log
    context.resources["cancel_event"] = cancel_event or threading.Event()
    old_history = read_history_xlsx(history_path)
    try:
        result = run_workflow(definition, build_registry("native"), context,
                              {"history": MetadataArtifact(payload={"history": old_history})})
        if result.status.value == "FAILED":
            raise RuntimeError(result.error or "DAG 流程执行失败")
        issues_artifact = result.outputs.get("issues")
        if issues_artifact is None:
            # 节点按失败策略跳过后流程可能不置 FAILED，但下游必备产物缺失；
            # 此时把节点失败原因直接带给用户，而不是抛裸 KeyError('issues')。
            detail = result.error or _describe_node_problems(result.node_runs)
            raise DagFlowError(f"流程未产出审核结果（缺少 issues 产物）：{detail}")
        issues = list(issues_artifact.issues)
        current = issues
        resolved = list(issues_artifact.metadata.get("resolved_issues") or ())
        # DAG 的历史富化节点已在导航/结果输出前完成分类；此处只持久化历史资产。
        write_history_xlsx(history_path, merge_history(old_history, current))
        final_workbooks = result.outputs.get("final_workbooks")
        files = _file_results_from_workbooks(final_workbooks, current)
        structure = result.outputs.get("structure_report")
        if structure is not None:
            files = list(structure.metadata.get("rejected_file_results") or ()) + files
        summary = result.outputs.get("summary_workbook")
        copies = {"公式校验复制": Path(final_workbooks.records[0].audit).parent} if final_workbooks and final_workbooks.records else {}
        response = AuditRunResult(batch_id=batch_id, period=period, output_dir=Path(output_dir),
            current_issues=current, resolved_issues=resolved, files=files,
            summary_path=Path(summary.path) if summary else None, history_path=history_path,
            performance_lines=("计算引擎 LibreOffice Calc",), copies=copies)
        if feature_log is not None:
            path = feature_log.write()
            if path is not None:
                response = replace(response, log_path=path)
        return response
    except Exception:
        if feature_log is not None:
            try: feature_log.write()
            except Exception: pass
        raise
    finally:
        workspace.cleanup()


def build_registry(platform_kind: str = "com") -> ModuleRegistry:
    from .adapters.audit_handlers import AUDIT_HANDLERS, AUDIT_MODULE_DEFINITIONS
    from .adapters.simple_handlers import SUMMARY_HANDLERS, SUMMARY_MODULE_DEFINITIONS
    if platform_kind == "native":
        from .adapters.native_audit_handlers import AUDIT_NATIVE_HANDLERS
        audit_handlers = AUDIT_NATIVE_HANDLERS
    else:
        audit_handlers = AUDIT_HANDLERS

    registry = ModuleRegistry()
    for definition in AUDIT_MODULE_DEFINITIONS:
        registry.register(definition, audit_handlers[definition.module_id])
    for definition in SUMMARY_MODULE_DEFINITIONS:
        registry.register(definition, SUMMARY_HANDLERS[definition.module_id])
    return registry


def _runtime_parameters(params: Dict) -> Dict:
    return {
        "input_dir": str(params["input_dir"]),
        "output_dir": str(params["output_dir"]),
        "template_path": str(params.get("template_path") or ""),
        "external_path": str(params.get("external_path") or ""),
        "selected_files": [str(p) for p in params["selected_files"]] if params.get("selected_files") is not None else None,
        "extra_files": [str(p) for p in params["extra_files"]] if params.get("extra_files") is not None else None,
        "recursive": params.get("recursive", True),
        "period": params.get("period") or "",
        "batch_id": params.get("batch_id") or "",
        "audit_time": params.get("audit_time") or "",
        "flow_name": params.get("flow_name") or "",
        "include_external": bool(params.get("external_path")),
    }


def _file_results_from_workbooks(workbooks, issues):
    """无审核导航节点时，直接由最终公式副本 Artifact 组装文件结果。"""
    from ..models import FileAuditResult

    if workbooks is None:
        return []
    by_source = {}
    for issue in issues:
        by_source.setdefault(str(Path(issue.source_file).resolve()), []).append(issue)
    return [
        FileAuditResult(
            source_path=Path(record.source), audit_path=Path(record.audit) if record.audit else None,
            org_code=record.org_code, org_name=record.org_name,
            issues=list(by_source.get(str(Path(record.source).resolve()), ())),
            error="；".join(record.stage_errors),
        )
        for record in workbooks.records
    ]


def run_dag_audit(
    *,
    service: AuditService,
    template_path: Path,
    input_dir: Path,
    output_dir: Path,
    period: str,
    history_path: Path,
    selected_files=None,
    external_path: Path | None = None,
    extra_files=None,
    recursive: bool = False,
    on_step: Optional[Callable[[str], None]] = None,
    write_flow_logs: bool = True,
    engine_preference: str | None = None,
    definition=None, cancel_event=None,
):
    """汇总核查表校验的 DAG 并行实现（Windows COM 管线）。"""
    from ..excel_com import ExcelSession

    started = time.monotonic()
    engine_preference = engine_preference or service.engine_preference
    if pipeline_kind(engine_preference) != "com":
        return run_dag_native_audit(
            service=service, template_path=template_path, input_dir=input_dir,
            output_dir=output_dir, period=period, history_path=history_path,
            selected_files=selected_files, external_path=external_path,
            extra_files=extra_files, recursive=recursive, on_step=on_step,
            write_flow_logs=write_flow_logs,
            definition=definition, cancel_event=cancel_event,
        )

    log = on_step or (lambda _text: None)
    batch_id = datetime.now().strftime("%Y%m%d%H%M%S")
    audit_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    history_path, history_sheet, legacy_history_path = service._history_storage(history_path)
    definition = definition or default_workflows()["dag:汇总核查表校验"]
    if definition.workflow_id == "dag:汇总核查表校验":
        # 与 native 入口同口径：先实测 Excel/WPS COM 可用再开跑。
        from ..engines import calculation_engine_preflight
        from ..system_info import system_environment_text
        log(f"[运行环境] {system_environment_text()}")
        ok, message = calculation_engine_preflight(engine_preference)
        log(f"[计算引擎预检] {message}")
        if not ok:
            raise DagFlowError(f"汇总核查表校验未开始：{message}")
    flow_name = (
        "汇总核查表校验" if definition.workflow_id == "dag:汇总核查表校验"
        else str(definition.settings.get("flow_name") or definition.name)
    )
    feature_log = FeatureLog(flow_name, output_dir) if write_flow_logs else None
    stage_workspace = tempfile.TemporaryDirectory(prefix="base-audit-dag-")
    runtime = _runtime_parameters({
        "input_dir": input_dir, "output_dir": output_dir, "template_path": template_path,
        "external_path": external_path, "selected_files": selected_files,
        "extra_files": extra_files, "recursive": recursive, "period": period,
        "batch_id": batch_id, "audit_time": audit_time, "flow_name": flow_name,
    })
    runtime["stage_dir"] = stage_workspace.name
    runtime["audit_output_dir"] = str(output_dir / (
        f"{_output_prefix(50, '公式校验复制', '机构审核副本')}_{batch_id}"
    ))
    definition = replace(
        definition,
        nodes=[
            replace(node, parameters={**runtime, **node.parameters})
            for node in definition.nodes
        ],
    )

    registry = build_registry()
    from .artifact_store import ArtifactStore

    context = RunContext(
        workflow_id=definition.workflow_id, store=ArtifactStore(),
        platform_capabilities=("com",), log=log,
    )
    context.resources["feature_log"] = feature_log
    context.resources["cancel_event"] = cancel_event or threading.Event()

    context.services["service"] = service
    startup_complete = threading.Event()
    history_load_seconds = 0.0

    def _report_slow_engine_startup() -> None:
        if not startup_complete.is_set():
            log(
                "表格引擎启动超过 15 秒：可能有 Excel/WPS 隐藏对话框或遗留自动化进程；"
                "请先关闭无响应的表格软件后重试。"
            )

    log(f"正在启动表格引擎：{engine_preference}……")
    startup_watchdog = threading.Timer(15, _report_slow_engine_startup)
    startup_watchdog.daemon = True
    startup_watchdog.start()
    try:
        with ExcelSession(engine_preference) as excel:
            startup_complete.set()
            context.services["excel"] = excel
            log(f"已连接表格引擎：{excel.engine_name}")
            # 历史表仅为值读取，不需要 COM；ExcelSession 只服务后续的公式重算
            # 与条件格式实际渲染两个节点。
            from ..native.history_io import read_history_xlsx

            history_started = time.monotonic()
            history = read_history_xlsx(history_path, sheet_name=history_sheet)
            if not history and legacy_history_path and legacy_history_path.exists():
                history = read_history_xlsx(legacy_history_path)
            history_load_seconds = time.monotonic() - history_started
            run_result = run_workflow(
                definition, registry, context,
                {"history": MetadataArtifact(payload={"history": history})},
            )
    except Exception:
        # 流程中途失败时也写一份已收集 sheet 的运行日志，便于排错（与旧入口一致）。
        if feature_log is not None:
            try:
                feature_log.write()
            except Exception:
                pass
        stage_workspace.cleanup()
        raise
    finally:
        startup_complete.set()
        startup_watchdog.cancel()

    if run_result.status.value == "FAILED":
        if feature_log is not None:
            try:
                feature_log.write()
            except Exception:
                pass
        stage_workspace.cleanup()
        raise RuntimeError(run_result.error or "DAG 流程执行失败")

    # ---- 组装与旧版同构的 AuditRunResult ----
    from ..models import AuditRunResult

    node_seconds = sum(
        float(run.metrics.get("elapsed_seconds") or 0.0)
        for run in run_result.node_runs
    )
    node_labels = {node.node_id: node.display_name or node.node_id for node in definition.nodes}
    final_workbooks = run_result.outputs.get("final_workbooks")
    copies = {}
    if final_workbooks and final_workbooks.records:
        copies["公式校验复制"] = Path(final_workbooks.records[0].audit).parent
    summary_path = None
    artifact = run_result.outputs.get("summary_workbook")
    if artifact is not None:
        summary_path = Path(artifact.path)
    issue_artifact = run_result.outputs.get("issues")
    structure_artifact = run_result.outputs.get("structure_report")
    files = _file_results_from_workbooks(
        final_workbooks, list(issue_artifact.issues) if issue_artifact is not None else [],
    )
    if structure_artifact is not None:
        files = list(structure_artifact.metadata.get("rejected_file_results") or ()) + files
    result = AuditRunResult(
        batch_id=batch_id,
        period=str(period).strip(),
        output_dir=Path(output_dir),
        current_issues=list(issue_artifact.issues) if issue_artifact is not None else [],
        resolved_issues=[],
        files=files,
        summary_path=summary_path,
        history_path=None,
        preflight_path=None,
        performance_lines=(),
        copies=copies,
    )
    feature_log_write_seconds = 0.0
    if feature_log is not None:
        feature_log_started = time.monotonic()
        log_path = feature_log.write()
        feature_log_write_seconds = time.monotonic() - feature_log_started
        if log_path is not None:
            result = replace(result, log_path=log_path)
    workspace_cleanup_started = time.monotonic()
    stage_workspace.cleanup()
    workspace_cleanup_seconds = time.monotonic() - workspace_cleanup_started
    # 运行日志落盘也属于一次完整执行；在最后重新计算总计，避免“总计”
    # 只覆盖 DAG 节点、不覆盖引擎启动和收尾。
    final_total_seconds = time.monotonic() - started
    final_performance = [f"计算引擎 {excel.engine_name or '未知'}"]
    for run in run_result.node_runs:
        if "elapsed_seconds" in run.metrics:
            final_performance.append(
                f"{node_labels.get(run.node_id, run.node_id)} {float(run.metrics['elapsed_seconds']):.2f} 秒"
            )
        if "formula_copy_seconds" in run.metrics:
            final_performance.append(
                    "  └ 公式写入 {:.2f} 秒；逐副本重算与保存 {:.2f} 秒".format(
                    float(run.metrics["formula_copy_seconds"]),
                    float(run.metrics.get("formula_recalculate_seconds") or 0.0),
                )
            )
        if "external_source_load_time" in run.metrics:
            final_performance.append(
                "  └ 外部源装载 {:.2f} 秒；表快照提取 {:.2f} 秒；写入 {:.2f} 秒；"
                "样式 {:.2f} 秒；保存 {:.2f} 秒；副本合计 {:.2f} 秒".format(
                    float(run.metrics["external_source_load_time"]),
                    float(run.metrics["sheet_extract_time"]),
                    float(run.metrics["sheet_write_time"]),
                    float(run.metrics["style_write_time"]),
                    float(run.metrics["xlsx_save_time"]),
                    float(run.metrics["per_workbook_total"]),
                )
            )
        if "external_direct_write_seconds" in run.metrics:
            final_performance.append(
                "  └ Direct OOXML 外部表准备 {:.2f} 秒；写入 {:.2f} 秒".format(
                    float(run.metrics.get("external_prepare_seconds") or 0.0),
                    float(run.metrics.get("external_direct_write_seconds") or 0.0),
                )
            )
    lifecycle = getattr(excel, "lifecycle_metrics", {})
    if lifecycle:
        final_performance.append(
            "引擎生命周期：" + "；".join(
                f"{name} {float(lifecycle.get(name) or 0.0):.2f} 秒"
                for name in (
                    "engine_detect_time", "excel_process_start_time", "com_connect_time",
                    "excel_ready_time", "template_com_open_time", "idle_wait_time",
                    "workbook_cleanup_time", "excel_quit_time", "com_release_time",
                    "final_gc_time",
                )
            )
        )
    final_performance.append(
        "非节点收尾：历史读取 {:.2f} 秒；运行日志写入 {:.2f} 秒；临时文件清理 {:.2f} 秒".format(
            history_load_seconds, feature_log_write_seconds, workspace_cleanup_seconds,
        )
    )
    final_performance.extend((
        f"引擎启动与流程收尾 {max(0.0, final_total_seconds - node_seconds):.2f} 秒",
        f"总计 {final_total_seconds:.2f} 秒",
    ))
    result = replace(result, performance_lines=tuple(final_performance))
    return result


def run_dag_simple(
    workflow_id: str,
    *,
    service: AuditService,
    template_path: Path | None,
    input_dir: Path,
    output_dir: Path,
    period: str,
    history_path: Path,
    selected_files=None,
    external_path: Path | None = None,
    recursive: bool = True,
    on_step: Optional[Callable[[str], None]] = None,
    write_flow_logs: bool = True,
    engine_preference: str | None = None,
    definition=None, cancel_event=None,
):
    """汇总/组合流程的 DAG 运行：节点内部直接调用 service 公共方法。

    汇总流程含 COM 专属的命名区域检查/结构比对/公式副本节点（复用审核链的
    ExcelSession 资源），因此需要按平台选择注册表并准备对应资源；组合流程
    只走平台无关的 openpyxl 节点。
    """
    log = on_step or (lambda _text: None)
    definition = definition or default_workflows()[workflow_id]
    execution_workflow_id = str(
        definition.settings.get("base_workflow_id") or workflow_id
    )
    flow_name = str(
        definition.settings.get("flow_name")
        or definition.name
        or workflow_id.split(":", 1)[-1]
    )
    feature_log = FeatureLog(flow_name, output_dir) if write_flow_logs else None
    params = {
        "input_dir": str(input_dir), "output_dir": str(output_dir),
        "template_path": str(template_path or ""), "external_path": str(external_path or ""),
        "selected_files": [str(p) for p in selected_files] if selected_files is not None else None,
        "recursive": recursive, "period": period,
        "flow_name": flow_name,
        "summary_feature_names": ["任意行汇总"],
    }
    is_summary = execution_workflow_id == "dag:汇总校验结果说明"
    # 汇总说明的读取方式独立于“汇总核查表校验”的公式计算引擎。默认走
    # 无 COM 的 openpyxl 节点，只有用户明确选 Excel/WPS 时才启动 COM。
    kind = (
        service.summary_pipeline_kind()
        if is_summary
        else pipeline_kind(engine_preference or service.engine_preference)
    )
    registry = build_registry("native" if kind == "native" else "com")
    context = RunContext(
        workflow_id=definition.workflow_id, store=ArtifactStore(),
        platform_capabilities=(kind,), log=log,
    )
    context.services["service"] = service
    context.resources["feature_log"] = feature_log
    context.resources["cancel_event"] = cancel_event or threading.Event()

    if is_summary and kind == "com":
        # 表结构比对等节点复用审核链的 ExcelSession：模板只读打开一次，
        # 节点通过 ctx.resource 取用。runtime 需带 batch_id/stage_dir 等。
        from ..excel_com import ExcelSession

        batch_id = datetime.now().strftime("%Y%m%d%H%M%S")
        with tempfile.TemporaryDirectory(prefix="base-audit-dag-summary-") as workspace:
            params.update({
                "batch_id": batch_id, "audit_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "stage_dir": workspace,
                "audit_output_dir": str(Path(output_dir) / f"{_output_prefix(60, '校验结果提取', '本期审核结果')}_{batch_id}"),
            })
            definition = replace(
                definition,
                nodes=[replace(node, parameters={**params, **node.parameters}) for node in definition.nodes],
            )
            try:
                with ExcelSession(engine_preference or service.engine_preference) as excel:
                    context.services["excel"] = excel
                    log(f"已连接表格引擎：{excel.engine_name}")
                    template_workbook = None
                    if template_path and Path(template_path).is_file():
                        template_workbook = excel.open_workbook(
                            Path(template_path), read_only=True, purpose="template",
                        )
                        context.services["template_workbook"] = template_workbook
                    try:
                        result = run_workflow(
                            definition, registry, context,
                            {"runtime": MetadataArtifact(payload=params)},
                        )
                    finally:
                        if template_workbook is not None:
                            # 汇总节点会在自己的 COM 会话里重新打开并关闭模板
                            # （同一 Excel 实例共享工作簿）；此处只是收尾，重复
                            # 关闭失效对象不应影响已完成的流程。ExcelSession
                            # 退出时也会统一关闭所有工作簿。
                            try:
                                excel.close_workbook(template_workbook)
                            except Exception:
                                pass
            except Exception:
                if feature_log is not None:
                    try: feature_log.write()
                    except Exception: pass
                raise
    else:
        if is_summary and kind == "native":
            params.update({
                "audit_output_dir": str(output_dir),
            })
            definition = replace(
                definition,
                nodes=[replace(node, parameters={**params, **node.parameters}) for node in definition.nodes],
            )
        try:
            result = run_workflow(
                definition, registry, context,
                {"runtime": MetadataArtifact(payload=params)},
            )
        except Exception:
            if feature_log is not None:
                try: feature_log.write()
                except Exception: pass
            raise

    if result.status.value == "FAILED":
        if feature_log is not None:
            try: feature_log.write()
            except Exception: pass
        raise RuntimeError(result.error or "DAG 流程执行失败")
    # 汇总流程的公开输出名为 summary_result；两个组合流程为 result。
    output_name = "summary_result" if execution_workflow_id == "dag:汇总校验结果说明" else "result"
    try:
        inner = result.outputs[output_name].payload["result"]
    except KeyError as exc:
        raise RuntimeError(f"DAG 流程“{flow_name}”未产出预期结果“{output_name}”") from exc
    if feature_log is not None:
        log_path = feature_log.write()
        if log_path is not None:
            inner = replace(inner, log_path=log_path)
    return inner


def run_dag_flow(
    workflow_id: str,
    *,
    service: AuditService,
    template_path: Path | None,
    input_dir: Path,
    output_dir: Path,
    period: str,
    history_path: Path,
    selected_files=None,
    external_path: Path | None = None,
    extra_files=None,
    recursive: bool = False,
    on_step: Optional[Callable[[str], None]] = None,
    write_flow_logs: bool = True,
    engine_preference: str | None = None,
    strict: bool = True,
    cancel_event=None,
):
    """按 workflow_id 分派的统一 DAG 入口。"""
    # DAG 拓扑、命名区域和失败策略均为已验收的代码默认值，避免将可执行
    # 数据流暴露为可编辑配置而出现“界面保存成功、执行语义已改变”的隐患。
    configured_defaults = default_workflows()
    if workflow_id == "dag:汇总核查表校验":
        return run_dag_audit(
            service=service, template_path=template_path, input_dir=input_dir,
            output_dir=output_dir, period=period, history_path=history_path,
            selected_files=selected_files, external_path=external_path,
            extra_files=extra_files, recursive=recursive, on_step=on_step,
            write_flow_logs=write_flow_logs, engine_preference=engine_preference,
            definition=configured_defaults[workflow_id], cancel_event=cancel_event,
        )
    if workflow_id in {"dag:汇总校验结果说明", "dag:组合联合核查表", "dag:组合工作表"}:
        return run_dag_simple(
            workflow_id, service=service, template_path=template_path,
            input_dir=input_dir, output_dir=output_dir, period=period,
            history_path=history_path, selected_files=selected_files,
            external_path=external_path, recursive=recursive, on_step=on_step,
            write_flow_logs=write_flow_logs, engine_preference=engine_preference,
            definition=configured_defaults[workflow_id], cancel_event=cancel_event,
        )
    if workflow_id.startswith("custom:"):
        raise DagFlowError("当前版本不支持自定义 DAG；请使用内置流程")
    raise DagFlowError(f"未知的 DAG 流程：{workflow_id}")
