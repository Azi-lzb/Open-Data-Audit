"""拓扑调度器：验证 → 拓扑序 → 逐节点执行 → 原子发布 → 失败策略 → 清理。"""

from __future__ import annotations

from dataclasses import dataclass, field
from time import perf_counter
from typing import Any, Callable, Dict, List, Optional, Tuple

from .artifacts import Artifact
from .context import NodeContext, RunContext
from .graph import WorkflowDefinition
from .registry import ModuleRegistry
from .types import (
    Cardinality,
    FailurePolicy,
    NodeRunStatus,
    Retention,
)


class WorkflowValidationError(RuntimeError):
    pass


class WorkflowExecutionError(RuntimeError):
    pass


@dataclass
class ModuleExecutionResult:
    """Handler 统一返回：状态 + 端口产物 + 事件/指标/警告/部分失败。"""

    status: NodeRunStatus = NodeRunStatus.SUCCEEDED
    outputs: Dict[str, Artifact] = field(default_factory=dict)
    events: List[str] = field(default_factory=list)
    metrics: Dict[str, Any] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    partial_failures: List[Tuple[str, str]] = field(default_factory=list)

    @classmethod
    def ok(cls, **outputs: Artifact) -> "ModuleExecutionResult":
        return cls(status=NodeRunStatus.SUCCEEDED, outputs=dict(outputs))

    @classmethod
    def partial(cls, failures, **outputs: Artifact) -> "ModuleExecutionResult":
        return cls(
            status=NodeRunStatus.PARTIAL,
            outputs=dict(outputs),
            partial_failures=list(failures),
        )

    @classmethod
    def failed(cls, reason: str = "") -> "ModuleExecutionResult":
        return cls(status=NodeRunStatus.FAILED, warnings=[reason] if reason else [])


@dataclass
class NodeRun:
    node_id: str
    status: NodeRunStatus
    reason: str = ""
    error: str = ""
    metrics: Dict[str, Any] = field(default_factory=dict)


@dataclass
class WorkflowRunResult:
    status: NodeRunStatus = NodeRunStatus.SUCCEEDED
    node_runs: List[NodeRun] = field(default_factory=list)
    outputs: Dict[str, Artifact] = field(default_factory=dict)   # 输出绑定名 → Artifact
    error: str = ""


Handler = Callable[[NodeContext, Dict[str, Any], Dict[str, Any]], ModuleExecutionResult]


def _topo_order(definition: WorkflowDefinition) -> List[str]:
    indegree = {node.node_id: 0 for node in definition.nodes}
    adjacency: Dict[str, List[str]] = {node.node_id: [] for node in definition.nodes}
    for edge in definition.edges:
        adjacency[edge.from_node].append(edge.to_node)
        indegree[edge.to_node] += 1
    queue = sorted(n for n, d in indegree.items() if d == 0)
    order: List[str] = []
    while queue:
        current = queue.pop(0)
        order.append(current)
        for nxt in sorted(adjacency[current]):
            indegree[nxt] -= 1
            if indegree[nxt] == 0:
                queue.append(nxt)
    if len(order) != len(definition.nodes):
        raise WorkflowValidationError("流程图中存在环，无法拓扑排序")
    return order


def run_workflow(
    definition: WorkflowDefinition,
    registry: ModuleRegistry,
    context: RunContext,
    inputs: Dict[str, Any],
) -> WorkflowRunResult:
    """执行一个工作流：下游不会读到尚未完成节点的输出。"""
    from .validation import validate_workflow

    errors = validate_workflow(definition, registry, context.platform_capabilities)
    if errors:
        raise WorkflowValidationError("流程校验未通过：" + "；".join(errors))

    published: Dict[Tuple[str, str], Artifact] = {}
    succeeded_nodes: set = set()
    runs: List[NodeRun] = []
    result = WorkflowRunResult()

    def publish_outputs(node, module, execution: ModuleExecutionResult) -> None:
        for port_name, artifact in execution.outputs.items():
            port = module.port(port_name, is_input=False)
            if port is None:
                raise WorkflowExecutionError(
                    f"节点“{node.label()}”声明了不存在的输出端口“{port_name}”"
                )
            if not isinstance(artifact, Artifact):
                raise WorkflowExecutionError(
                    f"节点“{node.label()}”的输出“{port_name}”不是 Artifact"
                )
            if artifact.artifact_type != port.artifact_type:
                raise WorkflowExecutionError(
                    f"节点“{node.label()}”的输出“{port_name}”类型错误："
                    f"实际 {artifact.artifact_type.value if artifact.artifact_type else '无'}，"
                    f"要求 {port.artifact_type.value}"
                )
            if port.subtype is not None and artifact.subtype != port.subtype:
                raise WorkflowExecutionError(
                    f"节点“{node.label()}”的输出“{port_name}”子类型错误："
                    f"实际 {artifact.subtype.value if artifact.subtype else '无'}，"
                    f"要求 {port.subtype.value}"
                )
            artifact.producer_node_id = node.node_id
            artifact.producer_port = port_name
            artifact.run_id = context.run_id
            policy = (node.output_policies or {}).get(port_name) or {}
            if policy.get("retention"):
                artifact.retention = Retention(str(policy["retention"]))
            elif artifact.retention == Retention.RUN and port.retention_default != Retention.RUN:
                artifact.retention = port.retention_default
            published_artifact = context.store.publish(artifact)
            published[(node.node_id, port_name)] = published_artifact
            # 普通模式的自动 CurrentFile 绑定：显式输出策略优先，其次模块默认。
            promote = bool(policy.get("promote_to_current"))
            if not policy and module.current_file_policy.get("promote_to_current"):
                promote = True
            if promote:
                context.promote_current_file(
                    published_artifact, str(policy.get("slot") or "current")
                )

    order = _topo_order(definition)
    for node_id in order:
        node = definition.node(node_id)
        cancel_event = context.resources.get("cancel_event")
        if cancel_event is not None and cancel_event.is_set():
            runs.append(NodeRun(node_id, NodeRunStatus.SKIPPED, reason="用户请求终止"))
            for later in order[order.index(node_id) + 1:]:
                runs.append(NodeRun(later, NodeRunStatus.SKIPPED, reason="用户请求终止"))
            result.status = NodeRunStatus.FAILED
            result.node_runs = runs
            result.error = "用户请求终止：已停止后续节点，已生成文件已保留"
            context.log("已收到终止请求，停止后续节点")
            context.store.cleanup_temporary()
            return result
        module = registry.get(node.module_id, node.module_version)
        if not node.enabled:
            runs.append(NodeRun(node_id, NodeRunStatus.DISABLED, reason="节点已禁用"))
            continue

        # ---- 解析节点输入 ----
        resolved: Dict[str, Any] = {}
        missing_required: List[str] = []
        edge_counts: Dict[str, int] = {}
        for edge in definition.incoming(node_id):
            edge_counts[edge.to_port] = edge_counts.get(edge.to_port, 0) + 1
            artifact = published.get((edge.from_node, edge.from_port))
            if artifact is None:
                continue
            port = module.port(edge.to_port, is_input=True)
            if port is not None and port.cardinality == Cardinality.MANY:
                resolved.setdefault(edge.to_port, []).append(artifact)
            else:
                resolved[edge.to_port] = artifact
        for port in module.input_ports:
            if port.required and port.name not in resolved:
                missing_required.append(port.name)
        # ---- 流程级输入注入（优先于内部解析的同名端口） ----
        for binding in definition.input_bindings:
            if binding.node_id == node_id and binding.name in inputs:
                resolved[binding.port] = inputs[binding.name]

        missing_required: List[str] = []
        for port in module.input_ports:
            if port.required and port.name not in resolved:
                missing_required.append(port.name)
        if missing_required:
            runs.append(NodeRun(
                node_id, NodeRunStatus.SKIPPED,
                reason="上游未产出必填输入：" + "、".join(missing_required),
            ))
            continue

        node_context = NodeContext(run=context, node=node)
        handler = registry.handler(node.module_id)
        node_context.log("开始执行")
        started = perf_counter()
        try:
            execution = handler(node_context, dict(node.parameters), resolved)
        except Exception as exc:  # Handler 未自行处理失败 → 按策略收敛
            execution = ModuleExecutionResult.failed(f"{type(exc).__name__}: {exc}")
            node_context.log(f"节点异常：{exc}")
        elapsed_seconds = perf_counter() - started
        # 只记录观测指标，不影响节点输出、调度或失败策略。
        execution.metrics = {
            **execution.metrics,
            "elapsed_seconds": elapsed_seconds,
        }

        if execution.status == NodeRunStatus.FAILED:
            if node.failure_policy == FailurePolicy.FAIL_WORKFLOW:
                runs.append(NodeRun(node_id, NodeRunStatus.FAILED, error="；".join(execution.warnings)))
                result.status = NodeRunStatus.FAILED
                # 剩余节点补记 SKIPPED，运行记录保持完整可追溯
                for later in order[order.index(node_id) + 1:]:
                    runs.append(NodeRun(later, NodeRunStatus.SKIPPED, reason="上游失败，工作流已中止"))
                result.node_runs = runs
                result.error = (
                    f"节点“{node.label()}”执行失败（失败策略={node.failure_policy.value}）："
                    + "；".join(execution.warnings)
                )
                context.store.cleanup_temporary()
                return result
            runs.append(NodeRun(node_id, NodeRunStatus.SKIPPED, reason="失败后按策略跳过该节点"))
            continue

        publish_outputs(node, module, execution)
        status = execution.status
        node_context.log(f"完成，耗时 {elapsed_seconds:.2f} 秒")
        runs.append(NodeRun(
            node_id, status,
            error="；".join(execution.partial_failures and [
                f"{item}: {reason}" for item, reason in execution.partial_failures
            ] or []),
            metrics=dict(execution.metrics),
        ))
        succeeded_nodes.add(node_id)

    # ---- 收集流程输出 ----
    for binding in definition.output_bindings:
        artifact = published.get((binding.node_id, binding.port))
        if artifact is not None:
            result.outputs[binding.name] = artifact

    if result.status != NodeRunStatus.FAILED:
        failed_runs = [run for run in runs if run.status == NodeRunStatus.FAILED]
        partial_runs = [run for run in runs if run.status == NodeRunStatus.PARTIAL]
        if failed_runs:
            result.status = NodeRunStatus.FAILED
        elif partial_runs:
            result.status = NodeRunStatus.PARTIAL
        else:
            result.status = NodeRunStatus.SUCCEEDED
    result.node_runs = runs
    context.store.cleanup_temporary()
    return result
