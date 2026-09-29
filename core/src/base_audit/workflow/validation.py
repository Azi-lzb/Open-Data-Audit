"""DAG 校验器：连接合法性、类型兼容、环、能力与冲突检查（集中实现）。"""

from __future__ import annotations

from typing import Dict, List, Set, Tuple

from .graph import WorkflowDefinition
from .modules import ModuleCategory
from .registry import ModuleRegistry
from .types import Cardinality, ModuleCategory


def _topo_order(nodes, edges) -> List[str] | None:
    """Kahn 拓扑排序；存在环时返回 None。"""
    indegree = {node: 0 for node in nodes}
    adjacency: Dict[str, List[str]] = {node: [] for node in nodes}
    for edge in edges:
        adjacency[edge.from_node].append(edge.to_node)
        indegree[edge.to_node] += 1
    queue = sorted([n for n, d in indegree.items() if d == 0])
    order: List[str] = []
    while queue:
        current = queue.pop(0)
        order.append(current)
        for nxt in sorted(adjacency[current]):
            indegree[nxt] -= 1
            if indegree[nxt] == 0:
                queue.append(nxt)
    if len(order) != len(nodes):
        return None
    return order


def _reachable_from_inputs(definition: WorkflowDefinition) -> Set[str]:
    """从流程输入绑定或入度为零的根节点出发沿边可达的节点集合。"""
    roots = {binding.node_id for binding in definition.input_bindings}
    indegree: Dict[str, int] = {}
    for edge in definition.edges:
        indegree[edge.to_node] = indegree.get(edge.to_node, 0) + 1
    for node in definition.nodes:
        if indegree.get(node.node_id, 0) == 0:
            roots.add(node.node_id)
    reachable: Set[str] = set()
    stack = sorted(roots)
    while stack:
        current = stack.pop()
        if current in reachable:
            continue
        reachable.add(current)
        for edge in definition.outgoing(current):
            if edge.to_node not in reachable:
                stack.append(edge.to_node)
    return reachable


def validate_workflow(
    definition: WorkflowDefinition,
    registry: ModuleRegistry,
    platform_capabilities: Tuple[str, ...] = (),
) -> List[str]:
    """返回全部校验错误；空列表表示通过。禁止静默放过任何一项。"""
    errors: List[str] = []
    node_ids = [node.node_id for node in definition.nodes]
    duplicates = {nid for nid in node_ids if node_ids.count(nid) > 1}
    if duplicates:
        errors.append("节点编号重复：" + "、".join(sorted(duplicates)))
        return errors

    module_of = {}
    for node in definition.nodes:
        try:
            module_of[node.node_id] = registry.get(node.module_id, node.module_version)
        except KeyError as exc:
            errors.append(f"节点“{node.label()}”的模块不存在：{exc}")
    if errors:
        return errors

    # ---- 边引用与类型兼容 ----
    for edge in definition.edges:
        source = definition.node(edge.from_node)
        target = definition.node(edge.to_node)
        if source is None or target is None:
            errors.append(
                f"连线引用了不存在的节点：{edge.from_node}.{edge.from_port} → {edge.to_node}.{edge.to_port}"
            )
            continue
        source_module = module_of[source.node_id]
        target_module = module_of[target.node_id]
        out_port = source_module.port(edge.from_port, is_input=False)
        in_port = target_module.port(edge.to_port, is_input=True)
        if out_port is None:
            errors.append(f"节点“{source.label()}”没有输出端口“{edge.from_port}”")
            continue
        if in_port is None:
            errors.append(f"节点“{target.label()}”没有输入端口“{edge.to_port}”")
            continue
        from .ports import is_compatible

        if not is_compatible(out_port, in_port):
            errors.append(
                f"连线类型不兼容：{source.label()}.{edge.from_port}（{out_port.artifact_type.value}"
                f"/{out_port.subtype.value if out_port.subtype else '无'}）→ "
                f"{target.label()}.{edge.to_port}（要求 {in_port.artifact_type.value}）"
            )

    # ---- 必填端口、ONE 单连接、MANY 兼容、禁用断裂 ----
    input_bindings = {
        (binding.node_id, binding.port) for binding in definition.input_bindings
    }
    inputs_by_node: Dict[str, List[Tuple[str, int]]] = {}
    for edge in definition.edges:
        inputs_by_node.setdefault(edge.to_node, []).append((edge.to_port, 1))
    for node in definition.nodes:
        module = module_of[node.node_id]
        if not node.enabled:
            continue  # 禁用节点不执行：不检查其输入连接完整性
        connected = {}
        for port_name, _one in inputs_by_node.get(node.node_id, []):
            connected[port_name] = connected.get(port_name, 0) + 1
        for port in module.input_ports:
            count = connected.get(port.name, 0)
            bound = (node.node_id, port.name) in input_bindings
            if port.required and count == 0 and not bound:
                errors.append(f"节点“{node.label()}”的必填输入“{port.name}”未连接")
            if port.cardinality == Cardinality.ONE and count > 1:
                errors.append(
                    f"节点“{node.label()}”的单值输入“{port.name}”接了 {count} 条连线（只允许一条）"
                )
            if count > 0 and not node.enabled:
                continue
            if count == 0 and not port.required and node.enabled:
                continue
        # 禁用节点导致必填连接断裂：上游全部禁用时，必填输入视为断裂。
        if node.enabled:
            for port in module.input_ports:
                if not port.required:
                    continue
                feeders = [e.from_node for e in definition.incoming(node.node_id) if e.to_port == port.name]
                if feeders and all(not (definition.node(f) or node).enabled for f in feeders):
                    errors.append(
                        f"节点“{node.label()}”的必填输入“{port.name}”的上游节点全部被禁用"
                    )

    # ---- 环 ----
    order = _topo_order([n.node_id for n in definition.nodes], definition.edges)
    if order is None:
        errors.append("流程图中存在环，无法拓扑排序")

    # ---- 流程输出必须可从流程输入到达（经过启用节点） ----
    reachable = _reachable_from_inputs(definition)
    for binding in definition.output_bindings:
        node = definition.node(binding.node_id)
        if node is None:
            errors.append(f"流程输出“{binding.name}”绑定了不存在的节点“{binding.node_id}”")
            continue
        if binding.node_id not in reachable:
            errors.append(
                f"流程输出“{binding.name}”（{node.label()}）无法从流程输入到达"
            )

    # ---- 平台能力 ----
    for node in definition.nodes:
        module = module_of[node.node_id]
        if module.platform_capabilities and platform_capabilities and not any(
            cap in module.platform_capabilities for cap in platform_capabilities
        ):
            errors.append(
                f"节点“{node.label()}”需要平台能力 {'/'.join(module.platform_capabilities)}，"
                f"当前平台不满足（{'/'.join(platform_capabilities) or '无'}）"
            )

    # ---- 隐式 CurrentFile 冲突：多个启用 MODIFY 同时升位同一“当前文件”槽位 ----
    promote_slots: Dict[str, List[str]] = {}
    for node in definition.nodes:
        module = module_of[node.node_id]
        if module.category != ModuleCategory.MODIFY or not node.enabled:
            continue
        for port_name, policy in (node.output_policies or {}).items():
            if policy.get("promote_to_current"):
                slot = str(policy.get("slot") or "current")
                promote_slots.setdefault(slot, []).append(node.label())
    for slot, labels in promote_slots.items():
        if len(labels) > 1:
            errors.append(
                f"存在多个修改节点把输出生效为同一“当前文件”（{slot}）且没有显式合并节点："
                + "、".join(labels)
            )

    # ---- 输出文件名冲突 ----
    output_names: Dict[str, List[str]] = {}
    for node in definition.nodes:
        if not node.enabled:
            continue
        name = str((node.parameters or {}).get("output_name") or "").strip()
        if name:
            output_names.setdefault(name, []).append(node.label())
    for name, labels in output_names.items():
        if len(labels) > 1:
            errors.append(f"输出文件名冲突：“{name}”被多个节点使用（{'、'.join(labels)}）")

    # ---- SINK 输入完整 / 结果汇聚至少一个输入 ----
    for node in definition.nodes:
        module = module_of[node.node_id]
        if module.category != ModuleCategory.SINK or not node.enabled:
            continue
        for port in module.input_ports:
            if not port.required:
                continue
            if not any(e.to_port == port.name for e in definition.incoming(node.node_id)):
                errors.append(f"出口节点“{node.label()}”的必填输入“{port.name}”未连接")
        if not definition.incoming(node.node_id):
            errors.append(f"出口节点“{node.label()}”没有任何输入")
    for node in definition.nodes:
        module = module_of[node.node_id]
        if module.category == ModuleCategory.AGGREGATE and node.enabled:
            many_inputs = [
                port for port in module.input_ports if port.cardinality == Cardinality.MANY
            ]
            if many_inputs and not definition.incoming(node.node_id):
                errors.append(f"汇聚节点“{node.label()}”至少需要一个输入")
    return errors
