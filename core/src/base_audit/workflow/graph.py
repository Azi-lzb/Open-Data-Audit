"""工作流图模型：节点、边、流程定义与 JSON（schemaVersion=2）序列化。

连接只依赖 node_id + port_name；禁止保存“输入=前序功能名”的旧模型。
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Any, Dict, List, Tuple

from .types import SCHEMA_VERSION, FailurePolicy


@dataclass
class WorkflowNode:
    node_id: str
    module_id: str
    display_name: str = ""
    module_version: str = ""          # 空 = 用注册表最新版本
    preset_id: str = ""
    enabled: bool = True
    parameters: Dict[str, Any] = field(default_factory=dict)
    failure_policy: FailurePolicy = FailurePolicy.FAIL_WORKFLOW
    # 输出策略：{"端口名": {"persist": bool, "retention": "RUN"|..., "promote_to_current": bool}}
    output_policies: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    ui_position: Dict[str, float] = field(default_factory=dict)

    def label(self) -> str:
        return self.display_name or self.node_id


@dataclass
class WorkflowEdge:
    from_node: str
    from_port: str
    to_node: str
    to_port: str
    edge_id: str = ""                 # 空 = 自动生成
    mapping: Dict[str, str] = field(default_factory=dict)
    condition: str = ""

    def identity(self) -> Tuple[str, str, str, str]:
        return (self.from_node, self.from_port, self.to_node, self.to_port)


@dataclass
class WorkflowBinding:
    """流程级输入/输出绑定：外部参数 ↔ 节点端口。"""

    name: str
    node_id: str
    port: str


@dataclass
class WorkflowDefinition:
    workflow_id: str
    name: str
    version: str = "1"
    description: str = ""
    schema_version: str = SCHEMA_VERSION
    nodes: List[WorkflowNode] = field(default_factory=list)
    edges: List[WorkflowEdge] = field(default_factory=list)
    input_bindings: List[WorkflowBinding] = field(default_factory=list)
    output_bindings: List[WorkflowBinding] = field(default_factory=list)
    settings: Dict[str, Any] = field(default_factory=dict)

    def node(self, node_id: str) -> WorkflowNode | None:
        for item in self.nodes:
            if item.node_id == node_id:
                return item
        return None

    def incoming(self, node_id: str) -> List[WorkflowEdge]:
        return [edge for edge in self.edges if edge.to_node == node_id]

    def outgoing(self, node_id: str) -> List[WorkflowEdge]:
        return [edge for edge in self.edges if edge.from_node == node_id]

    # ---- JSON 序列化（schemaVersion=2；不兼容旧流程配置，不做迁移） ----
    def to_dict(self) -> Dict[str, Any]:
        return {
            "schemaVersion": self.schema_version,
            "workflowId": self.workflow_id,
            "name": self.name,
            "version": self.version,
            "description": self.description,
            "settings": self.settings,
            "nodes": [
                {
                    "nodeId": item.node_id,
                    "moduleId": item.module_id,
                    "moduleVersion": item.module_version,
                    "presetId": item.preset_id,
                    "displayName": item.display_name,
                    "enabled": item.enabled,
                    "parameters": item.parameters,
                    "failurePolicy": item.failure_policy.value,
                    "outputPolicies": item.output_policies,
                    "uiPosition": item.ui_position,
                }
                for item in self.nodes
            ],
            "edges": [
                {
                    "edgeId": item.edge_id,
                    "fromNode": item.from_node,
                    "fromPort": item.from_port,
                    "toNode": item.to_node,
                    "toPort": item.to_port,
                    "mapping": item.mapping,
                    "condition": item.condition,
                }
                for item in self.edges
            ],
            "inputBindings": [
                {"name": item.name, "nodeId": item.node_id, "port": item.port}
                for item in self.input_bindings
            ],
            "outputBindings": [
                {"name": item.name, "nodeId": item.node_id, "port": item.port}
                for item in self.output_bindings
            ],
        }

    @classmethod
    def from_dict(cls, payload: Dict[str, Any]) -> "WorkflowDefinition":
        schema = str(payload.get("schemaVersion") or "")
        if schema != SCHEMA_VERSION:
            raise ValueError(f"不支持的流程 schemaVersion：{schema or '缺失'}（当前支持 {SCHEMA_VERSION}）")
        definition = cls(
            workflow_id=str(payload.get("workflowId") or ""),
            name=str(payload.get("name") or ""),
            version=str(payload.get("version") or "1"),
            description=str(payload.get("description") or ""),
            schema_version=schema,
            settings=dict(payload.get("settings") or {}),
        )
        for item in payload.get("nodes") or []:
            definition.nodes.append(WorkflowNode(
                node_id=str(item["nodeId"]),
                module_id=str(item["moduleId"]),
                module_version=str(item.get("moduleVersion") or ""),
                preset_id=str(item.get("presetId") or ""),
                display_name=str(item.get("displayName") or ""),
                enabled=bool(item.get("enabled", True)),
                parameters=dict(item.get("parameters") or {}),
                failure_policy=FailurePolicy(str(item.get("failurePolicy") or FailurePolicy.FAIL_WORKFLOW.value)),
                output_policies=dict(item.get("outputPolicies") or {}),
                ui_position=dict(item.get("uiPosition") or {}),
            ))
        for item in payload.get("edges") or []:
            definition.edges.append(WorkflowEdge(
                from_node=str(item["fromNode"]),
                from_port=str(item["fromPort"]),
                to_node=str(item["toNode"]),
                to_port=str(item["toPort"]),
                edge_id=str(item.get("edgeId") or ""),
                mapping=dict(item.get("mapping") or {}),
                condition=str(item.get("condition") or ""),
            ))
        for item in payload.get("inputBindings") or []:
            definition.input_bindings.append(WorkflowBinding(item["name"], item["nodeId"], item["port"]))
        for item in payload.get("outputBindings") or []:
            definition.output_bindings.append(WorkflowBinding(item["name"], item["nodeId"], item["port"]))
        return definition

    def replace(self, **changes: Any) -> "WorkflowDefinition":
        return replace(self, **changes)
