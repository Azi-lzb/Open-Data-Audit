"""把普通编辑器的业务流程草稿编译为显式端口 DAG。"""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field
from typing import Any, Dict, List

from .defaults import default_workflows
from .functions import build_business_flow_templates, build_function_catalog
from .graph import WorkflowDefinition
from .types import FailurePolicy


@dataclass
class BusinessStepDraft:
    node_id: str
    function_id: str
    display_name: str = ""
    enabled: bool = True
    parameters: Dict[str, Any] = field(default_factory=dict)
    failure_policy: str = ""

    @classmethod
    def from_dict(cls, value: Dict[str, Any]) -> "BusinessStepDraft":
        return cls(
            node_id=str(value.get("nodeId") or ""),
            function_id=str(value.get("functionId") or ""),
            display_name=str(value.get("displayName") or ""),
            enabled=bool(value.get("enabled", True)),
            parameters=dict(value.get("parameters") or {}),
            failure_policy=str(value.get("failurePolicy") or ""),
        )

    def to_dict(self) -> Dict[str, Any]:
        return {
            "nodeId": self.node_id,
            "functionId": self.function_id,
            "displayName": self.display_name,
            "enabled": self.enabled,
            "parameters": deepcopy(self.parameters),
            "failurePolicy": self.failure_policy,
        }


@dataclass
class BusinessWorkflowDraft:
    workflow_id: str
    name: str
    template_id: str
    steps: List[BusinessStepDraft] = field(default_factory=list)

    @classmethod
    def from_dict(cls, value: Dict[str, Any]) -> "BusinessWorkflowDraft":
        return cls(
            workflow_id=str(value.get("workflowId") or ""),
            name=str(value.get("name") or ""),
            template_id=str(value.get("templateId") or ""),
            steps=[BusinessStepDraft.from_dict(item) for item in value.get("steps") or []],
        )

    def to_dict(self) -> Dict[str, Any]:
        return {
            "workflowId": self.workflow_id,
            "name": self.name,
            "templateId": self.template_id,
            "steps": [item.to_dict() for item in self.steps],
        }


def default_business_draft(template_id: str, workflow_id: str, name: str = "") -> BusinessWorkflowDraft:
    templates = build_business_flow_templates()
    functions = build_function_catalog()
    try:
        template = templates[template_id]
    except KeyError:
        raise ValueError("未知的普通流程模板：{}".format(template_id))
    base = default_workflows()[template.base_workflow_id]
    steps = []
    for node_id, function_id in template.step_functions:
        node = base.node(node_id)
        if node is None:
            raise ValueError("流程模板缺少节点：{}".format(node_id))
        steps.append(BusinessStepDraft(
            node_id=node_id,
            function_id=function_id,
            display_name=node.display_name or functions[function_id].name,
            enabled=node.enabled,
            parameters=deepcopy(node.parameters),
            failure_policy=node.failure_policy.value,
        ))
    return BusinessWorkflowDraft(
        workflow_id=workflow_id,
        name=name or template.name,
        template_id=template_id,
        steps=steps,
    )


def compile_business_workflow(draft: BusinessWorkflowDraft) -> WorkflowDefinition:
    if not draft.workflow_id.startswith("custom:"):
        raise ValueError("普通编辑器创建的流程编号必须以 custom: 开头")
    if not draft.name.strip():
        raise ValueError("流程名称不能为空")
    templates = build_business_flow_templates()
    functions = build_function_catalog()
    try:
        template = templates[draft.template_id]
    except KeyError:
        raise ValueError("未知的普通流程模板：{}".format(draft.template_id))

    definition = WorkflowDefinition.from_dict(
        default_workflows()[template.base_workflow_id].to_dict()
    )
    definition.workflow_id = draft.workflow_id
    definition.name = draft.name.strip()
    definition.description = "普通流程编辑器创建：基于{}。".format(template.name)

    expected = dict(template.step_functions)
    supplied = {item.node_id: item for item in draft.steps}
    unknown = sorted(set(supplied) - set(expected))
    if unknown:
        raise ValueError("流程包含模板外节点：{}".format("、".join(unknown)))

    normalized_steps: List[BusinessStepDraft] = []
    for node in definition.nodes:
        function_id = expected[node.node_id]
        function = functions[function_id]
        step = supplied.get(node.node_id) or BusinessStepDraft(
            node.node_id, function_id, node.display_name, node.enabled,
            deepcopy(node.parameters), node.failure_policy.value,
        )
        if step.function_id and step.function_id != function_id:
            raise ValueError("节点{}的业务功能与模板不一致".format(node.node_id))
        if not step.enabled and not function.allow_disable:
            raise ValueError("业务功能“{}”是该流程的必需步骤，不能禁用".format(function.name))
        unknown_parameters = sorted(set(step.parameters) - set(node.parameters) - set(function.config_schema))
        if unknown_parameters:
            raise ValueError(
                "业务功能“{}”包含不支持的参数：{}".format(function.name, "、".join(unknown_parameters))
            )
        node.preset_id = function_id
        node.display_name = step.display_name.strip() or function.name
        node.enabled = step.enabled
        node.parameters.update(function.default_parameters)
        node.parameters.update(step.parameters)
        if step.failure_policy:
            node.failure_policy = FailurePolicy(step.failure_policy)
        normalized_steps.append(BusinessStepDraft(
            node.node_id, function_id, node.display_name, node.enabled,
            deepcopy(node.parameters), node.failure_policy.value,
        ))

    normalized = BusinessWorkflowDraft(
        workflow_id=definition.workflow_id,
        name=definition.name,
        template_id=draft.template_id,
        steps=normalized_steps,
    )
    definition.settings = dict(definition.settings)
    definition.settings.update({
        "base_workflow_id": template.base_workflow_id,
        "business_designer": normalized.to_dict(),
    })
    return definition


def business_draft_from_workflow(definition: WorkflowDefinition) -> Dict[str, Any] | None:
    value = definition.settings.get("business_designer")
    return deepcopy(value) if isinstance(value, dict) else None
