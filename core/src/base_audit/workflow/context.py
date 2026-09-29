"""运行上下文：RunContext（运行级资源）与 NodeContext（节点级视图）。"""

from __future__ import annotations

import uuid
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, Tuple

from .artifact_store import ArtifactStore
from .artifacts import Artifact
from .graph import WorkflowNode


@dataclass
class RunContext:
    workflow_id: str
    store: ArtifactStore
    platform_capabilities: Tuple[str, ...] = field(default_factory=tuple)
    settings: Dict[str, Any] = field(default_factory=dict)
    services: Dict[str, Any] = field(default_factory=dict)      # ExcelSession / AuditService 等
    resources: Dict[str, Any] = field(default_factory=dict)     # FeatureLog、计时等运行级资源
    run_id: str = field(default_factory=lambda: "run-" + uuid.uuid4().hex[:12])
    log: Callable[[str], None] = field(default=lambda _text: None)
    node_runs: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    # 普通流程编辑器的“当前文件”只是端口自动绑定的快捷视图；它从不替代
    # DAG 边。在高级模式中，调用方不使用这个槽位，必须显式连线。
    current_files: Dict[str, Artifact] = field(default_factory=dict)

    def require_service(self, name: str):
        """取运行级共享资源（如 Excel 会话）；缺失即明确失败，不静默重建。"""
        try:
            return self.services[name]
        except KeyError:
            raise RuntimeError(f"运行级资源“{name}”尚未初始化，节点无法执行")

    def promote_current_file(self, artifact: Artifact, slot: str = "current") -> None:
        self.current_files[slot] = artifact

    def current_file(self, slot: str = "current") -> Artifact | None:
        return self.current_files.get(slot)


@dataclass
class NodeContext:
    run: RunContext
    node: WorkflowNode

    @property
    def store(self) -> ArtifactStore:
        return self.run.store

    @property
    def parameters(self) -> Dict[str, Any]:
        return self.node.parameters

    def log(self, text: str) -> None:
        if self.run.log:
            self.run.log(f"[{self.node.label()}] {text}")

    def resource(self, name: str):
        return self.run.require_service(name)
