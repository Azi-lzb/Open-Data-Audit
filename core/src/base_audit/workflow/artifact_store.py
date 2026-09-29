"""Artifact 存储：发布、检索、谱系登记与 TEMPORARY 清理。

只保存路径/哈希/状态/谱系等元数据；Excel 二进制永远不进入本存储。
接口可替换为 SQLite 实现（persistence 预留）。
"""

from __future__ import annotations

from typing import Dict, List, Optional

from .artifacts import Artifact
from .types import ArtifactStatus, Retention


class ArtifactStore:
    def __init__(self) -> None:
        self._artifacts: Dict[str, Artifact] = {}
        self._by_port: Dict[str, List[Artifact]] = {}

    def publish(self, artifact: Artifact) -> Artifact:
        artifact = artifact.as_valid()
        self._artifacts[artifact.artifact_id] = artifact
        key = f"{artifact.producer_node_id}.{artifact.producer_port}"
        self._by_port.setdefault(key, []).append(artifact)
        return artifact

    def get(self, artifact_id: str) -> Optional[Artifact]:
        return self._artifacts.get(artifact_id)

    def latest(self, node_id: str, port: str) -> Optional[Artifact]:
        items = self._by_port.get(f"{node_id}.{port}") or []
        return items[-1] if items else None

    def all_of_node(self, node_id: str) -> List[Artifact]:
        return [
            item for key, items in self._by_port.items()
            if key.split(".", 1)[0] == node_id for item in items
        ]

    def artifacts(self) -> List[Artifact]:
        return list(self._artifacts.values())

    def cleanup_temporary(self) -> List[str]:
        """运行结束清理 TEMPORARY 产物：置为 DELETED 并返回编号列表。"""
        cleaned: List[str] = []
        for artifact in self._artifacts.values():
            if artifact.retention == Retention.TEMPORARY and artifact.status not in (
                ArtifactStatus.DELETED,
            ):
                artifact.status = ArtifactStatus.DELETED
                cleaned.append(artifact.artifact_id)
        return cleaned
