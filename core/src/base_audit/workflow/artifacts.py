"""Artifact 模型：节点间传递的数据及谱系/生命周期元数据。

Excel 文件本体留在文件系统，Artifact 只登记路径、哈希、状态、生产节点、
输出端口、生命周期与业务元数据；AUDIT_RESULT_SET 等内存对象直接挂 payload。
"""

from __future__ import annotations

import uuid
from dataclasses import dataclass, field, replace
from datetime import datetime
from typing import Any, Tuple

from .types import ArtifactStatus, ArtifactSubtype, ArtifactType, Retention


def _new_id(prefix: str) -> str:
    return "{}-{}".format(prefix, uuid.uuid4().hex[:12])


@dataclass
class Artifact:
    """所有 Artifact 的公共元数据基类；payload 由具体子类承载。"""

    artifact_type: ArtifactType | None = None
    producer_node_id: str = ""
    producer_port: str = ""
    subtype: ArtifactSubtype | None = None
    artifact_id: str = field(default_factory=lambda: _new_id("art"))
    run_id: str = ""
    created_at: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))
    status: ArtifactStatus = ArtifactStatus.CREATED
    content_uri: str = ""
    content_hash: str = ""
    metadata: dict = field(default_factory=dict)
    lineage: Tuple[str, ...] = field(default_factory=tuple)   # 上游 artifact_id
    retention: Retention = Retention.RUN
    schema_version: str = "1"

    def clone(self, **changes: Any) -> "Artifact":
        return replace(self, **changes)

    def as_valid(self) -> "Artifact":
        return replace(self, status=ArtifactStatus.VALID)


@dataclass
class FileArtifact(Artifact):
    """单个文件；Excel 报送/副本/模板按 subtype 区分。"""

    path: str = ""

    def __post_init__(self) -> None:
        self.artifact_type = ArtifactType.FILE


@dataclass
class FileRecord:
    """文件集里的一条记录：源文件 + 当前工作副本 + 机构信息（对齐旧 records）。"""

    source: str
    audit: str = ""
    org_code: str = ""
    org_name: str = ""
    version: int = 0            # 修改版本号：SOURCE=0，每个 MODIFY 节点 +1
    # 每个版本的实际文件路径。中间版本位于运行临时目录，最终版本位于用户输出目录。
    stage_paths: Tuple[str, ...] = field(default_factory=tuple)
    stage_errors: Tuple[str, ...] = field(default_factory=tuple)


@dataclass
class FileSetArtifact(Artifact):
    """有序文件集；顺序即处理顺序（旧实现按文件名字典序）。"""

    records: Tuple[FileRecord, ...] = field(default_factory=tuple)

    def __post_init__(self) -> None:
        self.artifact_type = ArtifactType.FILE_SET


@dataclass
class MetadataArtifact(Artifact):
    """任意业务元数据（模板定义、外部表计划、历史记录、配置快照等）。"""

    payload: Any = None

    def __post_init__(self) -> None:
        self.artifact_type = ArtifactType.METADATA


@dataclass
class AuditResultSetArtifact(Artifact):
    """审核结果集：Issue 列表（保持旧 result_sets 的顺序语义）。"""

    issues: Tuple[Any, ...] = field(default_factory=tuple)

    def __post_init__(self) -> None:
        self.artifact_type = ArtifactType.AUDIT_RESULT_SET


@dataclass
class StructureMatchArtifact(Artifact):
    """表结构比对结果：[(路径, SourceMatch), ...] 与通过清单。"""

    matches: Tuple[Any, ...] = field(default_factory=tuple)
    accepted: Tuple[str, ...] = field(default_factory=tuple)
    rejected: Tuple[str, ...] = field(default_factory=tuple)

    def __post_init__(self) -> None:
        self.artifact_type = ArtifactType.STRUCTURE_MATCH_RESULT


@dataclass
class OperationResultArtifact(Artifact):
    """操作型结果：修改/写入动作的回执（成功行、失败行、消息）。"""

    ok: bool = True
    message: str = ""
    rows: Tuple[Any, ...] = field(default_factory=tuple)

    def __post_init__(self) -> None:
        self.artifact_type = ArtifactType.OPERATION_RESULT


@dataclass
class ResultWorkbookArtifact(Artifact):
    """汇总/结果工作簿产物。"""

    path: str = ""
    sheet_name: str = ""

    def __post_init__(self) -> None:
        self.artifact_type = ArtifactType.RESULT_WORKBOOK


@dataclass
class LogEventArtifact(Artifact):
    """运行日志分片：一个功能的 (表头, 行) 集合，最终由 log.workbook_write 落盘。"""

    entries: Tuple[Tuple[str, Tuple[str, ...], Tuple[Any, ...]], ...] = field(default_factory=tuple)

    def __post_init__(self) -> None:
        self.artifact_type = ArtifactType.LOG_EVENT_SET
