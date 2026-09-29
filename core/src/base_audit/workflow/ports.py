"""端口定义与类型兼容判断（集中实现，Handler 不得自行判断类型）。"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Tuple

from .types import SUBTYPE_PARENT, ArtifactSubtype, ArtifactType, Cardinality, Retention


@dataclass(frozen=True)
class InputPortDefinition:
    name: str
    artifact_type: ArtifactType
    accepted_subtypes: Tuple[ArtifactSubtype, ...] = field(default=())
    required: bool = True
    cardinality: Cardinality = Cardinality.ONE
    default_binding: str = ""
    description: str = ""


@dataclass(frozen=True)
class OutputPortDefinition:
    name: str
    artifact_type: ArtifactType
    subtype: ArtifactSubtype | None = None
    persist_default: bool = False
    retention_default: Retention = None  # type: ignore[assignment]
    allow_multiple_consumers: bool = True
    description: str = ""

    def __post_init__(self) -> None:
        if self.retention_default is None:
            object.__setattr__(self, "retention_default", Retention.RUN)


def subtype_of(subtype: ArtifactSubtype | None, artifact_type: ArtifactType) -> bool:
    """子类型是否归属于给定基础类型。"""
    if subtype is None:
        return False
    return SUBTYPE_PARENT.get(subtype) == artifact_type


def is_compatible(
    output: OutputPortDefinition,
    input_def: InputPortDefinition,
    output_subtype: ArtifactSubtype | None = None,
) -> bool:
    """集中判断输出端口能否接入输入端口。

    规则：基础类型必须一致（或输出声明的实际子类型归属输入要求的基础
    类型）；输入声明了 accepted_subtypes 时，实际子类型必须在其中。
    """
    effective_subtype = output_subtype if output_subtype is not None else output.subtype
    if effective_subtype is not None:
        if not subtype_of(effective_subtype, input_def.artifact_type):
            return False
        if input_def.accepted_subtypes and effective_subtype not in input_def.accepted_subtypes:
            return False
        return True
    return output.artifact_type == input_def.artifact_type
