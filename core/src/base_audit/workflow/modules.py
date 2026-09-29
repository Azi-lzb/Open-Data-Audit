"""模块定义（代码注册的固定能力）与功能预置（业务实例）。"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Tuple

from .ports import InputPortDefinition, OutputPortDefinition
from .types import ExecutionScope, ModuleCategory


@dataclass(frozen=True)
class ModuleDefinition:
    module_id: str
    version: str
    name: str
    category: ModuleCategory
    description: str = ""
    handler_id: str = ""
    execution_scope: ExecutionScope = ExecutionScope.PER_BATCH
    platform_capabilities: Tuple[str, ...] = field(default_factory=tuple)   # 空 = 全平台
    input_ports: Tuple[InputPortDefinition, ...] = field(default_factory=tuple)
    output_ports: Tuple[OutputPortDefinition, ...] = field(default_factory=tuple)
    parameter_schema: Dict[str, Any] = field(default_factory=dict)
    side_effects: Tuple[str, ...] = field(default_factory=tuple)
    # current_file_policy：normal 模式下该模块是否把输出生效为“当前文件”。
    # 与指南第八节对应：MODIFY 默认 promote_to_current=True。
    current_file_policy: Dict[str, Any] = field(default_factory=dict)

    def port(self, name: str, is_input: bool):
        ports = self.input_ports if is_input else self.output_ports
        for item in ports:
            if item.name == name:
                return item
        return None

    def has_capability(self, capability: str) -> bool:
        return not self.platform_capabilities or capability in self.platform_capabilities


@dataclass(frozen=True)
class ModulePreset:
    """功能预置：某模块在业务上的一组默认参数（相当于旧“模块化功能”行）。"""

    preset_id: str
    module_id: str
    name: str
    description: str = ""
    parameters: Dict[str, Any] = field(default_factory=dict)
    default_output_policies: Dict[str, Any] = field(default_factory=dict)
