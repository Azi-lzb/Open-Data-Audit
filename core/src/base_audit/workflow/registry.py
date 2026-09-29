"""模块注册表：module_id + version → (定义, handler)。"""

from __future__ import annotations

from typing import Callable, Dict, Tuple

from .modules import ModuleDefinition


class ModuleRegistry:
    def __init__(self) -> None:
        self._modules: Dict[Tuple[str, str], ModuleDefinition] = {}
        self._handlers: Dict[str, Callable] = {}
        self._latest: Dict[str, str] = {}

    def register(self, definition: ModuleDefinition, handler: Callable) -> None:
        key = (definition.module_id, definition.version)
        if key in self._modules:
            raise ValueError(f"模块重复注册：{definition.module_id} v{definition.version}")
        self._modules[key] = definition
        self._handlers[definition.module_id] = handler
        known = self._latest.get(definition.module_id)
        if known is None or definition.version > known:
            self._latest[definition.module_id] = definition.version

    def get(self, module_id: str, version: str = "") -> ModuleDefinition:
        resolved = version or self._latest.get(module_id)
        if resolved is None:
            raise KeyError(f"模块未注册：{module_id}")
        try:
            return self._modules[(module_id, resolved)]
        except KeyError:
            raise KeyError(f"模块版本不存在：{module_id} v{resolved}")

    def handler(self, module_id: str) -> Callable:
        try:
            return self._handlers[module_id]
        except KeyError:
            raise KeyError(f"模块未注册：{module_id}")

    def module_ids(self) -> list:
        return sorted(self._latest)
