from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


class ConfiguredAuditError(Exception):
    """Base class for the independent XLSX-configured audit engine."""


class ConfigurationError(ConfiguredAuditError):
    """The workbook configuration or expression is invalid."""

    def __init__(self, message: str, *, locations: list[dict[str, Any]] | None = None):
        self.message = message
        self.locations = locations or []
        super().__init__(self._display_message())

    def _display_message(self) -> str:
        if not self.locations:
            return self.message
        first = self.locations[0]
        parts = [f"第{first['row']}行"] if first.get("row") else []
        if first.get("column"):
            parts.append(f"第{first['column']}列")
        if first.get("field"):
            parts.append(first["field"])
        if first.get("rule_id"):
            parts.append(f"规则 {first['rule_id']}")
        where = first.get("sheet", "") + (f"!{first['cell']}" if first.get("cell") else "")
        detail = f"{where}（{'，'.join(parts)}）：{self.message}"
        related = [f"{item['sheet']}!{item['cell']}" for item in self.locations[1:] if item.get("cell")]
        if related:
            detail += "；关联单元格：" + "、".join(dict.fromkeys(related))
        return detail

    def as_dict(self) -> dict[str, Any]:
        return {"message": self.message, "locations": self.locations}


class DSLParseError(ConfigurationError):
    """A DSL expression is not in the supported grammar."""


class EvaluationError(ConfiguredAuditError):
    """A well-formed expression could not be evaluated safely."""


@dataclass
class RetrievalError(ConfiguredAuditError):
    code: str
    message: str
    evidence: dict[str, Any] = field(default_factory=dict)

    def __str__(self) -> str:
        return self.message
