"""Host-injected paths and cooperative cancellation for the configured engine."""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


class AuditCancelled(RuntimeError):
    """Cancellation is a task state, never a business-rule failure."""


def check_cancel(cancel_event=None) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise AuditCancelled("用户请求终止；尚未处理的文件和规则未执行")


def default_data_dir() -> Path:
    return Path(os.environ.get("BASE_AUDIT_DATA_DIR") or Path.home() / ".audit-tools") / "s5"


@dataclass(frozen=True)
class RuntimePaths:
    template_path: Path | None = None
    template_dir: Path | None = None
    schema_snapshot_dir: Path | None = None
