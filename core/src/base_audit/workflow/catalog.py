"""DAG 工作流目录。

默认 DAG 由代码提供；组合工作表的用户方案存放在用户设置中。
本模块不再保存或读取自定义 DAG JSON。
"""

from __future__ import annotations

from pathlib import Path

from .defaults import default_workflows
from .graph import WorkflowDefinition


def list_workflows(history_config_path: Path) -> list[WorkflowDefinition]:
    """Return the fixed DAG catalog; custom graph persistence is disabled."""
    return list(default_workflows().values())


def get_workflow(history_config_path: Path, workflow_id: str) -> WorkflowDefinition:
    builtin = default_workflows().get(workflow_id)
    if builtin is not None:
        return builtin
    raise KeyError("未找到内置 DAG 工作流：{}".format(workflow_id))


def save_custom_workflow(history_config_path: Path, definition: WorkflowDefinition) -> Path:
    raise ValueError("当前版本不保存自定义 DAG；固定流程由程序维护")


def delete_custom_workflow(history_config_path: Path, workflow_id: str) -> None:
    raise ValueError("当前版本不支持删除 DAG；固定流程由程序维护")
