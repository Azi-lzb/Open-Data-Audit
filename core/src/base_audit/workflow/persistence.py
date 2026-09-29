"""工作流 JSON 导入导出（schemaVersion=2）与运行记录持久化接口。

ModuleDefinition 存代码注册表；ModulePreset/WorkflowDefinition 存 JSON；
运行状态与 Artifact 元数据经 ArtifactStore 接口保存，可替换为 SQLite
（禁止把 Excel 二进制存入数据库，文件只登记路径与哈希）。
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict

from .graph import WorkflowDefinition


def workflow_to_json(definition: WorkflowDefinition) -> str:
    return json.dumps(definition.to_dict(), ensure_ascii=False, indent=2)


def workflow_from_json(text: str) -> WorkflowDefinition:
    payload = json.loads(text)
    if not isinstance(payload, dict):
        raise ValueError("流程 JSON 必须是对象")
    return WorkflowDefinition.from_dict(payload)


def save_workflow(definition: WorkflowDefinition, path: Path) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(workflow_to_json(definition), encoding="utf-8")
    return path


def load_workflow(path: Path) -> WorkflowDefinition:
    return workflow_from_json(Path(path).read_text(encoding="utf-8"))


def run_record_payload(run_id: str, workflow_id: str, node_runs) -> Dict[str, Any]:
    """运行记录的 JSON 视图（workflow_run / node_run 两级），供 SQLite 落库。"""
    return {
        "runId": run_id,
        "workflowId": workflow_id,
        "nodeRuns": [
            {
                "nodeId": run.node_id,
                "status": run.status.value,
                "reason": run.reason,
                "error": run.error,
                "metrics": run.metrics,
            }
            for run in node_runs
        ],
    }
