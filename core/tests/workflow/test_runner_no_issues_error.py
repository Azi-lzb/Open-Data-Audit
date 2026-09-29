# -*- coding: utf-8 -*-
"""run_dag_native_audit 缺少 issues 产物时的报错路径回归。

节点按失败策略跳过后，流程整体不置 FAILED 但下游必备产物缺失；此时必须把
节点失败原因（如 LibreOffice 计算失败）通过 DagFlowError 带给用户，而不是
抛裸 KeyError('issues') 把真实错误吞掉。
"""

from __future__ import annotations

import pytest

from base_audit.workflow.runner import DagFlowError, run_dag_native_audit
from base_audit.workflow.scheduler import NodeRun, NodeRunStatus, WorkflowRunResult


def test_missing_issues_reports_node_reasons(monkeypatch, tmp_path):
    import base_audit.native.history_io as history_io
    import base_audit.workflow.runner as runner

    monkeypatch.setattr(history_io, "read_history_xlsx", lambda path: [])
    monkeypatch.setattr(
        "base_audit.engines.calculation_engine_preflight",
        lambda *a, **k: (True, "预检通过"),
    )
    result = WorkflowRunResult(
        status=NodeRunStatus.SUCCEEDED,
        node_runs=[
            NodeRun("公式校验复制", NodeRunStatus.FAILED,
                    error="LibreOffice 未能计算审核版.xlsx（GLIBC_2.34 not found）"),
            NodeRun("history-enrich", NodeRunStatus.SKIPPED,
                    reason="上游未产出必填输入：enriched"),
        ],
        outputs={},
    )
    monkeypatch.setattr(runner, "run_workflow", lambda *args, **kwargs: result)
    monkeypatch.setattr(runner, "build_registry", lambda *args, **kwargs: None)

    with pytest.raises(DagFlowError) as excinfo:
        run_dag_native_audit(
            service=object(), template_path=tmp_path / "t.xlsx", input_dir=tmp_path,
            output_dir=tmp_path, period="2026-07", history_path=tmp_path / "h.xlsx",
        )
    message = str(excinfo.value)
    assert "缺少 issues 产物" in message
    assert "公式校验复制" in message
    assert "LibreOffice" in message


def test_s1f1_preflight_blocks_before_workflow(monkeypatch, tmp_path):
    """引擎预检失败时，S1F1 在进入 DAG 之前即失败并携带原因。"""
    import base_audit.engines as engines
    import base_audit.native.history_io as history_io
    import base_audit.workflow.runner as runner

    monkeypatch.setattr(history_io, "read_history_xlsx", lambda path: [])
    monkeypatch.setattr(
        engines, "calculation_engine_preflight",
        lambda *a, **k: (False, "LibreOffice 已找到（/opt/libreoffice26.8）但无法启动：GLIBC_2.34 not found"),
    )
    called = []
    monkeypatch.setattr(runner, "run_workflow", lambda *a, **k: called.append(1))
    monkeypatch.setattr(runner, "build_registry", lambda *a, **k: None)

    with pytest.raises(DagFlowError) as excinfo:
        run_dag_native_audit(
            service=object(), template_path=tmp_path / "t.xlsx", input_dir=tmp_path,
            output_dir=tmp_path, period="2026-07", history_path=tmp_path / "h.xlsx",
        )
    assert not called, "预检失败时不得进入 DAG 调度"
    assert "未开始" in str(excinfo.value)
    assert "GLIBC_2.34" in str(excinfo.value)
