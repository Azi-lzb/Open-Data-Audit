"""workflow 引擎单元测试（阶段 1-2）。

覆盖实现指南第十八节要求的全部用例：合法线性图、分支、MANY 汇聚、环、
缺失必填端口、类型不兼容、ONE 多连接、禁用节点断路、模块缺失、平台能力
缺失、CurrentFile 冲突、四种失败策略、PARTIAL、临时产物清理、持久产物
保留和文件谱系。
"""

from __future__ import annotations

import sys
import threading
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from base_audit.workflow import scheduler as wf_scheduler
from base_audit.workflow.artifacts import (
    Artifact,
    FileArtifact,
    FileSetArtifact,
    FileRecord,
    MetadataArtifact,
)
from base_audit.workflow.context import RunContext
from base_audit.workflow.graph import (
    WorkflowBinding,
    WorkflowDefinition,
    WorkflowEdge,
    WorkflowNode,
)
from base_audit.workflow.modules import ModuleCategory, ModuleDefinition
from base_audit.workflow.artifact_store import ArtifactStore
from base_audit.workflow.registry import ModuleRegistry
from base_audit.workflow.scheduler import (
    ModuleExecutionResult,
    WorkflowExecutionError,
    WorkflowValidationError,
    run_workflow,
)
from base_audit.workflow.types import (
    ArtifactSubtype,
    ArtifactType,
    Cardinality,
    ExecutionScope,
    FailurePolicy,
    NodeRunStatus,
    Retention,
)
from base_audit.workflow.ports import InputPortDefinition, OutputPortDefinition
from base_audit.workflow.validation import validate_workflow


def test_union_combine_preset_reuses_combine_sheets_module():
    """联合核查表只是组合工作表的固定正则预置，不能再有第二套模块。"""
    from base_audit.workflow.defaults import default_workflows

    workflow = default_workflows()["dag:组合联合核查表"]
    assert len(workflow.nodes) == 1
    assert workflow.nodes[0].module_id == "excel.combine_sheets"
    # 联合核查表预置的方案名已中文化：plan_id 即“组合方案”表里的固定行。
    assert workflow.nodes[0].parameters["plan_id"] == "组合工作表（核查表）"


def _module(
    module_id="m",
    category=ModuleCategory.CONTROL,
    inputs=(),
    outputs=(),
    scope=ExecutionScope.PER_BATCH,
    capabilities=(),
    output_name=None,
):
    parameters = {"output_name": output_name} if output_name else {}
    return ModuleDefinition(
        module_id=module_id,
        version="1",
        name=module_id,
        category=category,
        execution_scope=scope,
        platform_capabilities=capabilities,
        input_ports=tuple(inputs),
        output_ports=tuple(outputs),
        parameter_schema=parameters,
    )


def _registry_with(handler, **modules) -> ModuleRegistry:
    registry = ModuleRegistry()
    for definition in modules.values():
        registry.register(definition, handler)
    return registry


IN_FILE = InputPortDefinition("in", ArtifactType.FILE_SET)
OUT_SET = OutputPortDefinition("out", ArtifactType.FILE_SET)


def _linear_definition(**node_overrides):
    module = _module("m", inputs=(IN_FILE,), outputs=(OUT_SET,))
    definition = WorkflowDefinition(workflow_id="wf", name="测试")
    definition.nodes = [
        WorkflowNode("a", "m", display_name="甲"),
        WorkflowNode("b", "m", display_name="乙"),
    ]
    definition.edges = [WorkflowEdge("a", "out", "b", "in")]
    for node_id, changes in node_overrides.items():
        for node in definition.nodes:
            if node.node_id == node_id:
                for key, value in changes.items():
                    setattr(node, key, value)
    return definition, module


def _run(definition, registry, handler, inputs=None, capabilities=()):
    context = RunContext(
        workflow_id=definition.workflow_id,
        store=ArtifactStore(),
        platform_capabilities=capabilities,
    )
    registry.register(registry.get("m"), handler) if "m" in registry.module_ids() else None
    return run_workflow(definition, registry, context, inputs or {})


def test_linear_graph_runs_in_order():
    calls = []

    def handler(ctx, parameters, inputs):
        calls.append(ctx.node.node_id)
        previous = inputs.get("in")
        return ModuleExecutionResult.ok(out=FileSetArtifact(records=tuple(previous.records) if previous else ()))

    module = _module("m", inputs=(IN_FILE,), outputs=(OUT_SET,))
    registry = _registry_with(handler, m=module)
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [WorkflowNode("a", "m"), WorkflowNode("b", "m"), WorkflowNode("c", "m")]
    definition.edges = [
        WorkflowEdge("a", "out", "b", "in"),
        WorkflowEdge("b", "out", "c", "in"),
    ]
    context = RunContext(workflow_id="wf", store=ArtifactStore())
    inputs = {"source": FileSetArtifact(records=(FileRecord(source="x.xlsx"),))}
    definition.input_bindings = [WorkflowBinding("source", "a", "in")]
    result = run_workflow(definition, registry, context, inputs)
    assert result.status == NodeRunStatus.SUCCEEDED
    assert calls == ["a", "b", "c"]
    assert all("elapsed_seconds" in run.metrics for run in result.node_runs)
    assert all(run.metrics["elapsed_seconds"] >= 0 for run in result.node_runs)


def test_cancel_request_stops_before_next_node():
    calls = []
    cancel_event = threading.Event()

    def handler(ctx, _parameters, _inputs):
        calls.append(ctx.node.node_id)
        cancel_event.set()
        return ModuleExecutionResult.ok(out=FileSetArtifact())

    module = _module("m", outputs=(OUT_SET,))
    registry = _registry_with(handler, m=module)
    definition = WorkflowDefinition(workflow_id="wf", name="终止测试")
    definition.nodes = [WorkflowNode("a", "m"), WorkflowNode("b", "m")]
    context = RunContext(workflow_id="wf", store=ArtifactStore())
    context.resources["cancel_event"] = cancel_event

    result = run_workflow(definition, registry, context, {})

    assert calls == ["a"]
    assert result.status == NodeRunStatus.FAILED
    assert result.error.startswith("用户请求终止")
    assert result.node_runs[-1].node_id == "b"
    assert result.node_runs[-1].reason == "用户请求终止"


def test_windows_formula_node_uses_com_copy_and_one_workbook_at_a_time(tmp_path):
    """回归：不得用 openpyxl 改写整份副本，也不得同时打开多个副本计算。"""
    from openpyxl import Workbook
    from base_audit.models import AuditRule, CopyRange, TemplateDefinition
    from base_audit.workflow.adapters import audit_handlers

    template = tmp_path / "模板.xlsx"
    book = Workbook(); book.active.title = "数据"; book.active["B1"] = '=IF(参照表!A1>0,"错|指标|说明|1|0|1","")'; book.save(template); book.close()
    records = []
    for name in ("甲.xlsx", "乙.xlsx"):
        path = tmp_path / name
        book = Workbook(); book.active.title = "数据"; book.save(path); book.close()
        records.append(FileRecord(source=str(path), audit=str(path), stage_paths=(str(path),)))
    definition = TemplateDefinition(
        rules=[AuditRule("R1", True, "", "数据", "B1", "B1", "错误", "说明", "B1")],
        copy_ranges=[CopyRange("数据", "B1")], structured=False,
    )

    class FakeWorkbook:
        def Save(self):
            pass

    class FakeExcel:
        def __init__(self):
            self.apply_calls = 0; self.calculate_calls = 0; self.active = 0; self.max_active = 0
            self.formula_overrides = []
        def open_workbook(self, _path, read_only=False, purpose=""):
            self.active += 1; self.max_active = max(self.max_active, self.active); return FakeWorkbook()
        def close_workbook(self, _book):
            self.active -= 1
        def apply_rules(self, _template, _audit, _definition, calculate=True, formula_overrides=None):
            self.apply_calls += 1
            self.formula_overrides.append(formula_overrides)
        def calculate_workbook(self, _workbook):
            self.calculate_calls += 1

    fake = FakeExcel()
    context = RunContext(workflow_id="wf", store=ArtifactStore(), services={"excel": fake})
    node_context = audit_handlers.NodeContext(context, WorkflowNode("formula", "excel.formula_copy_recalculate"))
    audit_handlers.handle_formula_copy(node_context, {
        "template_path": str(template), "audit_output_dir": str(tmp_path / "输出"),
        "feature_name": "公式校验复制", "on_failure": "跳过", "flow_name": "测试",
    }, {
        "template": MetadataArtifact(payload={"template_path": str(template), "definition": definition}),
        "workbooks": FileSetArtifact(records=tuple(records)),
    })

    assert fake.apply_calls == 2
    assert fake.calculate_calls == 2
    assert fake.max_active <= 2  # 公式模板 + 当前审核副本
    expected_formula = '=IF(参照表!A1>0,"错|指标|说明|1|0|1","")'
    assert all(item[("数据", "B1")] == expected_formula for item in fake.formula_overrides)


def test_branch_and_many_merge():
    """分支 + MANY 汇聚：两个上游产物都进入合并节点。"""
    seen = {}

    def source_handler(ctx, parameters, inputs):
        marker = ctx.node.node_id
        return ModuleExecutionResult.ok(out=FileSetArtifact(records=(FileRecord(source=marker),)))

    def merge_handler(ctx, parameters, inputs):
        items = inputs["in"] if isinstance(inputs["in"], list) else [inputs["in"]]
        merged = tuple(record for item in items for record in item.records)
        seen[ctx.node.node_id] = sorted(record.source for record in merged)
        return ModuleExecutionResult.ok(out=FileSetArtifact(records=merged))

    source_module = _module("src", category=ModuleCategory.SOURCE, outputs=(OUT_SET,))
    merge_module = _module(
        "merge", category=ModuleCategory.AGGREGATE,
        inputs=(InputPortDefinition("in", ArtifactType.FILE_SET, cardinality=Cardinality.MANY),),
        outputs=(OUT_SET,),
    )
    registry = ModuleRegistry()
    registry.register(source_module, source_handler)
    registry.register(merge_module, merge_handler)
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [
        WorkflowNode("a", "src"), WorkflowNode("b", "src"), WorkflowNode("m", "merge"),
    ]
    definition.edges = [
        WorkflowEdge("a", "out", "m", "in"),
        WorkflowEdge("b", "out", "m", "in"),
    ]
    context = RunContext(workflow_id="wf", store=ArtifactStore())
    result = run_workflow(definition, registry, context, {})
    assert result.status == NodeRunStatus.SUCCEEDED
    assert seen["m"] == ["a", "b"]


def test_cycle_detected_by_validation():
    module = _module("m", inputs=(IN_FILE,), outputs=(OUT_SET,))
    registry = _registry_with(lambda ctx, p, i: ModuleExecutionResult.ok(), m=module)
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [WorkflowNode("a", "m"), WorkflowNode("b", "m")]
    definition.edges = [
        WorkflowEdge("a", "out", "b", "in"),
        WorkflowEdge("b", "out", "a", "in"),
    ]
    errors = validate_workflow(definition, registry)
    assert any("环" in item for item in errors)


def test_missing_required_port_fails_validation():
    module = _module("m", inputs=(IN_FILE,), outputs=(OUT_SET,))
    registry = _registry_with(lambda ctx, p, i: ModuleExecutionResult.ok(), m=module)
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [WorkflowNode("a", "m")]
    errors = validate_workflow(definition, registry)
    assert any("必填输入" in item for item in errors)


def test_type_incompatible_edge_rejected():
    file_module = _module(
        "file", outputs=(OutputPortDefinition("out", ArtifactType.FILE),),
    )
    set_module = _module("set", inputs=(IN_FILE,), outputs=(OUT_SET,))
    registry = ModuleRegistry()
    registry.register(file_module, lambda ctx, p, i: ModuleExecutionResult.ok())
    registry.register(set_module, lambda ctx, p, i: ModuleExecutionResult.ok())
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [WorkflowNode("a", "file"), WorkflowNode("b", "set")]
    definition.edges = [WorkflowEdge("a", "out", "b", "in")]
    errors = validate_workflow(definition, registry)
    assert any("类型不兼容" in item for item in errors)


def test_one_input_rejects_multiple_edges():
    module = _module("m", inputs=(IN_FILE,), outputs=(OUT_SET,))
    registry = _registry_with(lambda ctx, p, i: ModuleExecutionResult.ok(), m=module)
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [
        WorkflowNode("a", "m"), WorkflowNode("b", "m"), WorkflowNode("c", "m"),
    ]
    definition.edges = [
        WorkflowEdge("a", "out", "c", "in"),
        WorkflowEdge("b", "out", "c", "in"),
    ]
    errors = validate_workflow(definition, registry)
    assert any("只允许一条" in item for item in errors)


def test_disabled_upstream_breaks_required_input():
    module = _module("m", inputs=(IN_FILE,), outputs=(OUT_SET,))
    registry = _registry_with(lambda ctx, p, i: ModuleExecutionResult.ok(), m=module)
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [WorkflowNode("a", "m", enabled=False), WorkflowNode("b", "m")]
    definition.edges = [WorkflowEdge("a", "out", "b", "in")]
    errors = validate_workflow(definition, registry)
    assert any("上游节点全部被禁用" in item for item in errors)

    # 必填改为可选时不再报错；运行时下游因上游禁用而跳过
    optional_module = _module(
        "mo",
        inputs=(InputPortDefinition("in", ArtifactType.FILE_SET, required=False),),
        outputs=(OUT_SET,),
    )
    registry2 = _registry_with(lambda ctx, p, i: ModuleExecutionResult.ok(), mo=optional_module)
    definition2 = WorkflowDefinition(workflow_id="wf", name="t")
    definition2.nodes = [WorkflowNode("a", "mo", enabled=False), WorkflowNode("b", "mo")]
    definition2.edges = [WorkflowEdge("a", "out", "b", "in")]
    assert not validate_workflow(definition2, registry2)
    context = RunContext(workflow_id="wf", store=ArtifactStore())
    result = run_workflow(definition2, registry2, context, {})
    statuses = {run.node_id: run.status for run in result.node_runs}
    assert statuses["a"] == NodeRunStatus.DISABLED
    # 可选输入缺失不阻断执行：b 正常运行，输入里没有 in
    assert statuses["b"] == NodeRunStatus.SUCCEEDED


def test_missing_module_reported():
    registry = ModuleRegistry()
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [WorkflowNode("a", "ghost")]
    errors = validate_workflow(definition, registry)
    assert any("模块不存在" in item for item in errors)


def test_platform_capability_missing_reported():
    module = _module("m", inputs=(IN_FILE,), outputs=(OUT_SET,), capabilities=("win32",))
    registry = _registry_with(lambda ctx, p, i: ModuleExecutionResult.ok(), m=module)
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [WorkflowNode("a", "m")]
    definition.input_bindings = [WorkflowBinding("src", "a", "in")]
    errors = validate_workflow(definition, registry, platform_capabilities=("native",))
    assert any("平台能力" in item for item in errors)
    assert not validate_workflow(definition, registry, platform_capabilities=("win32",))


def test_output_binding_must_be_reachable():
    module = _module("m", inputs=(IN_FILE,), outputs=(OUT_SET,))
    registry = _registry_with(lambda ctx, p, i: ModuleExecutionResult.ok(), m=module)
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [WorkflowNode("a", "m"), WorkflowNode("orphan", "m")]
    definition.edges = [WorkflowEdge("a", "out", "orphan", "in")]
    definition.input_bindings = [WorkflowBinding("src", "a", "in")]
    definition.output_bindings = []
    # 孤儿链可达（从 a 出发）→ 合法；换一个不可达节点 → 报错
    definition.output_bindings = [WorkflowBinding("result", "orphan", "out")]
    assert not validate_workflow(definition, registry)
    definition.output_bindings = [WorkflowBinding("result", "x-ghost", "out")]
    assert any("无法从流程输入到达" in item or "不存在" in item
               for item in validate_workflow(definition, registry))


def test_current_file_conflict_detected():
    modify = _module(
        "mod", category=ModuleCategory.MODIFY,
        inputs=(IN_FILE,), outputs=(OUT_SET,),
    )
    registry = _registry_with(lambda ctx, p, i: ModuleExecutionResult.ok(out=FileSetArtifact()), m=modify)
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [
        WorkflowNode("a", "mod", output_policies={"out": {"promote_to_current": True}}),
        WorkflowNode("b", "mod", output_policies={"out": {"promote_to_current": True}}),
    ]
    errors = validate_workflow(definition, registry)
    assert any("当前文件" in item for item in errors)
    # 分支到不同槽位不冲突
    for node, slot in zip(definition.nodes, ("out@1", "out@2")):
        node.output_policies = {"out": {"promote_to_current": True, "slot": slot}}
    assert not [e for e in validate_workflow(definition, registry) if "当前文件" in e]


def test_duplicate_output_filename_detected():
    module = _module("m", inputs=(IN_FILE,), outputs=(OUT_SET,))
    registry = _registry_with(lambda ctx, p, i: ModuleExecutionResult.ok(), m=module)
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [
        WorkflowNode("a", "m", parameters={"output_name": "汇总"}),
        WorkflowNode("b", "m", parameters={"output_name": "汇总"}),
    ]
    errors = validate_workflow(definition, registry)
    assert any("输出文件名冲突" in item for item in errors)


def test_sink_requires_inputs():
    sink = _module("sink", category=ModuleCategory.SINK, inputs=(IN_FILE,))
    registry = _registry_with(lambda ctx, p, i: ModuleExecutionResult.ok(), m=sink)
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [WorkflowNode("a", "sink")]
    errors = validate_workflow(definition, registry)
    assert any("必填输入" in item for item in errors)
    assert any("没有任何输入" in item for item in errors)


def test_failure_policy_fail_workflow_stops_run():
    ok_module = _module("ok", inputs=(IN_FILE,), outputs=(OUT_SET,))
    bad_module = _module("bad", inputs=(IN_FILE,), outputs=(OUT_SET,))
    registry = ModuleRegistry()
    registry.register(ok_module, lambda ctx, p, i: ModuleExecutionResult.ok(out=FileSetArtifact()))
    registry.register(bad_module, lambda ctx, p, i: ModuleExecutionResult.failed("引擎不可用"))
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [
        WorkflowNode("a", "ok"), WorkflowNode("b", "bad", failure_policy=FailurePolicy.FAIL_WORKFLOW),
        WorkflowNode("c", "ok"),
    ]
    definition.edges = [WorkflowEdge("a", "out", "b", "in"), WorkflowEdge("b", "out", "c", "in")]
    definition.input_bindings = [WorkflowBinding("src", "a", "in")]
    context = RunContext(workflow_id="wf", store=ArtifactStore())
    result = run_workflow(definition, registry, context, {"src": FileSetArtifact()})
    assert result.status == NodeRunStatus.FAILED
    statuses = {run.node_id: run.status for run in result.node_runs}
    assert statuses["a"] == NodeRunStatus.SUCCEEDED
    assert statuses["b"] == NodeRunStatus.FAILED
    assert statuses["c"] == NodeRunStatus.SKIPPED
    assert "引擎不可用" in result.error


def test_failure_policy_skip_node_lets_rest_continue():
    ok_module = _module("ok", inputs=(IN_FILE,), outputs=(OUT_SET,))
    bad_module = _module("bad", inputs=(IN_FILE,), outputs=(OUT_SET,))
    registry = ModuleRegistry()
    registry.register(ok_module, lambda ctx, p, i: ModuleExecutionResult.ok(out=FileSetArtifact()))
    registry.register(bad_module, lambda ctx, p, i: ModuleExecutionResult.failed("boom"))
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [
        WorkflowNode("a", "ok"), WorkflowNode("b", "bad", failure_policy=FailurePolicy.SKIP_NODE),
        WorkflowNode("c", "ok"),
    ]
    definition.edges = [WorkflowEdge("a", "out", "b", "in"), WorkflowEdge("b", "out", "c", "in")]
    definition.input_bindings = [WorkflowBinding("src", "a", "in")]
    context = RunContext(workflow_id="wf", store=ArtifactStore())
    result = run_workflow(definition, registry, context, {"src": FileSetArtifact()})
    statuses = {run.node_id: run.status for run in result.node_runs}
    assert statuses["a"] == NodeRunStatus.SUCCEEDED
    assert statuses["b"] == NodeRunStatus.SKIPPED
    # c 依赖 b 的输出，b 跳过后 c 也跳过
    assert statuses["c"] == NodeRunStatus.SKIPPED


def test_partial_status_with_item_failures():
    def handler(ctx, parameters, inputs):
        artifact = FileSetArtifact(records=(
            FileRecord(source="a.xlsx"), FileRecord(source="b.xlsx"),
        ))
        return ModuleExecutionResult.partial([("b.xlsx", "结构不匹配")], out=artifact)

    module = _module("m", inputs=(IN_FILE,), outputs=(OUT_SET,))
    registry = ModuleRegistry()
    registry.register(module, handler)
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [WorkflowNode("a", "m")]
    definition.input_bindings = [WorkflowBinding("src", "a", "in")]
    context = RunContext(workflow_id="wf", store=ArtifactStore())
    result = run_workflow(definition, registry, context, {"src": FileSetArtifact()})
    assert result.status == NodeRunStatus.PARTIAL
    assert result.node_runs[0].status == NodeRunStatus.PARTIAL
    assert "结构不匹配" in result.node_runs[0].error


def test_temporary_artifacts_cleaned_and_persistent_kept():
    def handler(ctx, parameters, inputs):
        temp = FileSetArtifact(records=())
        temp.retention = Retention.TEMPORARY
        keep = FileSetArtifact(records=())
        keep.retention = Retention.PERSISTENT
        return ModuleExecutionResult.ok(temp_out=temp, final=keep)

    module = _module(
        "m", inputs=(IN_FILE,), outputs=(
            OutputPortDefinition("temp_out", ArtifactType.FILE_SET),
            OutputPortDefinition("final", ArtifactType.FILE_SET),
        ),
    )
    registry = ModuleRegistry()
    registry.register(module, handler)
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [WorkflowNode("a", "m")]
    definition.input_bindings = [WorkflowBinding("src", "a", "in")]
    context = RunContext(workflow_id="wf", store=ArtifactStore())
    result = run_workflow(definition, registry, context, {"src": FileSetArtifact()})
    statuses = {a.producer_port: a.status for a in context.store.artifacts()}
    from base_audit.workflow.types import ArtifactStatus
    assert statuses["temp_out"] == ArtifactStatus.DELETED
    assert statuses["final"] == ArtifactStatus.VALID


def test_lineage_records_upstream():
    def handler(ctx, parameters, inputs):
        upstream = inputs["in"]
        child = FileSetArtifact(records=upstream.records)
        child.lineage = (upstream.artifact_id,)
        return ModuleExecutionResult.ok(out=child)

    module = _module("m", inputs=(IN_FILE,), outputs=(OUT_SET,))
    registry = ModuleRegistry()
    registry.register(module, handler)
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [WorkflowNode("a", "m"), WorkflowNode("b", "m")]
    definition.edges = [WorkflowEdge("a", "out", "b", "in")]
    definition.input_bindings = [WorkflowBinding("src", "a", "in")]
    context = RunContext(workflow_id="wf", store=ArtifactStore())
    run_workflow(definition, registry, context, {"src": FileSetArtifact()})
    artifacts = context.store.artifacts()
    child = next(a for a in artifacts if a.producer_node_id == "b")
    parent = next(a for a in artifacts if a.producer_node_id == "a")
    assert child.lineage == (parent.artifact_id,)


def test_scheduler_rejects_handler_output_with_wrong_type():
    module = _module("m", inputs=(IN_FILE,), outputs=(OUT_SET,))
    registry = _registry_with(
        lambda ctx, p, i: ModuleExecutionResult.ok(out=FileArtifact(path="wrong.xlsx")),
        m=module,
    )
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [WorkflowNode("a", "m")]
    definition.input_bindings = [WorkflowBinding("src", "a", "in")]
    context = RunContext(workflow_id="wf", store=ArtifactStore())
    with pytest.raises(WorkflowExecutionError, match="类型错误"):
        run_workflow(definition, registry, context, {"src": FileSetArtifact()})


def test_current_file_slot_is_promoted_by_modify_module():
    module = _module("mod", category=ModuleCategory.MODIFY, inputs=(IN_FILE,), outputs=(OUT_SET,))
    module = ModuleDefinition(
        **{**module.__dict__, "current_file_policy": {"promote_to_current": True}}
    )
    registry = _registry_with(
        lambda ctx, p, i: ModuleExecutionResult.ok(out=FileSetArtifact()), m=module,
    )
    definition = WorkflowDefinition(workflow_id="wf", name="t")
    definition.nodes = [WorkflowNode("a", "mod")]
    definition.input_bindings = [WorkflowBinding("src", "a", "in")]
    context = RunContext(workflow_id="wf", store=ArtifactStore())
    run_workflow(definition, registry, context, {"src": FileSetArtifact()})
    assert context.current_file() is not None
    assert context.current_file().producer_node_id == "a"


def test_legacy_chinese_skip_maps_to_per_item_policy():
    from base_audit.workflow.adapters.audit_handlers import _item_failure_policy

    assert _item_failure_policy({"on_failure": "跳过"}) == FailurePolicy.SKIP_ITEM
    assert _item_failure_policy({"on_failure": "停止"}) == FailurePolicy.FAIL_WORKFLOW


def test_native_dag_registry_uses_native_handlers_not_com_or_legacy_flow():
    """UOS DAG 的注册表必须替换节点处理器，而非委托旧流程组合器。"""
    from base_audit.workflow.runner import build_registry

    native = build_registry("native")
    com = build_registry("com")
    assert native.handler("excel.formula_copy_recalculate") is not com.handler("excel.formula_copy_recalculate")
    assert "excel.audit_navigation_write" not in native.module_ids()


def test_native_dag_runs_nodes_without_service_run_flow(tmp_path, monkeypatch):
    """不启动 soffice 的调度冒烟：旧入口一旦被碰到即失败，且保留 v4 审核副本。"""
    from openpyxl import Workbook
    from openpyxl.workbook.defined_name import DefinedName
    from base_audit.name_config import initialize_config
    from base_audit.service import AuditService
    from base_audit.workflow.runner import run_dag_native_audit

    def make_book(path, formula=False):
        book = Workbook(); sheet = book.active; sheet.title = "数据"
        sheet.append(("字段", "数值")); sheet.append(("贷款", 1))
        if formula:
            sheet["D2"] = '=IF(B2>0,"错误|贷款|不应有数|1","")'
            for name, area in (("校验区域", "$D$2"), ("表结构区域", "$A$1:$B$1"), ("条件格式区域", "$B$2")):
                book.defined_names.add(DefinedName(name, attr_text=f"'数据'!{area}"))
        book.save(path); book.close()

    source_dir = tmp_path / "源"; source_dir.mkdir()
    make_book(source_dir / "甲银行.xlsx")
    template = tmp_path / "模板.xlsx"; make_book(template, formula=True)
    config = initialize_config(tmp_path / "配置.xlsx")
    service = AuditService(config_path=config, engine_preference="自动")
    # 此处只替换计算器进程；其余所有节点、端口和文件版本真实运行。
    recalculated = []
    monkeypatch.setattr(
        "base_audit.workflow.adapters.native_audit_handlers.LibreOfficeCalculator.recalculate",
        lambda _self, path: recalculated.append(Path(path)),
    )
    service.run_flow = lambda **_kwargs: (_ for _ in ()).throw(AssertionError("不得调用旧 run_flow"))
    result = run_dag_native_audit(
        service=service, template_path=template, input_dir=source_dir, output_dir=tmp_path / "输出",
        period="2026-06", history_path=config,
    )
    assert result.summary_path and result.summary_path.is_file()
    assert result.files[0].audit_path and result.files[0].audit_path.is_file()
    # 审核导航已移除：最终交付副本就是公式复制后的 v3，仅需一次重算。
    assert len(recalculated) == 1
    from openpyxl import load_workbook
    book = load_workbook(result.files[0].audit_path, read_only=True)
    try:
        assert "审核导航" not in book.sheetnames
    finally:
        book.close()


def test_audit_runner_selects_native_adapter_without_com(monkeypatch, tmp_path):
    from base_audit.workflow import runner

    class NativeService:
        engine_preference = "自动"

    monkeypatch.setattr(runner, "pipeline_kind", lambda _engine: "native")
    monkeypatch.setattr(runner, "run_dag_native_audit", lambda **kwargs: ("native-dag", kwargs))
    result, kwargs = runner.run_dag_audit(
        service=NativeService(), template_path=tmp_path / "模板.xlsx", input_dir=tmp_path,
        output_dir=tmp_path / "输出", period="2026-06", history_path=tmp_path / "历史.xlsx",
    )
    assert result == "native-dag"
    assert kwargs["template_path"].name == "模板.xlsx"


def test_audit_runner_reports_engine_start_before_com_session(monkeypatch, tmp_path):
    """COM 启动卡住/失败时，界面必须先给出可定位的阶段提示。"""
    from base_audit.workflow import runner

    class Service:
        engine_preference = "Microsoft Excel"

        @staticmethod
        def _history_storage(path):
            return path, "核查表校验结果", None

    class FailingSession:
        def __init__(self, _preference):
            pass

        def __enter__(self):
            raise RuntimeError("模拟 COM 启动失败")

        def __exit__(self, *_args):
            return False

    logs = []
    monkeypatch.setattr(runner, "pipeline_kind", lambda _engine: "com")
    monkeypatch.setattr("base_audit.excel_com.ExcelSession", FailingSession)
    with pytest.raises(RuntimeError, match="模拟 COM 启动失败"):
        runner.run_dag_audit(
            service=Service(), template_path=tmp_path / "模板.xlsx", input_dir=tmp_path,
            output_dir=tmp_path / "输出", period="2026-09", history_path=tmp_path / "历史.xlsx",
            on_step=logs.append,
        )
    assert "正在启动表格引擎：Microsoft Excel……" in logs


def test_workflow_json_round_trip(tmp_path):
    from base_audit.workflow.persistence import load_workflow, save_workflow

    module = _module("m", inputs=(IN_FILE,), outputs=(OUT_SET,))
    definition = WorkflowDefinition(workflow_id="wf", name="序列化")
    definition.nodes = [WorkflowNode("a", "m", parameters={"output_name": "结果"})]
    definition.input_bindings = [WorkflowBinding("src", "a", "in")]
    path = save_workflow(definition, tmp_path / "wf.json")
    loaded = load_workflow(path)
    assert loaded.to_dict() == definition.to_dict()
    assert loaded.schema_version == "2"


def test_old_schema_version_rejected():
    from base_audit.workflow.persistence import workflow_from_json

    with pytest.raises(ValueError, match="schemaVersion"):
        workflow_from_json('{"schemaVersion": "1", "workflowId": "x"}')


def test_business_compiler_preserves_validated_audit_graph():
    from base_audit.workflow.compiler import (
        compile_business_workflow,
        default_business_draft,
    )
    from base_audit.workflow.defaults import default_workflows
    from base_audit.workflow.runner import build_registry

    original = default_workflows()["dag:汇总核查表校验"]
    draft = default_business_draft("audit", "custom:audit-copy", "我的审核流程")
    compiled = compile_business_workflow(draft)

    assert [node.module_id for node in compiled.nodes] == [node.module_id for node in original.nodes]
    assert [edge.identity() for edge in compiled.edges] == [edge.identity() for edge in original.edges]
    assert [binding.name for binding in compiled.output_bindings] == [
        binding.name for binding in original.output_bindings
    ]
    assert compiled.settings["base_workflow_id"] == "dag:汇总核查表校验"
    assert all(node.preset_id for node in compiled.nodes)
    assert not validate_workflow(compiled, build_registry(), ("com",))


def test_business_compiler_rejects_disabling_required_step():
    from base_audit.workflow.compiler import (
        compile_business_workflow,
        default_business_draft,
    )

    draft = default_business_draft("summary", "custom:bad", "错误流程")
    draft.steps[0].enabled = False
    with pytest.raises(ValueError, match="不能禁用"):
        compile_business_workflow(draft)


def test_business_compiler_binds_combine_plan_without_changing_module():
    from base_audit.workflow.compiler import (
        compile_business_workflow,
        default_business_draft,
    )

    draft = default_business_draft("combine_sheets", "custom:combine", "杭州组合")
    draft.steps[0].parameters["plan_id"] = "plan-hangzhou"
    compiled = compile_business_workflow(draft)
    assert compiled.nodes[0].module_id == "excel.combine_sheets"
    assert compiled.nodes[0].parameters["plan_id"] == "plan-hangzhou"


def test_custom_dag_route_is_disabled(tmp_path):
    from base_audit.workflow import runner

    with pytest.raises(runner.DagFlowError, match="不支持自定义 DAG"):
        runner.run_dag_flow(
            "custom:audit-route", service=object(), template_path=tmp_path / "模板.xlsx",
            input_dir=tmp_path, output_dir=tmp_path / "输出", period="2026-09",
            history_path=tmp_path / "历史.xlsx",
        )


def test_all_default_business_workflows_have_explicit_valid_inputs():
    from base_audit.workflow.defaults import default_workflows
    from base_audit.workflow.runner import build_registry

    registry = build_registry()
    for workflow_id, definition in default_workflows().items():
        errors = validate_workflow(definition, registry, ("com",))
        assert not errors, "{}: {}".format(workflow_id, errors)
