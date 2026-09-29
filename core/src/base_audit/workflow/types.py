"""DAG 工作流引擎的基础类型与枚举（阶段 1）。

职责边界见项目《逐笔统计系统——DAG 二次架构实现指南》：本包只提供
编排层的显式模型，Excel/WPS COM、UOS native 等业务能力仍由
``base_audit.excel_com`` / ``base_audit.native`` 提供，经 adapters 包装。
"""

from __future__ import annotations

from enum import Enum

SCHEMA_VERSION = "2"


class ModuleCategory(str, Enum):
    """模块类别：数据来源、检查、修改、汇聚、出口与控制。"""

    SOURCE = "SOURCE"
    CHECK = "CHECK"
    MODIFY = "MODIFY"
    AGGREGATE = "AGGREGATE"
    SINK = "SINK"
    CONTROL = "CONTROL"


class ExecutionScope(str, Enum):
    """执行粒度：每次运行、每批、每文件、归并和出口。"""

    PER_RUN = "PER_RUN"
    PER_BATCH = "PER_BATCH"
    PER_FILE = "PER_FILE"
    REDUCE = "REDUCE"
    SINK = "SINK"


class ArtifactType(str, Enum):
    """基础 Artifact 类型。"""

    FILE = "FILE"
    FILE_SET = "FILE_SET"
    METADATA = "METADATA"
    AUDIT_RESULT = "AUDIT_RESULT"
    AUDIT_RESULT_SET = "AUDIT_RESULT_SET"
    STRUCTURE_MATCH_RESULT = "STRUCTURE_MATCH_RESULT"
    OPERATION_RESULT = "OPERATION_RESULT"
    RESULT_WORKBOOK = "RESULT_WORKBOOK"
    LOG_EVENT_SET = "LOG_EVENT_SET"


class ArtifactSubtype(str, Enum):
    """Excel 业务子类型；归属某个基础类型（见 SUBTYPE_PARENT）。"""

    EXCEL_SOURCE = "EXCEL_SOURCE"
    EXCEL_SOURCE_SET = "EXCEL_SOURCE_SET"
    EXCEL_WORKBOOK = "EXCEL_WORKBOOK"
    EXCEL_WORKBOOK_SET = "EXCEL_WORKBOOK_SET"
    EXCEL_TEMPLATE = "EXCEL_TEMPLATE"
    EXTERNAL_WORKBOOK = "EXTERNAL_WORKBOOK"
    FINAL_AUDIT_WORKBOOK_SET = "FINAL_AUDIT_WORKBOOK_SET"
    RESULT_WORKBOOK = "RESULT_WORKBOOK"
    NAMED_RANGE_SET = "NAMED_RANGE_SET"


SUBTYPE_PARENT = {
    ArtifactSubtype.EXCEL_SOURCE: ArtifactType.FILE,
    ArtifactSubtype.EXCEL_SOURCE_SET: ArtifactType.FILE_SET,
    ArtifactSubtype.EXCEL_WORKBOOK: ArtifactType.FILE,
    ArtifactSubtype.EXCEL_WORKBOOK_SET: ArtifactType.FILE_SET,
    ArtifactSubtype.EXCEL_TEMPLATE: ArtifactType.FILE,
    ArtifactSubtype.EXTERNAL_WORKBOOK: ArtifactType.FILE,
    ArtifactSubtype.FINAL_AUDIT_WORKBOOK_SET: ArtifactType.FILE_SET,
    ArtifactSubtype.RESULT_WORKBOOK: ArtifactType.RESULT_WORKBOOK,
    ArtifactSubtype.NAMED_RANGE_SET: ArtifactType.METADATA,
}


class Cardinality(str, Enum):
    ONE = "ONE"
    OPTIONAL_ONE = "OPTIONAL_ONE"
    MANY = "MANY"


class ArtifactStatus(str, Enum):
    CREATED = "CREATED"
    VALID = "VALID"
    PARTIAL = "PARTIAL"
    FAILED = "FAILED"
    SKIPPED = "SKIPPED"
    DELETED = "DELETED"


class Retention(str, Enum):
    TEMPORARY = "TEMPORARY"
    RUN = "RUN"
    PERSISTENT = "PERSISTENT"
    EXTERNAL = "EXTERNAL"


class FailurePolicy(str, Enum):
    """节点失败策略；旧配置“停止/跳过”分别映射为前两种。"""

    FAIL_WORKFLOW = "FAIL_WORKFLOW"
    SKIP_NODE = "SKIP_NODE"
    SKIP_ITEM = "SKIP_ITEM"
    CONTINUE_WITH_PARTIAL = "CONTINUE_WITH_PARTIAL"


class NodeRunStatus(str, Enum):
    PENDING = "PENDING"
    RUNNING = "RUNNING"
    SUCCEEDED = "SUCCEEDED"
    PARTIAL = "PARTIAL"
    FAILED = "FAILED"
    SKIPPED = "SKIPPED"
    DISABLED = "DISABLED"


OLD_FAILURE_POLICY_MAP = {
    "停止": FailurePolicy.FAIL_WORKFLOW,
    "跳过": FailurePolicy.SKIP_ITEM,
}
