"""base_audit.workflow：DAG 二次架构的编排层。

分层约束（见项目实现指南）：
- types/ports/artifacts/modules/graph/registry/validation：显式模型与校验；
- context/artifact_store/scheduler/persistence：运行时；
- adapters/：把现有 Excel/native 能力包装成 Handler（不复制核心实现）；
- defaults.py：两条主流程等固定 DAG 的默认定义。
本包不判断操作系统或 COM/UNO，平台差异只在 engines 与 adapters。
"""
