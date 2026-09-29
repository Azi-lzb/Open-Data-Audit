# S3-F02 比较文件到说明文件

当前版本：v1.0.1

## v1.0.1 — 2026-09-19

变更类型：Refactor / PATCH

前版本：v1.0.0
新版本：v1.0.1

修改内容：
- 源码路径归位：core/src/base_audit/central_statistics → core/src/base_audit/systems/s3_central_statistics
- 实现登记路径更新

业务逻辑：
- 无变化（纯路径/命名重构，算法与配置口径未动）

配置影响：
- 无

测试：
- core 全量 pytest 通过；三系统烟测通过

GPT修改记录：
- GPT修改记录/20260919_110000_源码按S1S2S3归位并清理版本化命名.txt

Git提交：
- 见仓库日志（refactor: align source layout with S1 S2 S3 systems）

## v1.0.0 — 2026-09-19

变更类型：Baseline

说明：

V3 版本管理体系建立时的正式基线。此版本建立前的开发历史未逐项倒推
SemVer；可通过 `GPT修改记录/` 与 Git 历史继续追溯。

组件说明：
- 大集中统计系统：按“输出说明”标记导出简化说明文件。

当前实现：
- `core/src/base_audit/central_statistics/explanation_exporter.py`

相关配置：
- S3-CFG-00、S3-CFG-01、S3-CFG-02、S3-CFG-03（版本见各自 CHANGELOG）

测试基线：
- 以建立版本管理体系当日 core 全量 pytest 通过状态为基线（详见当次
  GPT修改记录）。
