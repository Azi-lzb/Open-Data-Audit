# S3-F03 机构说明到说明文件

当前版本：v1.1.0

## v1.1.0 — 2026-10-13 21:54

变更类型：Feature / MINOR

前版本：v1.0.1
新版本：v1.1.0

修改内容：
- 本条为汇总追溯 `40875d7..1b0ae87` 的已完成调整，并按当前确认口径登记。
- 反馈导入按机构代码、指标代码、数据值、增减额四项匹配；代码作文本规范化，金额按四位小数 `ROUND_HALF_UP` 比较，机构名称和指标名称差异不再阻断匹配。
- 导出的村镇银行及多地区机构长代码可作为机构代码匹配；重复目标或反馈原因冲突时报告冲突，不静默覆盖。

影响文件：
- core/src/base_audit/systems/s3_central_statistics/explanation_exporter.py
- core/tests/test_central_explanation_exporter.py

配置影响：业务配置未修改；相关配置版本为 S3-CFG-00 v2.0.2、S3-CFG-01 v3.0.3、S3-CFG-02 v1.0.3、S3-CFG-03 v2.0.0。

测试与限制：
- 历史验证见 GPT修改记录/20261007_190346_机构说明导入对齐VBA宽松匹配与村镇银行长码.txt（该记录载明说明导入测试 7 passed、S3 组 433 passed）；结果为既有记录，本次未重跑。
- 本次未重打包，未做 Office 真机验收。

GPT修改记录：
- GPT修改记录/20261013_215438_统一3.3写入口径并补齐发布组件版本.txt

Git提交：
- 未提交。

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
- 大集中统计系统：批量导入机构反馈说明文件并回写变动原因。

当前实现：
- `core/src/base_audit/central_statistics/explanation_exporter.py`

相关配置：
- S3-CFG-00、S3-CFG-01、S3-CFG-02、S3-CFG-03（版本见各自 CHANGELOG）

测试基线：
- 以建立版本管理体系当日 core 全量 pytest 通过状态为基线（详见当次
  GPT修改记录）。
