# S1-F02 汇总校验结果说明

当前版本：v1.2.0

## v1.2.0 — 2026-10-19 21:43

变更类型：Feature / MINOR

前版本：v1.1.2
新版本：v1.2.0

修改内容：
- 新增可选“兜底模板”工作表；只有模板包含该表且配置一个“任意行汇总区域”和一个“表头区域”时，才会把未匹配的源工作表按该区域及表头汇总；
- 未匹配工作表没有有效数据行时不生成结果工作表；
- 模板不含“兜底模板”时维持原行为，未匹配的源工作表仍忽略；
- 保留普通命名区域的表头结构检查，以及缺少模板目标工作表时按表继续汇总的行为；
- 更新 S1-F02 流程说明、模板命名区域提示和测试登记。

影响文件：
- core/src/base_audit/summary_fallback.py
- core/src/base_audit/native/summary.py
- core/src/base_audit/region_summary.py
- core/src/base_audit/preflight_xlsx.py
- core/src/base_audit/workflow/adapters/audit_handlers.py
- core/src/base_audit/workflow/adapters/native_audit_handlers.py
- core/src/base_audit/workflow/defaults.py
- core/src/base_audit/name_config.py
- core/src/base_audit/node_flow_config.py
- core/frontend/web/index.html
- core/tests/test_native_uos.py
- core/tests/test_preflight_xlsx.py
- core/tests/test_version_manifest.py
- core/tests/TEST_GROUPS.md
- 版本管理/VERSION_MANIFEST.json

配置影响：
- 无；S1-CFG-01 不升级。需要使用兜底时，在现有模板中添加“兜底模板”工作表并设置所需命名区域。

测试：
- 定向测试：12 passed；
- `python -m pytest -q -rs -m s1 core/tests`：477 passed、13 skipped、463 deselected；
- `python -m compileall -q core/src/base_audit` 与 `git diff --check` 通过；
- UOS/麒麟与 Windows Excel COM 真机路径未在本轮执行。

GPT修改记录：
- GPT修改记录/20261019_214318_S1F02兜底模板汇总.txt

Git提交：
- 7ebaa9e24810cad4d3a640a092281cbe505df88b

## v1.1.2 — 2026-10-19 18:06

变更类型：Bugfix / PATCH

前版本：v1.1.1
新版本：v1.1.2

修改内容：
- 源文件缺少模板中的一张或多张目标工作表时，只跳过对应工作表，继续汇总该文件中其他结构匹配的工作表；即使所有目标工作表都缺失，也保留带表头的汇总结果；
- 缺少命名区域字段来源工作表时，该字段留空并给出提示，不再使整批汇总中断；
- 保留现有工作表结构比对，已存在工作表表头不匹配的源文件仍跳过，避免列错位数据进入汇总；
- S1-F02 不再因未实际使用的可选外部工作簿检查失败而中止。

影响文件：
- core/src/base_audit/preflight_xlsx.py
- core/src/base_audit/region_summary.py
- core/src/base_audit/native/summary.py
- core/src/base_audit/workflow/defaults.py
- core/src/base_audit/workflow/adapters/audit_handlers.py
- core/src/base_audit/workflow/adapters/native_audit_handlers.py
- core/src/base_audit/workflow/adapters/simple_handlers.py
- core/src/base_audit/name_config.py
- core/frontend/web/index.html
- core/tests/test_preflight_xlsx.py
- core/tests/test_native_uos.py
- core/tests/TEST_GROUPS.md

配置影响：
- 无；S1-CFG-01 不升级。仅更新默认流程说明，不改变用户配置工作簿结构。

测试：
- 定向汇总测试：5 passed；
- `python -m pytest -q -rs -m s1 core/tests`：475 passed、13 skipped、463 deselected；跳过项为 Windows/Linux 平台限制或 Flask 依赖条件；
- UOS/麒麟真机及 Office/LibreOffice 实际运行未在本轮执行。

GPT修改记录：
- GPT修改记录/20261019_180630_S1F02缺少工作表继续汇总.txt

Git提交：
- 未提交

## v1.1.1 — 2026-09-28 11:11

变更类型：Bugfix / PATCH

前版本：v1.1.0
新版本：v1.1.1

修改内容：
- UOS 原生区域汇总按实际写入的数据行数更新结果提示，不再固定显示 0 行；
- 运行明细及导出的“任意行汇总”日志逐工作表记录数据行数，与 Windows COM 日志列一致。

影响文件：
- core/src/base_audit/native/summary.py
- core/src/base_audit/service.py
- core/src/base_audit/workflow/adapters/simple_handlers.py

配置影响：
- 无；S1-CFG-01 不升级。

测试：
- Python 编译检查及 git diff --check 通过；本轮未运行 pytest 或 UOS 真机复验。
- UOS 验收包重新生成，打包器 ZIP CRC 与 SHA-256 检查通过。

GPT修改记录：
- GPT修改记录/20260928_111100_UOS_S1F02汇总日志行数修复.txt

Git提交：
- 未提交

## v1.1.0 — 2026-09-23 13:19

变更类型：Feature / MINOR

前版本：v1.0.0
新版本：v1.1.0

修改内容：
- 运行日志 Excel 导出改由独立设置控制，页面明细显示不再决定是否导出；
- 旧 `write_flow_logs` 用户设置保持兼容；业务汇总结果不变。

影响文件：
- core/src/base_audit/web_app.py
- core/frontend/web/index.html

配置影响：
- 用户设置 JSON 增加独立日志显示字段；业务配置工作簿未修改。

测试：
- 日志体验定向测试：112 passed；compileall、前端 JavaScript 语法检查、diff 检查通过。core 剩余批次结果及 S2 配置筛选范围断言、COM/native_uos 环境限制详见 GPT 修改记录。

GPT修改记录：
- GPT修改记录/20260923_131900_日志展示与用户配置体验优化.txt

Git提交：
- 未提交

## v1.0.0 — 2026-09-19

变更类型：Baseline

说明：

V3 版本管理体系建立时的正式基线。此版本建立前的开发历史未逐项倒推
SemVer；可通过 `GPT修改记录/` 与 Git 历史继续追溯。

组件说明：
- 逐笔统计系统：汇总校验结果并生成校验结果说明（含区域汇总与历史状态比对）。

当前实现：
- `core/src/base_audit/region_summary.py`

相关配置：
- S1-CFG-01（版本见各自 CHANGELOG）

测试基线：
- 以建立版本管理体系当日 core 全量 pytest 通过状态为基线（详见当次
  GPT修改记录）。
