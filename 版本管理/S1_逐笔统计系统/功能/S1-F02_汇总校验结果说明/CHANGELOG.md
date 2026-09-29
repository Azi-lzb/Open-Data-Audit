# S1-F02 汇总校验结果说明

当前版本：v1.1.1

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
