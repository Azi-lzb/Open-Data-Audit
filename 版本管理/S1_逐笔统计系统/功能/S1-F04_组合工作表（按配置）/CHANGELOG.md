# S1-F04 组合工作表（按配置）

当前版本：v1.1.0

## v1.1.0 — 2026-09-23 13:19

变更类型：Feature / MINOR

前版本：v1.0.0
新版本：v1.1.0

修改内容：
- 运行日志 Excel 导出改由独立设置控制，页面明细显示不再决定是否导出；
- 旧 `write_flow_logs` 用户设置保持兼容；组合工作簿业务内容不变。

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
- 逐笔统计系统：按流程配置执行组合工作表 DAG 流程。

当前实现：
- `core/src/base_audit/service.py`

相关配置：
- S1-CFG-01（版本见各自 CHANGELOG）

测试基线：
- 以建立版本管理体系当日 core 全量 pytest 通过状态为基线（详见当次
  GPT修改记录）。
