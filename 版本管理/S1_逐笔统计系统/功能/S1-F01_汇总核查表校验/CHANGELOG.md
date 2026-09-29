# S1-F01 汇总核查表校验

当前版本：v1.2.0

## v1.2.0 — 2026-09-23 13:19

变更类型：Feature / MINOR

前版本：v1.1.1
新版本：v1.2.0

修改内容：
- 设置中心将页面运行明细与运行日志 Excel 导出拆为独立开关；
- 性能与运行诊断展示使用既有耗时及实际执行方式记录，不改变日志采集和业务执行；
- 旧用户设置 `write_flow_logs` 迁移为运行明细和 Excel 导出的兼容值；
- 相同输入下，日志开关不改变审核结果、结果工作簿数据或审核副本内容。

影响文件：
- core/src/base_audit/settings.py
- core/src/base_audit/web_app.py
- core/frontend/web/index.html

配置影响：
- 用户设置 JSON 增加独立日志显示字段；旧字段仍兼容；业务配置工作簿未修改。

测试：
- 日志体验定向测试：112 passed；compileall、前端 JavaScript 语法检查、diff 检查通过。core 剩余批次结果及 S2 配置筛选范围断言、COM/native_uos 环境限制详见 GPT 修改记录。

GPT修改记录：
- GPT修改记录/20260923_131900_日志展示与用户配置体验优化.txt

Git提交：
- 未提交

## v1.1.1 — 2026-09-21

变更类型：Bugfix / PATCH

前版本：v1.1.0
新版本：v1.1.1

修改内容：
- 条件格式多区域 `sqref` 的锚点不再依赖 XML 区域序列化顺序，Python 路径逐格
  判定与 Office 渲染基线一致；
- COM 求值按“规则 + 目标格 + 平移表达式”求值，避免一格命中后向整条规则区域
  广播；
- 条件格式规则说明改为当前命中格的相对引用口径；
- 数值比较按 Excel 15 位有效数字语义收敛，消除 D35/E35 这类二进制浮点尾差误报。

影响文件：
- core/src/base_audit/cf_reader.py
- core/src/base_audit/conditional_evaluators.py
- core/src/base_audit/conditional_engine.py
- core/src/base_audit/conditional_format.py
- core/tests/test_conditional_label_shift.py
- core/tests/test_conditional_excel_float.py

配置影响：
- S1-CFG-01 v1.0.0 → v1.0.0（未改）

测试：
- 条件格式定向测试 48 passed；真实报送样本逐格对拍通过；
- core 全量：678 passed、2 skipped、3 failed（均为本轮外的黄金基线/遗留目录问题）。

GPT修改记录：
- GPT修改记录/20260921_221958_S1F01条件格式修复版本登记与提交前校验.txt

Git提交：
- 未提交

## v1.1.0 — 2026-09-19

变更类型：Feature / MINOR（用户需求：检查模板公式 → 检查模板）

前版本：v1.0.1
新版本：v1.1.0

修改内容：
- 按钮与报告更名「检查模板」；
- 新增命名区域盘点（只提示、不阻断）：
  - 三族业务区域按族计数并列出 名称 = range（作用域：工作簿/工作表）：
    校验区域（含旧 VBA 变体 本地校验区域1 / 联表校验区域#1，名称含
    “校验区域”即归族）、表结构区域、条件格式区域；
  - Excel 内部名称（_FilterDatabase* / Print_Area / Print_Titles /
    _xlnm.*）单独忽略清单，不报错误；
  - 其余自定义名称列入「其他」供确认是否遗留；
  - 缺失业务命名区域不阻断——必需性校验仍由执行流程的模板体检完成；
- 命名区域中的错误值引用仍计入明确错误；内部名称的错误残留不误报。

影响文件：
- core/src/base_audit/formula_inspection.py
- core/src/base_audit/web_app.py
- core/frontend/web/index.html

配置影响：
- 无

测试：
- test_formula_inspection 重写 + 新增盘点测试（三族分类/旧变体归族/
  内部名忽略/其他列出）；真实 2026-07-31 模板实测（表结构区域 8 个、
  条件格式区域 4 个逐条列出）；core 全量通过

Git提交：
- 见仓库日志

## v1.0.1 — 2026-09-19

变更类型：Config / PATCH（用户要求的检查报告输出简化）

前版本：v1.0.0
新版本：v1.0.1

修改内容：
- 「检查模板公式」只报明确错误（公式文本/命名区域中的 #REF!、#NAME?、
  #N/A、#VALUE! 等），不再逐条列出「易失/动态引用」（INDIRECT/OFFSET
  等）与「查询引用」（INDEX/MATCH/VLOOKUP 等）需复核清单——公式能否
  计算一律以 Office 计算引擎真实重算为准
- 修复字符串字面量误报：公式常量里出现 "#N/A 文本" 等字样不再计为
  错误（剥离双引号字符串后识别错误 token）
- 同步：web_app 汇总文案、前端检查报告与按钮提示

影响文件：
- core/src/base_audit/formula_inspection.py
- core/src/base_audit/web_app.py
- core/frontend/web/index.html

配置影响：
- 无

测试：
- tests/test_formula_inspection.py 重写（INDIRECT/INDEX 不再产出警告、
  字符串字面量不误报、坏名称与 #REF! 仍报）；core 全量通过

Git提交：
- 见仓库日志

## v1.0.0 — 2026-09-19

变更类型：Baseline

说明：

V3 版本管理体系建立时的正式基线。此版本建立前的开发历史未逐项倒推
SemVer；可通过 `GPT修改记录/` 与 Git 历史继续追溯。

组件说明：
- 逐笔统计系统主流程：模板体检、表结构比对、公式复制、Excel/WPS 重算、问题提取与审核副本输出（DAG 编排）。

当前实现：
- `core/src/base_audit/service.py`

相关配置：
- S1-CFG-01（版本见各自 CHANGELOG）

测试基线：
- 以建立版本管理体系当日 core 全量 pytest 通过状态为基线（详见当次
  GPT修改记录）。
