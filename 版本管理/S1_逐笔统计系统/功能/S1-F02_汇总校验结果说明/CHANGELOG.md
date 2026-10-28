# S1-F02 汇总校验结果说明

当前版本：v1.3.1

## v1.3.1 — 2026-10-27 20:51

变更类型：Bugfix / PATCH

前版本：v1.3.0；新版本：v1.3.1。

- 待处理文件说明将“新增工作表”改为“模板外工作表”，避免把源文件已有但未纳入模板的表误称为本次新增；统一使用“所选模板、同名工作表、模板外工作表”表述。
- 界面不再显示 F1/F2 标签或兜底模板提示；模板外工作表只作为匹配提示，匹配情况以同名工作表为准。
- 提示区分普通隐藏与 `veryHidden` 工作表，并在悬浮详情说明普通取消隐藏列表为何看不到 `veryHidden` 表及其显示方法。
- 仅调整 F2 预览信息；F1 自动模板识别、文件选择、结构检查和汇总逻辑不变。

影响文件：`core/src/base_audit/discovery.py`、`core/frontend/web/index.html`、相关 S1 测试。

配置影响：无；S1-CFG-01 v2.0.0 不变。

测试：`test_discovery.py`、`test_web_app.py`、版本清单测试及 S1 回归结果见本轮 GPT 修改记录。

GPT修改记录：`GPT修改记录/20261027_205145_S1-F02模板外隐藏工作表提示.txt`。

Git提交：未提交。

## v1.3.0 — 2026-10-22 18:39

变更类型：Feature / MINOR

前版本：v1.2.2；新版本：v1.3.0。

- 纯 Python 汇总管线的源数据值读取新增 AUTO / Calamine / OpenPyXL 路由；AUTO 优先 Calamine，兼容性失败时说明原因并回退，手动模式严格执行。
- 新建汇总结果工作簿增加独立 AUTO / Rust xlsxwriter / OpenPyXL 路由；现有模板命名区域读取、结构检查和 Office 公式缓存刷新链路不变。
- 源数据位置、日期、文字、布尔值、空值、工作表及汇总结果与 OpenPyXL 对拍一致；运行记录显示实际引擎、耗时和自动回退原因。

影响文件：`core/src/base_audit/native/summary.py`、`summary_source_reader.py`、`s1_summary_writer.py`、S1-F02 设置与测试，以及 `core/rust/s1_summary_writer/`。

配置影响：用户设置 JSON 增加 `summary_source_reader` 和 `summary_result_writer`；业务配置工作簿与 S1-CFG-01 v2.0.0 不变。

性能基准：Windows 合成源数据为 3 个工作簿、9,000 行、12 列，Calamine/OpenPyXL 读取中位数 0.1661/1.1905 秒（7.17 倍）。结果表写入 1,000/5,000/20,000 行时，Rust/OpenPyXL 分别为 0.0444/0.1917、0.1439/0.8677、0.5256/3.5827 秒；每个规模均逐格对拍。

测试：S1 系统回归与 Windows/WSL core 全量回归结果见本轮 GPT 修改记录。

GPT修改记录：`GPT修改记录/20261022_182600_S1读写引擎路由与Rust结果写入.txt`。

Git提交：未提交。

## v1.2.2 — 2026-10-21 23:54

变更类型：Bugfix / PATCH

前版本：v1.2.1；新版本：v1.2.2。

- 模板中存在、源文件暂未出现的工作表不再显示为部分匹配。
- 源文件新增工作表改用中性提示；模板含“兜底模板”时提示可尝试兜底，否则明确说明新增表不会参与汇总。
- F1 模板识别仍与 F2 预览分开；实际汇总范围和结构检查逻辑不变。

影响文件：`core/src/base_audit/discovery.py`、`core/frontend/web/index.html`、相关 S1 测试。

配置影响：无；S1-CFG-01 v2.0.0 未变。

测试：S1 回归 501 passed / 13 skipped / 0 failed，3 subtests passed；最终模板表排除调整后 `test_discovery.py` 20 passed。

GPT修改记录：`GPT修改记录/20261021_235428_S1S5待处理文件匹配状态优化.txt`。

Git提交：未提交。

## v1.2.1 — 2026-10-21 11:02

变更类型：Bugfix / PATCH
前版本：v1.2.0；新版本：v1.2.1。

- 待处理文件列表增加相对当前模板的 F2 工作表名匹配提示；有部分同名表时显示“部分匹配（F2）”及匹配数，可以按匹配表汇总。
- F1 原分类结果仍独立显示，不放宽分类分数、自动模板推荐或实际表结构检查。
- 提示中明确表头仍需执行时核对，额外表仅在配置“兜底模板”时尝试汇总；未修改汇总算法、配置及原始文件。

影响文件：core/src/base_audit/discovery.py、web_app.py、core/frontend/web/index.html 与对应定向测试。
配置影响：无；S1-CFG-01 不升级。
测试及 GPT修改记录：见 GPT修改记录/20261021_110257_S1F2部分工作表匹配显示.txt。
Git提交：未提交。

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
