# S1-F01 汇总核查表校验

当前版本：v2.7.4

## v2.7.4 — 2026-11-07 16:40

变更类型：Bugfix / PATCH

前版本：v2.7.3；新版本：v2.7.4。

- 修复审核副本“悬空外部簿引用”缺陷：模板校验区域公式若携带 `[N]表名!` 外部簿前缀
  （模板作者用外部链接写规则，模板自带 `xl/externalLinks/*`），而副本把外部簿导入为
  本地 sheet 后不含 externalLink 部件，公式便指向不存在的外部簿——Excel COM
  `Workbooks.Open` 直接拒绝（HRESULT 0x80020009），WPS/LibreOffice 可打开但重算
  丢失公式或回落断链缓存值。
- 修复：公式写入后、Office 打开/重算前，剥离指向“本地已导入同名工作表”的 `[N]` 前缀
  （表名存在于本包 `<sheets>` 才剥；指向真外部簿的不动）。覆盖 `<f>`、条件格式/数据验证
  `<formula>`、图表 `<c:f>`、`<definedNames>` 四类载体，兼容 `'表名'!` 与裸 `表名!`
  变体，并兼容表名的字面 UTF-8 与 NCR 实体（`&#38598;`）两种编码——openpyxl 保存
  会写成 NCR，不兼容会漏修（本次实测踩到）。
- 接线：Windows `handle_formula_copy`（DIRECT_OOXML 写入后、COM 打开前）、
  UOS native `handle_formula_copy`（openpyxl 写公式后、LibreOffice 重算前）。
- 拒绝“保留外链定义”方案：保留虽可让 Excel 打开，但重算回落模板缓存的外部旧值，
  审核结论基于错误数据；剥离后公式指向本次导入的本地表，与旧 audit 结果一致。
  真·无关外链剥离后重算为 `#REF!`，按审核异常口径捕获，优于整文件拒绝打开。
- 机制文档与跨仓库结论：`docs/悬空外部引用缺陷_机制与跨仓库结论.md`
  （含 openpyxl 不剥离外链的实测、旧管线“校验区域范围隔离”而非代码归一的结论）；
  规则补充：`rules/03_Excel与OOXML边界.md` 第 2 节。

影响文件：`core/src/base_audit/excel_open_repair.py`、
`core/src/base_audit/workflow/adapters/audit_handlers.py`、
`core/src/base_audit/workflow/adapters/native_audit_handlers.py`、
`core/tests/test_excel_open_repair.py`、`core/tests/test_version_manifest.py`、
`rules/03_Excel与OOXML边界.md`、`docs/悬空外部引用缺陷_机制与跨仓库结论.md`、
`版本管理/VERSION_MANIFEST.json`。

配置影响：无；S1-CFG-01 v2.0.0 不变。无新增用户设置。

测试：新增 5 项（剥本地 `[N]`、保留真外链、引号变体、definedNames 上下文、
真实毒公式 business_templates 债券业务）；`test_excel_open_repair.py` 12 passed；
相关回归 108 passed；S1 回归 594 passed / 13 skipped。
验收四条实测：①修复副本 Excel 只读打开成功（7/7 表）②WPS 重算后公式保留
（`=INDEX(集中系统数据!$1:$1048576,1)`）③原始报送文件 SHA-256 不变
④公式指向本地已导入表。因果实验：A 干净副本 OK → B 加悬空 `[1]` FAIL(0x80020009)
→ C 剥离后 OK。

GPT修改记录：`GPT修改记录/20261107_164000_S1F1悬空外部引用修复.txt`。

Git提交：未提交。

## v2.7.3 — 2026-11-07 12:20

变更类型：Bugfix / PATCH

前版本：v2.7.2；新版本：v2.7.3。

- 条件格式结果提取在 Excel/WPS COM 模式下改为读取审核副本（copy-files 产物，已做 Excel 兼容修复），
  原始报送文件不再经 COM 打开。此前该节点直接打开原始报送文件，v2.7.2 的副本修复覆盖不到它，
  28 类 Excel 拒绝打开的文件在条件格式节点仍会失败。
- 副本与原始文件的条件格式语义完全一致（v2.7.2 修复只规范化引号与窗格，不改规则、值与渲染）；
  COM_DISPLAY 的真实渲染判定在副本上照常成立。
- PYTHON 模式仍读原始报送文件（无 COM 打开问题，保持与 129 份对拍基线同口径）。
- 结构比对拒绝的文件没有审核副本，COM 提取回退打开原始文件（若 Excel 仍打不开则按节点失败策略处理）。
- DAG 新增边 copy-files.workbooks → extract-conditional-issues.workbooks；模块新增可选输入端口
  （required=False），UOS native 图同样带边但 native handler 忽略该输入。

影响文件：`core/src/base_audit/workflow/adapters/audit_handlers.py`、
`core/src/base_audit/workflow/defaults.py`、`core/tests/test_external_copy_fallback.py`、
`core/tests/test_version_manifest.py`、`版本管理/VERSION_MANIFEST.json`。

配置影响：无；S1-CFG-01 v2.0.0 不变。

测试：handler 门控 3 项新增（COM 模式开副本、端口可选、DAG 边存在）；端到端实测
（Windows Excel COM）：置换旧债样本经复制修复后 COM_DISPLAY 从副本提取成功；
同业业务/代发工资/惠州农商代发工资的“副本提取 vs 原始提取”逐格一致（9/9、7/7、16/16）。
S1 回归 581 passed / 13 skipped。

GPT修改记录：`GPT修改记录/20261107_122000_S1F1条件格式提取改读审核副本.txt`。

Git提交：未提交。

## v2.7.2 — 2026-11-07 11:40

变更类型：Bugfix / PATCH

前版本：v2.7.1；新版本：v2.7.2。

- Windows Excel/自动引擎在“复制审核副本”节点对副本自动修复两类 Excel 拒绝打开的非规范内容：
  1. 公式引用含全角括号的表名未加单引号（如 `个人贷款发生额信息（隐藏发放）!A1`）→ 补引号，
     公式语义不变；载体含单元格公式、条件格式/数据验证公式、图表引用与定义名称。
  2. 冻结窗格 topLeftCell 行号落在冻结区内（如 `ySplit="6"` 配 `topLeftCell="C6"`）→
     行号修正为 ySplit+1。
- 门控：仅引擎为“Microsoft Excel/自动”时修复；“WPS 表格”引擎与 UOS native 管线不修复
  （WPS 本身能打开这些文件，修复只属于 Excel 路径）。修复只作用于审核副本，
  原始报送文件不改写；无缺陷文件字节不变；修复失败不打断流程（副本保持原样并按原路径继续）。
- 根因与证据：129 份真实核查表中 28 本 Excel COM 打不开（B00 三类 21 本 + 置换旧债 7 本），
  根因分析与属性级二分见 `test-output/s1f1_mode_parity_20261107/excel_open_root_cause.md`；
  两类修复应用后 28/28 可被 Excel 打开（excel_open_fix_28.json），修复模块与根因分析脚本
  在 28 本上计数逐本一致。
- 运行记录：复制节点输出 excel_open_repair_files/quoted/panes 指标，逐文件写日志说明。

影响文件：`core/src/base_audit/excel_open_repair.py`（新增）、
`core/src/base_audit/workflow/adapters/audit_handlers.py`、
`core/tests/test_excel_open_repair.py`（新增）、`core/tests/conftest.py`、
`core/tests/test_version_manifest.py`、`版本管理/VERSION_MANIFEST.json`。

配置影响：无；S1-CFG-01 v2.0.0 不变。无新增用户设置。

测试：新增 `test_excel_open_repair.py` 7 项（合成缺陷修复、干净文件字节不变、非 ZIP 跳过、
Excel/自动/WPS 引擎门控、原始文件保护、真实样本计数）；定向 25 passed（excel_open_repair +
version_manifest）；条件格式相关 59 passed；S1 回归 577 passed / 13 skipped
（1 项 distribution_profile 为既有偶发，单独与整文件复跑均通过）。

GPT修改记录：`GPT修改记录/20261107_114000_S1F1审核副本Excel兼容修复.txt`。

Git提交：未提交。

## v2.7.1 — 2026-11-07 10:55

变更类型：Bugfix / PATCH

前版本：v2.7.0；新版本：v2.7.1。

- 条件格式表达式求值对齐 Excel/WPS 真实口径，修正四类偏差（依据 129 份真实核查表、双引擎 COM 对拍）：
  1. 算术中空白单元格按 0 参与；除以 0（含 空/空）为 #DIV/0! 错误，错误在算术与比较中传播，
     含错误的规则整体不成立。原先 `空/空<=0.9` 被 Python 误判为触发（Excel/WPS 实测判不触发）。
  2. 空单元格参与比较按对方类型折算：对文本折空串、对数值折 0、对布尔折 FALSE。
  3. 数字与文本混合比较按 Excel 类型序：= 恒 FALSE、<> 恒 TRUE、数字恒小于文本；
     字面量 "" 是文本，`0<>""` 为 TRUE（此前会被折算成 0<>0 而误判）。
  4. t="e" 错误缓存值（#DIV/0! 等）参与比较 → 不成立（单位存款余额 F 列 5 格场景）。
- 对拍证据：对齐前 Python 相对真实 Office 有 92 处空白假阳性（66 vs Excel 可比集 + 26 vs WPS 可比集）；
  修正后 vs COM_EVALUATE(Excel) 101 文件、vs COM_DISPLAY(Excel) 101 文件、vs COM_EVALUATE(WPS) 129 文件
  全部 0 差异，且未误删任何正确命中（miss_to_hit=0）。逐文件数据见
  `test-output/s1f1_mode_parity_20261107/`（parity_report.md、checkpoint.jsonl、
  landed_semantics_verification.json、probe_divzero_com.py）。
- 语义边界保持不变：cellIs 空白目标格仍按不触发（当前语料 0 条 cellIs）；文本参与算术沿用现行为；
  文本不参与大小排序。Excel/WPS COM 对除零的实测返回（VARIANT 错误码 -2146826281 = #DIV/0!）
  见 probe_divzero_com.py，两家引擎逐项一致。

影响文件：`core/src/base_audit/conditional_format.py`、`core/tests/test_conditional_excel_float.py`、
`core/tests/test_version_manifest.py`、`版本管理/VERSION_MANIFEST.json`。

配置影响：无；S1-CFG-01 v2.0.0 不变。UOS native 与 Windows COM 回退路径共用本求值器，两侧口径同时受益。

测试：新增 TestExcelBlankAndErrorSemantics 14 项（定向 25 passed）；条件格式相关回归 59 passed；
S1 系统回归 571 passed / 13 skipped / 3 subtests passed；129 份金标准回归 0 差异
（`test-output/s1f1_mode_parity_20261107/landed_semantics_verification.json`）。

GPT修改记录：`GPT修改记录/20261107_105549_S1F1条件格式求值对齐Excel空白与除零语义.txt`。

Git提交：未提交。

## v2.7.0 — 2026-11-07 01:50

变更类型：Feature / MINOR

前版本：v2.6.5；新版本：v2.7.0。

- 高级设置新增“空白目标单元格”选项，可选择是否把值为空、但条件格式规则成立或 Office 显示填充变化的目标格纳入 S1-F01 结果；默认纳入，关闭后会排除并在运行记录中显示数量。
- WPS 无法读取 `DisplayFormat` 显示色时，继续使用 Python 条件规则求值兜底，并在运行记录明确标记该文件结果不代表 WPS 实际显示；Python 判定、Office 显示色判定和兜底结果保持各自来源，不混合成显示色证据。
- 设置保存在用户设置中；旧设置文件缺少该字段时按默认“纳入”读取。

影响文件：`core/src/base_audit/settings.py`、`core/src/base_audit/web_app.py`、`core/src/base_audit/service.py`、`core/src/base_audit/excel_com.py`、`core/src/base_audit/conditional_evaluators.py`、`core/src/base_audit/workflow/adapters/audit_handlers.py`、`core/frontend/web/index.html`、`core/tests/test_conditional_evaluators.py`、`core/tests/test_conditional_format_fallback.py`、`core/tests/test_web_app.py`、`core/tests/test_version_manifest.py`。

配置影响：无；S1-CFG-01 v2.0.0 不变。用户设置新增兼容字段，旧设置按默认值加载。

测试：定向测试 21 passed；S1 系统回归 563 passed / 7 skipped / 597 deselected / 3 subtests passed；`git diff --check` 通过。未执行 Excel/WPS 真机新增设置的 GUI 点选验收。

GPT修改记录：`GPT修改记录/20261107_0150_S1F1空白条件格式目标设置与WPS回退说明.txt`。

Git提交：未提交。

## v2.6.5 — 2026-11-06 23:03

变更类型：Bugfix / PATCH

前版本：v2.6.4；新版本：v2.6.5。

- Direct OOXML 外部表写入器支持保留公式文本、公式属性和缓存值；只含公式但无其他不支持对象的外部工作表不再因公式而退回 openpyxl。
- 对仍需 openpyxl 降级的外部表，在保存后从原工作簿 OOXML 中补回合并区域非左上角且实际有值的单元格；字符串转为 inlineStr，避免恢复整张工作表 XML 或受 sharedStrings 索引影响。
- 回归覆盖合并区域内的文本/数值/公式、工作表关系与重开；广东博罗个人贷款样例的 B8/C8 保留，Excel COM 重算 N8/O8 为 19057.46/37883.29，公式文本不变，原始输入 SHA-256 未变化。

影响文件：`core/src/base_audit/external_sheet_writer.py`、`core/src/base_audit/native/openpyxl_workbook.py`、`core/tests/test_external_sheet_writer.py`、`core/tests/test_external_copy_fallback.py`、`core/tests/test_version_manifest.py`。

配置影响：无；S1-CFG-01 v2.0.0 不变。

测试：定向测试 9 passed；版本清单测试 8 passed；S1 系统回归 561 passed / 7 skipped / 597 deselected / 3 subtests passed。真实样例原始文件在当前 Excel COM 环境中无法直接打开，因此 COM 数值验证使用同一真实样例的既有 Python 审核副本临时副本，未写入真实输入文件；重算结果及保存缓存均为 N8=19057.46、O8=37883.29。

GPT修改记录：`GPT修改记录/20261106_231326_S1F1外部表公式与合并区域数据保留.txt`。

Git提交：c350567（S1-F01 外部表数据保留与 V26.2.2.5 发布）。

## v2.6.4 — 2026-11-06 18:13

变更类型：Bugfix / PATCH

前版本：v2.6.3；新版本：v2.6.4。

- 条件格式指标定位忽略 Excel 错误缓存，不再把 `#DIV/0!` 等错误文本拼入指标名和问题标识；仅修正字段文字，不改变条件格式触发和触发值。
- 公式异常提取对“错误|差值为数值”应用绝对 `1e-10` 容差，过滤计算缓存中的极小浮点残值；其他公式结果和小额字段不套用此容差。
- 加固 Linux/native 缓存矩阵读取：矩阵缺少目标坐标时单格回读，避免 `IndexError` 丢失该工作簿的有效发现；若单格读取也失败，明确报告工作表和单元格。
- 定向只读复核当前 Python native 提取器对先前列出的 13 本副本可提取 95/95 条；另增加矩阵越界回读保护。

影响文件：`core/src/base_audit/template.py`、`core/src/base_audit/indicator_resolver.py`、`core/src/base_audit/excel_com.py`、`core/src/base_audit/native/openpyxl_workbook.py`。

配置影响：无；S1-CFG-01 v2.0.0 不变。

测试：四个修改模块 `py_compile` 通过；`git diff --check` 通过。未运行 pytest；S1 完整回归未执行。

GPT修改记录：`GPT修改记录/20261106_181336_S1F1提取错误过滤与浮点容差.txt`。

Git提交：未提交。

## v2.6.3 — 2026-10-28

变更类型：Bugfix / PATCH

前版本：v2.6.2；新版本：v2.6.3。

- 修复 UOS/native 管线无法添加外部文件导致整条 S1-F01 审核链被中止的问题：外部表含合并单元格、公式、数据验证或图片时，`DirectOoxmlExternalSheetWriter.prepare()` 抛出的 `UnsupportedExternalSheetOperation` 原先在 `handle_external_copy` 中没有捕获，异常冒泡直接终止 DAG。
- 外部文件添加改为三级降级：Direct OOXML 快路径不可用时自动改用 openpyxl 添加（保留合并单元格、公式、样式），快路径 `write()` 失败时对该文件单独改用 openpyxl 重试；两者都失败时仅把该文件记为失败、审核副本保留、节点状态降级为 PARTIAL，链路继续。
- Windows/COM 适配器同步修正：原先 prepare 失败会在每个文件上重新抛出同一异常，导致按“跳过”策略时所有文件都被跳过；现在改为整批降级到 openpyxl，保留原有“停止/跳过”策略只处理真正写入失败。
- 外部文件失败原因写入文件级 `stage_errors` 与节点部分失败记录，不再静默吞掉；外部表缺失可能使引用它的公式显示 `#REF!`，因此必须可见。
- 外部文件添加方式默认值不变，仍为 `DIRECT_OOXML`；纯数据外部表继续走快路径。
- 修复 `native/openpyxl_workbook._copy_sheet` 跨簿复制样式越界：原先直接搬 `_style` 索引，目标工作簿未登记对应字体/对齐，带样式的真实外部表复制后读回报 `IndexError: list index out of range`。现改为按原型逐项登记样式并按来源去重缓存，同时补齐条件格式与数据验证复制。
- 用 `config-real/外部文件-大集中相关数据.xlsx`（参照表 211 行、机构参照表 221 行、含 VLOOKUP 公式）实测：UOS 与 Windows 两条降级路径均成功添加工作表，公式与样式保留，审核副本可正常打开。

影响文件：`core/src/base_audit/workflow/adapters/native_audit_handlers.py`、`core/src/base_audit/workflow/adapters/audit_handlers.py`、`core/src/base_audit/native/openpyxl_workbook.py`、`core/tests/test_external_copy_fallback.py`、`core/tests/conftest.py`。

配置影响：无；S1-CFG-01 v2.0.0 不变。

测试：新增定向 `test_external_copy_fallback.py`：4 passed（纯数据走快路径、复杂外部表降级 openpyxl 并保留合并与公式、纯数据表与公式表混合计划整批降级后两表都在、双写入器失败时 PARTIAL 且副本保留）；相关回归 `test_external_sheet_writer.py`、`test_external.py`、`test_native_uos.py`、`core/tests/workflow/`：84 passed；S1 回归：524 passed / 13 skipped / 698 deselected / 3 subtests passed。

GPT修改记录：`GPT修改记录/20261028_113754_S1-F01外部文件添加降级与链路续跑.txt`。

Git提交：未提交。

## v2.6.2 — 2026-10-27 17:38

变更类型：Bugfix / PATCH

前版本：v2.6.1；新版本：v2.6.2。

- S1-F01 自动模板识别与待处理文件匹配排除源文件中 `veryHidden` 工作表，避免模板未配置的内部隐藏表拉低匹配率。
- Excel 普通隐藏工作表仍参与匹配；手动选择模板后的结构校验仍逐项检查模板要求的工作表和结构区域。
- 自动计算引擎遇到 Excel 无法打开的工作簿时，在当前会话没有打开其他工作簿的情况下自动改用 WPS 重试；用户明确选择 Excel 时仍按所选引擎执行。
- S1-F02 的工作表提示未调整。

影响文件：`core/src/base_audit/discovery.py`、`core/src/base_audit/excel_com.py`、`core/src/base_audit/workflow/runner.py`、相关 S1 测试。

配置影响：无；S1-CFG-01 v2.0.0 不变。

测试：定向 `test_discovery.py`、`test_preflight_xlsx.py`、`test_native_uos.py`：69 passed；S1 回归：518 passed / 13 skipped / 1 failed / 696 deselected / 3 subtests passed。唯一失败是既有 `test_version_manifest.py` 对 S5-F01 仍断言 v3.0.0，而当前清单为 v3.0.1，与本次 S1 改动无关。对 7 份真实审核副本的临时拷贝实测 Excel 打开失败后自动切换 WPS；7 份均打开、重算并保存成功，原始文件未改动。

GPT修改记录：`GPT修改记录/20261027_173844_S1-F01模板匹配与WPS打开兼容.txt`。

Git提交：未提交。

## v2.6.1 — 2026-10-22 19:54

变更类型：Bugfix / PATCH

前版本：v2.6.0；新版本：v2.6.1。

- Linux 原生审核流程将公式写入后的 LibreOffice 重算改为每批最多 8 份副本共用一次引擎启动；批次输出全部生成后才回写，降低重复启动开销并避免失败时只覆盖该批部分文件。
- 公式写入与 LibreOffice 批次重算分别记录耗时，重算日志显示工作簿数量和批次数；Windows Excel/WPS COM 路径、审核公式及结果口径不变。

影响文件：`core/src/base_audit/engines/libreoffice.py`、`core/src/base_audit/workflow/adapters/native_audit_handlers.py`、S1 native workflow 测试、`core/tests/TEST_GROUPS.md`。

配置影响：无；S1-CFG-01 v2.0.0 不变。

性能与兼容性验证：WSL Ubuntu AMD64 / LibreOffice 24.2.7.2，3 份合成审核副本（每份 1,000 行、2,000 个 INDEX/MATCH 与 INDIRECT 公式），五轮逐文件重算中位数 8.79 秒，批量重算 3.68 秒（约 2.39 倍）；公式缓存相同，并在修改数据值和动态引用后重新计算为新结果。另以 10 份中文文件名副本验证生产批量大小 8，实际分成 8+2 两批，10 份缓存均正确。

测试：Windows core 全量 1,207 passed / 13 skipped / 0 failed / 8 warnings / 5 subtests；WSL Ubuntu AMD64 core 全量 1,199 passed / 21 skipped / 0 failed / 8 warnings / 5 subtests；S1 定向新增批量成功/失败保护用例通过。8 条 warning 为测试环境未安装 `python-calamine` 时既有的自动回退提示。未在 Windows Excel/WPS COM 环境或 UOS 真机进行手工重算验收。

GPT修改记录：`GPT修改记录/20261022_195422_S1-F01_批量LibreOffice重算优化.txt`。

Git提交：未提交。

## v2.6.0 — 2026-10-22 18:39

## v2.6.0 — 2026-10-22 18:39

变更类型：Feature / MINOR

前版本：v2.5.0；新版本：v2.6.0。

- 新建审核结果工作簿支持独立 AUTO / Rust xlsxwriter / OpenPyXL 写入方式；自动优先可用的 Rust writer，失败时安全回退，手动选择失败时明确报错。
- 保留审核副本外部文件链接、目标工作表和单元格定位、可见单元格文字、表头样式、筛选、冻结及列宽；不涉及公式复制或 Office 重算。
- 设置中心独立保存 S1-F01 写入方式，不与 S1-F02 汇总结果写入方式共用设置。

影响文件：`core/src/base_audit/native/result_writer.py`、`core/src/base_audit/native/s1_summary_writer.py`、S1 workflow 结果输出节点、设置中心、`core/rust/s1_summary_writer/` 与相关测试。

配置影响：用户设置 JSON 增加 `audit_result_writer`；现有配置工作簿及 S1-CFG-01 v2.0.0 不变。

性能基准：Windows 合成结果表含逐行审核副本链接，1,000 条问题 Rust/OpenPyXL 中位数 0.419/1.532 秒（3.66 倍），3,000 条为 1.228/7.980 秒（6.50 倍）；对拍表格值、文件链接及目标单元格定位。

测试：S1 系统回归与 Windows/WSL core 全量回归结果见本轮 GPT 修改记录。

GPT修改记录：`GPT修改记录/20261022_182600_S1读写引擎路由与Rust结果写入.txt`。

Git提交：未提交。

## v2.5.0 — 2026-10-22 00:27

变更类型：Feature / MINOR

前版本：v2.4.0；新版本：v2.5.0。

- S1 待处理文件列表新增 F1 模板匹配状态和报表模板筛选；筛选只控制显示，原有勾选保留，全选/取消仅作用于当前显示文件。
- S1 文件选择读取限定在 S1 待处理列表，避免与隐藏的 S5 文件列表共享选择框。

影响文件：`core/frontend/web/index.html`、`core/tests/test_distribution_profile.py`、`core/tests/TEST_GROUPS.md`。

配置影响：无；S1-CFG-01 v2.0.0 未变。

测试：`python -X utf8 -m pytest -q -rs -m s1 core/tests`：502 passed / 13 skipped / 0 failed / 682 deselected，3 subtests passed。跳过项为 Linux/UOS 专属探测与缺少 Flask 的外壳路由测试。前端筛选和选择行为 Node 定向测试通过。

GPT修改记录：`GPT修改记录/20261022_002708_S1筛选与S5检查文案.txt`。

Git提交：未提交。

## v2.4.0 — 2026-10-21 23:54

变更类型：Feature / MINOR

前版本：v2.3.0；新版本：v2.4.0。

- 模板匹配按源工作簿的工作表覆盖判断；模板额外保留其他表时，不再因此降低匹配状态。
- 自动推荐仍优先选择工作表集合最相近的模板；多个候选相近时继续要求人工确认。

影响文件：`core/src/base_audit/discovery.py`、相关 S1 测试。

配置影响：无；S1-CFG-01 v2.0.0 未变。

测试：S1 回归 501 passed / 13 skipped / 0 failed，3 subtests passed；最终模板表排除调整后 `test_discovery.py` 20 passed。

GPT修改记录：`GPT修改记录/20261021_235428_S1S5待处理文件匹配状态优化.txt`。

Git提交：未提交。

## v2.3.0 — 2026-10-21

变更类型：Feature / MINOR

前版本：v2.2.1；新版本：v2.3.0。

修改内容：
- 首次启动时，若发行配置中带有“外部文件-大集中相关数据”工作簿，默认填入“大集中系统数据”路径。
- 首次启动时优先把“模板文件_命名区域”设为模板目录；用户选过且仍存在的模板目录优先保留。
- S1 选定的输入、模板目录、外部文件与输出路径立即持久化；用户清空外部文件后，不会在下次启动自动恢复默认文件。

影响文件：`core/src/base_audit/web_app.py`、`core/src/base_audit/settings.py`、`core/tests/test_web_app.py`。

配置影响：无工作簿或审核规则变化；S1-CFG-01 v2.0.0（未变）。

验证：路径与发行配置定向测试 85 passed、3 subtests passed；S1/S2 联合回归 733 passed、13 skipped、2 failed。两项失败均来自本机忽略的 config-real 内容与固定测试期望不一致（指标参照 0/105、通用单位设置 0/3）；本轮未改真实配置。未重新打包。

GPT修改记录：`GPT修改记录/20261021_221251_S1S2S5默认外部文件与命名区域模板路径.txt`。

Git提交：未提交。

## v2.2.1 — 2026-10-15 13:45

变更类型：Bugfix / PATCH

前版本：v2.2.0
新版本：v2.2.1

修改内容：
- 多个 TDF LibreOffice DEB 版本并存时，direct 引擎按数值版本优先选择最新的 `/opt/libreofficeX.Y/program/soffice`；`BASE_AUDIT_SOFFICE` 和 `BASE_AUDIT_LIBREOFFICE_DIR` 显式设置仍优先。
- 修复安装页已核验新版本，但 S1 计算引擎仍从 PATH 或系统目录选择旧版本的状态不一致。

影响文件：
- `core/src/base_audit/engines/libreoffice.py`
- `core/src/base_audit/native/libreoffice_install.py`
- `core/src/base_audit/web_app.py`
- `core/frontend/web/index.html`

配置影响：无工作簿、规则或业务配置变化。
相关配置：S1-CFG-01 v2.0.0（未变）。
测试：未运行自动化测试；Python 语法编译检查通过。多版本选择和 S1 计算需在 Linux/V11 真机确认。
GPT修改记录：GPT修改记录/20261015_134524_LinuxLibreOffice版本与麒麟维护模式管理.txt。
Git提交：未提交。

## v2.2.0 — 2026-10-15 09:43

变更类型：Feature / MINOR

前版本：v2.1.0
新版本：v2.2.0

修改内容：
- 支持从应用数据目录读取显式提供的 KARE 环境配置；没有该配置时，原有 direct/玲珑自动发现逻辑不变。
- KARE 配置读取不参与麒麟 V11 内置 LibreOffice 安装；V11 Maintain Mode 的主系统安装会使用原有 direct LibreOffice 发现路径。

影响文件：
- `core/src/base_audit/engines/libreoffice.py`
- `core/src/base_audit/web_app.py`
- `core/tests/test_native_uos.py`

配置影响：无工作簿或业务规则变化；可选读取用户显式提供的 `data/libreoffice-kare.json`，内置安装器不会自动创建该文件。
相关配置：S1-CFG-01 v2.0.0（未变）。
测试：KARE 引擎模拟用例通过；用户报告麒麟 V11 主系统 LibreOffice 已被 S1-F1 直接发现并完成实机测试。
GPT修改记录：GPT修改记录/20261015_麒麟V11改用主系统APT安装LibreOffice.txt。
Git提交：未提交。

## v2.1.0 — 2026-10-15 09:03

变更类型：Feature / MINOR

前版本：v2.0.0
新版本：v2.1.0

修改内容：
- 增加可选 KARE LibreOffice 计算引擎，命令采用已提供的 `kare run <app-id> ...` 调用形式。
- 临时工作簿、转换输出与独立 profile 可映射到 KARE upper backing directory；传给兼容环境的路径转换为其根文件系统内路径。
- KARE 入口要求显式配置 app ID；多个 KARE 环境时显式指定 upper 路径，不影响直接 LibreOffice 和玲珑引擎。

配置影响：无工作簿或业务规则配置变化；新增运行环境变量见 docs/麒麟V11_KARE安装LibreOffice与S1计算方案.md。
测试：KARE 模拟 Docker 内的项目真实计算器公式重算通过；定向 KARE/direct/玲珑/临时文件用例 5 passed；S1 组 445 passed / 13 skipped / 3 failed，3 项因 Windows 测试机无法发现 Linux LibreOffice 而在 native 公式预检失败。麒麟 V11 Guest 真机安装和调用未执行。
GPT修改记录：GPT修改记录/20261015_090302_麒麟V11_KARE与LibreOffice适配.txt。
Git提交：未提交。

已知限制：KARE DEB 安装流程与应用 ID/upper 路径需在麒麟 V11 中验证；本次不修改现有 pkexec/dpkg 安装器。

## v2.0.0 — 2026-10-13 21:54

变更类型：Breaking / MAJOR

前版本：v1.2.0
新版本：v2.0.0

修改内容：
- 汇总登记 2026-10-13 已完成的结果与公式协议调整，不倒填历史版本。
- 审核结果在“错误类型”后新增“校验类型”，输出由 12 列变为 13 列；COM 和 Native 提取/写出同步。
- 模板公式按“级别|校验类型|校验指标|描述|当前值|对比值|差值”依次解析；短公式按前几段填充，不按旧段序解释。
- 条件格式批注写入校验类型，规则公式写入描述，二者同时保留。
- 公式问题的非空校验类型进入身份键，升级首次历史匹配可能出现一次“新增”。

影响文件：
- core/src/base_audit/template.py、models.py、history.py、excel_com.py
- core/src/base_audit/conditional_context.py、native/openpyxl_workbook.py、native/history_io.py、native/result_writer.py
- core/src/base_audit/config_guide.py、name_config.py

配置影响：S1-CFG-01 v1.0.5 → v2.0.0，结果/历史表增加一列；已有模板需人工按七段顺序核对。
相关配置：S1-CFG-01 v2.0.0。

测试：本轮为版本登记，未运行自动化回归。原实现测试记录见 GPT修改记录/20261013_122608_审核结果新增校验类型列与模板公式七段式.txt，不作为本轮执行结果。
验收限制：未重新打包或执行 Office/平台真机验收。

GPT修改记录：GPT修改记录/20261013_215438_统一3.3写入口径并补齐发布组件版本.txt。
Git提交：未提交。

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
