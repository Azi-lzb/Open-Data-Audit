# S1-F01 汇总核查表校验

当前版本：v2.6.3

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
