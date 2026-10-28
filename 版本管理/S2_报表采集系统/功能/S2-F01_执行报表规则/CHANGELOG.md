# S2-F01 执行报表规则

## v10.8.1 — 2026-10-27 21:22

变更类型：Bugfix / PATCH

前版本：v10.8.0；新版本：v10.8.1。

- 大集中数据中的空白、非数值及其他可继续处理的数据提示只写入运行记录，不再写入 Excel“校验规则”工作表，也不再计入审核异常数。
- 运行记录保留逐条提示；审核完成时显示非阻断提示数量。提示不会中断审核。
- 外部工作表表头中的 `null`/`NULL` 视为空白占位，不作为指标列或生成空白值提示；普通报表和时序格式口径一致。
- 外部核对取数口径不变：匹配行中的空白或非数值仍按 0 参与核对。

影响文件：S2-F01 外部数据提示处理、完成状态与运行日志；相关测试。

配置影响：无；S2-CFG-01、S2-CFG-02 不变。

测试：S2 回归 464 passed / 6 skipped / 1 failed；唯一失败为既有版本清单测试中 S5-F01 期望 3.0.0、MANIFEST 为 3.0.1，与本次 S2 修改无关。定向测试与编译结果见本轮 GPT 修改记录。

补充记录：`GPT修改记录/20261027_212600_S2-F01_null表头按空白处理.txt`（同一未提交 v10.8.1 变更批次）。

## v10.8.0 — 2026-10-22

变更类型：Feature / MINOR

前版本：v10.7.0；新版本：v10.8.0。

- 高级设置新增大集中外部数据读取引擎选项：自动、Calamine、OpenPyXL。
- 自动模式优先 Calamine；仅自动模式会在快速读取失败时回退 OpenPyXL。显式选择的引擎失败时明确报错，不静默切换。
- 运行记录注明实际读取引擎及版本、读取耗时和自动回退原因。普通报表/时序格式判断、取数及结果语义不变。
- 基准结果：真实外部工作簿（1,488 个值）Calamine 中位数 0.399 秒、OpenPyXL 0.467 秒；10,000 行合成时序工作簿分别 1.117 秒、3.474 秒。两者业务值对拍一致。
- 结果写入实测未引入 Rust 后端：Direct OOXML 对小中型结果更快；Rust 原型仅在 5,000/10,000 行时节省约 0.064/0.159 秒端到端时间，尚不足以抵消跨平台构建与维护成本。

影响文件：S2-F01 读取路由、高级设置 UI 与相关测试。

配置影响：无；S2-CFG-01 与外部文件工作簿结构不变。

测试：Windows core 全量 1,196 passed、7 skipped；WSL Ubuntu AMD64 core 全量 1,182 passed、21 skipped；读取与写入基准按 GPT 修改记录复核。

GPT修改记录：`GPT修改记录/20261022_170852_S2_大集中数据Calamine路由与读写基准.txt`。

Git提交：未提交。

## v10.7.0 — 2026-10-21

变更类型：Feature / MINOR

前版本：v10.6.4；新版本：v10.7.0。

- 首次启动时，若发行配置中带有“大集中外部数据模板”，默认填入 S2 的大集中系统数据路径。
- 用户选定的其他文件仍优先使用并持久化；手动清空会保留为未启用状态，不会在重启后自动恢复默认文件。
- S2 外部文件读取、规则计算与结果格式不变；S2-CFG-01 v8.4.0、S2-CFG-02 v1.0.0 未变。

影响文件：`core/src/base_audit/web_app.py`、`core/src/base_audit/settings.py`、`shell-flask/Linux-2-打包DEB与TARGZ.sh`、相关测试。

验证：S1/S2 联合回归 733 passed、13 skipped、2 failed。两项失败均来自本机忽略的 config-real 内容与固定测试期望不一致（指标参照 0/105、通用单位设置 0/3）；本轮未改真实配置。路径与发行配置定向测试 85 passed、3 subtests passed。未重新打包。

GPT修改记录：`GPT修改记录/20261021_221251_S1S2S5默认外部文件与命名区域模板路径.txt`。

Git提交：未提交。

## v10.6.4 — 2026-10-21

变更类型：Bugfix / PATCH

前版本：v10.6.3；新版本：v10.6.4。

- 行名称为空时，只有指标列实际有值才提示行标签缺失并忽略该行。
- 时序表中日期、报表名称、报表代码、版本号、报表组或行顺序号等元数据存在，但行名称和指标列为空，按正常无数据行忽略，不再误报。
- 只有整行全空且位于有效数据区内部时，才给出空行已忽略的软提示；尾部空行不提示。空格和不间断空格按空白处理。
- 普通报表读取时不再把已知元数据列当作指标列。

影响文件：`core/src/base_audit/systems/s2_report_collection/importer.py`、`core/tests/test_period_v2.py`、`core/tests/TEST_GROUPS.md`。

配置影响：无；输出仍使用既有“校验规则”工作表记录提示。

测试：`test_v2_external_blank_row_label_is_ignored_with_soft_warning` 两种布局用例 2 passed；`python -m pytest -q -rs -m s2 core/tests`：445 passed / 6 skipped / 0 failed / 3 subtests passed（719 deselected）；版本清单 8 passed；compileall、MANIFEST JSON 与 `git diff --check` 通过。真实业务数据及平台真机未执行。

GPT修改记录：`GPT修改记录/20261021_213956_S2S5外部空标签行误报修正.txt`。

Git提交：未提交。

## v10.6.3 — 2026-10-21

变更类型：Bugfix / PATCH

前版本：v10.6.2；新版本：v10.6.3。

- 时序布局中有行名称、但数据日期空白或格式无效的记录不再导致整份外部文件导入中止；该行会被忽略，并在运行结果中给出行号和原因提示。
- 仍只取与本期完整报告日期匹配的数据，不会把异常日期或其他日期的数据当成本期；若没有本期有效记录，对应规则仍明确提示缺少本期数据。
- 重复时序记录和不完整时序表头仍作为结构问题阻断，避免猜测取数。
- 普通报表和时序布局的空值、非数值值及空行/空标签处理保持一致，按 v10.6.2 规则提示并继续处理。

影响文件：`core/src/base_audit/systems/s2_report_collection/importer.py`、`core/tests/test_period_v2.py`、`core/tests/TEST_GROUPS.md`。

配置影响：无；输出仍使用既有“校验规则”工作表记录提示。

测试：定向共用测试 159 passed；`python -m pytest -q -rs -m s2 core/tests`：445 passed / 6 skipped / 0 failed / 3 subtests passed；版本清单 8 passed；compileall、MANIFEST JSON 与 `git diff --check` 通过。平台真机及真实业务数据未执行。

GPT修改记录：`GPT修改记录/20261021_210609_S2S5外部异常数据继续处理与最终提示.txt`。

Git提交：未提交。

## v10.6.2 — 2026-10-21 17:43

变更类型：Bugfix / PATCH

前版本：v10.6.1；新版本：v10.6.2。

- 外部文件中报表项目已匹配时，空白或非数值指标单元格按 0 参与外部核对；结果“校验规则”增加软性提示，写明行号、报表项目、指标列及按 0 处理，请复核是否误删或改动。
- “行名称”为空且该行无指标值时忽略该行并软提示；“行名称”为空但有指标值时跳过该行、不阻断审核，并提示检查名称是否误删。普通报表中间全空行同样软提示，尾部格式空行不产生提示。
- 审核完成状态和运行记录显示外部数据提示条数，并引导查看“校验规则”工作表。
- 保留既有时序行为：仍规范化完整日期，并仅匹配与本期报送日期一致的数据；未新增日期校验。

影响文件：`core/src/base_audit/systems/s2_report_collection/{importer,external_engine,validation,service}.py`、`core/src/base_audit/web_app.py`、S2 回归测试与测试组清单。

配置影响：无（S2-CFG-01 v8.4.0 不变）。结果仍为三张分类审核表；新增提示写入既有“校验规则”工作表。

测试：S2 回归组 443 passed / 6 skipped / 0 failed / 3 subtests passed；新增普通/时序空值与非数值、空名称行、完成提示及结果工作簿验证。其余检查见本次 GPT 修改记录。未重新打包。

GPT修改记录：`GPT修改记录/20261021_174327_S2F1外部数据空值按零并提示.txt`。

Git提交：未提交。

## v10.6.1 — 2026-10-21 17:26

变更类型：Bugfix / PATCH

前版本：v10.6.0；新版本：v10.6.1。

- 外部文件时序格式仅在表头同时含“数据日期”和“行名称”时识别。纯报表格式若将原本空白的首列标题写成“行名称”，现在仍按普通报表解析，并使用该列匹配报表项目；只有“数据日期”而没有“行名称”仍提示时序表头不完整。
- 时序数据行有指标值但“行名称”为空时，错误提示说明无法匹配报表项目，并建议检查导出不完整或名称误删。
- 普通/时序格式中，报表项目已匹配但对应指标数值单元格为空时，结果明确提示该值为空白且不会按 0 处理，并建议检查数据或重新导出。缺少指标列也单独说明。空值仍按缺数审核，不作为 0 参与比较。

影响文件：`core/src/base_audit/systems/s2_report_collection/{importer,external_engine}.py`、`core/tests/test_period_v2.py`、`core/tests/TEST_GROUPS.md`、`core/tests/test_version_manifest.py`、`版本管理/VERSION_MANIFEST.json`。

配置影响：无（S2-CFG-01 v8.4.0 不变）。输出结构不变；外部数据仍只支持包含“集中系统数据”工作表的 `.xlsx` 文件；机构参照表可选，无外部映射时回退 S2 配置。

测试：`python -m pytest -q core/tests/test_period_v2.py core/tests/test_s2_units.py`，150 passed；`python -m pytest -q -m s2 core/tests`，439 passed / 6 skipped / 0 failed / 3 subtests passed。`compileall`、MANIFEST JSON 解析及目标文件 `git diff --check` 通过。未执行真实业务数据复核或重新打包。

GPT修改记录：`GPT修改记录/20261021_172655_S2F1外部文件识别与缺数提示.txt`。

Git提交：未提交。

## v10.6.0 — 2026-10-20 23:52

变更类型：Feature / MINOR

前版本：v10.5.0；新版本：v10.6.0。

- 新增 python-calamine 0.3.1（Rust 扩展）作为两期报表文件的快速读取引擎：`importer._open` 优先走 calamine，整包缺失（兼容构建未打包）时静默回退 xlrd/openpyxl，单个文件读取失败时逐文件提示“快速读取引擎处理失败，已用兼容引擎重读”并回退，最坏情况只是回到原速，不影响结果。
- 值形态与行号逐值对齐传统引擎：.xls 对齐 xlrd（整型与布尔还原为 xlrd 的 float/int 形态），.xlsx 对齐 openpyxl（整型 float 还原 int、空串对齐 None、date 补零点 datetime）；calamine `to_python` 会裁掉前导空行，已按 `sheet.start` 补回等量空行，保证 `rows[3:]` 表头跳过与 SourceRef 行号与旧引擎逐行一致（该问题由既有用例 test_v2_importer_preserves_percentage_data_type 捕获后修复，并把前导空行写入对拍夹具防回归）。
- 已知有意的边界差异：.xls 值列的日期单元格 xlrd 读作序列数 float 参与数值计算，快速引擎读作 date 走非数值路径（更符合业务语义）；报表值列不含日期，实际不受影响。
- 性能（本机 549 对/1098 文件合成样例，预热后两轮）：两期读取 4.26s → 1.02s（4.2x），记录数与警告两引擎完全一致（16470/0）；结合 v10.4.2 的 Direct OOXML 导出，同规模全程预计从 7.5s 降到约 4.6s。注：冷启动首轮会慢（新写文件触发 Defender 首扫），基准以预热后为准。
- 打包：requirements.txt 新增 `python-calamine==0.3.1`（钉版：新版要求 Python≥3.11，Win7 兼容包用 32 位 Python 3.8）；Windows build_exe.py COMMON_HIDDEN 与 Linux-2 脚本 `--collect-all python_calamine`、依赖导入检查均已登记；Win7 32 位构建 venv 已装入 cp38-win32 wheel 验证可用。Linux wheel 为 manylinux_2_17（glibc≥2.17），覆盖统信 UOS 2.28、麒麟 V10 SP1 2.31 及 ARM64 glibc 2.23 冒烟目标；每包增重约 1.5~2MB。

影响文件：`core/src/base_audit/systems/s2_report_collection/importer.py`、`core/tests/{test_fast_reader,conftest,TEST_GROUPS}`、`shell-flask/{requirements.txt,build_exe.py,Linux-2-打包DEB与TARGZ.sh}`、`版本管理/VERSION_MANIFEST.json`。

配置影响：无（S2-CFG-01 v8.4.0 不变）。打包影响：新增运行依赖 python-calamine==0.3.1（约 1.5~2MB/包），各平台打包脚本已接入；Linux DEB/tar.gz 与 Win7 包需重打包验证。

测试：S2 定向新增 test_fast_reader.py 6 项（xlsx/xls 逐值逐类型对拍、缺包静默回退、单文件失败提示回退、两引擎导入全量一致，夹具含前导空行/空值/整数/小数/布尔/日期/尾部空行）；S2 回归组 431 passed / 2 failed / 6 skipped（Flask 跳过），两项失败仍为本机既有“输出文件单位=万元 vs 测试断言元”口径差异。

提交前复核（2026-10-21）：按用户授权修正两处将输出单位写死为元的测试，保留合法单位及配置解析一致性验证，真实配置继续使用万元。S2 433 passed/6 skipped，core 全量 1054 passed/13 skipped/0 failed/0 errors；真实两期报表的单 Sheet、按机构和按报表三种布局分别使用两引擎运行，审核结论、汇总及全部工作簿值一致，输入哈希未变。未执行 UOS/麒麟、ARM64 或 Office 实机验收。

GPT修改记录：`GPT修改记录/20261020_235228_S2接入calamine快速读取引擎.txt`、`GPT修改记录/20261021_004437_S2输出单位测试按配置验证并同步主分支.txt`。

Git提交：7ebaa9e24810cad4d3a640a092281cbe505df88b。

## v10.5.0 — 2026-10-20 22:47（23:05 补缺失级别设置，23:16 补缺行默认执行）

变更类型：Feature / MINOR

前版本：v10.4.2；新版本：v10.5.0。

- 结果工作簿“环比规则”表在“指标代码”与“指标名称”之间新增“数据属性”列，取值来自指标参照配置（文字/余额/累发/百分数/个数），未配置指标时回落记录值类型，与策略组匹配口径一致；“外部核对规则”“校验规则”两表列不变。
- 通用设置新增可选设置项“文字变动默认级别”（硬性/软性，缺省按硬性执行）：文字指标两期不一致（“文字指标变化”）与文字指标单侧缺失（“本期无，上期有”/“本期有，上期无”）行按该设置定级。
- 通用设置新增可选设置项“本期缺失默认级别”“上期缺失默认级别”（硬性/软性，缺省按硬性执行）：仅控制非文字指标（余额/累发/百分数/个数），按审核说明方向分别生效——“本期无，上期有”（本期整行缺指标、本期无数值或本期为 0 按无处理）跟随“本期缺失默认级别”，“本期有，上期无”（上期整行缺指标、上期无数值或上期为 0 按无处理）跟随“上期缺失默认级别”。此前非文字单侧缺失固定软性，现默认输出硬性（用户定稿口径）。
- 通用设置全部设置行（单位、隐藏、三项级别）缺失时不再报“缺少设置项”阻断：按默认执行（源/外部/输出单位=元、隐藏=否、级别=硬性），运行结果“校验规则”表与配置检查中逐项提示补齐（warning 级，不升阻断）；“换算范围”为纯说明行，缺行不提示；填写了非法值的行仍明确报错（含工作表与行号）。两期表单不一致仍恒为硬性。
- 使用说明（period 指南）补充三项设置、缺行默认口径与新输出列说明，CONFIG_VERSION 2.2 → 2.3 触发现有工作簿说明页刷新。

影响文件：`core/src/base_audit/systems/s2_report_collection/{models,config,period_engine,exporter}.py`、`core/src/base_audit/config_guide.py`、`core/tests/{test_period_v2,test_s2_units,test_s2_config_check,test_version_manifest}.py`、`版本管理/VERSION_MANIFEST.json`。

配置影响：S2-CFG-01 v8.3.0 → v8.4.0（通用设置新增三行可选设置，B8:B10 带硬性/软性下拉，全部设置行可缺省按默认执行）；本机正式/默认配置已同步加行（均=硬性）；正式配置“校验规则”级别下拉范围 E2:E998 对齐为 E2:E1000。产品整包及其他系统版本不变。

测试：S2 定向（test_period_v2、test_s2_units、test_s2_config_check）165 passed / 2 failed；S2 回归组 426 passed / 2 failed / 6 skipped（Flask 环境跳过）。两项失败均为既有本机配置差异（正式配置输出文件单位为万元而测试断言元，用户已知，见 v10.4.2 记录）；硬性/软性下拉范围差异已随本次配置修正消失。

GPT修改记录：`GPT修改记录/20261020_224706_S2环比数据属性列与文字变动默认级别.txt`、`GPT修改记录/20261020_230533_S2本期与上期缺失默认级别设置.txt`、`GPT修改记录/20261020_231657_S2通用设置缺行按默认执行.txt`。

Git提交：未提交。

## v10.4.2 — 2026-10-20 21:42

变更类型：Bugfix / PATCH

前版本：v10.4.1；新版本：v10.4.2。

- 运行记录现在按审核阶段说明正在处理的内容；本期/上期文件读取和结果表写入按约10%进度记录计数，并附当前文件名，避免每个文件都刷屏。
- 记录导入指标数、环比/校验/外部核对结果数、工作簿保存阶段，以及最终机构/表单/指标/异常汇总和结果文件。
- 进度与运行记录使用同一后端事件，不另行重复读取源文件；审核结果与输出结构不变。
- 固定三张审核结果表改用 Python 标准库 zipfile + XML 直接写 OOXML，避免 openpyxl 为每个单元格建立对象；异常日期/自定义数据类型仍回退 openpyxl。读取和其他通用工作簿操作不变，不需要 LibreOffice。
- 在 549 对/1098 个样例文件、19032 条指标、28182 条审核结果上，完整运行从 15.694 秒降至 7.534 秒；导出阶段从 9.4869 秒降至 2.042 秒。finding、summary、三张表的全部单元格、冻结窗格、筛选、隐藏行和样式均与原结果一致；输出约小 3.7%，输入 SHA-256 未变。该扩展样例不是用户当前原始数据或 Linux 实机承诺值。

影响文件：`core/src/base_audit/web_app.py`、`core/src/base_audit/systems/s2_report_collection/{service,exporter}.py`、`core/tests/{test_period_v2,test_web_app,test_version_manifest}.py`、`core/tests/TEST_GROUPS.md`、`版本管理/VERSION_MANIFEST.json`。

配置影响：无；S2-CFG-01 v8.3.0、S2-CFG-02 v1.0.0、产品整包及其他系统版本不变。未改真实配置、原始业务工作簿或打包依赖。

测试：S2-F01 定向 97 passed，S2 回归组 419 passed / 6 skipped / 2 failed / 0 errors，628 deselected、3 个子场景 passed。两项失败为既有本机真实配置差异：“校验规则”缺少硬性/软性下拉、输出单位为万元而测试要求元；6 项因缺少 Flask 跳过。Direct OOXML 输出由 openpyxl 重新打开并逐格比较，值、数据类型、样式 ID、列宽、冻结窗格、筛选、隐藏行均无差异。Python 编译与 `git diff --check` 通过；未执行 Office/Linux 实机测试或平台打包。

GPT修改记录：`GPT修改记录/20261020_211421_S2F01运行记录增加阶段与进度细节.txt`。

Git提交：未提交。

## v10.4.1 — 2026-10-20 20:53

变更类型：Bugfix / PATCH

前版本：v10.4.0；新版本：v10.4.1。

- 按用户要求移除执行按钮旁的独立进度框及其样式/渲染逻辑，只保留运行记录上方的一条进度条。
- 阶段、完成数量和耗时合并显示在该进度条同一行，当前文件或错误详情留tooltip；无计数阶段不显示0/0。实际百分比、失败完成量、终态收起计时器及其他任务旧值隔离保持。
- 比较了1098个XLS的直接读、xlrd取行/按需/稀疏选项、Python纯数值转换XLSX后读取，以及临时安装的Calamine候选。直接xlrd约2.25–2.64秒，转写并读XLSX下界约33.47秒，Calamine约0.22–0.30秒。正式读取引擎和依赖不变；仅将独立试验作为后续优化参考。

影响文件：`core/frontend/web/index.html`、`core/tests/test_distribution_profile.py`、`core/tests/test_version_manifest.py`、`core/tests/TEST_GROUPS.md`、`版本管理/VERSION_MANIFEST.json`。

配置影响：无；S2-CFG-01 v8.3.0、S2-CFG-02 v1.0.0、产品整包及其他系统版本不变。未改配置、原始文件或打包依赖，未打包。

测试：前端定向20 passed、0 skipped、0 failed、0 errors（含完整JS语法检查）；最终S2回归418 passed、6 skipped、2 failed、0 errors，628 deselected及3个subtests passed。失败仍为既有硬性/软性下拉和输出单位配置差异，6项因缺少Flask跳过，无新增失败。浏览器已确认仅一条底部进度、详情在同一行；编译和 `git diff --check` 通过。

GPT修改记录：`GPT修改记录/20261020_205332_S2F01仅保留底部进度条并比较XLS读取方案.txt`。

Git提交：未提交。

## v10.4.0 — 2026-10-20 20:32

变更类型：Feature / MINOR

前版本：v10.3.1；新版本：v10.4.0。

- S2 页面新增阶段进度条，显示读取本期/上期文件数、环比/校验/外部核对阶段、三张结果表写入行数、保存阶段和耗时。保存并完成后才显示100%，失败保留已有进度；其他系统运行时隐藏旧的S2进度。
- 底部通用进度在S2运行时采用后端进度快照，其他任务继续使用原逻辑；终态按既有停留时间收起，避免800ms轮询反复延后计时器。
- 对比对象表头在列表内滚动时固定；数据及筛选条件未变时复用DOM，重建时保留滚动位置。后台已有配对列表不再于每次审核结束重扫，空清单仍重试元数据扫描。
- 每个工作簿的绝对路径只解析一次，xlsx导入限定既有A:C取数范围，导出复用相同样式对象，业务规则及输出协议不变。

影响文件：`core/frontend/web/index.html`、`core/src/base_audit/web_app.py`、`core/src/base_audit/systems/s2_report_collection/{service,importer,exporter}.py`、`core/tests/{test_distribution_profile,test_period_v2,test_web_app,test_version_manifest}.py`、`core/tests/TEST_GROUPS.md`、`版本管理/VERSION_MANIFEST.json`。

配置影响：无；S2-CFG-01 v8.3.0、S2-CFG-02 v1.0.0、产品整包版本不变。未修改真实配置或原始业务工作簿，未打包。

性能：由现有样例扩展的549对/1098个文件（约10.93MiB），19032条本期指标、28182条审核结果；无采样分析器的同机基线29.940秒，优化后18.407秒和15.694秒。结果模型、汇总和全部工作簿单元格一致；工作表、样式、关系和包结构逐部件一致，仅文档时间戳不同，输入SHA-256不变。该结果来自扩展样例，不代表用户当前549对原始文件或Linux实机的保证耗时。

测试：定向211 passed、0 skipped、0 failed、0 errors，1 deselected及3个subtests passed；前端独立20 passed，含Node脚本语法和进度/筛选行为。最终S2回归418 passed、6 skipped、2 failed、0 errors，628 deselected及3个subtests passed；失败仍为校验规则缺少硬性/软性下拉、本机输出单位万元与测试要求元不一致，6项因缺少Flask跳过，无新增失败。Python编译与 `git diff --check` 通过。已在浏览器验证运行/完成/失败进度、夜间主题及列表内滚动固定表头。

GPT修改记录：`GPT修改记录/20261020_203256_S2F01阶段进度与固定表头及批量性能优化.txt`。

Git提交：未提交。

## v10.3.1 — 2026-10-20 20:23

变更类型：Bugfix / PATCH

前版本：v10.3.0；新版本：v10.3.1。

- 外部核对缺数且报表项目不为空时，审核说明改为“大集中系统数据未找到该报表项目，请先完善大集中数据的报表定义”。
- 报表项目为空时，审核说明改为“请完善报表采集系统配置或外部文件中的机构参照映射”。普通和时序数据使用同一口径，结果工作簿同步显示。
- 普通格式的计算过程补充已选报表项目与指标；时序格式已有的具体日期、项目及指标缺数信息继续保留。缺数仍留空并提示异常，真实0、缺少机构标签列时的既有按0兜底及比较逻辑不变。

影响文件：`core/src/base_audit/systems/s2_report_collection/external_engine.py`、`core/tests/test_s2_units.py`、`core/tests/test_period_v2.py`、`core/tests/test_version_manifest.py`、`core/tests/TEST_GROUPS.md`、`版本管理/VERSION_MANIFEST.json`。

配置影响：无；S2-CFG-01 v8.3.0、S2-CFG-02 v1.0.0、产品整包版本不变。未修改配置或原始工作簿，未打包。

测试：外部核对及版本定向 68 passed、0 skipped、0 failed、0 errors，80 deselected；最终 S2 回归 413 passed、6 skipped、2 failed、0 errors，628 deselected。失败仍为既有配置差异（校验规则缺少硬性/软性下拉、本机输出单位万元与测试要求元不一致），6项因缺少Flask跳过，没有新增失败。Python 编译、WIP 增量复核和 `git diff --check` 通过；未执行真实业务文件、Office 或 Linux 实机验收。

GPT修改记录：`GPT修改记录/20261020_202339_S2F01外部缺数说明区分报表项目与机构映射.txt`。

Git提交：未提交。

## v10.3.0 — 2026-10-20 20:18

变更类型：Feature / MINOR

前版本：v10.2.0；新版本：v10.3.0。

- 对比对象筛选新增“配对状态”：全部报表、单边报表、本期缺少、上期缺少、两期齐全；可与机构名称、信用代码、表单条件组合，统一清空并在常规刷新时保留。
- 本期缺少对应仅上期存在，上期缺少对应仅本期存在；缺失列直接标明“本期缺少”或“上期缺少”，方便定位。
- 仍按现有工作簿/工作表存在性判断，空表算存在；不读取正文或更改后台配对、审核范围与结果输出。页面补充单边定义，保持显示筛选的范围说明。

影响文件：
- `core/frontend/web/index.html`
- `core/tests/test_distribution_profile.py`
- `core/tests/test_version_manifest.py`、`core/tests/TEST_GROUPS.md`
- `版本管理/VERSION_MANIFEST.json`

配置影响：无；S2-CFG-01 v8.3.0、S2-CFG-02 v1.0.0、产品整包版本不变。没有改写配置或原始工作簿，未打包。

测试：前端行为和版本定向 27 passed、0 skipped、0 failed、0 errors；最终 S2 回归 412 passed、6 skipped、2 failed、0 errors。两项失败仍为既有配置差异（校验规则缺少硬性/软性下拉、本机输出单位万元与测试要求元不一致），6项因缺少Flask跳过，没有新增功能失败。Node.js 完整脚本语法、各状态筛选和四条件组合/清空检查通过，测试 Python 编译和 `git diff --check` 通过；未执行浏览器视觉或 Linux 实机验收。

GPT修改记录：`GPT修改记录/20261020_201828_S2F01单边报表与缺少期别筛选.txt`。

Git提交：未提交。

## v10.2.0 — 2026-10-20 20:07

变更类型：Feature / MINOR

前版本：v10.1.0；新版本：v10.2.0。

- 在“文件列表与对比对象”标题右侧增加“筛选”按钮，可展开/收起筛选面板。支持机构名称、统一社会信用代码关键词和表单下拉条件，多个条件同时生效；名称和代码忽略首尾空格及大小写。
- 提供清空筛选、已启用条件数量、筛选后/全部对象数和无匹配提示。表单选项由当前清单去重生成，常规刷新保留条件；表单集合变化时清除失效选择，选项及显示文本正确转义。
- 筛选只作用于当前配对清单的显示，不重新读取工作簿或正文，不改变正式审核范围、源配对数据或 Excel 输出。页面明确说明正式审核仍处理所选目录中的全部报表。

影响文件：
- `core/frontend/web/index.html`
- `core/tests/test_distribution_profile.py`
- `core/tests/test_version_manifest.py`、`core/tests/TEST_GROUPS.md`
- `版本管理/VERSION_MANIFEST.json`

配置影响：无；S2-CFG-01 v8.3.0、S2-CFG-02 v1.0.0 及产品整包版本未变。没有修改配置或原始工作簿，没有打包。

测试：前端行为和版本定向 22 passed、0 skipped、0 failed、0 errors；S2 回归 407 passed、6 skipped、2 failed、0 errors。失败仍为既有配置差异：缺硬性/软性下拉、输出单位万元与测试要求元不一致；6项因缺少Flask跳过。Node.js 完整脚本语法、交付范围显示及筛选交互检查通过，测试 Python 编译和 `git diff --check` 通过；未执行浏览器视觉验收或 Linux 实机验收。

GPT修改记录：`GPT修改记录/20261020_200716_S2F01对比对象组合筛选.txt`。

Git提交：未提交。

## v10.1.0 — 2026-10-20 19:57

变更类型：Feature / MINOR

前版本：v10.0.0；新版本：v10.1.0。

- 文件配对清单的“对比对象”在机构名称下增加信用代码，保留原表单行；代码可选中复制、自动换行，悬停显示完整“统一社会信用代码”。
- 配对数据独立传递 `socialCreditCode`，仅来自标准导出文件名或 banks 工作表名。空表、单边以及未取得机构名称的 banks 子表均保留代码；没有识别出代码时显示“未识别”，不用机构名称代替代码。
- 配对键、数量、正文读取边界、banks 提醒、正式审核和 Excel 输出保持原有行为；不增加机构参照读取，不依赖名称映射是否成功。

影响文件：
- `core/src/base_audit/period_compare.py`
- `core/frontend/web/index.html`
- `core/tests/test_period_compare.py`、`core/tests/test_distribution_profile.py`、`core/tests/test_web_app.py`
- `core/tests/test_version_manifest.py`、`core/tests/TEST_GROUPS.md`
- `版本管理/VERSION_MANIFEST.json`

配置影响：无；S2-CFG-01 v8.3.0、S2-CFG-02 v1.0.0 不变。原始工作簿和产品整包版本未修改；未打包。

测试：配对、前端行为、Web API 和版本定向 113 passed、0 skipped、0 failed、0 errors；S2 回归 406 passed、6 skipped、2 failed、0 errors。失败仍为既有配置差异：缺少硬性/软性下拉、输出单位万元与测试要求元不一致；6项因缺少Flask跳过。Python 编译、Python 3.8 语法及 `git diff --check` 通过；未执行浏览器视觉验收或 Linux 实机验收。

GPT修改记录：`GPT修改记录/20261020_195747_S2F01对比对象显示统一社会信用代码.txt`。

Git提交：未提交。

## v10.0.0 — 2026-10-20 19:50

变更类型：Feature / Breaking / MAJOR

前版本：v9.3.1；新版本：v10.0.0。

- “外部核对规则”结果表在“社会信用代码”右侧插入“报表项目”，展示引擎本次选定的机构映射标签；只改变该结果表，由 15 列变为 16 列，原“表单”及后续列右移一列。
- 保持外部文件机构参照映射优先、CF2 机构参照回退的取数顺序。已选项目但缺少日期或指标数值时仍显示该项目供核对；没有机构映射时留空。标签列缺失按0兜底的原有说明继续保留。
- 新字段由外部核对引擎随审核发现传递，导出器直接展示；不从机构名称、审核说明或其他字段猜测项目，不改比较公式、金额换算、审核结论或配置业务数据。
- 报表项目自动换行并按文本保留，等号开头的标签不会被写成公式。更新列位变化后的隐藏行回归及普通/时序端到端断言。

版本说明：插入列改变固定列坐标，按 rules/11 的输出契约规则升 MAJOR；按列名读取结果的逻辑可继续识别。配置工作簿结构未变。

影响文件：
- `core/src/base_audit/systems/s2_report_collection/models.py`
- `core/src/base_audit/systems/s2_report_collection/external_engine.py`
- `core/src/base_audit/systems/s2_report_collection/exporter.py`
- `core/src/base_audit/config_guide.py`
- `core/tests/test_period_v2.py`、`core/tests/test_s2_units.py`
- `core/tests/TEST_GROUPS.md`、`core/tests/test_version_manifest.py`
- `版本管理/VERSION_MANIFEST.json`

配置影响：无；S2-CFG-01 仍为 v8.3.0，S2-CFG-02 仍为 v1.0.0。整包产品版本不变；未打包或修改原始工作簿。

测试：外部核对/导出、版本与说明定向 85 passed、0 skipped、0 failed、0 errors；S2 回归 402 passed、6 skipped、2 failed、0 errors。两项失败仍为既有本机配置差异：校验规则缺硬性/软性下拉、输出单位万元与测试要求元不一致；6项因缺少Flask跳过。Python 编译、Python 3.8 语法及 `git diff --check` 通过；未执行 Python 3.8.10 运行时或 Linux 实机验收。

GPT修改记录：`GPT修改记录/20261020_195038_S2F01外部核对结果显示报表项目.txt`。

Git提交：未提交。

## v9.3.1 — 2026-10-20 19:42

变更类型：Bugfix / PATCH

前版本：v9.3.0；新版本：v9.3.1。

- 两期文件配对和正式导入共用文件名、工作表名身份协议；正文的机构名称、机构代码仅作为指标值审核，空值、`--` 或误填内容不再覆盖文件身份。
- 配对预览只读取工作表目录，xlsx 不扫描正文或共享字符串，xls 按需读取工作表清单；完全空白的工作表也保留。数据日期继续不参与跨期配对。
- 标准逐机构逐表及 reports 文件从文件名保留机构代码和名称；banks 从工作表名取得机构代码，由机构参照补充名称，没有参照时显示代码。按名称备用匹配配置不能覆盖文件名中的另一机构代码。
- 页面推荐使用包含机构名称的两种文件名格式；任一期选择 banks 文件后，在文件列表上方及执行确认中提示改用“默认单sheet”或“按报表单sheet”重新导出。更换目录或清除选择后同步更新提醒；保留已有 banks 兼容。
- 指标值、空值和真实0的审核含义保持不变；缺报数继续给出相应提示，结果的机构字段来自稳定文件身份。补充说明及回归测试。

影响文件：
- `core/src/base_audit/report_source.py`（新增）
- `core/src/base_audit/period_compare.py`（配对预览）
- `core/src/base_audit/systems/s2_report_collection/importer.py`
- `core/src/base_audit/web_app.py`
- `core/frontend/web/index.html`
- `core/src/base_audit/config_guide.py`
- `core/tests/test_period_compare.py`、`core/tests/test_period_v2.py`
- `core/tests/test_web_app.py`、`core/tests/test_distribution_profile.py`
- `core/tests/TEST_GROUPS.md`、`core/tests/test_version_manifest.py`
- `版本管理/VERSION_MANIFEST.json`

配置影响：无；S2-CFG-01 保持 v8.3.0，S2-CFG-02 保持 v1.0.0。未修改公开或真实配置工作簿，原始业务文件未写入。整包版本未变，未重新打包。

测试：定向 204 passed（含 xls/xlsx、三种导出、空表/正文占位符、结果机构字段、只读哈希、banks 选择/清除与 Node.js 前端行为）；S2 回归 395 passed、6 skipped、2 failed、0 errors。两项失败仍为既有本机配置差异：校验规则缺少硬性/软性下拉，输出单位万元与测试要求元不一致；6 项因缺少 Flask 跳过。Python 3.8 语法检查、编译及 `git diff --check` 通过；未执行 Python 3.8.10 运行时或 Linux 实机验收。

GPT修改记录：`GPT修改记录/20261020_194259_S2F01按文件名表名配对空表并提示banks导出.txt`。

Git提交：未提交。

## v9.3.0 — 2026-10-20 19:10

变更类型：Feature / MINOR

前版本：v9.2.0；新版本：v9.3.0。

- 同一“集中系统数据”工作表按表头自动兼容普通布局与时序布局；在前 20 行识别“数据日期”与“行名称”，允许前置空行/说明行和字段调序，不按第一行是否非空猜格式。
- 时序数据以本期报送文件的完整日期、机构报表项目、外部指标名称取数；同一机构多个日期互不覆盖，不借用最新日期或跨期合计。
- 时序日期/机构存在性与数值分别保存，真实 0 正常比较；缺日期、缺机构行或空值继续输出未匹配说明，不能进入普通格式的“标签列缺失按0”兜底。
- 时序表头不完整/重复、同日同机构重复行、非法日期或数值明确报告工作表和行/字段，不静默覆盖或降级。
- 仍只读取正式数据表“集中系统数据”；样例或其他数据表不混用。保留普通布局、外部机构参照优先级与金额单位换算；补充使用说明。

影响文件：
- `core/src/base_audit/systems/s2_report_collection/importer.py`
- `core/src/base_audit/systems/s2_report_collection/external_engine.py`
- `core/src/base_audit/config_guide.py`
- `core/tests/test_period_v2.py`
- `core/tests/TEST_GROUPS.md`
- `core/tests/test_version_manifest.py`
- `版本管理/VERSION_MANIFEST.json`

配置影响：无；S2-CFG-01 保持 v8.3.0，S2-CFG-02 保持 v1.0.0；原始工作簿未修改。

测试：导入/导出、使用说明与版本清单定向回归 88 passed；S2 组 358 passed、6 skipped、2 failed。失败为既有本机配置差异：校验规则缺少硬性/软性下拉；本机输出单位为万元而测试固定断言元。6 项因缺少 Flask 跳过。真实样例结构只读解析及 SHA-256 保护通过；样例时序数值区为空，未完成真实数值端到端对拍。受影响 Python 编译、Python 3.8 语法解析及 `git diff --check` 通过；未执行 Python 3.8.10 运行时或麒麟/UOS 实机验收。

GPT修改记录：`GPT修改记录/20261020_191005_S2F01兼容大集中时序格式并按本期日期匹配.txt`。

Git提交：未提交。

## v9.2.0 — 2026-10-20 16:11

变更类型：Feature / MINOR

前版本：v9.1.0；新版本：v9.2.0。

- “集中系统数据”未能识别机构行标签列时不再中止 S2-F01；外部核对继续执行，并将对应大集中数按 0 参与比较。
- 审核说明明确标注“机构行标签列未识别，外部数按0参与比较”；计算过程仍使用原有差异绝对值格式。
- 该兜底仅由导入器明确标记的“机构行标签列缺失”场景触发；普通机构映射缺失或指标数值缺失仍按未匹配处理，不自动当作 0。

影响文件：
- `core/src/base_audit/systems/s2_report_collection/importer.py`
- `core/src/base_audit/systems/s2_report_collection/external_engine.py`
- `core/tests/test_period_v2.py`
- `core/tests/test_s2_units.py`
- `core/tests/test_version_manifest.py`
- `版本管理/VERSION_MANIFEST.json`

配置影响：无；S2-CFG-01 保持 v8.3.0，S2-CFG-02 保持 v1.0.0。

测试：导入定向测试 4 passed；零值兜底与普通缺数测试 2 passed；版本清单测试 8 passed；S2 组 333 passed、6 skipped、2 failed。两项失败与本次修改无关：本机配置缺少“校验规则”硬性/软性下拉验证；本机输出单位为“万元”，而现有测试固定断言“元”。6 项因缺少 Flask 跳过。受影响 Python 文件编译及 `git diff --check` 通过；未执行真实报表数据端到端对拍。

GPT修改记录：`GPT修改记录/20261020_161115_S2F01机构标签列缺失时外部数按0继续执行.txt`。

Git提交：未提交。

## v9.1.0 — 2026-10-20 09:26

变更类型：Feature / MINOR

前版本：v9.0.0；新版本：v9.1.0。

- 外部核对机构映射优先读取所选大集中工作簿的“机构参照表”，按“社会信用代码”匹配“报表项目”，并跳过标记为禁用的行。
- 保留旧工作簿兼容：缺少“机构参照表”时读取“参照表”；两个工作表同时存在时，以“机构参照表”为准。
- 外部映射没有对应机构时，继续使用 S2-CFG-01“机构参照”的“报表项目”作为回退值；S1 工作簿处理未变。
- 新参照表表头不完整时明确报错，不静默退回旧表。

影响文件：
- `core/src/base_audit/systems/s2_report_collection/importer.py`
- `core/tests/test_period_v2.py`
- `core/tests/test_s2_units.py`
- `core/tests/test_version_manifest.py`
- `版本管理/VERSION_MANIFEST.json`

配置影响：无，S2-CFG-01 与 S2-CFG-02 不变。

测试：`test_period_v2.py -k external_import` 3 passed；S2 组 331 passed、6 skipped、2 failed；版本清单测试 8 passed；Python 编译及 `git diff --check` 通过。两项失败是既有配置差异：本机“校验规则”没有硬性/软性下拉验证；本机输出单位为“万元”，而测试固定断言“元”。本机外部工作簿读取成功，解析 1,464 个数值和 188 条机构映射。未执行真实报表端到端对拍。

GPT修改记录：`GPT修改记录/20261020_092656_S2F01外部机构映射接入机构参照表.txt`。

Git提交：未提交。

## v9.0.0 — 2026-10-19 23:15

变更类型：Breaking

前版本：v8.0.0；新版本：v9.0.0。

- 环比结果表删除“规则编号”列；环比规则身份仍在引擎内部用于匹配，不再作为审核输出内容。
- 环比结果表删除“取数说明”列；外部核对结果表删除“规则编号”和“取数说明”列。
- 未触发环比规则、未达到最小变动门槛或命中规则未配置说明时，“审核说明”留空；命中且配置了分类说明时只显示该说明。
- 环比计算过程只保留计算公式，不附带策略组、规则命中或门槛说明；数值统一保留两位小数，百分数指标显示百分点差公式。
- 外部核对计算过程统一只显示“差异绝对值= abs(报表值 - 大集中值) = 差异值”，金额附输出单位；公式中的数值保留两位小数，不再显示容差比较或“成立/不成立”。规则比较方式、容差和审核状态计算不变。
- 校验规则的“取值明细”和“计算过程”中指标值统一保留两位小数；配置表达式原样显示，原始数值精度继续用于规则判断。
- 校验规则的“计算过程”同时显示配置原表达式与指标值代入后的表达式，便于对照公式和取数。
- 不改变规则匹配、最小变动门槛、审核级别或现有异常计算逻辑；是否通过由人工判断。

影响文件：
- `core/src/base_audit/systems/s2_report_collection/period_engine.py`
- `core/src/base_audit/systems/s2_report_collection/external_engine.py`
- `core/src/base_audit/systems/s2_report_collection/exporter.py`
- `core/src/base_audit/systems/s2_report_collection/rule_engine.py`
- `core/tests/test_period_v2.py`
- `core/tests/test_s2_units.py`
- `core/tests/test_s2_config_check.py`
- `版本管理/VERSION_MANIFEST.json`

配置影响：无；S2-CFG-01 保持 v8.3.0。

验证：`test_period_v2.py` 与 `test_s2_units.py`（排除读取本机配置并固定断言“元”的用例）86 passed、1 deselected。S2 组回归 328 passed、6 skipped、2 failed；失败分别是本机 `config-real/` 的“校验规则”缺少硬性/软性下拉验证，以及配置输出单位为“万元”而测试断言“元”，均属真实配置与测试固定值不一致；6 项因本机缺少 Flask 跳过。相关 Python 语法检查、`VERSION_MANIFEST.json` 解析及 `git diff --check` 均通过。

GPT修改记录：`GPT修改记录/20261019_221933_S2F01环比审核说明去规则编号和门槛判定.txt`、`GPT修改记录/20261019_222534_S2F01精简环比外部列并显示校验原表达式.txt`、`GPT修改记录/20261019_231515_S2F01环比计算过程仅保留公式并保留两位小数.txt`、`GPT修改记录/20261019_231824_S2F01外部核对差异绝对值公式精简.txt`、`GPT修改记录/20261019_232138_S2F01校验规则取值与计算展示两位小数.txt`。

Git提交：未提交。

## v8.0.0 — 2026-10-19 22:05

变更类型：Breaking

前版本：v7.2.0；新版本：v8.0.0。

- 三类审核结果统一只输出“硬性/软性”：原“严重/错误”映射为硬性，原“提示/关注/核实”合并为软性。
- 旧 S2 配置仍可读取；结果级别只用于分类和筛选，不改变规则执行、审核状态或流程阻断。
- 导出层拒绝写出旧级别，避免同一份新结果中重新出现三档值。

配置影响：S2-CFG-01 v8.2.1 → v8.3.0。
验证：`test_period_v2.py`、`test_s2_units.py`、`test_s2_config_check.py`；完整验证结果见 GPT修改记录。Git提交：未提交。

## v7.2.0 — 2026-10-16 16:04

变更类型：Feature / MINOR（外部数据模板维护）

前版本：v7.1.0；新版本：v7.2.0。

- 设置中心增加大集中外部数据模板的打开、设为默认和重置入口，在S1/S2交付可见。
- 模板路径由版本清单登记定位，独立默认快照恢复前必须成功备份；无快照明确失败，不重建或清空业务文件。
- 不改变S1/S2当前选择的外部运行文件，不自动启用外部核对；保持既有表头识别和代码→报表项目匹配协议。

配置影响：新增 S2-CFG-02 v1.0.0；S2-CFG-01 v8.2.0 不变。
验证：见 GPT修改记录/20261016_160450_大集中外部模板纳管与参照同步.txt。Git提交：1f83ef8a8c5758e6115bea42a62be19d926729ec。

## v7.1.0 — 2026-09-24

变更类型：Feature（向后兼容的配置检查展示增强）

前版本：v7.0.0
新版本：v7.1.0

修改内容：

- S2“检查配置”逐项展示文件读取、必需工作表、指标参照、环比策略以及正式配置解析的检查状态和问题明细。
- 明示检查范围：只读配置工作簿，不读取真实报表或大集中数据，也不执行审核规则；避免将“配置可解析”误认为“业务数据审核通过”。
- 保留原有检查报告字段，新增结构化检查明细供 Flask 与 pywebview 共用的页面展示；不改变业务规则、配置协议和正式配置工作簿。
- 用当前真实 S2 数据将 D20201017 在内存配置副本中分别试配“等于”“大于”“不等于”，验证左右值方向、严格比较和审核结论；正式配置保持“等于”。

配置影响：S2-CFG-01 保持 v8.0.0。

Git提交：未提交，待用户确认范围。

## v7.0.0 — 2026-09-23

变更类型：Breaking

前版本：v6.1.0
新版本：v7.0.0

修改内容：

- 外部核对固定使用大集中数据，配置列调整后支持等于、不等于、大于、大于等于、小于、小于等于；左侧始终是报表指标值，右侧始终是大集中指标值，比较成立为正常。
- “等于”使用绝对差异不超过容差判断；其他五种比较方式严格比较，不使用容差。
- 余额、累发的双方数值先转换到“通用设置”的输出文件单位；金额容差也按该输出单位填写和比较，不做四舍五入。
- 等于仍保留“与大集中系统数据一致/不一致”审核说明；其他方式使用“满足/不满足外部核对规则”，避免把大小关系误写成数值相等。

配置影响：S2-CFG-01 v7.0.0 → v8.0.0；旧“外部系统”列不再支持。

测试：详见 `GPT修改记录/20260923_235908_S2外部核对比较方式与配置列调整.txt`。

Git提交：未提交。

## v6.1.0 — 2026-09-23

变更类型：Feature

前版本：v6.0.2
新版本：v6.1.0

修改内容：

- 根据“通用设置”的两个独立开关，导出报表审核清单时隐藏“环比规则”中审核说明恰为“无变动”的行，以及“外部核对规则”中审核说明恰为“与大集中系统数据一致”的行。
- 默认两个开关均为“否”；隐藏只设置工作表行的隐藏属性，保留全部审核记录、单元格值和汇总数量，用户可以取消隐藏。
- 以完整文案精确匹配，避免把“与大集中系统数据不一致”误判为一致。

配置影响：S2-CFG-01 v6.0.0 → v7.0.0。

测试：详见 `GPT修改记录/20260923_222207_S2通用设置与审核清单行隐藏.txt`。

Git提交：未提交。

## v6.0.2 — 2026-09-23

变更类型：Bugfix

前版本：v6.0.1
新版本：v6.0.2

修改内容：

- 配置加载、配置检查和设置中心默认/重置流程改为从主 S2 配置工作簿读取和维护“单位设置”工作表。
- 移除对独立单位配置工作簿的运行时依赖；不改变金额换算精度、规则判断或审核结果语义。

影响文件：

- `core/src/base_audit/systems/s2_report_collection/config.py`
- `core/src/base_audit/systems/s2_report_collection/config_check.py`
- `core/src/base_audit/web_app.py`
- `core/src/base_audit/config_guide.py`

配置影响：S2-CFG-01 v5.0.0 → v6.0.0。

测试：详见本次 GPT 修改记录。

GPT修改记录：

- `GPT修改记录/20260923_210937_S2单位设置合并至主配置工作簿.txt`

Git提交：未提交。

## v6.0.1 — 2026-09-23

变更类型：Bugfix

前版本：v6.0.0
新版本：v6.0.1

修改内容：

- 修复旧版 `period_compare` 入口读取新版指标参照表时，将百分数、个数、文字等非金额指标误按元/万元换算的问题；仅余额、累发执行金额单位换算。
- 更新百分数环比回归断言，确保原始百分数值不被缩放。
- 修正资产负债平衡精度测试夹具中同一指标代码同时标成文字和金额的冲突。

影响文件：

- `core/src/base_audit/period_compare.py`
- `core/tests/test_percentage_point_compare.py`
- `core/tests/test_period_compare.py`
- `core/tests/test_config_guide.py`

配置影响：无；S2-CFG-01 保持 v5.0.0。

测试：定向与全量结果见对应 GPT 修改记录。

GPT修改记录：

- `GPT修改记录/20260923_181700_修复S2百分数单位误换算及6项测试失败.txt`

Git提交：未提交。

## v6.0.0 — 2026-09-23

变更类型：Breaking

前版本：v5.0.0
新版本：v6.0.0

修改内容：

- 本期、上期余额/累发金额在导入时先从源数据单位换算到输出文件单位，再进入环比、校验规则和后续计算。
- 外部文件余额/累发金额先换算到同一输出单位，再与报表值比较；环比最小变动值和外部核对容差均按输出文件单位解释。
- 元、万元、亿元换算直接调整十进制位，不量化、不四舍五入；外部金额导入和比较也使用 Decimal，避免先转换为二进制浮点。
- 规则表达式中的数值字面量按原始文本构造 Decimal，不先经过 Python float 或按 15 位有效数字量化。
- 百分数、个数和文字不进行金额换算；真实 0 保持为 0。“Thd值(万元)”仍按万元录入，并在表达式计算前精确换算到输出单位。

配置影响：S2-CFG-01 v4.0.0 → v5.0.0。

测试：S2 单位转换、导入、环比、外部比较、规则阈值、配置检查及默认配置恢复定向测试通过；完整测试结果详见 GPT 修改记录。

Git提交：未提交。

## v5.0.0 — 2026-09-23

变更类型：Breaking

前版本：v4.0.0
新版本：v5.0.0

修改内容：

- 新单位配置下，余额和累发源值按“源数据单位”换算为元；大集中值按“大集中数据单位”换算为元后比较。
- S2 环比/外部核对的金额结果值改为元口径；百分数、个数和文字不作金额换算。
- 最小变动值按源数据单位换算；大集中容差按大集中数据单位换算；“Thd值(万元)”规则阈值继续按万元解释。
- 删除旧单位字段的配置兼容；缺少新单位列或填入未知单位时配置检查/读取明确失败。

配置影响：S2-CFG-01 v3.0.3 → v4.0.0。

测试：S2 单位配置与引擎定向测试通过；详见 GPT 修改记录。

Git提交：未提交。

## v4.0.0 — 2026-09-22

变更类型：Breaking

前版本：v3.0.0
新版本：v4.0.0

修改内容：

- 环比结果的规则编号改为“策略组-规则编号”，例如 `PCT01-B005`、`BAL01-B005`，避免不同策略组复用 Bxxx 时无法区分。
- 审核说明改为稳定、可筛选的分类：优先使用「环比策略」的分类说明；命中规则但分类说明为空时统一写“命中规则”。
- 审核说明不再拼接每条记录的实际变动值；具体数值继续保留在「变动幅度」和「计算过程」。
- 未命中策略统一写“未命中规则”；最小变动门槛未通过且无分类说明时写“未达到最小变动值”。

配置影响：S2-CFG-01 v3.0.0 → v3.0.1，仅修正工作簿显示样式，不改变策略内容。

## v3.0.0 — 2026-09-22

变更类型：Breaking

前版本：v2.1.0
新版本：v3.0.0

修改内容：

- 新“环比策略”的上下限统一改用百分数数值口径：`30` 表示 `30%`，配置单元格不显示 `%`。
- 普通数值指标把相对变动率乘以 100 后匹配策略；百分数指标直接用百分点差匹配，因此两类指标共用直观刻度。
- 旧“环比规则”和无策略组的历史调用继续按小数口径兼容，避免旧测试/调用无提示失效。
- 正式 Bxxx 边界整体乘以 100，普通数值与百分数指标的既有审核结论保持不变。

配置影响：S2-CFG-01 v2.0.0 → v3.0.0，正式配置与默认配置同步更新。

测试：见本轮 GPT 修改记录。

## v2.1.0 — 2026-09-22

变更类型：Minor

前版本：v2.0.1
新版本：v2.1.0

修改内容：

- 环比配置改为指标级“环比启用/环比策略组/最小变动值”，运行时按数据属性绑定策略组。
- 环比策略采用每组连续、无重叠的 `[下限, 上限)` 区间；命中区间后还必须满足绝对变动门槛，最小变动值为 0 表示关闭第二道门槛。
- 百分数指标继续按百分点差匹配，余额/累发/个数按各自业务单位计算绝对变动值；文字指标明确为“不适用”。
- 策略组 0 条或多条命中不再静默默认，改为阻断并报告配置问题。
- 增加共享只读配置检查器、S2 设置页“检查配置”按钮和执行前阻断。

配置影响：S2-CFG-01 升级为 v2.0.0；正式配置与默认配置同步迁移。

测试与真实验收：

- S2-F01、配置检查和版本定向测试：43 passed。
- 真实目录 + 外部大集中工作簿：环比 728、外部核对 357、校验规则 3，合计 1088；相对 1096 基线减少的 8 条属于前一版本已确认的 Decimal 浮点尾差误报，不是策略组迁移造成。
- 非 COM core 测试：655 passed / 3 skipped；完整 COM/golden 集合因当前 Excel RPC 环境异常未能取得汇总。

## v2.0.1 — 2026-09-22

变更类型：Bugfix

前版本：v2.0.0
新版本：v2.0.1

修改内容：

- 校验规则表达式求值改用 Decimal，并按 Excel 风格 15 位有效数字归一化，避免 `9697998.71 <> 8875290.77 + 822707.94` 因二进制浮点尾差误报。
- 不修改 S2 配置中的业务容差；真实差异仍按规则正常命中。
- 增加资产负债等式回归测试：精确相等不生成 finding，差异 0.01 仍生成 finding。

影响文件：

- core/src/base_audit/systems/s2_report_collection/rule_engine.py
- core/tests/test_period_v2.py
- core/tests/test_version_manifest.py
- 版本管理/VERSION_MANIFEST.json

真实数据验收：

- 环比 400、外部核对 264、校验规则 3，合计 1088。
- 相比 v2.0.0 的 1096 条，减少 8 条浮点尾差导致的校验误报；剩余 3 条校验规则均为实际差异。

测试：

- S2-F01 定向测试：32 passed。
- core 全量 pytest：686 passed / 2 skipped / 0 failed。

## v2.0.0 — 2026-09-22

## v2.0.0 — 2026-09-22

变更类型：Breaking

前版本：v1.0.7
新版本：v2.0.0

修改内容：

- 报表审核清单固定输出“环比规则”“外部核对规则”“校验规则”三个工作表，停止生成旧的“审核清单”工作表。
- 删除最终 Excel 中的“审核类型”和“状态”列；审核正常使用空白“审核级别”表示，异常使用提示/关注/严重。
- 三类工作表使用独立表头，环比“变动率(%)”改为“变动幅度”，保留普通数值相对变动率与百分数百分点差的格式语义。
- 环比和外部核对增加“取数说明”；校验规则增加“涉及表单”“规则类型”“取值明细”，明确区分规则依赖与缺失原因。
- 未知审核类型和审核级别/内部状态冲突不再静默导出，改为明确失败并保留定位信息；校验规则涉及表单从全部相关记录去重、稳定排序。

影响文件：

- core/src/base_audit/systems/s2_report_collection/exporter.py
- core/src/base_audit/systems/s2_report_collection/models.py
- core/src/base_audit/systems/s2_report_collection/period_engine.py
- core/src/base_audit/systems/s2_report_collection/external_engine.py
- core/src/base_audit/systems/s2_report_collection/rule_engine.py
- core/src/base_audit/systems/s2_report_collection/validation.py
- core/tests/test_period_v2.py
- README.md
- tools/build_user_guide.py
- tools/build_system_operation_guides.py
- core/src/base_audit/config_guide.py
- core/frontend/web/index.html

配置影响：无。S2-CFG-01 保持 v1.0.1。

测试：

- Python compileall：通过。
- S2-F01/版本定向测试：39 passed。
- core 全量 pytest：685 passed / 2 skipped / 0 failed。
- 真实业务数据：环比 728、外部核对 357、校验规则 11，合计 1096；最终工作簿可由 openpyxl 重新打开。
- DOCX 已做定点内容更新；当前环境没有可用的 bundled LibreOffice/soffice，未完成 render_docx PNG 视觉验收。

GPT修改记录：

- GPT修改记录/20260922_111011_重构S2F01报表审核清单三Sheet输出.txt

Git提交：未提交

## v1.0.7 — 2026-09-21

变更类型：Bugfix

前版本：v1.0.6
新版本：v1.0.7

修改内容：

- Pywebview 源码启动前校验 `xlrd`，缺失时明确提示用 `requirements.txt` 安装完整依赖。
- PyInstaller 显式包含延迟导入的 `xlrd`，确保 S2 读取 `.xls` 的发行 EXE 不遗漏依赖。

影响文件：

- shell-pywebview/启动pywebview.bat
- shell-pywebview/build_exe.py

配置影响：无。

相关配置：S2-CFG-01 v1.0.0（未变）

测试：Python 3.12 中安装并导入 xlrd 2.0.2；pip check 通过；打包清单静态断言通过。

GPT修改记录：GPT修改记录/20260921_*.txt

Git提交：未提交

当前版本：v1.0.6

## v1.0.6 — 2026-09-20

变更类型：PATCH（契约澄清：外部数据收敛为 .xlsx 单后缀 + 明确报错；配套界面修复）

前版本：v1.0.5
新版本：v1.0.6

背景（UOS 2026-09-19 验收 DEFECT-4 / DEFECT-5）：
- 「报表采集·大集中数据」选择器口径为 .csv/.xlsx/.xls，但
  `import_external_workbook` 仅走 openpyxl，选 .csv 提交即报
  `InvalidFileException`——过滤器承诺与实现脱节；
- 浏览器内置文件对话框行元素 onclick 拼成 `browseToggleFile(this)(this)`
  （双重调用），多选对话框无法勾选。

用户定案（2026-09-20，取代同日早间「三后缀兼容」草案）：
- 外部核对依赖「参照表」工作表做机构映射：.xls（xlrd 链路）与 .csv
  （无工作表概念）都会丢「参照表」→ 外部数据收敛为**仅 .xlsx**，
  其余后缀明确报错；本期/上期目录维持 .xls/.xlsx 不变。

修改内容：
- importer：`_open_external` 前置后缀守卫——非 .xlsx 报
  「大集中数据仅支持 .xlsx 工作簿（.xls/.csv 会丢失「参照表」机构映射）」；
  读取路径与 v1.0.5 一致（openpyxl）；
- 过滤两侧同步收敛为仅 .xlsx：Tk 选择器 `*.xlsx`、浏览器内置 `report_data`
  模式 `(".xlsx",)`；
- 前端（共用件，不在 17 组件清单）：多选对话框确认时强制「同一批只能选择
  同一种后缀」（toast 提示、弹框保持打开）；修复 onclick 双重拼接
  （`const click=multi?'browseToggleFile':'browsePickOne'`，只补一次 `(this)`）。
  core/frontend 与 shell-flask/、shell-pywebview/ dist 快照同步。

影响文件：
- core/src/base_audit/systems/s2_report_collection/importer.py
- core/src/base_audit/web_app.py（Tk 过滤器与注释口径）
- core/src/base_audit/path_browser.py（report_data 模式收敛 .xlsx）
- core/frontend/web/index.html（+ shell-flask/、shell-pywebview/ dist 副本）

配置影响：
- 无

测试：
- tests/test_period_v2.py：外部数据 xlsx 正常解析（含参照表 org_map）；
  .xls/.csv/.txt 明确拒绝；缺「集中系统数据」工作表明确报错；
- 同轮完成黄金基线更新（`2026.08 比较文件.xls` 重生成为 2026-09-19 Windows
  当前产出冻结快照，VBA 旧件存档于 大集中系统工具/VBA基线存档/），
  test_central_comparison 表头断言升 22 列、黄金备注分布纳入断言；
- core 全量 651 passed / 1 skipped / 0 failed。

## v1.0.5 — 2026-09-19

变更类型：Config / PATCH（用户口径：审核说明文案简化）

前版本：v1.0.4
新版本：v1.0.5

修改内容（环比审核说明统一映射）：
- 本期与上期完全一致（含两侧均为 0、两侧均无数据、文字指标未变化、
  百分数 0 个百分点）→ 统一「无变动」，状态=正常、不进环比档位；
- 单侧存在（一侧有值另一侧空/零，含记录级缺失）→ 明确标识
  「本期有，上期无」/「本期无，上期有」，级别提示，环比不计算
  （变动率列为空），说明不带“未计算普通环比”解释尾巴；
- 增幅/降幅命中环比规则 → 保持规则描述不变。
- 此前的「本期缺失，上期存在」「上期缺失，本期存在」「本期、上期均为0」
  「上期为0、本期非0，未计算普通环比」「本期为0、上期非0」「本期、上期
  均无数据，未参与环比」「空值或非数字指标」「较上期持平」「文字指标
  未变化」九种文案收敛为三类。

影响文件：
- core/src/base_audit/systems/s2_report_collection/period_engine.py

配置影响：
- 无

测试：
- 新增审核说明文案映射测试（9 场景矩阵 + 单侧环比不计算断言）；
  边界档位测试适配（0 变动短路后用 0.0001 检验 B006 下边界）；
  core 全量 627 passed

Git提交：
- 见仓库日志

## v1.0.4 — 2026-09-19

变更类型：Config / PATCH（用户口径：两侧无数据不应标异常）

前版本：v1.0.3
新版本：v1.0.4

修改内容：
- 环比引擎：本期、上期都没有有效数值 → 状态=正常、说明
  「本期、上期均无数据，未参与环比」；单侧有值另一侧空 → 保留
  异常/关注（数据存在性问题）。

影响文件：
- core/src/base_audit/systems/s2_report_collection/period_engine.py

配置影响：
- 无

测试：
- 新增定向（两侧空=正常 / 单侧空=异常保留）；core 全量通过

Git提交：
- 见仓库日志

## v1.0.3 — 2026-09-19

变更类型：Bugfix / PATCH（用户报障：文件列表与对比对象 UI 不见）

前版本：v1.0.2
新版本：v1.0.3

排查结论：
- 源码无回归——用用户真实目录（2026.06/2026.03 季报，21 个 .xls）
  实测 list_period_pairs：21 对全部正常；
- 空白原因：配对清单只在「选择目录/程序启动」时计算，早前某次配对
  失败（如文件被占用）被 catch 清空后不再重试，旧进程界面保持空白。

修改内容：
- 执行报表规则完成后（_period_compare_v2_worker finally）自动刷新
  两期配对清单，旧空状态自愈；
- 前端摘要分类文案同步「校验规则」（v1.0.2 改名遗漏处）。

配置影响：
- 无

测试：
- test_period_v2 + test_web_app 56 passed；core 全量通过

GPT修改记录：
- GPT修改记录/20260919_124500_两期配对清单自愈刷新.txt

Git提交：
- 见仓库日志

## v1.0.2 — 2026-09-19

变更类型：Bugfix / PATCH（用户报障：校验规则取不到数据 全是 0）

前版本：v1.0.1
新版本：v1.0.2

修改内容：
- 根因：规则引擎不支持旧 VBA 八段占位符语法。真实 2.配置的校验规则是
  `[,,,20203045,]`（逗号分段，第 4 段才是指标代码），引擎把整段
  `,,,20203045,` 当 lookup key → 永远 miss → 全部表达式规则
  “缺少指标”，计算过程呈现 `0.0 < 0.055 * 0.0`。新增 _indicator_code
  八段解析（与 V1 period_compare 同口径），兼容简写 `[20201028]`。
- 缺失/空值保护：指标行存在但值为空（空报）时与缺失同等处理进
  「规则无法执行」，禁止拿 0 参与比较——否则“充足率低于监管线”类
  规则会把空报机构全部误判成异常。
- 呈现修正：规则无法执行行（原「数据质量」）改审核类型「校验规则」
  （用户要求，validation/regression 同步），审核说明写真实原因
  （缺少哪些指标的数据 + 规则描述括注），不再用规则描述冒充
  “低于监管线 异常”。

影响文件：
- core/src/base_audit/systems/s2_report_collection/rule_engine.py
- core/src/base_audit/systems/s2_report_collection/validation.py
- core/src/base_audit/systems/s2_report_collection/regression.py

配置影响：
- 无（2.报表采集系统_配置.xlsx 零修改）

测试：
- 新增 2 个定向：八段占位符取数（数据齐全不误报/真低于线命中/
  缺指标出校验规则行）；缺指标与空值的审核说明断言
- core 全量：625 passed / 0 failed
- 真实配置三场景复现：数据齐全不误报、真低于线命中
  （10.0 < 0.055 * 2000.0）、缺指标显示“规则无法执行：缺少指标
  20203045 的当期数据”

GPT修改记录：
- GPT修改记录/20260919_123000_校验规则八段占位符与呈现修复.txt

Git提交：
- 见仓库日志

## v1.0.1 — 2026-09-19

变更类型：Refactor / PATCH

前版本：v1.0.0
新版本：v1.0.1

修改内容：
- 源码路径归位：core/src/base_audit/period_v2 → core/src/base_audit/systems/s2_report_collection
- 去除包名中的历史版本后缀 v2

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
- 报表采集系统唯一功能：两期报送数据导入、环比计算、Bxxx 警戒档、校验规则与外部核对，输出报表审核清单。

当前实现：
- `core/src/base_audit/period_v2`

相关配置：
- S2-CFG-01（版本见各自 CHANGELOG）

测试基线：
- 以建立版本管理体系当日 core 全量 pytest 通过状态为基线（详见当次
  GPT修改记录）。
