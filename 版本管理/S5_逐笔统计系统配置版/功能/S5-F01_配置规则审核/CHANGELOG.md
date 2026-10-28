# S5-F01 汇总核查表校验

当前版本：v3.0.1

## v3.0.1 — 2026-10-22 21:10

变更类型：性能优化 / PATCH

前版本：v3.0.0；新版本：v3.0.1。

- 缓存只读审核表达式 AST、取值字典解析和字典依赖检查结果；按业务、行列路径、工作表及行表头为模板指标建立索引，减少多机构审核中的重复解析与全目录扫描。
- 审核范围、规则、表达式语义、取数值、错误状态和结果列不变。
- 同一组 16 个本地样例的旧/新路径 A/B：规则审核 19.92 → 13.56 秒（约快 32%）；整批 24.08 → 17.79 秒（约快 26%）。14 个可识别文件共 5,279 条结果的逐条签名一致。
- 测试：`python -X utf8 -m pytest -q -rs -m s5 core/tests`：339 passed / 884 deselected / 6 warnings / 3 subtests passed。6 条 warning 为测试环境缺少 python-calamine 后按预期回退 openpyxl。

GPT修改记录：`GPT修改记录/20261022_211051_S5-F1当前业务审核性能优化.txt`。

## v3.0.0 — 2026-10-22 20:37

变更类型：Breaking / MAJOR

前版本：v2.7.1；新版本：v3.0.0。

- F1 每次只在“执行结果”目录直接生成一份带时间戳的汇总 XLSX；不再创建 `batch_*`、单机构子目录，不再写单机构结果、审核副本或运行总账 JSON。
- 结果中的定位链接直接指向原始源文件；F1 不计算输入文件内容哈希，因此不再校验文件在运行期间是否被外部修改。
- 运行记录显示配置加载、机构参照、文件识别、规则审核、模板公式扫描、条件格式扫描、结果导出和总耗时。
- S5-F1 返回值不再包含 `ledger_path` 或单机构 `result_path`；配置与审核规则语义不变。
- 主面板不再重复显示 F1 结束统计卡片；配置检查主面板改为简短状态，完整明细仍保存在可打开的检查报告中。

影响文件：S5-F1 服务、批量执行器、通用审核引擎的可选完整性检查、结果链接生成、运行记录及前端结果入口。

配置影响：无；S5-CFG-01 v2.0.0 未变。

测试：`python -X utf8 -m pytest -q -rs -m s5 tests`：331 passed / 6 skipped / 0 failed，3 subtests passed；6 项需 Flask 的测试按环境条件跳过，6 条 Calamine 缺包回退告警符合预期。

GPT修改记录：`GPT修改记录/20261022_203700_S5F1耗时与单工作簿输出.txt`。

Git提交：未提交。

## v2.7.1 — 2026-10-22 00:27

变更类型：Bugfix / PATCH

前版本：v2.7.0；新版本：v2.7.1。

- 配置检查对模板路径的提示和统计统一称为“模板取数地址”，不再显示 V_PRE，避免用户误以为需要选择尚未提供的上期源数据目录。表达式执行和 V_PRE 取数能力不变。

影响文件：`core/src/base_audit/systems/s5_configured_statistics/service.py`、相关 S5 测试与版本清单。

配置影响：S5-CFG-01 v2.0.0 未变。

测试：`python -X utf8 -m pytest -q -rs -m s5 core/tests`：326 passed / 6 skipped / 0 failed / 865 deselected，3 subtests passed；6 项因缺少 Flask 跳过，8 条 Calamine 缺包告警触发预期的 openpyxl 回退。

GPT修改记录：`GPT修改记录/20261022_002708_S1筛选与S5检查文案.txt`。

Git提交：未提交。

## v2.7.0 — 2026-10-22 00:19

变更类型：Feature / MINOR

前版本：v2.6.0；新版本：v2.7.0。

- 配置检查拆分为“检查当前业务”和“检查全量”两种范围。当前业务检查从当前匹配模板识别业务，只加载并核对该业务启用规则；全量模式一次解析启用规则，再逐业务寻找唯一匹配模板并检查地址。
- 全量报告按业务列出模板匹配和地址覆盖情况；模板缺失、匹配歧义或动态地址未展开均标为未检查，报告显示不完整，避免把未覆盖范围当作通过。
- 全量模式复用一次配置解析与规则引用提取，外部地址仅扫描一次；进度按业务更新。

影响文件：`core/src/base_audit/systems/s5_configured_statistics/{service.py,web_api.py}`、`core/src/s1audit_configured/workbook_config.py`、`core/frontend/web/index.html`、相关 S5 测试与版本清单。

配置影响：S5-CFG-01 v2.0.0 未变；S1/S2/S3 业务逻辑和配置未修改。

测试：`python -X utf8 -m pytest -q -rs -m s5 core/tests`：325 passed / 6 skipped / 0 failed / 865 deselected，3 subtests passed，72.53 秒；版本清单测试 8 passed，Python 编译和 `git diff --check` 通过。6 项跳过为缺少 Flask 的外壳路由测试；8 条 Calamine 缺包告警触发预期的 openpyxl 回退。真实配置全量检查约 10.5 秒，19 个业务模板均唯一匹配，2,392 条启用规则、外部地址 43/43 定位；模板地址 4,151/4,215 成功，64 项未找到，集中在单位存款、票据业务、股权和特定目的载体投资，结果按配置问题标记而非误报通过。

GPT修改记录：`GPT修改记录/20261022_001931_S5双模式配置检查.txt`。

Git提交：未提交。

## v2.6.0 — 2026-10-21 23:54

变更类型：Feature / MINOR

前版本：v2.5.0；新版本：v2.6.0。

- 待处理文件匹配按源工作簿工作表覆盖判断；模板里未出现在源文件中的表不再降级为部分匹配。
- 模板之外的源工作表单独显示为“含模板外工作表”；只有完全没有同名数据表时才标为未匹配。
- 工作表匹配数量以源工作表为分母展示，避免模板保留的其他表造成低比例误读。
- 自动识别仍保留候选相近时的歧义保护。

影响文件：`core/src/base_audit/discovery.py`、`core/src/base_audit/systems/s5_configured_statistics/recognition.py`、`core/frontend/web/index.html`、相关 S5 测试。

配置影响：S5-CFG-01 v2.0.0 未变；S1/S2/S3 未修改。

测试：`python -X utf8 -m pytest -q -rs -m s5 tests`：322 passed / 6 skipped / 0 failed，3 subtests passed；8 条 Calamine fallback warning 为测试环境缺少 python-calamine 时的预期回退。

GPT修改记录：`GPT修改记录/20261021_235428_S1S5待处理文件匹配状态优化.txt`。

Git提交：未提交。

## v2.5.0 — 2026-10-21 23:31

变更类型：Feature / MINOR

前版本：v2.4.1；新版本：v2.5.0。

- S5 页面打开、选择目录/配置和开始执行时不再自动检查配置或识别模板；“检查配置”和“识别模板”均由用户手动触发，手动选择模板后可保留该模板识别业务与报告期。
- 待处理文件显示模板匹配状态、业务类型和报告日期；增加匹配、部分匹配、业务类型和日期的组合筛选。筛选只控制当前显示，隐藏文件原有勾选会保留，全选/取消作用于可见项。
- 输出默认改为源数据目录下的“执行结果”；F1 批次直接在该目录生成，不再多建 `S5-F01` 子目录。旧默认路径 `S5_执行结果` 会自动迁移到同级“执行结果”，固定输出目录不受影响。

影响文件：`core/frontend/web/index.html`、`core/src/base_audit/systems/s5_configured_statistics/{web_api.py,recognition.py,service.py}`、相关 S5 测试与说明。

配置影响：S5-CFG-01 v2.0.0 未变；S1/S2/S3 未修改。

测试：`python -X utf8 -m pytest -q -rs -m s5 core/tests`：317 passed / 6 skipped / 0 failed，3 subtests passed；6 项因未安装 Flask 跳过。8 条 Calamine 回退 warning 为测试环境缺包时的预期行为。`compileall` 与 `git diff --check` 通过。

GPT修改记录：`GPT修改记录/20261021_233119_S5手动检查与待处理文件筛选.txt`。

Git提交：未提交。

## v2.4.1 — 2026-10-21 23:13

变更类型：Bugfix / PATCH

前版本：v2.4.0；新版本：v2.4.1。

- 修复大配置检查在规则解析后长时间停留于 58%：规则错误定位索引改为只读顺序扫描，增加索引与地址引用提取进度阶段。
- 模板和外部文件地址扫描共用一次规则引用解析，避免重复解析启用规则中的字典和表达式。
- 首次进入 S5 仍自动检查并显示弹窗；同一页面会话已检查过的配置不再因每次切换导航重复自动检查，重新选择配置仍会检查。
- 真实配置只读验证：2,392 条规则全部启用，表达式/字典检查通过；未提供模板和外部文件路径时，这两项保持“未检查”。检查总耗时约 4.2 秒。

影响文件：`core/src/base_audit/systems/s5_configured_statistics/service.py`、`core/frontend/web/index.html`、相关测试。

配置影响：S5-CFG-01 v2.0.0 未变；未修改 `config-real/`。

测试：S5 回归 311 passed / 6 skipped / 0 failed / 3 subtests passed；配置检查与前端自动检查定向测试 9 passed。S5 回归中的 8 条 Calamine fallback warning 由测试场景预期触发。

GPT修改记录：`GPT修改记录/20261021_231346_S5配置检查进度与性能优化.txt`。

Git提交：未提交。

## v2.4.0 — 2026-10-21

变更类型：Feature / MINOR

前版本：v2.3.2；新版本：v2.4.0。

- “检查配置”改为后台任务，显示实时阶段进度与逐项结果，并生成可打开的时间戳检查报告。
- 检查报告写入通用“执行结果”目录，与 F1 审核输出目录分开。
- 表达式解析、取值字典键引用和模板/外部文件地址只检查启用规则；禁用规则跳过。没有启用规则时标记未检查，不会显示为完整通过。
- 模板地址检查继续使用运行时定位器；共享配置加载器仅新增可选进度回调，不改变其他调用者行为。

影响文件：`core/frontend/web/index.html`、`core/src/base_audit/systems/s5_configured_statistics/service.py`、`core/src/base_audit/systems/s5_configured_statistics/web_api.py`、`core/src/s1audit_configured/workbook_config.py`、相关测试与日志。

配置影响：S5-CFG-01 v2.0.0 未变；不修改 `config-real/` 中的正式配置。

测试：S5 回归 310 passed / 6 skipped / 0 failed / 3 subtests passed；S1 回归 498 passed / 13 skipped / 0 failed / 3 subtests passed；版本清单和更新日志测试 9 passed；异步检查与禁用规则定向测试 2 passed。Python `compileall`、前端 `node --check` 与 `git diff --check` 通过。未打包、未提交。

GPT修改记录：`GPT修改记录/20261021_230053_S5配置检查后台进度与报告.txt`。

Git提交：未提交。

## v2.3.2 — 2026-10-21

变更类型：Bugfix / PATCH

前版本：v2.3.1；新版本：v2.3.2。

- S5“检查配置”和“自动识别模板”按钮改用与 S1 相同的紧凑蓝色样式，并紧跟字段标题左侧排列；按钮尺寸、颜色和位置与 S1 对齐。

影响文件：`core/frontend/web/index.html`、`core/tests/s5/test_s5_integration.py`。

配置影响：无；S5-F01 执行及配置读取语义不变。

测试：S5 界面路径和按钮样式、五连击隐藏行为、交付版前端显示及版本清单定向测试 11 passed；`git diff --check` 通过。未打包、未提交。

GPT修改记录：`GPT修改记录/20261021_222925_S5按钮尺寸颜色位置对齐.txt`。

Git提交：未提交。

## v2.3.1 — 2026-10-21

变更类型：Bugfix / PATCH

前版本：v2.3.0；新版本：v2.3.1。

- S5 主面板和维护工具的路径标签统一使用 S1 的可点击蓝色链接样式。
- 默认审核配置优先使用 `5.逐笔统计系统_审核配置.xlsx`，兼容回退旧文件名；发行目录中保存的旧配置路径自动迁移到新文件。
- 当共享 `模板文件_命名区域` 存在时，将发行目录中保存的默认 `S5_模板文件` 自动迁移到该目录；用户自定义路径保持原样。
- S1/S2 交付包排除改名后的 S5 配置文件，完整版内部配置包继续带上重命名配置和共享命名区域模板目录。

影响文件：`core/src/base_audit/systems/s5_configured_statistics/web_api.py`、`core/frontend/web/index.html`、`shell-flask/distribution-profiles.json`、相关测试。

配置影响：S5-CFG-01 v2.0.0 未变；只调整本地配置文件的默认路径识别，不修改工作簿内容。

测试：`python -m pytest -q -rs -m s5 tests`：303 passed、6 skipped、865 deselected、3 subtests passed、0 failed；发行配置与版本清单测试 32 passed；`git diff --check` 和 JSON 清单解析通过。6 项跳过是缺少 Flask 的安装测试，8 条读取告警为可选 python-calamine 未安装后的预期回退。未打包，未做目标桌面实机验收。

GPT修改记录：`GPT修改记录/20261021_222539_S5界面颜色与重命名配置路径.txt`。

Git提交：未提交。

## v2.3.0 — 2026-10-21

变更类型：Feature / MINOR

前版本：v2.2.2；新版本：v2.3.0。

- 首次启动时，模板目录优先默认到共享的“模板文件_命名区域”，外部文件默认到随包大集中外部参照工作簿。
- 用户已选且仍存在的自定义路径继续优先使用；路径失效时回退到发行配置目录中的默认项。
- S5 审核和配置内容不变；S5-CFG-01 v2.0.0 未变。

影响文件：`core/src/base_audit/systems/s5_configured_statistics/web_api.py`、`core/tests/test_web_app.py`、发行打包脚本。

验证：S5 回归 301 passed、6 skipped、3 subtests passed、0 failed；路径与发行配置定向测试 85 passed、3 subtests passed。pytest 有可选 python-calamine 未安装并回退 openpyxl 的提示。未重新打包。

GPT修改记录：`GPT修改记录/20261021_221251_S1S2S5默认外部文件与命名区域模板路径.txt`。

Git提交：未提交。

## v2.2.2 — 2026-10-21

变更类型：Bugfix / PATCH

前版本：v2.2.1；新版本：v2.2.2。

- 将 F1 汇总核查表校验与隐藏的 F2、F3、F4 操作按钮放入同一张卡片，按钮按纵向排列。
- 移除维护功能标题/解锁提示及空闲状态下的说明卡片；F2-F4 仍默认隐藏，五击左侧 S5 导航的显示/隐藏行为不变。
- 审核状态与结果打开入口仅在存在实际任务结果时显示，并放在同一按钮卡片内；空闲时不显示占位说明。

影响文件：`core/frontend/web/index.html`、`core/tests/s5/test_s5_integration.py`、`core/tests/TEST_GROUPS.md`。

配置影响：无；F1-F4 的执行逻辑不变。

测试：`test_s5_maintenance_tools_are_hidden_by_default_and_toggle_with_five_nav_clicks` 1 passed；`python -m pytest -q -rs -m s5 core/tests`：299 passed / 6 skipped / 0 failed / 3 subtests passed（865 deselected）；版本清单 8 passed；`git diff --check` 通过。pytest 有 8 条可选 python-calamine 未安装并回退 openpyxl 的 RuntimeWarning；浏览器截图及桌面视觉验收未执行。

GPT修改记录：`GPT修改记录/20261021_214855_S5四功能按钮合一卡片.txt`。

Git提交：未提交。

## v2.2.1 — 2026-10-21

变更类型：Bugfix / PATCH

前版本：v2.2.0；新版本：v2.2.1。

- 对空行名称行按指标列与整行状态分别判断：指标有值时提示行名称缺失；日期、报表代码、行顺序号等元数据存在但指标为空时忽略；仅整行全空的内部行给出软提示。
- 普通报表和时序布局采用相同的空行口径，避免元数据列造成“行名称及指标数据均为空”的误报。

影响文件：`core/src/s1audit_configured/external_layout.py`、`core/tests/s5/test_s5_integration.py`、`core/tests/TEST_GROUPS.md`。

配置影响：无；数据质量提示仍按原方式输出。

测试：`test_external_layout_anomalies_warn_and_continue_for_both_layouts` 八种布局/值/输出引擎组合 8 passed；`python -m pytest -q -rs -m s5 core/tests`：299 passed / 6 skipped / 0 failed / 3 subtests passed（865 deselected）；版本清单 8 passed；compileall 与 `git diff --check` 通过。pytest 有 8 条可选 python-calamine 未安装并回退 openpyxl 的 RuntimeWarning；真实业务数据及平台真机未执行。

GPT修改记录：`GPT修改记录/20261021_213956_S2S5外部空标签行误报修正.txt`。

Git提交：未提交。

## v2.2.0 — 2026-10-21

变更类型：Feature / MINOR

前版本：v2.1.1；新版本：v2.2.0。

- S5 外部取数统一按表头识别普通报表和“数据日期/行名称”时序布局，即使配置工作簿包含固定数据源契约也不会绕开布局识别；时序值仍只匹配本期完整报告期。
- 对已匹配的数值指标，空白或非数值内容按 0 继续参与规则计算，并记录机构、文件、工作表、单元格和提示；有值但行名称为空的行跳过并提示检查误删，空白行忽略并软提示。
- 时序目标行的数据日期空白或无效时提示并忽略该行；若没有本期有效记录，仅对应规则报取数失败，批次继续处理其他规则和机构。重复目标、重复指标表头及不完整布局仍明确报错。
- 批次和单机构结果在存在异常提示时增加“数据质量提示”工作表；任务结束在运行记录中提示数量并列出定位信息。

配置影响：S5-CFG-01 不变。

测试：定向共用测试 159 passed；异常处理/输出引擎和无本期有效日期时继续批次处理的定向测试 11 passed；`python -m pytest -q -rs -m s5 core/tests`：299 passed / 6 skipped / 0 failed / 3 subtests passed；版本清单 8 passed；compileall 与 `git diff --check` 通过。平台真机及真实业务数据未执行。

GPT修改记录：`GPT修改记录/20261021_210609_S2S5外部异常数据继续处理与最终提示.txt`。

Git提交：未提交。

## v2.1.1 — 2026-10-21

变更类型：Bugfix / PATCH
前版本：v2.1.0；新版本：v2.1.1。

- 三个配置维护入口移至 F1 主按钮下方，名称为“F2 模板 V_CUR 表达式”“F3 外部文件 V_EXT 表达式”“F4 模板校验公式转配置”。
- 修复显示开关状态：每次启动默认隐藏，连续点击左侧 S5 导航五次显示或隐藏；隐藏时退出维护页并返回 F1。
- 显式隐藏 S5 内的 `.card[hidden]`，避免卡片默认 `display:flex` 覆盖隐藏状态。
- 显示状态仅保存在当前运行会话，不沿用上次启动时的浏览器存储状态。

配置影响：S5-CFG-01 不变。
测试：共享前端、S5 集成及版本清单 47 passed；8 个既有读取器告警为未安装可选 Calamine 时回退 openpyxl。
GPT修改记录：GPT修改记录/20261021_170941_S5维护功能编号与隐藏切换修复.txt。
Git提交：未提交。

## v2.1.0 — 2026-10-21

变更类型：Feature / MINOR
前版本：v2.0.0；新版本：v2.1.0。

- 在 S5 内新增三个隐藏的配置维护入口：V_CUR/V_PRE 表达式参考、V_EXT 外部表达式、校验公式转配置候选；连续点击 S5 导航 5 次（2 秒内）显示或隐藏。
- 三个维护工具统一沿用 S5 面板当前匹配模板和大集中系统数据；V_EXT 与 S5-F1 按表头识别普通报表及“数据日期/行名称”时序格式，并按本期报告期精确选取时序记录。
- 维护工具输出独立保存；公式候选默认禁用，必须人工筛选后再复制到正式 S5 配置。
- 输入工作簿保持只读；S1/S2/S3 原功能与配置未改。

配置影响：S5-CFG-01 不变。
测试：S5 组 284 passed / 6 skipped；S1 组 492 passed / 13 skipped；外部布局定向 12 passed；Python 编译、版本清单、前端 JavaScript 语法及差异检查通过。
GPT修改记录：GPT修改记录/20261021_125517_S5隐藏配置维护工具.txt、GPT修改记录/20261021_131134_S5维护工具沿用主面板路径.txt。
Git提交：未提交。

## v2.0.0 — 2026-10-21

变更类型：Breaking / MAJOR
前版本：v1.1.1；新版本：v2.0.0。

- 功能名称改为“汇总核查表校验”，S5 仅保留 F1；执行仍使用 S5 独立审核规则引擎，不调用 S1-F01 DAG。
- 移除 S5-F02 模板公式审核入口及适配代码；S1 原有工作流未修改。
- 配置工作簿影响见 S5-CFG-01 v2.0.0。

配置影响：S5-CFG-01 v1.2.0 → v2.0.0。
测试：未运行自动化测试；配置工作簿读回确认、静态引用检查及差异检查完成。
GPT修改记录：GPT修改记录/20261021_124319_S5单功能收敛与配置精简.txt。
Git提交：未提交。

## v1.1.1 — 2026-10-21

变更类型：Bugfix / PATCH
前版本：v1.1.0；新版本：v1.1.1。

- 将“大集中系统数据”与“匹配模板”在第二行的显示位置调整为与 S1 一致；执行参数和匹配逻辑不变。

配置影响：S5-CFG-01 v1.1.0 → v1.2.0，增加两个可选规则字段并导入待筛选候选。
测试：未运行自动化测试；配置导入后直接读取确认 2,392 条候选均为禁用。
GPT修改记录：GPT修改记录/20261021_123222_S5配置候选导入与路径布局.txt。
Git提交：未提交。

## v1.1.0 — 2026-10-21

变更类型：Feature / MINOR
前版本：v1.0.0；新版本：v1.1.0。

- S5 六框顺序与 S1 对齐；模板从目录自动匹配，业务和报告期自动识别，歧义明确提示。
- 移除业务/日期/单独 F2 配置框；S5 统一配置增加原 S1 格式工作表，原 S1 配置只读。
- 保留显式参数及 V_PRE 服务能力，未改原 S1/S2/S3/S4 业务实现。

配置影响：S5-CFG-01 v1.0.0 → v1.1.0，新增兼容工作表；F1 原审核规则保留。
测试：定向106 passed；真实 Office 对拍、Python3.8.10 完整自动识别审核、浏览器及配置样式保留验证通过。未跑全量或发布。
GPT修改记录：GPT修改记录/20261021_104142_S5六框布局与自动识别统一配置.txt。
Git提交：未提交。

## v1.0.0 — 2026-10-21

变更类型：Baseline / Feature

首次登记独立 S5 接入；既有 S1/S2/S3/S4 业务实现与配置未改动。

配置影响：新增独立 S5-CFG-01 v1.0.0；旧配置不变。
相关配置：S5-CFG-01 v1.0.0；F2 使用原 S1 格式配置的独立副本。
测试：定向 91 passed；真实 Office 对拍及快照复验 2 passed；Python3.8.10 完整 F1 合成审核通过。本轮未跑全量或发布。
GPT修改记录：GPT修改记录/20261021_095124_S5独立接入与迁移.txt。
Git提交：未提交。
