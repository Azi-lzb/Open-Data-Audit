# V3 项目开发总纲

本文件是 Codex/Claude 共用的唯一项目入口。`CLAUDE.md` 只指向本文件，禁止维护第二份规则副本。
开始任务先读本文件，再按任务类型只读对应的 `rules/`；不要默认加载全部规则。

## 项目与架构

- V3 服务 Windows 与 UOS/Linux，包含 S1 逐笔统计、S2 报表采集、S3 大集中统计。
- 只有一套业务核心：`core/`。`shell-flask/`、`shell-pywebview/` 是共用 core 的外壳；平台差异仅放在适配层和启动层。
- 技术顺序：原生 Python → 适用时 Direct OOXML → 通用静态读写用 openpyxl → 仅在真实计算或不可替代原生语义时使用 Office。openpyxl/OOXML 不计算公式。
- 配置检查只读；启用规则必须成功进入执行链或明确失败；迁移默认保留 VBA/现有配置业务语义，数量差异必须可解释；原始业务文件不得覆盖。
- 用户错误要说明对象、原因和处理方式；底层异常进入技术日志。

## 目录边界

`core/` 是唯一业务实现；`config/` 只保存公开空表头配置；本机真实配置放在不入库的 `config-real/`；`rules/` 保存按任务加载的细则；`GPT修改记录/` 保存修改记录。不得在历史打包副本或平台外壳中另建业务实现。

## 按需读取规则

| 任务 | 必读规则 |
|---|---|
| 项目范围、目录、模块归属 | `rules/01_项目范围与目录.md` |
| 架构、平台适配、技术选型 | `rules/02_核心架构与跨平台.md` |
| Excel/WPS/LibreOffice、OOXML、openpyxl | `rules/03_Excel与OOXML边界.md` |
| 模板、命名区域、DAG、数据边界 | `rules/04_模板工作流与数据边界.md` |
| UI、Bridge API、用户提示 | `rules/05_UI与接口.md` |
| VBA 行为复刻、配置检查、规则执行性 | `rules/06_VBA复刻与配置复核.md` |
| 新 VBA/旧配置迁移 | `rules/10_VBA与配置迁移规范.md` |
| 测试、真实数据、发布验收 | `rules/07_测试与发布.md` |
| 子 Agent 委派、WIP、修改记录 | `rules/08_AI协作与修改记录.md` |
| 平台实现参考（仅相关时） | `rules/09_技术实现参考.md` |
| 组件编号、版本升级 | `rules/11_版本管理与组件编号.md` |

任务执行顺序：识别 Sx-Fxx/Sx-CFG-xx → 检查 Git/WIP → 读相关规则与当前实现 → 修改 → 按影响范围验证 → 更新版本/CHANGELOG（适用时）→ 新增修改记录 → 报告结果。功能调整必须评估是否需新增/更新测试；按 `core/tests/TEST_GROUPS.md` 回归受影响系统。

复杂任务由主模型先查证实际结构、确定方案和验收标准；再按需要把文件边界清楚、可独立验收的工作委派给子 Agent。优先使用 `gpt-6-luna` + `xhigh`（不可用时取该模型支持的最高档）；只有并行收益高于委派开销时才拆分。主模型审查全部 diff 并负责最终集成验收，子 Agent 不获得额外 Git 权限。细节见 `rules/08_AI协作与修改记录.md`。

## 版本与留痕

- 产品名 V3；3.0/3.1/3.2/3.3 等历史编号不是版本号。版本唯一依据为 `版本管理/VERSION_MANIFEST.json`；功能版本与配置版本独立升级。
- 每次实际修改都要在 `GPT修改记录/` 新增时间戳记录；版本变更按 `rules/11` 同步 MANIFEST 与受影响组件 CHANGELOG。详细格式只以对应规则为准。

## 双仓库 Git 约定

- 本仓库的 `origin` 必须是公开代码仓库 `https://github.com/Azi-lzb/Open-Data-Audit.git`。这里使用独立的新 Git 历史，以 `main` 作为共同代码基线。
- `https://github.com/Azi-lzb/Close-Data-Audit.git` 是旧项目的私有历史仓库，保留原有 Git 历史、配置和测试数据；今后默认只读。不要将 Close 的分支、提交或历史合并、推送到 Open，也不要把它设为 Open 的上游或镜像。
- 真实配置、原始业务文件、测试数据、验收产出和构建产物禁止提交到 Open。公开 `config/` 仅含空表头工作簿；本机正式配置放在被忽略的 `config-real/`。`.gitignore` 无法保护已跟踪的空表头工作簿，提交前运行 `python tools/check_public_snapshot.py` 检查暂存内容。
- 两个仓库需要同步时，在各自工作目录分别核对 `git remote -v`、`git status --short`，只与各自远端同步。另一台机器开始开发前先 `git fetch origin` 更新 Open 的工作目录。

未经用户明确要求，不得 commit、push、merge、rebase、创建 PR 或修改 `main`；禁止强推。用户已确认的公开仓库首次代码快照按该次确认范围执行。后续功能修改先列明拟提交文件、验证结果和风险，等待用户确认提交范围。

## 冲突优先级

用户当前明确要求 > 本文件硬约束 > 对应 `rules/` > 当前实现与测试证据 > 历史文档。发现业务口径冲突时先报告，不自行猜测。
