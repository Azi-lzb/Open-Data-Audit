# S3-F05 比较文件到金融表单

当前版本：v1.1.0

## v1.1.0 — 2026-09-24

变更类型：Feature / MINOR

前版本：v1.0.2
新版本：v1.1.0

修改内容：
- 新增 3.3“检查配置”入口，先核查“转表设置”“报表清单”及运行时必需表头，再检查设置值、单位只允许元/万元/亿元、内嵌表单模板和报表代码引用。
- 联同 3.0 通用配置只读预检展示错误与提示；不逐格检查灵活的内嵌表单内容，不改变业务结果。

配置影响：S3-CFG-03 保持原版本。

Git提交：未提交。

## v1.0.2 — 2026-09-23

变更类型：Bugfix / PATCH

前版本：v1.0.1
新版本：v1.0.2

修改内容：

- 复用 S3-CFG-00 新的百分数数值警戒区间匹配：`99` 表示 99%；
- 表单着色在正式迁移后的配置下保持原有命中结论不变。

相关配置：

- S3-CFG-00 v1.0.2 → v2.0.0。

## v1.0.1 — 2026-09-19

变更类型：Refactor / PATCH

前版本：v1.0.0
新版本：v1.0.1

修改内容：
- 源码路径归位：core/src/base_audit/central_statistics → core/src/base_audit/systems/s3_central_statistics
- expression_v2.py 更名为 expression_lae.py（去版本化命名）

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
- 大集中统计系统：按机构回填金融表单并应用警戒颜色（FAST_OOXML 为固定默认渲染路径）。

当前实现：
- `core/src/base_audit/central_statistics/form_template.py`

相关配置：
- S3-CFG-00、S3-CFG-01、S3-CFG-02、S3-CFG-03（版本见各自 CHANGELOG）

测试基线：
- 以建立版本管理体系当日 core 全量 pytest 通过状态为基线（详见当次
  GPT修改记录）。
