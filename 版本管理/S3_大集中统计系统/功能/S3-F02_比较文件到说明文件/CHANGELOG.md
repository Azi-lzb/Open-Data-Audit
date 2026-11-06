# S3-F02 比较文件到说明文件

当前版本：v3.0.0

## v3.0.0 — 2026-11-06 12:53

变更类型：Breaking / MAJOR

前版本：v2.0.0
新版本：v3.0.0

修改内容：
- 月报工作表名、标题、8列表头、宋体 11 号、纵向居中、列宽及机构代码/名称合并方式按 2026.10 VBA 月报样例输出；不再额外生成“生成信息”工作表。
- 月报与日报均输出原机构代码和不带地区后缀的机构名称；日报仅在机构名称后新增“地区代码”“地区名称”两列，并保留筛选。
- 月报金额列按 `0.00` 显示；表头包含单位，单位从比较文件增减额表头取得。

影响文件：
- core/src/base_audit/systems/s3_central_statistics/explanation_exporter.py
- core/tests/test_central_explanation_exporter.py

配置影响：无。

验收：对比 `2026.10 说明文件.xls`，工作表名、标题、8 个表头、字体、对齐、金额格式、A–E 列宽、标题合并、筛选和冻结设置一致；当前程序 101 条、VBA 样例 78 条，行数差异来自 F1 规则标记口径，未宣称内容逐行一致。

GPT修改记录：
- GPT修改记录/20261106_115644_S3月报日报说明与对比字段维护.txt

Git提交：未提交。

## v2.0.0 — 2026-11-06

变更类型：Breaking / MAJOR

前版本：v1.1.0
新版本：v2.0.0

修改内容：
- 说明文件可选择月报或日报格式。月报沿用原 8 列；日报在机构名称后增加地区代码、地区名称两列。
- 两种格式的机构代码都保持原始机构代码，不再拼接地区代码；金额列仍仅显示两位小数。
- 指标名称默认保留原始空白与缩进，可在高级设置中选择 trim；结果可选完成后自动打开，默认关闭。

影响文件：
- core/src/base_audit/systems/s3_central_statistics/explanation_exporter.py
- core/src/base_audit/web_app.py
- core/frontend/web/index.html

配置影响：仅新增用户设置；业务配置工作簿未修改。

兼容性：月报保持既有 8 列结构，但机构代码不再追加地区码；旧版长机构代码反馈与新月报之间的兼容性需要实际样例确认。日报反馈支持地区码参与匹配。

测试：本轮未运行测试。

GPT修改记录：
- GPT修改记录/20261106_115644_S3月报日报说明与对比字段维护.txt

Git提交：未提交。

## v1.1.0 — 2026-10-13 21:54

变更类型：Feature / MINOR

前版本：v1.0.1
新版本：v1.1.0

修改内容：
- 本条为汇总追溯 `40875d7..1b0ae87` 的已完成调整，并按当前确认口径登记。
- 导出机构代码时，村镇银行始终拼接地区代码；其他机构仅在同机构覆盖多个地区时拼接，单地区机构保留原码。
- 说明文件的数据值、上期值和增减额显示两位小数，仅改变显示格式，不改写单元格数值。

影响文件：
- core/src/base_audit/systems/s3_central_statistics/explanation_exporter.py
- core/tests/test_central_explanation_exporter.py

配置影响：业务配置未修改；相关配置版本为 S3-CFG-00 v2.0.2、S3-CFG-01 v3.0.3、S3-CFG-02 v1.0.3、S3-CFG-03 v2.0.0。

测试与限制：
- 历史验证见 GPT修改记录/20261007_193703_说明文件多地区机构代码拼地区码.txt（说明导出测试 8 passed、S3 组 437 passed）；结果为既有记录，本次未重跑。
- 两位小数显示的历史记录见 GPT修改记录/20261007_192710_说明与核对金额两位显示及E-14判等阈值.txt。
- 本次未重打包，未做 Office 真机验收。

GPT修改记录：
- GPT修改记录/20261013_215438_统一3.3写入口径并补齐发布组件版本.txt

Git提交：
- 未提交。

## v1.0.1 — 2026-09-19

变更类型：Refactor / PATCH

前版本：v1.0.0
新版本：v1.0.1

修改内容：
- 源码路径归位：core/src/base_audit/central_statistics → core/src/base_audit/systems/s3_central_statistics
- 实现登记路径更新

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
- 大集中统计系统：按“输出说明”标记导出简化说明文件。

当前实现：
- `core/src/base_audit/central_statistics/explanation_exporter.py`

相关配置：
- S3-CFG-00、S3-CFG-01、S3-CFG-02、S3-CFG-03（版本见各自 CHANGELOG）

测试基线：
- 以建立版本管理体系当日 core 全量 pytest 通过状态为基线（详见当次
  GPT修改记录）。
