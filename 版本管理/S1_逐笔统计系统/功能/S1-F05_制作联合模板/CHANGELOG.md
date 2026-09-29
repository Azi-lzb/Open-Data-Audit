# S1-F05 制作联合模板

当前版本：v1.0.1

## v1.0.1 — 2026-09-19

变更类型：Bugfix / PATCH

前版本：v1.0.0
新版本：v1.0.1

修改内容：
- 修复 run_template_merge COM 分支缺少 `from .excel_com import ExcelSession`
  局部导入导致的 NameError（Windows 制作联合模板必坏；同文件
  _merge_one_org_com 一直有正确范式，本函数遗漏）

发现方式：
- 11 功能面板实测（本轮引入的功能级真实烟测），非用户报障

配置影响：
- 无

测试：
- COM 真实执行：复制 1 张工作表、迁移命名区域 3 个，成功
- core 全量 pytest 通过

Git提交：
- 见仓库日志（fix: S1-F05 制作联合模板 COM 路径缺失 ExcelSession 导入）

## v1.0.0 — 2026-09-19

变更类型：Baseline

说明：

V3 版本管理体系建立时的正式基线。此版本建立前的开发历史未逐项倒推
SemVer；可通过 `GPT修改记录/` 与 Git 历史继续追溯。

组件说明：
- 逐笔统计系统：由报送核查表与正式模板制作联合审核模板（UOS 为 openpyxl 实现，语义与 COM 版一致）。

当前实现：
- `core/src/base_audit/native/template_merge.py`

相关配置：
- S1-CFG-01（版本见各自 CHANGELOG）

测试基线：
- 以建立版本管理体系当日 core 全量 pytest 通过状态为基线（详见当次
  GPT修改记录）。
