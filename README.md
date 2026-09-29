# Open Data Audit

基础数据审核工具 V3 的公开代码仓库。`core/` 是唯一业务核心；`shell-flask/` 提供 Windows 与 UOS/Linux 入口，`shell-pywebview/` 提供 Windows 桌面入口。

## 仓库内容

- `core/src/`：审核、报表采集和大集中统计的后端实现。
- `core/frontend/`：本地工作台界面。
- `core/tests/`：不依赖私有业务文件的测试代码。
- `config/`：仅含工作表名和表头的公开占位配置。
- `rules/`、`AGENTS.md`：项目开发规则。
- `版本管理/`：组件版本清单。

真实配置、报送数据、验收结果及构建产物不在本仓库流转。请通过独立渠道取得真实配置，将其放在仓库根目录的 `config-real/`；源码运行时优先读取该目录。缺少真实配置时，工作台可以启动，但占位配置无法完成正式业务审核。**不要用真实工作簿覆盖 `config/` 中已跟踪的占位文件。**

使用与打包入口见 [Flask 外壳说明](shell-flask/README.md)。公开仓库没有附带 Word 使用说明，打包脚本允许在缺少该文档时继续构建。

提交前运行 `python tools/check_public_snapshot.py` 检查已暂存文件，防止把真实配置或数据文件带入公开仓库。
