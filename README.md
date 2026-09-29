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

## 运行依赖：LibreOffice Calc（离线安装包不入库）

逐笔统计（S1）、报表采集（S2）、大集中统计（S3）的 Excel 读写与条件格式处理依赖
**LibreOffice Calc**（Windows 侧对应 Excel/WPS）。目标机已装有 LibreOffice 或能联网
安装的，无需额外处理；目标机无网络且未安装时，需要用 LibreOffice 官方离线包安装
（本机使用版本 **26.8.0，Linux x86-64 deb 版**，约 210MB）。

离线安装包**不在本仓库**（超过 GitHub 单文件 100MB 上限，且第三方二进制按仓库
规则不入库），已放在 [GitHub Release「Libreoffice」](https://github.com/Azi-lzb/Open-Data-Audit/releases/tag/Needed)
（210MB，内网不便时也可用 U 盘/共享介质拷贝）。目标机安装：

```sh
tar -xzf LibreOffice_26.8.0_Linux_x86-64_deb.tar.gz -C /tmp
cd /tmp/LibreOffice*_deb/DEBS && sudo dpkg -i *.deb
```

装完 LibreOffice 后再安装审核工具的 DEB（`dist/` 构建产物同样不入库，见打包
脚本说明）。审核工具启动与冒烟不依赖 LibreOffice，仅执行涉及 Excel 的功能
（如 S1 汇总核查表校验）时需要。

提交前运行 `python tools/check_public_snapshot.py` 检查已暂存文件，防止把真实配置或数据文件带入公开仓库。
