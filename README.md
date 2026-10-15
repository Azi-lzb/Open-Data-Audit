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

审核工具 V26.1.2.0 起，Linux 用户可在 **设置中心 → 帮助与更新 → 安装 LibreOffice** 选择两种方式：查看终端安装步骤，或选择 DEB 格式的 tar.gz 安装包/解压目录，通过系统管理员授权窗口辅助安装。内置入口优先使用 APT 安装整组 DEB 并处理依赖。麒麟 V11 需先切换至 Maintain Mode 并重启；用户实机反馈主系统安装后 S1-F1 可直接调用 LibreOffice。Windows 不显示该入口。操作及限制见 [Linux 安装说明](docs/LibreOffice安装说明.md)。

逐笔统计（S1）、报表采集（S2）、大集中统计（S3）的 Excel 读写与条件格式处理依赖
**LibreOffice Calc**（Windows 侧对应 Excel/WPS）。目标机已装有 LibreOffice 或能联网
安装的，无需额外处理；目标机无网络且未安装时，需要用 LibreOffice 官方离线包安装
（本机使用版本 **26.8.0，Linux x86-64 deb 版**，约 210MB）。

离线安装包**不在本仓库**。请从 [LibreOffice 官方下载页](https://www.libreoffice.org/download/)
自行选择系统、CPU 架构及 Linux DEB 安装包。开发人员推荐 **25.8.7**；官网下载页
未列出该版本时，可在[官方历史版本目录](https://downloadarchive.documentfoundation.org/libreoffice/old/)
查找 25.8.7，维护状态以官方说明为准（内网可通过批准的共享介质拷贝）。目标机安装：

**版本与系统匹配（重要）**：LibreOffice 官方离线包同样受 glibc 约束。离线机
请按目标系统选择：

- **开发人员推荐 25.8.7**（此前 x86-64 DEB 包审计 GLIBC ≤2.28 / GLIBCXX ≤3.4.22）；
  审计仅针对已检查的包，仍需按实际架构、系统依赖及目标机器验收选择；
- 7.6.7.2 的历史 x86-64 包审计 GLIBC 最大 2.17、GLIBCXX 最大 3.4.19，
  仅作兼容性参考，不作为当前推荐版本或全平台兼容承诺；
- 26.8 按新系统基线构建（实测要求 glibc ≥2.34），**仅适用 UOS25 等新系统**，
  装到 UOS20 会在公式计算时报 GLIBC 版本缺失。

程序会自动发现安装的 LibreOffice（/opt/libreoffice*、PATH、
/usr/lib/libreoffice/program/soffice、玲珑商店版均在检索范围内）。**S1F1
（汇总核查表校验）开跑前会先做计算引擎预检**：引擎缺失或无法启动（如安装包
与系统 glibc 不匹配）会立即报出原因和处理建议，不再跑到中途才失败。

**优先用 deb 离线包安装，不要只用应用商店的玲珑版**：玲珑（ll-cli）每次启动
都要做容器装配，实测冷启动约 10 秒，而 S1F1 公式计算对每个工作簿各启动一次
LibreOffice，累积明显拖慢审核；deb 版直接运行二进制（实测 --version 约
0.2 秒）。同时安装两者时程序自动优先 /opt 的 deb 版，玲珑版保留兜底。

```sh
tar -xzf LibreOffice_25.8.7_Linux_x86-64_deb.tar.gz -C /tmp
cd /tmp/LibreOffice*_deb/DEBS && sudo apt install ./*.deb
```

可以先安装审核工具的 DEB，再通过内置入口安装 LibreOffice（`dist/` 构建产物同样不入库，见打包
脚本说明）。审核工具启动与冒烟不依赖 LibreOffice，仅执行涉及 Excel 的功能
（如 S1 汇总核查表校验）时需要。

提交前运行 `python tools/check_public_snapshot.py` 检查已暂存文件，防止把真实配置或数据文件带入公开仓库。
