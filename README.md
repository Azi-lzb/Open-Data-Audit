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

**版本与系统匹配（重要）**：LibreOffice 官方离线包同样受 glibc 约束。离线机
请按目标系统选择：

- **推荐统一用 7.6.7.2 离线包**——全部 42 个安装包 ELF 符号审计：GLIBC 最大
  2.17、GLIBCXX 最大 3.4.19，**UOS20（glibc 2.28）/ UOS25 等全系目标机通吃**；
- 25.8.7 为备选（审计 GLIBC ≤2.28 / GLIBCXX ≤3.4.22，UOS20 卡线通过）；
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
tar -xzf LibreOffice_26.8.0_Linux_x86-64_deb.tar.gz -C /tmp
cd /tmp/LibreOffice*_deb/DEBS && sudo dpkg -i *.deb
```

装完 LibreOffice 后再安装审核工具的 DEB（`dist/` 构建产物同样不入库，见打包
脚本说明）。审核工具启动与冒烟不依赖 LibreOffice，仅执行涉及 Excel 的功能
（如 S1 汇总核查表校验）时需要。

提交前运行 `python tools/check_public_snapshot.py` 检查已暂存文件，防止把真实配置或数据文件带入公开仓库。
