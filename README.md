# Open Data Audit

> **本项目已暂停更新。** 后续开发与功能更新转至 [Rust-Audit](https://github.com/Azi-lzb/Rust-Audit)，请前往该项目关注最新进展及使用说明。本仓库保留现有源码、历史记录和已发布安装包，供参考及兼容使用。

审核工具（AuditTools）V3 的公开代码仓库，本仓库保留版本为 **V26.2.2.5**。程序在本机处理 Excel 工作簿，提供 Windows 桌面窗口、Windows 浏览器工作台和 Linux 浏览器工作台。

## 下载与选择版本

前往 [审核工具 V26.2.2.5 Release](https://github.com/Azi-lzb/Open-Data-Audit/releases/tag/v26.2.2.5) 下载已打包程序；本项目历史发行包可在 [Releases](https://github.com/Azi-lzb/Open-Data-Audit/releases) 查看。Release 页面自动生成的 **Source code** 是源码，不是安装包。

- **Windows-Flask ZIP**：浏览器工作台，同一个压缩包内提供 Windows 10/11 64 位程序和 Windows 7 SP1 兼容程序。
- **Windows-PyWebView ZIP**：Windows 10/11 64 位桌面窗口版，需要 Microsoft Edge WebView2 Runtime。
- **Linux AMD64 DEB**：适用于满足运行依赖的 Debian/Ubuntu 系桌面系统，包括 UOS、麒麟对应发行版；通过系统安装器安装。
- **Linux AMD64 TAR.GZ**：便携版，解压后启动。Linux 发布架构为 x86_64/AMD64，构建兼容目标为 glibc 2.28，实际依赖下限以附件文件名和发布说明为准。

已打包版本自带 Python 运行时，使用者无需另装 Python。Linux 不发布 ARM64 包。Windows 7 兼容包的实际运行仍取决于系统补丁、运行库及 Office 环境。

**公开版仅附空表头配置，不含真实机构信息、业务规则、报送文件或正式模板。** 正式审核前需要通过独立渠道取得适用配置和模板，并在程序中选择。源码 ZIP 同样不包含这些私有资料。

## 首次使用

### Windows

1. 将所选 ZIP 完整解压到有写入权限的目录，保留 `exe`、`core/` 和 `config/` 的同级关系。
2. Flask 包：Windows 10/11 双击 `审核工具_Flask_V26.2.2.5.exe`；Windows 7 SP1 双击 `审核工具_Flask_Win7兼容_V26.2.2.5.exe`。启动后浏览器会打开本机工作台。
3. PyWebView 包：双击 `审核工具_V26.2.2.5.exe`。缺少 WebView2 时先安装对应运行时，或使用 Flask 包。
4. 按需准备 Microsoft Excel 或 WPS 表格，供需要公式重算、Office 自动化的功能使用。Win7 如提示缺少运行库，先运行 Flask 包内 `运行库修复/vc_redist_2019.x86.exe`，再启动程序。
5. 在所需系统中选择配置、源文件、模板或外部文件，再执行相应功能。空表头配置不能直接完成正式审核。

### Linux

DEB：下载后通过系统软件安装器打开，也可以在下载目录执行：

```sh
sudo apt install ./Audit-Tools_V26.2.2.5_*_amd64.deb
```

安装完成后，从应用菜单打开“审核工具”。涉及公式重算的功能还需安装匹配系统的 LibreOffice Calc。

TAR.GZ：解压后进入产品目录，运行其中的 **`shell-flask/Linux-1-启动审核工具.sh`**。文件管理器支持时，可右键选择“在终端中运行”；也可在产品目录打开终端执行：

```sh
sh ./shell-flask/Linux-1-启动审核工具.sh
```

Linux 用户配置和运行数据位于 `~/.local/share/base-audit-v3/`；升级安装不会自动覆盖已有用户配置。需要使用新配置时，在程序中重新选择或按配置管理流程更新。

## 功能范围

当前保留 S1、S2、S3，共 11 个顶级功能：

- **S1 逐笔统计系统**：核查表审核、校验结果说明汇总、工作簿及区域汇总、模板命名区域配置。
- **S2 报表采集系统**：报表校验、环比检查、外部数据核对，输出分类审核清单。
- **S3 大集中统计系统**：比较文件生成、说明文件导出及反馈回写、本期数值核对、金融表单拆分。

S4 地金预览及 S5 配置版已退出当前产品，发布包不含其界面与专属配置。

读取值、复制公式、写入文件与公式重算是不同操作。项目按能力使用 Calamine、Direct OOXML、openpyxl 和可用的 Rust 写入器；需要真实工作簿重算时使用 Excel/WPS 或 LibreOffice。普通值读取和新建结果文件不统一要求安装 Office。

## 仓库内容

- `core/src/`：审核、报表采集和大集中统计的后端实现。
- `core/frontend/`：本地工作台界面。
- `core/tests/`：不依赖私有业务文件的测试代码。
- `config/`：仅含工作表名和表头的公开占位配置。
- `rules/`、`AGENTS.md`：项目开发规则。
- `版本管理/`：组件版本清单。

`core/` 是唯一业务核心，两个外壳共用同一套审核逻辑。真实配置、报送数据、验收结果及构建产物不提交到 Git；公开发行包通过 Release 附件分发。

源码运行时，可将独立取得的真实配置放在仓库根目录的 `config-real/`，程序优先读取该目录。**不要用真实工作簿覆盖 `config/` 中已跟踪的占位文件。** 公开发行应从不含 `config-real/` 的干净源码快照构建；本机打包脚本默认优先携带 `config-real/`，不能直接将内部包上传公开 Release。

## 源码运行与验证

推荐使用 Python 3.11/3.12 创建独立虚拟环境；Windows 7 打包另有 Python 3.8.10 32 位依赖方案。Flask 源码入口：

```sh
python -m venv .venv
# Windows 使用 .venv\Scripts\activate；Linux 使用 source .venv/bin/activate
python -m pip install -r shell-flask/requirements.txt
python shell-flask/run.py
```

默认打开本机 `127.0.0.1:8750` 工作台；端口被占用时程序会按规则选择可用端口。PyWebView 源码入口及打包流程见其 [打包脚本](shell-pywebview/build_exe.py)，Flask 与 Linux 打包流程见 [Flask 外壳说明](shell-flask/README.md)。

回归入口为根目录 `运行全量测试.bat` 和 `run_full_test_linux.sh`，用于有测试依赖的源码环境；打包程序的使用者无需运行测试。版本唯一依据为 [VERSION_MANIFEST.json](版本管理/VERSION_MANIFEST.json)，历史功能变更见 [产品发布记录](版本管理/产品发布/CHANGELOG.md)。

## Linux 的 LibreOffice 安装

Linux 可在 **设置中心 → 帮助与更新 → 安装 LibreOffice** 选择 DEB 格式的 LibreOffice tar.gz 安装包、解压目录或 `DEBS` 文件夹，并通过系统管理员授权窗口辅助安装。这个入口用于 LibreOffice；它不是通用 DEB 安装器。具体操作见 [LibreOffice 安装说明](docs/LibreOffice安装说明.md)。

LibreOffice 离线包不与审核工具捆绑。可从 [LibreOffice 官方下载页](https://www.libreoffice.org/download/) 取得，仓库的 [LibreOffice 附件页](https://github.com/Azi-lzb/Open-Data-Audit/releases/tag/Needed) 也保留历史离线包。选择时核对 CPU 架构、glibc 和系统依赖，不应把适合新系统的包直接装到旧系统。

程序会发现 PATH、`/opt/libreoffice*`、常见系统路径及玲珑商店版中的 LibreOffice；S1-F1 执行前会预检计算引擎。引擎缺失或无法启动时，按程序提示安装适合目标系统的版本。启动审核工具或进行启动冒烟本身不要求 LibreOffice。

提交前运行 `python tools/check_public_snapshot.py` 检查已暂存文件，防止把真实配置或数据文件带入公开仓库。
