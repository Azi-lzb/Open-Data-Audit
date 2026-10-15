# Linux 内置 LibreOffice 安装入口

适用版本：审核工具 V26.1.5.0 起。位置：设置中心 → 帮助与更新 → 安装 LibreOffice。自 V26.1.5.1 起，Windows 7、Windows 10/11、Linux AMD64 与 Linux ARM64 使用相同页面布局；Windows 可查看 Linux 安装步骤与麒麟 V11 维护模式说明，但这些系统操作按钮不可用。

**麒麟 V11 提示：** 已有 V11 实机反馈确认：切换到维护模式并重启后，使用系统 APT 安装 LibreOffice DEB 可成功；安装后的主系统 LibreOffice 可被审核工具直接发现，S1-F1 实测成功，无需经 KARE 启动。帮助与更新页面可检查当前模式、提交进入维护模式或保存退出；操作需要系统授权，并由用户手动重启生效。终端安装前也可运行 `sudo mm-cli -s` 查看模式。详见[麒麟 V11 安装与 S1 验证记录](麒麟V11_KARE安装LibreOffice与S1计算方案.md)。

**审核工具自身也要在主系统运行。** 本次 V11 反馈已确认：双击旧审核工具 DEB 后，`kare -l` 列出了 `base-audit-v3`，主系统 `dpkg-query -W audit-tools` 显示没有匹配软件包；菜单启动的旧审核工具把主系统 LibreOffice/WPS 都显示为“未找到”，解压发行包运行 `Linux-1-启动审核工具.sh` 则可发现。KARE 中的 `/opt` 与主系统不同，修改 LibreOffice 搜索路径无法跨越文件系统隔离。切换到维护模式后，在只放有本次安装包的目录中执行 `sudo apt install ./Audit-Tools_V*.deb`，将新版审核工具安装到主系统。用 `dpkg-query -W audit-tools` 确认主系统已登记，再从主系统终端执行 `audit-tools` 核对环境信息。若菜单有两个同名入口，先核对新入口的版本和启动命令 `/usr/bin/audit-tools`；旧 KARE 包待新入口验收通过后再单独处理。

若收到完整的 Linux 离线交付 tar.gz，先核对文件名中的 `amd64` 或 `arm64` 与目标机器架构一致；解压后进入该目录，执行 `sh ./install.sh --check`，再执行 `sh ./install.sh`。UOS 20/25、麒麟 V10/V11 桌面版共用此入口；仅 V11 要先切换维护模式。交付包附同架构 LibreOffice 25.8.7 的 42 个 DEB 和审核工具 DEB；脚本核对版本、架构及 SHA-256 后调用主系统 APT，禁止联网下载缺失依赖。普通用户从主系统终端运行脚本，脚本只在必要步骤调用 `sudo`。公开包使用空表头配置；文件名带 `internal-config` 的包包含真实配置，仅供内部测试，禁止外发。ARM64 包可由 x86-64 开发机通过 Docker/QEMU 构建，但仍须在目标 ARM64 机器上验收。

## 两种安装方式，任选一种

### 方式一：终端安装

1. 从 [LibreOffice 官方下载页](https://www.libreoffice.org/download/)选择与本机架构、系统相符的 Linux DEB 安装包，右键解压后找到 `DEBS` 文件夹。开发人员推荐 **25.8.7**；官方下载页未列出该版本时，可在[官方历史版本目录](https://downloadarchive.documentfoundation.org/libreoffice/old/)查找 25.8.7，维护状态以官方说明为准。
2. 在该文件夹空白处右键选择“在终端中打开”。如果先打开了终端，执行下列命令，把示例路径替换为实际的 DEBS 目录；路径有空格时保留双引号。

```sh
cd -- "/实际路径/DEBS"
sudo apt install ./*.deb
```

`./*.deb` 会把当前 DEBS 目录中的包交给 APT 一次处理；APT 会按系统已配置的软件源检查依赖。密码由终端中的 sudo 接收，输入时通常不显示字符。执行命令前保存并关闭 Office 文档；若安装失败，请保留终端输出交给管理员。手动命令不由审核工具执行。

### 方式二：辅助安装

1. 从 [LibreOffice 官方下载页](https://www.libreoffice.org/download/)取得与系统、CPU 架构相符的 **DEB 格式**安装包。开发人员推荐 **25.8.7**，可在[官方历史版本目录](https://downloadarchive.documentfoundation.org/libreoffice/old/)查找；不强制版本，实际兼容性需在目标机器确认。
2. 点击“选择安装包”选择 `.tar.gz` / `.tgz`，或点击“选择文件夹”选择解压目录、包含一层安装包目录的文件夹、`DEBS` 文件夹。
3. 等待检查完成，核对显示的版本、架构及安装文件数量。保存并关闭正在编辑的 Office 文档。
4. 点击“开始安装”，确认后在 Linux 系统自己的管理员授权窗口输入密码。审核工具不读取、不保存密码。
5. 保持工具打开，等待安装结束。安装完成后会尝试启动新版本检查其是否可用；需要重算时再执行审核。

## 多版本检测与卸载

- 页面会列出已检测到的版本、来源和程序路径，并标记“当前用于审核”的版本。明确设置的 `BASE_AUDIT_SOFFICE` 或 `BASE_AUDIT_LIBREOFFICE_DIR` 优先；未明确设置时，审核工具优先使用数值版本最高的 `/opt/libreoffice主版本.次版本/program/soffice`。
- 不同主版本目录通常可并存；同一主版本目录的更新会替换对应 DEB 包。不同版本仍可能共用并更新公共依赖，不能保证完全隔离。安装新版本不会自动卸载旧版本。
- 对本页识别到的官方版本化 DEB，可选择卸载单个主版本。卸载前 APT 会模拟计划；如果计划包含所选版本之外的软件包，操作会中止。程序不执行自动清理；系统仓库版和手动安装目录只显示状态，不由本页卸载。
- 麒麟 V11 安装或卸载都需要当前已确认处于维护模式。切换模式后先重启，再打开页面继续操作。安装完成后可点“退出维护模式”保存更改并提交退出，再手动重启；不要使用会丢弃更改的 `mm-cli -c -n reboot`。

## 系统与安装限制

- 需要 DEB 包管理环境及系统管理员权限；采用 RPM 的 Linux 版本不支持此入口。
- 普通用户需要系统提供 `pkexec` 和桌面授权代理。没有授权窗口或管理员权限时，请联系管理员通过系统软件管理工具安装。
- 包名、CPU 架构检查不等于发布者身份验证，必须选择可信来源的安装包。架构相同也不保证包所需的系统库版本满足要求。
- 内置安装优先使用系统 `apt-get` 一次安装已校验的 DEB 包集并处理依赖；系统没有 APT 时回退到 `dpkg --install`。命令以参数列表执行，不拼接 shell 命令。APT 可能使用系统已配置的软件源获取缺少的依赖。
- 麒麟 V11 必须先进入维护模式并重启后再安装或卸载。界面只调用固定的 `mm-cli -s/-o/-c` 参数，不接受任意命令；离线交付脚本只查询模式，不切换模式。界面和脚本都不会自动重启。
- 系统缺依赖、包管理器被占用或包损坏时可能失败；部分软件包可能已解包。请查看安装日志交给管理员处理，不能将失败视为自动回滚。
- 安装后台独立运行；切换页签不会取消安装。安装期间不能启动新的审核任务，正在审核时不能开始安装。
- 卸载和维护模式切换也不能与审核任务并行；版本卸载只针对确认过的 DEB 软件包清单。
- 安装不修改审核配置工作簿、用户业务设置或原始数据。

## 桥接 API

共享实现位于 `core/src/base_audit/native/libreoffice_install.py`；Flask 和 pywebview 通过同一 `WebApi` 使用。

| 接口 | 参数 | 返回与含义 |
|---|---|---|
| `get_libreoffice_install()` | 无 | 独立安装状态；成功后刷新环境摘要 |
| `choose_libreoffice_source(folder, path)` | `folder` 为布尔，`path` 为本机绝对路径；留空调用当前原生/Tk 选择器 | 启动后台检查，原始包和目录只读 |
| `start_libreoffice_install()` | 无 | 仅安装已检查并暂存的文件；不能提交任意命令或额外安装参数 |
| `start_libreoffice_uninstall(release)` | 版本主目录，如 `25.8` | 仅允许卸载检测到的该版本化 DEB 包集；APT 模拟计划越界时中止 |
| `start_kylin_maintenance(action)` | `status / enable / exit` | 仅允许检查模式、运行固定 `mm-cli -o` 或保存退出 `mm-cli -c` |
| `browse_dir(path, "libreoffice_archive")` | 复用已有目录选择器 | 只显示 `.tar.gz` / `.tgz` 和文件夹 |

状态字段还包括 `installedVersions`、`activeEngine`、`kylinMaintenance`。`phase` 包含 `idle / checking / ready / installing / uninstalling / checking_maintenance / changing_maintenance / success / failed / cancelled`。用户页面只显示文件名，完整来源放悬停提示，技术细节保存在操作日志。

Flask 的安装、卸载和维护模式接口仅接受本机主机名、回环地址的 JSON 请求；带 Origin 的请求必须与审核工具页面同源。权限提升固定调用系统 `pkexec` 和系统包管理器，不以管理员身份运行整个审核工具。

## 验收范围

既有自动化测试覆盖安装输入检查、暂存保护、固定命令参数、状态转换、权限取消与失败，以及 UI/桥接边界。新增版本卸载、V11 模式操作和多版本活动引擎仍需在 Linux/V11 真机确认。麒麟 V11 用户已报告维护模式下通过 APT 安装及 S1-F1 实机成功；UOS 设备的授权弹窗、系统依赖和 S1 重算仍需在对应桌面环境确认。
