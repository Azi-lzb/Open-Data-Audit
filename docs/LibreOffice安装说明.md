# Linux 内置 LibreOffice 安装入口

适用版本：审核工具 V26.1.2.0 起。位置：设置中心 → 帮助与更新 → 安装 LibreOffice。Windows 不显示该入口。

## 两种安装方式，任选一种

### 方式一：终端安装

1. 下载与本机架构、系统相符的 Linux DEB 安装包，右键解压后找到 `DEBS` 文件夹。
2. 在该文件夹空白处右键选择“在终端中打开”。如果先打开了终端，执行下列命令，把示例路径替换为实际的 DEBS 目录；路径有空格时保留双引号。

```sh
cd -- "/实际路径/DEBS"
sudo dpkg -i *.deb
```

`-i` 后必须有空格。密码由终端中的 sudo 接收，输入时通常不显示字符。执行命令前保存并关闭 Office 文档；若安装失败，请保留终端输出交给管理员。手动命令不由审核工具执行。

### 方式二：辅助安装

1. 从[项目依赖包页面](https://github.com/Azi-lzb/Open-Data-Audit/releases/tag/Needed)或 LibreOffice 官方渠道取得与系统、CPU 架构相符的 **DEB 格式**安装包。
2. 点击“选择安装包”选择 `.tar.gz` / `.tgz`，或点击“选择文件夹”选择解压目录、包含一层安装包目录的文件夹、`DEBS` 文件夹。
3. 等待检查完成，核对显示的版本、架构及安装文件数量。保存并关闭正在编辑的 Office 文档。
4. 点击“开始安装”，确认后在 Linux 系统自己的管理员授权窗口输入密码。审核工具不读取、不保存密码。
5. 保持工具打开，等待安装结束。安装完成后会尝试启动新版本检查其是否可用；需要重算时再执行审核。已有版本和显式指定的引擎配置可能影响实际使用的版本，可在高级设置环境信息查看。

## 系统与安装限制

- 需要 DEB 包管理环境及系统管理员权限；采用 RPM 的 Linux 版本不支持此入口。
- 普通用户需要系统提供 `pkexec` 和桌面授权代理。没有授权窗口或管理员权限时，请联系管理员通过系统软件管理工具安装。
- 包名、CPU 架构检查不等于发布者身份验证，必须选择可信来源的安装包。架构相同也不保证包所需的系统库版本满足要求。
- 使用系统的 `dpkg --install` 安装选中的 LibreOffice 文件，不执行包外脚本，不自动卸载已有版本，不自动联网下载依赖。
- 系统缺依赖、包管理器被占用或包损坏时可能失败；部分软件包可能已解包。请查看安装日志交给管理员处理，不能将失败视为自动回滚。
- 安装后台独立运行；切换页签不会取消安装。安装期间不能启动新的审核任务，正在审核时不能开始安装。
- 安装不修改审核配置工作簿、用户业务设置或原始数据。

## 桥接 API

共享实现位于 `core/src/base_audit/native/libreoffice_install.py`；Flask 和 pywebview 通过同一 `WebApi` 使用。

| 接口 | 参数 | 返回与含义 |
|---|---|---|
| `get_libreoffice_install()` | 无 | 独立安装状态；成功后刷新环境摘要 |
| `choose_libreoffice_source(folder, path)` | `folder` 为布尔，`path` 为本机绝对路径；留空调用当前原生/Tk 选择器 | 启动后台检查，原始包和目录只读 |
| `start_libreoffice_install()` | 无 | 仅安装已检查并暂存的文件；不能提交任意命令或额外安装参数 |
| `browse_dir(path, "libreoffice_archive")` | 复用已有目录选择器 | 只显示 `.tar.gz` / `.tgz` 和文件夹 |

状态字段：`available`、`supported`、`busy`、`phase`、`message`、`source`、`sourceName`、`packageVersion`、`architecture`、`packageCount`、`installed`、`logPath`、`canInstall`；桥接额外提供 `auditBusy`。`phase` 包含 `idle / checking / ready / installing / success / failed / cancelled`。用户页面只显示文件名，完整来源放悬停提示，技术细节保存在安装日志。

Flask 的安装接口仅接受本机主机名、回环地址的 JSON 请求；带 Origin 的请求必须与审核工具页面同源。权限提升固定调用系统 `pkexec` 和 `dpkg`，不以管理员身份运行整个审核工具。

## 验收范围

自动测试覆盖输入检查、暂存保护、固定命令参数、状态转换、权限取消与失败，以及 UI/桥接边界。WSL 检查实际项目压缩包及解压目录；Docker 隔离安装验收结果见本轮 GPT 修改记录。仍需在 UOS/麒麟桌面真机检查授权弹窗、实际安装、系统依赖和安装后的 S1 LibreOffice 重算；容器及 Windows 模拟测试不能替代目标桌面验收。
