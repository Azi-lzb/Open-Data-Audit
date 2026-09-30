# 审核工具（Flask 外壳：Windows + 统信 UOS/麒麟）

`shell-flask` 是跨平台 Flask 外壳，与 `shell-pywebview`（仅
Windows）共用 `core` 的唯一业务核心（`base_audit` 后端 +
`frontend` 前端 + `tests` 测试）。本目录只放外壳与平台启动/打包脚本。

窗口标题、关于页面与程序文件名统一读取 `版本管理/VERSION_MANIFEST.json`。
当前整包版本为 `V26.1.0.0`（年份.大版本.中版本.小版本），对应：

- pywebview：`审核工具_V26.1.0.0.exe`
- Flask Win10/11：`审核工具_Flask_V26.1.0.0.exe`
- Flask Win7：`审核工具_Flask_Win7兼容_V26.1.0.0.exe`
- UOS/Linux 可执行程序：`审核工具_V26.1.0.0`

图标采用 `core/frontend/assets/app-icon.png` 的红色图形，Windows 程序使用同源 `.ico`。

## 架构

- **前后端唯一接口是桥接层**：页面注入脚本把 `bridge.api.方法名(...)`
  代理为 `POST /api/方法名`，前端 HTML 两个平台一字不改。
- **引擎按操作系统自动识别**（`core/base_audit/engines`）：
  - Windows：Microsoft Excel / WPS 表格 COM；
  - 统信 UOS/麒麟：LibreOffice Calc（`soffice --headless` 重算 +
    openpyxl 原生管线，条件格式走 OOXML 规则求值，与 WPS 兜底同族）。
- 平台差异收口：文件打开（`os.startfile` / `xdg-open`）、引擎列表、
  计算管线（COM / native）均在 core 的 engines 与 native 包，外壳不判断。

## Windows 运行与打包

```bat
Windows-1-启动审核工具.bat     :: 运行（源码模式，绑定根目录 core；--packed 用打包版）
Windows-2-打包双版本.bat       :: 构建 Win10/11 与 Win7 两个 EXE，放在同一个发布文件夹
```

### Windows 双发行形态（同一业务核心）

运行 `Windows-2-打包双版本.bat`。它在 `dist\` 下新建一个目录，把两个
EXE 与共用的 `core\`、`config\` 放在一起；用户按自己的 Windows 版本选择
对应 EXE。两套构建脚本保存在 `windows-build/`，由双版本入口依次调用。
不要把旧 `dist\core\data` 或旧 `dist\config` 当作发布内容；它们可能包含
本机运行数据。合并发布目录的布局为：

```
dist\
└─ windows-YYYYMMDD_HHMMSS\
   ├─ 审核工具_Flask_V26.1.0.0.exe          ← Win10/11
   ├─ 审核工具_Flask_Win7兼容_V26.1.0.0.exe  ← Win7
   ├─ core\                              ← src、frontend 与空的 data 目录
   ├─ config\                            ← 当前正式配置副本
   ├─ 运行库修复\                         ← Win7 缺运行库时使用
   └─ 运行环境说明-Win7.txt
```

两个 EXE 都从仓库根目录同一份 `core\` 与 `config\` 生成。Win10/11 用户
双击带 `Flask` 标记的推荐版；Win7 用户双击带
`Win7兼容` 标记的版本。若用启动器验证打包版：`Windows-1-启动审核工具.bat --packed`
启动推荐版，`Windows-1-启动审核工具.bat --packed win7` 启动 Win7 版。

双版本打包的 Win10/11 阶段（内部 `windows-build/build-modern.bat`）首次会从
本机 64 位 Python 3.11/3.12 创建仓库根目录
`.venv`，并安装 `requirements-win10win11-build.txt`；复制来的旧 `.venv`
不可运行时改用独立的 `.venv-flask-build`。本机已有的 Conda Python 3.11
也可用于创建该环境；其它位置可设置 `PYTHON_MODERN` 指向解释器。
Win7 使用独立的 32 位 Python 3.7/3.8 环境。

> **绿色便携版已移除**（2026-09-21 用户要求）：不再生成内嵌运行时的
> 便携目录/ZIP。`build_win7.py portable` 会明确报错；历史产物留在原处，
> 请只分发本次生成的时间戳目录。

Win7 阶段（内部 `windows-build/build-win7.bat`）使用已验证的 32 位 Python；旧虚拟环境失效时创建
`build-win7-venv-rebuilt`，不会复用原有坏环境。可通过 `PYTHON_WIN7`
指定其它 32 位 Python 3.7/3.8。离线依赖缓存包括：

- `runtime\python38-full`：完整版 3.8.10（PyInstaller 不支持 embeddable，EXE 构建专用）；
- `runtime\python38`：embeddable 缓存（仅引导用，不再随包发行）；
- `wheels\`：全套 win32 固定版本 wheels（flask 2.2.5 / openpyxl 3.1.3 / pywin32 306 /
  pyinstaller 4.10 等，见 `requirements-win7.txt`，纯 ASCII——老 pip 按 ANSI 读），
  构建机**完全离线**可用；
- `redist\vc_redist_2019.x86.exe`：VS2019 线（14.29，官方支持 Win7 SP1）离线运行库
  修复包，随发行包附带。

Win7 包不在 EXE 旁散放 DLL。目标机使用系统 UCRT；裸 Win7 报缺
`api-ms-win-crt-*.dll` 时，先安装包内「运行库修复\vc_redist_2019.x86.exe」，
然后重新启动工具。安装运行库可能需要管理员权限。

> **勿从构建机拷 ucrtbase.dll**：Win10/11 的 ucrtbase 自身 import Win8+ 的
> apiset（如 api-ms-win-core-sysinfo-l1-2-0），app-local 优先于系统目录，
> 在 Win7 上会让 python.exe 直接「无法启动」——2026-09-21 真机踩坑实证。
> 本打包脚本不再进行 app-local UCRT 分发。

**传输**：合并打包后只拷贝一个 `dist\windows-*\` 目录
（两个 EXE 必须与 core\ 同级）。外发前核对 `config\` 中的机构信息是否适合接收方。
`dist\` 根目录中的旧 EXE、`core\`、`config\` 等是历史产物，切勿与时间戳目录混发。
往目标机传大目录树时务必确认拷贝完整——真机曾因丢文件报
`can't open file 'run.py'`；新 EXE 已把全部业务模块打进包内，不再依赖
外部拷贝的完整性。

## 统信 UOS/麒麟 运行与打包

### 启动与打包

在文件管理器中打开 `shell-flask/`，右键 `Linux-1-启动审核工具.sh` 选
“在终端中运行”即可启动（加 `--background` 参数为后台启动并自动打开浏览器）。
其余脚本（打包、冒烟、源码验收包等）同样方式运行，用途速查见同目录 `文件说明.txt`。

说明：不再随仓库和发行包提供 `.desktop` 快捷方式——实测 UOS 桌面对该类入口
的路径解析不可靠（相对 Exec 拒载、启动工作目录不定），双击体验反而不如直接
运行脚本。安装 DEB 的用户从系统应用菜单启动「基础数据审核工具 V3」，该入口
由安装包注册，稳定可靠。入口文件在版本库中带 Linux 可执行权限，无需手动
chmod。纯源码启动需要 Python 3.8+ 及 `requirements.txt` 依赖；UOS20 自带的
Python 3.7 不满足源码启动要求（发行包不受影响，见下）。

### 冻结发行包（DEB + tar.gz，自带 Python）

终端入口按用途分开，共用同一构建引擎 `Linux-2-打包DEB与TARGZ.sh`：

```sh
/bin/sh ./Linux-1-启动审核工具.sh        # 前台调试运行（源码模式）
/bin/sh ./Linux-2-打包DEB与TARGZ.sh      # 冻结+审计，一次出 DEB + tar.gz
/bin/sh ./Linux-3-测试发行包启动.sh        # 本机实测 dist/ 最新 .deb 与 .tar.gz 能否启动
```

- **DEB**：文件管理器双击安装到 `/opt/base-audit-v3/`，注册应用菜单
  「基础数据审核工具 V3」（`base-audit-v3` 命令同效）。
- **tar.gz 便携包**：DEB 的备用方案——目标机无 root、包管理受限或不想安装时，
  解压到任意可写目录，运行其中 `shell-flask/Linux-1-启动审核工具.sh`
  （文件管理器右键“在终端中运行”或终端直接执行）；启动器自动识别发行包布局。
- **运行依赖**：目标机需桌面会话、xdg-utils、iproute2 和 LibreOffice Calc
  （S1/S2/S3 的 Excel 处理依赖；离线机用 LibreOffice 官方 deb 离线包安装，
  详见仓库根 README「运行依赖」一节——该安装包不入库，随介质分发）。
- **Linux-3-测试发行包启动.sh**：打包后、分发前的本机实测（解包→启动→HTTP 200 与页面
  标题校验→释放端口→汇总）；可随产物拷到 UOS/麒麟目标机做安装前预检，
  测试机不需要 Python。

引擎默认**不用构建机系统 Python**，而是下载 python-build-standalone 的自包含
CPython 3.12.14（glibc 2.17 基线构建，含 Tcl/Tk）并用 PyInstaller 冻结，因此
**在一台 UOS 25 上构建的包同样能运行在 UOS 20（glibc 2.28）/ 麒麟 V10 / UOS 25**；
构建机只需 `dpkg-deb`、`tar`、`objdump`（binutils）、`sha256sum` 和网络（首次，
官方 PyPI 失败自动改用镜像）。产物文件名含构建系统、Python 版本、**实测最低
glibc** 与架构（如 `..._uos-25_py3.12.14_glibc2.25_amd64.deb`）——构建后逐个
ELF 审计 GLIBC/GLIBCXX 符号版本，超过目标基线（默认 glibc 2.28，即 UOS20）时
构建失败并列出超标文件，取代旧的“到最老目标系统上构建”要求。

**跨机器/跨芯片**：PyInstaller 不能交叉编译，不同 CPU 架构需到对应机器上打包；
`amd64` 与 `arm64`（飞腾/鲲鹏）打包机自动下载对应架构解释器，其它架构回退该机
系统 Python（`--python` 可指定解释器）。UOS 与麒麟（Kylin V10 等 Debian 系）
均适用：DEB 依赖只有 libc6/libcrypt1/libstdc++6/xdg-utils/iproute2，tar.gz
完全免安装。

**配置边界**：默认打包使用仓库公开 `config/` 的空表头配置（构建时仍做隐私
清理校验，机构参照/审核历史所在表只保留表头）；`--with-local-config` 从本机
不入库的 `config-real/` 读取正式配置，仅供本人验证，禁止外发。首次启动把包内
六册配置复制到 `~/.local/share/base-audit-v3/config/`，运行数据保存在同目录下；
升级或换便携包目录不会覆盖已有用户配置。LibreOffice Calc 仍需目标系统提供；
本仓库未提供使用说明 DOCX 时发行包不含该文档（放回仓库根目录或 `core/` 即可
自动带入）。

**验收状态**：2026-09-29 在 UOS Desktop 25 上完成全流程构建并对 DEB 与 tar.gz
做了解包冒烟（均 HTTP 200、端口释放干净，tar.gz 便携模式自动识别通过），冻结包
实测最低 glibc 2.25（来源 lxml `getentropy`）。UOS 20 / 麒麟 V10 / arm64 实机
安装与业务全流程仍待真机验收。逐笔系统的组合流程与 S1-F05「制作联合模板」在
UOS 使用共享核心的 native/openpyxl 实现；只有确实依赖 Excel/WPS COM 原生语义
的能力限于 Windows。

## 测试

```sh
cd core && PYTHONPATH=src python3 -m pytest tests/ -q
```

## 与其他发行线的关系

- 业务功能一律改 `core`，两个外壳同时生效；归档线见
  `external/Flask`、`external/pywebview2`（各含归档说明）。
