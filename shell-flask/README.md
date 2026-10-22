# 审核工具（Flask 外壳：Windows + 统信 UOS/麒麟）

`shell-flask` 是跨平台 Flask 外壳，与 `shell-pywebview`（仅
Windows）共用 `core` 的唯一业务核心（`base_audit` 后端 +
`frontend` 前端 + `tests` 测试）。本目录只放外壳与平台启动/打包脚本。

窗口标题、关于页面与程序文件名统一读取 `版本管理/VERSION_MANIFEST.json`。
当前整包版本为 `V26.1.5.1`（年份.大版本.中版本.小版本），对应：

- pywebview：`审核工具_V26.1.5.1.exe`
- Flask Win10/11：`审核工具_Flask_V26.1.5.1.exe`
- Flask Win7：`审核工具_Flask_Win7兼容_V26.1.5.1.exe`
- UOS/Linux 可执行程序：`审核工具_V26.1.5.1`

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
Windows-2-打包双版本_S1-S2专用.bat :: 双击构建 S1/S2 专用版，包含 Win10/11 与 Win7 两个 EXE
```

### Windows 双发行形态（同一业务核心）

双击 `Windows-2-打包双版本.bat`，默认构建完整版。它在 `dist\` 下生成版本目录，把两个
EXE 与共用的 `core\`、`config\` 放在一起；用户按自己的 Windows 版本选择
对应 EXE。两套构建脚本保存在 `windows-build/`，由双版本入口依次调用。
交付范围与 Linux 共用 `distribution-profiles.json`。默认读取 `full` 方案；
要生成只显示 S1/S2 的双版本包，直接双击以下专用 BAT，无需输入参数：

```bat
Windows-2-打包双版本_S1-S2专用.bat
:: 等同于 Windows-2-打包双版本.bat --profile s1s2
```

两个 EXE 共用同一份处理后的前端；S3/S4 导航及相关指引、问答、常规和高级设置
按方案隐藏，输出目录追加 JSON 中的 `outputSuffix`（S1/S2 为 `_S1-S2`）。
S1/S2 包按 JSON 的 `excludedConfigPatterns` 排除 3.0—3.3 配置工作簿，
包含默认配置及备份目录中的同名副本；Windows、Linux 使用同一筛选实现。
代码仍完整保留。通过关于口令显示 S3/S4 后，使用这些系统须另行导入相应配置。
版本号自动读取版本清单；完整版目录为 `Windows-Flask_V26.1.5.1`，
S1/S2 版为 `Windows-Flask_V26.1.5.1_S1-S2`。两个版本先在临时目录完成构建，
成功后才发布到版本目录；同名旧包移入 `dist/归档-旧发行包-日期/`，失败时保留旧包。
Windows 与 Linux 共用导航及 S1/S2 交付规则；运行环境、计算引擎与系统操作按平台能力展示。
两个内部构建完成后才处理发行目录中的页面，仓库源页面始终保留完整界面。
Windows-1、Linux-1 的源码开发启动不应用交付隐藏；Linux-1 在发行包内启动时仍遵从包内方案。
pywebview 不应用这套交付隐藏。
隐藏交付版可打开“关于”，在 2 秒内连续点击标题“关于审核工具”5 次，输入交付显示口令后
显示全部系统及相关栏目。本次运行有效，重启后端恢复原交付范围；详见 `../docs/开发入口与本机开关.md`。

不要把旧 `dist\core\data` 或旧 `dist\config` 当作发布内容；它们可能包含
本机运行数据。合并发布目录的布局为：

```
dist\
└─ Windows-Flask_V26.1.5.1\                ← S1/S2 版追加 _S1-S2
   ├─ 审核工具_Flask_V26.1.5.1.exe          ← Win10/11
   ├─ 审核工具_Flask_Win7兼容_V26.1.5.1.exe  ← Win7
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
> 请只分发本次成功生成的版本目录。

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

**传输**：合并打包后只拷贝一个 `dist\Windows-Flask_V*\` 目录
（两个 EXE 必须与 core\ 同级）。外发前核对 `config\` 中的机构信息是否适合接收方。
`dist\` 根目录中的旧 EXE、`core\`、`config\` 等是历史产物，切勿与版本目录混发。
`windows-时间戳` 是构建中间目录；带 `BUILD_INCOMPLETE.txt` 的目录不可交付。
往目标机传大目录树时务必确认拷贝完整——真机曾因丢文件报
`can't open file 'run.py'`；新 EXE 已把全部业务模块打进包内，不再依赖
外部拷贝的完整性。

## 统信 UOS/麒麟 运行与打包

### 启动与打包

在文件管理器中打开 `shell-flask/`，右键 `Linux-1-启动审核工具.sh` 选
“在终端中运行”即可启动（加 `--background` 参数为后台启动并自动打开浏览器）。
其余脚本（打包、冒烟、源码验收包等）同样方式运行，用途速查见同目录 `文件说明.txt`。

DEB 安装包会注册“审核工具 V版本号”应用菜单和桌面快捷方式；简体中文环境显示
“审核工具”，其他语言环境显示 `AuditTools`。手动运行脚本时，入口文件在版本库中
带 Linux 可执行权限，无需手动 chmod。纯源码启动需要 Python 3.8+ 及 `requirements.txt` 依赖；UOS20 自带的
Python 3.7 不满足源码启动要求（发行包不受影响，见下）。

### 冻结发行包（DEB + tar.gz，自带 Python）

终端入口按用途分开，共用同一构建引擎 `Linux-2-打包DEB与TARGZ.sh`：

```sh
/bin/sh ./Linux-1-启动审核工具.sh        # 前台调试运行（源码模式）
/bin/sh ./Linux-2-打包DEB与TARGZ.sh      # 冻结+审计，一次出 DEB + tar.gz
/bin/sh ./Linux-2-打包DEB与TARGZ.sh --profile s1s2 # 按 JSON 配置生成 S1/S2 交付版
/bin/sh ./Linux-3-测试发行包启动.sh        # 本机实测 dist/ 最新 .deb 与 .tar.gz 能否启动
```

系统归属和交付方案集中在 `distribution-profiles.json`。Linux-2 与 Linux-7 都读取该文件；
调整某方案的 `enabledSystems` 后，重打包会同步控制侧栏、操作指引、问答、常规设置和高级设置中的对应内容。
目前 `full` 默认开放 S1–S4，`s1s2` 默认开放 S1、S2；隐藏仅用于交付界面范围，不拆分业务核心。

Windows 本机双击 `Linux-7-Docker双架构打包.bat`，会启动 Docker Desktop，
再通过 Ubuntu WSL 调用同名 SH，在 Docker 中依次构建 amd64、arm64 完整版。
构建 S1/S2 交付版时传入 `--profile s1s2`；Linux/WSL 终端也可给主脚本加同一参数。
旧的 S1/S2 专用入口仍可作为兼容快捷方式使用。
Ubuntu 须已启用 Docker Desktop WSL 集成及 ARM64 模拟支持。

完整版输出到 `dist/Linux-amd64_V版本/`、`dist/Linux-arm64_V版本/`；
专用版目录追加 `_S1-S2`，每目录仅一个 DEB、一个程序 TAR.GZ、一个程序＋LO TAR.GZ。
专用版隐藏未开放系统的侧栏菜单，以及对应的常规设置、高级设置、操作指引、问答和更新条目；
S1/S2 的外部数据参照保留。只调整交付显示范围，不拆分共用业务核心。
构建默认优先使用 `config-real/`，缺失才回退 `config/`；两个版本各自输出，不互相覆盖。

大集中外部数据模板已登记为 `S2-CFG-02`，文件名以版本清单为准。完整版和S1/S2版都携带它及独立默认快照；Windows两个外壳、Linux内部构建复制本机真实模板，公开构建仅携带第2行指标表头、参照表头及说明。设置中心提供打开、设为默认和重置；重置按整册快照恢复并先备份，运行时仍需选择已更新当期数据的外部文件。首行/A列为空、第2行表头、B列机构标签、C列起数值的位置须保持。

- **DEB**：完整发行包使用 `Audit-Tools_V版本_构建时间_系统_Python_glibc_架构.deb`；S1/S2 专用版以 `Audit-Tools-S1-S2_V` 开头；
  专用 DEB 包名为 `audit-tools-s1-s2`，与完整版互斥，安装时由 APT 替换，不可并存。
  UOS 20/25、麒麟 V10/V11 桌面版均推荐从主系统终端以 APT 安装到
  `/opt/base-audit-v3/`；麒麟 V11 先进入维护模式。V11 图形安装器可能把 DEB 放进
  KARE 兼容环境，届时菜单启动的审核工具
  可能看不到主系统的 LibreOffice/WPS。完整包管理器标识为 `audit-tools`，S1/S2 专用版为 `audit-tools-s1-s2`；注册应用菜单
  「审核工具 V版本」及桌面快捷方式。原 `base-audit-v3` 包会在升级时替换，旧命令保留为兼容入口。
- **tar.gz 便携包**：DEB 的备用方案——目标机无 root、包管理受限或不想安装时，
  解压到任意可写目录，运行其中 `shell-flask/Linux-1-启动审核工具.sh`
  （文件管理器右键“在终端中运行”或终端直接执行）；启动器自动识别发行包布局。
- **运行依赖**：目标机需桌面会话、xdg-utils、iproute2 和 LibreOffice Calc
  （S1/S2/S3 的 Excel 处理依赖；离线机用 LibreOffice 官方 deb 离线包安装，
  详见仓库根 README「运行依赖」一节——该安装包不入库，随介质分发）。
- **Linux-3-测试发行包启动.sh**：打包后、分发前的本机实测（解包→启动→HTTP 200 与页面
  标题校验→释放端口→汇总）；可随产物拷到 UOS、麒麟、Ubuntu、Debian 等兼容的 glibc Linux
  目标机做安装前预检。发行包自带 Python，测试机无需安装 Python；HTTP 检查优先使用
  curl/wget，两者都没有时才回退到系统 Python。DEB 解包需要 dpkg-deb，其他发行版使用
  同架构且运行库兼容的 TAR.GZ；Alpine/musl 不在当前兼容范围内。
  默认等待服务启动 30 秒；Docker/QEMU 的 ARM64 模拟检查可设置
  `FLASK_SMOKE_START_TIMEOUT=180` 延长等待时间，启动后的 HTTP、图标和端口释放检查保持一致。
- **通用离线交付**：`Linux-6-组装通用离线交付.sh` 将同架构审核工具 DEB
  与 LibreOffice 官方 25.8.7 主安装包组装为单个 tar.gz，包含 42 个 LibreOffice
  DEB、SHA-256 清单和 `install.sh`（源码为 `Linux-5-通用离线安装.sh`）。先核对
  LibreOffice 官方 SHA-256，目标机再执行 `sh ./install.sh --check` 和
  `sh ./install.sh`。适用于对应 CPU 架构的 UOS 20/25、麒麟 V10/V11 桌面版；仅 V11 须先进入
  维护模式。脚本通过主系统 APT 安装，禁止联网补依赖；不要双击 DEB 进入 KARE。
  组装沿用输入 DEB 的配置来源，支持 `audit-tools` 与 `audit-tools-s1-s2`；
  默认构建优先真实配置，缺失时回退公开空表头；含真实配置的产物仅供内部使用。
  安装内部包不会覆盖既有用户配置；已启动过公开包的测试机须在设置中心导入正式配置。
  amd64 与 arm64 分别打包；目标系统缺少基础依赖时，离线安装会明确失败，须另备匹配系统的依赖包。

引擎默认**不用构建机系统 Python**，而是下载 python-build-standalone 的自包含
CPython 3.12.14（glibc 2.17 基线构建，含 Tcl/Tk）并用 PyInstaller 冻结。
**能否运行由整个冻结包的原生依赖和目标机环境决定，不能只凭 Python 基线或“麒麟 V10”名称判断**；
构建机只需 `dpkg-deb`、`tar`、`objdump`（binutils）、`sha256sum` 和网络（首次，
官方 PyPI 失败自动改用镜像）。产物文件名含构建系统、Python 版本、**实测最低
glibc** 与架构（如 `..._uos-25_py3.12.14_glibc2.25_amd64.deb`）——构建后逐个
ELF 审计 GLIBC/GLIBCXX 符号版本，超过目标基线（默认 glibc 2.28，即 UOS20）时
构建失败并列出超标文件。较老目标环境还须确认 pip 选择的二进制 wheel 基线；
仅修改 `GLIBC_TARGET` 不会降低已有依赖的 glibc 要求。

**跨机器/跨芯片**：PyInstaller 不能直接在 x86-64 Python 中生成 ARM64 程序，
不同 CPU 架构需进入对应架构的运行环境打包（实体机或 QEMU 模拟容器）；
`amd64` 与 `arm64`（飞腾/鲲鹏）打包机自动下载对应架构解释器，其它架构回退该机
系统 Python（`--python` 可指定解释器）。UOS 与麒麟（Kylin V10 等 Debian 系）
均适用：DEB 依赖只有 libc6/libcrypt1/libstdc++6/xdg-utils/iproute2，tar.gz
完全免安装。

没有 ARM64 实体机器时，可在 x86-64 Windows 的 Docker Desktop 中用 QEMU 构建。
先用 `docker run --rm --platform linux/arm64 alpine:3.20 uname -m` 确认输出
`aarch64`，再以 `docker build --platform linux/arm64 -f shell-flask/Dockerfile.arm64-builder
-t audit-tools-arm64-builder .` 建立构建镜像。容器内运行 `Linux-2-打包DEB与TARGZ.sh
--skip-tgz` 输出 arm64 DEB；`Linux-6-组装通用离线交付.sh` 配上官方
`LibreOffice_25.8.7_Linux_aarch64_deb.tar.gz` 生成完整交付包。QEMU 构建只能验证
ARM64 包结构与模拟运行，最终仍须在目标 ARM64 UOS/麒麟机器验收安装和 S1-F01。

在项目根目录的 PowerShell 中可直接使用：

```powershell
$repo = (Get-Location).Path
docker build --platform linux/arm64 -f shell-flask/Dockerfile.arm64-builder -t audit-tools-arm64-builder .
docker run --rm --platform linux/arm64 --mount "type=bind,source=$repo,target=/work" --mount 'type=volume,source=audit-tools-arm64-cache,target=/root/.cache' --env FLASK_DIST_DIR=/work/shell-flask/dist/arm64 audit-tools-arm64-builder sh shell-flask/Linux-2-打包DEB与TARGZ.sh --skip-tgz
$deb = (Get-ChildItem shell-flask/dist/arm64 -Filter '*.deb' | Sort-Object LastWriteTime -Descending | Select-Object -First 1).Name
docker run --rm --platform linux/arm64 --mount "type=bind,source=$repo,target=/work" --mount 'type=bind,source=F:\VMware_share,target=/lo,readonly' audit-tools-arm64-builder sh shell-flask/Linux-6-组装通用离线交付.sh "shell-flask/dist/arm64/$deb" /lo/LibreOffice_25.8.7_Linux_aarch64_deb.tar.gz shell-flask/dist/arm64
```

内部测试版把输出目录改成 `shell-flask/dist/arm64/internal`，给 `Linux-2` 加
`--with-local-config`，给 `Linux-6` 加 `--internal-config`；该包不得上传公开仓库。

**配置边界**：默认打包使用仓库公开 `config/` 的空表头配置（构建时仍做隐私
清理校验，机构参照/审核历史所在表只保留表头）；`--with-local-config` 从本机
不入库的 `config-real/` 读取正式配置，仅供本人验证，禁止外发。首次启动把包内
六册配置复制到 `~/.local/share/base-audit-v3/config/`，运行数据保存在同目录下；
升级或换便携包目录不会覆盖已有用户配置。LibreOffice Calc 仍需目标系统提供；
本仓库未提供使用说明 DOCX 时发行包不含该文档（放回仓库根目录或 `core/` 即可
自动带入）。

**验收状态**：2026-09-29 在 UOS Desktop 25 上完成全流程构建并对 DEB 与 tar.gz
做了解包冒烟（均 HTTP 200、端口释放干净，tar.gz 便携模式自动识别通过）。完整版和
S1/S2 版均随包保留 `lxml`；openpyxl 在支持的 XML 操作中优先使用它，其他部分仍用
标准库等实现。此前普通环境构建的 AMD64 冻结包最低 glibc 2.25（来源 lxml `getentropy`）。
无 lxml 的 AMD64 实验包曾降到 glibc 2.17，但该实验策略目前不作为默认打包策略。
禁用 lxml 后曾出现的 9 项 S3-F05 测试失败，已定位为测试模板注入代码未兼容
`<definedNames />` 空节点写法；只在内存中调整测试样例后，相关 44 项定向测试通过，
不能据此认定 S3-F05 必须依赖 lxml，测试文件尚未正式修复。
Python 运行时及原生组件仍依赖系统 glibc；此前携带 lxml 的 AMD64 产物不满足 glibc
2.23 基线。ARM64 包需独立构建、审计，不能套用 AMD64 的最低版本结论。
2026-10-16 已在 WSL2 下建立 Ubuntu Base 16.04.6 ARM64 / glibc 2.23 用户空间，
用 QEMU 8.2.2 构建完整版。全新环境为同一个 lxml 6.1.3 选择了
`cp312-cp312-manylinux2014_aarch64.manylinux_2_17_aarch64` wheel，保留 lxml 的
冻结包实际审计最低 glibc 为 2.17；较新构建环境可能选择更高基线 wheel。
构建命令为 `PIP_ONLY_BINARY=:all: GLIBC_TARGET=2.23 GLIBCXX_TARGET=3.4.21 FLASK_DIST_DIR=/evidence/dist bash shell-flask/Linux-2-打包DEB与TARGZ.sh --python /opt/python/bin/python3 --profile full`，
须在上述 ARM64 隔离环境执行。旧版 dpkg-deb 不支持 `--root-owner-group` 时，
打包脚本使用 root 临时组包目录或 fakeroot 兼容构建。
该环境中 core 全量回归为 926 passed、11 skipped、2 subtests passed，0 failures/errors；
源码/冻结程序共 10 项专项检查通过，DEB 与 tar.gz 解包启动均通过。
冻结程序通过 HTTP 接口完成 S1-F01（真实 LibreOffice 重算）、S2-F01（三张结果表）和
S3-F05（两种渲染器），并确认服务进程实际加载 lxml.etree 与 libc-2.23.so。
S3-F05 各冻结渲染器与对应源码输出完全一致；两种渲染器之间仍有 13 处表头/空值格
格式或浮点尾数差异，已记录，未在本次修改其行为。715 个原始配置/数据文件哈希未变。
证据和公开配置安装包保存在本机忽略目录
`shell-flask/dist/ARM64-glibc223-validation-20261016/`，入口为 `验证说明.txt`。
本次公式重算使用 LibreOffice 5.1.6.2，不代表 LO 25.8/26.8 安装包兼容 glibc 2.23；
该模拟验证也不替代麒麟实体机的桌面安装、系统集成和性能验收。
随后在独立 Ubuntu Base 16.04.6 AMD64 / glibc 2.23 chroot 中完成同版本完整版构建，
同样保留 lxml 6.1.3，ELF 最低 glibc 为 2.17。AMD64 core 回归为
926 passed、11 skipped、2 subtests passed，源码/冻结专项 10 项及两种包启动均通过；
原始 715 个文件哈希未变。两个架构的 129 个运行源码文件归一化换行后内容一致。
双架构目录为本机忽略的 `shell-flask/dist/dual-arch-glibc223-20261016/`，
包含 `amd64/`、`arm64/` 各自的 DEB、tar.gz、报告，以及根目录 `先看这里.txt` 和
`SHA256SUMS`。目标机 `uname -m` 为 `x86_64` 选 amd64，为 `aarch64` 选 arm64。
包内为公开空表头配置，需导入正式配置；LibreOffice 不随包提供。
UOS 20 / 麒麟 V10 / arm64 实机
安装与业务全流程仍待验收。逐笔系统的组合流程与 S1-F05「制作联合模板」在
UOS 使用共享核心的 native/openpyxl 实现；只有确实依赖 Excel/WPS COM 原生语义
的能力限于 Windows。

## 测试

```sh
cd core && PYTHONPATH=src python3 -m pytest tests/ -q
```

## 与其他发行线的关系

- 业务功能一律改 `core`，两个外壳同时生效；归档线见
  `external/Flask`、`external/pywebview2`（各含归档说明）。

## Windows 打包环境三件套（换机器必读）

Win7 兼容包依赖三样**不入 git** 的本机资源（`wheels/` 与 `.exe` 在公开仓库被
`check_public_snapshot.py` 禁止，虚拟环境绑定本机解释器），缺失时的补齐方法：

1. **`wheels/`（离线依赖缓存）**：用任意带 SSL 的 Python 在 `shell-flask/` 下执行：

   ```sh
   python -m pip download --disable-pip-version-check \
     -r requirements-win7.txt -d wheels \
     --platform win32 --python-version 3.7 --only-binary=:all:
   python -m pip download --disable-pip-version-check \
     -r requirements-win7.txt -d wheels \
     --platform win32 --python-version 3.8 --only-binary=:all:
   # pip 可能漏下 <3.8 标记的包，再单独补一次：
   python -m pip download --disable-pip-version-check \
     importlib-metadata==4.13.0 packaging==24.0 zipp==3.15.0 \
     typing-extensions==4.7.1 pyinstaller-hooks-contrib==2024.7 \
     -d wheels --platform win32 --python-version 3.7 --only-binary=:all:
   ```

2. **`redist/`（vc_redist 运行库）**：从微软官方下载 `vc_redist.x86.exe` 与
   `vc_redist_2019.x86.exe` 放入 `shell-flask/redist/`（或从旧发行包复制）。
   缺失时 Win7 包仍会生成，但不含"运行库修复"附件。

3. **解释器（PYTHON_WIN7）**：指向 3.7/3.8 的 x86 Python（如 conda 环境
   `HZPBCwin7x86`）；注意此类环境通常无 SSL，依赖只能走 wheels 离线安装。

构建中断续跑：`FLASK_DIST_DIR` 指向既有输出目录重跑 `build-win7.bat` 即可补齐；
成功后手动删除目录里的 `BUILD_INCOMPLETE.txt`。
