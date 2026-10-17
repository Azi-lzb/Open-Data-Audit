# 麒麟 V11 LibreOffice 安装与 S1 计算记录（KARE 调查）

更新日期：2026-10-15

> **结论修订：** 用户已在银河麒麟 V11 实机验证，切换到 Maintain Mode 并重启后，`sudo apt install ./*.deb` 可以安装 LibreOffice；安装到主系统的 LO 可被审核工具直接发现，S1-F1 实测成功。V11 的内置安装入口应调用主系统 APT；无需走 KARE，也不需要为 S1-F1 增加 V11 专用调用。下方 KARE 内容是此前假设和模拟调查的过程记录，不再是当前安装步骤或验收前置条件。

## 早期 KARE 假设与模拟调查（已被上方实机结果取代）

麒麟 V11 的磐石架构通过 KARE 兼容旧系统应用；系统主环境中的 `soffice` 不一定能看到兼容环境内的 LibreOffice。因此把“安装到 KARE 环境”和“从 S1-F1 调用 KARE 里的 Calc”看成两个独立环节。

本机 WSL2 + Docker 已完成离线运行验证：LibreOffice 25.8.7 x86-64 DEB 安装到 Ubuntu 20.04 模拟容器后，可以无头重算 XLSX，也能把 DOCX 转 PDF；项目 S1-F1 `LibreOfficeCalculator` 通过模拟 `kare run <app-id>` 完成重算，缓存结果为 6，输入文件 SHA-256 未改变。

在这段早期调查期间，项目安装器调用主系统 `pkexec dpkg --install`；附件日志只记录返回码 255，没有 dpkg/pkexec 的错误正文，不能据此断言具体原因。后续的 V11 实机结论见本文件末尾。

## 已完成的项目适配

`core/src/base_audit/engines/libreoffice.py` 增加显式 KARE 引擎。S1-F1 在麒麟 V11 使用时需给启动进程提供：

```bash
export BASE_AUDIT_ENGINE=kare
export BASE_AUDIT_KARE_APP_ID='<实际安装后查询到的 LibreOffice 应用 ID>'
export BASE_AUDIT_KARE_UPPER_ROOT='/opt/kare-applications/<实际环境>/upper'
```

其中 `BASE_AUDIT_KARE_UPPER_ROOT` 可在只有一个 `upper` 目录时省略。程序把临时工作簿、输出和独立 LO profile 放进 upper backing directory，再把路径转换为 KARE 进程内的 `/home/<用户>/...` 路径后调用：

```text
kare run <app-id> --headless --nologo --nodefault --nolockcheck \
  --nofirststartwizard -env:UserInstallation=file:///home/<用户>/... \
  --convert-to "xlsx:Calc MS Excel 2007 XML" --outdir /home/<用户>/... \
  file:///home/<用户>/.../audit_input.xlsx
```

没有配置 `BASE_AUDIT_KARE_APP_ID` 时，不会猜测或自动运行未知 KARE 应用。桌面启动器如何持久设置这三个变量，要在取得 V11 的实际 app ID、upper 路径和权限后再接入；当前不建议把临时终端环境变量当成面向最终用户的配置方式。

## WSL2 + Docker 本机复测

验证脚本和临时样本保存在被 Git 忽略的 `core/data/kare-v11-sim/`，不包含真实配置。进入 WSL Ubuntu 后运行：

```bash
cd /path/to/open-data-audit/core/data/kare-v11-sim
./run-docker-smoke.sh
```

该脚本构建 Ubuntu 20.04 x86-64 镜像，在容器中安装官方 LO 25.8.7 DEB 及 Ubuntu 仓库中的运行库，然后禁网运行。容器 rootfs 只读，输入只读挂载，输出单独挂载；`/opt/kare-applications/v10sp1/upper` 用 tmpfs 模拟可写 upper 层，`/home/kare` 对应其容器内路径。

脚本通过以下断言才算通过：Calc 版本可读取、XLSX 公式缓存结果为 6、DOCX 输出 PDF 文件头有效、源工作簿 SHA-256 未变化、项目真实计算适配器报告的来源为 `kare run org.libreoffice.libreoffice`。

## 当前 VMware V11 虚拟机检查状态

已找到本机的银河麒麟 V11 虚拟机文件，`vmrun` 显示虚拟机运行中；Guest Tools 状态为 `unknown`，没有可用 SSH。用户提供 VNC 访问后，已成功认证并读取到 Guest 桌面，但画面停留在麒麟启动画面；随后 VNC 端口停止响应，当前无法继续控制 Guest。没有到达登录界面，也没有传输或安装 LO。KARE 环境、应用列表、upper 路径和预检命令已由用户终端截图确认，详见下方“V11 Guest 已确认信息”。

在 WSL 中的基本调用形式如下。VMX 路径和传输文件路径需用 `wslpath` 转换；`vmrun` 的 Guest 操作还需要 VMware Tools 正常运行及 Guest 账户认证：

```bash
VMRUN='/mnt/c/Program Files/VMware/VMware Workstation/vmrun.exe'
VMX='/path/to/VMware/Virtual Machines/Ubuntu -kyrinV11/Ubuntu -kyrinV11.vmx'
VMX_WIN="$(wslpath -w "$VMX")"

# 若虚拟机未运行则启动；已运行时跳过此行：
"$VMRUN" -T ws start "$VMX_WIN" gui
"$VMRUN" -T ws checkToolsState "$VMX_WIN"
"$VMRUN" -T ws getGuestIPAddress "$VMX_WIN" -wait

# 示例：把 LO 压缩包传给 Guest（需把变量替换为 Guest 登录信息）。
# 密码只在本地交给 vmrun，不要写入项目文件或 shell 历史。
LO_ARCHIVE='/path/to/LibreOffice_25.8.7_Linux_x86-64_deb.tar.gz'
read -r -p 'Guest 用户名: ' GUEST_USER
read -r -s -p 'Guest 密码: ' GUEST_PASSWORD; printf '\n'
GUEST_HOME="/home/$GUEST_USER"
KARE_UPPER='/opt/kare-applications/v10sp1/upper' # 替换为 Guest 中实际查到的目录
KARE_STAGE="$KARE_UPPER$GUEST_HOME/kare-probe"
"$VMRUN" -T ws -gu "$GUEST_USER" -gp "$GUEST_PASSWORD" runProgramInGuest \
  "$VMX_WIN" /usr/bin/bash -lc "mkdir -p '$KARE_STAGE/out'"
"$VMRUN" -T ws -gu "$GUEST_USER" -gp "$GUEST_PASSWORD" \
  CopyFileFromHostToGuest "$VMX_WIN" "$(wslpath -w "$LO_ARCHIVE")" \
  "$KARE_STAGE/LibreOffice_25.8.7_Linux_x86-64_deb.tar.gz"

# 转换调用示例：假设真实 app ID 已确认且 KARE 能访问该 home 路径。
# 先在 Guest 准备一个测试 DOCX；输出留在 upper backing directory 中，便于检查。
PROBE_DOCX='/path/to/open-data-audit/core/data/kare-v11-sim/input/writer-probe.docx'
"$VMRUN" -T ws -gu "$GUEST_USER" -gp "$GUEST_PASSWORD" CopyFileFromHostToGuest \
  "$VMX_WIN" "$(wslpath -w "$PROBE_DOCX")" "$KARE_STAGE/input.docx"
KARE_APP_ID='替换为真实LibreOffice-app-id'
"$VMRUN" -T ws -gu "$GUEST_USER" -gp "$GUEST_PASSWORD" runProgramInGuest \
  "$VMX_WIN" /usr/bin/bash -lc \
  "kare run '$KARE_APP_ID' --headless --nologo --nodefault --nolockcheck --nofirststartwizard -env:UserInstallation=file://$GUEST_HOME/kare-probe/profile --convert-to pdf --outdir '$GUEST_HOME/kare-probe/out' '$GUEST_HOME/kare-probe/input.docx'"
"$VMRUN" -T ws -gu "$GUEST_USER" -gp "$GUEST_PASSWORD" CopyFileFromGuestToHost \
  "$VMX_WIN" "$KARE_STAGE/out/input.pdf" \
  "$(wslpath -w /tmp/kare-v11-result/writer-probe.pdf)"

# 在 Guest 执行只读诊断并将结果文件复制回 WSL。
"$VMRUN" -T ws -gu "$GUEST_USER" -gp "$GUEST_PASSWORD" runProgramInGuest \
  "$VMX_WIN" /usr/bin/bash -lc \
  '{ cat /etc/os-release; uname -m; kare --help; kare run --help; find /opt/kare-applications -maxdepth 4 -type d -name upper -print; } > /tmp/kare-v11-diagnostic.txt 2>&1'
mkdir -p /tmp/kare-v11-result
"$VMRUN" -T ws -gu "$GUEST_USER" -gp "$GUEST_PASSWORD" CopyFileFromGuestToHost \
  "$VMX_WIN" '/tmp/kare-v11-diagnostic.txt' \
  "$(wslpath -w /tmp/kare-v11-result/kare-v11-diagnostic.txt)"
unset GUEST_PASSWORD
```

`vmrun` 文档要求认证参数放在命令前。示例假设用户 home 可写并由 KARE 映射到相同路径；实际 KARE 是否暴露 Guest home、upper 是否允许普通用户写、命令是否需要额外分隔符，必须在 V11 Guest 验证。KARE 安装命令也需实际查证后再放入安装器；不能从 WSL 主机直接执行 Guest 内的 `kare`。

### V11 Guest 已确认信息（2026-10-15）

用户从 Guest 桌面终端提供的截图确认：

- `/etc/os-release` 标识为银河麒麟桌面操作系统 V11，`uname -m` 为 `x86_64`。
- `kare -E` 可运行，输出 `v10sp1` 与 `shadow` 两个环境。
- `find /opt/kare-applications -maxdepth 5 -type d -name upper -print` 找到 `/opt/kare-applications/v10sp1/upper`。此前失败是 `find` 参数之间缺少空格，不是 KARE 报错。
- `kare -e` 单独执行提示 `argument -e: expected one argument`，说明该选项必须带环境名；不能把它当作无参数查询。
- `kare run --help` 显示 `run_args` 和 `--redirect_io`。这说明运行参数由 KARE 转交，但尚不能单凭帮助输出判断 LibreOffice 应该以 app ID 还是 `soffice` 命令启动。
- `kare -l` 只显示表头和分隔线，没有已登记应用；`kare -S libreoffice` 也未显示结果。当前 KARE 环境中尚无已登记的 LibreOffice。
- `kare try-install --help` 给出准确格式 `kare try-install --env ENV deb_file`，且根帮助说明该子命令只检查 DEB 与环境、不执行安装。实际安装仍未执行。

当前 S1 KARE 适配仍需 `BASE_AUDIT_KARE_APP_ID` 和 upper 路径。已确认的路径可先用于诊断：

```bash
export BASE_AUDIT_ENGINE=kare
export BASE_AUDIT_KARE_UPPER_ROOT=/opt/kare-applications/v10sp1/upper
```

暂不设置 `BASE_AUDIT_KARE_APP_ID`，因为当前 KARE 应用列表为空，LO 的运行入口尚不存在。CLI 帮助和列表已确认；下一步先在 Guest 找到已经传入的 LibreOffice DEB：

```bash
find ~/Downloads -type f -iname '*.deb' -print
```

把 LO 的 x86-64 DEB 文件传入 Guest 后，可先对 DEB 执行不安装预检：

```bash
kare try-install --env v10sp1 "/实际路径/DEBS/LibreOffice主程序.deb"
```

`try-install` 是单个 DEB 的试安装检查，不会真的安装。LibreOffice 包之间有依赖关系，单包检查不能代表完整包集的结果；安装器会先校验和暂存完整 DEBS 包集，再按 V11 根帮助显示的 `-e ENV -i INSTALL [INSTALL ...]` 一次交给 KARE 安装。

## 当前安装入口与实机结果

麒麟 V11 的当前路径是安装到主系统，而不是 KARE。系统需要先由管理员切换到 Maintain Mode 并重启；据用户在目标机验证，切换方式为先用 `sudo mm-cli -s` 查看状态，再运行 `sudo mm-cli -o` 并按系统提示重启。不要让审核工具自动切换维护模式或重启。

终端安装示例：

```bash
cd -- "/实际路径/DEBS"
sudo apt install ./*.deb
```

审核工具内置安装入口已调整为通过 `pkexec apt-get install` 一次安装整组已检查的 DEB，让系统 APT 使用本机依赖状态和已配置软件源完成处理；安装后读取主系统 LibreOffice 版本。若 APT 不存在，才回退到 `dpkg`。这段内置流程尚需用户在 V11 实机实际点选并安装后做最终验收。

用户已报告在 V11 主系统安装后，S1-F1 能直接找到 LibreOffice，实际测试成功。因此这次不要求用户修改 S1 配置或启用 `BASE_AUDIT_ENGINE=kare`。代码中保留的 KARE 适配仅是可选路径，V11 默认安装和计算都不依赖它。

### 内置安装入口最终验收

1. 在 V11 Maintain Mode 重启后，使用“安装 LibreOffice”选择完整 DEBS 目录或官方 tar.gz。
2. 检查显示的版本、架构和包数量；授权后确认 APT 完成安装且安装日志无未满足依赖。
3. 确认安装器读出的版本与包版本一致，并在高级设置/环境信息中看到可用 LibreOffice。
4. 使用现有测试工作簿跑 S1-F1，核对审核结果和输出文件；确认原始输入未被覆盖。
5. 若目标机离线，先确认所有系统依赖都已满足或已能从本机 APT 缓存/软件源取得；`.deb` 包本身完整不代表它包含了所有发行版系统依赖。

此前 VNC 曾停在系统启动画面，导致当时的 KARE 安装计划无法由本机代为执行。该状态记录仅说明当时我们没有远程操作成功；后续用户实机结果已解决了“是否必须使用 KARE”的关键问题。

并提供 `kare list`（若帮助中存在该命令）的输出。不要发送账号密码；有 VMware Tools 后可由本机 `vmrun` 或 SSH 密钥完成文件传输。
