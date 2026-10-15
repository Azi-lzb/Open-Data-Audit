# 国产 Linux 系统 deb 包安装机制速查指南

> 适用系统：UOS 20 / 麒麟 V10 / UOS 25 / 麒麟 V11
> 来源：2026-10 在麒麟 V11 上安装 LibreOffice 25.8 实战复盘
> 结论先行：**交付统一用标准 deb，用户侧统一 `sudo apt install ./xxx.deb` 装宿主**，不为单个系统单独写 sh 脚本。

---

## 一、四系统行为对比总表

| 系统 | 双击 deb 装到哪 | 命令行 apt/dpkg | 商店通道 | 特殊机制 |
|---|---|---|---|---|
| UOS 20 | 宿主 | 宿主，无阻碍 | deb 为主 | 无 |
| 麒麟 V10 | 宿主 | 宿主，无阻碍 | deb | kysec 安全策略可能拦**未签名**包（属安检拦截，包目的地仍是宿主） |
| UOS 25 | 宿主 | 宿主，无阻碍 | 商店部分应用为**玲珑**版（进玲珑沙箱） | 玲珑是与 deb 平行的另一条通道 |
| 麒麟 V11 | **正常模式下不进宿主，改道 KARE 容器** | 宿主安装需**维护模式** | 商店原生应用为**开明包**（kaiming） | 磐石架构：不可变系统 + 开明包格式 + KARE |

关键认知：**双击 deb 在所有系统上都进宿主，唯一的例外是麒麟 V11**。V11 正常模式下宿主系统不可写，双击会被麒麟安装器收进 KARE 容器——这就是"同一个 deb，命令行装和双击装去两个不同世界"现象的根源。

---

## 二、dpkg 与 apt 的区别

| | `sudo dpkg -i` | `sudo apt install ./xxx.deb` |
|---|---|---|
| 层级 | 底层包操作工具 | 高层前端（最终还是调 dpkg） |
| 依赖解析 | ❌ 缺依赖直接报错，包留在半安装态（`dpkg -l` 显示 `iU`） | ✅ 自动解析依赖，从仓库补齐缺的库 |
| 安装顺序 | 按参数顺序，不智能 | 自动排序 |
| 仓库感知 | 无 | 有，能升级/卸载/autoremove |

- **装本地 deb 必须带 `./` 路径**：`apt install ./xxx.deb`。不带路径 apt 会把参数当远程仓库包名去搜，报"无法定位软件包"。
- dpkg 装到一半依赖报错的补救：`sudo apt -f install`。
- **统一用 apt 装本地 deb，没有反例。** dpkg 只在"离线且 deb 自带全部依赖"的场景有优势（如 LibreOffice 官方 DEBS 目录自闭环）。

---

## 三、麒麟 V11 磐石架构详解

三件套：**不可变系统 + 开明（KaiMing）包格式 + KARE 兼容运行环境**。

### 3.1 不可变系统与维护模式

- 正常模式下宿主系统只读，dpkg/apt 解包被拦截。典型报错：
  ```
  ** 当前模式禁止执行（unpack）操作**
  dpkg: 错误: 执行钩子 /usr/bin/ostree-pkgs-guard 出错，退出状态为 256
  ```
- 维护模式下系统可写，是官方允许的手工安装通道。命令：

  ```bash
  sudo mm-cli -s      # 查看当前模式：Normal Mode / Maintain Mode
  sudo mm-cli -o      # 开启维护模式（需重启生效）
  sudo reboot
  # ……在 Maintain Mode 下安装……
  sudo mm-cli -c      # 保存修改并退出维护模式（然后重启）
  ```

  ⚠️ **千万别用 `mm-cli -c -n reboot`**——那是"不保存修改直接退出"，会丢掉维护模式期间的安装。安全中心 GUI 里也有维护模式开关，效果相同。
- 维护模式期间别顺手改其他系统文件，装完就退，保持不可变保护。

### 3.2 KARE（麒麟应用运行环境）

- 基于**容器/命名空间**的兼容环境，收容传统 deb：宿主装不上的、或正常模式双击的 deb 都进这里运行。
- 入口只有**双击 deb → 麒麟安装器**，正常模式即可（走容器通道不受宿主 unpack 守卫管）。**没有公开的命令行工具**（无 kare-cli 之类）。
- KARE 里的 deb 不出现在宿主 `dpkg -l`，也不在 `kaiming list`。
- openKylin 2.0 SP2 起有跨环境文件映射、共享目录等机制，但公开文档薄，不可依赖。

### 3.3 开明包（kaiming）

- V11 商店的原生格式，设计大量参考 Flatpak：沙箱隔离 + 权限管控 + 共享运行时层（类似 UOS 25 的玲珑）。
- `kaiming list` 可看已装的开明应用和运行时（如 `top.openkylin.base`、`ukyui runtime` 就是运行时层）。
- CLI 示例：`sudo kaiming install org.xxx.App`（反向域名 ID）。

### 3.4 判断一个程序住在哪个世界（排障第一步）

```bash
dpkg -l | grep -i <程序名>     # 查得到 → 宿主
kaiming list | grep -i <程序名> # 查得到 → 开明包
两者都查不到                    # 八成在 KARE 容器里
```

---

## 四、沙箱世界的规则（跨环境调用必读）

把系统想成三座岛：**宿主 / KARE 容器 / 开明沙箱**（UOS 25 上则是宿主 / 玲珑沙箱）。

1. **岛与岛互相看不见**：KARE 看不见宿主的 `/opt`、`/usr`，也看不见开明沙箱；开明内部两个应用之间默认也不互通。
2. **只有宿主站在岛外，能看见所有岛**；方向不可逆。
3. **跨岛调用二进制：没有官方通道，不现实**。共享目录/门户只适合传数据文件，拉不起对方的进程。这是可行性问题而非性能问题——容器是 namespace 隔离不是虚拟机，就算跨过去，进程本身也近原生。
4. **程序住哪座岛，它依赖的外部程序就得装进哪座岛**，没有跨岛借用捷径。
5. 代价：同一份依赖（如 LibreOffice）可能每座岛各存一份物理拷贝，磁盘翻倍是正常现象。
6. 性能上真正的大头是 `soffice` 这类进程自身 1~3 秒的启动时间，与沙箱无关。

---

## 五、实战案例复盘：V11 安装 LibreOffice 25.8

### 5.1 完整流程

```bash
# 0. 确认架构
uname -m        # x86_64 → 下 x86-64 deb 版；aarch64 → ARM64 版

# 1. 进维护模式
sudo mm-cli -o && sudo reboot
mm-cli -s       # 确认显示 Maintain Mode

# 2. 安装（DEBS 目录自带全部依赖，可一把梭）
cd LibreOffice_25.8.7.3_Linux_x86-64_deb/DEBS/
sudo apt install ./*.deb

# 3. 退出维护模式（保存！）
sudo mm-cli -c && sudo reboot
```

中途的无害报错（可忽略）：`debconf 无法初始化前端界面：Dialog`（退回 Readline 而已）；`_apt 权限不够`沙盒提示。

### 5.2 Calc 最小包集（43 个包 → 8 个）

只需 Calc 时，DEBS 目录里只装：

```
libreoffice25.8            # 主包，提供 soffice
libobasis25.8-ure          # UNO 运行时（地基）
libobasis25.8-core         # 核心库（体积大头，省不掉）
libobasis25.8-calc + libreoffice25.8-calc
libobasis25.8-images + libobasis25.8-graphicfilter
libreoffice25.8-debian-menus   # 菜单图标，纯命令行可省
```

可跳过：Writer/Base/Impress/Draw/Math 全套、词典、beanshell/javascript 扩展、postgresql-sdbc、report-builder、librelogo、onlineupdate、ogltrans。
体积预期：全量约 600 MB，最小集约省 200~300 MB（core 大头省不掉）。**省磁盘意义不大，省操作次数才有意义**（比如往 KARE 双击装时从 43 次降到 8 次）。

### 5.3 PATH 问题

TDF 官方 deb **从不**把 `soffice` 放进 PATH、不建 `/usr/bin` 软链，只装到 `/opt/libreoffice25.8/program/soffice` 并注册菜单图标。按命令名找 LO 的程序会扑空，修复：

```bash
sudo ln -sf /opt/libreoffice25.8/program/soffice /usr/local/bin/soffice
```

若 `/usr/local/bin` 只读，进维护模式做。路径写法在宿主和容器里相同（各自解析到各自岛上的那份）。

---

## 六、打包落地建议（正式结论）

1. **交付物**：单一标准 deb，依赖全部写进 `Depends` 让 apt 在各系统上自动补齐；不按系统定制安装脚本（apt 通吃 UOS/麒麟全系 deb 系）。
2. **README 统一安装命令**：`sudo apt install ./your-app.deb`（强调带 `./`）。
3. **V11 专属说明必须写进文档**：
   - 安装前先开维护模式（安全中心 GUI 或 `sudo mm-cli -o` 重启），装完 `sudo mm-cli -c` 退出；
   - **警告用户不要双击 deb**——正常模式双击会进 KARE 容器，导致与宿主程序/依赖互相不可见。
4. **程序若调用外部二进制**（如 `soffice`）：提供"完整路径"配置项而不是只按命令名搜索；文档写明路径示例。
5. **可选自检脚本**（安装前跑，仅提示不代劳）：
   ```bash
   uname -m                          # 架构匹配检查
   mm-cli -s 2>/dev/null | grep -q Maintain && echo "✅ 维护模式" || echo "⚠️ 麒麟V11需先开维护模式"
   ```
6. **不要假设跨环境调用**：程序与它调用的外部程序必须同岛。能上宿主就都上宿主——零新增开销、零额外拷贝，总成本最低。

---

## 附：本次结论与通道对照

| 用户动作 | UOS20 / V10 / UOS25 | 麒麟 V11 |
|---|---|---|
| `sudo apt install ./x.deb`（正常模式） | ✅ 装宿主 | ❌ 被 ostree-pkgs-guard 拦截 |
| `sudo apt install ./x.deb`（维护模式） | — | ✅ 装宿主（推荐） |
| 双击 deb | ✅ 装宿主 | ⚠️ 进 KARE 容器（非宿主） |
| 商店安装 | 视应用：deb 或玲珑 | 视应用：开明包 |
| 查安装位置 | `dpkg -l` | `dpkg -l` / `kaiming list` / 都没有→KARE |
