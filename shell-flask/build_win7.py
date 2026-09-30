"""Win7 发行包收尾工具（由 windows-build/build-win7.bat 调用，Python 3.7/3.8 可运行）。

唯一子命令：
  exe-extras   PyInstaller 产物收尾：生成《运行环境说明-Win7.txt》（GBK，
               Win7 记事本可读）、附带「运行库修复\」离线 vc_redist，
               将离线运行库安装包放入「运行库修复」子目录。

发行形态：dist\\win7-时间戳\\ 或 dist\\windows-时间戳\\ 下的 EXE
          + core\\ + config\\ + 运行库修复\\ + 说明文档。

Win7 SP1 依赖策略（目标机按概率分层，全部离线可用）：
目标机使用系统 UCRT；裸 Win7 SP1 若缺运行库，先运行包内
「运行库修复\\vc_redist_2019.x86.exe」。不在 EXE 同级散放 DLL。
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CORE = (ROOT.parent / "core") if (ROOT.parent / "core" / "src").is_dir() else (ROOT / "core")
CONFIG_TEMPLATE = ROOT.parent / "config"
DIST = Path(os.environ.get("FLASK_DIST_DIR") or ROOT / "dist")
REDIST = ROOT / "redist"
REDIST_2019 = REDIST / "vc_redist_2019.x86.exe"   # VS2019 线 14.29，官方支持 Win7 SP1
REDIST_LATEST = REDIST / "vc_redist.x86.exe"      # VS2022 线，新系统兜底
if str(CORE / "src") not in sys.path:
    sys.path.insert(0, str(CORE / "src"))

from base_audit.app_identity import app_title, executable_stem


# GBK + CRLF：Win7 记事本（ANSI）双击即可正常阅读。
README_TEXT = """\
{app_title}（Flask 版）Windows 7 运行说明
================================================

一、本包特点
1. 32 位程序，32 位 / 64 位 Windows 7 SP1 及以上系统均可运行；
2. 本包内嵌构建时使用的 32 位 Python 3.7/3.8 与全部组件，目标机
   无需联网、无需安装 Python；
3. 依赖 Windows 系统运行库；缺少时使用本包的离线安装程序修复。

二、启动方法
1. 双击 {executable_name}
   （exe 必须与本目录的 core 文件夹保持同级，升级 exe 时不要覆盖 core）；

三、常见问题
1. 双击后闪退、无反应，或提示缺少 api-ms-win-crt-*.dll 等运行库：
   打开本包的「运行库修复」文件夹，双击里面的
   vc_redist_2019.x86.exe 安装一次（离线安装包，无需联网），
   完成后重新启动程序即可。若还有问题，可再装 Windows 7 补丁
   KB2999226（效果相同）。
2. 文件选择：未内嵌 Tk 界面库的包会自动使用"浏览器内置"选目录模式，
   也可在 高级设置 中随时切换。
3. 计算引擎：完整审核需本机已安装 Microsoft Excel 或 WPS 表格（COM）。
4. 请勿把程序放在无写入权限的目录（如 C:\\Program Files），
   历史数据与用户设置保存在 core\\data 下，升级时注意保留。
"""


def write_readme(path: Path) -> None:
    with open(path, "w", encoding="gbk", newline="\r\n") as handle:
        handle.write(README_TEXT.format(
            app_title=app_title(),
            executable_name=executable_stem("flask", win7=True) + ".exe",
        ))


def copy_redist_repair(target: Path) -> None:
    """附带微软官方离线运行库修复包（裸 Win7 SP1 双击一次即修复）。

    优先 VS2019 线（14.29，官方支持 Win7 SP1）；缺失时退 VS2022 线。
    """
    repair = target / "运行库修复"
    repair.mkdir(parents=True, exist_ok=True)
    shipped = []
    for redist in (REDIST_2019, REDIST_LATEST):
        if redist.is_file():
            shutil.copy2(redist, repair / redist.name)
            shipped.append(redist.name)
    if not shipped:
        print("[提示] redist\\ 下没有 vc_redist，跳过运行库修复附件。")
        return
    print(f"[完成] 运行库修复附件：{'、'.join(shipped)}（Win7 装较早的那个）。")


def exe_extras() -> int:
    """在当前 Win7 独立目录或双 Windows 合并目录收尾。"""
    dist_root = (ROOT / "dist").resolve()
    if not os.environ.get("FLASK_DIST_DIR") or DIST.resolve().parent != dist_root or not DIST.name.startswith(("win7-", "windows-")):
        raise SystemExit("[错误] Win7 输出目录必须是 dist\\win7-时间戳\\ 或 dist\\windows-时间戳\\。")
    if not DIST.is_dir():
        raise SystemExit(f"[错误] 未找到打包输出目录：{DIST}")
    copy_redist_repair(DIST)
    write_readme(DIST / "运行环境说明-Win7.txt")
    print(f"[完成] 已生成 {DIST / '运行环境说明-Win7.txt'}（运行库附件位于子目录）。")
    return 0


def main(argv: list[str]) -> int:
    """唯一入口：exe-extras（便携版模式已移除）。"""
    if len(argv) == 2 and argv[1] == "portable":
        print("[错误] 便携版模式已移除（2026-09-21 用户要求：发行只保留 EXE 形态）。\n"
              "       请运行 Windows-2-打包双版本.bat 生成两个 EXE + core + config。")
        return 2
    if len(argv) != 2 or argv[1] != "exe-extras":
        print(__doc__)
        return 2
    return exe_extras()


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
