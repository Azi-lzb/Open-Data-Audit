"""Build one Windows release folder containing both Flask executables."""

from __future__ import annotations

import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parent
DIST_ROOT = ROOT / "dist"


def main() -> int:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = DIST_ROOT / ("windows-" + stamp)
    suffix = 2
    while output.exists():
        output = DIST_ROOT / ("windows-" + stamp + "-" + str(suffix))
        suffix += 1
    output.mkdir(parents=True)
    incomplete = output / "BUILD_INCOMPLETE.txt"
    incomplete.write_text(
        "Build is incomplete. Do not distribute this folder.\n", encoding="utf-8"
    )

    env = os.environ.copy()
    env["FLASK_DIST_DIR"] = str(output)
    for label, script in (
        ("Windows 10/11", "打包Flask.bat"),
        ("Windows 7", "打包Flask-Win7.bat"),
    ):
        print("[build] {} -> {}".format(label, output), flush=True)
        result = subprocess.run(
            ["cmd.exe", "/d", "/c", str(ROOT / script)],
            cwd=str(ROOT),
            env=env,
            stdin=subprocess.DEVNULL,
        )
        if result.returncode:
            print("[ERROR] {} build failed; partial output: {}".format(label, output))
            return result.returncode

    required = (
        "基础数据审核工具_Flask.exe",
        "基础数据审核工具_Flask_Win7兼容.exe",
        "core/frontend/web/index.html",
        "config",
        "运行库修复",
        "运行环境说明-Win7.txt",
    )
    missing = [name for name in required if not (output / name).exists()]
    if missing:
        print("[ERROR] Incomplete release folder: {}".format(", ".join(missing)))
        return 1

    incomplete.unlink()
    print("[DONE] Ship this one folder: {}".format(output))
    return 0


if __name__ == "__main__":
    sys.exit(main())
