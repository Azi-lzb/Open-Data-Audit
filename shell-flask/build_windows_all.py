"""Build one Windows release folder containing both Flask executables."""

from __future__ import annotations

import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parent
DIST_ROOT = ROOT / "dist"
CORE_SRC = ROOT.parent / "core" / "src"
if str(CORE_SRC) not in sys.path:
    sys.path.insert(0, str(CORE_SRC))

from base_audit.app_identity import executable_stem


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
        ("Windows 10/11", "windows-build/build-modern.bat"),
        ("Windows 7", "windows-build/build-win7.bat"),
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
        executable_stem("flask") + ".exe",
        executable_stem("flask", win7=True) + ".exe",
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
