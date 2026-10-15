"""Build one Windows release folder containing both Flask executables."""

from __future__ import annotations

import argparse
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

from base_audit.app_identity import executable_stem, load_product_info
from distribution_profile import DEFAULT_CONFIG, embed_profile, filter_release_config, get_profile, load_config


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--profile", default="full", help="distribution-profiles.json profile (default: full)")
    args = parser.parse_args(argv)
    try:
        config = load_config(DEFAULT_CONFIG)
        profile = get_profile(config, args.profile)
        version = load_product_info()["version"]
    except ValueError as exc:
        print("[ERROR] {}".format(exc))
        return 2

    print("[profile] {}: {}".format(args.profile, ", ".join(profile["enabledSystems"])), flush=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_base = "windows-" + stamp + profile["outputSuffix"]
    output = DIST_ROOT / output_base
    suffix = 2
    while output.exists():
        output = DIST_ROOT / (output_base + "-" + str(suffix))
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

    # Both builders replace the shared frontend. Apply the profile only after
    # the Win7 build has finished, to the release copy rather than source files.
    try:
        filter_release_config(output / "config", config, args.profile)
        embed_profile(output / "core/frontend/web/index.html", config, args.profile)
        (output / "distribution-profile.txt").write_text(args.profile + "\n", encoding="utf-8")
    except (OSError, ValueError) as exc:
        print("[ERROR] Cannot apply delivery profile: {}".format(exc))
        return 1
    # Publish a stable versioned name only after both builds are complete.
    # Keep the previous working delivery until then, including on build failure.
    target = DIST_ROOT / ("Windows-Flask_V" + version + profile["outputSuffix"])
    archived = None
    try:
        if target.resolve().parent != DIST_ROOT.resolve() or target.is_symlink():
            raise ValueError("Unsafe Windows delivery directory")
        if target.exists():
            archive_root = DIST_ROOT / ("归档-旧发行包-" + datetime.now().strftime("%Y%m%d"))
            if archive_root.resolve().parent != DIST_ROOT.resolve():
                raise ValueError("Unsafe Windows archive directory")
            archive_root.mkdir(exist_ok=True)
            archived = archive_root / (target.name + "-" + stamp)
            number = 2
            while archived.exists():
                archived = archive_root / (target.name + "-" + stamp + "-" + str(number))
                number += 1
            target.rename(archived)
        try:
            output.rename(target)
        except OSError:
            if archived is not None:
                archived.rename(target)
            raise
        output = target
        (output / incomplete.name).unlink()
    except (OSError, ValueError) as exc:
        print("[ERROR] Cannot publish delivery folder: {}".format(exc))
        return 1
    if archived is not None:
        print("[archive] Previous delivery: {}".format(archived))
    print("[DONE] Ship this one folder: {}".format(output))
    return 0


if __name__ == "__main__":
    sys.exit(main())
