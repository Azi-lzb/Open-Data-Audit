"""Build or locate the platform-native S3-F01 Rust writer for release bundling."""

from __future__ import annotations

import argparse
import os
import shlex
import shutil
import subprocess
import sys
from pathlib import Path


CRATE = Path(__file__).resolve().parent
TARGETS = {
    "windows-amd64": ("x86_64-pc-windows-gnu", "s3-comparison-writer.exe"),
    "linux-amd64": ("x86_64-unknown-linux-musl", "s3-comparison-writer"),
}


def _fresh(binary: Path) -> bool:
    if not binary.is_file():
        return False
    sources = [CRATE / "Cargo.toml", CRATE / "Cargo.lock", *sorted((CRATE / "src").glob("*.rs"))]
    return binary.stat().st_mtime >= max((path.stat().st_mtime for path in sources if path.is_file()), default=0)


def _wsl_crate_path() -> str:
    drive, tail = os.path.splitdrive(str(CRATE.resolve()))
    if not drive or len(drive) != 2 or drive[1] != ":":
        raise RuntimeError(f"项目路径无法映射到 WSL：{CRATE}")
    relative_path = tail.lstrip("\\/").replace("\\", "/")
    return f"/mnt/{drive[0].lower()}/{relative_path}"


def _build(platform_id: str, target: str) -> tuple[bool, str]:
    if sys.platform == "win32" and platform_id == "windows-amd64":
        cargo = shutil.which("cargo.exe") or shutil.which("cargo")
        if cargo:
            completed = subprocess.run(
                [cargo, "build", "--quiet", "--release", "--target", target],
                cwd=CRATE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
            )
        else:
            wsl = shutil.which("wsl.exe") or shutil.which("wsl")
            if not wsl:
                return False, "没有 Rust cargo 或 WSL；Windows 版 Rust 写入器未构建"
            try:
                linux_crate = _wsl_crate_path()
            except RuntimeError as exc:
                return False, str(exc)
            script = f"cd {shlex.quote(linux_crate)} && cargo build --quiet --release --target {target}"
            distro = os.environ.get("AUDITTOOLS_WSL_DISTRO", "Ubuntu")
            completed = subprocess.run(
                [wsl, "--distribution", distro, "--exec", "bash", "-lc", script],
                cwd=CRATE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
            )
        if completed.returncode:
            detail = (completed.stderr + completed.stdout).decode("utf-8", errors="replace").strip()
            return False, detail or f"cargo build 退出码 {completed.returncode}"
        return True, ""

    if sys.platform.startswith("linux") and platform_id == "linux-amd64":
        cargo = shutil.which("cargo")
        if not cargo:
            return False, "没有 Rust cargo；Linux 版 Rust 写入器未构建"
        if not (shutil.which("musl-gcc") or shutil.which("x86_64-linux-musl-gcc")):
            return False, "没有 musl-gcc；无法生成不依赖目标机 glibc 的 Linux AMD64 写入器"
        completed = subprocess.run(
            [cargo, "build", "--quiet", "--release", "--target", target],
            cwd=CRATE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        if completed.returncode:
            detail = (completed.stderr + completed.stdout).decode("utf-8", errors="replace").strip()
            return False, detail or f"cargo build 退出码 {completed.returncode}"
        return True, ""

    return False, f"当前构建平台 {sys.platform} 不支持 {platform_id} 目标"


def locate_or_build(platform_id: str) -> tuple[Path | None, str]:
    try:
        target, filename = TARGETS[platform_id]
    except KeyError:
        return None, f"不支持的平台目标：{platform_id}"
    binary = CRATE / "target" / target / "release" / filename
    if _fresh(binary):
        return binary, ""
    ok, reason = _build(platform_id, target)
    if ok and binary.is_file():
        return binary, ""
    if ok:
        reason = f"cargo 构建完成，但未生成预期文件：{binary}"
    return None, reason


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("platform", choices=tuple(TARGETS))
    parser.add_argument("--required", action="store_true", help="没有构建结果时返回错误")
    args = parser.parse_args()
    binary, reason = locate_or_build(args.platform)
    if binary is None:
        print(f"[提示] S3 Rust 写入器未随本次发行包提供：{reason}", file=sys.stderr)
        return 1 if args.required else 0
    print(binary)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
