"""Safe, Linux-only installation of LibreOffice Calc DEB release bundles.

This module deliberately does not download packages, verify signatures, or try
to repair package dependencies.  It accepts a user-selected LibreOffice DEB
tar.gz or an already extracted ``DEBS`` directory, validates the package
metadata, stages private copies, and asks the system package manager to install
those exact copies.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import tarfile
import tempfile
import threading
import traceback
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


DPKG_EXECUTABLE = Path("/usr/bin/dpkg")
DPKG_DEB_EXECUTABLE = Path("/usr/bin/dpkg-deb")
PKEXEC_EXECUTABLE = Path("/usr/bin/pkexec")

MAX_ARCHIVE_MEMBERS = 20000
MAX_SOURCE_DIRECTORY_ENTRIES = 2000
MAX_DEB_PACKAGES = 1000
MAX_SINGLE_FILE_BYTES = 2 * 1024 * 1024 * 1024
MAX_TOTAL_ARCHIVE_BYTES = 8 * 1024 * 1024 * 1024
MAX_TOTAL_DEB_BYTES = 4 * 1024 * 1024 * 1024
MIN_FREE_SPACE_BYTES = 64 * 1024 * 1024
DPKG_DEB_TIMEOUT_SECONDS = 30
ARCHITECTURE_TIMEOUT_SECONDS = 10
SOFFICE_VERSION_TIMEOUT_SECONDS = 30

_PACKAGE_NAME_RE = re.compile(r"^[a-z0-9][a-z0-9+.-]{0,254}$")
_ARCHITECTURE_RE = re.compile(r"^[a-z0-9][a-z0-9-]{0,63}$")
_VERSION_RE = re.compile(r"^[0-9][A-Za-z0-9.+:~_-]{0,127}$")
_NUMERIC_VERSION_RE = re.compile(r"^([0-9]+(?:\.[0-9]+)*)")
_SOFFICE_VERSION_RE = re.compile(r"LibreOffice\s+([0-9]+(?:\.[0-9]+)+)", re.IGNORECASE)


@dataclass(frozen=True)
class _StagedPackage:
    name: str
    version: str
    architecture: str
    path: Path
    sha256: str
    size: int


class _PreparationError(Exception):
    """An expected input or environment problem with a user-facing message."""

    def __init__(self, message: str, detail: str = "") -> None:
        super().__init__(message)
        self.message = message
        self.detail = detail


def _system_command_env() -> dict[str, str]:
    """Give host tools the host loader path, never the frozen app's bundled one."""
    env = os.environ.copy()
    original = env.pop("LD_LIBRARY_PATH_ORIG", None)
    if original:
        env["LD_LIBRARY_PATH"] = original
    else:
        env.pop("LD_LIBRARY_PATH", None)
    return env


def _is_linux() -> bool:
    return sys.platform.startswith("linux")


def _is_root() -> bool:
    try:
        return hasattr(os, "geteuid") and os.geteuid() == 0
    except OSError:
        return False


def _is_executable(path: Path) -> bool:
    try:
        return path.is_file() and os.access(str(path), os.X_OK)
    except OSError:
        return False


def _tool_support() -> tuple[bool, str]:
    """Return whether safe DEB inspection and an authorized install are possible."""
    if not _is_linux():
        return False, "此安装方式仅适用于 Linux。"
    missing = []
    if not _is_executable(DPKG_EXECUTABLE):
        missing.append("dpkg")
    if not _is_executable(DPKG_DEB_EXECUTABLE):
        missing.append("dpkg-deb")
    if missing:
        return False, "系统缺少 {}，无法安装 DEB 软件包。请使用系统软件管理器安装 LibreOffice，或联系管理员补齐系统工具。".format("、".join(missing))
    if not _is_root() and not _is_executable(PKEXEC_EXECUTABLE):
        return False, "当前系统没有可用的 pkexec 管理员授权程序。请使用系统软件管理器安装 LibreOffice。"
    return True, ""


def _architecture_from_dpkg() -> str:
    try:
        completed = subprocess.run(
            [str(DPKG_EXECUTABLE), "--print-architecture"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=ARCHITECTURE_TIMEOUT_SECONDS,
            check=False,
            env=_system_command_env(),
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise _PreparationError(
            "无法读取系统的 DEB 架构。请确认 dpkg 可用，或使用系统软件管理器安装 LibreOffice。",
            repr(exc),
        )
    value = (completed.stdout or "").strip()
    if completed.returncode != 0 or not _ARCHITECTURE_RE.fullmatch(value):
        raise _PreparationError(
            "无法识别系统的 DEB 架构。请确认系统软件包管理器工作正常后重试。",
            "dpkg --print-architecture exited {}: {}".format(completed.returncode, (completed.stderr or "")[:1000]),
        )
    return value


def _normalise_archive_name(name: str, is_directory: bool) -> str:
    if not name or len(name) > 4096 or "\x00" in name or "\\" in name:
        raise _PreparationError("归档中含有不安全的文件路径。请重新下载官方 LibreOffice Linux DEB 压缩包。", repr(name[:200]))
    if name.startswith("/") or re.match(r"^[A-Za-z]:", name):
        raise _PreparationError("归档中含有指向系统目录的文件路径，已停止检查。请重新下载官方 LibreOffice Linux DEB 压缩包。", repr(name[:200]))
    raw_parts = name.split("/")
    if is_directory and raw_parts and raw_parts[-1] == "":
        raw_parts = raw_parts[:-1]
    parts = []
    for index, part in enumerate(raw_parts):
        if part == "" or part == "..":
            raise _PreparationError("归档中含有越出压缩包目录的文件路径。请重新下载官方 LibreOffice Linux DEB 压缩包。", repr(name[:200]))
        if part == ".":
            if index == 0:
                continue
            raise _PreparationError("归档中含有不规范的文件路径。请重新下载官方 LibreOffice Linux DEB 压缩包。", repr(name[:200]))
        parts.append(part)
    if not parts:
        if is_directory:
            return ""
        raise _PreparationError("归档中含有空文件路径。请重新下载官方 LibreOffice Linux DEB 压缩包。", repr(name[:200]))
    if len(parts) > 8:
        raise _PreparationError("归档目录层级过深，无法安全检查。请重新下载官方 LibreOffice Linux DEB 压缩包。", repr(name[:200]))
    return "/".join(parts)


def _assert_free_space(directory: Path, required_bytes: int) -> None:
    try:
        free_bytes = shutil.disk_usage(str(directory)).free
    except OSError as exc:
        raise _PreparationError("无法检查 LibreOffice 暂存空间。请确认应用数据目录可用后重试。", repr(exc))
    required_with_reserve = required_bytes + MIN_FREE_SPACE_BYTES
    if required_with_reserve > free_bytes:
        raise _PreparationError(
            "可用磁盘空间不足以暂存 LibreOffice 安装包。请释放至少 {} MB 空间后重试。".format(
                max(1, (required_with_reserve - free_bytes + 1024 * 1024 - 1) // (1024 * 1024))
            ),
            "required={} free={}".format(required_with_reserve, free_bytes),
        )


def _copy_stream(source, destination: Path, expected_size: int, total_limit: int) -> tuple[str, int]:
    if expected_size < 0 or expected_size > MAX_SINGLE_FILE_BYTES or expected_size > total_limit:
        raise _PreparationError("安装包中的单个文件超过安全大小限制。请重新下载完整的 LibreOffice DEB 安装包。", str(expected_size))
    digest = hashlib.sha256()
    copied = 0
    try:
        with destination.open("xb") as target:
            while True:
                block = source.read(1024 * 1024)
                if not block:
                    break
                copied += len(block)
                if copied > expected_size or copied > total_limit:
                    raise _PreparationError("安装包实际大小超出检查范围。请重新下载完整的 LibreOffice DEB 安装包。")
                digest.update(block)
                target.write(block)
        if copied != expected_size:
            raise _PreparationError("安装包文件内容不完整。请重新下载完整的 LibreOffice DEB 安装包。", "expected={} copied={}".format(expected_size, copied))
        try:
            os.chmod(str(destination), 0o400)
        except OSError:
            pass
    except _PreparationError:
        try:
            destination.unlink()
        except OSError:
            pass
        raise
    except OSError as exc:
        try:
            destination.unlink()
        except OSError:
            pass
        raise _PreparationError("无法把安装包复制到安全暂存目录。请检查磁盘空间和文件权限。", repr(exc))
    return digest.hexdigest(), copied


def _copy_folder_file(
    directory: Path,
    filename: str,
    destination: Path,
    expected_stat,
    directory_fd: int | None = None,
) -> tuple[str, int]:
    flags = os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0)
    fd = None
    if directory_fd is not None and os.name == "posix":
        try:
            fd = os.open(filename, flags, dir_fd=directory_fd)
        except (OSError, TypeError):
            fd = None
    if fd is None:
        try:
            fd = os.open(str(directory / filename), flags)
        except OSError as exc:
            raise _PreparationError("所选 DEBS 文件夹中的安装包无法安全读取。请重新选择文件夹。", "{}: {}".format(filename, repr(exc)))
    try:
        opened = os.fstat(fd)
        if not stat.S_ISREG(opened.st_mode):
            raise _PreparationError("DEBS 文件夹中含有非普通安装包文件，已停止检查。请重新选择完整的官方包文件夹。", filename)
        identity_changed = False
        if expected_stat is not None:
            identity_changed = expected_stat.st_size != opened.st_size
            if os.name == "posix":
                identity_changed = identity_changed or (
                    getattr(expected_stat, "st_dev", None) != opened.st_dev
                    or getattr(expected_stat, "st_ino", None) != opened.st_ino
                )
        if identity_changed:
            raise _PreparationError("DEBS 文件夹中的文件在检查期间发生变化。请重新选择安装包。", filename)
        with os.fdopen(fd, "rb") as source:
            fd = None
            return _copy_stream(source, destination, opened.st_size, MAX_TOTAL_DEB_BYTES)
    finally:
        if fd is not None:
            try:
                os.close(fd)
            except OSError:
                pass


def _find_debs_directory(source: Path) -> Path:
    try:
        source_stat = source.lstat()
    except OSError as exc:
        raise _PreparationError("所选文件夹不存在或无法读取。请重新选择 LibreOffice 的 DEBS 文件夹。", repr(exc))
    if stat.S_ISLNK(source_stat.st_mode) or not stat.S_ISDIR(source_stat.st_mode):
        raise _PreparationError("所选位置不是普通文件夹。请选择 DEBS 文件夹、其直接上级文件夹，或单层解压目录。")
    if source.name.casefold() == "debs":
        return source

    immediate = source / "DEBS"
    try:
        immediate_stat = immediate.lstat()
    except FileNotFoundError:
        immediate_stat = None
    except OSError as exc:
        raise _PreparationError("无法读取所选文件夹中的 DEBS 子文件夹。请重新选择安装包。", repr(exc))
    if immediate_stat is not None:
        if stat.S_ISLNK(immediate_stat.st_mode) or not stat.S_ISDIR(immediate_stat.st_mode):
            raise _PreparationError("所选文件夹中的 DEBS 项不是普通文件夹。请重新选择 LibreOffice 安装包。")
        return immediate

    try:
        with os.scandir(str(source)) as entries:
            names = []
            directory_names = []
            for entry in entries:
                names.append(entry.name)
                if len(names) > MAX_SOURCE_DIRECTORY_ENTRIES:
                    raise _PreparationError("所选解压目录项目过多，无法安全扫描。请直接选择其 DEBS 文件夹。")
                try:
                    entry_stat = entry.stat(follow_symlinks=False)
                except OSError as exc:
                    raise _PreparationError("无法安全读取所选解压目录。请重新选择安装包。", repr(exc))
                if stat.S_ISDIR(entry_stat.st_mode):
                    directory_names.append(entry.name)
    except _PreparationError:
        raise
    except OSError as exc:
        raise _PreparationError("无法读取所选解压目录。请重新选择 LibreOffice 安装包。", repr(exc))
    if len(directory_names) != 1:
        raise _PreparationError("未找到唯一的 LibreOffice/DEBS 子目录。请直接选择 DEBS 文件夹或其单层解压目录。")
    wrapper = source / directory_names[0]
    try:
        wrapper_stat = wrapper.lstat()
        debs_stat = (wrapper / "DEBS").lstat()
    except OSError:
        raise _PreparationError("单层解压目录中没有可用的 DEBS 文件夹。请重新选择 LibreOffice 安装包。")
    if stat.S_ISLNK(wrapper_stat.st_mode) or not stat.S_ISDIR(wrapper_stat.st_mode):
        raise _PreparationError("解压目录包含不安全的子目录。请直接选择 DEBS 文件夹。")
    if stat.S_ISLNK(debs_stat.st_mode) or not stat.S_ISDIR(debs_stat.st_mode):
        raise _PreparationError("单层解压目录中没有普通的 DEBS 文件夹。请重新选择 LibreOffice 安装包。")
    return wrapper / "DEBS"


def _folder_package_entries(debs_directory: Path) -> list[tuple[str, os.stat_result]]:
    try:
        directory_stat = debs_directory.lstat()
    except OSError as exc:
        raise _PreparationError("DEBS 文件夹不存在或无法读取。请重新选择安装包。", repr(exc))
    if stat.S_ISLNK(directory_stat.st_mode) or not stat.S_ISDIR(directory_stat.st_mode):
        raise _PreparationError("DEBS 项不是普通文件夹。请重新选择 LibreOffice DEBS 文件夹。")
    entries = []
    scanned_entries = 0
    try:
        with os.scandir(str(debs_directory)) as scan:
            for entry in scan:
                scanned_entries += 1
                if scanned_entries > MAX_ARCHIVE_MEMBERS:
                    raise _PreparationError("DEBS 文件夹中的项目数量超过安全限制。请直接选择完整的官方安装包。")
                try:
                    entry_stat = entry.stat(follow_symlinks=False)
                except OSError as exc:
                    raise _PreparationError("无法安全检查 DEBS 文件夹中的项目。请重新选择安装包。", repr(exc))
                if stat.S_ISLNK(entry_stat.st_mode):
                    if entry.name.casefold().endswith(".deb"):
                        raise _PreparationError("DEBS 文件夹中的安装包是符号链接，已停止检查。请重新解压官方安装包。", entry.name)
                    continue
                if entry.name.casefold().endswith(".deb"):
                    if not stat.S_ISREG(entry_stat.st_mode):
                        raise _PreparationError("DEBS 文件夹中的 .deb 项不是普通文件。请重新解压官方安装包。", entry.name)
                    entries.append((entry.name, entry_stat))
    except _PreparationError:
        raise
    except OSError as exc:
        raise _PreparationError("无法读取 DEBS 文件夹。请检查文件权限后重试。", repr(exc))
    if not entries:
        raise _PreparationError("所选 DEBS 文件夹中没有 .deb 安装包。请重新选择 LibreOffice 的 DEBS 文件夹。")
    if len(entries) > MAX_DEB_PACKAGES:
        raise _PreparationError("DEBS 文件夹中的软件包数量超过安全限制。请重新选择完整的官方安装包。")
    return entries


def _stage_archive(source: Path, staging: Path, data_dir: Path) -> list[tuple[Path, str, int]]:
    try:
        source_stat = source.lstat()
    except OSError as exc:
        raise _PreparationError("所选压缩包不存在或无法读取。请重新选择 LibreOffice Linux DEB tar.gz 文件。", repr(exc))
    if stat.S_ISLNK(source_stat.st_mode) or not stat.S_ISREG(source_stat.st_mode):
        raise _PreparationError("所选压缩包不是普通文件。请重新选择 LibreOffice Linux DEB tar.gz 文件。")
    if source_stat.st_size > MAX_TOTAL_ARCHIVE_BYTES:
        raise _PreparationError("所选压缩包超过安全大小限制。请重新下载 LibreOffice Linux DEB 压缩包。")
    if not source.name.casefold().endswith((".tar.gz", ".tgz")):
        raise _PreparationError("只支持 LibreOffice Linux DEB 的 .tar.gz 或 .tgz 压缩包。请重新选择对应格式的文件。")

    selected = []
    seen_names = set()
    debs_directories = set()
    total_declared = 0
    total_deb_bytes = 0
    member_count = 0
    try:
        with _open_tar_no_follow(source) as archive:
            for member in archive:
                member_count += 1
                if member_count > MAX_ARCHIVE_MEMBERS:
                    raise _PreparationError("压缩包内文件数量超过安全限制。请重新下载官方 LibreOffice DEB 压缩包。")
                canonical = _normalise_archive_name(member.name, member.isdir())
                if canonical in seen_names:
                    raise _PreparationError("压缩包中含有重复文件路径。请重新下载官方 LibreOffice DEB 压缩包。", canonical)
                seen_names.add(canonical)
                if member.isdir():
                    if member.size != 0:
                        raise _PreparationError("压缩包目录条目大小异常。请重新下载官方 LibreOffice DEB 压缩包。", canonical)
                    continue
                if not member.isfile():
                    raise _PreparationError("压缩包包含链接或特殊文件，已停止检查。请重新下载官方 LibreOffice DEB 压缩包。", canonical)
                if member.size < 0 or member.size > MAX_SINGLE_FILE_BYTES:
                    raise _PreparationError("压缩包内存在超过安全大小限制的文件。请重新下载官方 LibreOffice DEB 压缩包。", canonical)
                total_declared += member.size
                if total_declared > MAX_TOTAL_ARCHIVE_BYTES:
                    raise _PreparationError("压缩包展开后的总大小超过安全限制。请重新下载官方 LibreOffice DEB 压缩包。")
                parts = canonical.split("/") if canonical else []
                if len(parts) >= 2 and parts[-2].casefold() == "debs":
                    if len(parts) not in (2, 3):
                        raise _PreparationError("DEBS 文件夹层级不受支持。请直接选择 LibreOffice 官方安装包。", canonical)
                    debs_path = "/".join(parts[:-1])
                    debs_directories.add(debs_path)
                    if parts[-1].casefold().endswith(".deb"):
                        total_deb_bytes += member.size
                        if total_deb_bytes > MAX_TOTAL_DEB_BYTES:
                            raise _PreparationError("DEBS 安装包总大小超过安全限制。请重新下载官方 LibreOffice DEB 压缩包。")
                        selected.append((member.name, canonical, member.size))
                        if len(selected) > MAX_DEB_PACKAGES:
                            raise _PreparationError("LibreOffice DEB 软件包数量超过安全限制。请重新下载官方安装包。")
    except _PreparationError:
        raise
    except (OSError, tarfile.TarError, EOFError) as exc:
        raise _PreparationError("所选文件不是完整、可读取的 LibreOffice DEB tar.gz 压缩包。请重新下载后重试。", repr(exc))

    if not selected:
        raise _PreparationError("压缩包中没有找到 DEBS 文件夹内的 .deb 软件包。请重新下载 LibreOffice Linux DEB 压缩包。")
    if len(debs_directories) != 1:
        raise _PreparationError("压缩包中存在多个 DEBS 目录或目录结构不完整。请重新下载官方 LibreOffice DEB 压缩包。")
    _assert_free_space(data_dir, total_deb_bytes)

    staged = []
    selected_by_name = {name: (canonical, size) for name, canonical, size in selected}
    copied_names = set()
    try:
        with _open_tar_no_follow(source) as archive:
            for member in archive:
                expected = selected_by_name.get(member.name)
                if expected is None:
                    continue
                canonical, expected_size = expected
                if not member.isfile() or member.size != expected_size or member.name in copied_names:
                    raise _PreparationError("压缩包内容在检查期间发生变化。请重新选择安装包。", canonical)
                stream = archive.extractfile(member)
                if stream is None:
                    raise _PreparationError("压缩包中的 DEB 文件无法读取。请重新下载官方安装包。", canonical)
                destination = staging / "package-{:04d}.deb".format(len(staged))
                with stream:
                    digest, size = _copy_stream(stream, destination, expected_size, MAX_TOTAL_DEB_BYTES)
                staged.append((destination, digest, size))
                copied_names.add(member.name)
    except _PreparationError:
        raise
    except (OSError, tarfile.TarError, EOFError) as exc:
        raise _PreparationError("无法把 DEB 文件安全复制到暂存目录。请重新下载或解压 LibreOffice 安装包。", repr(exc))
    if copied_names != set(selected_by_name):
        raise _PreparationError("压缩包内容在检查期间发生变化。请重新选择安装包。")
    return staged


@contextmanager
def _open_tar_no_follow(source: Path):
    flags = os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0)
    fd = os.open(str(source), flags)
    try:
        if not stat.S_ISREG(os.fstat(fd).st_mode):
            raise _PreparationError("所选压缩包不是普通文件。请重新选择 LibreOffice Linux DEB tar.gz 文件。")
        with os.fdopen(fd, "rb") as stream:
            fd = -1
            with tarfile.open(fileobj=stream, mode="r:gz") as archive:
                yield archive
    finally:
        if fd >= 0:
            try:
                os.close(fd)
            except OSError:
                pass


def _stage_folder(source: Path, staging: Path, data_dir: Path) -> list[tuple[Path, str, int]]:
    debs_directory = _find_debs_directory(source)
    entries = _folder_package_entries(debs_directory)
    total_bytes = sum(entry_stat.st_size for _, entry_stat in entries)
    if total_bytes > MAX_TOTAL_DEB_BYTES or any(size.st_size > MAX_SINGLE_FILE_BYTES for _, size in entries):
        raise _PreparationError("DEBS 文件夹中的安装包总大小超过安全限制。请重新选择完整的 LibreOffice 安装包。")
    _assert_free_space(data_dir, total_bytes)

    directory_fd = None
    if os.name == "posix":
        try:
            directory_fd = os.open(
                str(debs_directory),
                os.O_RDONLY | getattr(os, "O_DIRECTORY", 0) | getattr(os, "O_NOFOLLOW", 0),
            )
            opened_stat = os.fstat(directory_fd)
            if not stat.S_ISDIR(opened_stat.st_mode):
                raise _PreparationError("DEBS 文件夹在检查期间发生变化。请重新选择安装包。")
        except _PreparationError:
            if directory_fd is not None:
                os.close(directory_fd)
            raise
        except OSError as exc:
            raise _PreparationError("DEBS 文件夹无法安全打开。请重新选择安装包。", repr(exc))

    staged = []
    try:
        for index, (filename, expected_stat) in enumerate(entries):
            destination = staging / "package-{:04d}.deb".format(index)
            digest, size = _copy_folder_file(debs_directory, filename, destination, expected_stat, directory_fd)
            staged.append((destination, digest, size))
    finally:
        if directory_fd is not None:
            try:
                os.close(directory_fd)
            except OSError:
                pass
    return staged


def _numeric_version(value: str) -> tuple[int, ...] | None:
    # Debian versions may have an epoch; soffice reports only the upstream part.
    upstream = (value or "").split(":", 1)[-1]
    match = _NUMERIC_VERSION_RE.match(upstream)
    if match is None:
        return None
    try:
        return tuple(int(part) for part in match.group(1).split("."))
    except ValueError:
        return None


def _run_metadata(package_path: Path) -> tuple[str, str, str]:
    command = [str(DPKG_DEB_EXECUTABLE), "--field", str(package_path), "Package", "Version", "Architecture"]
    try:
        completed = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=DPKG_DEB_TIMEOUT_SECONDS,
            check=False,
            env=_system_command_env(),
        )
    except subprocess.TimeoutExpired as exc:
        raise _PreparationError("检查 LibreOffice 软件包信息超时。请重新下载完整安装包后重试。", repr(exc))
    except OSError as exc:
        raise _PreparationError("无法运行 dpkg-deb 检查软件包信息。请使用系统软件管理器安装 LibreOffice。", repr(exc))
    stdout = completed.stdout or ""
    stderr = completed.stderr or ""
    if completed.returncode != 0 or len(stdout) > 4096:
        raise _PreparationError("所选文件不是可识别的 DEB 软件包。请重新选择 LibreOffice 官方 Linux DEB 安装包。", "dpkg-deb exit={} stdout={} stderr={}".format(completed.returncode, stdout[:500], stderr[:500]))
    fields = [line.strip() for line in stdout.splitlines() if line.strip()]
    expected_fields = ("Package", "Version", "Architecture")
    if len(fields) == 3 and any(":" in line for line in fields):
        labelled = {}
        for line in fields:
            key, separator, value = line.partition(":")
            if not separator or key.strip() not in expected_fields or key.strip() in labelled:
                raise _PreparationError("DEB 软件包信息格式无效。请重新选择 LibreOffice 官方 Linux DEB 安装包。", repr(stdout[:1000]))
            labelled[key.strip()] = value.strip()
        if set(labelled) != set(expected_fields):
            raise _PreparationError("DEB 软件包信息不完整。请重新选择 LibreOffice 官方 Linux DEB 安装包。", repr(stdout[:1000]))
        fields = [labelled[name] for name in expected_fields]
    if len(fields) != 3:
        raise _PreparationError("DEB 软件包信息不完整。请重新选择 LibreOffice 官方 Linux DEB 安装包。", repr(stdout[:1000]))
    package_name, version, architecture = fields
    if not _PACKAGE_NAME_RE.fullmatch(package_name):
        raise _PreparationError("DEB 软件包名称无效。请重新选择 LibreOffice 官方 Linux DEB 安装包。", repr(package_name))
    if not _VERSION_RE.fullmatch(version):
        raise _PreparationError("DEB 软件包版本格式无效。请重新选择 LibreOffice 官方 Linux DEB 安装包。", repr(version))
    if not _ARCHITECTURE_RE.fullmatch(architecture):
        raise _PreparationError("DEB 软件包架构格式无效。请重新选择 LibreOffice 官方 Linux DEB 安装包。", repr(architecture))
    return package_name, version, architecture


def _validate_package_set(
    staged_files: list[tuple[Path, str, int]],
    host_architecture: str,
) -> tuple[list[_StagedPackage], str]:
    packages = []
    names = set()
    for path, digest, size in staged_files:
        package_name, version, architecture = _run_metadata(path)
        if package_name in names:
            raise _PreparationError("安装包中有重复的软件包名称，无法安全安装。请重新下载完整的官方 LibreOffice 安装包。", package_name)
        names.add(package_name)
        if architecture not in (host_architecture, "all"):
            raise _PreparationError(
                "安装包架构与本机不匹配：软件包 {} 为 {}，本机为 {}。请下载适用于本机的 Linux DEB 版本。".format(package_name, architecture, host_architecture),
                "package={} host={}".format(architecture, host_architecture),
            )
        packages.append(_StagedPackage(package_name, version, architecture, path, digest, size))
    if not packages:
        raise _PreparationError("没有可安装的 LibreOffice DEB 软件包。请重新选择完整安装包。")

    main_candidates = []
    for package in packages:
        main_name = re.fullmatch(r"libreoffice([0-9]+\.[0-9]+)", package.name)
        numeric = _numeric_version(package.version)
        if main_name and numeric and len(numeric) >= 2:
            name_release = tuple(int(part) for part in main_name.group(1).split("."))
            if numeric[:2] == name_release:
                main_candidates.append(package)
    if len(main_candidates) != 1:
        raise _PreparationError(
            "所选 DEBS 文件夹中没有唯一的 LibreOffice 主安装包。请重新选择完整的官方安装包。",
            "main packages={}".format([item.name for item in main_candidates]),
        )
    main_package = main_candidates[0]
    version = main_package.version
    numeric_version = _numeric_version(version)
    if numeric_version is None or len(numeric_version) < 2:
        raise _PreparationError("无法识别 LibreOffice 发行版版本。请重新下载官方安装包。", version)
    release = "{}.{}".format(numeric_version[0], numeric_version[1])

    # Official upstream DEB bundles ship this one Debian menu integration
    # package with a three-component upstream version (for example 7.6.7-2)
    # while the LibreOffice binaries use four components (7.6.7.2-2). Keep this
    # narrowly scoped to that package and require its first three components to
    # match the primary LibreOffice release exactly.
    menu_package_name = "libreoffice{}-debian-menus".format(release)
    allowed_unversioned = {"ure", "uno-libs-private"}
    for package in packages:
        prefix = "{}{}".format("libreoffice" if package.name.startswith("libreoffice") else "libobasis", release)
        if package.name in allowed_unversioned:
            if package.version != version:
                raise _PreparationError(
                    "DEB 软件包版本不一致。请重新下载同一发行版的完整 LibreOffice 安装包。",
                    "package={} version={} primary={}".format(package.name, package.version, version),
                )
            continue
        if package.name == prefix or package.name.startswith(prefix + "-"):
            if package.version == version:
                continue
            menu_numeric = _numeric_version(package.version)
            if (
                package.name == menu_package_name
                and menu_numeric is not None
                and len(menu_numeric) == 3
                and len(numeric_version) >= 4
                and menu_numeric == numeric_version[:3]
            ):
                continue
            raise _PreparationError(
                "DEB 软件包版本不一致。请重新下载同一发行版的完整 LibreOffice 安装包。",
                "package={} version={} primary={}".format(package.name, package.version, version),
            )
        raise _PreparationError(
            "所选安装包包含非 LibreOffice {} 软件包（{}）。请重新选择只包含该发行版 DEB 文件的官方安装包。".format(release, package.name),
            package.name,
        )

    required = {
        "libreoffice{}".format(release),
        "libreoffice{}-calc".format(release),
        "libobasis{}-core".format(release),
        "libobasis{}-calc".format(release),
    }
    missing = sorted(required - names)
    if missing:
        raise _PreparationError(
            "所选 DEBS 文件夹不完整，缺少 LibreOffice Calc 主程序或核心组件。请重新选择完整的官方 DEBS 安装包。",
            "missing={}".format(", ".join(missing)),
        )
    packages.sort(key=lambda item: (item.name, item.architecture, str(item.path)))
    return packages, version


def _verify_staged_package(package: _StagedPackage) -> bool:
    flags = os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0)
    try:
        fd = os.open(str(package.path), flags)
    except OSError:
        return False
    digest = hashlib.sha256()
    size = 0
    try:
        if not stat.S_ISREG(os.fstat(fd).st_mode):
            return False
        with os.fdopen(fd, "rb") as stream:
            fd = -1
            while True:
                block = stream.read(1024 * 1024)
                if not block:
                    break
                size += len(block)
                digest.update(block)
    finally:
        if fd >= 0:
            try:
                os.close(fd)
            except OSError:
                pass
    return size == package.size and digest.hexdigest() == package.sha256


def _probe_engine_version(expected_version: tuple[int, ...] | None = None) -> tuple[str, str]:
    """Refresh engine discovery and safely query the expected installed target."""
    engine = None
    discovery_error = ""
    try:
        from base_audit.engines.libreoffice import find_calc_engine

        engine = find_calc_engine()
    except Exception as exc:
        discovery_error = "find_calc_engine failed: {}".format(repr(exc))
    command_prefix = None
    display = ""
    if expected_version is not None:
        if len(expected_version) < 2:
            return "", "expected LibreOffice version is invalid"
        target = Path("/opt/libreoffice{}.{}".format(expected_version[0], expected_version[1])) / "program" / "soffice"
        if not _is_executable(target):
            return "", "expected target soffice was not found: {}".format(target)
        command_prefix = (str(target),)
        display = str(target)
    if command_prefix is None:
        if engine is None:
            return "", discovery_error or "find_calc_engine did not find a Calc engine"
        command_prefix = engine.command_prefix
        display = engine.display
    try:
        completed = subprocess.run(
            [*command_prefix, "--version"],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=SOFFICE_VERSION_TIMEOUT_SECONDS,
            check=False,
            env=_system_command_env(),
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return "", "soffice --version failed: {}".format(repr(exc))
    output = completed.stdout or ""
    match = _SOFFICE_VERSION_RE.search(output)
    if completed.returncode != 0 or match is None:
        return "", "soffice --version exited {}: {}".format(completed.returncode, output[:1000])
    return match.group(1), display


class LibreOfficeInstaller:
    """Asynchronous, serialized validator and installer for Linux LO DEB bundles."""

    def __init__(self, data_dir: Path) -> None:
        self.data_dir = Path(data_dir)
        self._lock = threading.RLock()
        self._log_lock = threading.Lock()
        self._phase = "idle"
        self._busy = False
        self._message = "请选择 LibreOffice 官方 Linux DEB tar.gz 文件或已解压的 DEBS 文件夹。"
        self._source = ""
        self._source_name = ""
        self._package_version = ""
        self._architecture = ""
        self._package_count = 0
        self._installed = ""
        self._log_path = ""
        self._staging_dir = None
        self._packages = []
        self._thread = None

    def snapshot(self) -> dict[str, object]:
        """Return the current state without probing or launching system tools."""
        available = _is_linux()
        supported, _ = _tool_support() if available else (False, "此安装方式仅适用于 Linux。")
        with self._lock:
            can_install = bool(
                self._phase == "ready"
                and not self._busy
                and self._packages
                and self._staging_dir is not None
                and supported
            )
            return {
                "available": available,
                "supported": supported,
                "busy": self._busy,
                "phase": self._phase,
                "message": self._message,
                "source": self._source,
                "sourceName": self._source_name,
                "packageVersion": self._package_version,
                "architecture": self._architecture,
                "packageCount": self._package_count,
                "installed": self._installed,
                "logPath": self._log_path,
                "canInstall": can_install,
            }

    def prepare(self, path: str) -> dict[str, object]:
        """Begin asynchronous validation/staging for one selected path."""
        source_text = os.fspath(path) if path is not None else ""
        if not _is_linux():
            with self._lock:
                if not self._busy:
                    self._source = source_text
                    self._source_name = Path(source_text).name if source_text else ""
                    self._phase = "failed"
                    self._message = "此安装方式仅适用于 Linux。请在 Linux 系统上使用 LibreOffice DEB 安装包。"
            return self.snapshot()

        with self._lock:
            if self._busy:
                return self.snapshot()
            old_staging = self._staging_dir
            self._staging_dir = None
            self._packages = []
            self._source = source_text
            self._source_name = Path(source_text).name if source_text else ""
            self._package_version = ""
            self._architecture = ""
            self._package_count = 0
            self._phase = "checking"
            self._busy = True
            self._message = "正在检查 LibreOffice 安装包…"
            worker = threading.Thread(
                target=self._prepare_worker,
                args=(source_text, old_staging),
                name="libreoffice-deb-prepare",
                daemon=True,
            )
            self._thread = worker
            worker.start()
        return self.snapshot()

    def start(self) -> dict[str, object]:
        """Begin installation of the exact staged files that passed validation."""
        with self._lock:
            if self._busy or self._phase != "ready" or not self._packages or self._staging_dir is None:
                return self.snapshot()
            supported, reason = _tool_support()
            if not supported:
                self._phase = "failed"
                self._message = reason
                return self.snapshot()
            packages = list(self._packages)
            staging_dir = self._staging_dir
            self._busy = True
            self._phase = "installing"
            self._message = "正在安装 LibreOffice。请等待系统安装完成。"
            worker = threading.Thread(
                target=self._install_worker,
                args=(packages, staging_dir),
                name="libreoffice-deb-install",
                daemon=True,
            )
            self._thread = worker
            worker.start()
        return self.snapshot()

    def _prepare_worker(self, source_text: str, old_staging) -> None:
        staging = None
        ready = False
        source = Path(source_text).expanduser() if source_text else Path()
        log_path = ""
        try:
            self._remove_staging(old_staging)
            log_path = self._create_log()
            self._set_log_path(log_path)
            self._log("Prepare started for source {!r}".format(source_text))
            supported, reason = _tool_support()
            if not supported:
                raise _PreparationError(reason)
            if not source_text:
                raise _PreparationError("尚未选择 LibreOffice 安装包。请选择官方 Linux DEB tar.gz 文件或 DEBS 文件夹。")
            source_stat = source.lstat()
            if stat.S_ISLNK(source_stat.st_mode):
                raise _PreparationError("所选安装包路径是符号链接。请重新选择原始文件或普通 DEBS 文件夹。")
            self.data_dir.mkdir(parents=True, exist_ok=True)
            staging = Path(tempfile.mkdtemp(prefix=".libreoffice-install-", dir=str(self.data_dir)))
            try:
                os.chmod(str(staging), 0o700)
            except OSError:
                pass
            host_architecture = _architecture_from_dpkg()
            suffix = source.name.casefold()
            if stat.S_ISDIR(source_stat.st_mode):
                staged_files = _stage_folder(source, staging, self.data_dir)
            elif stat.S_ISREG(source_stat.st_mode) and suffix.endswith((".tar.gz", ".tgz")):
                staged_files = _stage_archive(source, staging, self.data_dir)
            elif stat.S_ISREG(source_stat.st_mode):
                raise _PreparationError("所选文件不是 .tar.gz 压缩包。请选择 LibreOffice Linux DEB tar.gz 文件或已解压的 DEBS 文件夹。")
            else:
                raise _PreparationError("所选位置不是普通文件或文件夹。请重新选择 LibreOffice 安装包。")
            packages, package_version = _validate_package_set(staged_files, host_architecture)
            if len(packages) != len(staged_files):
                raise _PreparationError("暂存包数量与检查结果不一致。请重新选择安装包。")
            self._log("Validated {} packages; version={!r}, architecture={!r}".format(len(packages), package_version, host_architecture))
            for package in packages:
                self._log("package={} version={} architecture={} size={} sha256={}".format(package.name, package.version, package.architecture, package.size, package.sha256))
            with self._lock:
                self._staging_dir = staging
                self._packages = packages
                self._package_version = package_version
                self._architecture = host_architecture
                self._package_count = len(packages)
                self._phase = "ready"
                self._busy = False
                self._message = (
                    "已检查 {} 个 LibreOffice Calc 安装包，版本 {}，架构 {}。请仅安装来自可信来源的文件。".format(
                        len(packages), package_version, host_architecture
                    )
                )
                ready = True
        except _PreparationError as exc:
            if log_path:
                self._log("Preparation rejected: {}".format(exc.detail or str(exc)))
            with self._lock:
                self._phase = "failed"
                self._busy = False
                self._message = exc.message
                self._staging_dir = None
                self._packages = []
        except Exception as exc:
            if log_path:
                self._log("Preparation exception:\n{}".format(traceback.format_exc()))
            with self._lock:
                self._phase = "failed"
                self._busy = False
                self._message = "无法检查所选 LibreOffice 安装包。请重新选择文件，并查看技术日志中的详细原因。"
                self._staging_dir = None
                self._packages = []
        finally:
            if staging is not None and not ready:
                self._remove_staging(staging)

    def _install_worker(self, packages: list[_StagedPackage], staging_dir: Path) -> None:
        success = False
        try:
            self._log("Install started for {} packages".format(len(packages)))
            for package in packages:
                if not _verify_staged_package(package):
                    raise _PreparationError("暂存的 LibreOffice 安装包在安装前发生变化。请重新选择来源并重新检查。", str(package.path))
            if _is_root():
                command = [str(DPKG_EXECUTABLE), "--install"] + [str(package.path.resolve()) for package in packages]
                authorized = False
            else:
                command = [str(PKEXEC_EXECUTABLE), "--disable-internal-agent", str(DPKG_EXECUTABLE), "--install"]
                command.extend(str(package.path.resolve()) for package in packages)
                authorized = True
            self._log("Executing argv: {}".format(json.dumps(command, ensure_ascii=False)))
            try:
                with open(self._log_path, "a", encoding="utf-8") as log_file:
                    process = subprocess.Popen(
                        command,
                        stdout=log_file,
                        stderr=subprocess.STDOUT,
                        env=_system_command_env(),
                        shell=False,
                        close_fds=True,
                    )
                    return_code = process.wait()
            except OSError as exc:
                self._log("Unable to start package installation: {}".format(repr(exc)))
                raise _PreparationError("无法启动系统安装程序。请使用系统软件管理器安装 LibreOffice，并查看技术日志。", repr(exc))
            self._log("Install process exited with code {}".format(return_code))
            if authorized and return_code == 126:
                with self._lock:
                    self._phase = "cancelled"
                    self._message = "你取消了 LibreOffice 管理员授权，安装未开始。需要安装时请重新选择安装包并授权。"
                return
            if authorized and return_code == 127:
                raise _PreparationError("无法启动管理员授权或 dpkg。请使用系统软件管理器安装 LibreOffice，并查看技术日志。", "pkexec exited 127")
            if return_code != 0:
                raise _PreparationError(
                    "LibreOffice 软件包安装失败，系统返回错误码 {}。请查看技术日志，或在系统软件管理器中处理依赖后重试。".format(return_code),
                    "dpkg/pkexec exit={}".format(return_code),
                )

            expected_version = _numeric_version(self._package_version)
            actual_version, engine_display = _probe_engine_version(expected_version)
            self._log("Post-install engine version={!r} display={!r}".format(actual_version, engine_display))
            installed_version = _numeric_version(actual_version)
            if not actual_version or expected_version is None or installed_version != expected_version:
                raise _PreparationError(
                    "系统安装程序已结束，但无法确认本次目标 LibreOffice Calc 版本已启用。请查看技术日志，并使用系统软件管理器检查安装结果。",
                    "expected={!r} installed={!r} detail={!r}".format(self._package_version, actual_version, engine_display),
                )
            with self._lock:
                self._installed = "LibreOffice {} ({})".format(actual_version, engine_display)
                self._phase = "success"
                self._message = "LibreOffice Calc {} 已安装并通过版本检查。".format(actual_version)
                success = True
        except _PreparationError as exc:
            self._log("Installation did not succeed: {}".format(exc.detail or str(exc)))
            with self._lock:
                if self._phase != "cancelled":
                    self._phase = "failed"
                    self._message = exc.message
        except Exception:
            self._log("Installation exception:\n{}".format(traceback.format_exc()))
            with self._lock:
                self._phase = "failed"
                self._message = "LibreOffice 安装过程出现错误。请查看技术日志，或使用系统软件管理器检查安装结果。"
        finally:
            self._remove_staging(staging_dir)
            with self._lock:
                self._packages = []
                self._staging_dir = None
                if self._phase in ("success", "failed", "cancelled"):
                    self._busy = False

    def _create_log(self) -> str:
        logs_dir = self.data_dir / "logs"
        logs_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        path = logs_dir / "libreoffice-install_{}.log".format(timestamp)
        with open(path, "x", encoding="utf-8") as log_file:
            try:
                os.chmod(str(path), 0o600)
            except OSError:
                pass
            log_file.write("LibreOffice DEB installer technical log\n")
        return str(path)

    def _set_log_path(self, path: str) -> None:
        with self._lock:
            self._log_path = path

    def _log(self, message: str) -> None:
        with self._lock:
            path = self._log_path
        if not path:
            return
        line = "{} {}\n".format(datetime.now().isoformat(timespec="seconds"), message)
        try:
            with self._log_lock:
                with open(path, "a", encoding="utf-8") as log_file:
                    log_file.write(line)
        except OSError:
            pass

    @staticmethod
    def _remove_staging(path) -> None:
        if path is None:
            return

        def make_writable_and_retry(function, target, _error_info):
            try:
                os.chmod(target, stat.S_IREAD | stat.S_IWRITE | stat.S_IEXEC)
                function(target)
            except OSError:
                pass

        try:
            shutil.rmtree(str(path), onerror=make_writable_and_retry)
        except OSError:
            pass
