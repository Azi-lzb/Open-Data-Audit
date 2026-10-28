"""S3-F01 optional Rust writer for a newly-created comparison workbook."""

from __future__ import annotations

import json
import os
import platform
import subprocess
import sys
import time
from functools import lru_cache
from pathlib import Path
from typing import Callable


def writer_binary() -> Path | None:
    """Return a compatible bundled/development x64 writer, if available."""
    override = os.environ.get("AUDITTOOLS_S3_COMPARISON_WRITER", "").strip()
    if override:
        candidate = Path(override).expanduser()
        return candidate if candidate.is_file() else None

    if platform.machine().lower() not in {"amd64", "x86_64"}:
        return None
    if sys.platform == "win32":
        filename = "s3-comparison-writer.exe"
        target = "x86_64-pc-windows-gnu"
    elif sys.platform.startswith("linux"):
        filename = "s3-comparison-writer"
        target = "x86_64-unknown-linux-musl"
    else:
        return None

    candidates: list[Path] = []
    frozen_root = getattr(sys, "_MEIPASS", None)
    if frozen_root:
        candidates.append(Path(frozen_root) / "native" / filename)

    # core/native is the optional, external sidecar location for directory-based
    # development and future onedir releases.
    core_root = Path(__file__).resolve().parents[3]
    candidates.append(core_root / "native" / filename)
    candidates.append(
        core_root / "rust" / "s3_comparison_writer" / "target" / target
        / "release" / filename
    )
    if sys.platform.startswith("linux"):
        candidates.append(
            core_root / "rust" / "s3_comparison_writer" / "target" / "release" / filename
        )
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


def writer_available() -> bool:
    return writer_binary() is not None


@lru_cache(maxsize=8)
def _writer_version(binary: str) -> str:
    try:
        completed = subprocess.run(
            [binary, "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            check=False, timeout=5,
        )
    except (OSError, subprocess.SubprocessError):
        return "版本未报告"
    if completed.returncode:
        return "版本未报告"
    value = completed.stdout.decode("utf-8", errors="replace").strip()
    return value or "版本未报告"


def write_rust_payload(
    payload: dict,
    output_path: Path,
    *,
    mode: str = "AUTO",
    openpyxl_writer: Callable[[], Path],
    on_engine: Callable[[str], None] | None = None,
) -> Path:
    """Dispatch one S3-F01 new-workbook write and only fallback from a clean temp.

    A failed Rust process writes only to a unique sibling temporary file. AUTO can
    therefore safely retry through OpenPyXL; a final replace/permission error is
    allowed to propagate instead of being disguised as a fallback.
    """
    mode = str(mode or "AUTO").strip().upper()
    if mode not in {"AUTO", "RUST_XLSXWRITER", "OPENPYXL"}:
        raise ValueError("S3-F01 比较结果写入方式只能选择 AUTO、RUST_XLSXWRITER 或 OPENPYXL")
    output_path = Path(output_path)
    report = on_engine or (lambda _message: None)
    if mode == "OPENPYXL":
        started = time.perf_counter()
        result = openpyxl_writer()
        report(f"实际引擎=OpenPyXL；端到端耗时={time.perf_counter() - started:.3f} 秒")
        return result

    binary = writer_binary()
    if binary is None:
        reason = "当前程序未附带本平台的 Rust 写入器"
        if mode == "RUST_XLSXWRITER":
            raise RuntimeError(f"无法使用 Rust xlsxwriter：{reason}。请切换为自动或 OpenPyXL。")
        started = time.perf_counter()
        result = openpyxl_writer()
        report(
            f"实际引擎=OpenPyXL；自动回退原因={reason}；"
            f"端到端耗时={time.perf_counter() - started:.3f} 秒"
        )
        return result

    temporary = output_path.with_name(
        f".{output_path.stem}.rust-{os.getpid()}-{time.time_ns()}.tmp.xlsx"
    )
    started = time.perf_counter()
    rust_error = ""
    try:
        encoded = json.dumps(
            payload, ensure_ascii=False, separators=(",", ":"), allow_nan=False
        ).encode("utf-8")
        completed = subprocess.run(
            [str(binary), "write-comparison", str(temporary)],
            input=encoded,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
            timeout=1800,
        )
        if completed.returncode:
            detail = completed.stderr.decode("utf-8", errors="replace").strip()
            rust_error = detail or f"Rust 写入器退出码 {completed.returncode}"
        elif not temporary.is_file() or temporary.stat().st_size == 0:
            rust_error = "Rust 写入器没有生成有效的工作簿文件"
    except (OSError, subprocess.SubprocessError, TypeError, ValueError) as exc:
        rust_error = f"{type(exc).__name__}: {exc}"

    elapsed = time.perf_counter() - started
    if rust_error:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
        if mode == "RUST_XLSXWRITER":
            raise RuntimeError(f"Rust xlsxwriter 写入 S3-F01 比较结果失败：{rust_error}")
        fallback_started = time.perf_counter()
        result = openpyxl_writer()
        version = _writer_version(str(binary))
        report(
            f"实际引擎=OpenPyXL；Rust writer ({version}) 自动回退；"
            f"回退原因={rust_error}；Rust 尝试耗时={elapsed:.3f} 秒；"
            f"OpenPyXL 耗时={time.perf_counter() - fallback_started:.3f} 秒"
        )
        return result

    # Do not catch an os.replace failure and call it a compatible-backend
    # fallback: it normally signals an output permission/locking problem.
    os.replace(temporary, output_path)
    version = _writer_version(str(binary))
    report(
        f"实际引擎={version}；端到端耗时={time.perf_counter() - started:.3f} 秒"
    )
    return output_path
