"""Optional Rust writer for new S1-F02 summary workbooks."""

from __future__ import annotations

from datetime import date, datetime
import json
import os
import platform
import subprocess
import sys
import time
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile
from functools import lru_cache
from pathlib import Path
from typing import Callable


def writer_binary() -> Path | None:
    override = os.environ.get("AUDITTOOLS_S1_SUMMARY_WRITER", "").strip()
    if override:
        candidate = Path(override).expanduser()
        return candidate if candidate.is_file() else None
    if platform.machine().lower() not in {"amd64", "x86_64"}:
        return None
    if sys.platform == "win32":
        filename, target = "s1-summary-writer.exe", "x86_64-pc-windows-gnu"
    elif sys.platform.startswith("linux"):
        filename, target = "s1-summary-writer", "x86_64-unknown-linux-musl"
    else:
        return None
    core_root = Path(__file__).resolve().parents[3]
    candidates: list[Path] = []
    frozen_root = getattr(sys, "_MEIPASS", None)
    if frozen_root:
        candidates.append(Path(frozen_root) / "native" / filename)
    candidates.extend((
        core_root / "native" / filename,
        core_root / "rust" / "s1_summary_writer" / "target" / target / "release" / filename,
        core_root / "rust" / "s1_summary_writer" / "target" / "release" / filename,
    ))
    return next((candidate for candidate in candidates if candidate.is_file()), None)


def writer_available() -> bool:
    return writer_binary() is not None


def _json_default(value):
    if isinstance(value, datetime):
        return {"$audit_datetime": "datetime", "value": value.isoformat()}
    if isinstance(value, date):
        return {"$audit_datetime": "date", "value": value.isoformat()}
    raise TypeError(f"unsupported summary cell type: {type(value).__name__}")


@lru_cache(maxsize=8)
def _writer_version(binary: str) -> str:
    try:
        result = subprocess.run(
            [binary, "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            check=False, timeout=5,
        )
    except (OSError, subprocess.SubprocessError):
        return "版本未报告"
    if result.returncode:
        return "版本未报告"
    return result.stdout.decode("utf-8", errors="replace").strip() or "版本未报告"


def write_summary_payload(
    payload: dict,
    output_path: Path,
    *,
    mode: str = "AUTO",
    openpyxl_writer: Callable[[], Path],
    on_engine: Callable[[str], None] | None = None,
) -> Path:
    """Write to a sibling temp file; AUTO can safely retry from clean input."""
    mode = str(mode or "AUTO").strip().upper()
    if mode not in {"AUTO", "RUST_XLSXWRITER", "OPENPYXL"}:
        raise ValueError("S1-F02 汇总结果写入方式只能选择 AUTO、RUST_XLSXWRITER 或 OPENPYXL")
    output_path = Path(output_path)
    report = on_engine or (lambda _message: None)
    if mode == "OPENPYXL":
        started = time.perf_counter()
        result = openpyxl_writer()
        report(f"实际写入引擎=OpenPyXL；端到端耗时={time.perf_counter() - started:.3f} 秒")
        return result

    binary = writer_binary()
    if binary is None:
        reason = "当前程序未附带本平台的 Rust 写入器"
        if mode == "RUST_XLSXWRITER":
            raise RuntimeError(f"无法使用 Rust xlsxwriter：{reason}。请切换为自动或 OpenPyXL。")
        started = time.perf_counter()
        result = openpyxl_writer()
        report(
            f"实际写入引擎=OpenPyXL；自动回退原因={reason}；"
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
            payload, ensure_ascii=False, separators=(",", ":"),
            allow_nan=False, default=_json_default,
        ).encode("utf-8")
        completed = subprocess.run(
            [str(binary), "write-tables", str(temporary)],
            input=encoded, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            check=False, timeout=1800,
        )
        if completed.returncode:
            rust_error = completed.stderr.decode("utf-8", errors="replace").strip()
            rust_error = rust_error or f"Rust 写入器退出码 {completed.returncode}"
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
            raise RuntimeError(f"Rust xlsxwriter 写入 S1-F02 汇总结果失败：{rust_error}")
        fallback_started = time.perf_counter()
        result = openpyxl_writer()
        report(
            "实际写入引擎=OpenPyXL；Rust writer ({}) 自动回退；回退原因={}；"
            "Rust 尝试耗时={:.3f} 秒；OpenPyXL 耗时={:.3f} 秒".format(
                _writer_version(str(binary)), rust_error, elapsed,
                time.perf_counter() - fallback_started,
            )
        )
        return result

    os.replace(temporary, output_path)
    report(
        f"实际写入引擎={_writer_version(str(binary))}；"
        f"端到端耗时={time.perf_counter() - started:.3f} 秒"
    )
    return output_path


def _excel_column_name(number: int) -> str:
    result = ""
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _restore_audit_hyperlink_targets(workbook_path: Path, payload: dict) -> None:
    """Restore exact filesystem targets after the Rust writer creates links.

    rust_xlsxwriter normalizes file URLs differently for Windows and POSIX.
    The existing OpenPyXL contract stores the absolute filesystem path in the
    relationship Target and the sheet/cell in the hyperlink location. Keep
    that contract without re-writing workbook cells or formulas.
    """
    hyperlinks = payload.get("hyperlinks") or []
    if not hyperlinks:
        return
    relationship_namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
    spreadsheet_namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    office_relationship_namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    relation_tag = f"{{{relationship_namespace}}}Relationship"
    hyperlink_tag = f"{{{spreadsheet_namespace}}}hyperlink"
    relation_attribute = f"{{{office_relationship_namespace}}}id"

    ref_to_target = {
        f"{_excel_column_name(int(link['column']) + 1)}{int(link['row']) + 1}": str(link["target"])
        for link in hyperlinks
    }
    worksheet_name = "xl/worksheets/sheet1.xml"
    relations_name = "xl/worksheets/_rels/sheet1.xml.rels"
    temporary_zip = workbook_path.with_name(
        f".{workbook_path.name}.{os.getpid()}-{time.time_ns()}.zip"
    )
    updated_relations = False
    try:
        with ZipFile(workbook_path, "r") as source:
            worksheet_root = ET.fromstring(source.read(worksheet_name))
            relation_targets = {}
            for hyperlink in worksheet_root.iter(hyperlink_tag):
                ref = hyperlink.attrib.get("ref", "")
                target = ref_to_target.get(ref)
                relation_id = hyperlink.attrib.get(relation_attribute)
                if target is not None and relation_id:
                    relation_targets[relation_id] = target
            if len(relation_targets) != len(ref_to_target):
                raise ValueError("Rust 写入器生成的审核副本超链接数量与待写链接不一致")

            relations_root = ET.fromstring(source.read(relations_name))
            for relation in relations_root.findall(relation_tag):
                target = relation_targets.get(relation.attrib.get("Id", ""))
                if target is not None:
                    relation.set("Target", target)
                    relation.set("TargetMode", "External")
            if not all(
                any(relation.attrib.get("Target") == target for relation in relations_root.findall(relation_tag))
                for target in relation_targets.values()
            ):
                raise ValueError("无法恢复审核副本超链接的文件目标")

            ET.register_namespace("", relationship_namespace)
            ET.register_namespace("r", office_relationship_namespace)
            relations_data = ET.tostring(
                relations_root, encoding="utf-8", xml_declaration=True,
            )
            with ZipFile(temporary_zip, "w", compression=ZIP_DEFLATED, compresslevel=6) as target_zip:
                for info in source.infolist():
                    data = relations_data if info.filename == relations_name else source.read(info.filename)
                    target_zip.writestr(info, data)
        os.replace(temporary_zip, workbook_path)
        updated_relations = True
    finally:
        if not updated_relations:
            try:
                temporary_zip.unlink(missing_ok=True)
            except OSError:
                pass


def write_audit_summary_payload(
    payload: dict,
    output_path: Path,
    *,
    mode: str = "AUTO",
    openpyxl_writer: Callable[[], Path],
    on_engine: Callable[[str], None] | None = None,
) -> Path:
    """Write the S1-F01 audit result table, including its workbook hyperlinks."""
    mode = str(mode or "AUTO").strip().upper()
    if mode not in {"AUTO", "RUST_XLSXWRITER", "OPENPYXL"}:
        raise ValueError("S1-F01 审核结果写入方式只能选择 AUTO、RUST_XLSXWRITER 或 OPENPYXL")
    output_path = Path(output_path)
    report = on_engine or (lambda _message: None)
    if mode == "OPENPYXL":
        started = time.perf_counter()
        result = openpyxl_writer()
        report(f"实际写入引擎=OpenPyXL；端到端耗时={time.perf_counter() - started:.3f} 秒")
        return result

    binary = writer_binary()
    if binary is None:
        reason = "当前程序未附带本平台的 Rust 写入器"
        if mode == "RUST_XLSXWRITER":
            raise RuntimeError(f"无法使用 Rust xlsxwriter 写入 S1-F01 审核结果：{reason}。请切换为自动或 OpenPyXL。")
        started = time.perf_counter()
        result = openpyxl_writer()
        report(
            f"实际写入引擎=OpenPyXL；自动回退原因={reason}；"
            f"端到端耗时={time.perf_counter() - started:.3f} 秒"
        )
        return result

    temporary = output_path.with_name(
        f".{output_path.stem}.rust-audit-{os.getpid()}-{time.time_ns()}.tmp.xlsx"
    )
    started = time.perf_counter()
    rust_error = ""
    try:
        encoded = json.dumps(
            payload, ensure_ascii=False, separators=(",", ":"),
            allow_nan=False,
        ).encode("utf-8")
        completed = subprocess.run(
            [str(binary), "write-audit-summary", str(temporary)],
            input=encoded, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            check=False, timeout=1800,
        )
        if completed.returncode:
            rust_error = completed.stderr.decode("utf-8", errors="replace").strip()
            rust_error = rust_error or f"Rust 写入器退出码 {completed.returncode}"
        elif not temporary.is_file() or temporary.stat().st_size == 0:
            rust_error = "Rust 写入器没有生成有效的工作簿文件"
        else:
            _restore_audit_hyperlink_targets(temporary, payload)
    except (OSError, subprocess.SubprocessError, TypeError, ValueError) as exc:
        rust_error = f"{type(exc).__name__}: {exc}"
    elapsed = time.perf_counter() - started
    if rust_error:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
        if mode == "RUST_XLSXWRITER":
            raise RuntimeError(f"Rust xlsxwriter 写入 S1-F01 审核结果失败：{rust_error}")
        fallback_started = time.perf_counter()
        result = openpyxl_writer()
        report(
            "实际写入引擎=OpenPyXL；Rust writer ({}) 自动回退；回退原因={}；"
            "Rust 尝试耗时={:.3f} 秒；OpenPyXL 耗时={:.3f} 秒".format(
                _writer_version(str(binary)), rust_error, elapsed,
                time.perf_counter() - fallback_started,
            )
        )
        return result

    os.replace(temporary, output_path)
    report(
        f"实际写入引擎={_writer_version(str(binary))}；"
        f"端到端耗时={time.perf_counter() - started:.3f} 秒"
    )
    return output_path
