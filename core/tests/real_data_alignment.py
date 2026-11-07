"""真实数据对齐 runner：S1F1 / S1F2 / S2F1 / S3F1–F5 用真实数据跑生产入口，
把机器可读结果写入 test-output/rustaudit_alignment/ 供 RustAudit 对齐。

数据根：仓库 test-data/（被 .gitignore 忽略，真实数据不入库）。
- S1F1: test-data/S1F1_external_fallback_20261106/（129 份核查表 + 模板 + 外部文件）；
  文件→模板映射优先取 Rust-Audit 清单（RUST_AUDIT_DATA 环境变量可覆盖其根目录）。
- S1F2: test-data/S1F2/*/<…>校验结果及报送说明.xlsx（129 份）。
- S2F1: test-data/S2F1/2026.06 季报（本期）与 2026.03 季报（上期）。
- S3:   test-data/S3/2026.09|2026.10 月报二批中小法人查询.csv、2026.10 说明文件.xls。

配置经 config_paths.resolve_test_config 优先取 config-real（只读；写场景先复制到临时目录）。
引擎设置固定记录在 JSON 内（S1F1：自动引擎 + PYTHON 条件格式 + DIRECT_OOXML 公式/外部写入），
RustAudit 对齐时必须使用相同设置。

作为 pytest 辅助模块被 test_real_data_alignment.py 调用；也可直接运行：
    python core/tests/real_data_alignment.py [--feature s1f1|s1f2|s2f1|s3f1|s3f2|s3f3|s3f4|s3f5|all]
"""

from __future__ import annotations

import dataclasses
import hashlib
import json
import os
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "core" / "src"))
sys.path.insert(0, str(ROOT / "core" / "tests"))

ALIGN_DIR = ROOT / "test-output" / "rustaudit_alignment"
DATA = ROOT / "test-data"
S1F1_DIR = DATA / "S1F1_external_fallback_20261106"
S1F1_INPUTS = S1F1_DIR / "核查表"
S1F1_TEMPLATES = S1F1_DIR / "模板文件_命名区域"
S1F1_EXTERNAL = S1F1_DIR / "外部文件-大集中相关数据.xlsx"
S1F2_DIR = DATA / "S1F2"
S2F1_DIR = DATA / "S2F1"
S3_DIR = DATA / "S3"
RUST_AUDIT_DATA = Path(os.environ.get(
    "RUST_AUDIT_DATA",
    r"C:\Users\AZI\Desktop\VibeCode\Rust-Audit\测试数据"))
MANIFEST = RUST_AUDIT_DATA / "S1F1_新旧对比_20261105_231620" / "python-manifest-with-external.json"

RESULTS: dict[str, dict] = {}


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def write_alignment(name: str, payload: dict) -> Path:
    ALIGN_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "feature": name,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "environment": _environment(),
        **payload,
    }
    target = ALIGN_DIR / f"{name}.json"
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str),
                      encoding="utf-8")
    return target


def _environment() -> dict:
    import platform

    import openpyxl
    return {
        "python": sys.version.split()[0],
        "platform": platform.platform(),
        "openpyxl": openpyxl.__version__,
        "engine_settings_note": "S1F1: engine=自动(Excel COM), conditional_format_evaluator=PYTHON, "
                                "formula_region_writer=DIRECT_OOXML, external_sheet_writer=DIRECT_OOXML",
    }


def _count_fill_colors(rows) -> dict:
    counts: dict[str, int] = {}
    for row in rows:
        key = str(getattr(row, "fill_color", None))
        counts[key] = counts.get(key, 0) + 1
    return counts


def _norm_issue(issue) -> dict:
    d = dataclasses.asdict(issue)
    for key in ("issue_id", "period", "batch_id", "audit_time", "first_seen_period",
                "previous_seen_period", "consecutive_count", "audit_file"):
        d.pop(key, None)
    if d.get("source_file"):
        d["source_file_name"] = Path(str(d.pop("source_file")) or "").name
    return d


def _copy_history(tmp: Path) -> Path:
    from config_paths import resolve_test_config

    history = resolve_test_config("1.逐笔统计系统_配置.xlsx")
    target = tmp / "1.逐笔统计系统_配置.xlsx"
    shutil.copy2(history, target)
    return target


def available(name: str) -> bool:
    checks = {
        "s1f1": S1F1_INPUTS.is_dir() and S1F1_TEMPLATES.is_dir() and S1F1_EXTERNAL.is_file()
                and MANIFEST.is_file(),
        "s1f2": S1F2_DIR.is_dir(),
        "s2f1": (S2F1_DIR / "2026.06 季报").is_dir() and (S2F1_DIR / "2026.03 季报").is_dir(),
        "s3f1": (S3_DIR / "2026.10 月报二批中小法人查询.csv").is_file()
                and (S3_DIR / "2026.09 月报二批中小法人查询.csv").is_file(),
        "s3f4": (S3_DIR / "2026.10 月报二批中小法人查询.csv").is_file(),
        "s3f2": (S3_DIR / "2026.10 月报二批中小法人查询.csv").is_file(),
        "s3f3": (S3_DIR / "2026.10 月报二批中小法人查询.csv").is_file(),
        "s3f5": (S3_DIR / "2026.10 月报二批中小法人查询.csv").is_file(),
    }
    return checks.get(name, False)


# ---------------------------------------------------------------------------
# S1F1：129 份核查表全量审核（19 个模板批次）
# ---------------------------------------------------------------------------

def run_s1f1() -> dict:
    if "s1f1" in RESULTS:
        return RESULTS["s1f1"]
    from config_paths import resolve_test_config

    from base_audit.service import AuditService
    from base_audit.workflow.runner import run_dag_flow

    manifest = json.loads(MANIFEST.read_text(encoding="utf-8-sig"))
    batches = manifest["batches"]
    history = _copy_history(WORK_DIR())
    payload: dict = {"data_root": str(S1F1_DIR), "external": str(S1F1_EXTERNAL),
                     "external_sha256": sha(S1F1_EXTERNAL), "batches": []}
    started = time.monotonic()
    totals = {"files": 0, "successful": 0, "failed": 0, "issues": 0}
    source_hashes = {}
    for index, batch in enumerate(batches, 1):
        template = S1F1_TEMPLATES / Path(batch["template_path"]).name
        output = WORK_DIR() / "s1f1" / batch["batch_id"]
        output.mkdir(parents=True, exist_ok=True)
        selected = [S1F1_INPUTS / Path(p).name for p in batch["selected_files"]]
        for p in selected:
            if str(p) not in source_hashes:
                source_hashes[str(p)] = sha(p)
        started_batch = time.monotonic()
        service = AuditService(
            config_path=history,
            engine_preference="自动",
            conditional_format_evaluator="PYTHON",
            formula_region_writer="DIRECT_OOXML",
            external_sheet_writer="DIRECT_OOXML",
        )
        result = run_dag_flow(
            batch["workflow_id"], service=service, template_path=template,
            input_dir=S1F1_INPUTS, output_dir=output, period=batch["period"],
            history_path=history, selected_files=selected,
            external_path=S1F1_EXTERNAL, recursive=False,
            write_flow_logs=True, strict=True,
        )

        issues = [_norm_issue(i) for i in result.current_issues]
        entry = {
            "batch_id": batch["batch_id"], "name": batch["name"],
            "template": template.name, "period": batch["period"],
            "status": "completed",
            "successful_files": result.successful_files,
            "failed_files": result.failed_files,
            "issue_count": len(issues),
            "issues": issues,
            "files": [{"source": Path(f.source_path).name,
                       "audit": Path(f.audit_path).name if f.audit_path else None,
                       "error": f.error or ""} for f in result.files],
            "elapsed_seconds": round(time.monotonic() - started_batch, 1),
        }
        payload["batches"].append(entry)
        totals["files"] += result.successful_files + result.failed_files
        totals["successful"] += result.successful_files
        totals["failed"] += result.failed_files
        totals["issues"] += len(issues)
        print(f"[s1f1 {index}/{len(batches)}] {batch['name'][:20]} "
              f"files={entry['successful_files']}/{entry['successful_files'] + entry['failed_files']} "
              f"issues={entry['issue_count']} {entry['elapsed_seconds']}s", flush=True)
    changed = [p for p, h in source_hashes.items() if sha(Path(p)) != h]
    payload["totals"] = totals
    payload["elapsed_seconds"] = round(time.monotonic() - started, 1)
    payload["sources_unchanged"] = not changed
    payload["changed_sources"] = [Path(p).name for p in changed]
    RESULTS["s1f1"] = payload
    write_alignment("S1F1", payload)
    return payload


# ---------------------------------------------------------------------------
# S1F2：129 份校验结果及报送说明 → 汇总
# ---------------------------------------------------------------------------

def run_s1f2() -> dict:
    if "s1f2" in RESULTS:
        return RESULTS["s1f2"]
    from base_audit.service import AuditService
    from base_audit.workflow.runner import run_dag_flow

    template = S1F1_TEMPLATES / "！汇总校验结果及报送说明.xlsx"
    files = sorted(S1F2_DIR.glob("*/[!汇总]*校验结果及报送说明.xlsx"))
    files = [f for f in files if f.name != "汇总信息.xlsx" and not f.name.startswith("汇总")]
    # source.excel_files 只扫描根目录：把 129 份摊平到临时目录供来源节点识别。
    flat = WORK_DIR() / "s1f2_flat"
    flat.mkdir(parents=True, exist_ok=True)
    flat_files = []
    for f in files:
        target = flat / f.name
        if not target.is_file():
            shutil.copy2(f, target)
        flat_files.append(target)
    files = flat_files
    history = _copy_history(WORK_DIR())
    output = WORK_DIR() / "s1f2"
    output.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    service = AuditService(config_path=history, engine_preference="自动")
    result = run_dag_flow(
        "dag:汇总校验结果说明", service=service, template_path=template,
        input_dir=flat, output_dir=output, period="2026-02",
        history_path=history, selected_files=files, recursive=False,
        write_flow_logs=True, strict=True,
    )
    output_path = Path(result.output_path)
    payload = {
        "input_files": len(files),
        "summary_items": [{"feature": i.feature_name, "rows": i.row_count, "message": i.message}
                          for i in result.items],
        "rows_total": sum(i.row_count for i in result.items),
        "skipped_files": [{"name": path.name, "reason": match.details}
                          for path, match in result.skipped_files],
        "output": {"name": output_path.name, "sha256": sha(output_path)},
        "elapsed_seconds": round(time.monotonic() - started, 1),
    }
    RESULTS["s1f2"] = payload
    write_alignment("S1F2", payload)
    return payload


# ---------------------------------------------------------------------------
# S2F1：报表采集（2026.06 本期 vs 2026.03 上期）
# ---------------------------------------------------------------------------

def run_s2f1() -> dict:
    if "s2f1" in RESULTS:
        return RESULTS["s2f1"]
    from config_paths import resolve_test_config

    from base_audit.systems.s2_report_collection.service import PeriodAuditV2Service

    config = resolve_test_config("2.报表采集系统_配置.xlsx")
    central = resolve_test_config("外部文件-大集中相关数据20261016.xlsx")
    if not central.is_file():
        central = resolve_test_config("外部文件-大集中相关数据.xlsx")
    output = WORK_DIR() / "s2f1"
    output.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    result = PeriodAuditV2Service().run(
        current_dir=S2F1_DIR / "2026.06 季报", previous_dir=S2F1_DIR / "2026.03 季报",
        config_path=config, output_dir=output, central_path=central,
        external_reader="AUTO",
        on_progress=lambda p: None,
    )
    findings = result.findings
    by_status: dict[str, int] = {}
    for f in findings:
        by_status[f.status] = by_status.get(f.status, 0) + 1
    paths = dataclasses.asdict(result.paths) if dataclasses.is_dataclass(result.paths) else {}
    payload = {
        "config": config.name, "central": central.name,
        "findings_total": len(findings),
        "findings_by_status": by_status,
        "findings": [dataclasses.asdict(f) for f in findings],
        "external_warnings": list(result.external_warnings),
        "output_files": {k: {"path": str(v), "sha256": sha(Path(v)) if v and Path(v).is_file() else None}
                         for k, v in paths.items() if isinstance(v, (str, Path))},
        "elapsed_seconds": round(time.monotonic() - started, 1),
    }
    RESULTS["s2f1"] = payload
    write_alignment("S2F1", payload)
    return payload


# ---------------------------------------------------------------------------
# S3F1：两期比较（2026.10 vs 2026.09 CSV）
# ---------------------------------------------------------------------------

def run_s3f1() -> dict:
    if "s3f1" in RESULTS:
        return RESULTS["s3f1"]
    from config_paths import resolve_test_config

    from base_audit.systems.s3_central_statistics.service import run_comparison

    config = resolve_test_config("3.1大集中执行比较_配置.xlsx")
    output = WORK_DIR() / "s3f1" / "比较结果.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    result = run_comparison(
        current_csv=S3_DIR / "2026.10 月报二批中小法人查询.csv",
        previous_csv=S3_DIR / "2026.09 月报二批中小法人查询.csv",
        output_path=output, config_path=config,
    )
    comparison = Path(result.output_path)
    payload = {
        "config": config.name,
        "comparison": {"name": comparison.name, "sha256": sha(comparison)},
        "rows_total": len(result.rows),
        "rows_by_fill_color": _count_fill_colors(result.rows),
        "elapsed_seconds": round(time.monotonic() - started, 1),
    }
    RESULTS["s3f1"] = payload
    write_alignment("S3F1", payload)
    return payload


def _s3f1_comparison_path() -> Path:
    run_s3f1()
    return WORK_DIR() / "s3f1" / "比较结果.xlsx"


# ---------------------------------------------------------------------------
# S3F2：比较文件 → 说明文件
# ---------------------------------------------------------------------------

def run_s3f2() -> dict:
    if "s3f2" in RESULTS:
        return RESULTS["s3f2"]
    from config_paths import resolve_test_config

    from base_audit.systems.s3_central_statistics.explanation_exporter import (
        export_selected_explanations,
    )

    comparison = _s3f1_comparison_path()
    config = resolve_test_config("3.0大集中通用配置.xlsx")
    output = WORK_DIR() / "s3f2" / "说明文件.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    count = export_selected_explanations(comparison, output, config_path=config)
    payload = {
        "comparison": comparison.name, "config": config.name,
        "exported_rows": count,
        "output": {"name": output.name, "sha256": sha(output)},
        "elapsed_seconds": round(time.monotonic() - started, 1),
    }
    RESULTS["s3f2"] = payload
    write_alignment("S3F2", payload)
    return payload


# ---------------------------------------------------------------------------
# S3F3：机构说明反馈 → 说明文件回写
# ---------------------------------------------------------------------------

def run_s3f3() -> dict:
    if "s3f3" in RESULTS:
        return RESULTS["s3f3"]
    import dataclasses

    from base_audit.systems.s3_central_statistics.explanation_exporter import (
        import_explanation_feedback,
    )

    run_s3f2()
    target = WORK_DIR() / "s3f2" / "说明文件.xlsx"
    feedback_root = WORK_DIR() / "s3f3"
    feedback_root.mkdir(parents=True, exist_ok=True)
    feedback_copy = feedback_root / "2026.10 说明文件.xlsx"
    shutil.copy2(S3_DIR / "2026.10 说明文件.xls", feedback_copy)
    started = time.monotonic()
    result = import_explanation_feedback(target, [feedback_copy])
    payload = {
        "feedback": feedback_copy.name,
        "result": dataclasses.asdict(result) if dataclasses.is_dataclass(result) else str(result),
        "elapsed_seconds": round(time.monotonic() - started, 1),
    }
    RESULTS["s3f3"] = payload
    write_alignment("S3F3", payload)
    return payload


# ---------------------------------------------------------------------------
# S3F4：本期数值核对（跨期）
# ---------------------------------------------------------------------------

def run_s3f4() -> dict:
    if "s3f4" in RESULTS:
        return RESULTS["s3f4"]
    from config_paths import resolve_test_config

    from base_audit.systems.s3_central_statistics.service import run_cross_period_check

    config = resolve_test_config("3.2大集中本期数值核对_配置.xlsx")
    output = WORK_DIR() / "s3f4" / "核对结果.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    result = run_cross_period_check(
        current_csv=S3_DIR / "2026.10 月报二批中小法人查询.csv",
        previous_csv=S3_DIR / "2026.09 月报二批中小法人查询.csv",
        output_path=output, config_path=config,
    )
    check = Path(result.output_path)
    payload = {
        "config": config.name,
        "check": {"name": check.name, "sha256": sha(check)},
        "rows_total": len(result.rows),
        "elapsed_seconds": round(time.monotonic() - started, 1),
    }
    RESULTS["s3f4"] = payload
    write_alignment("S3F4", payload)
    return payload


# ---------------------------------------------------------------------------
# S3F5：比较结果 → 金融表单（每机构 xlsx）
# ---------------------------------------------------------------------------

def run_s3f5() -> dict:
    if "s3f5" in RESULTS:
        return RESULTS["s3f5"]
    from config_paths import resolve_test_config

    from base_audit.systems.s3_central_statistics.service import render_financial_forms

    comparison = _s3f1_comparison_path()
    config = resolve_test_config("3.3大集中指标比较拆分_配置.xlsx")
    output = WORK_DIR() / "s3f5"
    output.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    result = render_financial_forms(
        comparison_file=comparison, template_file=None, output_dir=output,
        config_path=config, render_mode="OPENPYXL",
    )
    rendered = sorted(p for p in output.rglob("*.xlsx"))
    payload = {
        "config": config.name,
        "rendered_files": [{"name": p.name, "sha256": sha(p)} for p in rendered],
        "result_fields": {k: str(v) for k, v in vars(result).items()
                          if isinstance(v, (int, float, str, bool))},
        "elapsed_seconds": round(time.monotonic() - started, 1),
    }
    RESULTS["s3f5"] = payload
    write_alignment("S3F5", payload)
    return payload


_WORK: Path | None = None


def WORK_DIR() -> Path:
    global _WORK
    if _WORK is None:
        _WORK = ALIGN_DIR / "_work"
        _WORK.mkdir(parents=True, exist_ok=True)
    return _WORK


RUNNERS = {
    "s1f1": run_s1f1, "s1f2": run_s1f2, "s2f1": run_s2f1,
    "s3f1": run_s3f1, "s3f2": run_s3f2, "s3f3": run_s3f3,
    "s3f4": run_s3f4, "s3f5": run_s3f5,
}


def main(argv: list[str]) -> int:
    args = argv[1:] or ["all"]
    targets = list(RUNNERS) if "all" in args else args
    failures = []
    for name in targets:
        runner = RUNNERS.get(name)
        if runner is None:
            print(f"[skip] 未知功能：{name}")
            continue
        if not available(name):
            print(f"[skip] {name}: test-data 缺少所需数据")
            continue
        try:
            runner()
            print(f"[ok] {name}")
        except Exception as exc:
            failures.append(name)
            print(f"[FAIL] {name}: {type(exc).__name__}: {exc}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
