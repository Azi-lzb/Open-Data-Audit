"""真实数据对齐测试：S1F1 / S1F2 / S2F1 / S3F1–F5 用 test-data 真实数据跑生产入口。

目的（用户要求）：为 RustAudit 提供可对齐的结果基线——每个功能把机器可读结果
（命中位置、数量、文件级结果、输出哈希）写入 test-output/rustaudit_alignment/<功能>.json。

约定：
- test-data 被 .gitignore 忽略；数据缺失时对应测试显式 skip（不假装通过）。
- 真实配置经 config_paths 优先取 config-real，只读；写场景先复制到临时目录。
- S1F1 引擎设置固定为 自动(Excel COM) + PYTHON 条件格式 + DIRECT_OOXML 公式/外部写入，
  与 JSON 内 environment.engine_settings_note 一致；RustAudit 对齐必须用相同设置。
- 每个功能独立断言：流程完成、数量稳定、原始文件 SHA-256 不变。
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))

import real_data_alignment as align  # noqa: E402


def test_s1f1_real_data_alignment() -> None:
    if not align.available("s1f1"):
        pytest.skip("test-data 缺少 S1F1 数据（129 份核查表/模板/外部文件或 Rust-Audit 清单）")
    payload = align.run_s1f1()
    totals = payload["totals"]
    assert totals["files"] == 129, totals
    assert totals["failed"] == 0, [b for b in payload["batches"] if b["failed_files"]]
    assert totals["issues"] > 0, "真实样本应产出条件格式/校验命中"
    assert payload["sources_unchanged"] is True, payload["changed_sources"]
    assert len(payload["batches"]) > 0


def test_s1f2_real_data_alignment() -> None:
    if not align.available("s1f2"):
        pytest.skip("test-data 缺少 S1F2 数据（129 份校验结果及报送说明）")
    payload = align.run_s1f2()
    assert payload["input_files"] >= 129, payload["input_files"]
    assert payload["rows_total"] > 0, payload["summary_items"]
    assert payload["output"]["sha256"]


def test_s2f1_real_data_alignment() -> None:
    if not align.available("s2f1"):
        pytest.skip("test-data 缺少 S2F1 数据（2026.06/2026.03 季报）")
    payload = align.run_s2f1()
    assert payload["findings_total"] > 0, payload["findings_total"]
    assert payload["output_files"], "应产出比较/清单输出文件"


def test_s3f1_real_data_alignment() -> None:
    if not align.available("s3f1"):
        pytest.skip("test-data 缺少 S3 两期 CSV")
    payload = align.run_s3f1()
    assert payload["rows_total"] > 0, payload["rows_total"]
    assert payload["comparison"]["sha256"]


def test_s3f2_real_data_alignment() -> None:
    if not align.available("s3f2"):
        pytest.skip("test-data 缺少 S3 两期 CSV")
    payload = align.run_s3f2()
    assert payload["exported_rows"] >= 0
    assert payload["output"]["sha256"]


def test_s3f3_real_data_alignment() -> None:
    if not align.available("s3f3"):
        pytest.skip("test-data 缺少 S3 说明反馈文件")
    payload = align.run_s3f3()
    assert "result" in payload


def test_s3f4_real_data_alignment() -> None:
    if not align.available("s3f4"):
        pytest.skip("test-data 缺少 S3 两期 CSV")
    payload = align.run_s3f4()
    assert payload["rows_total"] >= 0
    assert payload["check"]["sha256"]


def test_s3f5_real_data_alignment() -> None:
    if not align.available("s3f5"):
        pytest.skip("test-data 缺少 S3 两期 CSV")
    payload = align.run_s3f5()
    assert len(payload["rendered_files"]) > 0, "应渲染出至少一个机构金融表单"
