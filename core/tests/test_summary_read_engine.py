from __future__ import annotations

from base_audit.engines import (
    SUMMARY_READER_COM,
    SUMMARY_READER_PYTHON,
)
from base_audit.service import AuditService
from base_audit.settings import SettingsStore


def test_summary_reader_defaults_to_pure_python(tmp_path):
    settings = SettingsStore(tmp_path / "用户设置.json").load()

    assert settings.summary_read_engine == SUMMARY_READER_PYTHON
    assert AuditService(summary_read_engine=settings.summary_read_engine).summary_pipeline_kind() == "native"


def test_summary_reader_com_is_independent_of_formula_calculation_engine():
    service = AuditService(
        engine_preference="Microsoft Excel",
        summary_read_engine=SUMMARY_READER_PYTHON,
    )
    assert service.summary_pipeline_kind() == "native"

    service.summary_read_engine = SUMMARY_READER_COM
    assert service.summary_pipeline_kind() == "com"
