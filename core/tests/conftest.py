"""统一测试导入路径：无论从哪个目录运行 pytest，都能同时以
``base_audit`` 和 ``src.base_audit`` 两种风格导入共享核心。"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

_SRC = Path(__file__).resolve().parent.parent / "src"
for entry in (str(_SRC), str(_SRC.parent)):
    if entry not in sys.path:
        sys.path.insert(0, entry)


# 系统级回归入口由文件归属决定；共用 UI/配置/路径/版本测试进入三组。
# 新增测试模块必须在这里归类，避免系统测试静默漏收。
_SYSTEM_TEST_FILES = {
    "s1": {
        "test_combine_settings.py",
        "test_conditional_compare.py",
        "test_conditional_engine.py",
        "test_engine_preflight.py",
        "test_system_info.py",
        "test_conditional_evaluators.py",
        "test_conditional_excel_float.py",
        "test_conditional_format.py",
        "test_conditional_format_fallback.py",
        "test_conditional_label_shift.py",
        "test_conditional_real_data.py",
        "test_config_history.py",
        "test_external.py",
        "test_external_sheet_writer.py",
        "test_feature_log.py",
        "test_formula_inspection.py",
        "test_formula_region_writer.py",
        "test_history.py",
        "test_indicator_resolver.py",
        "test_merge_org.py",
        "test_native_template_merge.py",
        "test_native_uos.py",
        "test_preflight_xlsx.py",
        "test_region_summary.py",
        "test_result_writer.py",
        "test_summary_read_engine.py",
        "test_template.py",
        "test_template_snapshot.py",
        "test_golden_audit_flow.py",
        "test_golden_native_flow.py",
        "test_golden_summary_flow.py",
        "test_node_flow_config.py",
        "test_workflow_core.py",
        "test_runner_no_issues_error.py",
    },
    "s2": {
        "test_period_compare.py",
        "test_period_v2.py",
        "test_percentage_indicators.py",
        "test_percentage_point_compare.py",
        "test_s2_config_check.py",
        "test_s2_units.py",
    },
    "s3": {
        "test_central_comparison.py",
        "test_central_config_split.py",
        "test_central_cross_period.py",
        "test_central_csv_importer.py",
        "test_central_explanation_exporter.py",
        "test_central_expression.py",
        "test_central_five_segment.py",
        "test_central_forms.py",
        "test_central_indicator_rules.py",
        "test_central_renderers.py",
        "test_complex_rule_compile.py",
        "test_complex_rule_preflight.py",
        "test_expression_lae.py",
        "test_fast_delete_defined_names.py",
        "test_office_expression.py",
        "test_ooxml_structure.py",
        "test_rule_action_engine.py",
        "test_rule_applicability.py",
        "test_s3_common_config_check.py",
        "test_s3_comparison_config_check.py",
        "test_s3_config_check_ui.py",
        "test_s3_expression_numeric_equality.py",
        "test_s3_forms_config_check.py",
        "test_s3_rollover_dates.py",
    },
}
_SHARED_TEST_FILES = {
    "test_config_guide.py",
    "test_discovery.py",
    "test_path_browser.py",
    "test_version_manifest.py",
    "test_web_app.py",
}


def pytest_collection_modifyitems(items: list[pytest.Item]) -> None:
    for item in items:
        filename = Path(str(item.path)).name
        groups = (
            tuple(_SYSTEM_TEST_FILES)
            if filename in _SHARED_TEST_FILES
            else tuple(group for group, files in _SYSTEM_TEST_FILES.items() if filename in files)
        )
        if not groups:
            raise pytest.UsageError(
                f"测试文件 {filename} 未归入 S1/S2/S3；请更新 core/tests/conftest.py"
            )
        for group in groups:
            item.add_marker(getattr(pytest.mark, group))
