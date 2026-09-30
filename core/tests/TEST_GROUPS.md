# S1 / S2 / S3 系统回归测试分组

先区分三种范围：单文件或单用例是**定向测试**；`s1`、`s2`、`s3` 是**系统回归**；不带标记运行 `core/tests` 是**core 全量回归**。`test_rule_action_engine.py` 的 27 项只覆盖 S3 规则动作相关部分，不代表全部 11 个功能。

在仓库根目录执行：

```powershell
python -m pytest -q -rs -m s1 core/tests
python -m pytest -q -rs -m s2 core/tests
python -m pytest -q -rs -m s3 core/tests
python -m pytest -q -rs core/tests
```

分组由 `core/tests/conftest.py` 给现有测试模块自动标记，标记定义在 `core/pytest.ini`。设置、Web API、路径、配置指引、应用身份和版本清单等公共测试同时进入三组；因此三组测试数之和大于全量测试数。新增测试模块必须在 `conftest.py` 归类，否则收集时明确报错。不要把某一组通过写成“全部功能通过”。

## 现成配置的选择顺序

读取项目现成工作簿的测试统一使用 `config_paths.resolve_test_config()`：逐文件优先取仓库根目录 `config-real/`，对应文件不存在时才回退 `config/`。`默认配置/` 快照按同样的相对路径独立选择，不能用正在使用的正式配置替代默认快照。公开配置只保留空表头，回退后不具备正式规则的测试可能失败；失败应按实际配置来源解释，不能当作完整业务验收。

测试只读真实配置；需要删表、改规则或写入的场景先复制到临时目录。`test_config_paths.py` 进入三组，覆盖优先级、逐文件回退、默认快照及文件不变；这些测试使用临时样例文件，不依赖本机配置内容。

## S1：逐笔统计系统，5 个功能

| 功能 | 主要自动测试 | 不能由单测替代的验收 |
|---|---|---|
| S1-F01 汇总核查表校验 | `golden/test_golden_audit_flow.py`、`golden/test_golden_native_flow.py`、`test_conditional_*.py`、`test_preflight_xlsx.py`、`test_result_writer.py`、`test_history.py` | Windows Excel/WPS 条件格式渲染与公式重算；UOS LibreOffice 真机；真实核查表数量及结果对拍 |
| S1-F02 汇总校验结果说明 | `golden/test_golden_summary_flow.py`、`test_region_summary.py` | 真实说明文件与历史审核结果对拍；需要 Office 的路径须真机验证 |
| S1-F03 组合工作表（核查表） | `test_merge_org.py`、`test_external_sheet_writer.py`、`test_formula_region_writer.py`、`test_native_uos.py` | 多机构真实文件组合、公式/样式与源文件保护 |
| S1-F04 组合工作表（按配置） | `workflow/test_workflow_core.py`、`workflow/test_node_flow_config.py`、`test_combine_settings.py`、`test_native_uos.py` | 真实配置 DAG 和组合输出；目前没有独立的完整黄金流程对拍 |
| S1-F05 制作联合模板 | `test_native_template_merge.py`、`test_merge_org.py`、`test_native_uos.py` | Excel/WPS 与 LibreOffice 原生差异、命名区域及真实模板验收 |

## S2：报表采集系统，1 个功能

| 功能 | 主要自动测试 | 不能由单测替代的验收 |
|---|---|---|
| S2-F01 执行报表规则 | `test_period_v2.py`（正式 S2 包：导入、环比、外部核对、校验、导出）、`test_s2_units.py`、`test_s2_config_check.py` | 真实报表与大集中数据重跑，核对三 Sheet 的行数、结论、数值和配置；目前没有专属黄金端到端对拍 |

`test_period_compare.py`、`test_percentage_indicators.py`、`test_percentage_point_compare.py` 也归在 S2 组，但测的是历史 `base_audit.period_compare` 兼容路径，**不能替代**当前 `systems/s2_report_collection` 的测试。

## S3：大集中统计系统，5 个功能

| 功能 | 主要自动测试 | 不能由单测替代的验收 |
|---|---|---|
| S3-F01 两期数据到比较文件 | `test_central_comparison.py`、`test_rule_action_engine.py`、`test_rule_applicability.py`、`test_central_indicator_rules.py`、`test_central_expression.py`、`test_central_five_segment.py`、`test_s3_rollover_dates.py`、`test_s3_comparison_config_check.py` | 真实两期数据与 VBA/黄金结果对拍，核对规则命中、备注、说明及计算过程 |
| S3-F02 比较文件到说明文件 | `test_central_explanation_exporter.py` 中说明文件导出用例 | 真实比较文件到说明文件的完整输出核对 |
| S3-F03 机构说明到说明文件 | `test_central_explanation_exporter.py` 中反馈导入、身份键与冲突报告用例 | 多机构真实反馈文件回写及冲突对拍 |
| S3-F04 本期数值核对 | `test_central_cross_period.py`、`test_s3_config_check_ui.py`、表达式相关测试 | 真实日报/月报批次与规则结果对拍；目前以组件测试为主 |
| S3-F05 比较文件到金融表单 | `test_central_forms.py`、`test_central_renderers.py`、`test_s3_forms_config_check.py`、`test_ooxml_structure.py`、`test_fast_delete_defined_names.py` | 真实金融表单的 OOXML 结构、格式和办公套件打开结果 |

## 逐功能测试用例：供人工检查

下表写的是**实际存在的代表性 pytest 用例**与其断言目标，不是新写的测试需求。路径以 `core/tests/` 为起点；同一格中后续只写 `::...` 时沿用前一个文件，只写 `::test_...` 时还沿用前一个测试类。完整清单用 `python -m pytest --collect-only -q -m s1 core/tests`（或将 `s1` 换成 `s2`、`s3`）查看。一个测试文件可能覆盖多个功能；系统分组是回归入口，不能把每项文件数当作端到端验收数。

### S1-F01 汇总核查表校验

| 核对点 | pytest 用例 |
|---|---|
| 审核流程黄金输出及重复运行 | `golden/test_golden_audit_flow.py::test_audit_flow_golden_contract`、`::test_audit_flow_rerun_repeats_contract` |
| 原生 DAG 输出、日志开关不改变业务结果 | `golden/test_golden_native_flow.py::test_dag_native_flow_golden_contract`、`::test_s1_business_outputs_are_identical_when_run_log_export_changes` |
| 条件格式逐格触发、优先级、相对引用 | `test_conditional_engine.py::test_triggering_cell_reported_with_result_fields`、`::test_higher_priority_wins_over_lower`、`::test_expression_rule_triggers_with_relative_shift` |
| 没有命名区域时仍扫描“应用于”区域 | `test_conditional_format_fallback.py::test_fallback_scans_applies_to_without_named_range` |
| 模板结构差异及结果格式 | `test_preflight_xlsx.py::PreflightXlsxTests::test_structure_details_keep_sheet_name_and_cell_address`、`test_result_writer.py::test_openpyxl_result_writer_preserves_summary_contract` |

### S1-F02 汇总校验结果说明

| 核对点 | pytest 用例 |
|---|---|
| 说明文件黄金输出、无校验区场景 | `golden/test_golden_summary_flow.py::test_summary_flow_golden_contract`、`::test_summary_flow_without_validation_region` |
| 输出名称、历史说明、历史键、重复运行 | `test_region_summary.py::test_explanation_summary_uses_business_facing_output_name`、`::test_history_context_appends_by_builtin_rule`、`::test_history_key_matches_numeric_and_text_values`、`::test_enrich_history_columns_is_idempotent` |

### S1-F03 组合工作表（核查表）

| 核对点 | pytest 用例 |
|---|---|
| 多机构分组、工作表命名与公式保留 | `test_merge_org.py::test_organisation_key_is_first_filename_segment`、`::test_sheet_prefix_is_sanitized_and_deduplicated`、`::test_openpyxl_merge_keeps_all_sheets_and_formulas` |
| UOS/openpyxl 合并路径 | `test_native_uos.py::test_dag_combine_sheets_runs_with_openpyxl_on_linux` |

### S1-F04 组合工作表（按配置）

| 核对点 | pytest 用例 |
|---|---|
| 固定 DAG 配置不能被旧配置任意改写 | `workflow/test_node_flow_config.py::test_legacy_node_workbook_cannot_override_fixed_dag`、`::test_fixed_workflows_validate_for_com_and_native` |
| 组合预设与业务编译器绑定 | `workflow/test_workflow_core.py::test_union_combine_preset_reuses_combine_sheets_module`、`::test_business_compiler_binds_combine_plan_without_changing_module` |
| DAG 顺序、失败策略、输入校验 | `workflow/test_workflow_core.py::test_linear_graph_runs_in_order`、`::test_failure_policy_fail_workflow_stops_run`、`::test_all_default_business_workflows_have_explicit_valid_inputs` |

### S1-F05 制作联合模板

| 核对点 | pytest 用例 |
|---|---|
| 基准模板、命名区域不冲突、外部公式引用本地化 | `test_native_template_merge.py::test_base_template_is_authoritative`、`::test_named_ranges_get_sheet_suffix_and_no_collision`、`::test_cross_workbook_reference_is_localized` |
| 不覆盖原始模板、拒绝不受支持的源文件 | `test_native_template_merge.py::test_output_never_modifies_originals`、`::test_rejects_xls_and_empty_sources` |

### S2-F01 执行报表规则

| 核对点 | pytest 用例 |
|---|---|
| 三 Sheet 路由、真实 0 数值、未知类型与级别冲突 | `test_period_v2.py::test_v2_export_keeps_zero_numeric_and_routes_three_empty_sheets`、`::test_v2_export_rejects_unknown_audit_type`、`::test_v2_export_rejects_severity_conflict` |
| 普通数值环比与百分数百分点差 | `test_period_v2.py::test_v2_normal_amount_still_uses_relative_change`、`::test_v2_percentage_change_rate_is_percentage_point_difference` |
| 外部核对容差、缺失值、六种比较方式 | `test_period_v2.py::test_v2_external_rule_uses_its_own_tolerance_and_rule_identity`、`test_s2_units.py::test_external_missing_value_is_not_treated_as_real_zero`、`::test_external_comparison_methods_use_report_value_as_left_operand_and_test_boundaries` |
| 金额单位统一后比较、保持小数精度 | `test_s2_units.py::test_amount_comparison_converts_both_sides_to_output_unit_and_keeps_tolerance_in_that_unit`、`::test_rule_expression_preserves_decimal_digits_beyond_excel_float_precision` |
| 正式配置检查、环比策略区间与外部规则字段 | `test_s2_config_check.py::test_formal_and_default_s2_configs_pass_read_only_check`、`::test_checker_blocks_gap_and_duplicate_strategy`、`::test_checker_requires_external_rule_header_order` |
| 日志设置不改变业务输出 | `test_period_v2.py::test_s2_business_output_is_independent_of_unrelated_log_preferences` |

### S3-F01 两期数据到比较文件

| 核对点 | pytest 用例 |
|---|---|
| 比较文件表头、行数、值、备注、说明与计算过程黄金对拍 | `test_central_comparison.py::ComparisonGoldenTests::test_headers_match_golden`、`::test_row_count_matches_golden`、`::test_values_match_golden`、`::test_explain_and_process_match_golden` |
| 12MU2 一类的实际变动值判断、金额阈值 | `test_central_indicator_rules.py::test_change_action_requires_actual_change_even_without_threshold`、`::test_change_action_uses_change_amount_not_current_balance` |
| 结转规则与表达式只在结转日生效 | `test_rule_action_engine.py::test_real_config_all_rollover_scene_actions_are_date_gated`、`test_s3_rollover_dates.py::test_real_config_expression_groups_follow_date_gate` |
| 命中说明含规则编号、阈值、备注；空备注不产生尾部分隔符 | `test_rule_action_engine.py::test_action_output_includes_id_action_effective_threshold_and_config_note`、`::test_action_output_omits_empty_config_note_without_trailing_separator` |
| 五段/八段语法、配置动作检查 | `test_central_five_segment.py::TestGoldenParity::test_sbe_parity_row_by_row`、`test_s3_comparison_config_check.py::test_unsupported_enabled_action_reports_sheet_row_rule_and_original_action` |

### S3-F02 比较文件到说明文件

| 核对点 | pytest 用例 |
|---|---|
| 只导出标记为需要说明的行 | `test_central_explanation_exporter.py::ExplanationExporterTests::test_exports_only_rows_marked_for_explanation` |
| 机构精确/通配参照和源地区保留 | `test_central_explanation_exporter.py::ExplanationExporterTests::test_uses_precise_then_wildcard_org_reference_and_keeps_source_region` |

### S3-F03 机构说明到说明文件

| 核对点 | pytest 用例 |
|---|---|
| 按四字段业务身份匹配反馈 | `test_central_explanation_exporter.py::ExplanationExporterTests::test_imports_feedback_by_four_business_identity_columns` |
| 同名指标按代码区分、冲突报告保留来源与原因 | `test_central_explanation_exporter.py::ExplanationExporterTests::test_keeps_same_named_indicators_separate_by_indicator_code`、`::test_writes_conflict_report_with_feedback_source_and_reason` |

### S3-F04 本期数值核对

| 核对点 | pytest 用例 |
|---|---|
| 五段指标引用、频度与前提条件 | `test_central_cross_period.py::TokenTests::test_five_segment_token_joins_uni_code`、`::DataStoreTests::test_frequency_check`、`::RunCrossPeriodTests::test_premise_false_skips` |
| 阈值触发、取反、缺一侧与公式错误 | `test_central_cross_period.py::RunCrossPeriodTests::test_ratio_rule_hits_and_suppresses`、`::test_inverted_rule`、`::test_one_side_missing_message`、`::test_formula_error_still_outputs_and_logs` |
| 3.2 配置必需 Sheet 与语法检查 | `test_s3_config_check_ui.py::test_formal_32_check_reports_selected_workbooks_sheet_and_syntax_scope`、`::test_32_check_reports_a_malformed_enabled_rule_from_a_disposable_copy` |

### S3-F05 比较文件到金融表单

| 核对点 | pytest 用例 |
|---|---|
| 按机构回填及内嵌模板使用 | `test_central_forms.py::RenderFormsTests::test_renders_one_file_per_org_with_values_and_colors`、`::test_embedded_config_is_used_as_template_and_not_emitted` |
| FAST_OOXML 与 openpyxl 的表单、数值、格式、颜色、隐藏行一致 | `test_central_renderers.py::DualModeParityTests::test_sheet_names_and_order_match`、`::test_cell_values_match`、`::test_number_formats_match`、`::test_warning_fills_match`、`::test_hidden_rows_match` |
| 3.3 配置检查及内嵌表单引用 | `test_s3_forms_config_check.py::test_formal_config_passes_and_is_not_modified`、`::test_missing_referenced_embedded_form_sheet_is_reported` |

## 后续执行口径

1. 改动某一功能：先跑对应测试文件，再跑所属系统标记。跨系统公共代码：跑所有受影响系统标记。
2. 准备提交或合入 `main`：跑 `compileall`、相关系统标记和 core 全量测试；前端改动另跑 JavaScript 语法检查。
3. 涉及真实 Office 计算、条件格式或跨平台行为：额外跑相应 Windows/UOS 真机验收。涉及规则迁移：额外做配置可执行性检查与真实数据/VBA 对拍。
4. 报告每组的 passed、skipped、failed、errors，以及未执行的真机/真实数据项目。`0 skipped` 仅表示本次 pytest 未跳过，**不表示**所有平台或业务数据都已验收。
