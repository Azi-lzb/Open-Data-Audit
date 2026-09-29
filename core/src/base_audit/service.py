from __future__ import annotations

import hashlib
import shutil
import time
from dataclasses import replace
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional

from .discovery import source_workbooks
from .excel_com import ExcelSession
from .region_summary import RegionSummaryItem, RegionSummaryResult
from .external import ExternalSheetPlan, make_external_sheet_plan
from .history import HISTORY_AUDIT_SHEET
from .name_config import (
    FeatureMapping,
    USED_RANGE_SUMMARY_FUNCTION,
    FIXED_ROW_SUMMARY_FUNCTION,
    WORKBOOK_TABLE_MERGE_FUNCTION,
    load_feature_mappings,
)
from .combine_settings import get_combine_plan
from .models import (
    PreflightItem,
    PreflightRunResult,
    SourceMatch,
    TemplateDefinition,
)
from .feature_log import FeatureLog
from .merge_org import MergeOrgResult, TemplateMergeResult, run_combine_sheets, run_merge_org, run_template_merge
from .preflight_xlsx import (
    template_structure_values,
    validate_required_sheets_xlsx,
    validate_source_xlsx,
    write_preflight_report_xlsx,
)
from .template import TemplateError


CONFIG_HISTORY_SHEET = HISTORY_AUDIT_SHEET


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _organisation_from_name(path: Path) -> tuple[str, str]:
    parts = [part.strip() for part in path.stem.split("_") if part.strip()]
    if len(parts) >= 2:
        return parts[0], parts[1]
    return path.stem, path.stem


def _available_output_path(folder: Path, source: Path) -> Path:
    candidate = folder / f"{source.stem}_审核版.xlsx"
    if not candidate.exists():
        return candidate
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return folder / f"{source.stem}_审核版_{stamp}.xlsx"


def _output_prefix(order: int, feature_name: str, output_name: str = "") -> str:
    """输出前缀：填写“输出文件名”列时用该值，否则沿用“序号_功能名”。"""
    return output_name or f"{order:02d}_{feature_name}"


def _external_sheet_plan(
    excel: ExcelSession,
    template_workbook: object,
    definition: object,
    external_workbook: object,
) -> ExternalSheetPlan:
    # 模板规则上千条时逐格 COM 取公式是热路径（每格一次跨进程调用，可达数秒）。
    # 按工作表一次性读取 UsedRange.Formula 二维数组，再用单元格坐标取回规则
    # 公式；语义与逐格读取一致（只看启用规则的公式）。
    from .excel_com import _cell_position

    cells_by_sheet: dict[str, set[str]] = {}
    for rule in definition.rules:
        if rule.enabled:
            cells_by_sheet.setdefault(rule.sheet_name, set()).add(rule.formula_cell)
    formulas: list[object] = []
    for sheet_name, cells in sorted(cells_by_sheet.items()):
        sheet = template_workbook.Worksheets(sheet_name)
        used = sheet.UsedRange
        matrix = used.Formula
        top, left = int(used.Row), int(used.Column)
        if not isinstance(matrix, tuple):
            matrix = ((matrix,),)
        width = max(len(row) for row in matrix) if matrix else 0
        for address in sorted(cells):
            row, column = _cell_position(address)
            r, c = row - top, column - left
            if 0 <= r < len(matrix):
                row_values = matrix[r]
                if isinstance(row_values, tuple) and 0 <= c < len(row_values):
                    formulas.append(row_values[c])
                    continue
                if not isinstance(row_values, tuple) and c == 0:
                    formulas.append(row_values)
                    continue
            # 单元格在 UsedRange 之外（理论不应发生）：回退单格读取。
            formulas.append(sheet.Range(address).Formula)
    return make_external_sheet_plan(
        formulas=formulas,
        available_sheets=excel._worksheet_names(external_workbook),
    )
def _source_files(
    input_dir: Path,
    selected_files: list[Path] | None = None,
    *,
    recursive: "bool | int" = False,
    extra_files: list[Path] | None = None,
) -> list[Path]:
    """源目录扫描结果 + 用户手动追加的文件，共同构成待处理清单。"""
    all_files = list(source_workbooks(input_dir, recursive=recursive))
    extra = [
        Path(path)
        for path in (extra_files or [])
        if Path(path).is_file()
        and Path(path).resolve() not in {path.resolve() for path in all_files}
    ]
    all_files.extend(extra)
    all_files.sort(key=lambda path: path.name)
    if selected_files is None:
        return all_files
    allowed = {path.resolve() for path in all_files}
    requested = [Path(path).resolve() for path in selected_files]
    invalid = [path for path in requested if path not in allowed]
    if invalid:
        raise ValueError("选择的待审核文件不在源数据目录中：" + str(invalid[0]))
    return [path for path in all_files if path.resolve() in set(requested)]


class AuditService:
    def __init__(
        self, *, config_path: Path | None = None, engine_preference: str = "自动",
        summary_read_engine: str = "纯 Python",
        conditional_format_evaluator: str = "PYTHON",
        conditional_format_rule_reader: str = "DIRECT_OOXML",
        formula_region_writer: str = "DIRECT_OOXML",
        external_sheet_writer: str = "DIRECT_OOXML",
        combine_sheets_plans: list[dict[str, object]] | None = None,
        active_combine_sheets_plan_id: str = "default",
    ) -> None:
        self.config_path = config_path
        self.engine_preference = engine_preference
        self.summary_read_engine = summary_read_engine
        # 条件格式检测模式：NATIVE=Excel/WPS 真实渲染（Windows 默认）；
        # OOXML=纯规则求值。UOS（native 管线）固定 OOXML，与此设置无关。
        self.conditional_format_evaluator = conditional_format_evaluator
        self.conditional_format_rule_reader = conditional_format_rule_reader
        self.formula_region_writer = formula_region_writer
        self.external_sheet_writer = external_sheet_writer
        self.combine_sheets_plans = list(combine_sheets_plans or [])
        self.active_combine_sheets_plan_id = active_combine_sheets_plan_id

    def summary_pipeline_kind(self) -> str:
        """返回汇总流程的适配器类型，不复用审核主流程的计算引擎。"""
        from .engines import summary_pipeline_kind

        return summary_pipeline_kind(self.summary_read_engine)

    def get_combine_sheets_plan(self, plan_id: str | None = None) -> dict[str, object]:
        """读取设置中心保存的组合方案，不读取 DAG/流程配置文件。"""
        return get_combine_plan(
            self.combine_sheets_plans,
            self.active_combine_sheets_plan_id,
            plan_id,
        )

    def summarize_regions(
        self,
        *,
        template_path: Path,
        input_dir: Path,
        output_dir: Path,
        selected_files: list[Path] | None = None,
        flow_name: str | None = None,
        recursive: bool = True,
        on_step: Optional[Callable[[str], None]] = None,
        named_range_features: tuple[FeatureMapping, ...] = (),
        output_name: str | None = None,
        feature_log: Optional[FeatureLog] = None,
        copies_dir: Path | None = None,
        feature_names: tuple[str, ...] | None = None,
        with_history: bool = False,
        summary_features: tuple[FeatureMapping, ...] = (),
    ) -> "RegionSummaryResult":
        """区域汇总入口（DAG 汇总节点调用）。

        ``feature_names`` 由调用方（DAG 工作流定义）显式给出时不再读取
        旧“执行流程”配置；未提供时保持空（历史无参调用已随旧执行器移除）。
        ``summary_features`` 由 DAG 汇总节点按 config「区域组」解析后传入，
        此时不再从代码内置默认值取区域，保证“检查的目标=实际汇总的区域”。
        ``with_history`` 为 True 时在汇总输出上追加历史说明列；DAG 已把该
        富化拆为独立节点（``summary.history_enrich``），故默认 False。
        """
        from .engines import SUMMARY_READER_LIBREOFFICE

        if self.config_path is None:
            raise ValueError("汇总功能需要逐笔统计系统_历史审核配置.xlsx")
        flow_features = tuple(feature_names) if feature_names is not None else ()
        if not flow_features and not summary_features:
            raise ValueError("汇总流程没有启用的汇总功能")
        if self.summary_pipeline_kind() == "native":
            from .native.summary import merge_workbook_tables, run_region_summaries as native_run_region_summaries

            if summary_features:
                features = list(summary_features)
            elif named_range_features:
                features = list(named_range_features)
            else:
                names = flow_features
                mappings = {item.name: item for item in load_feature_mappings(self.config_path, template_path)}
                features = [mappings[name] for name in names if name in mappings]
            merge_features = [item for item in features if item.feature_type == WORKBOOK_TABLE_MERGE_FUNCTION]
            row_features = [item for item in features if item.feature_type != WORKBOOK_TABLE_MERGE_FUNCTION]
            output_dir = output_dir.resolve()
            if merge_features and not row_features:
                path = merge_workbook_tables(
                    input_dir=input_dir, output_dir=output_dir, recursive=recursive,
                    selected_files=selected_files, on_step=on_step,
                    output_name=output_name or "汇总表合并",
                )
                return RegionSummaryResult(output_path=path, items=[RegionSummaryItem("汇总表合并", 0, "")])
            if not row_features:
                raise ValueError("没有启用“汇总_任意行汇总”或“汇总_固定行汇总”模块")
            output_rows: list[tuple[str, int]] = []
            feature_counts: dict[str, int] = {}

            def record_result(rows: list[tuple[str, int]], counts: dict[str, int]) -> None:
                output_rows.extend(rows)
                feature_counts.update(counts)

            if self.summary_read_engine == SUMMARY_READER_LIBREOFFICE:
                # LibreOffice 只在临时副本上刷新公式缓存，原始报送文件和模板
                # 始终只读留存；随后复用与纯 Python 相同的区域/表头汇总口径。
                from .engines.libreoffice_adapter import LibreOfficeAdapter

                files = _source_files(input_dir, selected_files, recursive=recursive)
                if not files:
                    raise ValueError("源数据目录中没有可汇总的 .xlsx 文件")
                say = on_step or (lambda _text: None)
                with tempfile.TemporaryDirectory(prefix="base-audit-summary-lo-") as folder:
                    staged_root = Path(folder)
                    staged_template = staged_root / "template" / template_path.name
                    staged_template.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(template_path, staged_template)
                    staged_files: list[Path] = []
                    for index, source in enumerate(files, start=1):
                        # 用独立子目录保留原始文件名/机构名，避免同名文件相互覆盖。
                        staged = staged_root / "sources" / f"{index:04d}" / source.name
                        staged.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(source, staged)
                        staged_files.append(staged)
                    calculator = LibreOfficeAdapter()
                    say("正在通过 LibreOffice Calc 刷新汇总读取副本……")
                    calculator.recalculate(staged_template)
                    for staged in staged_files:
                        calculator.recalculate(staged)
                    path = native_run_region_summaries(
                        template_path=staged_template, input_dir=staged_root / "sources", output_dir=output_dir,
                        features=row_features, recursive=True, selected_files=staged_files,
                        on_step=on_step, output_name=output_name or "区域汇总",
                        on_result=record_result,
                        history_config_path=self.config_path if with_history else None,
                    )
            else:
                path = native_run_region_summaries(
                    template_path=template_path, input_dir=input_dir, output_dir=output_dir,
                    features=row_features, recursive=recursive, selected_files=selected_files,
                    on_step=on_step, output_name=output_name or "区域汇总",
                    on_result=record_result,
                    history_config_path=self.config_path if with_history else None,
                )
            if feature_log is not None:
                log_name = row_features[0].name if len(row_features) == 1 else "区域汇总"
                feature_log.add_sheet(
                    log_name, ("输出工作表", "行数"),
                    output_rows or [("（无汇总行）", 0)],
                )
            items = [
                RegionSummaryItem(item.name, feature_counts.get(item.name, 0), "")
                for item in row_features
            ]
            return RegionSummaryResult(output_path=path, items=items)
        from .region_summary import run_region_summaries
        return run_region_summaries(
            template_path=template_path,
            input_dir=input_dir,
            output_dir=output_dir,
            config_path=self.config_path,
            selected_files=selected_files,
            feature_names=flow_features,
            flow_name=flow_name,
            recursive=recursive,
            on_step=on_step,
            named_range_features=named_range_features,
            output_name=output_name,
            feature_log=feature_log,
            copies_dir=copies_dir,
            engine_preference=self.engine_preference,
            history_config_path=self.config_path if with_history else None,
            summary_features=summary_features,
        )

    def enrich_summary_history(
        self,
        *,
        summary_path: Path,
        on_step: Optional[Callable[[str], None]] = None,
    ) -> Path:
        """把历史表的人工说明列按复合键追加到已生成的汇总工作簿（原地）。

        DAG 的“历史说明富化”节点调用：对汇总输出的每张表，按内置追加规则
        （表名 → 去重列）匹配历史工作簿，缺失的说明列追加到右侧。历史表只读。
        """
        from .engines import pipeline_kind

        if self.config_path is None:
            raise ValueError("历史说明富化需要逐笔统计系统_配置.xlsx")
        summary_path = Path(summary_path).resolve()
        if not summary_path.is_file():
            raise FileNotFoundError(f"汇总工作簿不存在：{summary_path}")
        if on_step is not None:
            on_step("正在富化历史说明")
        if pipeline_kind(self.engine_preference) == "native":
            from .native.summary import enrich_history_columns

            result = enrich_history_columns(summary_path, self.config_path, on_step=on_step)
        else:
            from .region_summary import enrich_history_columns

            result = enrich_history_columns(summary_path, self.config_path, on_step=on_step)
        if on_step is not None:
            on_step("完成：历史说明富化")
        return result

    def merge_org_files(
        self,
        *,
        input_dir: Path,
        output_dir: Path,
        period: str = "",
        selected_files: list[Path] | None = None,
        flow_name: str | None = None,
        recursive: bool = True,
        on_step: Optional[Callable[[str], None]] = None,
        feature_log: Optional[FeatureLog] = None,
        output_name: str | None = None,
    ) -> MergeOrgResult:
        """Run the standalone preparation flow that merges each institution's workbooks."""
        return run_merge_org(
            input_dir=input_dir,
            output_dir=output_dir,
            period=period,
            selected_files=selected_files,
            flow_name=flow_name,
            recursive=recursive,
            on_step=on_step,
            feature_log=feature_log,
            output_name=output_name,
            engine_preference=self.engine_preference,
        )

    def combine_sheets(
        self,
        *,
        input_dir: Path,
        output_dir: Path,
        period: str = "",
        selected_files: list[Path] | None = None,
        flow_name: str | None = None,
        recursive: bool = True,
        on_step: Optional[Callable[[str], None]] = None,
        feature_log: Optional[FeatureLog] = None,
        output_name: str | None = None,
    ) -> MergeOrgResult:
        plan = self.get_combine_sheets_plan()
        if on_step is not None:
            on_step(f"组合分组方案：{plan['name']}（{plan['mode']}）")
        return run_combine_sheets(
            input_dir=input_dir, output_dir=output_dir, period=period,
            selected_files=selected_files, flow_name=flow_name, recursive=recursive,
            on_step=on_step, feature_log=feature_log, output_name=output_name,
            engine_preference=self.engine_preference, grouping_plan=plan,
        )

    def merge_template_files(
        self,
        *,
        base_template: Path,
        source_templates: list[Path],
        on_step: Optional[Callable[[str], None]] = None,
    ) -> TemplateMergeResult:
        """Create a combined template from explicitly selected workbooks.

        This is intentionally outside config-driven audit flows: it is a
        manual, low-frequency template authoring action and never examines the
        workbench's source-data or auto-matched template fields.
        """
        from .engines import pipeline_kind

        if pipeline_kind(self.engine_preference) == "native":
            # UOS/麒麟：openpyxl 实现，语义与 Windows 版一致（基准模板权威、
            # 命名区域改写为工作表后缀、跨簿引用重定向、检查报告同结构）。
            from .native.template_merge import run_template_merge_native

            return run_template_merge_native(
                base_template=base_template,
                source_templates=source_templates,
                on_step=on_step,
            )

        return run_template_merge(
            base_template=base_template,
            source_templates=source_templates,
            on_step=on_step,
            engine_preference=self.engine_preference,
        )

    def _history_storage(self, legacy_path: Path) -> tuple[Path, str, Path | None]:
        """Keep operational history in the configured history workbook.

        The supplied path is retained only as a one-time migration source, so
        existing users do not lose their old independent history workbook.
        """
        if self.config_path is not None:
            # 审核链历史固定为“核查表校验结果”表，按规则编号匹配人工说明。
            return self.config_path.resolve(), CONFIG_HISTORY_SHEET, legacy_path.resolve()
        return legacy_path.resolve(), "问题历史", None

    def preflight(
        self,
        *,
        template_path: Path,
        input_dir: Path,
        output_dir: Path,
        selected_files: list[Path] | None = None,
        external_path: Path | None = None,
        extra_files: list[Path] | None = None,
        recursive: bool = False,
        summary_feature_names: tuple[str, ...] | None = None,
        write_report: bool = True,
        on_step: Optional[Callable[[str], None]] = None,
        named_range_features: tuple[tuple[int, FeatureMapping], ...] = (),
        feature_log: Optional[FeatureLog] = None,
        feature_mappings: list[FeatureMapping] | None = None,
    ) -> PreflightRunResult:
        from .engines import pipeline_kind

        if pipeline_kind(self.engine_preference) == "native":
            from .native.preflight import run_native_preflight

            return run_native_preflight(
                template_path=template_path,
                input_dir=input_dir,
                output_dir=output_dir,
                config_path=self.config_path,
                selected_files=selected_files,
                external_path=external_path,
                recursive=recursive,
                summary_feature_names=summary_feature_names,
                write_report=write_report,
                on_step=on_step,
                named_range_features=named_range_features,
                feature_log=feature_log,
                feature_mappings=feature_mappings,
            )
        template_path = template_path.resolve()
        input_dir = input_dir.resolve()
        output_dir = output_dir.resolve()
        if external_path is not None:
            external_path = external_path.resolve()
        if not template_path.is_file():
            raise FileNotFoundError(f"模板不存在：{template_path}")
        if not input_dir.is_dir():
            raise FileNotFoundError(f"源数据目录不存在：{input_dir}")
        if external_path is not None and not external_path.is_file():
            raise FileNotFoundError(f"外部文件不存在：{external_path}")
        source_files = _source_files(
            input_dir, selected_files, recursive=recursive, extra_files=extra_files
        )
        if not source_files:
            raise ValueError("源数据目录中没有可检查的 .xlsx 文件")
        output_dir.mkdir(parents=True, exist_ok=True)
        batch_id = datetime.now().strftime("%Y%m%d%H%M%S")
        report_path = output_dir / f"审核前检查_{batch_id}.xlsx"

        with ExcelSession(self.engine_preference) as excel:
            template_workbook = excel.open_workbook(template_path, read_only=True)
            external_workbook = None
            external_plan = None
            try:
                try:
                    template_items: list[PreflightItem] = []
                    for _order, mapping in named_range_features:
                        if on_step is not None:
                            on_step(f"正在执行：{mapping.name}")
                        # 先调查再强校验：区域缺失时运行日志仍保留“缺失”记录。
                        if feature_log is not None:
                            survey = excel.survey_named_ranges(template_workbook, mapping)
                            feature_log.add_sheet(
                                mapping.name, ("工作表", "命名区域名", "覆盖区域", "结果"), survey
                            )
                            if on_step is not None:
                                found = [row for row in survey if row[3] == "通过"]
                                if found:
                                    per_sheet: dict[str, int] = {}
                                    for row in found:
                                        per_sheet[row[0]] = per_sheet.get(row[0], 0) + 1
                                    detail = "、".join(f"{sheet} {count} 处" for sheet, count in per_sheet.items())
                                    on_step(f"{mapping.name}：{detail}")
                                else:
                                    on_step(f"{mapping.name}：未找到 {mapping.range_names}")
                        ranges = excel.require_named_ranges(template_workbook, mapping)
                        template_items.append(
                            PreflightItem(
                                "命名区域检查", "提示", "通过", "", "",
                                f"模块“{mapping.name}”已找到 {len(ranges)} 个区域",
                            )
                        )
                        if on_step is not None:
                            on_step(f"完成：{mapping.name}")
                    if not summary_feature_names:
                        # 纯表结构检查流程没有汇总步骤（summary_feature_names 为空），
                        # 同样走模板体检，不做“汇总结构”准备。
                        definition = excel.read_template(
                            template_workbook, config_path=self.config_path
                        )
                        template_items.extend(excel.inspect_template_health(template_workbook, definition))
                    else:
                        mappings = feature_mappings or load_feature_mappings(self.config_path, template_path)
                        selected = [item for item in mappings if item.name in summary_feature_names]
                        data_ranges = []
                        for mapping in selected:
                            data_ranges.extend(excel._named_ranges_for_features(template_workbook, [mapping]))
                        headers = excel._named_ranges_for_features(
                            template_workbook,
                            [FeatureMapping("汇总表头", USED_RANGE_SUMMARY_FUNCTION, ("表头区域",), False, "", "")],
                        )
                        if not data_ranges:
                            raise TemplateError("汇总功能未找到命名区域")
                        if not headers:
                            raise TemplateError("汇总功能缺少“表头区域”命名区域")
                        definition = TemplateDefinition(
                            rules=[], copy_ranges=data_ranges, structured=False,
                            structure_ranges=headers,
                        )
                        template_items.append(
                            PreflightItem(
                                "汇总结构", "提示", "通过", "", "",
                                f"将按 {len(headers)} 个表头区域核对汇总文件结构",
                            )
                        )
                    if external_path is not None:
                        external_workbook = excel.open_workbook(external_path, read_only=True)
                        external_plan = _external_sheet_plan(
                            excel, template_workbook, definition, external_workbook,
                        )
                        template_items.append(
                            excel.inspect_external_workbook(
                                external_workbook, external_plan.sheet_names,
                                plan_source=external_plan.source,
                            )
                        )
                except Exception as exc:
                    template_items = [
                        PreflightItem("总体", "错误", "不通过", "", "", str(exc))
                    ]
                    if feature_log is not None:
                        feature_log.add_template_health(template_path, template_items)
                    if write_report:
                        write_preflight_report_xlsx(
                            report_path, template_path, template_items
                        )
                    raise TemplateError(
                        f"模板体检不通过：{exc}"
                        + (f"。体检报告：{report_path}" if write_report else "")
                    ) from exc

                if feature_log is not None:
                    feature_log.add_template_health(template_path, template_items)
                if any(item.level == "错误" for item in template_items):
                    if write_report:
                        write_preflight_report_xlsx(
                            report_path, template_path, template_items
                        )
                    raise TemplateError(
                        "模板体检不通过，请先处理错误项"
                        + (f"。体检报告：{report_path}" if write_report else "")
                    )

                source_matches: list[tuple[Path, SourceMatch]] = []
                expected_structure = template_structure_values(template_workbook, definition)
                if on_step is not None:
                    on_step("正在执行：表结构比对")
                for source_path in source_files:
                    try:
                        match_result = validate_source_xlsx(
                            source_path, definition, expected_structure,
                            external_sheet_names=external_plan.sheet_names if external_plan else (),
                        )
                    except Exception as exc:
                        match_result = SourceMatch(
                            False, 0.0, 0, 0, (), f"无法完成结构匹配：{exc}"
                        )
                    source_matches.append((source_path, match_result))
                if on_step is not None:
                    on_step("完成：表结构比对")
                if feature_log is not None:
                    feature_log.add_sheet(
                        "表结构比对",
                        ("报送文件", "匹配结果", "匹配率", "匹配标签数", "检查标签数", "缺少工作表", "说明"),
                        [
                            (
                                str(source_path),
                                "通过" if result.matched else "不通过",
                                f"{result.score:.0%}",
                                result.matched_labels,
                                result.checked_labels,
                                "、".join(result.missing_sheets) or "无",
                                result.details,
                            )
                            for source_path, result in source_matches
                        ],
                    )
                if write_report:
                    write_preflight_report_xlsx(
                        report_path, template_path, template_items
                    )
            finally:
                if external_workbook is not None:
                    excel.close_workbook(external_workbook)
                excel.close_workbook(template_workbook)

        result = PreflightRunResult(
            report_path=report_path if write_report else None,
            total_files=len(source_matches),
            matched_files=sum(1 for _, item in source_matches if item.matched),
            template_warnings=sum(
                1 for item in template_items if item.level in {"提示", "警告"}
            ),
            source_matches=tuple(source_matches),
            batch_id=batch_id,
        )
        if on_step is not None:
            on_step("完成：表结构比对")
        return result
