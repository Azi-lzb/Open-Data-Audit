from __future__ import annotations

import json
import os
import re
import ctypes
from dataclasses import asdict, dataclass, field
from pathlib import Path

from .combine_settings import (
    default_combine_plans,
    load_legacy_combine_settings,
    normalize_combine_settings,
)


MAX_RECENT_PATHS = 10
CALCULATION_ENGINES = ("自动", "Microsoft Excel", "WPS 表格", "LibreOffice Calc")
# XLSX 渲染模式：FAST_OOXML=OOXML/XML 高速渲染（默认，固定模板批量生成的
# 实测基准见 docs/大集中统计系统/FAST_OOXML高速渲染器说明.md）；
# OPENPYXL=标准 openpyxl 渲染（兼容性兜底）。
XLSX_RENDER_MODES = ("OPENPYXL", "FAST_OOXML")
DEFAULT_XLSX_RENDER_MODE = "FAST_OOXML"

# 大集中执行比较规则引擎固定为 V3：「规则动作」表 + 启动过滤 + 指标索引。
CENTRAL_RULE_ENGINES = ("v3",)
# V3 已转正：规则引擎固定为 v3。
DEFAULT_CENTRAL_RULE_ENGINE = "v3"

# 表达式 2×2：处理方式（SBE/LAE）× 求值后端（PYTHON/OFFICE）。
EXPRESSION_EVALUATION_MODES = ("SBE", "LAE")
EXPRESSION_EVALUATION_BACKENDS = ("PYTHON", "OFFICE")
DEFAULT_EXPRESSION_EVALUATION_MODE = "SBE"
DEFAULT_EXPRESSION_EVALUATION_BACKEND = "PYTHON"
SUMMARY_READ_ENGINES = ("纯 Python", "Excel/WPS", "LibreOffice Calc")
# 设置中心「背景颜色切换」：浅色（默认）/ 深色 / 纯黑，纯显示偏好。
UI_THEMES = ("浅色", "深色", "纯黑")
FILE_PICKER_MODES = ("系统原生", "Tk 对话框", "浏览器内置")
# 条件格式检测模式：NATIVE=Excel/WPS COM 真实渲染（Windows 默认，保持既有
# 稳定行为）；OOXML=纯规则求值。UOS 固定 OOXML，与该设置无关。
# 条件格式判定方式（三选一，决定“规则是否触发”）：PYTHON=OOXML 规则求值；
# COM_EVALUATE=Excel/WPS 公式引擎求值（规则仍来自 RuleReader）；
# COM_DISPLAY=Excel/WPS DisplayFormat 真实渲染（渲染基线）。
# 旧值 OOXML/NATIVE/DISPLAY_FORMAT_NATIVE 在载入时映射，保持已保存设置可用。
CONDITIONAL_FORMAT_EVALUATORS = ("PYTHON", "COM_EVALUATE", "COM_DISPLAY")
CONDITIONAL_FORMAT_ENGINE_MODES = ("OOXML", "NATIVE")   # 兼容旧设置文件
# 条件格式规则读取方式（与判定方式正交）：DIRECT_OOXML=直读 XML（性能优先）；
# OPENPYXL=对象模型读取（兼容/对照）。
CONDITIONAL_FORMAT_RULE_READERS = ("DIRECT_OOXML", "OPENPYXL")
# 公式校验复制方式（写入后端；计算策略不变）：DIRECT_OOXML=ZIP 级高速写入
# （默认，真实业务基准验证通过）；COM_RANGE=Range.Copy+Formula 矩阵（兼容/
# 结果对照）。UI 显示“Direct OOXML（推荐）/ COM Range（兼容）”。
FORMULA_REGION_WRITERS = ("COM_RANGE", "DIRECT_OOXML")
# 外部辅助表写入方式：DIRECT_OOXML 仅面向纯数值/文本工作表；当前正式外部表
# 已完成基准验证，故默认使用它；复杂工作表仍可切回 openpyxl。
EXTERNAL_SHEET_WRITERS = ("OPENPYXL", "DIRECT_OOXML")
_LEGACY_EVALUATOR_MAP = {
    "OOXML": "PYTHON", "NATIVE": "COM_DISPLAY", "DISPLAY_FORMAT_NATIVE": "COM_DISPLAY",
}


def normalize_conditional_evaluator(value: str) -> str:
    text = str(value or "").strip().upper()
    if text in CONDITIONAL_FORMAT_EVALUATORS:
        return text
    if text in _LEGACY_EVALUATOR_MAP:
        return _LEGACY_EVALUATOR_MAP[text]
    return "PYTHON"


def normalize_conditional_rule_reader(value: str) -> str:
    text = str(value or "").strip().upper()
    return text if text in CONDITIONAL_FORMAT_RULE_READERS else "DIRECT_OOXML"


# 兼容别名（第一阶段名称）
CONDITIONAL_FORMAT_EVALUATOR_MODES = CONDITIONAL_FORMAT_EVALUATORS


def normalize_conditional_evaluator_mode(value: str) -> str:
    return normalize_conditional_evaluator(value)


def hide_application_data_directory(path: Path) -> None:
    """Hide the application-managed JSON directory on Windows when possible.

    This is presentation-only: failure to set the attribute must never prevent
    the audit tool from reading or saving user settings.
    """
    if os.name != "nt" or not path.is_dir():
        return
    try:
        kernel32 = ctypes.windll.kernel32
        attributes = kernel32.GetFileAttributesW(str(path))
        if attributes == 0xFFFFFFFF:
            return
        hidden = 0x02
        if not attributes & hidden:
            kernel32.SetFileAttributesW(str(path), attributes | hidden)
    except (AttributeError, OSError):
        pass


@dataclass(frozen=True)
class FavoritePath:
    name: str
    path: str


@dataclass
class UserSettings:
    last_input_dir: str = ""
    last_template_dir: str = ""
    last_external_file: str = ""
    last_output_dir: str = ""
    history_config: str = ""
    output_pinned: bool = False
    recursive_folders: bool = True
    # 递归深度：-1=最深处（默认）；0=仅根目录；N=最多进入 N 层子目录。
    recursive_depth: int = -1
    # 手动追加进待处理清单的文件（在源数据目录之外），跨会话保留。
    extra_files: list[str] = field(default_factory=list)
    # 页面运行明细、运行日志工作簿和已有性能信息分别控制；
    # write_flow_logs 作为旧设置/旧 API 的兼容别名，规范化为 export_run_logs。
    show_run_detail_logs: bool = False
    export_run_logs: bool = False
    show_performance_diagnostics: bool = False
    write_flow_logs: bool = False
    # 执行主流程前弹出待处理文件清单让用户确认。
    confirm_before_run: bool = True
    # 设置中心「背景颜色切换」：界面配色主题，纯显示偏好，不影响审核行为。
    ui_theme: str = "浅色"
    # 文件/目录选择交互：由外壳传入各自的首次默认值，用户修改后持久保存。
    file_picker_mode: str = "系统原生"
    # 设置中心「条件格式判定方式」（高级设置）：PYTHON=OOXML/Python 规则求值
    # （默认，真实数据与 COM 三路对照一致）；COM_EVALUATE / COM_DISPLAY 为
    # Windows 兼容/渲染基线。仅 Windows 生效，UOS 固定 PYTHON。
    conditional_format_evaluator: str = "PYTHON"
    conditional_format_rule_reader: str = "DIRECT_OOXML"
    # 公式校验复制方式（写入后端）：DIRECT_OOXML 默认（真实业务基准验证）。
    formula_region_writer: str = "DIRECT_OOXML"
    # 外部辅助表写入：纯数据外部表默认使用 OOXML 高速路径。
    external_sheet_writer: str = "DIRECT_OOXML"
    calculation_engine: str = "自动"
    # 只影响“汇总校验结果说明”；默认不启动 Excel/WPS/LibreOffice。
    summary_read_engine: str = "纯 Python"
    favorite_input_dirs: list[FavoritePath] = field(default_factory=list)
    favorite_template_dirs: list[FavoritePath] = field(default_factory=list)
    favorite_external_files: list[FavoritePath] = field(default_factory=list)
    favorite_output_dirs: list[FavoritePath] = field(default_factory=list)
    recent_input_dirs: list[str] = field(default_factory=list)
    recent_template_dirs: list[str] = field(default_factory=list)
    recent_external_files: list[str] = field(default_factory=list)
    recent_output_dirs: list[str] = field(default_factory=list)
    # 报表采集系统（跨期比较）独立保存的最近路径，不与逐笔统计系统混用。
    period_current_dir: str = ""
    period_previous_dir: str = ""
    period_central_file: str = ""
    period_output_dir: str = ""
    period_output_auto: bool = True
    # 报表采集系统：报表采集系统_比较配置.xlsx 的绑定路径（空 = 使用程序目录内置配置）。
    period_config_file: str = ""
    # 大集中核对差异容差（元，默认 100 元 = 0.01 万元），设置中心可调。
    central_diff_tolerance_yuan: float = 100.0
    # 大集中统计系统：独立保存的路径，不与另两个系统混用。
    central_current_csv: str = ""
    central_previous_csv: str = ""
    central_output_dir: str = ""
    central_output_auto: bool = True
    central_cross_current_csv: str = ""
    central_cross_output_dir: str = ""
    central_cross_output_auto: bool = True
    central_compare_file: str = ""
    central_explanation_file: str = ""
    central_template_file: str = ""
    central_form_output_dir: str = ""
    central_form_output_auto: bool = True
    central_common_config: str = ""
    central_comparison_config: str = ""
    central_cross_config: str = ""
    central_forms_config: str = ""
    # 大集中转表等批量模板生成的 XLSX 渲染模式（见 XLSX_RENDER_MODES）。
    # 默认 FAST_OOXML：真实模板实测 8 机构 2.8 秒（并行），已通过 Excel 打开
    # 与内容对照验收；OPENPYXL 仅作为兼容性兜底保留。
    xlsx_render_mode: str = DEFAULT_XLSX_RENDER_MODE
    # 大集中执行比较规则引擎（见 CENTRAL_RULE_ENGINES）。
    central_rule_engine: str = DEFAULT_CENTRAL_RULE_ENGINE
    # 表达式 2×2（见 EXPRESSION_EVALUATION_MODES / BACKENDS）；OFFICE 需
    # 本机可用 Excel/WPS，UOS 选择时执行期明确报错。
    expression_evaluation_mode: str = DEFAULT_EXPRESSION_EVALUATION_MODE
    expression_evaluation_backend: str = DEFAULT_EXPRESSION_EVALUATION_BACKEND
    # 3.1 表达式规则语法（FIVE_SEGMENT_V1=读取「表达式校验」5段表，主推；
    # LEGACY_8=读取「表达式校验（兼容）」8段表）。缺字段按 FIVE_SEGMENT_V1。
    central_expression_schema: str = "FIVE_SEGMENT_V1"
    # 组合工作表是逐笔统计系统的用户业务设置，不属于可编辑 DAG 图。
    combine_sheets_plans: list[dict[str, object]] = field(default_factory=default_combine_plans)
    active_combine_sheets_plan_id: str = "default"


def _is_native_path(value: object) -> bool:
    r"""判断已保存路径是否属于当前平台（空值视为有效，保持不变）。

    Windows 保存的 ``C:\...`` 或 ``\\server\share`` 在 Linux 上无效；
    Linux 保存的 ``/home/...`` 在 Windows 上无效；相对路径一律视为无效。
    """
    text = str(value or "").strip()
    if not text:
        return True
    if os.name == "nt":
        is_drive = bool(re.match(r"^[A-Za-z]:[\\/]", text))
        is_unc = text.startswith("\\\\")
        return is_drive or is_unc
    return text.startswith("/")


class SettingsStore:
    def __init__(self, path: Path) -> None:
        self.path = path
        # load() 若清理了失效路径则置 True，调用方可在启动时回写一次。
        self.sanitized = False

    def load(self, *, file_picker_default: str = "系统原生") -> UserSettings:
        if file_picker_default not in FILE_PICKER_MODES:
            file_picker_default = "系统原生"
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except FileNotFoundError:
            # 首次启动也要有机会从退役 config.xlsx 迁走组合方案。
            payload = {}
        except (OSError, ValueError, TypeError):
            return UserSettings()
        if not isinstance(payload, dict):
            return UserSettings()
        # 跨平台迁移：Windows 路径（C:\...）在 Linux 上会被当成一个超长相对路径，
        # 触发 Errno 36 等异常。凡是“非本平台原生绝对路径”的已保存路径一律清空，
        # 让用户重新选择，而不是让启动或自动识别崩溃。
        for key in ("last_input_dir", "last_template_dir", "last_external_file",
                    "last_output_dir", "history_config", "period_current_dir",
                    "period_previous_dir", "period_central_file", "period_output_dir",
                    "period_config_file", "central_current_csv", "central_previous_csv",
                    "central_output_dir", "central_cross_current_csv", "central_cross_output_dir",
                    "central_compare_file", "central_explanation_file", "central_template_file",
                    "central_form_output_dir",
                    "central_common_config", "central_comparison_config",
                    "central_cross_config", "central_forms_config"):
            if key in payload and not _is_native_path(payload.get(key)):
                payload[key] = ""
                self.sanitized = True
        # 迁移后失效的配置路径：旧布局（如已取消的 unified/config/）被删除后，
        # 保存的路径会指向不存在的目录；若其所在目录也不存在，说明文件已被整体
        # 移走，直接清空回退到程序默认 config 目录，绝不在旧位置凭空重建目录。
        for key in ("history_config", "period_config_file",
                    "central_common_config", "central_comparison_config",
                    "central_cross_config", "central_forms_config"):
            saved = str(payload.get(key) or "").strip()
            if saved and not Path(saved).exists() and not Path(saved).parent.is_dir():
                payload[key] = ""
                self.sanitized = True
        last_input_dir = str(payload.get("last_input_dir") or "")
        last_output_dir = str(payload.get("last_output_dir") or "")
        # 旧设置没有书钉字段时，保留用户原先手动改过的非默认输出目录。
        # “审核结果”“执行结果”“执行结果_skip”均为历次默认值；都不应
        # 被误判为用户手动固定的目录。
        legacy_pinned = bool(
            last_input_dir
            and last_output_dir
            and last_output_dir not in {
                str(Path(last_input_dir) / "审核结果"),
                str(Path(last_input_dir) / "执行结果"),
                str(Path(last_input_dir) / "执行结果_skip"),
            }
        )
        raw_plans = payload.get("combine_sheets_plans")
        raw_active_plan = payload.get("active_combine_sheets_plan_id")
        migrated_combine_settings = False
        if not isinstance(raw_plans, list):
            legacy = load_legacy_combine_settings(self.path.parent.parent)
            if legacy is not None:
                raw_plans, raw_active_plan = legacy
                migrated_combine_settings = True
            else:
                raw_plans, raw_active_plan = default_combine_plans(), "default"
        combine_plans, active_combine_plan = normalize_combine_settings(raw_plans, raw_active_plan)
        if migrated_combine_settings or payload.get("combine_sheets_plans") != combine_plans or _text(payload.get("active_combine_sheets_plan_id")) != active_combine_plan:
            self.sanitized = True
        legacy_write_flow_logs = bool(payload.get("write_flow_logs", False))
        show_run_detail_logs = bool(payload.get("show_run_detail_logs", legacy_write_flow_logs))
        export_run_logs = bool(payload.get("export_run_logs", legacy_write_flow_logs))
        show_performance_diagnostics = bool(payload.get("show_performance_diagnostics", False))
        return UserSettings(
            last_input_dir=last_input_dir,
            last_template_dir=str(payload.get("last_template_dir") or ""),
            last_external_file=str(payload.get("last_external_file") or ""),
            last_output_dir=last_output_dir,
            history_config=str(payload.get("history_config") or ""),
            output_pinned=bool(payload.get("output_pinned", legacy_pinned)),
            # 说明报送文件通常按机构放在子目录，默认保持递归发现。
            recursive_folders=bool(payload.get("recursive_folders", True)),
            # 旧设置只有布尔递归开关；迁移：true→最深处，false→仅根目录。
            recursive_depth=int(payload.get("recursive_depth", -1 if bool(payload.get("recursive_folders", True)) else 0)),
            extra_files=[str(item) for item in payload.get("extra_files", []) if isinstance(item, str) and item.strip()],
            # 三项日志体验设置彼此独立；旧 write_flow_logs 迁移为前两项。
            show_run_detail_logs=show_run_detail_logs,
            export_run_logs=export_run_logs,
            show_performance_diagnostics=show_performance_diagnostics,
            # 旧字段规范化为“是否导出运行日志 Excel”，供旧 API/旧外壳读取。
            write_flow_logs=export_run_logs,
            # 执行前确认默认开启：误点主按钮时先看到文件清单，避免直接跑批。
            confirm_before_run=bool(payload.get("confirm_before_run", True)),
            ui_theme=(
                str(payload.get("ui_theme") or "浅色")
                if str(payload.get("ui_theme") or "浅色") in UI_THEMES
                else "浅色"
            ),
            file_picker_mode=(
                str(payload.get("file_picker_mode") or file_picker_default)
                if str(payload.get("file_picker_mode") or file_picker_default) in FILE_PICKER_MODES
                else file_picker_default
            ),
            conditional_format_evaluator=normalize_conditional_evaluator(
                payload.get("conditional_format_evaluator")
                or payload.get("conditional_format_engine_mode")),
            conditional_format_rule_reader=normalize_conditional_rule_reader(
                payload.get("conditional_format_rule_reader")),
            formula_region_writer=(
                str(payload.get("formula_region_writer") or "DIRECT_OOXML").strip().upper()
                if str(payload.get("formula_region_writer") or "DIRECT_OOXML").strip().upper()
                in FORMULA_REGION_WRITERS else "DIRECT_OOXML"
            ),
            external_sheet_writer=(
                str(payload.get("external_sheet_writer") or "DIRECT_OOXML").strip().upper()
                if str(payload.get("external_sheet_writer") or "DIRECT_OOXML").strip().upper()
                in EXTERNAL_SHEET_WRITERS else "OPENPYXL"
            ),
            calculation_engine=(
                str(payload.get("calculation_engine") or "自动")
                if str(payload.get("calculation_engine") or "自动") in CALCULATION_ENGINES
                else "自动"
            ),
            summary_read_engine=(
                str(payload.get("summary_read_engine") or "纯 Python")
                if str(payload.get("summary_read_engine") or "纯 Python") in SUMMARY_READ_ENGINES
                else "纯 Python"
            ),
            favorite_input_dirs=self._favorites(payload.get("favorite_input_dirs")),
            favorite_template_dirs=self._favorites(
                payload.get("favorite_template_dirs")
            ),
            favorite_external_files=self._favorites(
                payload.get("favorite_external_files")
            ),
            favorite_output_dirs=self._favorites(payload.get("favorite_output_dirs")),
            recent_input_dirs=self._paths(payload.get("recent_input_dirs")),
            recent_template_dirs=self._paths(payload.get("recent_template_dirs")),
            recent_external_files=self._paths(payload.get("recent_external_files")),
            recent_output_dirs=self._paths(payload.get("recent_output_dirs")),
            period_current_dir=str(payload.get("period_current_dir") or ""),
            period_previous_dir=str(payload.get("period_previous_dir") or ""),
            period_central_file=str(payload.get("period_central_file") or ""),
            period_output_dir=str(payload.get("period_output_dir") or ""),
            period_output_auto=bool(payload.get("period_output_auto", True)),
            period_config_file=str(payload.get("period_config_file") or ""),
            central_current_csv=str(payload.get("central_current_csv") or ""),
            central_previous_csv=str(payload.get("central_previous_csv") or ""),
            central_output_dir=str(payload.get("central_output_dir") or ""),
            central_output_auto=bool(payload.get("central_output_auto", True)),
            central_cross_current_csv=str(payload.get("central_cross_current_csv") or ""),
            central_cross_output_dir=str(payload.get("central_cross_output_dir") or ""),
            central_cross_output_auto=bool(payload.get("central_cross_output_auto", True)),
            central_compare_file=str(payload.get("central_compare_file") or ""),
            central_explanation_file=str(payload.get("central_explanation_file") or ""),
            central_template_file=str(payload.get("central_template_file") or ""),
            central_form_output_dir=str(payload.get("central_form_output_dir") or ""),
            central_form_output_auto=bool(payload.get("central_form_output_auto", True)),
            central_common_config=str(payload.get("central_common_config") or ""),
            central_comparison_config=str(payload.get("central_comparison_config") or ""),
            central_cross_config=str(payload.get("central_cross_config") or ""),
            central_forms_config=str(payload.get("central_forms_config") or ""),
            xlsx_render_mode=(
                str(payload.get("xlsx_render_mode") or DEFAULT_XLSX_RENDER_MODE)
                if str(payload.get("xlsx_render_mode") or DEFAULT_XLSX_RENDER_MODE) in XLSX_RENDER_MODES
                else DEFAULT_XLSX_RENDER_MODE
            ),
            central_rule_engine=(
                str(payload.get("central_rule_engine") or DEFAULT_CENTRAL_RULE_ENGINE)
                if str(payload.get("central_rule_engine") or DEFAULT_CENTRAL_RULE_ENGINE)
                in CENTRAL_RULE_ENGINES
                else DEFAULT_CENTRAL_RULE_ENGINE
            ),
            expression_evaluation_mode=(
                str(payload.get("expression_evaluation_mode") or DEFAULT_EXPRESSION_EVALUATION_MODE)
                if str(payload.get("expression_evaluation_mode") or DEFAULT_EXPRESSION_EVALUATION_MODE).upper()
                in EXPRESSION_EVALUATION_MODES
                else DEFAULT_EXPRESSION_EVALUATION_MODE
            ),
            central_expression_schema=(
                "LEGACY_8"
                if str(payload.get("central_expression_schema") or "").strip().upper()
                == "LEGACY_8"
                else "FIVE_SEGMENT_V1"
            ),
            expression_evaluation_backend=(
                str(payload.get("expression_evaluation_backend") or DEFAULT_EXPRESSION_EVALUATION_BACKEND)
                if str(payload.get("expression_evaluation_backend") or DEFAULT_EXPRESSION_EVALUATION_BACKEND).upper()
                in EXPRESSION_EVALUATION_BACKENDS
                else DEFAULT_EXPRESSION_EVALUATION_BACKEND
            ),
            central_diff_tolerance_yuan=self._tolerance(payload.get("central_diff_tolerance_yuan")),
            combine_sheets_plans=combine_plans,
            active_combine_sheets_plan_id=active_combine_plan,
        )

    @staticmethod
    def _tolerance(value: object) -> float:
        """大集中核对容差（元）：非法或越界回落到默认 100 元。"""
        try:
            tolerance = float(value)  # type: ignore[arg-type]
        except (TypeError, ValueError):
            return 100.0
        if tolerance < 0 or tolerance > 1_000_000:
            return 100.0
        return tolerance

    def save(self, settings: UserSettings) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = {"version": 1, **asdict(settings)}
        temporary = self.path.with_suffix(self.path.suffix + ".tmp")
        temporary.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        temporary.replace(self.path)

    @staticmethod
    def _paths(value: object) -> list[str]:
        if not isinstance(value, list):
            return []
        return [str(item) for item in value if isinstance(item, str) and item.strip()][
            :MAX_RECENT_PATHS
        ]

    @staticmethod
    def _favorites(value: object) -> list[FavoritePath]:
        if not isinstance(value, list):
            return []
        result = []
        for item in value:
            if not isinstance(item, dict) or not item.get("path"):
                continue
            result.append(
                FavoritePath(
                    str(item.get("name") or Path(str(item["path"])).name),
                    str(item["path"]),
                )
            )
        return result


def _text(value: object) -> str:
    return str(value or "").strip()


def remember_path(items: list[str], value: str) -> list[str]:
    normalized = str(Path(value).resolve())
    result = [normalized]
    result.extend(item for item in items if str(Path(item)) != normalized)
    return result[:MAX_RECENT_PATHS]


def add_favorite(
    items: list[FavoritePath], name: str, value: str
) -> list[FavoritePath]:
    normalized = str(Path(value).resolve())
    result = [item for item in items if str(Path(item.path)) != normalized]
    result.append(FavoritePath(name.strip() or Path(normalized).name, normalized))
    return result
