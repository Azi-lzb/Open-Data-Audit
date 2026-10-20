"""把现有 ``报表采集系统_比较配置.xlsx`` 转换为 V2 配置对象。"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .models import (
    ExternalCheckRule, IndicatorConfig, InstitutionProfile, PeriodBandConfig, ReportAuditConfig, UnitSettings,
    RuleConfig, SEVERITIES, SourceRef,
)
from .unit_conversion import UNIT_TO_YUAN


class V2ConfigError(ValueError):
    pass


_INDICATOR = {
    "指标代码": "code", "指标名称": "name", "表单代码": "form", "数据属性": "type",
    "环比启用": "period_enabled", "环比策略组": "period_group", "最小变动值": "min_change",
    "禁用": "disabled", "备注": "note",
}
_EXTERNAL_RULE = {
    "规则编号": "id", "指标代码": "indicator", "外部指标名称": "external_name",
    "比较方式": "method", "容差": "tolerance", "级别": "severity", "启用": "enabled",
}
EXTERNAL_RULE_HEADERS = ("规则编号", "指标代码", "比较方式", "外部指标名称", "容差", "级别", "启用")
EXTERNAL_COMPARISON_METHODS = frozenset({"等于", "不等于", "大于", "大于等于", "小于", "小于等于"})
_INSTITUTION = {
    "机构名称": "name", "社会信用代码": "code", "机构代码": "code",
    "机构类别": "type", "承接行": "branch", "地区": "region", "归属行": "region",
    "报表项目": "report_item", "禁用": "disabled",
}
_BAND = {
    "规则编号": "id", "下限": "lower", "上限": "upper", "分类说明": "desc",
    "级别": "severity", "禁用": "disabled",
}
_PERIOD = {
    "策略组": "group", "规则编号": "id", "数据属性": "type", "下限": "lower", "上限": "upper",
    "审核级别": "severity", "级别": "severity", "分类说明": "desc", "禁用": "disabled",
}
_RULE = {
    "规则编号": "id", "类型": "type", "描述": "desc", "规则内容": "expression",
    "Thd值(万元)": "threshold", "Thd值": "threshold", "阈值(万元)": "threshold",
    "取反": "invert", "取反标识": "invert", "级别": "severity", "禁用": "disabled", "备注": "note",
}

# 新版仅使用“硬性/软性”；读取旧配置时按原风险程度兼容映射。
LEGACY_SEVERITY_MAP = {
    "错误": "硬性", "严重": "硬性",
    "核实": "软性", "提示": "软性", "关注": "软性",
}
_SEVERITY_MAP = {**LEGACY_SEVERITY_MAP, "硬性": "硬性", "软性": "软性"}

GENERAL_SETTINGS_SHEET_NAME = "通用设置"
UNIT_SETTING_FIELDS = {
    "源数据单位": "source_unit",
    "外部文件单位": "external_file_unit",
    "输出文件单位": "output_file_unit",
}
BOOLEAN_SETTING_FIELDS = {
    "是否隐藏无变动指标": "hide_unchanged_period_rows",
    "是否隐藏与大集中一致指标": "hide_matching_external_rows",
}
ENUM_SETTING_FIELDS = {
    "文字变动默认级别": "text_change_severity",
    "本期缺失默认级别": "current_missing_severity",
    "上期缺失默认级别": "previous_missing_severity",
}
OPTIONAL_VALUE_SETTINGS = ("换算范围",)
# 可缺省的枚举设置：旧配置不写该行时按默认值执行，出现后取值必须合法。
OPTIONAL_ENUM_SETTINGS = {
    name: ("硬性", "软性") for name in ENUM_SETTING_FIELDS
}
DEFAULT_ENUM_SETTINGS = {
    "text_change_severity": "硬性",
    "current_missing_severity": "硬性",
    "previous_missing_severity": "硬性",
}
GENERAL_SETTINGS_FIELDS = (
    *UNIT_SETTING_FIELDS,
    *OPTIONAL_VALUE_SETTINGS,
    *BOOLEAN_SETTING_FIELDS,
    *ENUM_SETTING_FIELDS,
)


@dataclass(frozen=True)
class GeneralSettingsResult:
    """通用设置解析结果；缺失的设置行按默认执行并附提示。"""

    units: UnitSettings
    hide_unchanged_period_rows: bool
    hide_matching_external_rows: bool
    text_change_severity: str
    current_missing_severity: str
    previous_missing_severity: str
    missing_row_warnings: list[str] = field(default_factory=list)


def _load_general_settings_from_book(
    book, config_path: Path,
) -> GeneralSettingsResult:
    if GENERAL_SETTINGS_SHEET_NAME not in book.sheetnames:
        raise V2ConfigError(
            f"报表采集配置“{Path(config_path).name}”缺少必需工作表“{GENERAL_SETTINGS_SHEET_NAME}”；"
            "请在该配置工作簿中维护单位及显示设置。"
        )
    sheet = book[GENERAL_SETTINGS_SHEET_NAME]
    rows = sheet.iter_rows(values_only=True)
    header = next(rows, ())
    expected_header = ("设置项", "设置值", "说明")
    actual_header = tuple(_text(value) for value in header[:3])
    if actual_header != expected_header:
        raise V2ConfigError(
            f"“{GENERAL_SETTINGS_SHEET_NAME}”工作表表头必须为："
            "A1=设置项、B1=设置值、C1=说明。"
        )
    configured: dict[str, tuple[int, str]] = {}
    for row_number, row in enumerate(rows, start=2):
        key = _text(row[0] if len(row) > 0 else None)
        value = _text(row[1] if len(row) > 1 else None)
        if not key and not value:
            continue
        if not key:
            raise V2ConfigError(
                f"“{GENERAL_SETTINGS_SHEET_NAME}”工作表第{row_number}行设置项不能为空。"
            )
        if key not in GENERAL_SETTINGS_FIELDS:
            raise V2ConfigError(
                f"“{GENERAL_SETTINGS_SHEET_NAME}”工作表第{row_number}行设置项无效：{key}"
            )
        if key in configured:
            raise V2ConfigError(f"“{GENERAL_SETTINGS_SHEET_NAME}”工作表设置项重复：{key}")
        configured[key] = (row_number, value)

    # 设置行允许缺失：按默认执行并在结果中提示补齐，不阻断运行；
    # 换算范围是纯说明行，缺行不提示。提示文案避开“不存在/无效/已跳过/
    # 暂不支持”字样，保持配置检查中的 warning 级别，不升级为阻断。
    default_display = {
        **{name: "元" for name in UNIT_SETTING_FIELDS},      # 与 UnitSettings 字段默认一致
        **{name: "否" for name in BOOLEAN_SETTING_FIELDS},   # 与 False 默认一致
        **{name: DEFAULT_ENUM_SETTINGS[field] for name, field in ENUM_SETTING_FIELDS.items()},
    }
    missing_row_warnings = [
        f"通用设置缺少设置项“{name}”，已按默认“{default_display[name]}”执行；请补齐该行后重跑。"
        for name in GENERAL_SETTINGS_FIELDS
        if name not in configured and name not in OPTIONAL_VALUE_SETTINGS
    ]

    unit_settings: dict[str, str] = {}
    for name, field in UNIT_SETTING_FIELDS.items():
        if name not in configured:
            continue
        row_number, value = configured[name]
        if value not in UNIT_TO_YUAN:
            raise V2ConfigError(
                f"“{GENERAL_SETTINGS_SHEET_NAME}”工作表第{row_number}行“{name}”"
                f"必须为元、万元或亿元：{value or '空白'}"
            )
        unit_settings[field] = value

    boolean_settings: dict[str, bool] = {}
    for name, field in BOOLEAN_SETTING_FIELDS.items():
        if name not in configured:
            continue
        row_number, value = configured[name]
        if value not in {"是", "否"}:
            raise V2ConfigError(
                f"“{GENERAL_SETTINGS_SHEET_NAME}”工作表第{row_number}行“{name}”"
                f"必须为“是”或“否”：{value or '空白'}"
            )
        boolean_settings[field] = value == "是"

    enum_settings: dict[str, str] = {}
    for name, field in ENUM_SETTING_FIELDS.items():
        if name not in configured:
            enum_settings[field] = DEFAULT_ENUM_SETTINGS[field]
            continue
        row_number, value = configured[name]
        allowed = OPTIONAL_ENUM_SETTINGS[name]
        if value not in allowed:
            raise V2ConfigError(
                f"“{GENERAL_SETTINGS_SHEET_NAME}”工作表第{row_number}行“{name}”"
                f"必须为{'或'.join(allowed)}：{value or '空白'}"
            )
        enum_settings[field] = value

    return GeneralSettingsResult(
        units=UnitSettings(**unit_settings),
        hide_unchanged_period_rows=boolean_settings.get("hide_unchanged_period_rows", False),
        hide_matching_external_rows=boolean_settings.get("hide_matching_external_rows", False),
        text_change_severity=enum_settings["text_change_severity"],
        current_missing_severity=enum_settings["current_missing_severity"],
        previous_missing_severity=enum_settings["previous_missing_severity"],
        missing_row_warnings=missing_row_warnings,
    )


def load_unit_settings(config_path: Path) -> UnitSettings:
    """读取主 S2 配置工作簿“通用设置”中的单位设置。"""
    path = Path(config_path)
    if not path.is_file():
        raise V2ConfigError(f"V2 比较配置不存在：{path}")
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        unit_settings = _load_general_settings_from_book(book, path).units
        return unit_settings
    finally:
        book.close()


def _text(value: Any) -> str:
    return str(value or "").strip()


def _code(value: Any) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = _text(value)
    # 兼容字符串形态的浮点指标代码（如 Excel 文本单元格 "20201028.0"），
    # 与 float 读出形态归一到同一代码，避免配置 lookup miss。
    if re.fullmatch(r"\d+\.0+", text):
        return str(int(float(text)))
    return text


def _yes(value: Any) -> bool:
    return _text(value).casefold() in {"是", "1", "true", "y", "yes"}


def _enabled(value: Any) -> bool:
    """新规则表的空白启用列按“是”处理，便于逐步维护。"""
    return True if value in (None, "") else _yes(value)


def _number(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _severity(value: Any) -> str:
    return _SEVERITY_MAP.get(_text(value), _text(value))


def _sheet(book, name: str):
    return book[name] if name in book.sheetnames else None


def _rows(sheet, aliases: dict[str, str]) -> tuple[dict[str, int], Iterable[tuple[Any, ...]]]:
    rows = sheet.iter_rows(values_only=True)
    header = next(rows, ())
    mapping: dict[str, int] = {}
    for index, cell in enumerate(header):
        field = aliases.get(_text(cell))
        if field and field not in mapping:
            mapping[field] = index
    return mapping, rows


def _get(row: tuple[Any, ...], mapping: dict[str, int], field: str) -> Any:
    index = mapping.get(field)
    return row[index] if index is not None and index < len(row) else None


def _validate_band_text(config: ReportAuditConfig, band: PeriodBandConfig) -> None:
    """描述不参与计算，但发现明显的正负号笔误时给配置人员提示。"""
    text = band.description.replace(" ", "")
    if "降幅" not in text:
        return
    # 例如“降幅(-80%,50%]”很可能遗漏第二个负号；只提示，绝不改边界。
    match = re.search(r"\((-?\d+(?:\.\d+)?)%[,，](-?\d+(?:\.\d+)?)%", text)
    if match and float(match.group(1)) < 0 < float(match.group(2)):
        config.warnings.append(
            f"环比规则“{band.rule_id or band.description}”描述疑似正负号异常：{band.description}；"
            "V2 将只按下限/上限数值执行，未自动修改配置。"
        )


def load_config(path: Path) -> ReportAuditConfig:
    if not path.is_file():
        raise V2ConfigError(f"V2 比较配置不存在：{path}")
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        general = _load_general_settings_from_book(book, path)
        config = ReportAuditConfig(
            unit_settings=general.units,
            hide_unchanged_period_rows=general.hide_unchanged_period_rows,
            hide_matching_external_rows=general.hide_matching_external_rows,
            text_change_severity=general.text_change_severity,
            current_missing_severity=general.current_missing_severity,
            previous_missing_severity=general.previous_missing_severity,
        )
        # 缺失的设置行按默认执行的提示：运行时进校验规则表，配置检查里保持 warning 级。
        config.warnings.extend(general.missing_row_warnings)
        indicator_sheet = _sheet(book, "指标参照")
        if indicator_sheet is None:
            raise V2ConfigError("比较配置缺少“指标参照”")
        headers = {_text(value) for value in next(indicator_sheet.iter_rows(values_only=True), ())}
        inline_unit_headers = headers.intersection({"源数据单位", "大集中数据单位"})
        if inline_unit_headers:
            raise V2ConfigError(
                f"单位统一配置应位于本工作簿的“{GENERAL_SETTINGS_SHEET_NAME}”工作表；"
                f"请从“指标参照”删除旧字段：{'、'.join(sorted(inline_unit_headers))}"
            )
        mapping, rows = _rows(indicator_sheet, _INDICATOR)
        if "code" not in mapping:
            raise V2ConfigError("“指标参照”缺少“指标代码”")
        for index, row in enumerate(rows, start=2):
            code = _code(_get(row, mapping, "code"))
            if not code:
                continue
            if code in config.indicators:
                config.warnings.append(f"指标参照存在重复指标代码：{code}")
                continue
            form_code = _code(_get(row, mapping, "form"))
            if not form_code:
                config.warnings.append(f"指标参照第{index}行指标 {code} 未填写表单代码")
            config.indicators[code] = IndicatorConfig(
                indicator_code=code,
                indicator_name=_text(_get(row, mapping, "name")),
                form_code=form_code,
                data_type=_text(_get(row, mapping, "type")),
                disabled=_yes(_get(row, mapping, "disabled")),
                note=_text(_get(row, mapping, "note")),
                period_enabled=_enabled(_get(row, mapping, "period_enabled")),
                period_strategy_group=_text(_get(row, mapping, "period_group")),
                min_change_value=_number(_get(row, mapping, "min_change")),
            )

        external_sheet = _sheet(book, "外部核对规则")
        if external_sheet is None:
            raise V2ConfigError("比较配置缺少“外部核对规则”")
        actual_external_headers = [
            _text(value) for value in next(external_sheet.iter_rows(values_only=True), ())
        ]
        while actual_external_headers and not actual_external_headers[-1]:
            actual_external_headers.pop()
        if tuple(actual_external_headers) != EXTERNAL_RULE_HEADERS:
            raise V2ConfigError(
                "“外部核对规则”表头必须依次为："
                + "、".join(EXTERNAL_RULE_HEADERS)
                + "；不再使用“外部系统”列。"
            )
        mapping, rows = _rows(external_sheet, _EXTERNAL_RULE)
        inline_external_unit_headers = {"大集中数据单位"}.intersection(
            {_text(value) for value in next(external_sheet.iter_rows(values_only=True), ())}
        )
        if inline_external_unit_headers:
            raise V2ConfigError(
                f"单位统一配置应位于本工作簿的“{GENERAL_SETTINGS_SHEET_NAME}”工作表；"
                "请从“外部核对规则”删除“大集中数据单位”字段"
            )
        for index, row in enumerate(rows, start=2):
            rule_id = _text(_get(row, mapping, "id"))
            indicator_code = _code(_get(row, mapping, "indicator"))
            if not rule_id and not indicator_code:
                continue
            if not rule_id or not indicator_code:
                config.warnings.append(f"外部核对规则第{index}行缺少规则编号或指标代码，已跳过")
                continue
            if indicator_code not in config.indicators:
                config.warnings.append(f"外部核对规则 {rule_id} 引用不存在指标：{indicator_code}，已跳过")
                continue
            external_indicator_name = _text(_get(row, mapping, "external_name"))
            method = _text(_get(row, mapping, "method"))
            tolerance = _number(_get(row, mapping, "tolerance"))
            severity = _severity(_get(row, mapping, "severity"))
            if not external_indicator_name or method not in EXTERNAL_COMPARISON_METHODS or tolerance is None or not math.isfinite(tolerance) or tolerance < 0:
                config.warnings.append(
                    f"外部核对规则第{index}行 {rule_id} 配置无效：比较方式须为等于/不等于/大于/大于等于/小于/小于等于，"
                    "容差须为非负有限数，外部指标名称不能为空；已跳过"
                )
                continue
            if method != "等于" and tolerance != 0:
                config.warnings.append(f"外部核对规则第{index}行 {rule_id} 仅“等于”使用容差；其他比较方式容差必须为0，已跳过")
                continue
            if severity not in SEVERITIES:
                config.warnings.append(f"外部核对规则 {rule_id} 级别无效：{severity}，按硬性处理")
                severity = "硬性"
            config.external_rules.append(ExternalCheckRule(
                rule_id=rule_id, indicator_code=indicator_code, external_system="大集中",
                external_indicator_name=external_indicator_name, comparison_method=method,
                tolerance=tolerance, severity=severity, enabled=_enabled(_get(row, mapping, "enabled")),
                source=SourceRef(str(path.resolve()), "外部核对规则", index),
            ))

        institution_sheet = _sheet(book, "机构参照")
        if institution_sheet is not None:
            mapping, rows = _rows(institution_sheet, _INSTITUTION)
            for row in rows:
                if _yes(_get(row, mapping, "disabled")):
                    continue
                name, code = _text(_get(row, mapping, "name")), _code(_get(row, mapping, "code"))
                if not name and not code:
                    continue
                profile = InstitutionProfile(
                    institution_id=code or name, institution_name=name or code,
                    social_credit_code=code, institution_type=_text(_get(row, mapping, "type")),
                    handling_branch=_text(_get(row, mapping, "branch")),
                    region=_text(_get(row, mapping, "region")), report_item=_text(_get(row, mapping, "report_item")),
                )
                if code:
                    config.institutions_by_id[code] = profile
                if name:
                    config.institutions_by_name[name] = profile

        band_sheet_name = "环比策略" if _sheet(book, "环比策略") is not None else "环比规则"
        band_sheet = _sheet(book, band_sheet_name)
        if band_sheet is not None:
            is_strategy_sheet = band_sheet_name == "环比策略"
            mapping, raw_rows = _rows(band_sheet, _PERIOD if is_strategy_sheet else _BAND)
            pending = list(raw_rows)
            for index, row in enumerate(pending):
                # 下限空=无穷小，上限空=无穷大；两者都填才是有限区间。
                lower = _number(_get(row, mapping, "lower"))
                upper = _number(_get(row, mapping, "upper"))
                if lower is None and upper is None and not is_strategy_sheet:
                    continue
                band = PeriodBandConfig(
                    rule_id=_text(_get(row, mapping, "id")) or f"B{index + 1:03d}", lower=lower, upper=upper,
                    description=_text(_get(row, mapping, "desc")), severity=_text(_get(row, mapping, "severity")),
                    disabled=_yes(_get(row, mapping, "disabled")),
                    strategy_group=_text(_get(row, mapping, "group")) if is_strategy_sheet else "",
                    data_type=_text(_get(row, mapping, "type")) if is_strategy_sheet else "",
                )
                if band.lower is not None and band.upper is not None and band.upper <= band.lower:
                    config.warnings.append(f"环比规则 {band.rule_id} 上限必须大于下限，已跳过")
                    continue
                if band.severity and _severity(band.severity) not in SEVERITIES:
                    config.warnings.append(f"环比规则 {band.rule_id} 级别无效：{band.severity}，按软性处理")
                    band = PeriodBandConfig(**{**band.__dict__, "severity": "软性"})
                elif band.severity:
                    band = PeriodBandConfig(**{**band.__dict__, "severity": _severity(band.severity)})
                config.bands.append(band)
                _validate_band_text(config, band)
            config.bands.sort(key=lambda item: (item.strategy_group, float("-inf") if item.lower is None else item.lower))
            for left, right in zip(config.bands, config.bands[1:]):
                if left.strategy_group != right.strategy_group:
                    continue
                if right.lower is None:
                    config.warnings.append(f"环比规则重叠：{left.rule_id} 与 {right.rule_id}（右档下限为空=无穷小）")
                    continue
                if left.upper is not None and left.upper > right.lower:
                    config.warnings.append(f"环比规则重叠：{left.rule_id} 与 {right.rule_id}")
                if left.upper is not None and left.upper < right.lower:
                    config.warnings.append(f"环比规则空洞：{left.rule_id} 与 {right.rule_id}")

        rule_sheet = _sheet(book, "校验规则")
        if rule_sheet is not None:
            mapping, rows = _rows(rule_sheet, _RULE)
            for index, row in enumerate(rows, start=2):
                expression = _text(_get(row, mapping, "expression"))
                if not expression:
                    continue
                rule_id = _text(_get(row, mapping, "id")) or f"ROW{index}"
                severity = _severity(_get(row, mapping, "severity")) or "软性"
                if severity not in SEVERITIES:
                    config.warnings.append(f"规则 {rule_id} 级别无效：{severity}，按软性处理")
                    severity = "软性"
                rule_type = _text(_get(row, mapping, "type")) or "表达式"
                if rule_type not in {"表达式", "累计不降(当年)", "累计不降(历史)"}:
                    config.warnings.append(f"规则 {rule_id} 类型暂不支持：{rule_type}")
                    continue
                unknown = [code for code in re.findall(r"[\[{]([^\]}]+)[\]}]", expression) if code.strip().isdigit() and code.strip() not in config.indicators]
                if unknown:
                    config.warnings.append(f"规则 {rule_id} 引用不存在指标：{'、'.join(unknown)}")
                config.rules.append(RuleConfig(
                    rule_id=rule_id, rule_type=rule_type, description=_text(_get(row, mapping, "desc")) or expression,
                    expression=expression, severity=severity, threshold=_number(_get(row, mapping, "threshold")),
                    invert=_yes(_get(row, mapping, "invert")), disabled=_yes(_get(row, mapping, "disabled")),
                    note=_text(_get(row, mapping, "note")),
                    source=SourceRef(str(path.resolve()), "校验规则", index),
                ))
    finally:
        book.close()
    return config


# 默认配置表的表头（与读取器一致）；生成后可直接在 Excel 中维护。
_DEFAULT_HEADERS = {
    "指标参照": ["指标代码", "指标名称", "表单代码", "数据属性", "环比启用", "环比策略组", "最小变动值", "禁用", "备注"],
    "机构参照": ["机构名称", "社会信用代码", "机构类别", "承接行", "地区", "报表项目", "禁用"],
    "环比策略": ["策略组", "规则编号", "数据属性", "下限", "上限", "审核级别", "分类说明", "禁用"],
    "校验规则": ["规则编号", "类型", "描述", "规则内容", "级别", "禁用", "备注"],
    "外部核对规则": list(EXTERNAL_RULE_HEADERS),
    GENERAL_SETTINGS_SHEET_NAME: ["设置项", "设置值", "说明"],
}
_DEFAULT_BANDS = [
    (None, -96.0, "降幅低于-96%", "硬性"),
    (-96.0, -90.0, "降幅(-96%,-90%]", "硬性"),
    (-90.0, -80.0, "降幅(-90%,-80%]", "软性"),
    (-80.0, -50.0, "降幅(-80%,-50%]", "软性"),
    (-50.0, -30.0, "降幅(-50%,-30%]", "软性"),
    (-30.0, 0.0, "降幅[30%以内]", "软性"),
    (0.0, 30.0, "增幅[0,30%)", "软性"),
    (30.0, 50.0, "增幅[30%,50%)", "软性"),
    (50.0, 100.0, "增幅[50%,1倍)", "软性"),
    (100.0, 500.0, "增幅[1倍,5倍)", "软性"),
    (500.0, 1000.0, "增幅[5倍,10倍)", "硬性"),
    (1000.0, None, "增幅10倍以上，请核实", "硬性"),
]


def write_default_config(path: Path) -> Path:
    """生成一份 V2 可用的默认配置工作簿（5 张业务表、通用设置 + 使用说明）。

    用于“报表采集系统_配置.xlsx”缺失或用户在设置中心点击重置时重建；
    只写表头和少量明确标为“禁用”的示例行，不写入任何真实机构数据。
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    from ...config_guide import write_guide_sheet

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    try:
        first = True
        for sheet_name, headers in _DEFAULT_HEADERS.items():
            sheet = book.active if first else book.create_sheet(sheet_name)
            first = False
            sheet.title = sheet_name
            sheet.append(headers)
        book["指标参照"].append(["20201001", "示例指标（请替换）", "20201", "文字", "是", "不适用", "不适用", "是", "示例行：确认后清空“禁用”即可使用"])
        book["机构参照"].append(["示例银行股份有限公司", "910000000000000000", "农村商业银行", "", "", "", "是"])
        for index, (lower, upper, desc, severity) in enumerate(_DEFAULT_BANDS, start=1):
            book["环比策略"].append(["BAL01", f"B{index:03d}", "余额", lower, upper, severity, desc, "否"])
        book["校验规则"].append([
            "R001", "表达式", "示例：按需改写规则内容", "[示例指标（请替换）] <> 0", "软性", "是",
            "示例行：[代码]取当期、{代码}取上期，也可写指标名称；确认后把“禁用”清空即可启用",
        ])
        book["外部核对规则"].append(["D001", "20201001", "等于", "示例外部指标（请替换）", 0, "软性", "否"])
        general_settings = book[GENERAL_SETTINGS_SHEET_NAME]
        general_settings.append(["源数据单位", "元", "本期和上期报表中的余额、累发金额单位。"])
        general_settings.append(["外部文件单位", "元", "外部核对文件中的余额、累发金额单位。"])
        general_settings.append(["输出文件单位", "元", "源数据与外部数据先统一换算到此单位，再进行计算并输出。"])
        general_settings.append(["换算范围", "", "仅换算余额和累发；百分数、个数、文字不换算；换算不做四舍五入。"])
        general_settings.append(["是否隐藏无变动指标", "否", "选择“是”：隐藏结果工作簿“环比规则”中审核说明为“无变动”的整行；数据保留，可取消隐藏。"])
        general_settings.append(["是否隐藏与大集中一致指标", "否", "选择“是”：隐藏结果工作簿“外部核对规则”中审核说明为“与大集中系统数据一致”的整行；数据保留，可取消隐藏。"])
        general_settings.append(["文字变动默认级别", "硬性", "文字指标的默认审核级别：两期不一致（审核说明“文字指标变化”）或单侧缺失（本期无，上期有/本期有，上期无）时使用。可选硬性或软性，此行可不填，缺省按硬性执行。"])
        general_settings.append(["本期缺失默认级别", "硬性", "非文字指标（余额、累发、百分数、个数）审核说明为“本期无，上期有”的环比行的默认审核级别：本期整行缺少指标、本期无数值，或本期值为0按无处理（不计算变动幅度）均属此类；文字指标不由此项控制。可选硬性或软性，此行可不填，缺省按硬性执行。"])
        general_settings.append(["上期缺失默认级别", "硬性", "非文字指标（余额、累发、百分数、个数）审核说明为“本期有，上期无”的环比行的默认审核级别：上期整行缺少指标、上期无数值，或上期值为0按无处理（不计算变动幅度）均属此类；文字指标不由此项控制。可选硬性或软性，此行可不填，缺省按硬性执行。"])
        write_guide_sheet(book, "period", index=0)

        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="1F4E78")
        body_font = Font(color="000000")
        body_fill = PatternFill(fill_type=None)
        for sheet in book.worksheets:
            sheet.freeze_panes = "A2"
            if sheet.title != "使用说明":
                for cell in sheet[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                for row in sheet.iter_rows(min_row=2):
                    for cell in row:
                        cell.font = body_font
                        cell.fill = body_fill
                for column in sheet.columns:
                    width = max(12, min(46, max(len(str(cell.value or "")) for cell in column) + 2))
                    sheet.column_dimensions[column[0].column_letter].width = width
        general_settings.column_dimensions["A"].width = 28
        general_settings.column_dimensions["B"].width = 14
        general_settings.column_dimensions["C"].width = 76
        general_settings.row_dimensions[5].height = 34
        from openpyxl.worksheet.datavalidation import DataValidation
        unit_validation = DataValidation(type="list", formula1='"元,万元,亿元"', allow_blank=False)
        unit_validation.error = "单位只能选择元、万元或亿元。"
        unit_validation.errorTitle = "单位无效"
        unit_validation.showErrorMessage = True
        unit_validation.prompt = "金额单位仅支持元、万元、亿元。"
        unit_validation.promptTitle = "选择金额单位"
        unit_validation.showInputMessage = True
        general_settings.add_data_validation(unit_validation)
        unit_validation.add("B2:B4")
        visibility_validation = DataValidation(type="list", formula1='"是,否"', allow_blank=False)
        visibility_validation.error = "只能选择是或否。"
        visibility_validation.errorTitle = "设置值无效"
        visibility_validation.showErrorMessage = True
        visibility_validation.prompt = "选择是否隐藏对应的正常结果。"
        visibility_validation.promptTitle = "选择是或否"
        visibility_validation.showInputMessage = True
        general_settings.add_data_validation(visibility_validation)
        visibility_validation.add("B6:B7")
        general_severity_validation = DataValidation(type="list", formula1='"硬性,软性"', allow_blank=True)
        general_severity_validation.error = "审核级别只能选择硬性或软性。"
        general_severity_validation.errorTitle = "审核级别无效"
        general_severity_validation.showErrorMessage = True
        general_severity_validation.prompt = "文字变动与本期/上期缺失的默认级别；未填写的设置行按硬性执行。"
        general_severity_validation.promptTitle = "选择默认审核级别"
        general_severity_validation.showInputMessage = True
        general_settings.add_data_validation(general_severity_validation)
        general_severity_validation.add("B8:B10")
        comparison_validation = DataValidation(
            type="list", formula1='"等于,不等于,大于,大于等于,小于,小于等于"', allow_blank=False,
        )
        comparison_validation.error = "比较方式只能选择等于、不等于、大于、大于等于、小于或小于等于。"
        comparison_validation.errorTitle = "比较方式无效"
        comparison_validation.showErrorMessage = True
        comparison_validation.prompt = "左边为报表指标值，右边为大集中外部指标值。"
        comparison_validation.promptTitle = "选择比较方式"
        comparison_validation.showInputMessage = True
        book["外部核对规则"].add_data_validation(comparison_validation)
        comparison_validation.add("C2:C1000")
        for sheet_name, header_name in (("环比策略", "审核级别"), ("校验规则", "级别"), ("外部核对规则", "级别")):
            sheet = book[sheet_name]
            column = next(cell.column for cell in sheet[1] if cell.value == header_name)
            used_row_count = sheet.max_row
            last_column = get_column_letter(sheet.max_column)
            severity_validation = DataValidation(type="list", formula1='"硬性,软性"', allow_blank=True)
            severity_validation.error = "审核级别只能选择硬性或软性。"
            severity_validation.errorTitle = "审核级别无效"
            severity_validation.showErrorMessage = True
            severity_validation.prompt = "硬性对应旧严重/错误；软性合并旧提示/关注/核实。该级别用于筛选，不改变流程是否阻断。"
            severity_validation.promptTitle = "选择审核级别"
            severity_validation.showInputMessage = True
            sheet.add_data_validation(severity_validation)
            column_letter = get_column_letter(column)
            severity_validation.add(f"{column_letter}2:{column_letter}1000")
            sheet.auto_filter.ref = f"A1:{last_column}{used_row_count}"
        book.save(target)
    except PermissionError as exc:
        raise V2ConfigError(f"无法写入报表采集配置：请先关闭“{target.name}”后重试") from exc
    finally:
        book.close()
    return target
