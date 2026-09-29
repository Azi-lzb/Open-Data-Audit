"""大集中统计系统 V3 配置工作簿。

运行时只读取 3.0 通用配置与当前功能配置（3.1/3.2/3.3）。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

SYSTEM_DESCRIPTION_HEADER = "系统生成说明"

COMMON_CONFIG_NAME = "3.0大集中通用配置.xlsx"
COMPARISON_CONFIG_NAME = "3.1大集中执行比较_配置.xlsx"
CROSS_CONFIG_NAME = "3.2大集中本期数值核对_配置.xlsx"
FORM_CONFIG_NAME = "3.3大集中指标比较拆分_配置.xlsx"

RUN_SHEET = "运行参数"
INDICATOR_SHEET = "指标参照"
ALERT_SHEET = "环比警戒"
UNIT_EXCEPTION_SHEET = "单位换算例外"
CUMULATIVE_RULE_SHEET = "累计指标规则"
SPECIAL_RULE_SHEET = "特殊指标规则"
# LEGACY_8 旧 8 段表：主名「表达式校验（兼容）」；老配置回退名「表达式校验」。
EXPRESSION_RULE_SHEET = "表达式校验（兼容）"
EXPRESSION_RULE_SHEET_OLD = "表达式校验"
# FIVE_SEGMENT_V1 新 5 段表（默认模式）：主名「表达式校验」；
# 本轮迁移版回退名「表达式校验5段式」。机构/地区为规则级作用域列，
# token [指标,数据属性,币种,频度,批次]。与旧 8 段表并存，均为正式路径；
# 名字冲突由选择顺序（兼容名优先、专名优先）与 5 段结构体检双重防呆。
FIVE_SEGMENT_RULE_SHEET = "表达式校验"
FIVE_SEGMENT_RULE_SHEET_MIGRATED = "表达式校验5段式"
ACTION_RULE_SHEET = "规则动作"
CROSS_SHEET = "跨期数值核对"
FORM_SHEET = "报表清单"
TOFORM_SHEET = "转表设置"
# 机构说明导出/回写时按 机构类代码+地区代码 映射机构名称（explanation_exporter）。
ORG_REFERENCE_SHEET = "机构地区参照"

# 历史 8 段表以旧名「表达式校验」存在时的旧列形态。它和现在的 5 段
# 主表同名，不能放入 SHEET_HEADERS（字典键会冲突），仅在读取回退时使用。
LEGACY_EXPRESSION_RULE_OLD_HEADERS = (
    "规则编号", "来源分组", "来源表单", "规则说明", "校验规则", "触发方式",
    "Thd值(万元)", "取反标识", "容差值(万元)", "启用", "停用原因", "备注",
)

# 各表必需表头；读取时自动在数据区上方定位表头行。
SHEET_HEADERS: dict[str, tuple[str, ...]] = {
    RUN_SHEET: ("参数", "值", "说明"),
    INDICATOR_SHEET: ("指标代码", "指标名称", "表单代码", "行序号", "指标别名"),
    ALERT_SHEET: (
        "序号", "下限", "上限", "备注文字", "填充颜色", "是否整行填充",
        "金额变动阈值", "是否说明",
    ),
    UNIT_EXCEPTION_SHEET: (
        "规则编号", "来源表单", "指标代码", "指标名称", "启用", "备注",
    ),
    CUMULATIVE_RULE_SHEET: (
        "规则编号", "指标代码", "指标名称", "数据属性", "频度", "来源表单", "启用", "备注",
    ),
    SPECIAL_RULE_SHEET: (
        "规则编号", "来源分组", "适用表单", "指标代码", "指标名称", "数据属性", "币种", "频度", "批次",
        "适用场景", "规则动作", "人民币阈值(亿元)", "美元阈值(亿美元)", "变幅阈值(%)",
        "比较指标代码", "来源说明", "详细说明", "补充说明", "启用", "停用原因", "修改日期",
        SYSTEM_DESCRIPTION_HEADER, "备注",
    ),
    EXPRESSION_RULE_SHEET: (
        "规则编号", "来源分组", "来源表单", "规则说明", "校验表达式", "是否取反",
        "容差值(万元)", "启用", "停用原因", "备注",
    ),
    # V3「规则动作」表：可选（存在即启用 V3 引擎路径）；三列 适用频度/批次/场景
    # 承接 VBA 特殊指标表名解析结果，“当年累计”来源分组的行走累计动作。
    ACTION_RULE_SHEET: (
        "规则编号", "规则说明", "规则动作", "指标代码", "数据属性", "币种", "比较指标代码",
        "人民币阈值(亿元)", "美元阈值(亿美元)", "变幅阈值(%)",
        "来源分组", "适用表单", "适用频度", "适用批次", "适用场景",
        "详细说明", "启用", "停用原因", "备注",
    ),
    CROSS_SHEET: (
        "校验编码", "校验名称", "机构类代码", "地区代码", "前提条件", "left", "right",
        "校验公式", "校验提示", "取反标识", "无误是否提示", "是否逻辑校验", "禁用", "备注",
    ),
    FIVE_SEGMENT_RULE_SHEET: (
        "规则编号", "来源分组", "来源表单", "机构类代码", "地区代码", "规则说明",
        "校验表达式", "是否取反", "容差值(万元)", "启用", "停用原因", "备注",
    ),
    FORM_SHEET: ("报表代码", "报表名称", "频度", "批次"),
    TOFORM_SHEET: ("参数", "值", "说明"),
    # 仅表头示例，机构数据由用户维护（程序不得内置机构名称）。
    ORG_REFERENCE_SHEET: ("机构类代码", "机构类名称", "地区代码", "地区名称", "机构名称", "备注"),
}

# S3-CFG-01 v3.0.0 前的 3.1 配置用“触发方式”表达同一语义。保留读取
# 兼容，让历史配置可继续使用；正式新建/维护配置统一使用“是否取反”。
EXPRESSION_RULE_HEADERS_TRIGGER_MODE = tuple(
    "触发方式" if header == "是否取反" else header
    for header in SHEET_HEADERS[EXPRESSION_RULE_SHEET]
)
FIVE_SEGMENT_RULE_HEADERS_TRIGGER_MODE = tuple(
    "触发方式" if header == "是否取反" else header
    for header in SHEET_HEADERS[FIVE_SEGMENT_RULE_SHEET]
)

COMMON_SHEETS = (RUN_SHEET, ALERT_SHEET, UNIT_EXCEPTION_SHEET, ORG_REFERENCE_SHEET)
# V3 正式形态：动作规则统一维护在「规则动作」表。
COMPARISON_SHEETS = (
    INDICATOR_SHEET, ACTION_RULE_SHEET, EXPRESSION_RULE_SHEET,
)
CROSS_SHEETS = (CROSS_SHEET,)
FORM_SHEETS = (FORM_SHEET, TOFORM_SHEET)

CONFIG_PROFILES: dict[str, tuple[str, tuple[str, ...]]] = {
    "common": (COMMON_CONFIG_NAME, COMMON_SHEETS),
    "comparison": (COMPARISON_CONFIG_NAME, COMPARISON_SHEETS),
    "cross": (CROSS_CONFIG_NAME, CROSS_SHEETS),
    "forms": (FORM_CONFIG_NAME, FORM_SHEETS),
}

# 规则编号唯一性按 (工作表, 编号列) 校验。
_ID_COLUMNS = {
    UNIT_EXCEPTION_SHEET: "规则编号",
    ACTION_RULE_SHEET: "规则编号",
    EXPRESSION_RULE_SHEET: "规则编号",
    CROSS_SHEET: "校验编码",
}
COMPLEX_SOURCES = ("单频", "跨期", "年报", "结转", "自定义")

DEFAULT_RUN_PARAMS = {
    "默认源数据单位": "元",
    "默认目标单位": "亿元",
    "复杂校验软性颜色": "46",
}

DEFAULT_TOFORM_PARAMS = {
    "隐藏空行": "是",
    "空表删除": "是",
    "分析文件数据单位": "亿元",
}

DEFAULT_ALERTS = [
    # 配置口径：99 表示 99%，不使用 0.99 这类小数；运行时环比仍为 0.99。
    (0, None, -99, "缩小100倍以上", 10, 1, 0.5,
     "绝对值变动超过{金额变动阈值}{目标单位}，且{备注文字}。"),
    (1, -99, -90, "缩小10倍-100倍", 10, 1, "", ""),
    (2, -90, -80, "缩小5倍-10倍", 43, 1, "", ""),
    (3, -80, -50, "缩小1倍-5倍", 35, 1, "", ""),
    (4, -50, -30, "缩小1倍以内", 35, 0, "", ""),
    (5, -30, 0, "", 0, 0, "", ""),
    (6, 0, 30, "", 0, 0, "", ""),
    (7, 30, 50, "", 36, 0, "", ""),
    (8, 50, 100, "", 36, 0, "", ""),
    (9, 100, 500, "增加1倍-5倍", 38, 0, "", ""),
    (10, 500, 1000, "增加5倍-10倍", 38, 0, "", ""),
    (11, 1000, 10000, "增加10倍-100倍", 7, 0, "", ""),
    (12, 10000, None, "请核实", 3, 0, 0.5,
     "绝对值变动超过{金额变动阈值}{目标单位}，且{备注文字}。"),
]


class CentralConfigError(ValueError):
    """配置缺失/损坏的明确失败（不静默回退）。"""


#: 供界面层捕获的稳定别名（V3 迁移后旧表缺失等提示走同一类型）。
CENTRAL_CONFIG_ERROR = CentralConfigError


def check_central_config_file(path: Path | str | None) -> str:
    """检查配置工作簿是否缺失或损坏；正常返回空串，否则返回提示语。"""
    target = Path(path) if path else None
    if target is None or not target.is_file():
        return "大集中统计系统配置文件缺失"
    try:
        book = load_central_config(target)
    except CentralConfigError as exc:
        return str(exc)
    if not book.rules and not book.complex_rules and not book.cross_rules:
        return f"大集中统计系统配置没有任何规则：{target.name}"
    return ""


@dataclass
class CentralConfig:
    path: Path
    paths: tuple[Path, ...] = field(default_factory=tuple)
    run_params: dict[str, str] = field(default_factory=dict)
    indicators: list[dict[str, str]] = field(default_factory=list)
    alerts: list[dict] = field(default_factory=list)
    rules: list[dict[str, str]] = field(default_factory=list)
    action_rules: list[dict[str, str]] = field(default_factory=list)
    complex_rules: list[dict[str, str]] = field(default_factory=list)
    # 5 段式规则（FIVE_SEGMENT_V1）：机构/地区为规则级作用域列，
    # token [指标,属性,币种,频度,批次]。仅当工作簿含该表时加载。
    five_segment_rules: list[dict[str, str]] = field(default_factory=list)
    cross_rules: list[dict[str, str]] = field(default_factory=list)
    forms: list[dict[str, str]] = field(default_factory=list)
    toform_params: dict[str, str] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)

    def param(self, name: str, default: str = "") -> str:
        value = self.run_params.get(name, "")
        return str(value).strip() if str(value).strip() else default

    def exempt_indicators(self) -> set[str]:
        return {
            row.get("指标代码", "").replace("'", "").strip()
            for row in self.rules
            if row.get("类型") == "单位不转换" and not _is_disabled(row.get("禁用"))
        }


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def _is_disabled(cell) -> bool:
    """禁用列：填「是/1/true/y」为禁用；空或其他为启用。"""
    text = _text(cell)
    return text in {"是", "1", "true", "True", "Y", "y"}


def _locate_header(rows, required: tuple[str, ...], sheet: str) -> tuple[int, dict[str, int]]:
    for index, values in enumerate(rows):
        possible = {_text(v): col for col, v in enumerate(values) if _text(v)}
        if set(required) <= set(possible):
            return index, possible
    raise CentralConfigError(f"配置表「{sheet}」缺少必需列：{'、'.join(required)}")


def _read_rows(book, sheet: str, headers: tuple[str, ...] | None = None) -> list[dict[str, str]]:
    if sheet not in book.sheetnames:
        return []
    raw = [list(row) for row in book[sheet].iter_rows(values_only=True)]
    # “系统生成说明”是程序可随时补齐的展示列，不是执行所需业务输入；
    # 保持旧版配置也能被正常读取和执行。
    required_headers = tuple(
        header for header in (headers or SHEET_HEADERS[sheet])
        if not (sheet == SPECIAL_RULE_SHEET and header == SYSTEM_DESCRIPTION_HEADER)
    )
    header_index, positions = _locate_header(raw, required_headers, sheet)
    result: list[dict[str, str]] = []
    for offset, values in enumerate(raw[header_index + 1:], start=header_index + 2):
        if not any(_text(v) for v in values):
            continue
        row = {
            name: _text(values[col]) if col < len(values) else ""
            for name, col in positions.items()
        }
        row["__行号__"] = str(offset)
        result.append(row)
    return result


def _disabled_from_enabled(value) -> str:
    """新版配置统一填“启用”；只有明确的否值才视为禁用。"""
    return "是" if _text(value).casefold() in {"否", "0", "false", "n", "no"} else ""


def _with_origin(row: dict[str, str], sheet: str) -> dict[str, str]:
    result = dict(row)
    result["__工作表__"] = sheet
    return result


def _normalize_unit_exception(row: dict[str, str]) -> dict[str, str]:
    return _with_origin({
        "规则编号": row.get("规则编号", ""),
        "类型": "单位不转换",
        "指标代码": row.get("指标代码", ""),
        "指标名称": row.get("指标名称", ""),
        "说明/比较指标": row.get("来源表单", ""),
        "备注": row.get("备注", ""),
        "禁用": _disabled_from_enabled(row.get("启用")),
        "__行号__": row.get("__行号__", ""),
    }, UNIT_EXCEPTION_SHEET)


def _normalize_cumulative_rule(row: dict[str, str]) -> dict[str, str]:
    return _with_origin({
        "规则编号": row.get("规则编号", ""),
        "类型": "累计不应下降",
        "指标代码": row.get("指标代码", ""),
        "指标名称": row.get("指标名称", ""),
        "数据属性": row.get("数据属性", ""),
        "频度": row.get("频度", ""),
        "表单": row.get("来源表单", ""),
        "备注": row.get("备注", ""),
        "禁用": _disabled_from_enabled(row.get("启用")),
        "__行号__": row.get("__行号__", ""),
    }, CUMULATIVE_RULE_SHEET)


def _normalize_special_rule(row: dict[str, str]) -> dict[str, str]:
    comparison = row.get("比较指标代码", "") or row.get("来源说明", "")
    return _with_origin({
        "规则编号": row.get("规则编号", ""),
        "类型": "特殊阈值",
        "来源分组": row.get("来源分组", ""),
        "指标代码": row.get("指标代码", ""),
        "指标名称": row.get("指标名称", ""),
        "数据属性": row.get("数据属性", ""),
        "币种": row.get("币种", ""),
        "频度": row.get("频度", ""),
        "批次": row.get("批次", ""),
        "场景": row.get("适用场景", ""),
        "人民币阀值(亿元)": row.get("人民币阈值(亿元)", ""),
        "美元阀值(亿美元)": row.get("美元阈值(亿美元)", ""),
        "绝对值变幅(%)": row.get("变幅阈值(%)", ""),
        "说明/比较指标": comparison,
        "__比较指标代码__": row.get("比较指标代码", ""),
        "表单": row.get("适用表单", ""),
        "详细说明": row.get("详细说明", ""),
        "备注": row.get("规则动作", ""),
        "禁用": _disabled_from_enabled(row.get("启用")),
        "__行号__": row.get("__行号__", ""),
    }, SPECIAL_RULE_SHEET)


def _normalize_action_rule(row: dict[str, str]) -> dict[str, str]:
    """V3「规则动作」行 → 与旧三表完全相同的内部字典形状。

    来源分组=当年累计 → 累计不应下降形状（accu_keys 匹配）；
    其余 → 特殊阈值形状（动作文本在“备注”键，与旧特殊表归一化一致）。
    """
    if row.get("来源分组") == "当年累计":
        return _with_origin({
            "规则编号": row.get("规则编号", ""),
            "类型": "累计不应下降",
            "指标代码": row.get("指标代码", ""),
            "指标名称": row.get("规则说明", ""),
            "数据属性": row.get("数据属性", ""),
            "频度": row.get("适用频度", ""),
            "场景": row.get("适用场景", ""),
            "表单": row.get("适用表单", ""),
            "规则动作": row.get("规则动作", ""),
            "备注": row.get("备注", ""),
            "配置备注": row.get("备注", ""),
            "禁用": _disabled_from_enabled(row.get("启用")),
            "__行号__": row.get("__行号__", ""),
        }, ACTION_RULE_SHEET)
    return _with_origin({
        "规则编号": row.get("规则编号", ""),
        "类型": "特殊阈值",
        "来源分组": row.get("来源分组", ""),
        "指标代码": row.get("指标代码", ""),
        "指标名称": row.get("规则说明", ""),
        "数据属性": row.get("数据属性", ""),
        "币种": row.get("币种", ""),
        "频度": row.get("适用频度", ""),
        "批次": row.get("适用批次", ""),
        "场景": row.get("适用场景", ""),
        "人民币阀值(亿元)": row.get("人民币阈值(亿元)", ""),
        "美元阀值(亿美元)": row.get("美元阈值(亿美元)", ""),
        "绝对值变幅(%)": row.get("变幅阈值(%)", ""),
        "说明/比较指标": row.get("比较指标代码", ""),
        "__比较指标代码__": row.get("比较指标代码", ""),
        "表单": row.get("适用表单", ""),
        "详细说明": row.get("详细说明", ""),
        "备注": row.get("规则动作", ""),
        "配置备注": row.get("备注", ""),
        "禁用": _disabled_from_enabled(row.get("启用")),
        "__行号__": row.get("__行号__", ""),
    }, ACTION_RULE_SHEET)


def _inverted_from_expression_row(row: dict[str, str]) -> str:
    """归一化正式“是否取反”，并兼容旧“触发方式/取反标识”。"""
    value = _text(row.get("是否取反")).casefold()
    if value:
        return "1" if value in {"是", "1", "true", "y"} else ""
    trigger = _text(row.get("触发方式"))
    if trigger:
        return "1" if trigger == "表达式不成立" else ""
    return "1" if _text(row.get("取反标识")) in {"1", "是"} else ""


def _normalize_expression_rule(row: dict[str, str]) -> dict[str, str]:
    return _with_origin({
        "规则编号": row.get("规则编号", ""),
        "校验表单": row.get("来源表单", ""),
        "校验描述": row.get("规则说明", ""),
        "校验规则": row.get("校验表达式", ""),
        "取反标识": _inverted_from_expression_row(row),
        "Thd值(万元)": row.get("容差值(万元)", ""),
        "禁用": _disabled_from_enabled(row.get("启用")),
        "来源": row.get("来源分组", ""),
        "备注": row.get("备注", ""),
        "__行号__": row.get("__行号__", ""),
    }, EXPRESSION_RULE_SHEET)


def _read_alert_rows(book) -> list[dict[str, str]]:
    """读取 V3 明确上下限的警戒区间；不再兼容旧“仅下限”格式。"""
    return _read_rows(book, ALERT_SHEET)


def _format_percent_bound(value: float) -> str:
    """将配置中的百分数数值显示为百分比文字（99 → 99%）。"""
    return f"{value:g}%"


def format_alert_band_remark(band: dict | None) -> str:
    """生成“自动区间文字 | 人工备注”的半手写警戒说明。

    区间文字始终按程序实际口径 ``下限 <= 环比 < 上限`` 生成；
    ``备注文字`` 只维护业务补充，例如“缩小5倍-10倍”或“请核实”。
    普通无颜色、无备注、无金额门槛的区间不输出说明。
    """
    if not band:
        return ""
    manual = _text(band.get("备注文字"))
    active = any((
        manual,
        _text(band.get("填充颜色")) not in {"", "0"},
        _text(band.get("是否整行填充")) not in {"", "0"},
        _text(band.get("金额变动阈值")),
        _text(band.get("是否说明")),
    ))
    if not active:
        return ""
    lower = _as_number(band.get("下限"))
    upper = _as_number(band.get("上限"))
    if lower is None and upper is not None:
        auto = (f"降幅<{_format_percent_bound(upper)}" if upper <= 0 else f"变幅<{_format_percent_bound(upper)}")
    elif lower is not None and upper is None:
        auto = f"增幅≥{_format_percent_bound(lower)}" if lower >= 0 else f"变幅≥{_format_percent_bound(lower)}"
    elif lower is not None and upper is not None:
        prefix = "增幅" if lower >= 0 else ("降幅" if upper <= 0 else "变幅")
        auto = f"{prefix}[{_format_percent_bound(lower)},{_format_percent_bound(upper)})"
    else:
        auto = ""
    if auto and manual:
        return f"{auto} | {manual}"
    return auto or manual


def alert_band_for(ratio: float, alerts: list[dict]) -> dict | None:
    """返回命中档位；配置用百分数数值，运行环比仍为小数。

    例如运行时 ``ratio=0.99`` 表示 99%，匹配配置区间 ``[50, 100)``。
    """
    ratio_percent = ratio * 100.0
    for band in alerts:
        try:
            lower_text = _text(band.get("下限"))
            lower = float(lower_text) if lower_text else None
            upper_text = _text(band.get("上限"))
            upper = float(upper_text) if upper_text else None
        except (TypeError, ValueError):
            continue
        if lower is not None and round(ratio_percent - lower, 10) < 0:
            continue
        if upper is None or round(ratio_percent - upper, 10) < 0:
            return band
    return None


def _config_paths(path) -> tuple[Path, ...]:
    if isinstance(path, (str, Path)) or path is None:
        return (Path(path),) if path else ()
    return tuple(Path(item) for item in path if item)


def load_central_config(path) -> CentralConfig:
    """读取一册配置，或按顺序合并“通用配置 + 功能配置”。

    行式工作表按文件顺序追加；参数式工作表后册覆盖前册同名参数。拆分后的
    四册默认没有重名业务表，这条覆盖规则只用于将来的功能级参数扩展。
    """
    from openpyxl import load_workbook

    targets = _config_paths(path)
    if not targets:
        raise CentralConfigError("未找到大集中统计系统配置")
    config = CentralConfig(path=targets[0], paths=targets)
    row_targets = (
        (INDICATOR_SHEET, "indicators"), (ALERT_SHEET, "alerts"),
        (CROSS_SHEET, "cross_rules"), (FORM_SHEET, "forms"),
    )
    for target in targets:
        if not target.is_file():
            raise CentralConfigError(f"未找到大集中统计系统配置：{target}")
        try:
            book = load_workbook(target, read_only=True, data_only=True)
        except Exception as exc:
            raise CentralConfigError(f"大集中统计系统配置无法读取（{target.name}）：{exc}") from exc
        try:
            config.run_params.update({
                row["参数"]: row["值"]
                for row in _read_rows(book, RUN_SHEET) if row.get("参数")
            })
            for sheet_name, attribute in row_targets:
                rows = _read_alert_rows(book) if sheet_name == ALERT_SHEET else _read_rows(book, sheet_name)
                getattr(config, attribute).extend(rows)
            config.rules.extend(
                _normalize_unit_exception(row) for row in _read_rows(book, UNIT_EXCEPTION_SHEET)
            )
            # V3 可选「规则动作」表：存在即加载，供 rule_action_engine 使用；
            # 旧三表路径不受影响（双轨并存，见 V3 合并计划 §七）。
            config.action_rules.extend(
                _normalize_action_rule(row) for row in _read_rows(book, ACTION_RULE_SHEET)
            )
            # 旧 8 段表：优先主名「表达式校验（兼容）」（改名后），回退旧名
            # 「表达式校验」（未升级的老配置）。
            legacy_sheet = (
                EXPRESSION_RULE_SHEET if EXPRESSION_RULE_SHEET in book.sheetnames
                else (EXPRESSION_RULE_SHEET_OLD if EXPRESSION_RULE_SHEET_OLD in book.sheetnames
                      else None))
            if legacy_sheet:
                # 旧表列形态有两种变体（迁移前校验表达式/容差值 列 vs 更早的
                # 校验规则/Thd值(万元)/取反标识 列）；先试新变体，失败回退旧变体。
                variant_errors = []
                legacy_rows = None
                for legacy_headers in (
                    SHEET_HEADERS[EXPRESSION_RULE_SHEET],
                    EXPRESSION_RULE_HEADERS_TRIGGER_MODE,
                    LEGACY_EXPRESSION_RULE_OLD_HEADERS,
                ):
                    try:
                        legacy_rows = _read_rows(
                            book, legacy_sheet, headers=legacy_headers)
                        break
                    except CentralConfigError as exc:
                        variant_errors.append(str(exc))
                if legacy_rows is None:
                    raise CentralConfigError(variant_errors[-1])
                config.complex_rules.extend(
                    _normalize_expression_rule(row) for row in legacy_rows
                )
            # 新 5 段表：优先迁移版名「表达式校验5段式」；主名「表达式校验」
            # 仅在「表达式校验（兼容）」同时存在（改名后双表配置）时才认作
            # 5 段来源——只有旧名单表的老配置不误选（由 5 段体检兜底报缺）。
            five_sheet = None
            if FIVE_SEGMENT_RULE_SHEET_MIGRATED in book.sheetnames:
                five_sheet = FIVE_SEGMENT_RULE_SHEET_MIGRATED
            elif (FIVE_SEGMENT_RULE_SHEET in book.sheetnames
                    and EXPRESSION_RULE_SHEET in book.sheetnames):
                five_sheet = FIVE_SEGMENT_RULE_SHEET
            if five_sheet:
                five_errors = []
                for five_headers in (
                    SHEET_HEADERS[FIVE_SEGMENT_RULE_SHEET],
                    FIVE_SEGMENT_RULE_HEADERS_TRIGGER_MODE,
                ):
                    try:
                        config.five_segment_rules.extend(
                            _read_rows(book, five_sheet, headers=five_headers))
                        break
                    except CentralConfigError as exc:
                        five_errors.append(str(exc))
                else:
                    raise CentralConfigError(five_errors[-1])
            config.toform_params.update({
                row["参数"]: row["值"]
                for row in _read_rows(book, TOFORM_SHEET) if row.get("参数")
            })
        finally:
            book.close()
    return config


def check_central_config_bundle(common_path, feature_path, profile: str) -> str:
    """检查某功能所需的两册配置及必需工作表。"""
    from openpyxl import load_workbook

    if profile not in CONFIG_PROFILES:
        return f"未知大集中配置类型：{profile}"
    targets = _config_paths((common_path, feature_path))
    expected = set(COMMON_SHEETS + CONFIG_PROFILES[profile][1])
    actual: set[str] = set()
    try:
        for target in targets:
            if not target.is_file():
                return f"大集中配置文件缺失：{target.name}"
            book = load_workbook(target, read_only=True, data_only=True)
            try:
                actual.update(book.sheetnames)
            finally:
                book.close()
        missing = expected - actual
        if missing:
            return f"大集中配置缺少工作表：{'、'.join(sorted(missing))}"
        loaded = load_central_config(targets)
        errors = validate_central_config(loaded)
        if not errors:
            return ""
        summary = ""
        if profile == "comparison":
            from .complex_rule_engine import inspect_complex_rules

            compile_result = inspect_complex_rules(loaded)
            summary = (
                f"3.1 表达式校验：启用 {compile_result.enabled_rows}，"
                f"成功编译 {compile_result.compiled_rules}，失败 {compile_result.failed_rules}。"
            )
        return summary + f"共 {len(errors)} 项：" + "；".join(errors)
    except Exception as exc:
        return f"大集中配置无法读取：{exc}"


def check_action_rules(config: CentralConfig, *, stats_out: dict | None = None) -> list[str]:
    """只读核对 3.1 启用动作是否能进入正式动作分派。"""
    from .indicator_rule_engine import _split_multi_codes, supports_special_action

    errors: list[str] = []
    enabled = cumulative = special = recognized = 0
    for row in config.action_rules:
        if _is_disabled(row.get("禁用")):
            continue
        enabled += 1
        rule_id = row.get("规则编号") or "未填写编号"
        location = f"「{ACTION_RULE_SHEET}」第{row.get('__行号__') or '?'}行规则 {rule_id}"
        codes = _split_multi_codes(row.get("指标代码"))
        if not codes:
            errors.append(f"{location}：指标代码为空，规则无法进入动作执行索引；请人工核对。")
        if row.get("类型") == "累计不应下降":
            cumulative += 1
            if codes:
                recognized += 1
            continue
        special += 1
        action = str(row.get("备注") or "").strip()
        if not supports_special_action(action):
            errors.append(
                f"{location}：规则动作“{action or '空白'}”无法由当前系统动作分派器识别；"
                "请人工核对规则动作，程序未自动修改配置。"
            )
        elif codes:
            recognized += 1
    if stats_out is not None:
        stats_out.update({
            "total": len(config.action_rules), "enabled": enabled,
            "disabled": len(config.action_rules) - enabled,
            "cumulative": cumulative, "special": special,
            "recognized": recognized, "errors": len(errors),
        })
    return errors


def validate_central_config(config: CentralConfig) -> list[str]:
    """配置体检：编号重复、枚举、数值、正则、表达式语法、引用完整性。"""
    import re

    from .expression_parser import ExpressionError, evaluate_expression

    def check_expression(sheet: str, row: dict, column: str, expression: str) -> None:
        # [...]/​{...} 指标引用在求值前由引擎替换为数值；体检用哑值 1 代入。
        probe = re.sub(r"\[[^\[\]]*\]|\{[^{}]*\}", "1", expression)
        try:
            evaluate_expression(probe, {"left": 1.0, "right": 1.0, "Thd": 1.0})
        except ExpressionError as exc:
            errors.append(
                f"「{sheet}」第{row['__行号__']}行{column}表达式有误"
                f"（{row.get('规则编号') or row.get('校验编码') or row.get('校验名称')}）：{exc}"
            )

    errors: list[str] = list(config.warnings)
    errors.extend(check_action_rules(config))
    # 编号查重：只针对启用的行（禁用行为历史遗留时不应阻塞）。
    rows_by_sheet = {
        UNIT_EXCEPTION_SHEET: [row for row in config.rules if row.get("__工作表__") == UNIT_EXCEPTION_SHEET],
        ACTION_RULE_SHEET: list(config.action_rules),
        EXPRESSION_RULE_SHEET: config.complex_rules,
        CROSS_SHEET: config.cross_rules,
    }
    for sheet, id_column in _ID_COLUMNS.items():
        rows = rows_by_sheet[sheet]
        seen: dict[str, str] = {}
        for row in rows:
            if _is_disabled(row.get("禁用")):
                continue
            rule_id = row.get(id_column, "")
            if not rule_id:
                errors.append(f"「{sheet}」第{row['__行号__']}行缺少{id_column}")
                continue
            if rule_id in seen:
                errors.append(f"「{sheet}」第{row['__行号__']}行{id_column}重复：{rule_id}（首次出现在第{seen[rule_id]}行）")
            else:
                seen[rule_id] = row["__行号__"]
    for row in config.complex_rules:
        expression = row.get("校验规则", "")
        if expression and not _balanced_brackets(expression):
            errors.append(f"「表达式校验」第{row['__行号__']}行规则括号不配对：{row.get('规则编号')}")
        hint = _ifs_fallback_hint(expression)
        if hint:
            errors.append(
                f"「表达式校验」第{row['__行号__']}行{row.get('规则编号')}：{hint}")

    # 3.1 表达式规则的“编译存活”必须复用正式运行时解析逻辑。
    # 配置检查只读：报告问题和人工建议，不修改 Excel，也不在内存里偷偷修补。
    from .complex_rule_engine import inspect_complex_rules

    compile_result = inspect_complex_rules(config)
    errors.extend(issue.format_user_message() for issue in compile_result.issues)
    for row in config.cross_rules:
        for column in ("前提条件", "left", "right", "校验公式"):
            expression = row.get(column, "")
            if not expression:
                continue
            check_expression("跨期数值核对", row, column, expression)
    pair_rule_marks = (
        "指标相除值核查", "太小核查", "指标需要对应存在",
        "指标应相等核查", "指标相减应大于核查",
    )
    for sheet in (ACTION_RULE_SHEET,):
        for row in rows_by_sheet[sheet]:
            amplitude = _text(row.get("绝对值变幅(%)"))
            if re.fullmatch(r"20\d{2}-\d{2}-\d{2}(?:\s+\d{2}:\d{2}:\d{2})?", amplitude):
                errors.append(
                    f"「{sheet}」第{row['__行号__']}行变幅阈值不能填写日期：{amplitude}"
                )
            if _is_disabled(row.get("禁用")):
                continue
            action = row.get("备注", "")
            try:
                has_amplitude = float(amplitude) > 0
            except (TypeError, ValueError):
                has_amplitude = False
            if action == "指标变幅异常需说明。" and not has_amplitude:
                errors.append(
                    f"「{sheet}」第{row['__行号__']}行“指标变幅异常需说明。”必须填写变幅阈值(%)"
                )
            if any(mark in action for mark in pair_rule_marks) and not row.get("__比较指标代码__"):
                errors.append(
                    f"「{sheet}」第{row['__行号__']}行规则动作需要填写比较指标代码：{action}"
                )
    # “指标参照”只负责输出名称与排序，并不是业务数据字典。比较指标只要能在
    # 当期/上期 CSV 中按完整维度键找到即可，因此不能要求它必须出现在指标参照中。
    return errors


def _ifs_fallback_hint(expression: str) -> str:
    """IFS 未以 ``TRUE, 兜底值`` 结尾时的提示（不阻断执行）。

    Excel 与程序口径一致：IFS 无匹配返回 #N/A（按规则错误显形）。若规则未写
    TRUE 兜底，正常数据落进无匹配分支时会报 #N/A；此处只提示，让作者自查。
    """
    import re

    text = str(expression or "")
    if "IFS" not in text.upper():
        return ""
    for match in re.finditer(r"IFS\s*\(", text, re.IGNORECASE):
        depth = 1
        index = match.end()
        while index < len(text) and depth:
            if text[index] == "(":
                depth += 1
            elif text[index] == ")":
                depth -= 1
            index += 1
        inner = text[match.end():index - 1]
        depth = 0
        parts, current = [], ""
        for char in inner:
            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
            if char == "," and depth == 0:
                parts.append(current)
                current = ""
            else:
                current += char
        parts.append(current)
        if len(parts) % 2:
            continue
        last_condition = parts[-2].strip().upper() if len(parts) >= 2 else ""
        if last_condition != "TRUE":
            return "IFS 最后一对条件不是 TRUE（建议以「TRUE, 兜底值」结尾，否则无匹配时按 #N/A 报规则错误）"
    return ""


def _balanced_brackets(text: str) -> bool:
    normalized = text.replace("（", "(").replace("）", ")").replace("，", ",")
    return normalized.count("[") == normalized.count("]") and normalized.count("{") == normalized.count("}")


def _as_number(value) -> float | None:
    try:
        return float(_text(value))
    except (TypeError, ValueError):
        return None


def _format_number(value) -> str:
    number = _as_number(value)
    if number is None:
        return ""
    return f"{number:g}"


def _format_percent(value) -> str:
    number = _as_number(value)
    return f"{number * 100:g}%" if number is not None else ""


def _amount_threshold_text(rmb, usd, *, change: bool = True) -> str:
    subject = "较上期变动金额" if change else "金额"
    clauses = []
    if _as_number(rmb) not in (None, 0):
        clauses.append(f"人民币指标{subject}达到{_format_number(rmb)}亿元")
    if _as_number(usd) not in (None, 0):
        clauses.append(f"美元合计指标{subject}达到{_format_number(usd)}亿美元")
    return "，或".join(clauses)


def build_special_rule_description(row: dict[str, object]) -> str:
    """把已支持的特殊规则动作翻译为不参与执行的维护说明。"""
    action = _text(row.get("规则动作"))
    rmb = row.get("人民币阈值(亿元)")
    usd = row.get("美元阈值(亿美元)")
    amplitude = row.get("变幅阈值(%)")
    amount = _amount_threshold_text(rmb, usd)
    percent = _format_percent(amplitude)
    target = _text(row.get("比较指标代码"))

    if action == "指标不应有数，需说明。":
        return "该指标本期不应有数；出现任何数值时需说明。"
    if action == "指标有变动需说明。":
        if amount and percent:
            return f"当{amount}且该指标较上期变动幅度超过{percent}时，需说明。"
        if amount:
            return f"当{amount}时，需说明。"
        if percent:
            return f"该指标较上期变动幅度超过{percent}时，需说明。"
        return "该指标本期数值较上期发生变动时，需说明。"
    if action == "指标变动或变幅超过阈值需说明。":
        conditions = [amount, f"该指标较上期变动幅度超过{percent}" if percent else ""]
        conditions = [item for item in conditions if item]
        return (
            f"当{'，或'.join(conditions)}时，需说明。"
            if conditions else "系统无法自动解释：该规则未填写金额或变幅阈值。"
        )
    if action == "指标变幅异常需说明。":
        return (
            f"该指标较上期变动幅度达到或超过{percent}时，需说明。"
            if percent else "系统无法自动解释：该规则未填写变幅阈值(%)。"
        )
    if action in ("指标比上期增加需说明。", "指标比上期减少需说明。"):
        direction = "增加" if "增加" in action else "减少"
        threshold = _amount_threshold_text(rmb, usd)
        return f"该指标较上期{direction}{'且' if threshold else ''}{threshold}时，需说明。".replace("增加且", "增加且").replace("减少且", "减少且")
    if action in ("指标新增需说明。", "指标结清需说明。"):
        event = "由零变为有数" if "新增" in action else "由有数变为零"
        threshold = _amount_threshold_text(rmb, usd)
        return f"该指标较上期{event}{'且' if threshold else ''}{threshold}时，需说明。"
    if action == "指标为负数需核实。":
        return "该指标本期数值小于 0 时，需核实。"
    if action == "指标应为整数。":
        return "该指标本期数值不是整数时，需核实。"
    if action == "指标应能被某数整除。":
        divisor = _format_number(rmb)
        return (
            f"该指标按输出单位换算为元后，应能被{divisor}整除；不能整除时需说明。"
            if divisor and _as_number(rmb) not in (None, 0)
            else "系统无法自动解释：整除规则需在人民币阈值(亿元)列填写非零除数。"
        )
    if action == "指标本期余额应小于（上期余额+本期发生额）。":
        return "在相邻报告期内，校验本期余额不超过上期余额加本期发生额；余额或发生额不匹配时提示核实。"
    if action == "指标本期余额应小于（上期余额+本期当年发生额-上期当年累计发生额）。":
        return "在同年或年末衔接期间，校验本期余额不超过上期余额加本期当年发生额减上期当年累计发生额；不匹配时提示核实。"
    if action == "指标需要对应存在。":
        return (
            f"该指标与比较指标代码“{target}”应同步有数或无数；任一方有数、另一方无数时需说明。"
            if target else "系统无法自动解释：该规则未填写比较指标代码。"
        )
    if action == "指标应相等核查。":
        return (
            f"该指标应与比较指标代码“{target}”数值相等；存在差异时需核实。"
            if target else "系统无法自动解释：该规则未填写比较指标代码。"
        )
    if action in ("指标相除值核查。", "指标相除值太小核查。"):
        threshold = _format_number(usd if _as_number(usd) not in (None, 0) else rmb)
        if target and threshold:
            operator = "小于" if "太小" in action else "超过"
            return f"该指标与比较指标代码“{target}”的比值{operator}{threshold}时，需说明。"
        return "系统无法自动解释：相除规则需填写比较指标代码和比较阈值。"
    if action == "指标相减应大于核查。":
        threshold = _amount_threshold_text(rmb, usd, change=False)
        if target and threshold:
            return f"该指标减去比较指标代码“{target}”后的差额小于{threshold.replace('人民币指标金额达到', '').replace('美元合计指标金额达到', '')}时，需核实。"
        return "系统无法自动解释：相减规则需填写比较指标代码和金额阈值。"
    if action in ("当年累计指标比上期不应减少。", "历史累计指标比上期不应减少。"):
        return "该累计指标较上期减少时，需说明。"
    return f"系统无法自动解释规则动作“{action or '未填写'}”；请核对规则动作、阈值和比较指标代码。"


def generate_special_rule_descriptions(path: Path) -> dict[str, int]:
    """写入“系统生成说明”列，绝不修改用户维护的“备注”列。"""
    from copy import copy

    from openpyxl import load_workbook
    from openpyxl.comments import Comment

    target = Path(path)
    if not target.is_file():
        raise FileNotFoundError(f"执行比较配置不存在：{target}")
    book = load_workbook(target)
    try:
        if SPECIAL_RULE_SHEET not in book.sheetnames:
            raise CentralConfigError(f"配置文件缺少工作表：{SPECIAL_RULE_SHEET}")
        sheet = book[SPECIAL_RULE_SHEET]
        raw = [list(row) for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 12), values_only=True)]
        header_index, positions = _locate_header(raw, tuple(
            item for item in SHEET_HEADERS[SPECIAL_RULE_SHEET] if item != SYSTEM_DESCRIPTION_HEADER
        ), SPECIAL_RULE_SHEET)
        header_row = header_index + 1
        if SYSTEM_DESCRIPTION_HEADER not in positions:
            remark_column = positions["备注"] + 1
            sheet.insert_cols(remark_column)
            source = sheet.cell(header_row, remark_column + 1)
            generated = sheet.cell(header_row, remark_column)
            generated._style = copy(source._style)
            generated.number_format = source.number_format
            generated.alignment = copy(source.alignment)
            generated.fill = copy(source.fill)
            generated.font = copy(source.font)
            generated.border = copy(source.border)
            generated.protection = copy(source.protection)
            generated.value = SYSTEM_DESCRIPTION_HEADER
            generated.comment = Comment("程序按规则动作、阈值和比较指标代码自动生成的说明。重新生成时会覆盖本列，不会修改右侧“备注”列。", "基础数据审核工具")
            old_remark_index = remark_column - 1
            positions = {
                name: column + (1 if column >= old_remark_index else 0)
                for name, column in positions.items()
            }
            positions[SYSTEM_DESCRIPTION_HEADER] = remark_column - 1
        generated_column = positions[SYSTEM_DESCRIPTION_HEADER] + 1
        action_column = positions["规则动作"] + 1
        count = unrecognized = 0
        for row_number in range(header_row + 1, sheet.max_row + 1):
            if not any(sheet.cell(row_number, column + 1).value not in (None, "") for column in positions.values()):
                continue
            row = {
                header: sheet.cell(row_number, column + 1).value
                for header, column in positions.items()
            }
            description = build_special_rule_description(row)
            sheet.cell(row_number, generated_column).value = description
            count += 1
            if description.startswith("系统无法自动解释"):
                unrecognized += 1
        if action_column <= 0:
            raise CentralConfigError("配置表缺少规则动作列")
        book.save(target)
        return {"generated": count, "unrecognized": unrecognized}
    finally:
        book.close()


def _expression_completion_suggestion(token: str) -> str | None:
    """3.1 占位符补全建议：段数不足 8 段时推导缺失空段的位置。

    指标代码的形态是「含数字的字母数字串」（37103 / 12A24 / 20203045），
    必须落在第 4 段（段3）；数据属性（余额）、币种（人民币）、频度（日）
    均不含数字。因此：在原 token 中找唯一的指标样段，按其应在位置在
    **前方补足空段**、尾部补足剩余空段到 8 段——空段=继承当前行，
    语义不变。识别不出唯一指标段、或指标段在第 4 段之后（段序错乱，
    无法安全推导）→ 返回 None，转人工检查。
    """
    import re

    text = token.strip()
    segments = [part.strip() for part in text.strip("[]{}").split(",")]
    if len(segments) >= 8:
        return None
    if len(segments) == 1:
        code = segments[0].replace("'", "")
        if not code or not re.search(r"\d", code):
            return None
        return f"{text[0]},,,{code},,,,{text[-1]}"
    indicator_pattern = re.compile(r"^(?=.*\d)[0-9A-Za-z]{4,}$")
    positions = [index for index, part in enumerate(segments)
                 if indicator_pattern.match(part)]
    if len(positions) != 1:
        return None
    indicator_index = positions[0]
    if indicator_index > 3:
        return None                      # 指标在段 3 之后：段序错乱，不猜
    front = 3 - indicator_index
    tail = 8 - len(segments) - front
    if front < 0 or tail < 0:
        return None
    completed = segments[:indicator_index] + [""] * front + [segments[indicator_index]] + segments[indicator_index + 1:] + [""] * tail
    return f"{text[0]}{','.join(completed)}{text[-1]}"


def _expression_overlong_suggestion(token: str) -> str | None:
    """超长 token（>8 段）建议：删除多余空段，唯一合法结果才给出。

    逐个尝试删除一个空段，保留「结果为 8 段且指标段（第 4 段）非空」的
    候选；候选去重后唯一 → 返回该串（相邻空段删除结果相同，天然去重）；
    多个不同合法结果 → 语义不定，返回 None 转人工。
    """
    text = token.strip()
    segments = [part.strip() for part in text.strip("[]{}").split(",")]
    extra = len(segments) - 8
    if extra <= 0:
        return None
    opener, closer = text[0], text[-1]
    candidates: set[str] = set()
    empty_positions = [index for index, part in enumerate(segments) if not part]
    for position in empty_positions:
        trimmed = segments[:position] + segments[position + 1:]
        if len(trimmed) == 8 and trimmed[3]:
            candidates.add(f"{opener}{','.join(trimmed)}{closer}")
    return candidates.pop() if len(candidates) == 1 else None


# 5 段 token 词表（用户口径，只提示不算错误）：段值不在词表内仅提示人工确认。
FIVE_SEGMENT_VOCABULARY = {
    0: ("指标代码", re.compile(r"^[0-9A-Za-z]{3,}$")),
    1: ("数据属性", ("余额", "发生额")),
    2: ("币种", ("本外币", "人民币", "美元合计", "美元", "人民币合计")),
    3: ("频度", ("日", "月", "季", "半年", "年")),
    4: ("批次", ("1", "2", "3")),
}


def _five_segment_vocabulary_notes(token: str) -> list[str]:
    """5 段 token 词表提示（只提示不算错误）：逐段核对值域，给出人工确认说明。"""
    notes: list[str] = []
    segments = [part.strip() for part in token.strip().strip("[]{}").split(",")]
    for index, (label, allowed) in FIVE_SEGMENT_VOCABULARY.items():
        value = segments[index] if index < len(segments) else ""
        if not value:
            continue          # 空段=继承/当前约定，不算问题
        if isinstance(allowed, tuple):
            if value not in allowed:
                notes.append(f"{token} 第{index + 1}段「{value}」不在常用{label}中"
                             f"（{'、'.join(allowed)}），请人工确认")
        elif not allowed.match(value):
            notes.append(f"{token} 第{index + 1}段「{value}」不是常规{label}"
                         "（字母数字组合），请人工确认")
    return notes


def check_expression_rules(
        config: CentralConfig, *, schema: str = "LEGACY_8",
        stats_out: dict | None = None) -> list[dict]:
    """3.1 表达式逐条体检：按当前语法模式检查对应工作表（只读）。

    LEGACY_8 → 「表达式校验」（8 段）；FIVE_SEGMENT_V1 → 「表达式校验5段式」
    （token 5 段 + 机构/地区列正则）。检查器复用运行时 parser/预编译，
    不另立业务真相。
    """
    import re

    from .complex_rule_engine import (
        ComplexRuleCompileError,
        SCHEMA_FIVE_SEGMENT_V1,
        _clear_rule,
        inspect_complex_rules,
        inspect_five_segment_rules,
    )
    from .expression_parser import ExpressionError, evaluate_expression

    if schema == SCHEMA_FIVE_SEGMENT_V1:
        result = inspect_five_segment_rules(config)
        if stats_out is not None:
            stats_out.update({
                "total": result.total_rows, "enabled": result.enabled_rows,
                "disabled": result.disabled_rows,
                "compiled": result.compiled_rules, "failed": result.failed_rules,
            })
        by_id: dict[str, dict] = {}
        vocabulary_notes: list[str] = []
        for issue in result.issues:
            entry = by_id.setdefault(issue.rule_id, {
                "规则编号": issue.rule_id, "规则说明": "", "启用": "是",
                "问题": [], "建议": [],
            })
            entry["问题"].append(issue.format_user_message())
        for row in config.five_segment_rules:
            rule_id = str(row.get("规则编号") or "").strip()
            if rule_id in by_id:
                by_id[rule_id]["规则说明"] = str(row.get("规则说明") or "").strip()
                by_id[rule_id]["启用"] = "否" if str(row.get("启用") or "").strip() in ("否", "0", "") else "是"
        # 词表提示（只提示不算错误）：逐 token 核对五段值域，经 stats_out 输出；
        # 禁用行不扫描；已进【错误】的 token 不重复进提示。
        problem_tokens = {i.token for i in result.issues if getattr(i, "token", "")}
        for row5 in config.five_segment_rules:
            if str(row5.get("启用") or "").strip() in {"否", "0", "false", "False"}:
                continue
            rule_id5 = str(row5.get("规则编号") or "").strip()
            for token_match in re.finditer(
                    r"\[[^\[\]]*\]|\{[^{}]*\}", str(row5.get("校验表达式") or "")):
                token = token_match.group(0)
                if token in problem_tokens:
                    continue
                vocabulary_notes.extend(
                    f"{rule_id5}：{note}"
                    for note in _five_segment_vocabulary_notes(token))
        if stats_out is not None:
            stats_out["vocabulary"] = vocabulary_notes
        return list(by_id.values())

    issues: list[dict] = []
    try:
        result = inspect_complex_rules(config)
    except ComplexRuleCompileError:
        result = None
    if stats_out is not None and result is not None:
        stats_out.update({
            "total": result.total_rows, "enabled": result.enabled_rows,
            "disabled": result.disabled_rows,
        })
    for row in config.complex_rules:
        rule_id = str(row.get("规则编号") or "").strip()
        expression = _clear_rule(str(row.get("校验规则") or "").strip())
        if not expression:
            continue
        row_issues: list[str] = []
        row_suggestions: list[str] = []
        seen_tokens: set[str] = set()
        for token_match in re.finditer(r"\[[^\[\]]*\]|\{[^{}]*\}", expression):
            token = token_match.group(0)
            if token in seen_tokens:
                continue
            seen_tokens.add(token)
            simple = _expression_completion_suggestion(token) or _expression_overlong_suggestion(token)
            if simple:
                row_suggestions.append(f"{token} → {simple}（空段=继承当前行，语义不变）")
            # 8 段完整（含指定值/^）即为正式形态，不做建议。
        # 中文逗号/全角括号已由 _clear_rule 规范化为半角：指标段混入中文
        # 逗号会把几个值拆成多余段，编译存活检查必然以“段数≠8”报出，
        # 无需重复检测。执行期该 token 查不到数据、按缺失静默取 0。
        trigger = str(row.get("取反标识") or "").strip()
        if trigger not in ("", "0", "1"):
            row_issues.append(f"取反标识应为空或 1，当前：{trigger!r}")
        try:
            evaluate_expression(
                re.sub(r"\[[^\[\]]*\]|\{[^{}]*\}", "1", expression), {"Thd": 1.0})
        except ExcelNaError:
            pass  # IFS 无匹配 #N/A 属业务语义，由撰写规范提示兜底
        except ExpressionError as exc:
            row_issues.append(f"语法/求值：{exc}")
        hint = _ifs_fallback_hint(expression)
        if hint:
            row_issues.append(hint)
        issues.append({
            "规则编号": rule_id, "规则说明": row.get("校验描述") or "",
            "启用": "否" if row.get("禁用") else "是",
            "问题": row_issues, "建议": row_suggestions,
        })
    # 编译失败明细（来自共享 inspect，含完整定位）
    if result is not None:
        by_id: dict[str, list[str]] = {}
        for issue in result.issues:
            by_id.setdefault(issue.rule_id, []).append(issue.format_user_message())
        for item in issues:
            if item["规则编号"] in by_id and item["启用"] == "是":
                item["问题"].append("编译存活检查未通过（详见编译错误条目）")
    return issues


def check_cross_formulas(config: CentralConfig, *, stats_out: dict | None = None) -> list[dict]:
    """3.2 跨期数值核对逐条体检：token 结构/词表/作用域正则/公式语法（只读）。

    检查覆盖全部四列（前提条件/left/right/校验公式）的 token 结构；
    语法探针把 token 替换为哑值只验骨架，token 内的分段/取值问题必须
    由结构检查承担。启用/禁用规则都检查（禁用行同样展示，标注（否））。
    """
    from .expression_parser import ExpressionError, evaluate_expression
    import re

    issues: list[dict] = []
    vocabulary_notes: list[str] = []
    enabled_count = 0
    for row in config.cross_rules:
        rule_id = row.get("校验编码") or row.get("校验名称") or ""
        disabled = str(row.get("禁用") or "").strip() in {"是", "1", "true", "True", "Y", "y"}
        if not disabled:
            enabled_count += 1
        row_issues: list[str] = []
        columns = {
            "前提条件": row.get("前提条件"), "left": row.get("left"),
            "right": row.get("right"), "校验公式": row.get("校验公式"),
        }
        # 作用域列：机构/地区正则必须可编译；左中/右中是 3.1 兼容转写，3.2 应直接写标准正则。
        for column, label in (("机构类代码", "机构类"), ("地区代码", "地区")):
            pattern = str(row.get(column) or "").strip()
            if not pattern:
                continue
            if "左中" in pattern or "右中" in pattern:
                row_issues.append(
                    f"{label}代码列包含「左中/右中」转写：{pattern}；"
                    "3.2 应直接使用标准正则表达式")
                continue
            try:
                re.compile(pattern)
            except re.error as exc:
                row_issues.append(f"{label}代码列正则无法编译：{exc}")

        bad_tokens: set[str] = set()      # 已进【错误】的 token 不再进词表提示
        for name in ("前提条件", "left", "right", "校验公式"):
            body = str(columns[name] or "")
            for match in re.finditer(r"\[[^\[\]]*\]|\{[^{}]*\}", body):
                token = match.group(0)
                segments = [part.strip() for part in token[1:-1].split(",")]
                # 段数硬校验前置：5 段式必须有 4 个逗号（恰好 5 段）。
                # 段数不对大概率取不到数据，直接进【错误】。
                if len(segments) != 5:
                    row_issues.append(
                        f"{name} {token} 段数不对：5 段式必须有 4 个逗号，"
                        f"当前 {len(segments) - 1} 个逗号（{len(segments)} 段），"
                        "大概率取不到数据，请人工检查")
                    bad_tokens.add(token)
                    continue
                if len(segments) == 5:
                    # 五段：逐项报出全部缺陷（不隐藏第二个缺陷）。
                    if not segments[0].strip():
                        row_issues.append(f"{name} token 缺指标代码：{token}")
                    if not segments[1].strip():
                        row_issues.append(f"{name} token 缺数据属性：{token}")
                    if not segments[2].strip():
                        row_issues.append(f"{name} token 缺币种：{token}")
                    if not segments[3].strip() and not segments[4].strip():
                        # 频度+批次为空 → fres 校验不含空键，规则执行时整条静默跳过。
                        row_issues.append(
                            f"{name} token 频度/批次为空：{token}（执行时该规则会因"
                            "匹配不到频度被跳过，请补全 频度,批次）")
                    else:
                        if not segments[3].strip():
                            row_issues.append(f"{name} token 缺频度：{token}")
                        if not segments[4].strip():
                            row_issues.append(f"{name} token 缺批次：{token}")
                # 词表提示（只提示不算错误）；禁用行与已进【错误】的 token 不重复。
                if not disabled and token not in bad_tokens:
                    vocabulary_notes.extend(
                        f"{rule_id}：{note}" for note in _five_segment_vocabulary_notes(token))
        for name in ("前提条件", "left", "right", "校验公式"):
            body = str(columns[name] or "").strip()
            if not body:
                continue
            probe = re.sub(r"\[[^\[\]]*\]|\{[^{}]*\}", "1", body)
            try:
                evaluate_expression(probe, {"left": 1.0, "right": 1.0, "Thd": 1.0})
            except ExpressionError as exc:
                row_issues.append(f"{name} 语法/求值：{exc}")
        issues.append({
            "规则编号": rule_id, "规则说明": row.get("校验名称") or "",
            "启用": "否" if disabled else "是",
            "问题": row_issues, "建议": [],
        })
    if stats_out is not None:
        stats_out.update({
            "total": len(config.cross_rules),
            "enabled": enabled_count,
            "disabled": len(config.cross_rules) - enabled_count,
            "vocabulary": vocabulary_notes,
        })
    return issues


def write_default_central_config(path: Path) -> Path:
    """生成默认配置工作簿（表头+默认参数/警戒区间+示例行）。"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    from ...config_guide import write_guide_sheet

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    try:
        first = True
        for sheet_name, headers in SHEET_HEADERS.items():
            sheet = book.active if first else book.create_sheet(sheet_name)
            first = False
            sheet.title = sheet_name
            sheet.append(list(headers))
        for name, value in DEFAULT_RUN_PARAMS.items():
            book[RUN_SHEET].append((name, value, ""))
        for row in DEFAULT_ALERTS:
            book[ALERT_SHEET].append(row)
        book[UNIT_EXCEPTION_SHEET].append((
            "U001", "A0000", "99999999", "示例指标（请替换）", "否", "示例行",
        ))
        book[ACTION_RULE_SHEET].append((
            "AC001", "示例累计指标（请替换）", "当年累计指标比上期不应减少。",
            "99999999", "发生额", "", "", "", "", "",
            "当年累计", "A0000", "月", "", "",
            "", "否", "", "示例行",
        ))
        book[ACTION_RULE_SHEET].append((
            "SP001", "示例特殊指标（请替换）", "指标不应有数，需说明。",
            "99999998", "余额", "人民币", "", "", "", "",
            "自定义", "A0000", "", "", "",
            "", "否", "", "示例行",
        ))
        book[EXPRESSION_RULE_SHEET].append((
            "C001", "自定义", "A0000", "示例：单期校验（请替换）",
            "[人民币,,^44,,,,,,]<>0", "表达式成立", "", "否", "示例行", "",
        ))
        book[CROSS_SHEET].append((
            "K001", "示例：信贷资产占比", "", "", "", "[33370,余额,人民币,月,2]",
            "{33370,余额,人民币,月,2}", "0.8 * left > right", "信贷资产占比小于80%", "", "", "", "", "示例行",
        ))
        book[TOFORM_SHEET].append(("隐藏空行", "是", "数据列全空的行整行隐藏"))
        book[TOFORM_SHEET].append(("空表删除", "是", "无任何数据的报表删除而非隐藏"))
        book[TOFORM_SHEET].append(("分析文件数据单位", "亿元", "比较结果数值单位"))
        write_guide_sheet(book, "central", index=0)
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="1F4E78")
        for sheet in book.worksheets:
            sheet.freeze_panes = "A2"
            if sheet.title != "使用说明":
                for cell in sheet[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                for column in sheet.columns:
                    width = max(10, min(44, max(len(str(c.value or "")) for c in column) + 2))
                    sheet.column_dimensions[column[0].column_letter].width = width
        book.save(target)
    except PermissionError as exc:
        raise RuntimeError(f"无法写入大集中统计系统配置：请先关闭“{target.name}”后重试") from exc
    finally:
        book.close()
    return target


# 用户 WIP 过渡别名：语义与 write_default_central_config 完全一致。
write_default_central_profile = write_default_central_config


def ensure_split_central_configs(config_dir: Path) -> dict[str, Path]:
    """确保 3.0/3.1/3.2/3.3 当前配置存在。

    不再读取或迁移旧 ``大集中统计系统_配置.xlsx``。缺册时仅用程序内置
    默认结构生成当前分册，正式发行包应始终直接携带四册配置。
    """
    from openpyxl import load_workbook

    from ...config_guide import write_guide_sheet

    config_dir = Path(config_dir)
    config_dir.mkdir(parents=True, exist_ok=True)
    outputs = {profile: config_dir / file_name for profile, (file_name, _) in CONFIG_PROFILES.items()}
    missing = [profile for profile, target in outputs.items() if not target.is_file()]
    if not missing:
        return outputs

    temporary_source = config_dir / ".大集中当前默认结构临时.xlsx"
    write_default_central_config(temporary_source)
    guide_keys = {
        "common": "central_common",
        "comparison": "central_comparison",
        "cross": "central_cross",
        "forms": "central_forms",
    }
    try:
        for profile in missing:
            target = outputs[profile]
            _, sheets = CONFIG_PROFILES[profile]
            book = load_workbook(temporary_source)
            try:
                keep = set(sheets)
                for sheet_name in list(book.sheetnames):
                    if sheet_name != "使用说明" and sheet_name not in keep:
                        del book[sheet_name]
                write_guide_sheet(book, guide_keys[profile], index=0)
                book.save(target)
            finally:
                book.close()
    finally:
        if temporary_source.is_file():
            temporary_source.unlink()
    return outputs
