"""大集中统计系统：执行比较引擎（两期配对、环比、警戒区间）。

对应 VBA ``DoCompare`` 基础管线（B导数比较.bas 385-466、1545-1577）。
单位换算在 CSV 导入时完成（豁免清单指标除外）；本引擎全部使用目标单位值。
"""

from __future__ import annotations

from .config import CentralConfig, alert_band_for, format_alert_band_remark
from .models import (
    CentralDataset,
    CentralRecord,
    ComparisonRow,
)

# VBA dbl_Epsilon（B导数比较.bas 42）。
DBL_EPSILON = 2.22044604925031e-13

REMARK_CUR_ONLY = "本期有，上期无"
REMARK_PRE_ONLY = "本期无，上期有"
REMARK_DIV_ZERO = "除数为0"
REMARK_SHRINK_100 = "缩小100倍以上"
EXPLANATION_JOIN = "    ||    "

COLOR_RED = 3     # 本期有上期无（VBA colorindex）
COLOR_YELLOW = 6  # 本期无上期有
COLOR_BELOW_BANDS = 10  # 降幅超出全部警戒档


def excel_round(value: float, digits: int) -> float:
    """WorksheetFunction.Round：四舍五入（.5 远离零），与 Python round 的银行家舍入不同。"""
    scaled = value * (10 ** digits)
    floor_part = float(int(scaled))
    fraction = scaled - floor_part
    if scaled >= 0:
        rounded = floor_part + (1.0 if fraction >= 0.5 else 0.0)
    else:
        # 负数按 |x| 四舍五入后取号（1.5 → -2）
        rounded = floor_part - (1.0 if -fraction >= 0.5 else 0.0)
    return rounded / (10 ** digits)


def vba_round(value: float, digits: int) -> float:
    """VBA Round：银行家舍入。"""
    factor = 10 ** digits
    scaled = value * factor
    floor_part = float(int(scaled))
    fraction = scaled - floor_part
    if fraction > 0.5:
        result = floor_part + 1.0
    elif fraction < -0.5:
        result = floor_part - 1.0
    elif fraction == 0.5:
        result = floor_part + 1.0 if floor_part % 2 == 0 else floor_part
    elif fraction == -0.5:
        result = floor_part if floor_part % 2 == 0 else floor_part - 1.0
    else:
        result = floor_part
    return result / factor


def get_round_digits(target_unit: str) -> int:
    """VBA GetRoundNum（277-303 行）：单位 -> 增减额保留小数位。"""
    return {
        "元": 2, "十元": 3, "百元": 4, "千元": 5, "万元": 6,
        "十万元": 7, "百万元": 8, "千万": 9, "亿元": 10,
    }.get(target_unit, 2)


def get_unit_factor(unit: str) -> float:
    """VBA GetUnitVal（通用.bas 442-468）。"""
    return {
        "元": 1.0, "十元": 10.0, "百元": 100.0, "千元": 1000.0, "万元": 1e4,
        "十万元": 1e5, "百万元": 1e6, "千万": 1e7, "千万元": 1e7, "亿元": 1e8,
    }.get(unit, 1.0)


def _band_remark(ratio: float, alerts: list[dict]) -> tuple[str, int, int, int]:
    """警戒区间命中：返回 (备注, 颜色, 是否整行填充, 命中)。

    使用配置的 ``下限 <= 环比 < 上限`` 区间；空上限表示无穷大。
    全不满足 → 「缩小100倍以上」整行填色 10。
    """
    if vba_round(ratio, 6) == 0:
        return "", 0, 0, 0
    band = alert_band_for(ratio, alerts)
    if band is not None:
        try:
            color = int(band.get("填充颜色") or 0)
        except (TypeError, ValueError):
            color = 0
        try:
            full_row = int(band.get("是否整行填充") or 0)
        except (TypeError, ValueError):
            full_row = 0
        return format_alert_band_remark(band), color, full_row, 1
    return REMARK_SHRINK_100, COLOR_BELOW_BANDS, 1, 1


def _band_explanation(change: float, band: dict | None, target_unit: str) -> str:
    """按警戒档的金额门槛生成“是否说明”；空配置不触发。"""
    if band is None:
        return ""
    try:
        threshold = float(band.get("金额变动阈值") or "")
    except (TypeError, ValueError):
        return ""
    configured = str(band.get("是否说明") or "").strip()
    if not configured or configured in {"否", "0", "false", "False"}:
        return ""
    if abs(change) <= abs(threshold):
        return ""
    if configured in {"是", "1", "true", "True"}:
        configured = format_alert_band_remark(band)
    threshold_text = str(int(threshold)) if threshold.is_integer() else str(threshold)
    return configured.replace("{金额变动阈值}", threshold_text) \
        .replace("{目标单位}", target_unit) \
        .replace("{备注文字}", format_alert_band_remark(band))


def _append_explanation(existing: str, added: str) -> str:
    """合并说明并保留全部来源；支持复合文本，相同内容只保留一次。"""
    parts = [part.strip() for part in str(existing or "").split("||") if part.strip()]
    for added_part in str(added or "").split("||"):
        added_part = added_part.strip()
        if added_part and added_part not in parts:
            parts.append(added_part)
    return EXPLANATION_JOIN.join(parts)


def build_order_replacements(config: CentralConfig) -> dict[str, tuple[str, str]]:
    """指标参照：指标代码 -> (表单_行序号 顺序码, 指标别名)。"""
    replacements: dict[str, tuple[str, str]] = {}
    for row in config.indicators:
        code = str(row.get("指标代码") or "").replace("'", "").strip()
        if not code:
            continue
        form = str(row.get("表单代码") or "").strip()
        try:
            line_no = int(float(str(row.get("行序号") or "0")))
        except (TypeError, ValueError):
            line_no = 0
        order_code = f"{form}_{line_no:04d}"
        alias = str(row.get("指标别名") or "").strip()
        replacements[code] = (order_code, alias)
    return replacements


def build_comparison(
    current: CentralDataset,
    previous: CentralDataset,
    config: CentralConfig,
    *,
    target_unit: str = "亿元",
) -> list[ComparisonRow]:
    """两期比较主管线；行序为文件序，排序由导出层按 VBA 键执行。"""
    alerts = [row for row in config.alerts]
    order_map = build_order_replacements(config)
    digits = get_round_digits(target_unit)
    rows: list[ComparisonRow] = []
    seen_keys: set[str] = set()

    for record in current.records:
        key = record.key()
        seen_keys.add(key)
        prev_record = previous.key_index.get(key)
        row = ComparisonRow(record=record)
        if prev_record is not None:
            row.prev_value = prev_record.value
            cur_number = record.value if isinstance(record.value, (int, float)) else 0.0
            prev_number = prev_record.value if isinstance(prev_record.value, (int, float)) else 0.0
            if abs(cur_number - prev_number) > DBL_EPSILON:
                row.change = excel_round(cur_number - prev_number, digits)
                if prev_number != 0:
                    row.ratio = row.change / prev_number
                    remark, color, full_row, _hit = _band_remark(float(row.ratio), alerts)
                    row.remark = remark
                    band = alert_band_for(float(row.ratio), alerts)
                    row.need_explain = _append_explanation(
                        row.need_explain,
                        _band_explanation(float(row.change), band, target_unit),
                    )
                    if color:
                        row.fill_color = color
                        row.fill_scope = "row" if full_row else "remark"
                else:
                    row.ratio = REMARK_DIV_ZERO
        else:
            # 本期有上期无：增减额=当期原值（未舍入），整行(1-18)填红。
            row.prev_value = None
            row.change = record.value if isinstance(record.value, (int, float)) else None
            row.ratio = None
            row.remark = REMARK_CUR_ONLY
            row.fill_color = COLOR_RED
            row.fill_scope = "row"
        _apply_output_names(row, order_map)
        rows.append(row)

    for prev_record in previous.records:
        key = prev_record.key()
        if key in seen_keys:
            continue
        seen_keys.add(key)
        # 本期无上期有：追加行（数据值空、增减=-上期、环比=-1、填黄）。
        ghost = CentralRecord(
            biz_class=prev_record.biz_class,
            record_date=current.record_date or prev_record.record_date,
            org_code=prev_record.org_code, org_name=prev_record.org_name,
            region_code=prev_record.region_code, region_name=prev_record.region_name,
            order_code=prev_record.order_code, indicator=prev_record.indicator,
            indicator_name=prev_record.indicator_name, data_attr=prev_record.data_attr,
            currency=prev_record.currency, frequency=prev_record.frequency,
            batch=prev_record.batch, value=None,
        )
        row = ComparisonRow(record=ghost)
        prev_value = prev_record.value if isinstance(prev_record.value, (int, float)) else 0.0
        row.prev_value = prev_value
        row.change = -prev_value
        row.ratio = -1
        row.remark = REMARK_PRE_ONLY
        row.fill_color = COLOR_YELLOW
        row.fill_scope = "row"
        _apply_output_names(row, order_map)
        rows.append(row)

    return rows


def _apply_output_names(row: ComparisonRow, order_map: dict[str, tuple[str, str]]) -> None:
    """VBA 477-486：顺序码/名称替换，无命中时顺序码回退为指标代码。"""
    lookup = order_map.get(row.record.indicator)
    if lookup is not None:
        row.order_code_out = lookup[0]
        if lookup[1]:
            row.indicator_name_out = lookup[1]
    else:
        row.order_code_out = row.record.indicator


def sort_comparison_rows(rows: list[ComparisonRow]) -> list[ComparisonRow]:
    """VBA EndCompare 排序：Key1=C 机构类代码、Key2=E 地区代码、Key3=G 顺序码。"""
    return sorted(
        rows,
        key=lambda row: (
            row.record.org_code,
            row.record.region_code,
            row.order_code_out or row.record.indicator,
        ),
    )
