"""S2 金额单位换算；余额和累发换算到统一输出单位，保留 Decimal 精度。"""

from __future__ import annotations

from decimal import Decimal, InvalidOperation
from typing import Any


AMOUNT_DATA_TYPES = frozenset({"余额", "累发"})
UNIT_TO_YUAN = {
    "元": Decimal("1"),
    "万元": Decimal("10000"),
    "亿元": Decimal("100000000"),
}
_UNIT_EXPONENT = {"元": 0, "万元": 4, "亿元": 8}


def as_decimal(value: Any) -> Decimal:
    """以十进制文本解释源数值，避免先转二进制浮点再换算。"""
    if isinstance(value, Decimal):
        result = value
    else:
        try:
            result = Decimal(str(value).strip().replace(",", ""))
        except (InvalidOperation, TypeError, ValueError) as exc:
            raise ValueError(f"金额不是有效数值：{value!r}") from exc
    if not result.is_finite():
        raise ValueError(f"金额不是有限数值：{value!r}")
    return result


def convert_amount(value: Any, source_unit: str, target_unit: str) -> Decimal:
    """按精确十进制比例转换单位，不量化、不四舍五入。"""
    source = str(source_unit or "").strip()
    target = str(target_unit or "").strip()
    if source not in UNIT_TO_YUAN:
        raise ValueError(f"不支持的金额单位：{source or '空白'}（支持：元、万元、亿元）")
    if target not in UNIT_TO_YUAN:
        raise ValueError(f"不支持的金额单位：{target or '空白'}（支持：元、万元、亿元）")
    number = as_decimal(value)
    # 元/万元/亿元之间的换算倍率始终是 10 的整数次幂。直接调整 Decimal
    # 的 exponent，避免乘除运算受 Decimal 当前 context 精度限制而舍入。
    shift = _UNIT_EXPONENT[source] - _UNIT_EXPONENT[target]
    sign, digits, exponent = number.as_tuple()
    return Decimal((sign, digits, exponent + shift))


def to_yuan(value: Any, unit: str) -> Decimal:
    """兼容旧内部调用名；新逻辑应显式指定目标单位。"""
    return convert_amount(value, unit, "元")
