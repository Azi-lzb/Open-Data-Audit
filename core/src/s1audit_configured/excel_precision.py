"""Excel numeric precision semantics shared by formula evaluators."""

from __future__ import annotations

import math
from decimal import Decimal
from typing import Any


def excel_round15(value: Any) -> Any:
    """Converge numeric values to Excel's 15 significant decimal digits.

    Excel formula results are stored at 15 significant digits. Keeping this
    normalization in one helper lets the audit DSL and conditional-format
    evaluator agree on comparisons and intermediate arithmetic results.
    Non-numeric values, booleans, zero, and non-finite values pass through.
    """
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        return value
    number = float(value)
    if number == 0.0 or not math.isfinite(number):
        return value
    return float("%.15g" % number)
