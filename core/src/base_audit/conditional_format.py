"""OOXML 条件格式规则统一求值器。

本模块是全部条件格式规则求值的唯一入口（引擎层不得另写解析器副本）：

- ``evaluate_expression_formula``：``expression`` 规则的公式解析器；
- ``evaluate_cellis_rule``：``cellIs`` 规则的统一比较求值；
- ``evaluate_expression_condition``：把公式求值包装为四态结果。

四态语义（引擎层必须区分，禁止把「不支持」当成「未触发」）：
  TRUE         规则成立
  FALSE        规则不成立
  UNSUPPORTED  规则/操作数无法由本求值器解析（需真实办公软件或人工核实）
  ERROR        求值过程本身出错

Supported subset (matches every shape observed in the real templates):
  AND / OR / NOT / ABS / SUM, comparisons = <> < > <= >=, arithmetic + - * /,
  numeric & string literals incl. "", single cells and A1:B2 ranges, with
  $ -locking.  Values are read from the audit copy's already-calculated
  (data_only) cache, so formula cells referenced by a rule already carry the
  computed result.
"""

from __future__ import annotations

import math
import re
import operator as _operator
from dataclasses import dataclass


# ---------------------------------------------------------------------------
# Excel 浮点语义：15 位有效数字舍入
# ---------------------------------------------------------------------------


def _excel_round15(value):
    """按 Excel 公式引擎语义把数值舍入到 15 位有效十进制数字。

    Excel 每次浮点运算后都会把结果收敛到 15 位有效数字，因此
    ``=0.1+0.2=0.3`` 在 Excel 中为 TRUE（标准 IEEE 754 比较为 FALSE）。
    「合计 = 子项之和」类条件格式规则依赖该语义：真机实测
    ``$D35<>$D27+$D28+$D34`` 两侧数学相等（31925.83 vs 31925.83），
    但双精度表示相差 3.6e-12；Excel 判「相等 → 不触发」，本求值器若用
    精确比较则误报（测试中 D35 / E35 的浮点比较问题）。

    仅吸收 1e-12 级尾差：真正不等的值（如差 0.01）不受影响，不会漏报。
    非数值与 NaN/Inf 原样返回（由调用方按既有语义处理）。
    """
    if isinstance(value, bool):
        return value
    if not isinstance(value, (int, float)):
        return value
    number = float(value)
    if number == 0.0 or not math.isfinite(number):
        return value
    return float("%.15g" % number)


class UnsupportedFormulaError(ValueError):
    """条件格式公式使用了本求值器不支持的表达（函数/操作数形态）。

    属于 ValueError 子类：既有 ``except ValueError`` 的调用方行为不变，
    但引擎层可单独捕获它，把规则记为 UNSUPPORTED 而不是未触发。
    """


# 四态求值结果状态（见模块 docstring）。
STATUS_TRUE = "TRUE"
STATUS_FALSE = "FALSE"
STATUS_UNSUPPORTED = "UNSUPPORTED"
STATUS_ERROR = "ERROR"


@dataclass
class RuleEvaluation:
    """单条规则在单个单元格上的求值结果。"""

    status: str
    reason: str = ""

    @property
    def triggered(self) -> bool:
        return self.status == STATUS_TRUE

    @property
    def supported(self) -> bool:
        return self.status != STATUS_UNSUPPORTED


def rule_true() -> RuleEvaluation:
    return RuleEvaluation(STATUS_TRUE)


def rule_false() -> RuleEvaluation:
    return RuleEvaluation(STATUS_FALSE)


def rule_unsupported(reason: str) -> RuleEvaluation:
    return RuleEvaluation(STATUS_UNSUPPORTED, reason=reason)


def rule_error(reason: str) -> RuleEvaluation:
    return RuleEvaluation(STATUS_ERROR, reason=reason)


def evaluate_expression_condition(formula: str, ws, anchor_row: int, anchor_col: int,
                                  cur_row: int, cur_col: int) -> RuleEvaluation:
    """把 ``expression`` 公式求值包装为四态结果。"""
    try:
        result = evaluate_expression_formula(formula, ws, anchor_row, anchor_col, cur_row, cur_col)
    except UnsupportedFormulaError as exc:
        return rule_unsupported(str(exc))
    except ValueError as exc:
        return rule_error(str(exc))
    return rule_true() if _expr_truthy(result) else rule_false()


# 公式里的单元格/区域引用（与 _EXPR_TOKEN_RE 的 ref 形态一致）。
# 两侧都要设边界：前邻不能是字母/数字/$（避免切进标识符中部），后邻不能
# 是字母/数字（避免把 LOG10 的 "1" 当行号），也不能直接跟 "("（函数名）。
_FORMULA_REF_RE = re.compile(
    r'"(?:[^"])*"'
    r"|(?<![A-Za-z0-9_$])(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)"
    r"(?!\w)(?!\s*\()"
)


def _col_to_letters(col: int) -> str:
    """1 基列号 → 列字母（26 进制，A=1）。"""
    letters = ""
    number = col
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _shift_ref_text(ref: str, dr: int, dc: int) -> str:
    """单个引用按 (dr, dc) 平移；$ 锁定维不动；越界（行/列 <1）记 #REF!。

    注意列号必须转换回字母（早期版本只改了行号、列平移时原样返回，
    导致 ``B2`` 平移后仍是 ``B2``）。
    """
    col_lock = "" if ref.startswith("$") else None
    body = ref.lstrip("$")
    match = re.match(r"([A-Za-z]+)(\$?)(\d+)$", body)
    if not match:
        return ref
    letters, row_lock, digits = match.groups()
    col = 0
    for ch in letters.upper():
        col = col * 26 + (ord(ch) - 64)
    row = int(digits)
    if col_lock is None:
        col += dc
    if not row_lock:
        row += dr
    if row < 1 or col < 1:
        return "#REF!"
    prefix = "$" if col_lock is not None else ""
    return "{}{}{}{}".format(prefix, _col_to_letters(col), row_lock, row)


def translate_formula_refs(formula: str, dr: int, dc: int) -> str:
    """把公式文本里的相对引用按 (dr, dc) 平移，返回用于展示的公式。

    坐标语义与求值器 ``_ExprRefs`` 完全一致（``$J8`` 列锁行相对；字符串
    字面量原样保留）——用于把规则描述「改到对应的单元格」，即 Office 对
    单格公式渲染的口径。无法识别的片段原样保留。
    """
    if not dr and not dc:
        return formula

    def _sub(match: re.Match) -> str:
        ref = match.group(1)
        if ref is None:                     # 字符串字面量
            return match.group(0)
        if ":" in ref:
            left, right = ref.split(":")
            return "{}:{}".format(_shift_ref_text(left, dr, dc),
                                  _shift_ref_text(right, dr, dc))
        return _shift_ref_text(ref, dr, dc)

    return _FORMULA_REF_RE.sub(_sub, formula)


_EXPR_TOKEN_RE = re.compile(r"""\s*(?:
   (?P<num>\d+(?:\.\d+)?) |
   (?P<str>"(?:[^"])*") |
   (?P<ref>\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?) |
   (?P<name>[A-Za-z_][A-Za-z0-9_.]*) |
   (?P<op><>|<=|>=|[<>+\-*/=()%,])
 )""", re.VERBOSE)

_EXPR_ARITH = {'+': _operator.add, '-': _operator.sub, '*': _operator.mul, '/': _operator.truediv}


def _expr_number(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        try:
            if text.endswith('%'):
                return float(text[:-1]) / 100.0
            return float(text)
        except ValueError:
            return None
    return None


def _expr_truthy(value) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, str):
        return value != ""
    return value is not None


class _ExprRefs:
    """Resolve cell / range references inside an expression rule."""

    def __init__(self, ws, anchor_row: int, anchor_col: int):
        self.ws = ws
        self.ar = anchor_row
        self.ac = anchor_col

    @staticmethod
    def _coord(ref: str):
        col_abs = ref.startswith("$")
        rest = ref.lstrip("$")
        match = re.match(r"([A-Za-z]+)\$?(\d+)$", rest)
        col_letters, row_text = match.group(1), match.group(2)
        col = 0
        for ch in col_letters.upper():
            col = col * 26 + (ord(ch) - 64)
        row = int(row_text)
        row_abs = ref.count("$") >= 2
        return col, row, col_abs, row_abs

    def cell_value(self, ref, dr, dc):
        col, row, col_abs, row_abs = self._coord(ref)
        col = col if col_abs else col + dc
        row = row if row_abs else row + dr
        return self.ws.cell(row=row, column=col).value

    def range_values(self, ref, dr, dc):
        a, b = ref.split(":")
        c1, r1, c1a, r1a = self._coord(a)
        c2, r2, c2a, r2a = self._coord(b)
        if not c1a:
            c1 += dc
        if not c2a:
            c2 += dc
        if not r1a:
            r1 += dr
        if not r2a:
            r2 += dr
        lo_c, hi_c = sorted((c1, c2))
        lo_r, hi_r = sorted((r1, r2))
        return [
            self.ws.cell(row=row, column=col).value
            for row in range(lo_r, hi_r + 1)
            for col in range(lo_c, hi_c + 1)
        ]


def _expr_call(name, args):
    if name == "AND":
        return all(_expr_truthy(a) for a in args)
    if name == "OR":
        return any(_expr_truthy(a) for a in args)
    if name == "NOT":
        return not _expr_truthy(args[0]) if args else False
    if name == "ABS":
        v = _expr_number(args[0])
        return abs(v) if v is not None else None
    if name == "SUM":
        total = 0.0
        for arg in args:
            if isinstance(arg, list):
                for item in arg:
                    n = _expr_number(item)
                    if n is not None:
                        total += n
            else:
                n = _expr_number(arg)
                if n is not None:
                    total += n
        return _excel_round15(total)
    raise UnsupportedFormulaError("暂不支持的条件格式函数：{}".format(name))


def evaluate_expression_formula(formula: str, ws, anchor_row: int, anchor_col: int,
                                cur_row: int, cur_col: int):
    """Evaluate an Excel CF ``expression`` at one cell (anchor-translated).

    ``ws`` may be any object exposing ``cell(row=…, column=…).value`` (an
    openpyxl worksheet).  Relative references are shifted by ``cur`` - ``anchor``
    while ``$``-locked references stay fixed, so a rule anchored at C5 and
    applied to C26 reads C26 rather than C5.
    """
    dr = cur_row - anchor_row
    dc = cur_col - anchor_col
    refs = _ExprRefs(ws, anchor_row, anchor_col)
    tokens = list(_EXPR_TOKEN_RE.finditer(formula))
    pos = [0]

    def peek():
        return tokens[pos[0]].group() if pos[0] < len(tokens) else None

    def peek_kind():
        return tokens[pos[0]].lastgroup if pos[0] < len(tokens) else None

    def eat():
        token = tokens[pos[0]]
        pos[0] += 1
        return token.group()

    def fail():
        raise ValueError("无法解析的条件格式公式：{}".format(formula))

    def expr():
        left = term()
        while peek_kind() == "op" and peek() in ("<>", "<=", ">=", "<", ">", "="):
            op = eat()
            right = term()
            left = _expr_compare(op, left, right)
        return left

    def term():
        left = factor()
        while peek_kind() == "op" and peek() in ("+", "-"):
            op = eat()
            right = factor()
            if isinstance(left, str) or isinstance(right, str):
                if op == "+":
                    left = str(left or "") + str(right or "")
                else:
                    left = None
            else:
                # Excel 对每次运算结果做 15 位有效数字收敛（见 _excel_round15）。
                left = _excel_round15(
                    _EXPR_ARITH[op](_expr_number(left) or 0.0, _expr_number(right) or 0.0))
        return left

    def factor():
        left = unary()
        while peek_kind() == "op" and peek() in ("*", "/"):
            op = eat()
            right = unary()
            ln, rn = _expr_number(left), _expr_number(right)
            if ln is None or rn is None:
                left = None
            elif op == "/" and rn == 0:
                left = None
            else:
                left = _excel_round15(_EXPR_ARITH[op](ln, rn))
        return left

    def unary():
        if peek_kind() == "op" and peek() == "-":
            eat()
            return -(_expr_number(unary()) or 0.0)
        if peek() == "(":
            eat()
            value = expr()
            if peek() != ")":
                fail()
            eat()
            return value
        kind = peek_kind()
        if kind == "num":
            raw = eat()
            return float(raw) if "." in raw else int(raw)
        if kind == "str":
            return eat()[1:-1]
        if kind == "ref":
            ref = eat()
            if ":" in ref:
                return refs.range_values(ref, dr, dc)
            return refs.cell_value(ref, dr, dc)
        if kind == "name":
            name = eat().upper()
            if name == "TRUE":
                return True
            if name == "FALSE":
                return False
            if peek() != "(":
                fail()
            eat()
            args = []
            if peek() != ")":
                while True:
                    args.append(expr())
                    if peek() == ",":
                        eat()
                    else:
                        break
            if peek() != ")":
                fail()
            eat()
            return _expr_call(name, args)
        fail()

    def _expr_compare(op, a, b):
        a_empty = a is None or a == ""
        b_empty = b is None or b == ""
        # 数值比较前按 Excel 15 位有效数字收敛（两侧同为数值时）；
        # 字符串/布尔比较不受影响。
        if not a_empty and not b_empty and not isinstance(a, bool) and not isinstance(b, bool):
            na15, nb15 = _expr_number(a), _expr_number(b)
            if na15 is not None and nb15 is not None:
                a, b = _excel_round15(na15), _excel_round15(nb15)
        if op == "=":
            if a_empty and b_empty:
                return True
            if a_empty or b_empty:
                return False
            return a == b
        if op == "<>":
            if a_empty and b_empty:
                return False
            if a_empty or b_empty:
                return True
            return a != b
        na = _expr_number(0 if a_empty else a)
        nb = _expr_number(0 if b_empty else b)
        if na is None or nb is None:
            return False
        return {"<": na < nb, ">": na > nb, "<=": na <= nb, ">=": na >= nb}[op]

    value = expr()
    return _expr_truthy(value)


_CELLIS_OPERATORS = (
    "equal", "notequal", "greaterthan", "greaterthanorequal",
    "lessthan", "lessthanorequal", "between", "notbetween",
)


def evaluate_cellis_rule(operator, formulas, value, resolve_operand) -> RuleEvaluation:
    """``cellIs`` 规则统一求值（OOXML 引擎与 Windows WPS 兜底共用）。

    ``resolve_operand(formulas, index)`` 解析第 index 个操作数（数值/文本/
    布尔字面量或单格引用）；解析不了时必须抛 :class:`UnsupportedFormulaError`，
    本函数将其转为 UNSUPPORTED——操作数是公式表达式时无法离线判定，
    当作「未触发」会掩盖真问题。

    类型语义与 Excel 一致：数值比数值、文本比文本、布尔比布尔；
    数值与文本混型视为不匹配（FALSE），不是错误。
    """
    op = str(operator or "").strip().casefold()
    if op not in _CELLIS_OPERATORS:
        return rule_unsupported("未知 cellIs 运算符：{}".format(operator or "空"))
    if value is None or value == "":
        # Excel 对空单元格不触发 cellIs 比较。
        return rule_false()
    try:
        first = resolve_operand(formulas, 0)
        second = resolve_operand(formulas, 1) if op in ("between", "notbetween") else None
    except UnsupportedFormulaError as exc:
        return rule_unsupported(str(exc))
    if first is None:
        return rule_unsupported("cellIs 操作数无法解析为字面量或单格引用")
    if op in ("between", "notbetween") and second is None:
        return rule_unsupported("cellIs {} 缺少第二操作数".format(op))

    def _compare(a, b):
        return {
            "equal": a == b, "notequal": a != b,
            "greaterthan": a > b, "greaterthanorequal": a >= b,
            "lessthan": a < b, "lessthanorequal": a <= b,
        }[op]

    value_num, first_num, second_num = _expr_number(value), _expr_number(first), _expr_number(second)
    if value_num is not None and first_num is not None:
        # 数值比较按 Excel 15 位有效数字收敛（与表达式求值器的比较同源）。
        value_num = _excel_round15(value_num)
        first_num = _excel_round15(first_num)
        if second_num is not None:
            second_num = _excel_round15(second_num)
        if op in ("between", "notbetween"):
            if second_num is None:
                return rule_unsupported("cellIs between 第二操作数不是数值")
            inside = first_num <= value_num <= second_num
            return rule_true() if (inside if op == "between" else not inside) else rule_false()
        return rule_true() if _compare(value_num, first_num) else rule_false()
    if isinstance(value, str) and isinstance(first, str):
        if op in ("between", "notbetween"):
            return rule_unsupported("cellIs between 暂不支持文本比较")
        return rule_true() if _compare(value, first) else rule_false()
    if isinstance(value, bool) and isinstance(first, bool):
        return rule_true() if _compare(value, first) else rule_false()
    # 数值/文本/布尔混型：Excel 比较为 FALSE。
    return rule_false()
