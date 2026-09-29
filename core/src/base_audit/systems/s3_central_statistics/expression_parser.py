"""大集中统计系统：业务规则表达式求值（受限解析器，禁止 eval）。

对应 VBA ``Evaluate(...)`` 的受控替代。语法与 Excel 公式一致：

- 数字（含 ``1.5E+3`` 科学计数法）、括号、四则运算与乘方 ``^``；
- 比较运算 ``= <> < <= > >=``（左结合，返回 TRUE/FALSE）；
- 函数（大小写不敏感）：逻辑 ``AND OR NOT IF IFS IFERROR ISERROR``；
  数值 ``ABS ROUND ROUNDUP ROUNDDOWN TRUNC INT MOD CEILING FLOOR SIGN SQRT
  POWER MAX MIN SUM AVERAGE COUNT``；
- 命名值 ``left right Thd`` 及其他通过 ``variables`` 传入的变量；
- 布尔量 ``TRUE FALSE``；算术中 TRUE=1、FALSE=0。

Excel 语义细节：``-2^2 = 4``（负号先于乘方作用于底数）、``2^3^2 = 64``
（乘方左结合）、``2^-2 = 0.25``、``1/0`` 为除零错误；ROUND 四舍五入（半值
远离零）、INT 向下取整、MOD 实数取模（符号随除数）——与 Excel 16/WPS 12
实测对齐（见 ``excel_round`` 等共享内核，expression_lae 同源复用）。

任何无法解析、未知名称、未知函数、除零都抛出 :class:`ExpressionError`，
绝不静默按 0 处理；调用方负责把异常写入规则错误日志。
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from decimal import Decimal


class ExpressionError(Exception):
    """表达式无法解析或求值失败；message 面向运行日志，可直接输出。"""


class ExcelCalculationError(ExpressionError):
    """Excel 求值期错误（#DIV/0!、#NUM! 等口径）。

    与解析/能力错误区分：IFERROR/ISERROR 只捕获本类（及子类），
    未知函数等能力问题不受 IFERROR 吞掉。
    """


class ExcelNaError(ExcelCalculationError):
    """#N/A 语义（Excel ``NA()`` / ``IFS`` 无匹配）。

    审核工具口径：**“规则没有得出结论”必须显形**——按错误处理并写入运行日志，
    不得静默当作 FALSE（漏报是审核工具最贵的失败模式）。
    ISNA 精确捕获本类；IFERROR/ISERROR 作为求值期错误一并捕获。
    """


@dataclass(frozen=True)
class _Token:
    kind: str   # num / name / op / lparen / rparen / comma
    text: str
    pos: int


_TOKEN_RE = re.compile(
    r"""
    (?P<space>\s+)
  | (?P<num>(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)
  | (?P<name>[A-Za-z_][A-Za-z0-9_]*)
  | (?P<op><>|<=|>=|[-+*/^=<>])
  | (?P<lparen>\()
  | (?P<rparen>\))
  | (?P<comma>,)
    """,
    re.VERBOSE,
)


def _tokenize(text: str) -> list[_Token]:
    tokens: list[_Token] = []
    pos = 0
    while pos < len(text):
        match = _TOKEN_RE.match(text, pos)
        if match is None:
            raise ExpressionError(
                f"表达式包含无法识别的字符「{text[pos]}」（位置 {pos + 1}）"
            )
        pos = match.end()
        kind = match.lastgroup
        if kind == "space":
            continue
        tokens.append(_Token(kind, match.group(), match.start()))
    return tokens


_FUNCTIONS = {
    "AND", "OR", "NOT", "IF", "IFS", "IFERROR", "ISERROR", "ISNA",
    "ABS", "ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC", "INT", "MOD",
    "CEILING", "FLOOR", "SIGN", "SQRT", "POWER", "MAX", "MIN", "SUM",
    "AVERAGE", "COUNT",
}

# 求值期错误的统一错误类型（ExcelCalculationError）；
# expression_lae 的 _EvalError 即本类别名，保证 IFERROR/ISERROR 语义跨引擎一致。
_EVAL_ERROR = ExcelCalculationError


# ---------------------------------------------------------------------------
# Excel 数值语义共享内核（SBE 与 expression_lae 同源复用；语义经 Excel 16 /
# WPS 12 双引擎实测对齐）
# ---------------------------------------------------------------------------

def excel_round(value: float, digits: int, rounding: str) -> float:
    """按 Excel 口径做十进制舍入（在 15 位有效数字表示上舍入，非二进制位）。

    ROUND_HALF_UP=四舍五入（ROUND，半值远离零），ROUND_UP=远离零
    （ROUNDUP），ROUND_DOWN=趋零（ROUNDDOWN/TRUNC）。digits 可为负。
    """
    quantum = Decimal(1).scaleb(-digits)
    return float(Decimal(repr(value)).quantize(quantum, rounding=rounding))


def excel_int(value: float) -> float:
    """Excel INT：向下取整（INT(-5.5)=-6）。"""
    return float(math.floor(value))


def excel_mod(number: float, divisor: float) -> float:
    """Excel MOD：实数取模，结果符号随除数；MOD(n,0) → #DIV/0!。"""
    if divisor == 0:
        raise _EVAL_ERROR("MOD 除零")
    return number - divisor * math.floor(number / divisor)


def excel_ceiling(number: float, significance: float) -> float:
    """Excel 旧版 CEILING：s=0 → 0；n>0 且 s<0 → #NUM!；否则 s*ceil(n/s)。"""
    if significance == 0:
        return 0.0
    if number > 0 and significance < 0:
        raise _EVAL_ERROR("CEILING 符号不符")
    return significance * math.ceil(number / significance)


def excel_floor(number: float, significance: float) -> float:
    """Excel 旧版 FLOOR：s=0 → #DIV/0!；n>0 且 s<0 → #NUM!；否则 s*floor(n/s)。

    已知 provider 分歧：FLOOR(-2.5,1) 在 Excel=-3、WPS 12=#NUM!（取 Excel 口径）。
    """
    if significance == 0:
        raise _EVAL_ERROR("FLOOR 除零")
    if number > 0 and significance < 0:
        raise _EVAL_ERROR("FLOOR 符号不符")
    return significance * math.floor(number / significance)


def excel_sign(value: float) -> float:
    return float((value > 0) - (value < 0))


def excel_sqrt(value: float) -> float:
    if value < 0:
        raise _EVAL_ERROR("SQRT 负数无实数解")
    return math.sqrt(value)


def excel_power(number: float, exponent: float) -> float:
    """Excel POWER：POWER(0,0) 与 负数底非整数指数 → #NUM!；溢出 → #NUM!。

    实测口径修正（2026-09-18 Windows 复核）：Excel Evaluate 对 POWER(0,0)
    返回 #NUM!（UOS 报告的「Excel 口径即 1」不成立，LO=1 属 LO 自身差异）。
    """
    if number == 0 and exponent == 0:
        raise _EVAL_ERROR("POWER(0,0) 未定义")
    try:
        result = number ** exponent
    except (OverflowError, ZeroDivisionError) as exc:
        raise _EVAL_ERROR(f"POWER 运算错误：{exc}") from exc
    if isinstance(result, complex):
        raise _EVAL_ERROR("POWER 无实数解")
    if result != result or result in (float("inf"), float("-inf")):
        raise _EVAL_ERROR("数值溢出")
    return result


# ---- AST 节点（元组形式：("num", v) / ("var", name) / ("un", op, a) /
# ("bin", op, a, b) / ("cmp", op, a, b) / ("call", name, args)） ----


class _Parser:
    def __init__(self, tokens: list[_Token]) -> None:
        self.tokens = tokens
        self.index = 0

    def _peek(self) -> _Token | None:
        return self.tokens[self.index] if self.index < len(self.tokens) else None

    def _next(self) -> _Token:
        token = self._peek()
        if token is None:
            raise ExpressionError("表达式意外结束")
        self.index += 1
        return token

    def parse(self):
        node = self._comparison()
        leftover = self._peek()
        if leftover is not None:
            raise ExpressionError(f"表达式有多余内容「{leftover.text}」")
        return node

    def _comparison(self):
        node = self._additive()
        while (token := self._peek()) is not None and token.kind == "op" and token.text in {"=", "<>", "<", "<=", ">", ">="}:
            self._next()
            right = self._additive()
            node = ("cmp", token.text, node, right)
        return node

    def _additive(self):
        node = self._multiplicative()
        while (token := self._peek()) is not None and token.kind == "op" and token.text in {"+", "-"}:
            self._next()
            node = ("bin", token.text, node, self._multiplicative())
        return node

    def _multiplicative(self):
        node = self._unary()
        while (token := self._peek()) is not None and token.kind == "op" and token.text in {"*", "/"}:
            self._next()
            node = ("bin", token.text, node, self._unary())
        return node

    def _unary(self):
        # 符号不在此处消费：交给 _power 的 _signed_primary，保证 -2^2 = (-2)^2 = 4
        # （Excel 语义）；对 * / 而言 _signed_primary 的结果等价于普通一元负号。
        return self._power()

    def _power(self):
        # Excel：负号先于乘方作用于底数（-2^2=4），乘方左结合（2^3^2=64），
        # 指数自身可带符号（2^-2）。
        node = self._signed_primary()
        while (token := self._peek()) is not None and token.kind == "op" and token.text == "^":
            self._next()
            node = ("bin", "^", node, self._signed_primary())
        return node

    def _signed_primary(self):
        sign = 1
        while (token := self._peek()) is not None and token.kind == "op" and token.text in {"+", "-"}:
            self._next()
            if token.text == "-":
                sign = -sign
        node = self._primary()
        return ("un", "-", node) if sign < 0 else node

    def _primary(self):
        token = self._next()
        if token.kind == "num":
            return ("num", float(token.text))
        if token.kind == "name":
            upper = token.text.upper()
            if upper == "TRUE":
                return ("bool", True)
            if upper == "FALSE":
                return ("bool", False)
            if upper in _FUNCTIONS:
                return self._call(upper)
            return ("var", token.text)
        if token.kind == "lparen":
            node = self._comparison()
            closing = self._next()
            if closing.kind != "rparen":
                raise ExpressionError(f"括号不匹配：「{closing.text}」（位置 {closing.pos + 1}）")
            return node
        raise ExpressionError(f"意外的符号「{token.text}」（位置 {token.pos + 1}）")

    _ARITY = {
        "NOT": (1, 1), "IF": (2, 3), "IFS": (2, None), "IFERROR": (1, 2),
        "ISERROR": (1, 1), "ISNA": (1, 1), "ABS": (1, 1), "ROUND": (1, 2), "ROUNDUP": (2, 2),
        "ROUNDDOWN": (2, 2), "TRUNC": (1, 2), "INT": (1, 1), "MOD": (2, 2),
        "CEILING": (2, 2), "FLOOR": (2, 2), "SIGN": (1, 1), "SQRT": (1, 1),
        "POWER": (2, 2), "AVERAGE": (1, None), "COUNT": (1, None),
    }

    def _call(self, name: str):
        opening = self._next()
        if opening.kind != "lparen":
            raise ExpressionError(f"函数 {name} 后缺少括号")
        args = []
        if (peek := self._peek()) is not None and peek.kind == "rparen":
            self._next()
        else:
            while True:
                args.append(self._comparison())
                token = self._next()
                if token.kind == "comma":
                    continue
                if token.kind == "rparen":
                    break
                raise ExpressionError(f"函数 {name} 参数列表格式错误：「{token.text}」")
        low, high = self._ARITY.get(name, (1, None))
        if len(args) < low or (high is not None and len(args) > high):
            expected = f"{low} 个参数" if high == low else (
                f"{low}~{high} 个参数" if high else f"至少 {low} 个参数")
            raise ExpressionError(f"{name} 需要 {expected}，实际 {len(args)} 个")
        if name in {"AND", "OR", "MAX", "MIN", "SUM"} and not args:
            raise ExpressionError(f"{name} 至少需要 1 个参数")
        return ("call", name, args)


def _to_number(value) -> float:
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    return float(value)


def _to_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    return float(value) != 0.0


class _Evaluator:
    def __init__(self, variables: dict[str, float]) -> None:
        self.variables = {name.casefold(): value for name, value in variables.items()}

    def evaluate(self, node):
        kind = node[0]
        if kind == "num":
            return node[1]
        if kind == "bool":
            return node[1]
        if kind == "var":
            name = node[1]
            if name.casefold() not in self.variables:
                raise ExpressionError(f"表达式引用了未知名称「{name}」")
            return float(self.variables[name.casefold()])
        if kind == "un":
            return -_to_number(self.evaluate(node[2]))
        if kind == "bin":
            return self._binary(node[1], node[2], node[3])
        if kind == "cmp":
            return self._compare(node[1], node[2], node[3])
        if kind == "call":
            return self._call(node[1], node[2])
        raise ExpressionError(f"内部错误：未知节点 {kind}")

    def _binary(self, op: str, left_node, right_node) -> float:
        left = _to_number(self.evaluate(left_node))
        right = _to_number(self.evaluate(right_node))
        if op == "+":
            result = left + right
        elif op == "-":
            result = left - right
        elif op == "*":
            result = left * right
        elif op == "/":
            if right == 0:
                raise _EVAL_ERROR("除零错误")
            result = left / right
        elif op == "^":
            try:
                result = left ** right
            except (OverflowError, ValueError, ZeroDivisionError) as exc:
                raise _EVAL_ERROR(f"乘方计算失败：{left}^{right}") from exc
            if isinstance(result, complex):
                raise _EVAL_ERROR(f"乘方结果为复数：{left}^{right}")
        else:
            raise ExpressionError(f"内部错误：未知运算 {op}")
        # Excel：溢出为 #NUM! 错误，不是 inf。
        if result != result or result in (float("inf"), float("-inf")):
            raise _EVAL_ERROR("数值溢出")
        return result

    def _compare(self, op: str, left_node, right_node) -> bool:
        # Excel workbook values arrive here as Python floats. Comparing the
        # result of a financial subtraction as a float can expose binary tails
        # (for example, two decimal differences that both equal -5443738.79).
        # Re-evaluate simple arithmetic comparison operands with Decimal made
        # from their shortest decimal spellings. This is exact decimal equality,
        # not a tolerance: meaningful differences remain different. Expressions
        # involving functions or other operators retain the established float
        # evaluation path.
        left_decimal = self._decimal_comparison_value(left_node)
        right_decimal = self._decimal_comparison_value(right_node)
        if left_decimal is not None and right_decimal is not None:
            left, right = left_decimal, right_decimal
        else:
            left = _to_number(self.evaluate(left_node))
            right = _to_number(self.evaluate(right_node))
        return {
            "=": left == right,
            "<>": left != right,
            "<": left < right,
            "<=": left <= right,
            ">": left > right,
            ">=": left >= right,
        }[op]

    def _decimal_comparison_value(self, node) -> Decimal | None:
        """Return exact decimal value for a simple arithmetic comparison tree.

        Workbook cell values are floats by the time they reach this parser, but
        ``str(float)`` yields Python's shortest round-trippable decimal spelling
        (for example ``100584251.06`` rather than its binary expansion). Using
        that spelling for +/- avoids subtraction-only noise without applying a
        broad tolerance or rounding away small, real differences.
        """
        kind = node[0]
        if kind == "num":
            value = Decimal(str(node[1]))
            return value if value.is_finite() else None
        if kind == "bool":
            return Decimal(1 if node[1] else 0)
        if kind == "var":
            name = node[1].casefold()
            if name not in self.variables:
                return None
            try:
                value = Decimal(str(self.variables[name]))
                return value if value.is_finite() else None
            except (ArithmeticError, TypeError, ValueError):
                return None
        if kind == "un":
            value = self._decimal_comparison_value(node[2])
            if value is None:
                return None
            return -value
        if kind == "bin" and node[1] in {"+", "-"}:
            left = self._decimal_comparison_value(node[2])
            right = self._decimal_comparison_value(node[3])
            if left is None or right is None:
                return None
            return left + right if node[1] == "+" else left - right
        return None

    def _call(self, name: str, args):
        # ---- 惰性函数：先判条件，未选中的分支不求值（Excel/IFERROR 口径）----
        if name == "NOT":
            return not _to_bool(self.evaluate(args[0]))
        if name == "IF":
            if _to_bool(self.evaluate(args[0])):
                return self.evaluate(args[1])
            return self.evaluate(args[2]) if len(args) == 3 else False
        if name == "IFS":
            for index in range(0, len(args), 2):
                if _to_bool(self.evaluate(args[index])):
                    return self.evaluate(args[index + 1])
            # Excel 口径：IFS 无匹配 → #N/A。审核工具不得静默当 FALSE（漏报）：
            # #N/A 使规则在运行日志中显形，倒逼作者补 TRUE 兜底分支。
            raise ExcelNaError("IFS 无匹配分支（#N/A）：请在最后一对写 TRUE, 兜底值")
        if name == "IFERROR":
            # 只吞求值期错误（ExcelCalculationError，如除零/溢出/#NUM! 类）；
            # 未知名称等解析/引用错误照常抛出。
            try:
                return self.evaluate(args[0])
            except ExcelCalculationError:
                return self.evaluate(args[1]) if len(args) == 2 else 0.0
        if name == "ISERROR":
            try:
                self.evaluate(args[0])
                return False
            except ExcelCalculationError:
                return True
        if name == "ISNA":
            # Excel：仅 #N/A 为真；#DIV/0! 等其他错误为假。
            try:
                self.evaluate(args[0])
                return False
            except ExcelNaError:
                return True
            except ExcelCalculationError:
                return False
        # ---- 严格函数 ----
        if name == "COUNT":
            # 布尔与数值都计入；文本在本解析器中不存在，等价于计参数个数。
            return float(sum(1 for arg in args if isinstance(self.evaluate(arg), (int, float, bool))))
        values = [_to_number(self.evaluate(item)) for item in args]
        if name == "AND":
            return all(v != 0 for v in values)
        if name == "OR":
            return any(v != 0 for v in values)
        if name == "ABS":
            return abs(values[0])
        if name == "ROUND":
            return excel_round(values[0], int(values[1]) if len(values) > 1 else 0, "ROUND_HALF_UP")
        if name == "ROUNDUP":
            return excel_round(values[0], int(values[1]), "ROUND_UP")
        if name == "ROUNDDOWN":
            return excel_round(values[0], int(values[1]), "ROUND_DOWN")
        if name == "TRUNC":
            return excel_round(values[0], int(values[1]) if len(values) > 1 else 0, "ROUND_DOWN")
        if name == "INT":
            return excel_int(values[0])
        if name == "MOD":
            return excel_mod(values[0], values[1])
        if name == "CEILING":
            return excel_ceiling(values[0], values[1])
        if name == "FLOOR":
            return excel_floor(values[0], values[1])
        if name == "SIGN":
            return excel_sign(values[0])
        if name == "SQRT":
            return excel_sqrt(values[0])
        if name == "POWER":
            return excel_power(values[0], values[1])
        if name == "MAX":
            return max(values)
        if name == "MIN":
            return min(values)
        if name == "SUM":
            return sum(values)
        if name == "AVERAGE":
            return sum(values) / len(values)
        raise ExpressionError(f"内部错误：未知函数 {name}")


#: AST 解析缓存（与 expression_lae._AST_CACHE 同步口径）：同一规则文本
# Parse 一次，多机构/多行复用；AST 求值只读，缓存安全。
_PARSE_CACHE: dict[str, tuple] = {}
_PARSE_CACHE_LIMIT = 8192
_PARSE_STATS = {"hits": 0, "misses": 0}


def reset_cache_stats() -> None:
    """清零解析缓存命中统计（每次流程启动时调用）。"""
    _PARSE_STATS["hits"] = 0
    _PARSE_STATS["misses"] = 0


def cache_stats() -> dict[str, int]:
    """返回 {"hits", "misses"}（性能观测：expression_parse_count 口径）。"""
    return dict(_PARSE_STATS)


def parse_cached(text: str) -> tuple:
    """解析（带缓存）表达式文本为 AST；语法/未知函数错误抛 ExpressionError。"""
    cached = _PARSE_CACHE.get(text)
    if cached is not None:
        _PARSE_STATS["hits"] += 1
        return cached
    _PARSE_STATS["misses"] += 1
    tokens = _tokenize(text)
    ast = _Parser(tokens).parse()
    if len(_PARSE_CACHE) < _PARSE_CACHE_LIMIT:
        _PARSE_CACHE[text] = ast
    return ast


def evaluate_expression(text: str, variables: dict[str, float] | None = None) -> "float | bool":
    """求值一个受限表达式；失败抛 :class:`ExpressionError`。

    ``variables`` 提供命名值（如 left/right/Thd）；未提供的名称一律报错。
    返回 float 或 bool；调用方按需用 ``_to_bool`` 语义（非 0 即真）判定。
    """
    if text is None or not str(text).strip():
        raise ExpressionError("表达式为空")
    ast = parse_cached(str(text))
    return _Evaluator(variables or {}).evaluate(ast)


def evaluate_boolean(text: str, variables: dict[str, float] | None = None) -> bool:
    """求值并按 Excel ``CBool`` 语义转为布尔（非 0 即真，错误抛出）。"""
    return _to_bool(evaluate_expression(text, variables))
