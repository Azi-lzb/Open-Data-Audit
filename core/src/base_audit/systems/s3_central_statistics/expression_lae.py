"""统一表达式校验 Expression V2（并行验证版；不接生产路径）。

目标：用一套「Excel 兼容表达式 + 本期/上期引用 + 本期日期/上期日期」覆盖
现有「累计指标规则 + 特殊指标规则 + 表达式校验」三类规则的表达能力。

本轮定位（严格遵循实施要求）：

- **纯 Python**：不依赖 Excel/WPS COM、LibreOffice UNO、soffice；三平台共用。
- **禁止 eval/exec**：Tokenizer → Parser（AST）→ Evaluator，函数/变量/操作符三重白名单。
- **三态结果**：OK / UNSUPPORTED（未知函数等）/ ERROR（解析或求值失败）；
  UNSUPPORTED 与 ERROR 绝不静默当成 FALSE。
- **惰性求值**：IF / IFS / AND / OR 短路，未命中的分支不求值（含引用 token），
  因此 ``IF({A}=0, TRUE, [A]/{A}<1.3)`` 在 {A}=0 时不产生除零。
- **引用语义与既有实现一致**：``[...]`` = 本期、``{...}`` = 上期，8 段 token
  （业务类,机构类代码,地区代码,指标代码,数据属性,币种,频度,批次），空段继承当前行；
  缺失记录→0，存在为 0→0.01 哨兵，存在非 0→×单位因子（与 complex_rule_engine 相同）。
- **日期是独立上下文变量**：``本期日期`` / ``上期日期``，读取 CSV 阶段解析一次放入
  ``ExpressionContext``，规则执行时直接读取，禁止逐条重解析。

与 VBA 的差异不在此处修正——差异报告负责记录，业务确认后再定口径。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from typing import Any

from .expression_parser import (
    ExcelCalculationError,
    ExcelNaError,
    ExpressionError,
    excel_ceiling,
    excel_floor,
    excel_int,
    excel_mod,
    excel_power,
    excel_round,
    excel_sign,
    excel_sqrt,
)

# ---------------------------------------------------------------------------
# 白名单
# ---------------------------------------------------------------------------

#: 支持的函数（大小写不敏感）。未列入者 → UNSUPPORTED（不得当 FALSE）。
SUPPORTED_FUNCTIONS = frozenset({
    "AND", "OR", "NOT", "IF", "IFS",
    "ABS", "ROUND", "INT", "MOD", "MIN", "MAX", "SUM",
    "YEAR", "MONTH", "DAY",
    # 引用状态函数（§C：严格区分 缺失记录/存在且为0/空值/非零值）
    "EXISTS", "MISSING", "ISBLANK", "RAW",
    # Excel 精确语义函数库（2026-09 对齐；语义以 Excel 16 / WPS 12 实测为准）
    "ROUNDUP", "ROUNDDOWN", "TRUNC", "CEILING", "FLOOR",
    "SIGN", "SQRT", "POWER", "AVERAGE", "COUNT",
    "IFERROR", "ISERROR", "ISNA",
})

#: 惰性求值函数：条件为假时不求值对应结果分支。
LAZY_FUNCTIONS = frozenset({"IF", "IFS", "AND", "OR", "IFERROR", "ISERROR", "ISNA"})

#: 上下文变量白名单（除 8 段引用 token 外可用的裸名称）。
SUPPORTED_VARIABLES = frozenset({"本期日期", "上期日期", "TRUE", "FALSE"})
#: 容差变量：大小写不敏感（Thd/thd），取值来自 ExpressionContext.thd。
#: 与 SBE 的 {"Thd": 值} 变量口径一致（复杂校验规则表「容差值(万元)」×10000）。
THD_VARIABLE = "THD"

#: 二元操作符。
_BINARY_OPS = frozenset({"+", "-", "*", "/", "^"})
_COMPARE_OPS = frozenset({"=", "<>", ">", ">=", "<", "<="})

STATUS_OK = "OK"
STATUS_UNSUPPORTED = "UNSUPPORTED"
STATUS_ERROR = "ERROR"


class UnsupportedFunctionError(ExpressionError):
    """表达式使用了白名单外的函数/变量/操作符。"""


# ---------------------------------------------------------------------------
# Tokenizer（支持 8 段引用 token 与中文标识符）
# ---------------------------------------------------------------------------

_TOKEN_RE = re.compile(
    r"""
    (?P<space>\s+)
  | (?P<ref>[\[\{][^\[\]\{\}]*[\]\}])
  | (?P<date_var>本期日期|上期日期)
  | (?P<num>(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)
  | (?P<name>[A-Za-z_\u4e00-\u9fff][A-Za-z0-9_\u4e00-\u9fff]*)
  | (?P<op><>|<=|>=|[-+*/^=<>])
  | (?P<lparen>\()
  | (?P<rparen>\))
  | (?P<comma>,)
    """,
    re.VERBOSE,
)


@dataclass(frozen=True)
class _Tok:
    kind: str      # ref / date_var / num / name / op / lparen / rparen / comma
    text: str
    pos: int


def tokenize(text: str) -> list[_Tok]:
    toks: list[_Tok] = []
    pos = 0
    while pos < len(text):
        m = _TOKEN_RE.match(text, pos)
        if m is None:
            raise ExpressionError(f"表达式包含无法识别的字符「{text[pos]}」（位置 {pos + 1}）")
        pos = m.end()
        if m.lastgroup == "space":
            continue
        toks.append(_Tok(m.lastgroup, m.group(), m.start()))
    return toks


# ---------------------------------------------------------------------------
# AST：元组节点
#   ("num", v) ("ref", raw, side) ("date", "本期日期"/"上期日期") ("bool", v)
#   ("un", op, a) ("bin", op, a, b) ("cmp", op, a, b) ("call", NAME, [args])
# ---------------------------------------------------------------------------


class _Parser:
    def __init__(self, toks: list[_Tok]) -> None:
        self.toks = toks
        self.i = 0

    def peek(self) -> _Tok | None:
        return self.toks[self.i] if self.i < len(self.toks) else None

    def take(self) -> _Tok:
        tok = self.peek()
        if tok is None:
            raise ExpressionError("表达式意外结束")
        self.i += 1
        return tok

    def expect(self, kind: str, what: str) -> _Tok:
        tok = self.peek()
        if tok is None or tok.kind != kind:
            raise ExpressionError(f"期望{what}")
        return self.take()

    def parse(self) -> tuple:
        node = self.comparison()
        if self.peek() is not None:
            raise ExpressionError(f"表达式在位置 {self.peek().pos + 1} 处有多余内容")
        return node

    def comparison(self) -> tuple:
        # Excel/VBA 口径：链式比较左结合（a>b>0 → (a>b)>0，布尔参与后续比较）。
        node = self.additive()
        while (tok := self.peek()) is not None and tok.kind == "op" and tok.text in _COMPARE_OPS:
            self.take()
            node = ("cmp", tok.text, node, self.additive())
        return node

    def additive(self) -> tuple:
        node = self.multiplicative()
        while True:
            tok = self.peek()
            if tok is not None and tok.kind == "op" and tok.text in ("+", "-"):
                self.take()
                node = ("bin", tok.text, node, self.multiplicative())
            else:
                return node

    def multiplicative(self) -> tuple:
        node = self.power()
        while True:
            tok = self.peek()
            if tok is not None and tok.kind == "op" and tok.text in ("*", "/"):
                self.take()
                node = ("bin", tok.text, node, self.power())
            else:
                return node

    def power(self) -> tuple:
        node = self.unary()
        tok = self.peek()
        if tok is not None and tok.kind == "op" and tok.text == "^":
            self.take()
            return ("bin", "^", node, self.unary())
        return node

    def unary(self) -> tuple:
        tok = self.peek()
        if tok is not None and tok.kind == "op" and tok.text in ("+", "-"):
            self.take()
            return ("un", tok.text, self.unary())
        return self.atom()

    def atom(self) -> tuple:
        tok = self.take()
        if tok.kind == "num":
            return ("num", float(tok.text))
        if tok.kind == "ref":
            side = "prev" if tok.text.startswith("{") else "cur"
            return ("ref", tok.text, side)
        if tok.kind == "date_var":
            return ("date", tok.text)
        if tok.kind == "lparen":
            node = self.comparison()
            self.expect("rparen", "右括号「)」")
            return node
        if tok.kind == "name":
            upper = tok.text.upper()
            if upper == "TRUE":
                return ("bool", True)
            if upper == "FALSE":
                return ("bool", False)
            if upper == THD_VARIABLE:
                return ("thd",)
            nxt = self.peek()
            if nxt is not None and nxt.kind == "lparen":
                self.take()
                args: list[tuple] = []
                if self.peek() is not None and self.peek().kind != "rparen":
                    args.append(self.comparison())
                    while self.peek() is not None and self.peek().kind == "comma":
                        self.take()
                        args.append(self.comparison())
                self.expect("rparen", "右括号「)」")
                return ("call", upper, args)
            # 裸名称：只允许上下文变量
            if tok.text in SUPPORTED_VARIABLES:
                return ("date", tok.text)
            raise UnsupportedFunctionError(f"不支持的变量或函数「{tok.text}」")
        raise ExpressionError(f"位置 {tok.pos + 1} 处语法错误：意外的「{tok.text}」")


_AST_CACHE: dict[str, tuple] = {}
_AST_CACHE_LIMIT = 8192
_CACHE_STATS = {"hits": 0, "misses": 0}


def reset_cache_stats() -> None:
    """清零 AST 缓存命中统计（每次流程启动时调用，观测一次流程的口径）。"""
    _CACHE_STATS["hits"] = 0
    _CACHE_STATS["misses"] = 0


def cache_stats() -> dict[str, int]:
    """返回 {"hits", "misses"}（性能观测：ast_cache_hit_count 等）。"""
    return dict(_CACHE_STATS)


def parse_expression(text: str) -> tuple:
    """解析为 AST；语法错误抛 ExpressionError，白名单外名称抛 UnsupportedFunctionError。

    同一表达式文本 Parse 一次后缓存复用（LAE：一条规则多机构共享 AST）；
    AST 只读求值，缓存安全。超限即停止新增（表达式集合有限，正常不会触顶）。
    """
    cached = _AST_CACHE.get(text)
    if cached is not None:
        _CACHE_STATS["hits"] += 1
        return cached
    _CACHE_STATS["misses"] += 1
    # 归一全角括号（VBA Evaluate 原生兼容全角；部分配置表达式混用全角/半角括号）
    body = str(text or "").strip().lstrip("=")
    body = body.replace("\uff08", "(").replace("\uff09", ")")
    if not body:
        raise ExpressionError("表达式为空")
    ast = _Parser(tokenize(body)).parse()
    if len(_AST_CACHE) < _AST_CACHE_LIMIT:
        _AST_CACHE[text] = ast
    return ast


# ---------------------------------------------------------------------------
# 引用解析（与 complex_rule_engine 同语义）
# ---------------------------------------------------------------------------


@dataclass
class ExpressionContext:
    """一次规则求值的上下文：日期解析一次、数据索引复用。"""

    current_date: str = ""
    previous_date: str = ""
    current_index: dict[str, Any] = field(default_factory=dict)
    previous_index: dict[str, Any] = field(default_factory=dict)
    row: Any = None                 # ComparisonRow（提供 8 段空位继承）
    unit_factor: float = 1.0        # 单位换算因子（亿元口径）
    exempt_indicators: frozenset[str] = frozenset()
    thd: float = 0.0
    zero_sentinel: float = 0.01
    # 性能观测（可选）：非 None 时累计 LAE 运行时计数（引用解析/短路/跳过分支）。
    stats: dict | None = None

    def date_of(self, side: str) -> date | None:
        raw = self.current_date if side == "cur" else self.previous_date
        try:
            return date.fromisoformat(str(raw)[:10])
        except (TypeError, ValueError):
            return None


def _target_key(token: str, row: Any) -> str:
    """8 段 token → 目标键；空段继承当前行，含 ^ 的机构/地区段不继承。"""
    body = token.strip().strip("[]{}").replace(" ", "")
    segments = [seg.strip() for seg in body.split(",")]
    while len(segments) < 8:
        segments.append("")
    record = row.record

    def pick(position: int, inherited: str) -> str:
        value = segments[position]
        if value and "^" not in value:
            return value.replace("'", "")
        return inherited

    return "".join((
        segments[0] or record.biz_class,
        pick(1, record.org_code),
        pick(2, record.region_code),
        segments[3].replace("'", "") or record.indicator,
        segments[4] or record.data_attr,
        segments[5] or record.currency,
        segments[6] or record.frequency,
        segments[7] or record.batch,
    ))


def _token_indicator(token: str) -> str:
    body = token.strip().strip("[]{}")
    segments = [seg.strip() for seg in body.split(",")]
    return segments[3].replace("'", "").strip() if len(segments) > 3 else ""


def reference_value(token: str, side: str, ctx: ExpressionContext) -> float:
    """引用求值：缺失记录→0；存在为 0→0.01 哨兵；存在非 0→×单位因子。"""
    index = ctx.current_index if side == "cur" else ctx.previous_index
    key = _target_key(token, ctx.row)
    record = index.get(key)
    if record is None:
        return 0.0
    value = record.value if isinstance(record.value, (int, float)) else 0.0
    if value == 0:
        return ctx.zero_sentinel
    factor = 1.0 if _token_indicator(token) in ctx.exempt_indicators else ctx.unit_factor
    return float(value) * factor


# ---------------------------------------------------------------------------
# Evaluator（惰性）
# ---------------------------------------------------------------------------


class _EvalError(ExcelCalculationError):
    """求值期错误（除零等）；与解析错误分开以便调用方标注 ERROR。

    统一继承共享内核的 ExcelCalculationError：IFERROR/ISERROR 的「只吞求值
    期错误、不吞能力错误」语义跨 SBE/LAE 一致。
    """


def _num(value: Any) -> float:
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if value is None:
        return 0.0
    raise _EvalError(f"无法把「{value}」当作数值")


def _bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if value is None:
        return False
    raise _EvalError(f"无法把「{value}」当作逻辑值")


def _finite(value: Any) -> float:
    """归一数值并拒绝非有限结果（Excel 溢出报 #NUM!，不是 inf）。"""
    number = _num(value)
    if number != number or number in (float("inf"), float("-inf")):
        raise _EvalError("数值溢出")
    return number


def evaluate_ast(node: tuple, ctx: ExpressionContext) -> Any:
    kind = node[0]
    if kind == "num":
        return node[1]
    if kind == "bool":
        return node[1]
    if kind == "thd":
        return _num(ctx.thd)
    if kind == "ref":
        if ctx.stats is not None:
            ctx.stats["reference_resolve_count"] = ctx.stats.get("reference_resolve_count", 0) + 1
        return reference_value(node[1], node[2], ctx)
    if kind == "date":
        return ctx.date_of("cur" if node[1] == "本期日期" else "prev")
    if kind == "un":
        value = _num(evaluate_ast(node[2], ctx))
        return value if node[1] == "+" else -value
    if kind == "bin":
        op = node[1]
        left = _num(evaluate_ast(node[2], ctx))
        right = _num(evaluate_ast(node[3], ctx))
        result: Any
        if op == "+":
            result = left + right
        elif op == "-":
            result = left - right
        elif op == "*":
            result = left * right
        elif op == "/":
            if right == 0:
                raise _EvalError("除零")
            result = left / right
        elif op == "^":
            try:
                result = left ** right
            except (OverflowError, ZeroDivisionError) as exc:
                raise _EvalError(f"幂运算错误：{exc}") from exc
            if isinstance(result, complex):
                raise _EvalError("幂运算无实数解")
        else:
            raise ExpressionError(f"未知操作符「{op}」")
        return _finite(result)
    if kind == "cmp":
        op = node[1]
        left = _num(evaluate_ast(node[2], ctx))
        right = _num(evaluate_ast(node[3], ctx))
        if op == "=":
            return left == right
        if op == "<>":
            return left != right
        if op == ">":
            return left > right
        if op == ">=":
            return left >= right
        if op == "<":
            return left < right
        if op == "<=":
            return left <= right
        raise ExpressionError(f"未知比较操作符「{op}」")
    if kind == "call":
        if ctx.stats is not None:
            ctx.stats["function_eval_count"] = ctx.stats.get("function_eval_count", 0) + 1
        return _call(node[1], node[2], ctx)
    raise ExpressionError(f"未知语法节点「{kind}」")


def _count_nodes(node: tuple) -> int:
    """AST 节点计数（性能观测：跳过分支的「少求值节点数」）。"""
    total = 1
    for part in node:
        if isinstance(part, tuple):
            total += _count_nodes(part)
        elif isinstance(part, list):
            total += sum(_count_nodes(item) for item in part)
    return total


def _call(name: str, args: list[tuple], ctx: ExpressionContext) -> Any:
    if name not in SUPPORTED_FUNCTIONS:
        raise UnsupportedFunctionError(f"不支持的函数「{name}」")

    # ---- 引用状态函数：参数必须是引用 token，且不做哨兵/单位换算 ----
    if name in ("EXISTS", "MISSING", "ISBLANK", "RAW"):
        if len(args) != 1 or args[0][0] != "ref":
            raise ExpressionError(f"{name} 的参数必须是一个 [...] 或 {{...}} 引用")
        _, token, side = args[0]
        index = ctx.current_index if side == "cur" else ctx.previous_index
        key = _target_key(token, ctx.row)
        record = index.get(key)
        exists = record is not None
        if name == "EXISTS":
            return exists
        if name == "MISSING":
            return not exists
        if name == "ISBLANK":
            if not exists:
                return False
            value = getattr(record, "value", None)
            return value is None or (isinstance(value, str) and not value.strip())
        # RAW：原始数值（无 0.01 哨兵、无单位换算）；缺失 → 0（与 VBA 查表缺失同义）
        if not exists:
            return 0.0
        value = getattr(record, "value", None)
        if value is None or (isinstance(value, str) and not value.strip()):
            return 0.0
        return float(value) if isinstance(value, (int, float)) else 0.0

    # ---- 惰性函数 ----
    if name == "IF":
        if len(args) not in (2, 3):
            raise ExpressionError("IF 需要 2~3 个参数")
        if _bool(evaluate_ast(args[0], ctx)):
            return evaluate_ast(args[1], ctx)
        if ctx.stats is not None:
            ctx.stats["if_skipped_branch_count"] = ctx.stats.get("if_skipped_branch_count", 0) + 1
            ctx.stats["skipped_ast_node_count"] = ctx.stats.get("skipped_ast_node_count", 0) + (
                _count_nodes(args[1]) if len(args) > 1 else 0)
        return evaluate_ast(args[2], ctx) if len(args) == 3 else False
    if name == "IFS":
        if not args or len(args) % 2:
            raise ExpressionError("IFS 需要成对的「条件, 结果」参数")
        for i in range(0, len(args), 2):
            if _bool(evaluate_ast(args[i], ctx)):
                return evaluate_ast(args[i + 1], ctx)
            if ctx.stats is not None:
                ctx.stats["ifs_skipped_branch_count"] = ctx.stats.get("ifs_skipped_branch_count", 0) + 1
                ctx.stats["skipped_ast_node_count"] = ctx.stats.get("skipped_ast_node_count", 0) + (
                    _count_nodes(args[i + 1]) if i + 1 < len(args) else 0)
        # Excel 口径（同 SBE）：IFS 无匹配 → #N/A，不静默当 FALSE（漏报）。
        raise ExcelNaError("IFS 无匹配分支（#N/A）：请在最后一对写 TRUE, 兜底值")
    if name == "AND":
        for arg in args:
            if not _bool(evaluate_ast(arg, ctx)):
                if ctx.stats is not None:
                    ctx.stats["and_short_circuit_count"] = ctx.stats.get("and_short_circuit_count", 0) + 1
                    ctx.stats["skipped_ast_node_count"] = ctx.stats.get("skipped_ast_node_count", 0) + sum(
                        _count_nodes(item) for item in args[args.index(arg) + 1:])
                return False
        return True
    if name == "OR":
        for arg in args:
            if _bool(evaluate_ast(arg, ctx)):
                if ctx.stats is not None:
                    ctx.stats["or_short_circuit_count"] = ctx.stats.get("or_short_circuit_count", 0) + 1
                    ctx.stats["skipped_ast_node_count"] = ctx.stats.get("skipped_ast_node_count", 0) + sum(
                        _count_nodes(item) for item in args[args.index(arg) + 1:])
                return True
        return False
    if name == "IFERROR":
        # 求值期错误（除零/溢出/无实数解…）转 fallback；未提供 fallback → 0（Excel 语义）。
        if len(args) not in (1, 2):
            raise ExpressionError("IFERROR 需要 1~2 个参数")
        try:
            return evaluate_ast(args[0], ctx)
        except ExcelCalculationError:
            return evaluate_ast(args[1], ctx) if len(args) == 2 else 0.0
    if name == "ISERROR":
        if len(args) != 1:
            raise ExpressionError("ISERROR 需要 1 个参数")
        try:
            evaluate_ast(args[0], ctx)
            return False
        except ExcelCalculationError:
            return True
    if name == "ISNA":
        if len(args) != 1:
            raise ExpressionError("ISNA 需要 1 个参数")
        try:
            evaluate_ast(args[0], ctx)
            return False
        except ExcelNaError:
            return True
        except ExcelCalculationError:
            return False

    # ---- 严格函数 ----
    # 日期函数必须早于数值归一化：参数是 date 对象，_num() 会失败。
    if name in ("YEAR", "MONTH", "DAY"):
        if len(args) != 1:
            raise ExpressionError(f"{name} 需要 1 个参数")
        d = _as_date_value(args[0], ctx)
        return float({"YEAR": d.year, "MONTH": d.month, "DAY": d.day}[name])

    values = [_num(evaluate_ast(arg, ctx)) for arg in args]
    if name == "NOT":
        if len(values) != 1:
            raise ExpressionError("NOT 需要 1 个参数")
        return not _bool(values[0])
    if name == "ABS":
        return abs(values[0])
    if name == "ROUND":
        # Excel：四舍五入（半值远离零），digits 可为负；共享内核（SBE 同源）。
        if not values or len(values) > 2:
            raise ExpressionError("ROUND 需要 1~2 个参数")
        return excel_round(values[0], int(values[1]) if len(values) > 1 else 0, "ROUND_HALF_UP")
    if name == "ROUNDUP":
        if len(values) != 2:
            raise ExpressionError("ROUNDUP 需要 2 个参数")
        return excel_round(values[0], int(values[1]), "ROUND_UP")
    if name == "ROUNDDOWN":
        if len(values) != 2:
            raise ExpressionError("ROUNDDOWN 需要 2 个参数")
        return excel_round(values[0], int(values[1]), "ROUND_DOWN")
    if name == "TRUNC":
        if not values or len(values) > 2:
            raise ExpressionError("TRUNC 需要 1~2 个参数")
        return excel_round(values[0], int(values[1]) if len(values) > 1 else 0, "ROUND_DOWN")
    if name == "INT":
        return excel_int(values[0])
    if name == "MOD":
        if len(values) != 2:
            raise ExpressionError("MOD 需要 2 个参数")
        return excel_mod(values[0], values[1])
    if name == "CEILING":
        if len(values) != 2:
            raise ExpressionError("CEILING 需要 2 个参数")
        return excel_ceiling(values[0], values[1])
    if name == "FLOOR":
        if len(values) != 2:
            raise ExpressionError("FLOOR 需要 2 个参数")
        return excel_floor(values[0], values[1])
    if name == "SIGN":
        if len(values) != 1:
            raise ExpressionError("SIGN 需要 1 个参数")
        return excel_sign(values[0])
    if name == "SQRT":
        if len(values) != 1:
            raise ExpressionError("SQRT 需要 1 个参数")
        return excel_sqrt(values[0])
    if name == "POWER":
        if len(values) != 2:
            raise ExpressionError("POWER 需要 2 个参数")
        return excel_power(values[0], values[1])
    if name == "AVERAGE":
        if not values:
            raise ExpressionError("AVERAGE 需要 1 个以上参数")
        return sum(values) / len(values)
    if name == "COUNT":
        # Excel：直接给出的数值与布尔都计入（COUNT(1,TRUE)=2；文本不支持）。
        counted = 0.0
        for arg in args:
            value = evaluate_ast(arg, ctx)
            if isinstance(value, (int, float, bool)):
                counted += 1.0
        return counted
    if name == "MIN":
        return min(values)
    if name == "MAX":
        return max(values)
    if name == "SUM":
        return sum(values)
    raise UnsupportedFunctionError(f"不支持的函数「{name}」")


def _as_date_value(node: tuple, ctx: ExpressionContext) -> date:
    value = evaluate_ast(node, ctx)
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return date.fromisoformat(value[:10])
        except ValueError as exc:
            raise _EvalError(f"无法把「{value}」当作日期") from exc
    raise _EvalError("YEAR/MONTH/DAY 需要日期参数（本期日期/上期日期）")


# ---------------------------------------------------------------------------
# 对外入口：三态结果
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ExpressionOutcome:
    """V2 表达式求值结果（三态）。status 取 OK / UNSUPPORTED / ERROR。"""

    status: str
    value: Any = None
    triggered: bool = False
    reason: str = ""

    @property
    def ok(self) -> bool:
        return self.status == STATUS_OK


def evaluate_expression_lae(text: str, ctx: ExpressionContext,
                           trigger_mode: str = "表达式成立") -> ExpressionOutcome:
    """求值一条 V2 表达式并按触发方式判定。

    ``trigger_mode``：``表达式成立`` → 结果为真时触发；``表达式不成立`` → 结果为假时触发。
    UNSUPPORTED / ERROR 一律不触发，但 status 如实回传（调用方必须记录，不得当 FALSE）。
    """
    try:
        ast = parse_expression(text)
    except UnsupportedFunctionError as exc:
        return ExpressionOutcome(STATUS_UNSUPPORTED, reason=str(exc))
    except ExpressionError as exc:
        return ExpressionOutcome(STATUS_ERROR, reason=f"解析失败：{exc}")
    try:
        value = evaluate_ast(ast, ctx)
    except UnsupportedFunctionError as exc:
        return ExpressionOutcome(STATUS_UNSUPPORTED, reason=str(exc))
    except ExpressionError as exc:
        return ExpressionOutcome(STATUS_ERROR, reason=f"求值失败：{exc}")
    except Exception as exc:                                  # noqa: BLE001
        return ExpressionOutcome(STATUS_ERROR, reason=f"求值异常：{type(exc).__name__}: {exc}")

    truth = _bool(value)
    want = str(trigger_mode or "表达式成立").strip()
    triggered = truth if want == "表达式成立" else (not truth)
    return ExpressionOutcome(STATUS_OK, value=value, triggered=triggered)
