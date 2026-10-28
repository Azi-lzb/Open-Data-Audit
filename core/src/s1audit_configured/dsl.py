from __future__ import annotations

from dataclasses import dataclass, replace
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from functools import lru_cache
from typing import Any, Callable, Mapping

from .errors import ConfigurationError, DSLParseError, EvaluationError
from .model import EvalValue
from .excel_precision import excel_round15


@dataclass(frozen=True)
class Token:
    kind: str
    value: str
    position: int


@dataclass(frozen=True)
class Node:
    pass


@dataclass(frozen=True)
class Literal(Node):
    value: Any
    kind: str


@dataclass(frozen=True)
class Variable(Node):
    name: str


@dataclass(frozen=True)
class TupleLiteral(Node):
    values: tuple[str, ...]


@dataclass(frozen=True)
class Unary(Node):
    operator: str
    operand: Node


@dataclass(frozen=True)
class Binary(Node):
    operator: str
    left: Node
    right: Node


@dataclass(frozen=True)
class Call(Node):
    name: str
    arguments: tuple[Node, ...]
    keywords: tuple[tuple[str, str], ...] = ()


_ALLOWED_CALLS = {
    "V_CUR", "V_PRE", "V_EXT", "ABS", "ROUND", "MIN", "MAX",
    "IF", "AND", "OR", "NOT", "SUM", "ISBLANK", "ISNUMBER",
}
_COMPARISONS = {"=", "==", "<>", "!=", "<", "<=", ">", ">="}
_TWO_CHAR_OPERATORS = {"==", "<>", "!=", "<=", ">="}


def _is_identifier_start(char: str) -> bool:
    return char == "_" or char.isalpha() or "\u3400" <= char <= "\u9fff"


def _is_identifier_char(char: str) -> bool:
    return _is_identifier_start(char) or char.isdigit()


def tokenize(text: str) -> tuple[Token, ...]:
    tokens: list[Token] = []
    index = 0
    while index < len(text):
        char = text[index]
        if char.isspace():
            index += 1
            continue
        if char == '"':
            start = index
            index += 1
            parts: list[str] = []
            while index < len(text):
                current = text[index]
                if current == '"':
                    index += 1
                    break
                if current == "\\":
                    index += 1
                    escapes = {'"': '"', "\\": "\\", "n": "\n", "r": "\r", "t": "\t"}
                    if index >= len(text) or text[index] not in escapes:
                        raise DSLParseError(f"字符串包含不支持的转义，位置 {index}")
                    parts.append(escapes[text[index]])
                    index += 1
                    continue
                parts.append(current)
                index += 1
            else:
                raise DSLParseError(f"字符串缺少结束引号，位置 {start}")
            tokens.append(Token("STRING", "".join(parts), start))
            continue
        if char.isdigit() or (char == "." and index + 1 < len(text) and text[index + 1].isdigit()):
            start = index
            saw_dot = char == "."
            index += 1
            while index < len(text):
                current = text[index]
                if current.isdigit():
                    index += 1
                elif current == "." and not saw_dot:
                    saw_dot = True
                    index += 1
                else:
                    break
            tokens.append(Token("NUMBER", text[start:index], start))
            continue
        if _is_identifier_start(char):
            start = index
            index += 1
            while index < len(text) and _is_identifier_char(text[index]):
                index += 1
            value = text[start:index]
            keyword = value.upper()
            tokens.append(Token(keyword if keyword in {"AND", "OR", "NOT", "TRUE", "FALSE"} else "IDENT", value, start))
            continue
        if char in "=<>!" and index + 1 < len(text) and text[index:index + 2] in _TWO_CHAR_OPERATORS:
            tokens.append(Token("OP", text[index:index + 2], index))
            index += 2
            continue
        if char in "+-*/=<>(){},:":
            kind = "OP" if char in "+-*/=" else {
                "<": "OP", ">": "OP", "!": "OP",
                "(": "LPAREN", ")": "RPAREN", "{": "LBRACE",
                "}": "RBRACE", ",": "COMMA", ":": "COLON",
            }[char]
            tokens.append(Token(kind, char, index))
            index += 1
            continue
        raise DSLParseError(f"不支持的字符 {char!r}，位置 {index}")
    tokens.append(Token("EOF", "", len(text)))
    return tuple(tokens)


class Parser:
    def __init__(self, text: str):
        self.text = text
        self.tokens = tokenize(text)
        self.index = 0

    @property
    def current(self) -> Token:
        return self.tokens[self.index]

    def peek(self, distance: int = 1) -> Token:
        return self.tokens[min(self.index + distance, len(self.tokens) - 1)]

    def advance(self) -> Token:
        token = self.current
        self.index += 1
        return token

    def accept(self, kind: str, value: str | None = None) -> Token | None:
        token = self.current
        if token.kind == kind and (value is None or token.value == value):
            return self.advance()
        return None

    def expect(self, kind: str, value: str | None = None) -> Token:
        token = self.accept(kind, value)
        if token is None:
            expected = value if value is not None else kind
            raise DSLParseError(f"期望 {expected}，位置 {self.current.position}")
        return token

    def parse_expression(self) -> Node:
        node = self.parse_or()
        self.expect("EOF")
        return node

    def parse_dictionary(self) -> dict[str, Node]:
        self.expect("LBRACE")
        result: dict[str, Node] = {}
        if self.accept("RBRACE"):
            self.expect("EOF")
            return result
        while True:
            key = self.expect("STRING").value
            if key in result:
                raise ConfigurationError(f"取值字典重复变量：{key}")
            self.expect("COLON")
            result[key] = self.parse_or()
            if self.accept("COMMA"):
                if self.accept("RBRACE"):
                    break
                continue
            self.expect("RBRACE")
            break
        self.expect("EOF")
        for node in result.values():
            validate_node(node)
        return result

    def parse_or(self) -> Node:
        node = self.parse_and()
        while self.accept("OR"):
            node = Binary("OR", node, self.parse_and())
        return node

    def parse_and(self) -> Node:
        node = self.parse_comparison()
        while self.accept("AND"):
            node = Binary("AND", node, self.parse_comparison())
        return node

    def parse_comparison(self) -> Node:
        node = self.parse_additive()
        if self.current.kind == "OP" and self.current.value in _COMPARISONS:
            operator = self.advance().value
            node = Binary(operator, node, self.parse_additive())
            if self.current.kind == "OP" and self.current.value in _COMPARISONS:
                raise DSLParseError(f"不支持连续比较，位置 {self.current.position}")
        return node

    def parse_additive(self) -> Node:
        node = self.parse_multiplicative()
        while self.current.kind == "OP" and self.current.value in {"+", "-"}:
            operator = self.advance().value
            node = Binary(operator, node, self.parse_multiplicative())
        return node

    def parse_multiplicative(self) -> Node:
        node = self.parse_unary()
        while self.current.kind == "OP" and self.current.value in {"*", "/"}:
            operator = self.advance().value
            node = Binary(operator, node, self.parse_unary())
        return node

    def parse_unary(self) -> Node:
        if self.current.kind in {"NOT"}:
            operator = self.advance().kind
            return Unary(operator, self.parse_unary())
        if self.current.kind == "OP" and self.current.value in {"+", "-"}:
            operator = self.advance().value
            return Unary(operator, self.parse_unary())
        return self.parse_primary()

    def parse_primary(self) -> Node:
        token = self.current
        if token.kind == "NUMBER":
            self.advance()
            try:
                return Literal(Decimal(token.value), "number")
            except InvalidOperation as exc:
                raise DSLParseError(f"无效数字，位置 {token.position}") from exc
        if token.kind == "STRING":
            self.advance()
            return Literal(token.value, "string")
        if token.kind in {"TRUE", "FALSE"}:
            self.advance()
            return Literal(token.kind == "TRUE", "boolean")
        if token.kind == "LBRACE":
            self.advance()
            name = self.expect("IDENT").value
            self.expect("RBRACE")
            return Variable(name)
        if token.kind == "IDENT":
            name = self.advance().value.upper()
            name = {"VCUR": "V_CUR", "VPRE": "V_PRE", "VEXT": "V_EXT", "EXT": "V_EXT"}.get(name, name)
            if not self.accept("LPAREN"):
                raise DSLParseError(f"标识符必须是受支持的函数调用，位置 {token.position}")
            arguments: list[Node] = []
            keywords: list[tuple[str, str]] = []
            if not self.accept("RPAREN"):
                while True:
                    if self.current.kind == "IDENT" and self.peek().kind == "OP" and self.peek().value == "=":
                        key = self.advance().value
                        self.advance()
                        if self.current.kind != "STRING":
                            raise DSLParseError(f"参数 {key}= 的值必须是字符串，位置 {self.current.position}")
                        keywords.append((key, self.advance().value))
                    else:
                        arguments.append(self.parse_argument())
                    if self.accept("COMMA"):
                        if self.accept("RPAREN"):
                            break
                        continue
                    self.expect("RPAREN")
                    break
            call = Call(name, tuple(arguments), tuple(keywords))
            validate_call(call)
            return call
        if token.kind == "LPAREN":
            self.advance()
            first = self.parse_or()
            self.expect("RPAREN")
            return first
        raise DSLParseError(f"此处需要值，位置 {token.position}")

    def parse_argument(self) -> Node:
        if (self.current.kind == "LPAREN" and self.peek().kind == "STRING"
                and self.peek(2).kind == "COMMA"):
            self.advance()
            values = [self.expect("STRING").value]
            while self.accept("COMMA"):
                if self.current.kind == "RPAREN":
                    break
                if self.current.kind != "STRING":
                    raise DSLParseError(f"路径元组只能包含字符串，位置 {self.current.position}")
                values.append(self.advance().value)
            self.expect("RPAREN")
            if any(not value.strip() for value in values):
                raise DSLParseError("路径元组各层不能为空")
            return TupleLiteral(tuple(values))
        return self.parse_or()


def parse_expression(text: str) -> Node:
    if not text or not text.strip():
        raise DSLParseError("表达式不能为空")
    return _parse_expression_cached(text)


@lru_cache(maxsize=8192)
def _parse_expression_cached(text: str) -> Node:
    node = Parser(text).parse_expression()
    validate_node(node)
    return node


def parse_dictionary(text: str | None) -> dict[str, Node]:
    if text is None or not str(text).strip():
        return {}
    return dict(_parse_dictionary_items(str(text)))


@lru_cache(maxsize=4096)
def _parse_dictionary_items(text: str) -> tuple[tuple[str, Node], ...]:
    # Keep cached AST values immutable. Callers receive a fresh dict so one
    # consumer cannot mutate the parsed dictionary seen by another rule.
    return tuple(Parser(text).parse_dictionary().items())


def validate_call(call: Call) -> None:
    if call.name not in _ALLOWED_CALLS:
        raise ConfigurationError(f"不支持的函数：{call.name}")
    keyword_names = [name for name, _ in call.keywords]
    if len(keyword_names) != len(set(keyword_names)):
        raise ConfigurationError(f"函数 {call.name} 存在重复具名参数")
    if any(name not in {"source", "sheet", "unit", "source_unit", "standard_unit", "metric_type"} for name in keyword_names):
        raise ConfigurationError(f"函数 {call.name} 的具名参数不受支持")
    if call.name not in {"V_CUR", "V_PRE", "V_EXT"} and call.keywords:
        raise ConfigurationError(f"函数 {call.name} 不接受具名参数")
    expected = {"V_CUR": (2, 3), "V_PRE": (2, 3), "V_EXT": (1, 1),
                "ABS": (1, 1), "ROUND": (2, 2), "MIN": (1, 99), "MAX": (1, 99),
                "IF": (3, 3), "AND": (1, 99), "OR": (1, 99), "NOT": (1, 1),
                "SUM": (1, 99), "ISBLANK": (1, 1), "ISNUMBER": (1, 1)}[call.name]
    if not expected[0] <= len(call.arguments) <= expected[1]:
        raise ConfigurationError(f"函数 {call.name} 参数个数错误")
    if call.name in {"V_CUR", "V_PRE", "V_EXT"}:
        path_arguments = call.arguments
        if call.name in {"V_CUR", "V_PRE"} and len(call.arguments) == 3:
            sheet_argument = call.arguments[0]
            if not isinstance(sheet_argument, Literal) or sheet_argument.kind != "string" or not sheet_argument.value.strip():
                raise ConfigurationError(f"函数 {call.name} 的第一个参数必须是非空工作表名称")
            if "sheet" in keyword_names:
                raise ConfigurationError(f"函数 {call.name} 已通过第一个参数指定工作表，不能再使用 sheet=")
            path_arguments = call.arguments[1:]
        for arg in path_arguments:
            if not isinstance(arg, (Literal, TupleLiteral)):
                raise ConfigurationError(f"函数 {call.name} 的定位参数必须是字符串常量或字符串元组")
            if isinstance(arg, Literal) and arg.kind != "string":
                raise ConfigurationError(f"函数 {call.name} 的定位参数必须是字符串")
        for _, value in call.keywords:
            if not value.strip():
                raise ConfigurationError("source= 和 sheet= 不能为空")


def validate_node(node: Node) -> None:
    if isinstance(node, Call):
        validate_call(node)
        for arg in node.arguments:
            validate_node(arg)
    elif isinstance(node, Unary):
        validate_node(node.operand)
    elif isinstance(node, Binary):
        validate_node(node.left)
        validate_node(node.right)


def variables_in(node: Node) -> set[str]:
    found: set[str] = set()
    if isinstance(node, Variable):
        found.add(node.name)
    elif isinstance(node, Call):
        for arg in node.arguments:
            found.update(variables_in(arg))
    elif isinstance(node, Unary):
        found.update(variables_in(node.operand))
    elif isinstance(node, Binary):
        found.update(variables_in(node.left))
        found.update(variables_in(node.right))
    return found


def calls_in(node: Node) -> tuple[Call, ...]:
    found: list[Call] = []
    if isinstance(node, Call):
        found.append(node)
        for argument in node.arguments:
            found.extend(calls_in(argument))
    elif isinstance(node, Unary):
        found.extend(calls_in(node.operand))
    elif isinstance(node, Binary):
        found.extend(calls_in(node.left))
        found.extend(calls_in(node.right))
    return tuple(found)


def validate_dictionary_dependencies(dictionary: Mapping[str, Node]) -> None:
    items = tuple(dictionary.items())
    try:
        error = _cached_dictionary_dependency_error(items)
    except TypeError:
        # Parsed DSL nodes are hashable, but retain compatibility for callers
        # that construct custom Node values containing unhashable literals.
        error = _dictionary_dependency_error(items)
    if error:
        raise ConfigurationError(error)


@lru_cache(maxsize=4096)
def _cached_dictionary_dependency_error(items: tuple[tuple[str, Node], ...]) -> str | None:
    return _dictionary_dependency_error(items)


def _dictionary_dependency_error(items: tuple[tuple[str, Node], ...]) -> str | None:
    dictionary = dict(items)
    dependencies = {name: variables_in(node) for name, node in items}
    for name, names in dependencies.items():
        undefined = names - dictionary.keys()
        if undefined:
            return f"变量 {name} 引用了未定义变量：{', '.join(sorted(undefined))}"
    visiting: list[str] = []
    visited: set[str] = set()

    def visit(name: str) -> str | None:
        if name in visiting:
            cycle = visiting[visiting.index(name):] + [name]
            return "取值字典存在循环引用：" + " → ".join(cycle)
        if name in visited:
            return None
        visiting.append(name)
        for dependency in dependencies[name]:
            error = visit(dependency)
            if error:
                return error
        visiting.pop()
        visited.add(name)
        return None

    for name in dependencies:
        error = visit(name)
        if error:
            return error
    return None


def static_kind(node: Node) -> str | None:
    if isinstance(node, Literal):
        return node.kind
    if isinstance(node, TupleLiteral):
        return "tuple"
    if isinstance(node, Unary):
        return "boolean" if node.operator == "NOT" else "number"
    if isinstance(node, Binary):
        if node.operator in {"AND", "OR", *_COMPARISONS}:
            return "boolean"
        return "number"
    if isinstance(node, Call):
        if node.name in {"V_CUR", "V_PRE", "V_EXT"}:
            return None
        if node.name in {"MIN", "MAX", "ABS", "ROUND", "SUM"}:
            return "number"
        if node.name in {"ISBLANK", "ISNUMBER"}:
            return "boolean"
        if node.name == "IF":
            yes, no = static_kind(node.arguments[1]), static_kind(node.arguments[2])
            return yes if yes == no else None
        if node.name in {"AND", "OR", "NOT"}:
            return "boolean"
    return None


def validate_static_types(node: Node, *, allow_string_literal: bool = False) -> None:
    """Reject statically obvious DSL type errors before a rule reads any source."""
    if isinstance(node, Literal):
        if node.kind == "string" and not allow_string_literal:
            raise ConfigurationError("字符串常量只能用于 V 函数参数或 IF 的结果分支")
        return
    if isinstance(node, (Variable, TupleLiteral)):
        return
    if isinstance(node, Unary):
        validate_static_types(node.operand)
        expected = "boolean" if node.operator == "NOT" else "number"
        actual = static_kind(node.operand)
        if actual is not None and actual != expected:
            raise ConfigurationError(f"{node.operator} 运算要求 {expected}，实际是 {actual}")
        return
    if isinstance(node, Binary):
        validate_static_types(node.left)
        validate_static_types(node.right)
        left_kind, right_kind = static_kind(node.left), static_kind(node.right)
        if node.operator in {"AND", "OR"}:
            for kind in (left_kind, right_kind):
                if kind is not None and kind != "boolean":
                    raise ConfigurationError(f"{node.operator} 运算只接受布尔值")
        elif node.operator in {"+", "-", "*", "/"}:
            for kind in (left_kind, right_kind):
                if kind is not None and kind != "number":
                    raise ConfigurationError(f"{node.operator} 运算只接受数值")
        elif left_kind is not None and right_kind is not None and left_kind != right_kind:
            raise ConfigurationError(f"比较两侧类型不一致：{left_kind} / {right_kind}")
        elif node.operator in {"<", "<=", ">", ">="} and left_kind not in {None, "number"}:
            raise ConfigurationError("字符串和布尔值只支持相等比较")
        return
    if isinstance(node, Call):
        validate_call(node)
        if node.name in {"V_CUR", "V_PRE", "V_EXT"}:
            return
        if node.name == "IF":
            validate_static_types(node.arguments[0])
            condition_kind = static_kind(node.arguments[0])
            if condition_kind is not None and condition_kind != "boolean":
                raise ConfigurationError("IF 的第一个参数必须是布尔值")
            validate_static_types(node.arguments[1], allow_string_literal=True)
            validate_static_types(node.arguments[2], allow_string_literal=True)
            yes_kind, no_kind = static_kind(node.arguments[1]), static_kind(node.arguments[2])
            if yes_kind is not None and no_kind is not None and yes_kind != no_kind:
                raise ConfigurationError(f"IF 的两个结果分支类型不一致：{yes_kind} / {no_kind}")
            return
        if node.name in {"ISBLANK", "ISNUMBER"}:
            validate_static_types(node.arguments[0], allow_string_literal=True)
            return
        expected = "boolean" if node.name in {"AND", "OR", "NOT"} else "number"
        for argument in node.arguments:
            validate_static_types(argument)
            actual = static_kind(argument)
            if actual is not None and actual != expected:
                raise ConfigurationError(f"{node.name} 参数要求 {expected}，实际是 {actual}")
        return
    raise ConfigurationError(f"未知表达式节点：{type(node).__name__}")


def _merge_evidence(*values: EvalValue) -> tuple[Any, ...]:
    merged: list[Any] = []
    seen: set[tuple[Any, ...]] = set()
    for value in values:
        for item in value.evidence:
            key = (getattr(item, "source_id", ""), getattr(item, "source_sha256", ""),
                   getattr(item, "sheet", ""), getattr(item, "cell", ""))
            if key not in seen:
                seen.add(key)
                merged.append(item)
    return tuple(merged)


def _require_number(value: EvalValue, function: str = "运算") -> Decimal | float:
    # Excel coerces a blank numeric cell to zero in arithmetic (for example,
    # =A1/10000 evaluates to 0 when A1 is empty). Keep the original blank kind
    # for ISBLANK(), but apply that coercion at numeric operator boundaries.
    if value.kind == "blank":
        return 0.0 if value.numeric_mode == "excel_binary64" else Decimal(0)
    if value.kind != "number" or isinstance(value.value, bool):
        raise EvaluationError(f"TYPE_ERROR：{function}需要数值")
    if value.numeric_mode == "excel_binary64":
        return float(value.value)
    return Decimal(value.value)


def _with_numeric_mode(value: EvalValue, numeric_mode: str) -> EvalValue:
    if numeric_mode != "excel_binary64" or value.kind not in {"number", "blank"}:
        return value
    raw = excel_round15(float(value.value)) if value.kind == "number" else None
    return replace(value, value=raw, numeric_mode=numeric_mode)


def _require_boolean(value: EvalValue, function: str = "逻辑运算") -> bool:
    if value.kind != "boolean" or not isinstance(value.value, bool):
        raise EvaluationError(f"TYPE_ERROR：{function}只接受布尔值")
    return value.value


def _combine_units(left: str | None, right: str | None, operator: str) -> str | None:
    if operator == "*":
        if left is None:
            return right
        if right is None:
            return left
        return f"{left}·{right}"
    if operator == "/":
        if right is None:
            return left
        if left == right:
            return None
        if left is None:
            return f"1/{right}"
        return f"{left}/{right}"
    return left


def _combine_precision(*values: EvalValue) -> int | None:
    precision_values = [value.precision for value in values if value.precision is not None]
    return max(precision_values) if precision_values else None


def _numeric_comparison_values(left: EvalValue, right: EvalValue) -> tuple[Decimal | float, Decimal | float]:
    left_number, right_number = _require_number(left, "比较"), _require_number(right, "比较")
    if left.unit == right.unit:
        pass
    elif left.unit is None and left.is_literal:
        pass
    elif right.unit is None and right.is_literal:
        pass
    else:
        raise EvaluationError(f"UNIT_MISMATCH：不能比较 {left.unit or '无单位'} 与 {right.unit or '无单位'}")
    if "excel_binary64" in {left.numeric_mode, right.numeric_mode}:
        return excel_round15(float(left_number)), excel_round15(float(right_number))
    return left_number, right_number


def _compare(operator: str, left: EvalValue, right: EvalValue) -> bool:
    if left.kind != right.kind:
        if left.kind in {"number", "blank"} and right.kind in {"number", "blank"}:
            pass
        else:
            raise EvaluationError(f"TYPE_ERROR：不能比较 {left.kind} 与 {right.kind}")
    if left.kind in {"number", "blank"}:
        lvalue, rvalue = _numeric_comparison_values(left, right)
    elif operator in {"=", "==", "<>", "!="}:
        if left.kind not in {"string", "boolean"}:
            raise EvaluationError("TYPE_ERROR：此类型只支持相等比较")
        lvalue, rvalue = left.value, right.value
    else:
        raise EvaluationError("TYPE_ERROR：字符串和布尔值只支持相等比较")
    result = {
        "=": lambda: lvalue == rvalue,
        "==": lambda: lvalue == rvalue,
        "<>": lambda: lvalue != rvalue,
        "!=": lambda: lvalue != rvalue,
        "<": lambda: lvalue < rvalue,
        "<=": lambda: lvalue <= rvalue,
        ">": lambda: lvalue > rvalue,
        ">=": lambda: lvalue >= rvalue,
    }[operator]()
    return bool(result)


def evaluate(node: Node, *, resolve_variable: Callable[[str], EvalValue],
             call_value: Callable[[str, tuple[Any, ...], Mapping[str, str]], EvalValue],
             numeric_mode: str = "decimal") -> EvalValue:
    if isinstance(node, Literal):
        if node.kind == "number":
            return EvalValue.number(node.value, is_literal=True, numeric_mode=numeric_mode)
        if node.kind == "string":
            return EvalValue.string(node.value, is_literal=True)
        return EvalValue.boolean(node.value, is_literal=True)
    if isinstance(node, Variable):
        return _with_numeric_mode(resolve_variable(node.name), numeric_mode)
    if isinstance(node, TupleLiteral):
        return EvalValue(tuple(node.values), "tuple", is_literal=True)
    if isinstance(node, Unary):
        value = evaluate(node.operand, resolve_variable=resolve_variable, call_value=call_value, numeric_mode=numeric_mode)
        if node.operator == "NOT":
            return EvalValue.boolean(not _require_boolean(value, "NOT"), evidence=value.evidence)
        number = _require_number(value, node.operator)
        return EvalValue.number(number if node.operator == "+" else -number,
                                unit=value.unit, metric_type=value.metric_type, evidence=value.evidence,
                                precision=value.precision, numeric_mode=value.numeric_mode)
    if isinstance(node, Binary):
        left = evaluate(node.left, resolve_variable=resolve_variable, call_value=call_value, numeric_mode=numeric_mode)
        if node.operator == "AND":
            if not _require_boolean(left, "AND"):
                return EvalValue.boolean(False, evidence=left.evidence)
            right = evaluate(node.right, resolve_variable=resolve_variable, call_value=call_value, numeric_mode=numeric_mode)
            return EvalValue.boolean(_require_boolean(right, "AND"), evidence=_merge_evidence(left, right))
        if node.operator == "OR":
            if _require_boolean(left, "OR"):
                return EvalValue.boolean(True, evidence=left.evidence)
            right = evaluate(node.right, resolve_variable=resolve_variable, call_value=call_value, numeric_mode=numeric_mode)
            return EvalValue.boolean(_require_boolean(right, "OR"), evidence=_merge_evidence(left, right))
        right = evaluate(node.right, resolve_variable=resolve_variable, call_value=call_value, numeric_mode=numeric_mode)
        if node.operator in _COMPARISONS:
            return EvalValue.boolean(_compare(node.operator, left, right), evidence=_merge_evidence(left, right))
        lnumber, rnumber = _require_number(left, node.operator), _require_number(right, node.operator)
        if node.operator in {"+", "-"}:
            left_is_unitless_literal = left.unit is None and left.is_literal
            right_is_unitless_literal = right.unit is None and right.is_literal
            if left.unit != right.unit and not left_is_unitless_literal and not right_is_unitless_literal:
                raise EvaluationError(f"UNIT_MISMATCH：不能进行 {node.operator} 运算：{left.unit or '无单位'} / {right.unit or '无单位'}")
            value = lnumber + rnumber if node.operator == "+" else lnumber - rnumber
            result_mode = "excel_binary64" if "excel_binary64" in {left.numeric_mode, right.numeric_mode} else "decimal"
            return EvalValue.number(value, unit=left.unit or right.unit,
                                    metric_type=left.metric_type or right.metric_type,
                                    evidence=_merge_evidence(left, right),
                                    precision=_combine_precision(left, right), numeric_mode=result_mode)
        if node.operator == "*":
            result_mode = "excel_binary64" if "excel_binary64" in {left.numeric_mode, right.numeric_mode} else "decimal"
            return EvalValue.number(lnumber * rnumber, unit=_combine_units(left.unit, right.unit, "*"),
                                    metric_type="derived", evidence=_merge_evidence(left, right),
                                    precision=_combine_precision(left, right), numeric_mode=result_mode)
        if node.operator == "/":
            if rnumber == 0:
                raise EvaluationError("DIVISION_BY_ZERO：除数为零")
            result_mode = "excel_binary64" if "excel_binary64" in {left.numeric_mode, right.numeric_mode} else "decimal"
            return EvalValue.number(lnumber / rnumber, unit=_combine_units(left.unit, right.unit, "/"),
                                    metric_type="derived", evidence=_merge_evidence(left, right),
                                    precision=_combine_precision(left, right), numeric_mode=result_mode)
    if isinstance(node, Call):
        name = node.name
        if name == "IF":
            condition = evaluate(node.arguments[0], resolve_variable=resolve_variable, call_value=call_value, numeric_mode=numeric_mode)
            branch = node.arguments[1] if _require_boolean(condition, "IF") else node.arguments[2]
            result = evaluate(branch, resolve_variable=resolve_variable, call_value=call_value, numeric_mode=numeric_mode)
            return EvalValue(result.value, result.kind, result.unit, result.metric_type,
                             _merge_evidence(condition, result), result.is_literal, result.precision, result.numeric_mode)
        if name == "AND":
            evaluated: list[EvalValue] = []
            for arg in node.arguments:
                value = evaluate(arg, resolve_variable=resolve_variable, call_value=call_value, numeric_mode=numeric_mode)
                if not _require_boolean(value, "AND"):
                    return EvalValue.boolean(False, evidence=_merge_evidence(*evaluated, value))
                evaluated.append(value)
            return EvalValue.boolean(True, evidence=_merge_evidence(*evaluated))
        if name == "OR":
            evaluated = []
            for arg in node.arguments:
                value = evaluate(arg, resolve_variable=resolve_variable, call_value=call_value, numeric_mode=numeric_mode)
                if _require_boolean(value, "OR"):
                    evaluated.append(value)
                    return EvalValue.boolean(True, evidence=_merge_evidence(*evaluated))
                evaluated.append(value)
            return EvalValue.boolean(False, evidence=_merge_evidence(*evaluated))
        if name == "NOT":
            value = evaluate(node.arguments[0], resolve_variable=resolve_variable, call_value=call_value, numeric_mode=numeric_mode)
            return EvalValue.boolean(not _require_boolean(value, "NOT"), evidence=value.evidence)
        if name in {"V_CUR", "V_PRE", "V_EXT"}:
            args: list[Any] = []
            for arg in node.arguments:
                if isinstance(arg, Literal):
                    args.append(arg.value)
                elif isinstance(arg, TupleLiteral):
                    args.append(arg.values)
                else:
                    raise ConfigurationError(f"函数 {name} 的定位参数必须是常量")
            return _with_numeric_mode(call_value(name, tuple(args), dict(node.keywords)), numeric_mode)
        values = [evaluate(arg, resolve_variable=resolve_variable, call_value=call_value, numeric_mode=numeric_mode)
                  for arg in node.arguments]
        if name == "ISBLANK":
            return EvalValue.boolean(values[0].kind == "blank", evidence=values[0].evidence)
        if name == "ISNUMBER":
            return EvalValue.boolean(values[0].kind == "number", evidence=values[0].evidence)
        if name == "SUM":
            numeric = [value for value in values if value.kind == "number"]
            if any(value.kind not in {"number", "blank", "string"} for value in values):
                raise EvaluationError("TYPE_ERROR：SUM 只接受数值、文本或空值")
            measured = [value for value in values if value.unit and not value.is_literal]
            unit_source = measured[0] if measured else (numeric[0] if numeric else None)
            unit = unit_source.unit if unit_source else None
            for value in values:
                if value is unit_source:
                    continue
                if value.kind not in {"number", "blank"}:
                    continue
                if value.unit != unit and not (value.unit is None and value.is_literal):
                    raise EvaluationError("UNIT_MISMATCH：SUM 的参数单位不一致")
            result_mode = "excel_binary64" if any(value.numeric_mode == "excel_binary64" for value in values) else "decimal"
            if result_mode == "excel_binary64":
                total = sum((float(value.value) for value in numeric), 0.0)
            else:
                total = sum((Decimal(value.value) for value in numeric), Decimal(0))
            return EvalValue.number(total, unit=unit,
                                    metric_type=(numeric[0].metric_type if numeric else unit_source.metric_type if unit_source else ""),
                                    evidence=_merge_evidence(*values),
                                    precision=_combine_precision(*values), numeric_mode=result_mode)
        if name == "ABS":
            value = values[0]
            return EvalValue.number(abs(_require_number(value, "ABS")), unit=value.unit,
                                    metric_type=value.metric_type, evidence=value.evidence,
                                    precision=value.precision, numeric_mode=value.numeric_mode)
        if name == "ROUND":
            value, digits_value = values
            digits_number = Decimal(str(_require_number(digits_value, "ROUND")))
            if digits_number != digits_number.to_integral_value():
                raise EvaluationError("TYPE_ERROR：ROUND 的 digits 必须是整数")
            digits = int(digits_number)
            if abs(digits) > 12:
                raise EvaluationError("ARGUMENT_ERROR：ROUND 的 digits 范围为 -12 到 12")
            number = Decimal(str(_require_number(value, "ROUND")))
            quantum = Decimal(1).scaleb(-digits)
            return EvalValue.number(number.quantize(quantum, rounding=ROUND_HALF_UP),
                                    unit=value.unit, metric_type=value.metric_type, evidence=value.evidence,
                                    precision=max(digits, 0), numeric_mode=value.numeric_mode)
        if name in {"MIN", "MAX"}:
            numbers = [_require_number(value, name) for value in values]
            unit = values[0].unit
            for value in values[1:]:
                if value.unit != unit and not (value.unit is None and value.is_literal) and not (unit is None and values[0].is_literal):
                    raise EvaluationError(f"UNIT_MISMATCH：{name} 的参数单位不一致")
            picked = min(range(len(numbers)), key=numbers.__getitem__) if name == "MIN" else max(range(len(numbers)), key=numbers.__getitem__)
            chosen = values[picked]
            return EvalValue.number(numbers[picked], unit=chosen.unit or unit,
                                    metric_type=chosen.metric_type, evidence=_merge_evidence(*values),
                                    precision=_combine_precision(*values), numeric_mode=chosen.numeric_mode)
        raise ConfigurationError(f"不支持的函数：{name}")
    raise ConfigurationError(f"未知表达式节点：{type(node).__name__}")


def evaluate_expression(text: str, *, variables: Mapping[str, EvalValue] | None = None,
                        call_value: Callable[[str, tuple[Any, ...], Mapping[str, str]], EvalValue] | None = None,
                        numeric_mode: str = "decimal") -> EvalValue:
    dictionary = variables or {}
    node = parse_expression(text)

    def resolve(name: str) -> EvalValue:
        if name not in dictionary:
            raise EvaluationError(f"UNDEFINED_VARIABLE：未定义变量 {name}")
        return dictionary[name]

    def no_call(name: str, args: tuple[Any, ...], keywords: Mapping[str, str]) -> EvalValue:
        if call_value is None:
            raise EvaluationError(f"未配置取值函数：{name}")
        return call_value(name, args, keywords)

    return evaluate(node, resolve_variable=resolve, call_value=no_call, numeric_mode=numeric_mode)
