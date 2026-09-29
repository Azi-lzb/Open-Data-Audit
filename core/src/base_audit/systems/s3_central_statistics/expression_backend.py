"""表达式执行 2×2 接口：SBE/LAE × PYTHON/OFFICE（V3 合并计划 §五之二）。

两个正交维度：
- 表达式处理方式：SBE（字符串代入求值）/ LAE（惰性 AST 求值）；
- 求值后端：PYTHON（纯 Python evaluator）/ OFFICE（办公套件自动路由）。

正式路径默认仍是 SBE + PYTHON（现行 complex_rule_engine 行为，零改动）；
LAE + PYTHON 使用 expression_lae（纯 Python、AST 缓存、短路求值）。
OFFICE 后端本版本未实现：选择即明确失败，禁止静默回退 Python。
"""

from __future__ import annotations

from dataclasses import dataclass

MODE_SBE = "SBE"
MODE_LAE = "LAE"
BACKEND_PYTHON = "PYTHON"
BACKEND_OFFICE = "OFFICE"

# 运行参数（配置簿「运行参数」表；缺省 = 现行正式组合）。
MODE_PARAM = "表达式处理方式"
BACKEND_PARAM = "表达式求值方式"

_VALID_MODES = {MODE_SBE, MODE_LAE}
_VALID_BACKENDS = {BACKEND_PYTHON, BACKEND_OFFICE}


class ExpressionBackendError(RuntimeError):
    """表达式执行配置不可用（明确失败，不静默回退）。"""


@dataclass
class ExpressionRunResult:
    """统一结果：hit=触发；process_extra 追加到计算过程；status 记录四态。"""

    hit: bool
    substituted: str = ""       # SBE：代入后的公式文本（计算过程用）
    status: str = "OK"          # OK / UNSUPPORTED / ERROR / OFFICE_UNAVAILABLE
    reason: str = ""


def resolve_expression_settings(params: dict[str, str] | None) -> tuple[str, str]:
    """从运行参数解析 (mode, backend)；非法值明确报错。"""
    mode = str((params or {}).get(MODE_PARAM, "") or MODE_SBE).strip().upper()
    backend = str((params or {}).get(BACKEND_PARAM, "") or BACKEND_PYTHON).strip().upper()
    if mode not in _VALID_MODES:
        raise ExpressionBackendError(
            f"{MODE_PARAM} 仅支持 {'/'.join(sorted(_VALID_MODES))}，当前：{mode}")
    if backend not in _VALID_BACKENDS:
        raise ExpressionBackendError(
            f"{BACKEND_PARAM} 仅支持 {'/'.join(sorted(_VALID_BACKENDS))}，当前：{backend}")
    return mode, backend


# ---------------------------------------------------------------------------
# LAE → Office 公式编译（LAE+Office：绑定引用后编译为 Excel/WPS 公式）
# ---------------------------------------------------------------------------

def _office_number(value: float) -> str:
    number = float(value)
    if number != number or number in (float("inf"), float("-inf")):
        raise ExpressionBackendError(f"无法把 {value!r} 编译为 Office 数值字面量")
    if number == int(number) and abs(number) < 1e15:
        return str(int(number))
    return repr(number)


def compile_ast_to_office(node: tuple, context) -> str:
    """把 V2 AST 编译为 Office 公式文本（引用经 ExpressionContext 绑定为字面量）。

    EXISTS/MISSING/ISBLANK/RAW 在 Excel/WPS 中不存在，按 V2 语义在编译期折叠为
    字面量（复用 expression_lae 同一实现，保证折叠结果与 LAE+Python 一致）；
    其余节点一一映射为 Excel 语法并全括号化，最终语义由 Office provider 决定
    ——这正是 LAE+Office 作为「Parser/Binding/编译链验证路径」的定位。
    IFS 编译为嵌套 IF（兼容 Excel 2016 及 WPS）。
    """
    from .expression_lae import evaluate_ast, reference_value

    kind = node[0]
    if kind == "num":
        return _office_number(node[1])
    if kind == "bool":
        return "TRUE" if node[1] else "FALSE"
    if kind == "ref":
        return _office_number(reference_value(node[1], node[2], context))
    if kind == "date":
        serial = _date_serial(context, node[1])
        return _office_number(serial)
    if kind == "thd":
        return _office_number(context.thd)
    if kind == "un":
        return f"({node[1]}{compile_ast_to_office(node[2], context)})"
    if kind == "bin":
        return f"({compile_ast_to_office(node[2], context)}{node[1]}{compile_ast_to_office(node[3], context)})"
    if kind == "cmp":
        return f"({compile_ast_to_office(node[2], context)}{node[1]}{compile_ast_to_office(node[3], context)})"
    if kind == "call":
        name, args = node[1], node[2]
        if name in ("EXISTS", "MISSING", "ISBLANK", "RAW"):
            folded = evaluate_ast(node, context)
            if isinstance(folded, bool):
                return "TRUE" if folded else "FALSE"
            return _office_number(folded)
        if name == "IF":
            parts = [compile_ast_to_office(arg, context) for arg in args]
            if len(parts) == 2:
                parts.append("FALSE")
            return f"IF({','.join(parts)})"
        if name == "IFS":
            # 兜底编译为 NA()（#N/A），与 Python 侧 IFS 无匹配语义一致；
            # 用 FALSE 会让配置缺陷静默漏报（Excel 实测 IF(...) 两参= FALSE、
            # NA() = #N/A，两者语义不同，此处必须用 NA()）。
            if not args or len(args) % 2:
                raise ExpressionBackendError("IFS 需要成对的「条件, 结果」参数")
            compiled = "NA()"
            for i in range(len(args) - 2, -1, -2):
                compiled = (
                    f"IF({compile_ast_to_office(args[i], context)},"
                    f"{compile_ast_to_office(args[i + 1], context)},{compiled})"
                )
            return compiled
        if name == "ROUND" and len(args) == 1:
            return f"ROUND({compile_ast_to_office(args[0], context)},0)"
        return f"{name}({','.join(compile_ast_to_office(arg, context) for arg in args)})"
    raise ExpressionBackendError(f"未知 AST 节点：{kind}")


def _date_serial(context, name: str) -> float:
    from datetime import date

    value = context.date_of("cur" if name == "本期日期" else "prev")
    if value is None:
        raise ExpressionBackendError(f"上下文变量「{name}」无法解析为日期")
    return float((value - date(1899, 12, 30)).days)


def evaluate_expression_rule(
    expression: str,
    *,
    trigger_inverted: bool,
    mode: str = MODE_SBE,
    backend: str = BACKEND_PYTHON,
    thd: float = 0.01,
    # SBE 输入
    substitute=None,            # callable(expression) -> 代入文本
    evaluate_text=None,         # callable(代入文本) -> 值（raise ExpressionError）
    # LAE 输入
    context=None,               # expression_lae.ExpressionContext
    # OFFICE 输入
    office_evaluator=None,      # office_eval.OfficeEvaluationAdapter
) -> ExpressionRunResult:
    """按 2×2 组合求值一条规则表达式并应用触发方式。

    SBE：先由调用方提供 substitute/evaluate_text（保持现行代入口径），
    再交 Python evaluator 或 Office；LAE：经 ExpressionContext 解析后由
    Python 惰性求值，或编译为 Office 公式交给 provider。
    UNSUPPORTED/ERROR 不静默当 FALSE（如实回传 status，由调用方按口径处理）；
    OFFICE 必须注入任务级 OfficeEvaluationAdapter（provider 一次解析、
    全程固定），缺失即明确失败。
    """
    if mode not in _VALID_MODES:
        raise ExpressionBackendError(f"未知表达式处理方式：{mode}")
    if backend not in _VALID_BACKENDS:
        raise ExpressionBackendError(f"未知求值后端：{backend}")
    if backend == BACKEND_OFFICE and office_evaluator is None:
        raise ExpressionBackendError(
            "表达式求值方式=OFFICE：未提供任务级 Office 求值适配器"
            "（provider 须在任务启动时一次解析并固定）；不静默回退 Python。")

    if mode == MODE_SBE:
        from .expression_parser import ExpressionError

        substituted = substitute(expression)
        if backend == BACKEND_OFFICE:
            outcome = office_evaluator.evaluate(substituted)
        else:
            try:
                value = evaluate_text(substituted)
            except ExpressionError as exc:
                # 现行语义：求值错误按触发处理并在计算过程标注（VBA 对齐）。
                return ExpressionRunResult(
                    hit=True, substituted=substituted, status="ERROR", reason=str(exc))
            if isinstance(value, bool):
                hit = (not trigger_inverted and value) or (trigger_inverted and not value)
                return ExpressionRunResult(hit=hit, substituted=substituted)
            # 非逻辑值：按现行 VBA 语义输出提示（拼接后为「校验结果并不是逻辑值:…--。」）。
            return ExpressionRunResult(
                hit=True, substituted=substituted, status="OK",
                reason=f"结果并不是逻辑值:{value}")
        if outcome.status == "ERROR":
            # VBA IsError(result) 口径：命中并标注「校验时发生错误--。」。
            return ExpressionRunResult(
                hit=True, substituted=substituted, status="ERROR", reason=outcome.reason)
        if isinstance(outcome.value, bool):
            hit = (not trigger_inverted and outcome.value) or (trigger_inverted and not outcome.value)
            return ExpressionRunResult(hit=hit, substituted=substituted)
        return ExpressionRunResult(
            hit=True, substituted=substituted, status="OK",
            reason=f"结果并不是逻辑值:{outcome.value}")

    # ---- LAE ----
    from .expression_lae import (
        ExpressionError,
        UnsupportedFunctionError,
        evaluate_ast,
        parse_expression,
    )

    try:
        ast = parse_expression(expression)
    except UnsupportedFunctionError as exc:
        return ExpressionRunResult(status="UNSUPPORTED", reason=str(exc))
    except ExpressionError as exc:
        return ExpressionRunResult(status="ERROR", reason=f"解析失败：{exc}")

    if backend == BACKEND_OFFICE:
        try:
            formula = compile_ast_to_office(ast, context)
        except UnsupportedFunctionError as exc:
            return ExpressionRunResult(status="UNSUPPORTED", reason=str(exc))
        except ExpressionError as exc:
            return ExpressionRunResult(status="ERROR", reason=f"绑定失败：{exc}")
        outcome = office_evaluator.evaluate(formula)
        if outcome.status == "ERROR":
            return ExpressionRunResult(status="ERROR", reason=outcome.reason)
        truth = bool(outcome.value)
        hit = truth if not trigger_inverted else not truth
        return ExpressionRunResult(hit=hit, substituted=formula)

    want = "表达式不成立" if trigger_inverted else "表达式成立"
    try:
        value = evaluate_ast(ast, context)
    except UnsupportedFunctionError as exc:
        return ExpressionRunResult(status="UNSUPPORTED", reason=str(exc))
    except ExpressionError as exc:
        return ExpressionRunResult(status="ERROR", reason=f"求值失败：{exc}")
    from .expression_lae import _bool

    truth = _bool(value)
    hit = truth if want == "表达式成立" else (not truth)
    return ExpressionRunResult(hit=hit, status="OK")
