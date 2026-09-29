"""Office 表达式后端：结果解释、LAE→Office 编译器、2×2 OFFICE 分支、装载归一。"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from base_audit.systems.s3_central_statistics.config import CentralConfig
from base_audit.systems.s3_central_statistics.expression_backend import (
    BACKEND_OFFICE,
    BACKEND_PYTHON,
    ExpressionBackendError,
    MODE_LAE,
    MODE_SBE,
    compile_ast_to_office,
    evaluate_expression_rule,
    resolve_expression_settings,
)
from base_audit.systems.s3_central_statistics.office_eval import interpret_office_value
from base_audit.systems.s3_central_statistics.expression_lae import ExpressionContext, parse_expression
from base_audit.systems.s3_central_statistics.models import CentralRecord, ComparisonRow


@dataclass
class _FakeOutcome:
    status: str
    value: object = None
    reason: str = ""


class _FakeOfficeEvaluator:
    """记录编译/代入产物并回放预设结果的 Office 适配器替身。"""

    def __init__(self, outcomes: list) -> None:
        self.outcomes = list(outcomes)
        self.received: list[str] = []

    def evaluate(self, formula: str) -> _FakeOutcome:
        self.received.append(formula)
        return _FakeOutcome(*self.outcomes.pop(0))


def _ctx_with(value: float, *, previous: float | None = None) -> ExpressionContext:
    def record(v):
        return CentralRecord(
            biz_class="人民币", record_date="2026-08-31", org_code="6k0i", org_name="测试",
            region_code="4400000", region_name="广东省", order_code="1", indicator="A",
            indicator_name="A", data_attr="余额", currency="人民币",
            frequency="月", batch="1", value=v,
        )

    row = ComparisonRow(record=record(value), prev_value=None, change=0.0, ratio=0.0)
    from base_audit.systems.s3_central_statistics.complex_rule_engine import _target_key

    current_index = {_target_key("[,,,A,余额,人民币,月,1]", row): record(value)}
    previous_index = {}
    if previous is not None:
        previous_index[_target_key("{,,,A,余额,人民币,月,1}", row)] = record(previous)
    return ExpressionContext(
        current_date="2026-08-31", previous_date="2026-07-31",
        current_index=current_index, previous_index=previous_index,
        row=row, unit_factor=1.0,
    )


# ---------------------------------------------------------------------------
# interpret_office_value
# ---------------------------------------------------------------------------

def test_interpret_office_value_shapes() -> None:
    assert interpret_office_value(True) == ("OK", True, "")
    assert interpret_office_value(2.5)[0] == "OK"
    status, value, _ = interpret_office_value(-2146826281)
    assert status == "ERROR" and value == -2146826281
    assert interpret_office_value("文本")[0] == "OK"


# ---------------------------------------------------------------------------
# compile_ast_to_office
# ---------------------------------------------------------------------------

def test_compile_binds_references_to_literals() -> None:
    ctx = _ctx_with(123.45, previous=7.0)
    ast = parse_expression("[,,,A,余额,人民币,月,1]-{,,,A,余额,人民币,月,1}>1")
    formula = compile_ast_to_office(ast, ctx)
    assert formula == "((123.45-7)>1)"


def test_compile_ifs_to_nested_if_and_round_single_arg() -> None:
    ctx = _ctx_with(1.0)
    ast = parse_expression("IFS(1<0,1,2>0,3)")
    # IFS 兜底编译为 NA()（#N/A）：与 Python 侧无匹配语义一致，严禁 FALSE 静默漏报。
    assert compile_ast_to_office(ast, ctx) == "IF((1<0),1,IF((2>0),3,NA()))"
    ast = parse_expression("ROUND(ABS(-2.567),2)")
    assert compile_ast_to_office(ast, ctx) == "ROUND(ABS((-2.567)),2)"
    ast = parse_expression("ROUND(2.5)")
    assert compile_ast_to_office(ast, ctx) == "ROUND(2.5,0)"


def test_compile_folds_raw_and_exists() -> None:
    ctx = _ctx_with(0.0)   # 当前值为 0：普通引用=哨兵 0.01，RAW=真 0
    ast = parse_expression("RAW([,,,A,余额,人民币,月,1])=0")
    assert compile_ast_to_office(ast, ctx) == "(0=0)"
    ast = parse_expression("EXISTS({,,,A,余额,人民币,月,1})")
    assert compile_ast_to_office(ast, ctx) == "FALSE"


def test_compile_date_variable_to_excel_serial() -> None:
    ctx = _ctx_with(1.0)
    ast = parse_expression("YEAR(本期日期)=2026")
    assert compile_ast_to_office(ast, ctx) == "(YEAR(46265)=2026)"


# ---------------------------------------------------------------------------
# evaluate_expression_rule：OFFICE 分支
# ---------------------------------------------------------------------------

def test_office_without_evaluator_fails_loudly() -> None:
    with pytest.raises(ExpressionBackendError):
        evaluate_expression_rule("1>0", trigger_inverted=False,
                                 mode=MODE_SBE, backend=BACKEND_OFFICE)


def test_sbe_office_bool_and_error_semantics() -> None:
    # 布尔真 → 不取反命中。
    fake = _FakeOfficeEvaluator([("OK", True, "")])
    result = evaluate_expression_rule(
        "1>0", trigger_inverted=False, mode=MODE_SBE, backend=BACKEND_OFFICE,
        substitute=lambda text: text, office_evaluator=fake)
    assert result.hit is True and fake.received == ["1>0"]
    # Excel 错误值 → 按 VBA IsError 口径命中并标注。
    fake = _FakeOfficeEvaluator([("ERROR", -2146826281, "Excel错误代码 -2146826281")])
    result = evaluate_expression_rule(
        "1/0", trigger_inverted=False, mode=MODE_SBE, backend=BACKEND_OFFICE,
        substitute=lambda text: text, office_evaluator=fake)
    assert result.hit is True and result.status == "ERROR"


def test_lae_office_compiles_and_hits() -> None:
    fake = _FakeOfficeEvaluator([("OK", False, "")])
    ctx = _ctx_with(5.0, previous=5.0)
    result = evaluate_expression_rule(
        "[,,,A,余额,人民币,月,1]={,,,A,余额,人民币,月,1}",
        trigger_inverted=False, mode=MODE_LAE, backend=BACKEND_OFFICE,
        context=ctx, office_evaluator=fake)
    assert fake.received == ["(5=5)"]
    assert result.hit is False


def test_settings_roundtrip() -> None:
    assert resolve_expression_settings({"表达式处理方式": "lae"}) == (MODE_LAE, BACKEND_PYTHON)
    with pytest.raises(ExpressionBackendError):
        resolve_expression_settings({"表达式求值方式": "LIBREOFFICE"})


# ---------------------------------------------------------------------------
# 装载期全角归一（VBA InitComplexRef_ClearRule 对齐）
# ---------------------------------------------------------------------------

def test_complex_rules_normalize_full_width_characters() -> None:
    from base_audit.systems.s3_central_statistics.complex_rule_engine import compile_complex_rules

    config = CentralConfig(path=Path("c.xlsx"), complex_rules=[{
        "规则编号": "CX测试", "校验表单": "A0", "校验描述": "全角",
        "校验规则": "（[,,,A,余额,人民币,月,1]）＞0".replace("＞", ">"),
        "取反标识": "", "Thd值(万元)": "", "禁用": "",
    }])
    index = compile_complex_rules(config)
    assert len(index["A"]) == 1
    rule = index["A"][0]
    assert rule.expression.startswith("([,,,A")


# ---------------------------------------------------------------------------
# Excel 精确函数库（语义以 Excel 16 / WPS 12 实测为准）
# ---------------------------------------------------------------------------

_ALIGN_CTX = ExpressionContext(current_date="2026-08-31", previous_date="2026-07-31")


def _eval(expression: str):
    from base_audit.systems.s3_central_statistics.expression_lae import evaluate_ast

    return evaluate_ast(parse_expression(expression), _ALIGN_CTX)


@pytest.mark.parametrize("expression,expected", [
    ("ROUND(2.5,0)", 3.0), ("ROUND(-2.5,0)", -3.0),
    ("ROUND(2.675,2)", 2.68), ("ROUND(25,-1)", 30.0),
    ("ROUNDUP(2.511,2)", 2.52), ("ROUNDUP(-2.511,2)", -2.52), ("ROUNDUP(25,-1)", 30.0),
    ("ROUNDDOWN(3.999,2)", 3.99), ("ROUNDDOWN(-3.999,2)", -3.99),
    ("TRUNC(-5.9)", -5.0), ("TRUNC(3.999,2)", 3.99),
    ("INT(-5.5)", -6.0),
    ("MOD(5.5,2)", 1.5), ("MOD(-5,3)", 1.0), ("MOD(5,-3)", -1.0), ("MOD(-5,-3)", -2.0),
    ("CEILING(2.5,1)", 3.0), ("CEILING(-2.5,1)", -2.0), ("CEILING(-2.5,-1)", -3.0),
    ("CEILING(5,0)", 0.0),
    ("FLOOR(2.5,1)", 2.0), ("FLOOR(-2.5,1)", -3.0), ("FLOOR(-2.5,-1)", -2.0),
    ("SIGN(-7)", -1.0), ("SIGN(0)", 0.0),
    ("POWER(2,10)", 1024.0),
    ("AVERAGE(1,2,3)", 2.0), ("COUNT(1,TRUE)", 2.0),
    ("IFERROR(5,99)", 5.0), ("ISERROR(1)", False),
])
def test_excel_exact_function_semantics(expression: str, expected: float) -> None:
    assert _eval(expression) == pytest.approx(expected)


@pytest.mark.parametrize("expression", [
    "MOD(5,0)", "CEILING(2.5,-1)", "FLOOR(5,0)", "SQRT(-1)",
    "POWER(-2,0.5)", "POWER(0,0)", "1E300*1E300", "(-2)^0.5",
])
def test_excel_error_semantics(expression: str) -> None:
    from base_audit.systems.s3_central_statistics.expression_lae import ExpressionError

    with pytest.raises(ExpressionError):
        _eval(expression)


def test_iferror_is_error_lazy() -> None:
    assert _eval("IFERROR(1/0,99)") == 99.0
    assert _eval("IFERROR(1/0)") == 0.0
    assert _eval("ISERROR(1/0)") is True


def test_new_functions_compile_through_generic_path() -> None:
    ctx = _ctx_with(2.0)
    assert compile_ast_to_office(parse_expression("ROUNDUP([,,,A,余额,人民币,月,1],1)"), ctx) \
        == "ROUNDUP(2,1)"
    assert compile_ast_to_office(parse_expression("IFERROR([,,,A,余额,人民币,月,1]/0,9)"), ctx) \
        == "IFERROR((2/0),9)"


# ---------------------------------------------------------------------------
# SBE+Python（expression_parser）：IF/IFS/IFERROR/ISERROR 与新函数库
# （与 LAE 共享 Excel 精确内核，语义同表复用）
# ---------------------------------------------------------------------------

def _eval_sbe(expression: str):
    from base_audit.systems.s3_central_statistics.expression_parser import evaluate_expression

    return evaluate_expression(expression)


def test_sbe_if_ifs_lazy_branches() -> None:
    assert _eval_sbe("IF(1>0,2,3)") == 2.0
    assert _eval_sbe("IF(1<0,2)") is False
    # 未选中的分支不得求值（除零不触发）。
    assert _eval_sbe("IF(1<0,1/0,99)") == 99.0
    assert _eval_sbe("IFS(1<0,1,1>0,2)") == 2.0
    # 无匹配 → #N/A（Excel 口径），由 test_ifs_no_match_is_na_not_false 专门覆盖。
    with pytest.raises(Exception):
        _eval_sbe("IFS(1<0,1,2<0,3)")


def test_sbe_true_false_are_booleans() -> None:
    assert _eval_sbe("TRUE") is True
    assert _eval_sbe("IF(TRUE,1,2)") == 1.0
    assert _eval_sbe("IF(FALSE,1,2)") == 2.0
    assert _eval_sbe("TRUE+1") == 2.0


@pytest.mark.parametrize("expression,expected", [
    ("ROUND(-2.5,0)", -3.0), ("ROUND(25,-1)", 30.0),
    ("ROUNDUP(-2.511,2)", -2.52), ("ROUNDDOWN(3.999,2)", 3.99),
    ("TRUNC(-5.9)", -5.0), ("INT(-5.5)", -6.0), ("MOD(-5,-3)", -2.0),
    ("CEILING(-2.5,1)", -2.0), ("FLOOR(-2.5,1)", -3.0), ("SIGN(-7)", -1.0),
    ("POWER(2,10)", 1024.0), ("AVERAGE(1,2,3)", 2.0), ("COUNT(1,TRUE)", 2.0),
])
def test_sbe_excel_exact_function_semantics(expression: str, expected: float) -> None:
    assert _eval_sbe(expression) == pytest.approx(expected)


def test_sbe_iferror_only_swallows_calculation_errors() -> None:
    from base_audit.systems.s3_central_statistics.expression_parser import ExpressionError

    assert _eval_sbe("IFERROR(1/0,99)") == 99.0
    assert _eval_sbe("IFERROR(1/0)") == 0.0
    assert _eval_sbe("ISERROR(1/0)") is True
    with pytest.raises(ExpressionError):
        _eval_sbe("IFERROR(未知名称+1,99)")


@pytest.mark.parametrize("expression", [
    "MOD(5,0)", "FLOOR(5,0)", "SQRT(-1)", "1E300*1E300", "(-2)^0.5",
])
def test_sbe_error_semantics(expression: str) -> None:
    # POWER(0,0)=1（Office 口径）不在此列。
    from base_audit.systems.s3_central_statistics.expression_parser import ExpressionError

    with pytest.raises(ExpressionError):
        _eval_sbe(expression)


# ---------------------------------------------------------------------------
# LibreOffice provider（UOS 路由；本机无 soffice 时验证解释与临时簿写出逻辑）
# ---------------------------------------------------------------------------

def test_interpret_calc_value_error_texts() -> None:
    from base_audit.systems.s3_central_statistics.office_eval import interpret_calc_value

    assert interpret_calc_value(True) == ("OK", True, "")
    assert interpret_calc_value(2.5) == ("OK", 2.5, "")
    status, value, _ = interpret_calc_value("#NUM!")
    assert status == "ERROR" and value == "#NUM!"
    assert interpret_calc_value("文本")[0] == "OK"


def test_libreoffice_provider_evaluates_via_scratch_recalc(tmp_path, monkeypatch) -> None:
    """无 soffice 环境下用假重算验证：公式写出→读缓存→解释 的完整链路。"""
    from base_audit.systems.s3_central_statistics import office_eval as module
    from openpyxl import Workbook, load_workbook

    def fake_recalculate(path):
        book = load_workbook(path)
        formula = book.active["A1"].value
        book.close()
        book = Workbook()
        sheet = book.active
        sheet["A1"] = 3.0 if "1+2" in formula else "#DIV/0!"
        book.save(path)
        book.close()
        return object()

    provider = module._LibreOfficeProvider.__new__(module._LibreOfficeProvider)
    provider.provider = module.PROVIDER_LIBREOFFICE
    provider._workbook_path = None
    provider._calculator = type("C", (), {"recalculate": staticmethod(fake_recalculate),
                                          "require_available": staticmethod(lambda: object())})()
    outcome = provider.evaluate("1+2")
    assert outcome.status == "OK" and outcome.value == 3.0 and outcome.provider == "LIBREOFFICE"
    outcome = provider.evaluate("1/0")
    assert outcome.status == "ERROR" and outcome.reason.startswith("Calc错误")
    provider.stop(None, None, None)
    assert provider._workbook_path is None or not provider._workbook_path.exists()


def test_expression_performance_counters() -> None:
    """性能观测：AST 缓存命中与 LAE 短路/跳过分支计数（V3 计划 §十）。"""
    from base_audit.systems.s3_central_statistics import expression_lae
    from base_audit.systems.s3_central_statistics.expression_lae import evaluate_ast

    expression_lae.reset_cache_stats()
    ctx = _ctx_with(5.0)
    ctx.stats = {}

    # AND 短路：第二参数（含除零）不求值 → 无除零错误。
    assert evaluate_ast(parse_expression("AND(0>1, 1/0)"), ctx) is False
    assert ctx.stats["and_short_circuit_count"] == 1
    assert ctx.stats["skipped_ast_node_count"] >= 1

    # OR 短路 + IF 跳过分支。
    assert evaluate_ast(parse_expression("OR(0>1, IF(1>2, 9, 8)>0)"), ctx) is True
    assert ctx.stats["or_short_circuit_count"] == 1
    assert ctx.stats["if_skipped_branch_count"] == 1
    assert ctx.stats["function_eval_count"] >= 3

    # AST 缓存：同一文本第二次解析命中。
    hits_before = expression_lae.cache_stats()["hits"]
    evaluate_ast(parse_expression("AND(0>1, 1/0)"), ctx)
    assert expression_lae.cache_stats()["hits"] == hits_before + 1


def test_libreoffice_provider_has_start_protocol() -> None:
    """UOS 验收 DEFECT-1 回归钉：_LibreOfficeProvider 必须实现 start() 协议。"""
    from base_audit.systems.s3_central_statistics import office_eval as module

    assert hasattr(module._LibreOfficeProvider, "start")
    provider = object.__new__(module._LibreOfficeProvider)
    provider.start()   # no-op，不得抛异常


def test_calc_formula_text_adds_xlfn_prefix() -> None:
    """UOS 验收 FINDING-1/2：仅 IFS 加 _xlfn. 前缀且大小写不敏感。

    UOS 26.8 实测口径：裸 IFS→#NAME?（需前缀）；_xlfn.IFERROR→#NAME?
    （不加前缀，裸 IFERROR 本可用）。
    """
    from base_audit.systems.s3_central_statistics.office_eval import _calc_formula_text

    assert _calc_formula_text("IFS(1<0,1,1>0,2)") == "_xlfn.IFS(1<0,1,1>0,2)"
    assert _calc_formula_text("ifs(1<0,1,1>0,2)") == "_xlfn.ifs(1<0,1,1>0,2)"
    assert _calc_formula_text("IFERROR(5,99)") == "IFERROR(5,99)"
    assert _calc_formula_text("AND(ABS(1-2)>0,1>0)") == "AND(ABS(1-2)>0,1>0)"
    # 已带前缀不重复叠加。
    assert _calc_formula_text("_xlfn.IFS(1,2)") == "_xlfn.IFS(1,2)"


def test_power_zero_zero_follows_office_semantics() -> None:
    """UOS 验收 DIFF-1 修正：POWER(0,0) 按 Excel Evaluate 实测口径报 #NUM!（错误）。

    Windows 复核（2026-09-18）：Excel Evaluate('power(0,0)')=#NUM!；
    LibreOffice 返回 1 属 LO 自身差异（F026，provider 语义项）。
    """
    from base_audit.systems.s3_central_statistics.expression_parser import ExpressionError

    with pytest.raises(ExpressionError):
        _eval("POWER(0,0)")


def test_libreoffice_provider_batch_single_recalc(tmp_path, monkeypatch) -> None:
    """LO 批量求值：N 条公式合并一次重算，顺序与结果一一对应。"""
    from base_audit.systems.s3_central_statistics import office_eval as module
    from openpyxl import Workbook, load_workbook

    calls = {"recalc": 0}

    def fake_recalculate(path):
        calls["recalc"] += 1
        book = load_workbook(path)
        sheet = book.active
        values = []
        for row in sheet.iter_rows(min_col=1, max_col=1, values_only=True):
            (text,) = row
            values.append(3.0 if "1+2" in str(text) else "#DIV/0!")
        book.close()
        book = Workbook()
        for index, value in enumerate(values, start=1):
            book.active.cell(row=index, column=1).value = value
        book.save(path)
        book.close()
        return object()

    provider = module._LibreOfficeProvider.__new__(module._LibreOfficeProvider)
    provider.provider = module.PROVIDER_LIBREOFFICE
    provider._workbook_path = None
    provider._calculator = type("C", (), {"recalculate": staticmethod(fake_recalculate),
                                          "require_available": staticmethod(lambda: object())})()
    outcomes = provider.evaluate_batch(["1+2", "1/0", "1+2"])
    assert calls["recalc"] == 1
    assert [o.status for o in outcomes] == ["OK", "ERROR", "OK"]
    assert outcomes[0].value == 3.0 and outcomes[1].value == "#DIV/0!"


def test_thd_supported_across_paths() -> None:
    """Thd 容差变量：SBE（variables）与 LAE（ctx.thd）同语义，Office 编译为字面量。"""
    from base_audit.systems.s3_central_statistics.expression_lae import evaluate_ast

    # LAE：ctx.thd 注入，四则与比较可用。
    ctx = _ctx_with(5.0)
    ctx.thd = 0.05
    ev = lambda expr: evaluate_ast(parse_expression(expr), ctx)
    assert ev("ABS(Thd - 0.05) < 0.001") is True
    assert ev("thd > 0") is True            # 大小写不敏感
    assert ev("Thd * 2") == 0.1
    # SBE：variables 传入（现行生产口径）。
    from base_audit.systems.s3_central_statistics.expression_parser import evaluate_expression

    assert evaluate_expression("Thd > 0", {"Thd": 0.05}) is True
    # LAE→Office：编译为字面量。
    assert compile_ast_to_office(parse_expression("Thd * 2"), ctx) == "(0.05*2)"


def test_chained_comparison_left_assoc() -> None:
    """链式比较左结合（Excel/VBA 口径）：(a>b)>0，布尔参与后续比较。"""
    ctx = _ctx_with(5.0)
    assert _eval("1>0>0") is True          # (True)>0
    assert _eval("0>1>0") is False         # (False)>0
    assert _eval("(5>3)>0") is True


def test_ifs_no_match_is_na_not_false() -> None:
    """A004 语义（业务定案）：IFS 无匹配 → #N/A（ERROR），严禁静默 FALSE。

    审核工具最贵的失败是漏报；#N/A 让配置缺陷在运行日志显形。
    四条路径（SBE/LAE × Python/Office）必须一致。
    """
    from base_audit.systems.s3_central_statistics.expression_parser import (
        ExcelNaError,
        evaluate_expression,
    )
    from base_audit.systems.s3_central_statistics.expression_lae import evaluate_ast

    ctx = _ctx_with(1.0)
    expression = "IFS(1<0,1,2<0,3)"
    # SBE 与 LAE 均抛 #N/A 语义错误。
    with pytest.raises(ExcelNaError):
        evaluate_expression(expression)
    with pytest.raises(ExcelNaError):
        evaluate_ast(parse_expression(expression), ctx)
    # IFS 有 TRUE 兜底 → 正常取值。
    assert evaluate_expression("IFS(1<0,1,TRUE,2)") == 2.0
    # ISNA 精确捕获 #N/A，对 #DIV/0! 为假（Excel 口径）。
    assert evaluate_expression("ISNA(IFS(1<0,1,2<0,3))") is True
    assert evaluate_expression("ISNA(1/0)") is False
    # IFERROR 捕获 #N/A 走兜底值。
    assert evaluate_expression("IFERROR(IFS(1<0,1,2<0,3),99)") == 99.0
    # LAE→Office：IFS 兜底编译为 NA()（不是 FALSE）。
    assert compile_ast_to_office(parse_expression(expression), ctx) ==         "IF((1<0),1,IF((2<0),3,NA()))"


def test_ifs_fallback_hint_detects_missing_true_branch() -> None:
    """配置体检提示：IFS 未以 TRUE 兜底结尾（只提示不阻断）。"""
    from base_audit.systems.s3_central_statistics.config import _ifs_fallback_hint

    assert _ifs_fallback_hint("IFS(1<0,1,2<0,3)")
    assert not _ifs_fallback_hint("IFS(1<0,1,TRUE,2)")
    assert not _ifs_fallback_hint("IFS(AND(1>0,2>1),1,TRUE,0)")
    assert not _ifs_fallback_hint("AND(1>0,2>1)")


# ---------------------------------------------------------------------------
# 集成冒烟：真实 Excel 可用时才执行（无 Office 环境自动跳过，UOS 上跳过）
# ---------------------------------------------------------------------------

def _excel_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client

        app = win32com.client.DispatchEx("Excel.Application")
        app.Quit()
        return True
    except Exception:
        return False


EXCEL = pytest.mark.skipif(not _excel_available(), reason="本机无 Excel COM")


@EXCEL
def test_adapter_evaluate_and_cell_fallback_with_real_excel() -> None:
    from base_audit.systems.s3_central_statistics.office_eval import OfficeEvaluationAdapter

    with OfficeEvaluationAdapter("Microsoft Excel") as adapter:
        assert adapter.provider == "EXCEL"
        outcome = adapter.evaluate("AND(1>0,2>1)")
        assert outcome.status == "OK" and outcome.value is True
        assert outcome.path == "evaluate"
        long_formula = "AND(" + ",".join(["1>0"] * 70) + ")"
        outcome = adapter.evaluate(long_formula)
        assert outcome.status == "OK" and outcome.value is True
        assert outcome.path == "cell"
        error = adapter.evaluate("1/0")
        assert error.status == "ERROR"
