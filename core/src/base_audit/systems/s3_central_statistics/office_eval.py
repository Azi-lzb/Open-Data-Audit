"""Office 求值适配器：表达式 2×2 的 OFFICE 后端（provider 自动路由）。

对应 V3 合并计划 §五之二：SBE/LAE 都可以把准备好的公式交给办公套件公式引擎
计算；provider 必须记录（EXCEL/WPS/LIBREOFFICE），Office 失败不得静默回退
Python。

provider 路由（任务启动时解析一次，全程固定）：

- Windows：Excel.Application → ket/KET.Application（复用 ExcelSession 隔离
  实例与消息过滤器）；
- UOS/麒麟等 Linux：LibreOffice Calc（复用 engines.libreoffice 的私有
  profile + ``--convert-to`` 重算管线；公式写入临时工作簿后重算并读回缓存值，
  每条求值一次 soffice 进程，适合验证与少量规则，大批量对拍请用批处理工具）。

Windows 求值口径完全对齐原 VBA（``B导数比较.bas`` 1673-1704）：公式 <250 字符
走工作表级 Evaluate，否则写入辅助单元格计算（VBA yc辅助区）。错误值经 COM 是
负整数、经 LibreOffice 缓存是 "#DIV/0!" 类文本，两种形态统一解释为 ERROR 状态，
由调用方决定命中语义。
"""

from __future__ import annotations

import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

EVALUATE_MAX_LENGTH = 250   # VBA：Len(roundRule) < 250 才走 Evaluate
PROVIDER_EXCEL = "EXCEL"
PROVIDER_WPS = "WPS"
PROVIDER_LIBREOFFICE = "LIBREOFFICE"

#: LibreOffice 写回的错误缓存值（openpyxl 读为文本）。
_ERROR_TEXTS = frozenset({
    "#DIV/0!", "#NAME?", "#NUM!", "#VALUE!", "#REF!", "#NULL!", "#N/A",
})


class OfficeEvaluationError(RuntimeError):
    """Office 求值后端不可用或求值失败（明确失败，不静默回退）。"""


#: Excel 2010+ 新函数在存储/写入公式时的 OOXML 内部前缀；LibreOffice 写入
#: 公式必须带该前缀才被识别（UOS 26.8 实测：裸 IFS 报 #NAME?，_xlfn.IFS=2 ✅；
#: _xlfn.IFERROR 反而 #NAME?，裸 IFERROR 可用——IFERROR 不入清单）。
#: 大小写不敏感：SBE 送审文本经小写化，必须 re.IGNORECASE 才能命中。
_XLFN_FUNCTIONS = ("IFS",)
_XLFN_PREFIX = "_xlfn."


def _calc_formula_text(formula: str) -> str:
    r"""把 Excel 惯用公式文本转换为 LibreOffice Calc 写入形态（_xlfn. 前缀）。

    大小写不敏感匹配（SBE 送审文本已小写化）；``(?<![\w.])`` 防止对
    已带 ``_xlfn.`` 前缀的函数二次叠加。
    """
    import re

    return re.sub(
        r"(?<![\w.])(" + "|".join(_XLFN_FUNCTIONS) + r")\s*\(",
        lambda m: _XLFN_PREFIX + m.group(1) + "(",
        formula,
        flags=re.IGNORECASE,
    )


@dataclass(frozen=True)
class OfficeEvaluation:
    """一次 Office 求值的结果。status：OK / ERROR。"""

    status: str
    value: Any = None          # OK：bool/float/str；ERROR：错误代码或错误文本
    provider: str = ""
    path: str = ""             # evaluate / cell / recalc
    reason: str = ""


def interpret_office_value(value: Any) -> tuple[str, Any, str]:
    """把 COM 返回的 Variant 解释为 (status, value, reason)。

    实测（Excel 16 / WPS 12 + pywin32）：布尔→Python bool；数值→float；错误
    值→负整数（如 #DIV/0! = -2146826281）；文本→str。
    """
    if isinstance(value, bool):
        return "OK", value, ""
    if isinstance(value, int):
        return "ERROR", value, f"Excel错误代码 {value}"
    if isinstance(value, float):
        return "OK", value, ""
    if isinstance(value, str):
        return "OK", value, ""
    return "ERROR", value, f"无法解释的返回类型 {type(value).__name__}"


def interpret_calc_value(value: Any) -> tuple[str, Any, str]:
    """把 openpyxl 读回的 LibreOffice 缓存值解释为 (status, value, reason)。"""
    if isinstance(value, bool):
        return "OK", value, ""
    if isinstance(value, (int, float)):
        return "OK", float(value), ""
    if isinstance(value, str):
        stripped = value.strip()
        if stripped in _ERROR_TEXTS:
            return "ERROR", stripped, f"Calc错误 {stripped}"
        return "OK", value, ""
    if value is None:
        return "OK", None, "空结果"
    return "ERROR", value, f"无法解释的返回类型 {type(value).__name__}"


class _WindowsProvider:
    """Windows Excel/WPS 隔离实例（工作表级 Evaluate + 单元格回退）。"""

    def __init__(self, engine_preference: str) -> None:
        self.engine_preference = engine_preference
        self._session = None
        self._sheet = None
        self.provider = ""

    def start(self) -> None:
        from ...excel_com import ExcelSession

        session = ExcelSession(self.engine_preference)
        session.__enter__()
        self._session = session
        self.provider = PROVIDER_EXCEL if session.engine_name == "Microsoft Excel" else PROVIDER_WPS
        self._ensure_scratch()

    def stop(self, exc_type, exc, tb) -> None:
        if self._session is not None:
            self._session.__exit__(exc_type, exc, tb)
        self._session = None
        self._sheet = None

    def _ensure_scratch(self) -> None:
        """懒创建隐藏工作簿的辅助单元格（不落盘，关闭时一并丢弃）。"""
        if self._sheet is None:
            workbook = self._session.excel.Workbooks.Add()
            self._sheet = workbook.Worksheets(1)

    def evaluate(self, formula: str) -> OfficeEvaluation:
        text = str(formula or "").strip()
        self._ensure_scratch()
        return self._evaluate_one(text)

    def evaluate_batch(self, formulas: list[str]) -> list[OfficeEvaluation]:
        return [self.evaluate(item) for item in formulas]

    def _evaluate_one(self, text: str) -> OfficeEvaluation:
        self._ensure_scratch()
        if len(text) < EVALUATE_MAX_LENGTH:
            try:
                raw = self._sheet.Evaluate(text)
            except Exception as exc:   # COM 调用被拒/参数被拒：按 VBA err.Number 口径
                return OfficeEvaluation(
                    "ERROR", None, self.provider, "evaluate", f"{type(exc).__name__}: {exc}")
            status, value, reason = interpret_office_value(raw)
            return OfficeEvaluation(status, value, self.provider, "evaluate", reason)
        # 长公式：写入辅助单元格计算（VBA yc辅助区 同款）。
        try:
            cell = self._sheet.Range("A1")
            # ClearContents/NumberFormat 在部分受控 Excel 实例（OLAP/策略加固）
            # 上被 COM 拒绝；全新工作表的 A1 本就为空且 General 格式，
            # 因此这两步只是尽力而为，Formula/Calculate/Value 才是必需链路。
            for tidy in (cell.ClearContents, lambda: setattr(cell, "NumberFormat", "General")):
                try:
                    tidy()
                except Exception:
                    pass
            cell.Formula = "=" + text
            self._sheet.Calculate()
            raw = cell.Value
        except Exception as exc:
            return OfficeEvaluation(
                "ERROR", None, self.provider, "cell", f"{type(exc).__name__}: {exc}")
        status, value, reason = interpret_office_value(raw)
        return OfficeEvaluation(status, value, self.provider, "cell", reason)


class _LibreOfficeProvider:
    """Linux LibreOffice Calc：临时工作簿写入公式 → soffice 重算 → 读缓存值。

    soffice 可用性在构造期解析（缺失即明确失败）；``start()`` 为协议空操作——
    Windows provider 在 start 里建 COM 会话，本 provider 无会话概念
    （每次 evaluate 独立起 soffice）。
    """

    def __init__(self) -> None:
        from ...engines.libreoffice import LibreOfficeCalculator

        self._calculator = LibreOfficeCalculator()
        self._engine = self._calculator.require_available()
        self.provider = PROVIDER_LIBREOFFICE
        self._workbook_path: Path | None = None
        self._scratch_text = "1"   # A1 单格；重算管线按整簿覆盖写

    def start(self) -> None:
        """协议对齐：无会话需建立，保持空操作。"""

    def stop(self, exc_type, exc, tb) -> None:
        if self._workbook_path is not None:
            try:
                self._workbook_path.unlink(missing_ok=True)
            except OSError:
                pass
        self._workbook_path = None

    def _ensure_workbook(self) -> None:
        if self._workbook_path is None:
            handle, name = tempfile.mkstemp(prefix="expr_eval_", suffix=".xlsx")
            import os
            os.close(handle)
            self._workbook_path = Path(name)

    def evaluate(self, formula: str) -> OfficeEvaluation:
        return self.evaluate_batch([formula])[0]

    def evaluate_batch(self, formulas: list[str]) -> list[OfficeEvaluation]:
        """批量求值：N 条公式一次 soffice 重算（UOS 批量对拍的关键路径）。"""
        from openpyxl import Workbook, load_workbook

        texts = [str(item or "").strip() for item in formulas]
        if any(not item for item in texts):
            raise OfficeEvaluationError("Office 求值公式为空")
        # LibreOffice 写入公式需 _xlfn. 前缀（IFS/IFERROR 等 2010+ 函数）。
        texts = [_calc_formula_text(item) for item in texts]
        self._ensure_workbook()
        book = Workbook()
        sheet = book.active
        for index, item in enumerate(texts, start=1):
            sheet.cell(row=index, column=1).value = "=" + item
        book.save(self._workbook_path)
        book.close()
        try:
            self._calculator.recalculate(self._workbook_path)
        except Exception as exc:
            return [
                OfficeEvaluation("ERROR", None, self.provider, "recalc", f"{type(exc).__name__}: {exc}")
                for _ in texts
            ]
        book = load_workbook(self._workbook_path, data_only=True)
        try:
            raw_values = [book.active.cell(row=index, column=1).value
                          for index in range(1, len(texts) + 1)]
        finally:
            book.close()
        results = []
        for raw in raw_values:
            status, value, reason = interpret_calc_value(raw)
            results.append(OfficeEvaluation(status, value, self.provider, "recalc", reason))
        return results


class OfficeEvaluationAdapter:
    """任务级 Office 求值适配器：provider 一次解析并固定，with 作用域内复用。"""

    def __init__(self, engine_preference: str = "自动") -> None:
        self.engine_preference = engine_preference
        self._provider_impl: Any = None
        self.provider = ""
        self.evaluation_count = 0
        self.cell_path_count = 0

    def __enter__(self) -> "OfficeEvaluationAdapter":
        if sys.platform == "win32":
            impl = _WindowsProvider(self.engine_preference)
        else:
            impl = _LibreOfficeProvider()
        impl.start()
        self._provider_impl = impl
        self.provider = impl.provider
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self._provider_impl is not None:
            self._provider_impl.stop(exc_type, exc, tb)
        self._provider_impl = None

    def evaluate(self, formula: str) -> OfficeEvaluation:
        """求值一条已代入/已编译的公式文本（不带前导 =）。"""
        return self.evaluate_batch([formula])[0]

    def evaluate_batch(self, formulas: list[str]) -> list[OfficeEvaluation]:
        """批量求值；LibreOffice provider 下 N 条公式合并为一次 soffice 重算。"""
        if self._provider_impl is None:
            raise OfficeEvaluationError("Office 求值适配器未启动（须在 with 中使用）")
        texts = [str(item or "").strip() for item in formulas]
        if any(not item for item in texts):
            raise OfficeEvaluationError("Office 求值公式为空")
        self.evaluation_count += len(texts)
        outcomes = self._provider_impl.evaluate_batch(texts)
        if self.provider == PROVIDER_LIBREOFFICE:
            self.cell_path_count += len(texts)
        else:
            self.cell_path_count += sum(1 for item in outcomes if item.path in ("cell", "recalc"))
        return outcomes
