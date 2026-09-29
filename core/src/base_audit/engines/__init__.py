"""计算引擎适配层：按操作系统提供可用引擎清单与提示。

业务层与前端只通过本模块（及 ``protocol.Calculator``）认识引擎；
不得在业务代码中直接判断 COM、UNO 或操作系统。
"""

from __future__ import annotations

import sys

from .excel_wps import ENGINE_AUTO, ENGINE_EXCEL, ENGINE_WPS, com_available
from .libreoffice_adapter import LibreOfficeAdapter
from .protocol import CalculationResult, Calculator, scan_formula_errors

ENGINE_LIBREOFFICE = "LibreOffice Calc"

# “汇总校验结果说明”只需要读取已保存的单元格值和命名区域；它不复制公式、
# 不提取条件格式。因此它的读取方式必须和审核主流程的计算引擎分开保存。
SUMMARY_READER_PYTHON = "纯 Python"
SUMMARY_READER_COM = "Excel/WPS"
SUMMARY_READER_LIBREOFFICE = ENGINE_LIBREOFFICE


def available_engines() -> list[dict[str, str]]:
    """当前环境的计算引擎选项（启动时按操作系统识别）。

    高级设置统一展示四个引擎：本平台可用的可保存执行；非本平台的项带
    ``disabled`` 标记与简短说明（便于开发者知道选项存在），保存校验
    （``valid_engine_values``）仍按平台过滤，保证不会选到未实现的适配器。
    """
    if sys.platform == "win32":
        return [
            {"value": ENGINE_AUTO, "label": "自动（推荐）"},
            {"value": ENGINE_EXCEL, "label": "EXCEL"},
            {"value": ENGINE_WPS, "label": "WPS"},
            {"value": ENGINE_LIBREOFFICE, "label": "LIBREOFFICE", "disabled": True},
        ]
    return [
        {"value": ENGINE_AUTO, "label": "自动（推荐）"},
        {"value": ENGINE_LIBREOFFICE, "label": "LIBREOFFICE"},
        {"value": ENGINE_EXCEL, "label": "EXCEL", "disabled": True},
        {"value": ENGINE_WPS, "label": "WPS", "disabled": True},
    ]


def engine_hint() -> str:
    if sys.platform == "win32":
        return (
        "自动模式优先使用 EXCEL，其次 WPS。指定引擎后不再自动切换。"
        )
    return "自动模式会选择当前可用的计算引擎；指定引擎后不再自动切换。"


def valid_engine_values() -> set[str]:
    """本平台可保存执行的引擎（不含仅展示用途的禁用项）。"""
    return {item["value"] for item in available_engines() if not item.get("disabled")}


def available_summary_readers() -> list[dict[str, str]]:
    """返回“汇总校验结果说明”可选的读取实现。

    默认的纯 Python 路径以 openpyxl 读取工作簿中已保存的值，适合当前没有
    公式的报送说明/本地校验结果文件。另两项只在各自的平台展示，避免用户把
    Windows COM 或 UOS LibreOffice 当作审核主流程的全局计算设置。
    """
    readers = [{
        "value": SUMMARY_READER_PYTHON,
        "label": "OPENPYXL（推荐）",
        "description": "直接读取已保存的数据，速度更快；不启动办公软件。",
    }]
    if sys.platform == "win32":
        readers.append({
            "value": SUMMARY_READER_COM,
            "label": "Office（兼容）",
            "description": "使用当前计算引擎刷新公式缓存后，再读取汇总结果。",
        })
    else:
        readers.append({
            "value": SUMMARY_READER_LIBREOFFICE,
            "label": "Office（兼容）",
            "description": "使用当前计算引擎刷新公式缓存后，再读取汇总结果。",
        })
    return readers


def valid_summary_reader_values() -> set[str]:
    return {item["value"] for item in available_summary_readers()}


def summary_pipeline_kind(reader: str = SUMMARY_READER_PYTHON) -> str:
    """汇总流程选用的 DAG 适配器类型。

    ``native`` 在这里是“无 COM 的工作簿路径”：纯 Python 与 UOS 的
    LibreOffice 临时计算后读取都使用同一套 openpyxl 汇总节点。
    """
    return "com" if reader == SUMMARY_READER_COM else "native"


def pipeline_kind(engine_preference: str = ENGINE_AUTO) -> str:
    """当前环境下应使用的审核管线类型，业务层只问结果不判系统。

    - ``"com"``：Windows Excel/WPS 会话（excel_com.ExcelSession 全功能门面）；
    - ``"native"``：openpyxl + soffice 重算的原生管线（native 包）。
    """
    return "com" if sys.platform == "win32" else "native"


def probe_engines() -> list[dict[str, str]]:
    """逐项探测引擎真实可用性；返回 value/available/detail 供诊断界面使用。"""
    status: list[dict[str, str]] = []
    if sys.platform == "win32":
        for preference in (ENGINE_EXCEL, ENGINE_WPS):
            ok, detail = com_available(preference)
            status.append({"value": preference, "available": "1" if ok else "0", "detail": detail})
    else:
        ok, detail = LibreOfficeAdapter().available()
        status.append({"value": ENGINE_LIBREOFFICE, "available": "1" if ok else "0", "detail": detail})
    return status


__all__ = [
    "CalculationResult",
    "Calculator",
    "scan_formula_errors",
    "available_engines",
    "engine_hint",
    "valid_engine_values",
    "available_summary_readers",
    "valid_summary_reader_values",
    "summary_pipeline_kind",
    "SUMMARY_READER_PYTHON",
    "SUMMARY_READER_COM",
    "SUMMARY_READER_LIBREOFFICE",
    "pipeline_kind",
    "probe_engines",
]
