from __future__ import annotations

from pathlib import Path
import sys

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))


def test_template_formula_check_reports_only_hard_errors(tmp_path):
    """检查模板公式只报明确错误（#REF!/#VALUE! 等），不列易失/查询引用提示。

    用户口径：公式能否计算一律以 Office 计算引擎真实重算为准，静态层
    不再输出 INDIRECT/OFFSET「易失/动态引用」与 INDEX/MATCH/VLOOKUP
    「查询引用」的需复核清单。
    """
    from base_audit.formula_inspection import inspect_template_formulas

    path = tmp_path / "模板.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "数据"
    sheet["A1"] = "=INDEX(B:B,MATCH(1,C:C,0))"
    sheet["A2"] = '=INDIRECT("B2")'
    sheet["A3"] = "=#REF!"
    sheet["A4"] = "=IF(A1>0,\"ok\",\"#N/A 文本\")"      # 文本含错误字样不算公式错误
    book.defined_names.add(DefinedName("业务坏名称", attr_text="#VALUE!"))
    book.defined_names.add(DefinedName("_FilterDatabase_数据", attr_text="#REF!"))
    book.save(path)
    book.close()

    report = inspect_template_formulas(path)

    assert report["formulaCount"] == 4
    # 只报真实错误：坏名称 + 公式文本中的 #REF!；查询/易失函数不再产出。
    assert {item["location"] for item in report["errors"]} == {"数据!A3", "业务坏名称"}
    assert report.get("warnings") is None


def test_template_check_inventories_named_ranges_by_family(tmp_path):
    """检查模板：业务命名区域按族盘点（校验区域/表结构区域/条件格式区域），
    列出名称+range+作用域；Excel 内部名称单独忽略；其他名称列出让用户确认；
    只提示不阻断。含旧 VBA 变体（本地校验区域1 / 联表校验区域#1）。
    """
    from base_audit.formula_inspection import inspect_template_formulas

    path = tmp_path / "模板.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "数据"
    sheet["A1"] = "=1+1"
    names = [
        ("校验区域", "'数据'!$D$2"),
        ("校验区域#1", "'数据'!$D$5"),
        ("本地校验区域1", "'数据'!$D$8"),
        ("表结构区域", "'数据'!$A$1:$B$1"),
        ("条件格式区域", "'数据'!$B$2"),
        ("我的自定义区域", "'数据'!$F$1"),
        ("_FilterDatabase_数据", "'数据'!$A$1:$B$2"),
    ]
    for name, ref in names:
        book.defined_names.add(DefinedName(name, attr_text=ref))
    book.save(path)
    book.close()

    report = inspect_template_formulas(path)
    groups = {g["family"]: g for g in report["namedRangeGroups"]}
    assert groups["校验区域"]["count"] == 3          # 校验区域 + #1 + 本地校验区域1
    assert {i["name"] for i in groups["校验区域"]["items"]} == {
        "校验区域", "校验区域#1", "本地校验区域1"}
    assert groups["表结构区域"]["items"][0]["range"] == "'数据'!$A$1:$B$1"
    assert groups["表结构区域"]["items"][0]["scope"] == "工作簿"
    assert groups["条件格式区域"]["count"] == 1
    # openpyxl 会丢弃 _xlnm.Print_Area 等内置名（由 print_area 属性管理），
    # 实际保留在 defined_names 的内部名是自动筛选残留 _FilterDatabase*。
    internal = [i["name"] for i in report["internalNames"]]
    assert internal == ["_FilterDatabase_数据"]
    assert [i["name"] for i in report["otherNames"]] == ["我的自定义区域"]
    # 命名区域错误仍计入 errors（原测试场景），盘点本身只提示不阻断。
