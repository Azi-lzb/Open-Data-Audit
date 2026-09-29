from pathlib import Path
from unittest.mock import patch

from base_audit.models import CopyRange
from base_audit.region_summary import (
    _combine_header_rows,
    _header_signature,
    _history_context,
    _metadata_plan,
    _output_filename_prefix,
    _output_sheet_name,
    _sort_extra_fields,
    _source_files,
)


def test_multirow_header_is_joined_top_to_bottom_with_underscore() -> None:
    assert _combine_header_rows([
        ["贷款", "贷款", ""],
        ["余额", "笔数", "备注"],
    ]) == ["贷款_余额", "贷款_笔数", "备注"]


def test_recursive_summary_prefers_institution_subdirectories() -> None:
    root = Path("C:/说明目录")
    root_file = root / "汇总信息.xlsx"
    nested = root / "机构A" / "报送说明.xlsx"

    with patch.object(Path, "rglob", return_value=[root_file, nested]):
        assert _source_files(root, None, recursive=True) == [nested]


def test_recursive_summary_uses_root_files_when_no_subdirectory_files() -> None:
    root = Path("C:/说明目录")
    source = root / "报送说明.xlsx"

    with patch.object(Path, "rglob", return_value=[source]):
        assert _source_files(root, None, recursive=True) == [source]


def test_non_recursive_summary_only_reads_current_folder() -> None:
    root = Path("C:/说明目录")
    root_file = root / "报送说明.xlsx"
    nested = root / "机构A" / "报送说明.xlsx"

    with patch.object(Path, "glob", return_value=[root_file]):
        assert _source_files(root, None, recursive=False) == [root_file]


def test_distinct_layout_uses_original_sheet_name() -> None:
    assert _output_sheet_name("说明汇总", ["本地校验结果"], set()) == "本地校验结果"


def test_matching_layout_uses_module_name() -> None:
    assert _output_sheet_name("说明汇总", ["本地校验结果", "在线抽验结果"], set()) == "本地校验结果"


def test_merge_headers_only_when_all_header_cells_match() -> None:
    assert _header_signature(["序号", "金额", None]) == _header_signature(["序号", "金额", ""])
    assert _header_signature(["序号", "金额"]) != _header_signature(["序号", "余额"])


def test_explanation_summary_uses_business_facing_output_name() -> None:
    assert _output_filename_prefix("汇总校验结果说明") == "校验结果与报送说明汇总"
    assert _output_filename_prefix("其他汇总流程") == "区域汇总"


def test_history_context_appends_by_builtin_rule(tmp_path: Path) -> None:
    """追加规则内置：业务说明按 4 个去重列匹配，自动带出历史说明列。"""
    from openpyxl import Workbook

    config = tmp_path / "逐笔统计系统_配置.xlsx"
    book = Workbook()
    history = book.active
    history.title = "业务说明"
    history.append(["来源文件", "来源工作表", "机构名称", "统一社会信用代码", "报送报表",
                    "数据日期", "序号", "业务表单名称", "业务表单英文名称", "币种", "说明内容", "历史说明"])
    history.append(["机构A_202606", "业务说明", "机构A", "CODE", "金融基础数据-单位贷款",
                    "2026-06-30", "1", "存量单位贷款信息", "CLDWDK", "", "本期余额为 1", "上期已说明"])
    book.save(config)
    book.close()

    keys = ["来源文件", "来源工作表", "业务表单名称", "业务表单英文名称"]
    base_headers = [*keys, "币种", "说明内容"]
    context = _history_context(config, "业务说明", base_headers)
    assert context is not None
    lookup, fields, resolved_keys = context
    assert resolved_keys == tuple(keys)
    assert fields == ("历史说明",)
    assert lookup[("机构A_202606", "业务说明", "存量单位贷款信息", "CLDWDK")] == ["上期已说明"]


def test_history_key_matches_numeric_and_text_values(tmp_path: Path) -> None:
    """复合键数值归一化：汇总侧 COM 读到 1.0，历史侧是文本 '1'，必须配对。"""
    from openpyxl import Workbook

    from base_audit.history import _composite_key
    from base_audit.region_summary import _history_context

    config = tmp_path / "逐笔统计系统_配置.xlsx"
    book = Workbook()
    append_sheet = book.active
    append_sheet.title = "历史追加配置"
    append_sheet.append(("结果工作表", "历史工作表", "去重列", "启用", "历史列", "说明"))
    append_sheet.append((
        "业务说明", "业务说明",
        "来源文件、来源工作表、机构名称、统一社会信用代码、报送报表、数据日期、序号、业务表单名称、业务表单英文名称",
        "是", "历史说明", "",
    ))
    history = book.create_sheet("业务说明")
    history.append(["来源文件", "来源工作表", "机构名称", "统一社会信用代码", "报送报表",
                    "数据日期", "序号", "业务表单名称", "业务表单英文名称", "历史说明"])
    # 历史侧序号是文本 '1'
    history.append(["机构A_说明", "业务说明", "机构A", "CODE", "报表", "2026-06-30", "1", "存量", "CLDWDK", "已说明"])
    book.save(config)
    book.close()

    keys = ["来源文件", "来源工作表", "机构名称", "统一社会信用代码", "报送报表",
            "数据日期", "序号", "业务表单名称", "业务表单英文名称"]
    base_headers = [*keys, "币种", "说明内容"]
    context = _history_context(config, "业务说明", base_headers)
    assert context is not None
    lookup, fields, resolved_keys = context
    assert fields == ("历史说明",)
    # 汇总侧序号是 COM 读出的 float 1.0 → 归一化后与文本 '1' 配对
    summary_row = ["机构A_说明", "业务说明", "机构A", "CODE", "报表", "2026-06-30", 1.0, "存量", "CLDWDK"]
    key = _composite_key(base_headers, summary_row, resolved_keys)
    assert lookup[key] == ["已说明"]


def test_extra_fields_are_sorted_top_to_bottom_then_left_to_right() -> None:
    fields = [
        ("数据日期", CopyRange("说明", "F3")),
        ("机构名称", CopyRange("说明", "B3")),
        ("备注", CopyRange("说明", "A4")),
    ]
    assert [name for name, _ in _sort_extra_fields(fields)] == ["机构名称", "数据日期", "备注"]


def test_global_field_on_own_sheet_is_used_without_duplicate_output_header() -> None:
    global_fields = [("机构名称", CopyRange("任务说明", "B3")), ("数据日期", CopyRange("任务说明", "F3"))]
    output_fields, repeated = _metadata_plan(global_fields, [], ["序号", "机构名称"], "任务说明")
    assert [name for name, _ in output_fields] == ["机构名称", "数据日期"]
    assert repeated == ["机构名称"]


def test_enrich_history_columns_appends_by_builtin_rule(tmp_path: Path) -> None:
    """历史说明富化已拆为独立节点：对已生成的汇总表按复合键补人工列。"""
    from openpyxl import Workbook, load_workbook
    from base_audit.region_summary import enrich_history_columns

    config = tmp_path / "逐笔统计系统_配置.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "业务说明"
    sheet.append(["来源文件", "来源工作表", "业务表单名称", "业务表单英文名称", "历史说明"])
    sheet.append(["甲银行", "业务说明", "表A", "TAB_A", "往期人工说明"])
    book.save(config)
    book.close()

    summary = tmp_path / "汇总.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "业务说明"
    sheet.append(["来源文件", "来源工作表", "业务表单名称", "业务表单英文名称", "本期情况"])
    sheet.append(["甲银行", "业务说明", "表A", "TAB_A", 123])
    sheet.append(["乙银行", "业务说明", "表B", "TAB_B", 456])
    book.save(summary)
    book.close()

    enrich_history_columns(summary, config)
    out = load_workbook(summary, read_only=True)
    rows = list(out["业务说明"].iter_rows(values_only=True))
    out.close()
    assert rows[0][-1] == "历史说明"
    assert rows[1][-1] == "往期人工说明"
    assert rows[2][-1] is None  # 无匹配键不填


def test_enrich_history_columns_is_idempotent(tmp_path: Path) -> None:
    from openpyxl import Workbook, load_workbook
    from base_audit.region_summary import enrich_history_columns

    config = tmp_path / "逐笔统计系统_配置.xlsx"
    book = Workbook(); book.active.title = "业务说明"
    book["业务说明"].append(["来源文件", "来源工作表", "业务表单名称", "业务表单英文名称", "历史说明"])
    book.save(config); book.close()

    summary = tmp_path / "汇总.xlsx"
    book = Workbook(); book.active.title = "业务说明"
    book["业务说明"].append(["来源文件", "来源工作表", "业务表单名称", "业务表单英文名称"])
    book["业务说明"].append(["甲银行", "业务说明", "表A", "TAB_A"])
    book.save(summary); book.close()

    enrich_history_columns(summary, config)
    enrich_history_columns(summary, config)  # 二次不应重复叠加
    out = load_workbook(summary, read_only=True)
    headers = [c.value for c in out["业务说明"][1]]
    out.close()
    assert headers == ["来源文件", "来源工作表", "业务表单名称", "业务表单英文名称", "历史说明"]
