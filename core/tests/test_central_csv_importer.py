"""大集中导数导入测试：CSV 与 XLSX 同数据等价、后缀与表头校验、异常值保留。"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "core" / "src"))

from base_audit.systems.s3_central_statistics.csv_importer import read_central_csv

CSV_HEADERS = (
    "业务类,数据日期,机构类代码,机构类名称,地区代码,地区名称,指标顺序码,"
    "指标代码,指标名称,数据属性,币种,频度,批次,数据值"
)

ROW_A = ("人民币", "2026-08-31", "6k0i", "示例甲银行", "4400000", "广东省",
         "1", "33370", "各项贷款", "本期数", "人民币", "月", "2", "12345.67")
ROW_B = ("人民币", "2026-08-31", "6k0i", "示例甲银行", "4400000", "广东省",
         "2", "33371", "各项贷款余额", "本期数", "人民币", "月", "2", "无")


def _write_csv(path: Path) -> None:
    lines = [CSV_HEADERS, ",".join(ROW_A), ",".join(ROW_B)]
    path.write_text("\n".join(lines), encoding="gbk")


def _write_xlsx(path: Path, *, leading_cover_sheet: bool = False) -> None:
    from openpyxl import Workbook

    book = Workbook()

    def _fill(sheet) -> None:
        sheet.append(CSV_HEADERS.split(","))
        for row in (ROW_A, ROW_B):
            cells = list(row)
            # 数值/批次按 Excel 真实形态存：整数列写 int、数据值写 float。
            cells[6] = int(cells[6])
            cells[12] = int(cells[12])
            cells[13] = 12345.67 if row is ROW_A else row[13]
            sheet.append(cells)

    if leading_cover_sheet:
        book.active.title = "说明"   # 数据表之前的无关工作表应被跳过
        _fill(book.create_sheet("导数"))
    else:
        _fill(book.active)
    book.save(path)
    book.close()


class CentralImporterTests(unittest.TestCase):
    def test_csv_and_xlsx_produce_same_dataset(self) -> None:
        with TemporaryDirectory() as folder:
            csv_path = Path(folder) / "本期.csv"
            xlsx_path = Path(folder) / "本期.xlsx"
            _write_csv(csv_path)
            _write_xlsx(xlsx_path)
            csv_data = read_central_csv(csv_path, unit_factor=0.0001)
            xlsx_data = read_central_csv(xlsx_path, unit_factor=0.0001)

            self.assertEqual(csv_data.record_date, "2026-08-31")
            self.assertEqual(xlsx_data.record_date, csv_data.record_date)
            self.assertEqual(
                [(r.indicator, r.value) for r in xlsx_data.records],
                [(r.indicator, r.value) for r in csv_data.records],
            )
            # 12345.67 元 → 亿元；“无”为异常文本原样保留且计入 issues。
            self.assertEqual(xlsx_data.records[0].value, 12345.67 * 0.0001)
            self.assertEqual(xlsx_data.records[1].value, "无")
            self.assertEqual(len(xlsx_data.issues), 1)
            self.assertIn("无", xlsx_data.issues[0].message)

    def test_xlsx_skips_non_data_sheets(self) -> None:
        with TemporaryDirectory() as folder:
            xlsx_path = Path(folder) / "本期.xlsx"
            _write_xlsx(xlsx_path, leading_cover_sheet=True)
            data = read_central_csv(xlsx_path)
            self.assertEqual(data.record_date, "2026-08-31")
            self.assertEqual(len(data.records), 2)

    def test_exempt_indicator_skips_unit_factor(self) -> None:
        with TemporaryDirectory() as folder:
            xlsx_path = Path(folder) / "本期.xlsx"
            _write_xlsx(xlsx_path)
            data = read_central_csv(xlsx_path, unit_factor=0.0001, exempt_indicators={"33370"})
            self.assertEqual(data.records[0].value, 12345.67)
            self.assertEqual(data.records[1].value, "无")

    def test_unsupported_suffix_is_rejected(self) -> None:
        with TemporaryDirectory() as folder:
            path = Path(folder) / "本期.xlsm"
            path.write_bytes(b"not a real file")
            with self.assertRaises(ValueError) as ctx:
                read_central_csv(path)
            self.assertIn("csv、xlsx 或 xls", str(ctx.exception))

    def test_corrupt_xls_reports_friendly_error(self) -> None:
        with TemporaryDirectory() as folder:
            path = Path(folder) / "本期.xls"
            path.write_bytes(b"not a real xls")
            with self.assertRaises(ValueError) as ctx:
                read_central_csv(path)
            self.assertIn("XLS 无法读取", str(ctx.exception))

    def test_xlsx_wrong_header_is_rejected(self) -> None:
        from openpyxl import Workbook

        with TemporaryDirectory() as folder:
            path = Path(folder) / "其他表.xlsx"
            book = Workbook()
            book.active.append(["甲", "乙", "丙"])
            book.save(path)
            book.close()
            with self.assertRaises(ValueError) as ctx:
                read_central_csv(path)
            self.assertIn("14 列", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()


def _write_query_xls(path, rows) -> None:
    """xlwt 生成 14 列查询导出 .xls（测试夹具；xlwt 仅为测试依赖）。"""
    import datetime

    import xlwt

    book = xlwt.Workbook()
    sheet = book.add_sheet("查询")
    date_style = xlwt.easyxf(num_format_str="YYYY-MM-DD")
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            if isinstance(value, datetime.datetime):
                sheet.write(r, c, value, date_style)
            elif isinstance(value, (int, float)):
                sheet.write(r, c, float(value))
            else:
                sheet.write(r, c, value)
    book.save(str(path))


def test_read_central_xls_query_export(tmp_path) -> None:
    """纯读取入口兼容 .xls：xlrd 读取 + 日期单元格转文本 + 数值解析。"""
    import datetime

    import pytest

    xlwt = pytest.importorskip("xlwt")
    assert xlwt
    headers = [
        "业务类", "数据日期", "机构类代码", "机构类名称", "地区代码", "地区名称",
        "指标顺序码", "指标代码", "指标名称", "数据属性", "币种", "频度", "批次", "数据值",
    ]
    rows = [headers, [
        "人民币", datetime.datetime(2026, 8, 31), "'6k0i", "测试机构", "'4400000", "广东省",
        1, "'3375V", "贷款", "余额", "人民币", "月", 2, 123.45,
    ]]
    path = tmp_path / "查询导出.xls"
    _write_query_xls(path, rows)
    dataset = read_central_csv(path)
    record = dataset.records[0]
    assert dataset.record_date == "2026-08-31"
    assert record.org_code == "6k0i"          # 前导引号清理
    assert record.indicator == "3375V"
    assert record.frequency == "月"
    assert record.value == 123.45
