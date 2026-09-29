"""大集中统计系统：领域模型。

数据结构与 VBA ``B导数比较.bas`` 的数组/字典语义一一对应，字段名保留业务
口径（业务类/机构类代码/…），便于与 VBA 黄金结果逐列对照。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Union

# 数据值：正常为 float；异常值（空/非数值文本）按 VBA 语义原样保留（不转单位）。
CentralValue = Union[float, str, None]


def make_record_key(
    biz_class: str,
    org_code: str,
    region_code: str,
    indicator: str,
    data_attr: str,
    currency: str,
    frequency: str,
    batch: str,
) -> str:
    """逻辑唯一键：8 段直连（无分隔符），与前导单引号清理（VBA 386/941/1025 行）。"""
    return (
        f"{biz_class}{org_code}{region_code}{indicator}"
        f"{data_attr}{currency}{frequency}{batch}"
    )


@dataclass
class CentralRecord:
    """一条大集中导数（CSV 一行）。"""

    biz_class: str          # 业务类：人民币/外币/本外币
    record_date: str        # 数据日期 YYYY-MM-DD
    org_code: str           # 机构类代码（已去前导单引号）
    org_name: str
    region_code: str        # 地区代码
    region_name: str
    order_code: str         # 指标顺序码
    indicator: str          # 指标代码
    indicator_name: str
    data_attr: str          # 数据属性：余额/发生额
    currency: str           # 币种：人民币/美元合计/本外币
    frequency: str          # 频度：月/季…
    batch: str              # 批次
    value: CentralValue     # 数据值（换算后；异常文本原样）

    def key(self) -> str:
        return make_record_key(
            self.biz_class, self.org_code, self.region_code, self.indicator,
            self.data_attr, self.currency, self.frequency, self.batch,
        )

    def attr_key(self) -> str:
        """指标+属性+币种+频度+批次（VBA uni_code，5 段直连）。"""
        return f"{self.indicator}{self.data_attr}{self.currency}{self.frequency}{self.batch}"


@dataclass
class ImportIssue:
    """导入期的数据异常（不中断，进运行日志）。"""

    row: int
    message: str


@dataclass
class CentralDataset:
    """一期数据：行序保持文件顺序；键索引重复时后行覆盖前行（VBA 语义）。"""

    records: list[CentralRecord] = field(default_factory=list)
    issues: list[ImportIssue] = field(default_factory=list)
    record_date: str = ""    # 本期唯一日期（CheckDateUnieque 校验后）
    key_index: dict[str, CentralRecord] = field(default_factory=dict)
    freq_dates: dict[str, str] = field(default_factory=dict)   # 频度 -> 日期

    def build_index(self) -> None:
        self.key_index = {}
        for record in self.records:
            self.key_index[record.key()] = record   # 后行覆盖前行


@dataclass
class ComparisonRow:
    """比较结果一行（黄金输出 21 列中的可变部分）。"""

    record: CentralRecord                 # 第 1-14 列的基础字段
    order_code_out: str = ""              # 第 7 列输出值（顺序参照替换后的值）
    indicator_name_out: str = ""          # 第 9 列输出值（别名替换后的值）
    prev_value: CentralValue = None       # 上期值
    change: CentralValue = None           # 增减额（已舍入）
    ratio: CentralValue = None            # 环比（小数或「除数为0」）
    remark: str = ""                      # 备注
    need_explain: str = ""                # 是否说明（各来源拼接）
    explain_content: str = ""             # 说明内容
    process: str = ""                     # 计算过程
    fill_color: int | None = None         # 填充色（VBA colorindex）
    fill_scope: str = ""                  # "row"=整行(1-18) / "remark"=仅备注列


@dataclass
class ComparisonResult:
    output_path: Path
    rows: list[ComparisonRow] = field(default_factory=list)
    log_rows: list[tuple] = field(default_factory=list)
    candidate_rows: list[tuple] = field(default_factory=list)   # 说明候选
    seconds: float = 0.0


@dataclass
class CheckResult:
    output_path: Path
    rows: list[dict] = field(default_factory=list)   # 对比结果行（列名->值）
    log_rows: list[tuple] = field(default_factory=list)
    seconds: float = 0.0


@dataclass
class FileSetResult:
    directory: Path
    files: list[Path] = field(default_factory=list)
    log_rows: list[tuple] = field(default_factory=list)
    seconds: float = 0.0
    stats: dict = field(default_factory=dict)


def parse_central_date(value) -> str:
    """把 CSV/单元格日期统一为 YYYY-MM-DD；失败返回空串（由调用方记日志）。"""
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip().lstrip("'")
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return ""


def parse_number(value) -> "float | None":
    """VBA ``val()`` 语义：从左侧解析数字前缀，非数字开头返回 None。"""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    best: float | None = None
    for end in range(len(text), 0, -1):
        try:
            best = float(text[:end])
            break
        except ValueError:
            continue
    return best
