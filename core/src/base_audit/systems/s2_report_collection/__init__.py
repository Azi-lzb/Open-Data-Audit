"""报表采集审核 V2。

本包与 :mod:`base_audit.period_compare` 平行存在。它只处理标准化记录、
审核发现和导出，不修改 V1 的运行入口或结果格式。
"""

from .service import PeriodAuditV2Service, V2RunResult

__all__ = ("PeriodAuditV2Service", "V2RunResult")
