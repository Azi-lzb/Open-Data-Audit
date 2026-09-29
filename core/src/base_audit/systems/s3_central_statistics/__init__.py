"""大集中统计系统：独立业务系统的业务层实现。

对外仅暴露三个稳定接口（见 service.py）：
``run_comparison`` / ``run_cross_period_check`` / ``render_financial_forms``。
UI 与调度层不得了解 VBA 子过程级别的实现细节。
"""
