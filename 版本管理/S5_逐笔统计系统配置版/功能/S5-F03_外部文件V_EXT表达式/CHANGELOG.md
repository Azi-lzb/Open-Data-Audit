# S5-F03 外部文件 V_EXT 表达式

## v1.0.0 — 2026-10-21

变更类型：Baseline / Feature

首次登记 S5 配置维护工具：沿用 S5 面板当前选择的大集中系统数据，生成 V_EXT 表达式参考副本；兼容普通报表与按本期日期匹配的时序布局，原文件保持只读。入口默认隐藏，通过连续点击 S5 导航五次切换显示。

相关配置：S5-CFG-01 v2.0.0，不修改正式配置。
测试：S5 维护入口默认隐藏及五击切换行为由 `test_s5_maintenance_tools_are_hidden_by_default_and_toggle_with_five_nav_clicks` 覆盖。
Git提交：未提交。
