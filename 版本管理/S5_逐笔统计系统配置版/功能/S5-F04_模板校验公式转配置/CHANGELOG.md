# S5-F04 模板校验公式转配置

## v1.0.0 — 2026-10-21

变更类型：Baseline / Feature

首次登记 S5 配置维护工具：沿用 S5 面板当前匹配模板，将模板校验公式转换为独立配置候选；候选默认禁用，需人工审核后再复制到正式配置。模板和正式配置保持只读。入口默认隐藏，通过连续点击 S5 导航五次切换显示。

相关配置：S5-CFG-01 v2.0.0，不修改正式配置。
测试：S5 维护入口默认隐藏及五击切换行为由 `test_s5_maintenance_tools_are_hidden_by_default_and_toggle_with_five_nav_clicks` 覆盖。
Git提交：未提交。
