# S5-F02 模板 V_CUR 表达式

## v1.0.0 — 2026-10-21

变更类型：Baseline / Feature

首次登记 S5 配置维护工具：沿用 S5 面板的当前匹配模板，生成包含 V_CUR 表达式参考的独立副本；原模板保持只读。入口默认隐藏，通过连续点击 S5 导航五次切换显示。

说明：早先同编号的“模板公式审核”历史流程已退役；本组件是模板 V_CUR 表达式维护工具，不恢复旧审核流程。

相关配置：S5-CFG-01 v2.0.0，不修改正式配置。
测试：S5 维护入口默认隐藏及五击切换行为由 `test_s5_maintenance_tools_are_hidden_by_default_and_toggle_with_five_nav_clicks` 覆盖。
Git提交：未提交。
