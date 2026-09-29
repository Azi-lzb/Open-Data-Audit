# s1_transaction_statistics（S1 逐笔统计系统）

**本目录当前为空（技术债登记占位）。**

S1 的业务实现仍在 `core/src/base_audit/` 根层：

| 功能 | 当前实现 | 迁移阻塞原因 |
|---|---|---|
| S1-F01/F02/F03/F04 | `service.py`、`region_summary.py`、`merge_org.py` | 与 `history.py`、`excel_com.py`、`name_config.py`、`template.py` 等互相交织，且被共享层 `native/`、`workflow/` 与 `web_app.py` 深度引用；`excel_com.py` 同时承担 COM 基础设施与 S1 业务（混合职责） |
| S1-F05 | `native/template_merge.py` | 位于共享平台层 native/，其结果模型来自 `merge_org.py`（S1 根模块） |

按 rules/11 与本轮重构约定：**混合职责模块不得整文件粗暴搬迁**，
须先做职责拆分（COM 基础设施归 `engines/`，S1 业务归本目录）；
拆分完成前保留原位，形成技术债：

- 后续应把 `excel_com.py` 的通用 COM 会话能力抽到 `engines/`；
- `service.py` 的 DAG 流程编排抽离后，S1 业务处理器迁入本目录；
- 拆分属独立重构任务，禁止在业务修改轮次中顺手进行。
