"""DAG 节点流程配置工作簿读取器。

配置工作簿位于发行根目录的 ``config/config.xlsx``。它只覆盖
DAG 节点的运行选项。工作簿的说明行可以自由增减；程序会自动定位实际表头行。
"""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from typing import Any

from .workflow.graph import WorkflowDefinition
from .workflow.types import FailurePolicy


# exe 级唯一节点配置工作簿：所有 DAG 节点的启用/失败后处理/区域组，
# 以及组合方案、组合分组与输出命名，都在 config.xlsx 中维护。
# （逐笔统计系统_配置.xlsx 另行承载逐笔流程专属配置与审核历史。）
CONFIG_NAME = "config.xlsx"
NODE_SHEET = "节点配置"
REGION_SHEET = "区域组"
PLAN_SHEET = "组合方案"
GROUP_SHEET = "组合分组"
PLAN_KEY = "方案名称"
NODE_HEADERS = ("流程", "顺序", "节点名称", "启用", "失败后处理", "区域组", "说明")

# 默认区域组登记：区域组的区域类型与适用功能。检查类节点与汇总节点的目标
# 区域都通过「区域组」列引用这里的组名，用户可增改。
_DEFAULT_REGION_GROUPS = (
    ("默认校验区域", "校验区域", "检查校验区域、公式校验复制、校验结果提取"),
    ("默认表结构区域", "表结构区域", "检查表结构区域、表结构比对"),
    ("默认条件格式区域", "条件格式区域", "条件格式结果提取"),
    ("默认表头区域", "表头区域", "检查表头区域、任意行汇总、固定行汇总"),
    ("默认任意行汇总区域", "任意行汇总区域", "检查汇总区域、任意行汇总"),
    ("默认固定行汇总区域", "固定行汇总区域", "固定行汇总"),
    ("默认单元格区域", "单元格区域", "单元格提取、单项检查"),
    ("默认全局单元格区域", "全局单元格区域", "跨文件全局单元格提取"),
)

# 节点说明：默认说明适用于所有节点，个别节点写得更具体，方便使用者在
# 配置表里看懂该节点实际做了什么（尤其是内嵌了历史说明富化的汇总节点）。
_NODE_NOTES = {
    "检查校验区域": "只检查模块化功能“检查校验区域”所需的命名区域；不复制、不计算。",
    "检查表结构区域": "只检查表结构比对所需的命名区域；不比对文件。",
    "检查汇总区域": "检查“任意行汇总区域”命名区域是否存在；查哪个区域由本行“区域组”决定。",
    "检查表头区域": "检查“表头区域”命名区域是否存在；汇总表头由此定位，缺了无法汇总。",
    "表结构比对": "逐格比对固定表头，不匹配的报送文件跳过。",
    "任意行汇总": "按“任意行汇总区域”汇总；只产基础列，历史列由“历史说明富化”节点追加。",
    "历史说明富化": "按复合键把历史表的人工说明列追加到汇总输出。",
}
_DEFAULT_NODE_NOTE = "节点顺序、模块和数据流由系统固定；此处只调整运行选项。"

# 各节点默认「区域组」：新配置/重置时写入，用户可改为其它已登记的区域组或用
# 逗号分隔直接填写命名区域名。检查类节点的目标区域因此完全由 config 决定。
_NODE_REGION_DEFAULTS = {
    "检查校验区域": "默认校验区域",
    "检查表结构区域": "默认表结构区域",
    "检查汇总区域": "默认任意行汇总区域",
    "检查表头区域": "默认表头区域",
    "表结构比对": "默认表结构区域",
    "公式校验复制": "默认校验区域",
    "校验结果提取": "默认校验区域",
    "条件格式结果提取": "默认条件格式区域",
    "任意行汇总": "默认任意行汇总区域",
    "固定行汇总": "默认固定行汇总区域",
}

_FLOW_ALIASES = {
    "汇总核查表校验": "dag:汇总核查表校验",
    "汇总核查表校验（DAG）": "dag:汇总核查表校验",
    "汇总校验结果说明": "dag:汇总校验结果说明",
    "汇总校验结果说明（DAG）": "dag:汇总校验结果说明",
    "组合工作表（核查表）": "dag:组合联合核查表",
    "组合工作表（按配置）": "dag:组合工作表",
}
_FAILURES = {
    "停止": FailurePolicy.FAIL_WORKFLOW,
    "终止流程": FailurePolicy.FAIL_WORKFLOW,
    "跳过": FailurePolicy.SKIP_ITEM,
    "跳过当前文件": FailurePolicy.SKIP_ITEM,
    "跳过当前步骤": FailurePolicy.SKIP_NODE,
    "保留部分结果继续": FailurePolicy.CONTINUE_WITH_PARTIAL,
}


def release_config_dir(project_root: Path) -> Path:
    """存在 config-real（root 或其上级）时优先绑定本机私有配置；
    否则回退随包/仓库的公开空表头配置目录。"""
    root = Path(project_root)
    for candidate in (
        root / "config-real", root.parent / "config-real",
        root / "config", root.parent / "config",
    ):
        if candidate.is_dir():
            return candidate
    return root / "config"


def node_flow_config_path(history_path: Path) -> Path:
    """Locate the release-level node config workbook (config/config.xlsx)."""
    history_path = Path(history_path).resolve()
    # 正式发行线的配置位于 config；测试或便携运行时也允许把
    # config 放在历史工作簿同级目录，便于使用一套明确的 Excel 配置驱动。
    candidates = (
        history_path.parent.parent / "config" / CONFIG_NAME,
        history_path.parent / "config" / CONFIG_NAME,
    )
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return candidates[0]


def _node_note(name: str) -> str:
    return _NODE_NOTES.get(name) or _DEFAULT_NODE_NOTE


def _lookup_override(overrides: dict, node_name: str) -> dict | None:
    """按节点名查找配置行；新节点名比旧配置更长时按前缀兼容匹配。"""
    exact = next((row for key, row in overrides.items() if key == node_name), None)
    if exact is not None:
        return exact
    return next((row for key, row in overrides.items() if node_name.startswith(key)), None)


def check_dag_config(history_path: Path) -> str:
    """检查 exe 级节点配置是否缺失或损坏；正常返回空串，否则返回提示语。"""
    target = node_flow_config_path(history_path)
    if not target.is_file():
        return f"节点配置文件缺失：{target.name}"
    try:
        rows = _node_rows(target)
    except Exception as exc:  # 损坏到无法解析表头
        return f"节点配置文件无法读取：{exc}"
    if not rows:
        return "节点配置缺少“节点配置”表或有效表头"
    return ""


def ensure_dag_config_guide(history_path: Path) -> bool:
    """给已存在的 config.xlsx 补齐/刷新“使用说明”；返回是否写入。"""
    from .config_guide import write_guide_to_file

    return write_guide_to_file(node_flow_config_path(history_path), "config")


def reset_dag_config(history_path: Path) -> Path:
    """重建 exe 级节点配置工作簿 config.xlsx；返回其路径。

    全部 DAG 节点的启用/失败后处理/区域组与组合方案/组合分组
    都在这里维护；审核结果始终输出，运行日志由设置中心“全量运行日志”控制；逐笔流程专属配置与审核历史在
    config/逐笔统计系统_配置.xlsx。重建只写节点口径的默认值，
    不会生成无效图或改变审核业务逻辑。
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from .config_guide import write_guide_sheet

    target = node_flow_config_path(history_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    try:
        nodes = book.active
        nodes.title = NODE_SHEET
        nodes.append(NODE_HEADERS)
        from .workflow.defaults import default_workflows
        for workflow in default_workflows().values():
            flow_name = workflow.name.replace("（DAG）", "")
            for order, node in enumerate(workflow.nodes, start=1):
                node_name = node.display_name or node.node_id
                nodes.append((
                    flow_name, order, node_name, "是",
                    "停止" if node.failure_policy == FailurePolicy.FAIL_WORKFLOW else "跳过当前文件",
                    _NODE_REGION_DEFAULTS.get(node_name, ""), _node_note(node_name),
                ))

        # 区域组：把「区域组」列填的组名映射到实际命名区域。检查类与汇总类节点
        # 都从这里取目标区域，保证“检查什么”与“实际汇总什么”一致。
        regions = book.create_sheet(REGION_SHEET)
        regions.append(("区域组", "命名区域", "区域类型", "适用功能", "说明"))
        for group_name, region_name, feature in _DEFAULT_REGION_GROUPS:
            regions.append((group_name, region_name, region_name, feature, ""))

        plans = book.create_sheet(PLAN_SHEET)
        plans.append(("方案名称", "分组方式", "方案状态", "文件名正则", "输出标识", "说明"))
        plans.append(("组合工作表（核查表）", "文件名正则", "固定只读", "(?P<组合>.+?)_金融基础数据-", "核查表合并", "固定预置，不读取用户方案"))
        plans.append(("按文件名正则提取组合", "文件名正则", "当前使用", "(?P<组合>.+?)_金融基础数据-", "正则组合", "文件名格式稳定时使用"))
        plans.append(("按关键字分组", "文件名关键字", "", "", "按关键字组合", "需要时将本行改为当前使用"))

        groups = book.create_sheet(GROUP_SHEET)
        groups.append(("方案名称", "组合名称", "匹配关键字", "启用", "说明"))
        groups.append(("按关键字分组", "请先配置", "请替换", "否", "新增一行填写组合名称和文件名关键字"))

        write_guide_sheet(book, "config", index=0)  # 说明表已有自己的样式与列宽
        for sheet in book.worksheets:
            if sheet.title == "使用说明":
                continue
            sheet.freeze_panes = "A2"
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
            for column in sheet.columns:
                width = max(12, min(42, max(len(str(cell.value or "")) for cell in column) + 2))
                sheet.column_dimensions[column[0].column_letter].width = width
        _save_workbook(book, target)
    except PermissionError as exc:
        raise RuntimeError("无法重置 config.xlsx：请先关闭已打开的配置文件后重试") from exc
    finally:
        book.close()
    return target


def _save_workbook(book, target: Path) -> None:
    temporary = target.with_name(target.stem + ".tmp.xlsx")
    book.save(temporary)
    temporary.replace(target)


def simplify_dag_config(history_path: Path) -> Path:
    """Replace technical identifiers in an existing config workbook with Chinese business names.

    This is a one-time, lossless layout upgrade: node settings and combination
    rules are retained, while node IDs, module IDs and plan IDs disappear from
    the user-facing sheets.  The fixed DAG remains code-owned.
    """
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill

    target = node_flow_config_path(history_path)
    if not target.is_file():
        return reset_dag_config(history_path)

    nodes_before = _node_rows(target)
    plans_before = _table_rows(target, PLAN_SHEET, {PLAN_KEY, "分组方式", "方案状态"})
    groups_before = _table_rows(target, GROUP_SHEET, {PLAN_KEY, "组合名称", "匹配关键字", "启用"})
    # Old workbooks retain both an ID and a readable plan name.  Group rows
    # formerly point at that ID, so convert them to the readable name.
    plan_names = {
        item.get("方案编号"): item.get(PLAN_KEY)
        for item in plans_before if item.get("方案编号") and item.get(PLAN_KEY)
    }
    node_options = {
        (item.get("流程"), item.get("节点名称") or item.get("节点编号")): item
        for item in nodes_before
    }

    book = load_workbook(target)
    try:
        def replace_sheet(name: str, headers: tuple[str, ...]):
            index = book.sheetnames.index(name) if name in book.sheetnames else len(book.sheetnames)
            if name in book.sheetnames:
                book.remove(book[name])
            sheet = book.create_sheet(name, index)
            sheet.append(headers)
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
            sheet.freeze_panes = "A2"
            return sheet

        nodes = replace_sheet(NODE_SHEET, NODE_HEADERS)
        from .workflow.defaults import default_workflows
        for workflow in default_workflows().values():
            flow_name = workflow.name.replace("（DAG）", "")
            flow_options = {key[1]: row for key, row in node_options.items() if key[0] == flow_name}
            for order, node in enumerate(workflow.nodes, start=1):
                node_name = node.display_name or node.node_id
                old = _lookup_override(flow_options, node_name) or {}
                nodes.append((
                    flow_name, old.get("顺序") or order, node_name, old.get("启用") or "是",
                    old.get("失败后处理") or ("停止" if node.failure_policy == FailurePolicy.FAIL_WORKFLOW else "跳过当前文件"),
                    old.get("区域组") or _NODE_REGION_DEFAULTS.get(node_name, ""),
                    old.get("说明") or _node_note(node_name),
                ))

        plans = replace_sheet(PLAN_SHEET, ("方案名称", "分组方式", "方案状态", "文件名正则", "输出标识", "说明"))
        for item in plans_before:
            name = item.get(PLAN_KEY) or item.get("方案编号")
            if not name:
                continue
            plans.append((name, item.get("分组方式", ""), item.get("方案状态", ""), item.get("文件名正则", ""), item.get("输出标识", ""), item.get("说明", "")))

        groups = replace_sheet(GROUP_SHEET, ("方案名称", "组合名称", "匹配关键字", "启用", "说明"))
        for item in groups_before:
            name = item.get(PLAN_KEY) or item.get("方案编号")
            name = plan_names.get(name, name)
            if name:
                groups.append((name, item.get("组合名称", ""), item.get("匹配关键字", ""), item.get("启用", ""), item.get("说明", "")))

        for sheet in (nodes, plans, groups):
            for column in sheet.columns:
                sheet.column_dimensions[column[0].column_letter].width = max(12, min(42, max(len(str(cell.value or "")) for cell in column) + 2))
        # 补齐/刷新“使用说明”，让升级后的配置也带上版本号与维护指引。
        from .config_guide import write_guide_sheet
        write_guide_sheet(book, "config", index=0)
        temporary = target.with_name(target.stem + ".tmp.xlsx")
        book.save(temporary)
        temporary.replace(target)
    except PermissionError as exc:
        raise RuntimeError("无法更新 config.xlsx：请先关闭已打开的配置文件后重试") from exc
    finally:
        book.close()
    return target


def needs_dag_config_simplification(history_path: Path) -> bool:
    """Whether a workbook still needs a one-time layout upgrade.

    需要升级的两种情况：仍暴露旧技术编号列；或节点配置表缺少“顺序”列与
    “使用说明”表。前者是去编号迁移，后者是补齐顺序列与维护指引。
    """
    path = node_flow_config_path(history_path)
    if not path.is_file():
        return False
    from .config_guide import GUIDE_SHEET
    from openpyxl import load_workbook

    book = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet_name, obsolete in ((NODE_SHEET, "节点编号"), (PLAN_SHEET, "方案编号"), (GROUP_SHEET, "方案编号")):
            if sheet_name not in book.sheetnames:
                continue
            for values in book[sheet_name].iter_rows(values_only=True):
                if obsolete in {_text(value) for value in values}:
                    return True
        if GUIDE_SHEET not in book.sheetnames:
            return True
        if NODE_SHEET not in book.sheetnames:
            return True
        for values in book[NODE_SHEET].iter_rows(max_row=5, values_only=True):
            labels = {_text(value) for value in values}
            if {"流程", "节点名称", "启用"} <= labels:
                return "顺序" not in labels
        return False
    finally:
        book.close()


def _text(value: Any) -> str:
    return str(value or "").strip()


def _enabled(value: Any, default: bool) -> bool:
    text = _text(value)
    if not text:
        return default
    return text not in {"否", "不", "0", "false", "False", "停用"}


def _failure(value: Any, default: FailurePolicy) -> FailurePolicy:
    text = _text(value)
    if not text:
        return default
    return _FAILURES.get(text, FailurePolicy(text) if text in {item.value for item in FailurePolicy} else default)


def _node_rows(path: Path) -> list[dict[str, str]]:
    if not path.is_file():
        return []
    from openpyxl import load_workbook

    book = load_workbook(path, read_only=True, data_only=True)
    try:
        if NODE_SHEET not in book.sheetnames:
            return []
        sheet = book[NODE_SHEET]
        header_row = None
        headers: dict[str, int] = {}
        for row_index, values in enumerate(sheet.iter_rows(values_only=True), start=1):
            possible = {_text(value): column for column, value in enumerate(values) if _text(value)}
            if {"流程", "节点名称", "启用"}.issubset(possible):
                header_row, headers = row_index, possible
                break
        if header_row is None:
            return []
        result: list[dict[str, str]] = []
        for values in sheet.iter_rows(min_row=header_row + 1, values_only=True):
            row = {name: _text(values[index]) if index < len(values) else "" for name, index in headers.items()}
            if row.get("流程") and (row.get("节点名称") or row.get("节点编号")):
                result.append(row)
        return result
    finally:
        book.close()


def _table_rows(path: Path, sheet_name: str, required: set[str]) -> list[dict[str, str]]:
    """Read a labelled worksheet whose header may follow one or more guide rows."""
    if not path.is_file():
        return []
    from openpyxl import load_workbook

    book = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in book.sheetnames:
            return []
        sheet = book[sheet_name]
        header_row = None
        headers: dict[str, int] = {}
        for row_index, values in enumerate(sheet.iter_rows(values_only=True), start=1):
            possible = {_text(value): column for column, value in enumerate(values) if _text(value)}
            # 旧版“组合分组”以方案编号关联；读取时把它视作方案名称，
            # 让升级前的配置文件仍可被新读取器识别。
            if PLAN_KEY in required and PLAN_KEY not in possible and "方案编号" in possible:
                possible[PLAN_KEY] = possible["方案编号"]
            if required.issubset(possible):
                header_row, headers = row_index, possible
                break
        if header_row is None:
            return []
        rows: list[dict[str, str]] = []
        for values in sheet.iter_rows(min_row=header_row + 1, values_only=True):
            row = {name: _text(values[index]) if index < len(values) else "" for name, index in headers.items()}
            if any(row.values()):
                rows.append(row)
        return rows
    finally:
        book.close()


def load_combine_sheets_plan(history_path: Path, plan_id: str | None = None) -> dict[str, object] | None:
    """Read one active combination plan from the release-level config workbook.

    ``None`` means the workbook does not exist.  A present but invalid table
    raises a clear error instead of silently using any other configuration.
    """
    path = node_flow_config_path(history_path)
    if not path.is_file():
        return None
    plans = _table_rows(path, PLAN_SHEET, {PLAN_KEY, "分组方式", "方案状态"})
    groups = _table_rows(path, GROUP_SHEET, {PLAN_KEY, "组合名称", "匹配关键字", "启用"})
    if not plans and not groups:
        return None
    if not plans:
        raise ValueError("config.xlsx 缺少“组合方案”表头或有效方案")

    requested = _text(plan_id)
    if requested:
        selected = next((item for item in plans if item.get(PLAN_KEY) == requested or item.get("方案编号") == requested), None)
        if selected is None:
            raise ValueError("组合工作表方案不存在：{}".format(requested))
    else:
        active = [item for item in plans if item.get("方案状态") == "当前使用"]
        if len(active) != 1:
            raise ValueError("组合方案必须且只能有一条“当前使用”记录")
        selected = active[0]

    selected_id = selected.get(PLAN_KEY) or selected.get("方案编号", "")
    legacy_plan_names = {
        item.get("方案编号"): item.get(PLAN_KEY)
        for item in plans if item.get("方案编号") and item.get(PLAN_KEY)
    }
    mode_text = selected.get("分组方式", "")
    if mode_text in {"文件名正则", "正则", "regex"}:
        pattern = selected.get("文件名正则", "")
        if not pattern:
            raise ValueError("组合方案“{}”未填写文件名正则".format(selected.get("方案名称") or selected_id))
        return {
            "id": selected_id, "name": selected.get("方案名称") or selected_id,
            "mode": "regex", "pattern": pattern, "groups": [],
            "output_marker": selected.get("输出标识") or "组合",
        }
    if mode_text in {"文件名关键字", "按关键字分组", "按名称分组", "name"}:
        grouped = []
        for item in groups:
            group_plan = item.get(PLAN_KEY) or item.get("方案编号")
            group_plan = legacy_plan_names.get(group_plan, group_plan)
            if group_plan != selected_id or item.get("启用") in {"否", "不", "0", "false", "False"}:
                continue
            words = [word.strip() for word in item.get("匹配关键字", "").replace("、", ",").replace("，", ",").split(",") if word.strip()]
            if item.get("组合名称") and words:
                grouped.append({"name": item["组合名称"], "keywords": words})
        if not grouped:
            raise ValueError("组合方案“{}”没有启用的关键字分组".format(selected.get("方案名称") or selected_id))
        return {
            "id": selected_id, "name": selected.get("方案名称") or selected_id,
            "mode": "name", "pattern": "", "groups": grouped,
            "output_marker": selected.get("输出标识") or "组合",
        }
    raise ValueError("组合方案“{}”的分组方式不支持：{}".format(selected.get("方案名称") or selected_id, mode_text or "空"))


def load_combine_output_marker(history_path: Path, plan_id: str, default: str) -> str:
    """Read only the output label for a fixed preset without making it editable."""
    path = node_flow_config_path(history_path)
    rows = _table_rows(path, PLAN_SHEET, {PLAN_KEY, "分组方式", "方案状态"})
    row = next((item for item in rows if item.get(PLAN_KEY) == plan_id or item.get("方案编号") == plan_id), None)
    return _text((row or {}).get("输出标识")) or default


def _region_group_map(history_path: Path) -> dict[str, str]:
    """读取 config「区域组」sheet：区域组名 -> 命名区域（可含多个，用、分隔）。"""
    mapping: dict[str, str] = {}
    for row in _table_rows(node_flow_config_path(history_path), "区域组", {"区域组", "命名区域"}):
        name = _text(row.get("区域组"))
        regions = _text(row.get("命名区域"))
        if name and regions:
            mapping[name] = regions
    return mapping


def _region_names_for_node(value: str, group_map: dict[str, str]) -> tuple[str, ...]:
    """把节点「区域组」列的值解析为命名区域名元组。

    支持三种写法，可混用、分隔符支持 ``、，,；;`` 与换行：
    - 区域组 sheet 中登记的组名（如 ``默认校验区域``）：按 sheet 的「命名区域」列展开；
    - 组名去掉开头「默认」后的裸区域名（历史约定）；
    - 直接填写命名区域名。
    """
    tokens: list[str] = []
    for chunk in value.replace("；", ",").replace(";", ",").replace("，", ",").replace("\n", ",").split(","):
        token = chunk.strip()
        if not token:
            continue
        mapped = group_map.get(token)
        if mapped is None and token.startswith("默认"):
            mapped = group_map.get(token) or group_map.get(token[2:].strip())
        if mapped is not None:
            tokens.extend(item.strip() for item in mapped.replace("、", ",").replace("，", ",").split(",") if item.strip())
            continue
        tokens.append(token[2:].strip() if token.startswith("默认") else token)
    # 去重且保持顺序
    return tuple(dict.fromkeys(name for name in tokens if name))


def load_dag_feature_mappings(history_path: Path | None = None) -> list[object]:
    """返回代码内置的命名区域协议。

    ``config/config.xlsx`` 已退役。保留这个函数名仅为了让 DAG handler 的
    调用点保持稳定；区域组不再由用户配置覆盖。
    """
    from .name_config import DEFAULT_FEATURES, RANGE_FREE_FEATURE_TYPES, FeatureMapping

    mappings = {
        row[0]: FeatureMapping(
            row[0], row[1],
            () if row[1] in RANGE_FREE_FEATURE_TYPES else tuple(
                item.strip() for item in row[2].replace("、", ",").split(",") if item.strip()
            ),
            False, "", row[5], row[4], row[3],
        )
        for row in DEFAULT_FEATURES
    }
    return list(mappings.values())


def apply_node_flow_config(definition: WorkflowDefinition, history_path: Path) -> WorkflowDefinition:
    """兼容旧调用点；固定 DAG 不再接受工作簿覆盖。"""
    return definition
