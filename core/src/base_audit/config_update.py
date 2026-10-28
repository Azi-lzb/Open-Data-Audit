"""发行包配置更新检查：对比随包配置与用户配置，备份后再更新。

安装包（DEB/tar.gz）每次升级都会把新版配置放进 ``/opt/.../config``，但启动脚本
只在目标不存在时复制到用户目录，因此已安装用户永远拿不到新版配置。本模块提供
"发现了什么、要不要更新、更新前先备份"的显式流程，供界面按用户选择执行。

版本判定口径（外置配置没有可靠的版本记录，必须如实区分）：

- 随包版本来自内嵌 ``VERSION_MANIFEST.json``，可信。
- 用户侧版本来自本模块记录的 ``.配置版本.json``；首次记录时只在本机文件与随包
  文件内容完全一致的情况下写入，不会凭空声明版本。
- 两侧内容不同且用户侧无记录时，状态为 ``unknown``，如实提示"无法确认版本"，
  由用户决定是否更新；绝不谎称"有新版"。
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

from .app_identity import manifest_path

RECORD_NAME = ".配置版本.json"
BACKUP_PREFIX = "备份-配置更新-"
# 外部文件在 config-real 中可能使用不带日期的稳定名；清单登记的是带日期名称。
_ALIASES = {"外部文件-大集中相关数据": "外部文件-大集中相关数据.xlsx"}


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _manifest_configs() -> dict[str, dict[str, str]]:
    """按文件名索引清单里的配置组件。"""
    try:
        manifest = json.loads(manifest_path().read_text(encoding="utf-8"))
        configs = manifest.get("configs") or {}
    except (OSError, json.JSONDecodeError):
        return {}
    indexed: dict[str, dict[str, str]] = {}
    for component_id, item in configs.items():
        relative = str(item.get("file") or "").replace("\\", "/")
        name = Path(relative).name
        if not name:
            continue
        indexed[name] = {
            "component": component_id,
            "name": str(item.get("name") or name),
            "version": str(item.get("version") or ""),
        }
    return indexed


def _version_tuple(text: str) -> tuple[int, ...] | None:
    parts = str(text or "").strip().split(".")
    if not parts or not all(part.isdigit() for part in parts):
        return None
    return tuple(int(part) for part in parts)


def _is_newer(bundled: str, local: str) -> bool:
    left, right = _version_tuple(bundled), _version_tuple(local)
    if left is None or right is None:
        return False
    width = max(len(left), len(right))
    return left + (0,) * (width - len(left)) > right + (0,) * (width - len(right))


def _is_frozen_release() -> bool:
    """只有安装包/便携包才有独立的随包配置；源码运行不参与。"""
    return bool(os.environ.get("BASE_AUDIT_FROZEN")) or bool(getattr(sys, "frozen", False))


def supported() -> bool:
    """配置更新只用于 Linux 安装包。

    Windows 是免安装目录布局：config/ 与程序同级、由用户自己保管，升级换目录即可，
    不存在"随包目录更新但用户目录不覆盖"的问题；该页签在 Windows 上一律不显示。
    """
    return sys.platform != "win32"


def unsupported_reason() -> str:
    return "Windows 版为免安装目录布局，配置随程序目录一并更换，不需要单独更新配置。"


def bundled_config_dir(project_root: Path) -> Path | None:
    """随包配置目录：安装根下的 config（只读，由包管理器维护）。

    冻结包由启动脚本设置 ``BASE_AUDIT_INSTALL_ROOT``；直接运行冻结程序时该变量
    不存在，需要从可执行文件位置推断。注意不能把 ``project_root.parent/config``
    当作随包目录——DEB 安装后那正是用户配置目录（``$DATA_HOME/config``），
    误判会让检查误以为"随包目录就是本机目录"。
    """
    for raw in (
        os.environ.get("BASE_AUDIT_INSTALL_ROOT"),
        os.environ.get("BASE_AUDIT_CORE_ROOT"),
    ):
        root = str(raw or "").strip()
        if not root:
            continue
        for candidate in (Path(root) / "config", Path(root).parent / "config"):
            if candidate.is_dir():
                return candidate
    # 冻结程序位于 <安装根>/app/，其上一级的 config 才是随包配置。
    if getattr(sys, "frozen", False):
        executable_root = Path(sys.executable).resolve().parent
        for candidate in (
            executable_root.parent / "config",
            executable_root.parent.parent / "config",
        ):
            if candidate.is_dir():
                return candidate
    return None


def _resolve(directory: Path, name: str) -> Path | None:
    target = directory / name
    if target.is_file():
        return target
    for prefix, alias in _ALIASES.items():
        if name.startswith(prefix):
            alternative = directory / alias
            if alternative.is_file():
                return alternative
    return None


def _record_path(config_dir: Path) -> Path:
    return config_dir / RECORD_NAME


def _load_record(config_dir: Path) -> dict[str, str]:
    try:
        payload = json.loads(_record_path(config_dir).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    files = payload.get("files")
    if not isinstance(files, dict):
        return {}
    return {str(key): str(value) for key, value in files.items() if isinstance(value, str)}


def _write_record(config_dir: Path, versions: dict[str, str]) -> None:
    target = _record_path(config_dir)
    temporary = target.with_name(target.name + ".tmp")
    payload = {"updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "files": versions}
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    temporary.replace(target)


def record_known_versions(config_dir: Path, bundled_dir: Path) -> None:
    """把内容完全一致的随包配置登记为当前版本；不修改任何配置文件。

    只在两侧哈希相同（即用户文件就是随包文件）时登记，避免把用户改过的文件
    误标成某个版本。
    """
    configs = _manifest_configs()
    if not configs:
        return
    recorded = _load_record(config_dir)
    changed = False
    for name, meta in configs.items():
        local = _resolve(config_dir, name)
        bundled = _resolve(bundled_dir, name)
        if local is None or bundled is None:
            continue
        if _sha256(local) != _sha256(bundled):
            continue
        if recorded.get(name) != meta["version"]:
            recorded[name] = meta["version"]
            changed = True
    if changed:
        _write_record(config_dir, recorded)


def inspect(project_root: Path) -> dict:
    """检查随包配置与用户配置的差异；只读，不修改任何文件。"""
    from .node_flow_config import release_config_dir

    if not supported():
        return {"applicable": False, "reason": unsupported_reason(), "entries": []}
    if not _is_frozen_release():
        return {"applicable": False, "reason": "源码运行不检查发行包配置更新。", "entries": []}
    bundled_dir = bundled_config_dir(project_root)
    if bundled_dir is None:
        return {"applicable": False, "reason": "未找到随包配置目录。", "entries": []}
    config_dir = release_config_dir(project_root)
    if bundled_dir.resolve() == config_dir.resolve():
        return {"applicable": False, "reason": "当前配置目录就是随包目录。", "entries": []}
    configs = _manifest_configs()
    if not configs:
        return {"applicable": False, "reason": "版本清单未登记配置组件。", "entries": []}
    recorded = _load_record(config_dir)
    entries = []
    for name, meta in sorted(configs.items(), key=lambda item: item[0]):
        bundled = _resolve(bundled_dir, name)
        local = _resolve(config_dir, name)
        if bundled is None:
            continue
        bundled_hash = _sha256(bundled)
        local_version = recorded.get(name, "")
        if local is None:
            status, available = "missing", True
        elif _sha256(local) == bundled_hash:
            status, available = "same", False
        elif local_version and _is_newer(meta["version"], local_version):
            status, available = "newer", True
        elif local_version:
            status, available = "local_equal_or_newer", False
        else:
            status, available = "unknown", True
        entries.append({
            "name": name,
            "component": meta["component"],
            "label": meta["name"],
            "bundled_version": meta["version"],
            "local_version": local_version,
            "status": status,
            "update_available": available,
            "local_path": str(local) if local else "",
            "bundled_path": str(bundled),
        })
    pending = [item for item in entries if item["update_available"]]
    return {
        "applicable": True,
        "config_dir": str(config_dir),
        "bundled_dir": str(bundled_dir),
        "entries": entries,
        "pending_count": len(pending),
    }


def apply_updates(project_root: Path, names: list[str]) -> dict:
    """备份后把选定配置替换为随包版本；任一文件失败即停止并报告已完成项。"""
    from .node_flow_config import release_config_dir

    report = inspect(project_root)
    if not report.get("applicable"):
        return {"ok": False, "message": str(report.get("reason") or "当前环境不支持配置更新。"), "updated": []}
    selected = {str(name) for name in names or []}
    if not selected:
        return {"ok": False, "message": "未选择要更新的配置文件。", "updated": []}
    config_dir = Path(report["config_dir"])
    by_name = {item["name"]: item for item in report["entries"]}
    unknown = selected - set(by_name)
    if unknown:
        return {"ok": False, "message": "未登记的配置文件：" + "、".join(sorted(unknown)), "updated": []}
    backup_dir = config_dir / (BACKUP_PREFIX + datetime.now().strftime("%Y%m%d_%H%M%S"))
    bucket = 1
    while backup_dir.exists():
        bucket += 1
        backup_dir = config_dir / (BACKUP_PREFIX + datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + str(bucket))
    updated: list[dict] = []
    recorded = _load_record(config_dir)
    for name in sorted(selected):
        item = by_name[name]
        bundled = Path(item["bundled_path"])
        local = Path(item["local_path"]) if item["local_path"] else config_dir / name
        backup = ""
        try:
            if local.is_file():
                backup_dir.mkdir(parents=True, exist_ok=True)
                target = backup_dir / local.name
                shutil.copy2(local, target)
                if _sha256(target) != _sha256(local):
                    raise OSError("备份校验失败：" + str(target))
                backup = str(target)
            config_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(bundled, local)
            if _sha256(local) != _sha256(bundled):
                raise OSError("写入校验失败：" + str(local))
        except OSError as exc:
            return {
                "ok": False,
                "message": "更新“" + name + "”失败：" + str(exc) + "；已完成的文件保持不变。",
                "updated": updated,
                "backup_dir": str(backup_dir),
            }
        recorded[name] = item["bundled_version"]
        updated.append({"name": name, "version": item["bundled_version"], "backup": backup})
    _write_record(config_dir, recorded)
    return {
        "ok": True,
        "message": "已更新 " + str(len(updated)) + " 个配置文件。",
        "updated": updated,
        "backup_dir": str(backup_dir),
        "restart_required": True,
    }
