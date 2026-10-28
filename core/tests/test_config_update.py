"""发行包配置更新：检测、备份、应用与失败处理。

安装包升级只把新版配置放进 ``/opt/.../config``，启动脚本遇同名文件不覆盖，
因此已安装用户拿不到新版配置。这里的用例固定"检测要说真话、更新必先备份、
失败不破坏已有文件"三条约定。
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

from base_audit import config_update


def _write_workbook(path: Path, marker: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"PK\x03\x04" + marker.encode("utf-8"))


def _manifest(tmp_path: Path, monkeypatch, configs: dict[str, dict[str, str]]) -> Path:
    manifest = tmp_path / "VERSION_MANIFEST.json"
    manifest.write_text(json.dumps({"configs": configs}, ensure_ascii=False), encoding="utf-8")
    monkeypatch.setattr(config_update, "manifest_path", lambda: manifest)
    return manifest


def _env(tmp_path: Path, monkeypatch) -> tuple[Path, Path]:
    """构造 安装根/config（随包）与 用户 config 双目录。

    release_config_dir 会从 project_root 逐级向上找 config，冻结环境下
    project_root 是 ``$DATA_HOME/core``，因此用户配置落在其上级的 config。
    """
    release = tmp_path / "release"
    bundled = release / "config"
    bundled.mkdir(parents=True)
    user = tmp_path / "user" / "config"
    user.mkdir(parents=True)
    monkeypatch.setenv("BASE_AUDIT_INSTALL_ROOT", str(release))
    monkeypatch.setenv("BASE_AUDIT_FROZEN", "1")
    # 该能力只对 Linux 安装包提供；用例在任意开发机上都要以 Linux 口径运行。
    monkeypatch.setattr(config_update.sys, "platform", "linux")
    return bundled, user


def _user_root(user: Path) -> Path:
    """冻结环境下的 project_root：$DATA_HOME/core，其上级即用户配置目录。"""
    root = user.parent / "core"
    root.mkdir(parents=True, exist_ok=True)
    return root


CONFIGS = {
    "S2-CFG-02": {
        "name": "大集中外部数据模板",
        "version": "1.2.0",
        "file": "config/外部文件-大集中相关数据20261016.xlsx",
    },
    "S1-CFG-01": {
        "name": "逐笔统计系统配置",
        "version": "2.1.0",
        "file": "config/1.逐笔统计系统_配置.xlsx",
    },
}


def test_missing_local_file_is_reported_as_installable(tmp_path, monkeypatch) -> None:
    bundled, user = _env(tmp_path, monkeypatch)
    _manifest(tmp_path, monkeypatch, CONFIGS)
    _write_workbook(bundled / "外部文件-大集中相关数据20261016.xlsx", "new")

    report = config_update.inspect(_user_root(user))

    assert report["applicable"] is True
    entry = next(item for item in report["entries"] if item["component"] == "S2-CFG-02")
    assert entry["status"] == "missing"
    assert entry["update_available"] is True


def test_identical_file_needs_no_update(tmp_path, monkeypatch) -> None:
    bundled, user = _env(tmp_path, monkeypatch)
    _manifest(tmp_path, monkeypatch, CONFIGS)
    _write_workbook(bundled / "1.逐笔统计系统_配置.xlsx", "same")
    _write_workbook(user / "1.逐笔统计系统_配置.xlsx", "same")

    report = config_update.inspect(_user_root(user))

    entry = next(item for item in report["entries"] if item["component"] == "S1-CFG-01")
    assert entry["status"] == "same"
    assert entry["update_available"] is False


def test_unknown_version_is_not_claimed_as_newer(tmp_path, monkeypatch) -> None:
    """用户改过的文件没有版本记录时必须如实说"无法确认"，不能自称有新版本。"""
    bundled, user = _env(tmp_path, monkeypatch)
    _manifest(tmp_path, monkeypatch, CONFIGS)
    _write_workbook(bundled / "1.逐笔统计系统_配置.xlsx", "bundled")
    _write_workbook(user / "1.逐笔统计系统_配置.xlsx", "user-edited")

    report = config_update.inspect(_user_root(user))

    entry = next(item for item in report["entries"] if item["component"] == "S1-CFG-01")
    assert entry["status"] == "unknown"
    # 允许用户主动更新，但状态必须与"确认有新版"区分开。
    assert entry["local_version"] == ""
    assert entry["update_available"] is True


def test_recorded_older_version_is_reported_as_newer(tmp_path, monkeypatch) -> None:
    bundled, user = _env(tmp_path, monkeypatch)
    _manifest(tmp_path, monkeypatch, CONFIGS)
    _write_workbook(bundled / "1.逐笔统计系统_配置.xlsx", "bundled")
    _write_workbook(user / "1.逐笔统计系统_配置.xlsx", "user-edited")
    (user / config_update.RECORD_NAME).write_text(
        json.dumps({"files": {"1.逐笔统计系统_配置.xlsx": "2.0.0"}}), encoding="utf-8"
    )

    report = config_update.inspect(_user_root(user))

    entry = next(item for item in report["entries"] if item["component"] == "S1-CFG-01")
    assert entry["status"] == "newer"
    assert entry["local_version"] == "2.0.0"


def test_apply_creates_timestamped_backup_before_overwrite(tmp_path, monkeypatch) -> None:
    bundled, user = _env(tmp_path, monkeypatch)
    _manifest(tmp_path, monkeypatch, CONFIGS)
    _write_workbook(bundled / "1.逐笔统计系统_配置.xlsx", "bundled-v2")
    local = user / "1.逐笔统计系统_配置.xlsx"
    _write_workbook(local, "user-v1")

    result = config_update.apply_updates(_user_root(user), ["1.逐笔统计系统_配置.xlsx"])

    assert result["ok"] is True
    assert local.read_bytes().endswith(b"bundled-v2")
    backup_dir = Path(result["backup_dir"])
    assert backup_dir.name.startswith(config_update.BACKUP_PREFIX)
    backup = backup_dir / local.name
    assert backup.is_file()
    assert backup.read_bytes().endswith(b"user-v1"), "备份必须是更新前的内容"
    assert result["restart_required"] is True


def test_apply_rejects_unregistered_name(tmp_path, monkeypatch) -> None:
    bundled, user = _env(tmp_path, monkeypatch)
    _manifest(tmp_path, monkeypatch, CONFIGS)

    result = config_update.apply_updates(_user_root(user), ["不存在.xlsx"])

    assert result["ok"] is False
    assert "未登记" in result["message"]
    assert result["updated"] == []


def test_apply_requires_selection(tmp_path, monkeypatch) -> None:
    _bundled, user = _env(tmp_path, monkeypatch)
    _manifest(tmp_path, monkeypatch, CONFIGS)

    result = config_update.apply_updates(_user_root(user), [])

    assert result["ok"] is False
    assert "未选择" in result["message"]


def test_source_run_is_not_applicable(tmp_path, monkeypatch) -> None:
    """源码运行没有独立随包配置，不应报告"有配置可更新"。"""
    monkeypatch.delenv("BASE_AUDIT_FROZEN", raising=False)
    monkeypatch.setattr(sys, "frozen", False, raising=False)
    monkeypatch.setattr(config_update.sys, "platform", "linux")

    report = config_update.inspect(tmp_path)

    assert report["applicable"] is False
    assert "源码" in report["reason"]


def test_record_known_versions_only_for_identical_files(tmp_path, monkeypatch) -> None:
    bundled, user = _env(tmp_path, monkeypatch)
    _manifest(tmp_path, monkeypatch, CONFIGS)
    _write_workbook(bundled / "1.逐笔统计系统_配置.xlsx", "same")
    _write_workbook(user / "1.逐笔统计系统_配置.xlsx", "same")
    _write_workbook(bundled / "外部文件-大集中相关数据20261016.xlsx", "bundled")
    _write_workbook(user / "外部文件-大集中相关数据20261016.xlsx", "edited")

    config_update.record_known_versions(user, bundled)

    record = json.loads((user / config_update.RECORD_NAME).read_text(encoding="utf-8"))
    assert record["files"] == {"1.逐笔统计系统_配置.xlsx": "2.1.0"}


def test_windows_platform_is_not_supported(tmp_path, monkeypatch) -> None:
    """Windows 是免安装目录布局，配置随程序整体更换，不提供配置更新入口。"""
    bundled, user = _env(tmp_path, monkeypatch)
    _manifest(tmp_path, monkeypatch, CONFIGS)
    _write_workbook(bundled / "1.逐笔统计系统_配置.xlsx", "bundled")
    _write_workbook(user / "1.逐笔统计系统_配置.xlsx", "user-edited")
    monkeypatch.setattr(config_update.sys, "platform", "win32")

    assert config_update.supported() is False
    report = config_update.inspect(_user_root(user))
    assert report["applicable"] is False
    assert "Windows" in report["reason"]


def test_windows_apply_is_refused(tmp_path, monkeypatch) -> None:
    bundled, user = _env(tmp_path, monkeypatch)
    _manifest(tmp_path, monkeypatch, CONFIGS)
    _write_workbook(bundled / "1.逐笔统计系统_配置.xlsx", "bundled")
    local = user / "1.逐笔统计系统_配置.xlsx"
    _write_workbook(local, "user-v1")
    monkeypatch.setattr(config_update.sys, "platform", "win32")

    result = config_update.apply_updates(_user_root(user), ["1.逐笔统计系统_配置.xlsx"])

    assert result["ok"] is False
    assert result["updated"] == []
    assert local.read_bytes().endswith(b"user-v1"), "Windows 下不得改动配置"


def test_linux_platform_is_supported(monkeypatch) -> None:
    monkeypatch.setattr(config_update.sys, "platform", "linux")
    assert config_update.supported() is True


def test_bundled_dir_inferred_from_frozen_executable(tmp_path, monkeypatch) -> None:
    """直接运行冻结程序（无 BASE_AUDIT_INSTALL_ROOT）时按可执行文件位置推断。"""
    release = tmp_path / "release"
    app_dir = release / "app"
    app_dir.mkdir(parents=True)
    (release / "config").mkdir()
    exe = app_dir / "审核工具_V26.1.8.0"
    exe.write_bytes(b"bin")
    monkeypatch.delenv("BASE_AUDIT_INSTALL_ROOT", raising=False)
    monkeypatch.delenv("BASE_AUDIT_CORE_ROOT", raising=False)
    monkeypatch.setattr(config_update.sys, "frozen", True, raising=False)
    monkeypatch.setattr(config_update.sys, "executable", str(exe), raising=False)
    core = release / "core"
    core.mkdir()

    assert config_update.bundled_config_dir(core) == release / "config"


def test_user_config_parent_is_not_mistaken_for_bundled(tmp_path, monkeypatch) -> None:
    """DEB 安装后 $DATA_HOME/config 是用户目录，不能被当成随包目录。"""
    release = tmp_path / "release"
    (release / "app").mkdir(parents=True)
    (release / "config").mkdir()
    data_home = tmp_path / "home"
    user_config = data_home / "config"
    user_config.mkdir(parents=True)
    core = data_home / "core"
    core.mkdir()
    exe = release / "app" / "审核工具"
    exe.write_bytes(b"bin")
    monkeypatch.delenv("BASE_AUDIT_INSTALL_ROOT", raising=False)
    monkeypatch.delenv("BASE_AUDIT_CORE_ROOT", raising=False)
    monkeypatch.setattr(config_update.sys, "frozen", True, raising=False)
    monkeypatch.setattr(config_update.sys, "executable", str(exe), raising=False)

    resolved = config_update.bundled_config_dir(core)

    assert resolved == release / "config"
    assert resolved != user_config, "不能把用户配置目录当作随包配置目录"
