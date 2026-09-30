# -*- coding: utf-8 -*-
"""产品身份、版本来源与启动器名称约定。"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

from base_audit import app_identity


ROOT = Path(__file__).resolve().parents[2]


def _write_manifest(path: Path, *, version: object = "26.1.0.0", name: str = "审核工具") -> Path:
    path.write_text(
        json.dumps(
            {
                "product": {
                    "id": "V3",
                    "name": name,
                    "version": version,
                    "changelog": "版本管理/产品发布/CHANGELOG.md",
                }
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    return path


def test_manifest_path_resolves_source_and_frozen_locations(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setattr(sys, "frozen", False, raising=False)
    monkeypatch.delattr(sys, "_MEIPASS", raising=False)
    assert app_identity.manifest_path() == ROOT / "版本管理" / "VERSION_MANIFEST.json"

    bundle_root = tmp_path / "bundle"
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(bundle_root), raising=False)
    assert app_identity.manifest_path() == bundle_root / "版本管理" / "VERSION_MANIFEST.json"


def test_product_info_reads_name_and_version_from_requested_manifest(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    manifest = _write_manifest(tmp_path / "VERSION_MANIFEST.json", version="26.2.4.1")

    info = app_identity.load_product_info(manifest)
    assert {key: info[key] for key in ("id", "name", "version", "title")} == {
        "id": "V3",
        "name": "审核工具",
        "version": "26.2.4.1",
        "title": "审核工具 V26.2.4.1",
    }

    monkeypatch.setattr(app_identity, "manifest_path", lambda: manifest)
    assert app_identity.app_title() == "审核工具 V26.2.4.1"
    assert app_identity.executable_stem("linux") == "审核工具_V26.2.4.1"


@pytest.mark.parametrize(
    "version",
    ["v26.1.0.0", "26.1.0", "2026.1.0.0", "26.1.0.0.1", "26.x.0.0", "", None, 1],
)
def test_invalid_product_version_fails_with_runtime_error(
    tmp_path: Path, version: object
) -> None:
    manifest = _write_manifest(tmp_path / "invalid.json", version=version)
    with pytest.raises(RuntimeError, match="版本|version|清单"):
        app_identity.load_product_info(manifest)


def test_missing_or_unreadable_manifest_fails_explicitly(tmp_path: Path) -> None:
    with pytest.raises(RuntimeError, match="清单|manifest|版本"):
        app_identity.load_product_info(tmp_path / "missing.json")

    malformed = tmp_path / "malformed.json"
    malformed.write_text("{not json", encoding="utf-8")
    with pytest.raises(RuntimeError, match="清单|manifest|JSON|json"):
        app_identity.load_product_info(malformed)


def test_executable_stems_distinguish_shell_and_win7_names() -> None:
    info = app_identity.load_product_info()
    version = "V" + info["version"]
    name = info["name"]
    pywebview = app_identity.executable_stem("pywebview")
    flask = app_identity.executable_stem("flask")
    win7 = app_identity.executable_stem("flask", win7=True)
    linux = app_identity.executable_stem("linux")

    assert pywebview == f"{name}_{version}"
    assert flask == f"{name}_Flask_{version}"
    assert win7 == f"{name}_Flask_Win7兼容_{version}"
    assert linux == f"{name}_{version}"
    assert win7 != flask
    with pytest.raises(ValueError):
        app_identity.executable_stem("linux", win7=True)
    with pytest.raises(ValueError):
        app_identity.executable_stem("unknown")


def test_icon_paths_follow_source_and_frozen_bundle_layout(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setattr(sys, "frozen", False, raising=False)
    monkeypatch.delattr(sys, "_MEIPASS", raising=False)
    assert app_identity.icon_path() == ROOT / "core" / "frontend" / "assets" / "app-icon.png"
    assert app_identity.icon_path(".ico") == ROOT / "core" / "frontend" / "assets" / "app-icon.ico"

    bundle_root = tmp_path / "bundle"
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(bundle_root), raising=False)
    assert app_identity.icon_path() == bundle_root / "assets" / "app-icon.png"
    assert app_identity.icon_path(".ico") == bundle_root / "assets" / "app-icon.ico"


def test_rendered_page_uses_safe_title_and_absolute_icon_uri(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    manifest = _write_manifest(
        tmp_path / "VERSION_MANIFEST.json", version="26.1.0.0", name="审核 & 工具"
    )
    icon = tmp_path / "temporary app icon.png"
    monkeypatch.setattr(app_identity, "manifest_path", lambda: manifest)
    monkeypatch.setattr(app_identity, "icon_path", lambda suffix=".png": icon)

    rendered = app_identity.render_app_page(
        "<html><head><title>__AUDIT_APP_TITLE__</title>"
        "<link rel='icon' href='__AUDIT_ICON_URL__'></head></html>"
    )

    assert "<title>审核 &amp; 工具 V26.1.0.0</title>" in rendered
    assert "审核 & 工具" not in rendered
    assert icon.as_uri() in rendered
    assert "__AUDIT_APP_TITLE__" not in rendered
    assert "__AUDIT_ICON_URL__" not in rendered


def _run_identity_cli(*args: str) -> subprocess.CompletedProcess[str]:
    child_env = os.environ.copy()
    child_env["PYTHONIOENCODING"] = "utf-8"
    return subprocess.run(
        [sys.executable, str(Path(app_identity.__file__).resolve()), *args],
        cwd=ROOT,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=child_env,
    )


@pytest.mark.parametrize(
    ("arguments", "expected"),
    [
        (("--version",), app_identity.load_product_info()["version"]),
        (("--title",), app_identity.app_title()),
        (("--linux-name",), app_identity.executable_stem("linux")),
        (("--manifest",), str(ROOT / "版本管理" / "VERSION_MANIFEST.json")),
    ],
)
def test_cli_reports_product_identity_fields(
    arguments: tuple[str, ...], expected: str
) -> None:
    completed = _run_identity_cli(*arguments)
    assert completed.returncode == 0, completed.stderr
    assert completed.stdout.strip() == expected
