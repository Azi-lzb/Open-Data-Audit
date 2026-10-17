"""安装桥接契约与本机请求边界；不启动 Office、不写真实配置。"""
from __future__ import annotations

import ast
import threading
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import urlsplit

import pytest

from base_audit import web_app
from base_audit.engines import libreoffice as libreoffice_engine
from base_audit.native import libreoffice_install
from base_audit.path_browser import browse_directory


class FakeInstaller:
    def __init__(self):
        self.data = {"available": True, "busy": False, "phase": "ready", "canInstall": True}
        self.prepared = []
        self.started = 0

    def snapshot(self):
        return dict(self.data)

    def prepare(self, path):
        self.prepared.append(path)
        return dict(self.data, source=path)

    def start(self):
        self.started += 1
        return dict(self.data, phase="installing", busy=True)


def test_download_guidance_uses_official_sources_only():
    root = Path(__file__).resolve().parents[2]
    html = (root / "core/frontend/web/index.html").read_text(encoding="utf-8")
    assert html.count('href="https://www.libreoffice.org/download/"') == 2
    assert html.count('href="https://downloadarchive.documentfoundation.org/libreoffice/old/"') == 2
    assert html.count("开发人员推荐 LibreOffice 25.8.7") == 2
    assert "github.com/Azi-lzb" not in html
    for filename in ("README.md", "docs/LibreOffice安装说明.md"):
        text = (root / filename).read_text(encoding="utf-8")
        assert "https://www.libreoffice.org/download/" in text
        assert "25.8.7" in text
        assert "releases/tag/Needed" not in text


def test_help_update_keeps_libreoffice_and_kylin_sections_visible_on_every_platform():
    root = Path(__file__).resolve().parents[2]
    html = (root / "core/frontend/web/index.html").read_text(encoding="utf-8")
    updates = html.split("async function renderUpdates(){", 1)[1].split("/* Linux 安装状态", 1)[0]
    installer = html.split("function renderLibreOfficeInstall(data){", 1)[1].split("async function refreshLibreOfficeInstall", 1)[0]

    assert "安装 LibreOffice</button>" in updates
    assert "state.libreOfficeInstallAvailable?" not in updates
    assert "const maintenance=LO_INSTALL.kylinMaintenance||{};" in installer
    assert "const maintenanceSupported=!!maintenance.supported;" in installer
    assert "!maintenanceSupported?' disabled':''" in installer
    assert "当前运行环境不支持 DEB 安装" in installer


def make_api():
    api = web_app.WebApi.__new__(web_app.WebApi)
    api.state = {"busy": False}
    api._execution_start_lock = threading.RLock()
    api._libreoffice_installer = FakeInstaller()
    api._libreoffice_install_refreshed = ""
    return api


def test_installer_command_prefers_apt_for_all_validated_debs(monkeypatch, tmp_path):
    package_paths = [tmp_path / "uno-libs-private.deb", tmp_path / "libreoffice25.8-calc.deb"]
    packages = [
        libreoffice_install._StagedPackage("uno-libs-private", "25.8.7.3", "amd64", path, "a", 1)
        for path in package_paths
    ]
    apt_get = tmp_path / "apt-get"
    monkeypatch.setattr(libreoffice_install, "APT_GET_EXECUTABLE", apt_get)
    monkeypatch.setattr(libreoffice_install.shutil, "which", lambda name: None)
    monkeypatch.setattr(libreoffice_install, "_is_executable", lambda path: Path(path) == apt_get)
    monkeypatch.setattr(libreoffice_install, "_is_root", lambda: True)

    command, authorized = libreoffice_install._build_install_command(packages)

    assert command == [str(apt_get), "install", "--yes", "--no-install-recommends", *(str(path.resolve()) for path in package_paths)]
    assert not authorized


def test_installer_command_uses_pkexec_for_apt_when_app_runs_unprivileged(monkeypatch, tmp_path):
    package = libreoffice_install._StagedPackage(
        "libreoffice25.8", "25.8.7.3", "amd64", tmp_path / "libreoffice.deb", "a", 1
    )
    apt_get = tmp_path / "apt-get"
    pkexec = tmp_path / "pkexec"
    monkeypatch.setattr(libreoffice_install, "APT_GET_EXECUTABLE", apt_get)
    monkeypatch.setattr(libreoffice_install, "PKEXEC_EXECUTABLE", pkexec)
    monkeypatch.setattr(libreoffice_install.shutil, "which", lambda name: None)
    monkeypatch.setattr(libreoffice_install, "_is_root", lambda: False)
    monkeypatch.setattr(libreoffice_install, "_is_executable", lambda path: Path(path) in (apt_get, pkexec))

    command, authorized = libreoffice_install._build_install_command([package])

    assert command == [
        str(pkexec), "--disable-internal-agent", str(apt_get), "install", "--yes",
        "--no-install-recommends", str(package.path.resolve()),
    ]
    assert authorized


def test_installer_falls_back_to_dpkg_when_apt_is_unavailable(monkeypatch, tmp_path):
    package = libreoffice_install._StagedPackage(
        "libreoffice25.8", "25.8.7.3", "amd64", tmp_path / "libreoffice.deb", "a", 1
    )
    monkeypatch.setattr(libreoffice_install, "APT_GET_EXECUTABLE", tmp_path / "apt-get")
    monkeypatch.setattr(libreoffice_install.shutil, "which", lambda name: None)
    monkeypatch.setattr(libreoffice_install, "_is_executable", lambda path: False)
    monkeypatch.setattr(libreoffice_install, "_is_root", lambda: True)

    command, authorized = libreoffice_install._build_install_command([package])

    assert command == [str(libreoffice_install.DPKG_EXECUTABLE), "--install", str(package.path.resolve())]
    assert not authorized


def test_kylin_v11_detection_mentions_maintenance_mode(monkeypatch):
    original = Path.read_text

    def fake_read_text(path, *args, **kwargs):
        if path == Path("/etc/os-release"):
            return 'NAME="Kylin"\nVERSION="银河麒麟桌面操作系统V11"\nID=kylin\n'
        return original(path, *args, **kwargs)

    monkeypatch.setattr(Path, "read_text", fake_read_text)

    assert libreoffice_install._is_kylin_v11()


def test_install_version_probe_falls_back_to_direct_host_engine(monkeypatch, tmp_path):
    direct = libreoffice_engine.CalcEngine(
        "LibreOffice Calc", "系统路径", (str(tmp_path / "soffice"),), str(tmp_path / "soffice")
    )
    monkeypatch.setattr(libreoffice_engine, "find_calc_engine", lambda: None)
    monkeypatch.setattr(libreoffice_engine, "_find_direct_engine", lambda: direct)
    monkeypatch.setattr(
        libreoffice_install, "_is_executable",
        lambda path: Path(path) == Path(direct.command_prefix[0]),
    )
    monkeypatch.setattr(
        libreoffice_install.subprocess, "run",
        lambda command, **kwargs: SimpleNamespace(returncode=0, stdout="LibreOffice 25.8.7.3\n"),
    )

    version, display = libreoffice_install._probe_engine_version((25, 8, 7, 3))

    assert version == "25.8.7.3"
    assert display == direct.display


def test_explicit_source_keeps_audit_state_unchanged(monkeypatch, tmp_path):
    api = make_api()
    monkeypatch.setattr(web_app.sys, "platform", "linux")
    package = tmp_path / "中文 $安装包.tar.gz"
    before = dict(api.state)
    result = api.choose_libreoffice_source(False, str(package))
    assert result["source"] == str(package)
    assert api.state == before
    assert api._libreoffice_installer.prepared == [str(package)]


def test_native_picker_tar_filter_and_cancel(monkeypatch):
    api = make_api()
    monkeypatch.setattr(web_app.sys, "platform", "linux")
    calls = []
    api._select_path = lambda **kw: calls.append(kw) or ()
    assert api.choose_libreoffice_source()["phase"] == "ready"
    assert calls[0]["file_types"] == (("LibreOffice Linux DEB 压缩包", "*.tar.gz *.tgz"),)
    assert api._libreoffice_installer.prepared == []


def test_windows_and_busy_audit_do_not_prepare_or_install(monkeypatch):
    api = make_api()
    monkeypatch.setattr(web_app.sys, "platform", "win32")
    api.choose_libreoffice_source(path="anything")
    assert not api._libreoffice_installer.prepared
    api.state["busy"] = True
    api.start_libreoffice_install()
    assert api._libreoffice_installer.started == 0


@pytest.mark.parametrize("name,args", [
    ("start", ("audit",)), ("start_dag_flow", ("test",)),
    ("create_template_merge", ()), ("start_central_comparison", ()),
    ("start_central_cross_period", ()), ("start_central_forms", ()),
    ("start_central_explanation_export", ()), ("import_central_explanation_feedback", ()),
    ("start_period_compare", ()), ("start_period_compare_v2", ()),
])
def test_installing_blocks_new_audit_jobs(name, args):
    api = make_api()
    api._libreoffice_installer.data["phase"] = "installing"
    api._log = lambda *_: None
    assert getattr(api, name)(*args) is False
    assert "正在管理 LibreOffice 或系统维护模式" in api.state["status"]


def test_finished_install_refreshes_environment_once(monkeypatch):
    from base_audit import system_info
    api = make_api()
    api._libreoffice_installer.data.update(phase="success", logPath="one.log")
    api.developer_access_store = SimpleNamespace(load=lambda: {key: True for key in web_app.DEVELOPER_SYSTEMS}, message="")
    api._developer_unlocked = set()
    calls = []
    monkeypatch.setattr(system_info, "system_environment_summary", lambda **kw: calls.append(kw) or {"libreoffice": "7.6"})
    monkeypatch.setattr(system_info, "system_environment_text", lambda: "LibreOffice 7.6")
    api.get_state()
    api.get_libreoffice_install()
    assert calls == [{"probe_version": True}]
    assert api.state["systemEnvironmentInfo"]["libreoffice"] == "7.6"


def test_install_and_audit_start_checks_are_serialized():
    api = make_api()
    entered = threading.Event()
    release = threading.Event()
    audit_finished = threading.Event()
    results = []
    def install_start():
        entered.set()
        assert release.wait(3)
        api._libreoffice_installer.data.update(phase="installing", busy=True)
        return api._libreoffice_installer.snapshot()
    def audit_start():
        results.append(api.start("audit"))
        audit_finished.set()
    api._libreoffice_installer.start = install_start
    installing = threading.Thread(target=api.start_libreoffice_install, daemon=True)
    auditing = threading.Thread(target=audit_start, daemon=True)
    installing.start()
    assert entered.wait(3)
    auditing.start()
    try:
        assert not audit_finished.wait(0.1)
    finally:
        release.set()
        installing.join(3)
        auditing.join(3)
    assert not installing.is_alive() and not auditing.is_alive()
    assert results == [False]


def test_browser_filter_supports_compound_suffix(tmp_path):
    (tmp_path / "LibreOffice.tar.gz").write_bytes(b"x")
    (tmp_path / "other.zip").write_bytes(b"x")
    (tmp_path / "package.tgz").write_bytes(b"x")
    result = browse_directory(tmp_path, mode="libreoffice_archive")
    assert [p["name"] for p in result["files"]] == ["LibreOffice.tar.gz", "package.tgz"]
    assert browse_directory(tmp_path, str(tmp_path / "LibreOffice.tar.gz"), "libreoffice_archive")["selectedFile"]
    assert "不符合" in browse_directory(tmp_path, str(tmp_path / "other.zip"), "libreoffice_archive")["error"]


def flask_client():
    flask = pytest.importorskip("flask", reason="Flask 外壳路由验收需要 Flask 环境")
    # 直接编译真实 create_app，避免外壳模块末尾实例化真实用户工作目录。
    root = Path(__file__).resolve().parents[2]
    tree = ast.parse((root / "shell-flask/app.py").read_text(encoding="utf-8"))
    fn = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "create_app")
    fake = SimpleNamespace(start_libreoffice_install=lambda: {"phase": "installing"})
    namespace = dict(__name__="installer_route_test", Flask=flask.Flask, jsonify=flask.jsonify,
                     request=flask.request, send_from_directory=flask.send_from_directory,
                     urlsplit=urlsplit, FlaskApi=lambda _: fake, PROJECT_ROOT=root)
    exec(compile(ast.Module(body=[fn], type_ignores=[]), "shell-flask/app.py", "exec"), namespace)
    return namespace["create_app"]().test_client()


@pytest.mark.parametrize("kwargs", [
    {"data": "[]", "content_type": "text/plain"},
    {"json": [], "headers": {"Origin": "https://malicious.example"}},
    {"json": [], "headers": {"Origin": "http://[broken"}},
    {"json": [], "base_url": "http://rebind.example"},
    {"json": [], "environ_overrides": {"REMOTE_ADDR": "192.168.1.5"}},
])
def test_flask_install_api_rejects_cross_site_and_remote_requests(kwargs):
    assert flask_client().post("/api/start_libreoffice_install", **kwargs).status_code == 403


def test_flask_install_api_accepts_local_json():
    result = flask_client().post("/api/start_libreoffice_install", json=[], headers={"Origin": "http://localhost"})
    assert result.status_code == 200
    assert result.json["result"]["phase"] == "installing"
