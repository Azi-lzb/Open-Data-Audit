"""安全 LibreOffice Linux DEB 安装后端的隔离测试。

测试只使用小型模拟 DEB 文件，并替换 dpkg/授权/版本探测；不会调用真实系统安装器。
"""

from __future__ import annotations

import hashlib
import io
import os
import tarfile
import threading
import time
from pathlib import Path
from types import SimpleNamespace

import pytest

from base_audit.native import libreoffice_install as installer


VERSION = "25.8.7.2-1"
RELEASE = "25.8"
REQUIRED_PACKAGES = (
    "libreoffice25.8",
    "libreoffice25.8-calc",
    "libobasis25.8-core",
    "libobasis25.8-calc",
)


@pytest.fixture
def mocked_linux(monkeypatch):
    monkeypatch.setattr(installer, "_is_linux", lambda: True)
    monkeypatch.setattr(installer, "_tool_support", lambda: (True, ""))
    monkeypatch.setattr(installer, "_architecture_from_dpkg", lambda: "amd64")

    def read_metadata(package_path):
        fields = Path(package_path).read_text(encoding="utf-8").split("|")
        return fields[0], fields[1], fields[2]

    monkeypatch.setattr(installer, "_run_metadata", read_metadata)
    return read_metadata


def _package_payload(name, version=VERSION, architecture="amd64"):
    return "{}|{}|{}".format(name, version, architecture).encode("utf-8")


def _make_debs_folder(root, names=REQUIRED_PACKAGES, version=VERSION, architecture="amd64"):
    debs = Path(root) / "DEBS"
    debs.mkdir(parents=True)
    for index, name in enumerate(names):
        arch = "all" if name.endswith("debian-menus") else architecture
        content = _package_payload(name, version, arch)
        (debs / "package-{}-{}.deb".format(index, name)).write_bytes(content)
    return debs


def _make_tar_gz(path, members):
    """members is an iterable of (archive_name, bytes_or_tarinfo)."""
    with tarfile.open(str(path), mode="w:gz") as archive:
        for name, value in members:
            if isinstance(value, tarfile.TarInfo):
                info = value
                info.name = name
                archive.addfile(info)
            else:
                data = bytes(value)
                info = tarfile.TarInfo(name)
                info.size = len(data)
                archive.addfile(info, io.BytesIO(data))
    return Path(path)


def _archive_for_packages(path, names=REQUIRED_PACKAGES, version=VERSION, architecture="amd64"):
    members = []
    for index, name in enumerate(names):
        arch = "all" if name.endswith("debian-menus") else architecture
        members.append(("LibreOffice_25.8_Linux_x86-64_deb/DEBS/package-{}.deb".format(index), _package_payload(name, version, arch)))
    members.append(("LibreOffice_25.8_Linux_x86-64_deb/readme.txt", b"ignored release notes"))
    return _make_tar_gz(path, members)


def _wait_for(service, phase, timeout=4):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        snapshot = service.snapshot()
        if snapshot["phase"] == phase and not snapshot["busy"]:
            return snapshot
        time.sleep(0.01)
    raise AssertionError("installer did not reach {!r}: {!r}".format(phase, service.snapshot()))


def _prepare_ready(service, source):
    service.prepare(str(source))
    result = _wait_for(service, "ready")
    assert result["canInstall"] is True
    return result


@pytest.mark.parametrize("layout", ["debs", "immediate", "wrapper"])
def test_folder_source_accepts_only_supported_shallow_layouts(tmp_path, layout):
    if layout == "debs":
        source = _make_debs_folder(tmp_path)
    elif layout == "immediate":
        wrapper = tmp_path / "extracted"
        source = _make_debs_folder(wrapper).parent
    else:
        source = tmp_path / "single-child"
        source.mkdir()
        _make_debs_folder(source / "LibreOffice_25.8_Linux_x86-64_deb")
    assert installer._find_debs_directory(source).name == "DEBS"


def test_folder_prepare_stages_only_debs_and_preserves_original_files(tmp_path, mocked_linux):
    source_root = tmp_path / "LibreOffice_25.8_Linux_x86-64_deb"
    debs = _make_debs_folder(source_root)
    (debs / "notes.txt").write_text("do not stage", encoding="utf-8")
    source_hashes = {
        path.name: hashlib.sha256(path.read_bytes()).hexdigest()
        for path in debs.iterdir()
        if path.is_file()
    }
    service = installer.LibreOfficeInstaller(tmp_path / "app-data")

    result = _prepare_ready(service, source_root)

    assert result["packageVersion"] == VERSION
    assert result["architecture"] == "amd64"
    assert result["packageCount"] == 4
    assert result["sourceName"] == source_root.name
    assert result["source"] == str(source_root)
    assert "可信来源" in result["message"]
    assert "signature" not in result["message"].casefold()
    assert {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in debs.iterdir() if path.is_file()} == source_hashes
    staged = list(service._staging_dir.iterdir())
    assert len(staged) == 4
    assert all(path.name.startswith("package-") and path.suffix == ".deb" for path in staged)
    assert all(path.name != "notes.txt" for path in staged)
    if os.name == "posix":
        assert oct(service._staging_dir.stat().st_mode & 0o777) == "0o700"


def test_archive_prepare_copies_only_direct_debs_and_preserves_archive(tmp_path, mocked_linux):
    archive = _archive_for_packages(tmp_path / "LibreOffice_25.8_linux_deb.tar.gz")
    original_hash = hashlib.sha256(archive.read_bytes()).hexdigest()
    service = installer.LibreOfficeInstaller(tmp_path / "app-data")

    result = _prepare_ready(service, archive)

    assert result["packageCount"] == 4
    assert result["packageVersion"] == VERSION
    assert hashlib.sha256(archive.read_bytes()).hexdigest() == original_hash
    assert len(list(service._staging_dir.iterdir())) == 4


def test_tgz_extension_is_accepted(tmp_path, mocked_linux):
    archive = _archive_for_packages(tmp_path / "LibreOffice_25.8_linux_deb.tgz")
    service = installer.LibreOfficeInstaller(tmp_path / "app-data")

    result = _prepare_ready(service, archive)

    assert result["packageCount"] == 4


@pytest.mark.parametrize(
    "stdout",
    [
        "libreoffice25.8-calc\n25.8.7.2-1\namd64\n",
        "Package: libreoffice25.8-calc\nVersion: 25.8.7.2-1\nArchitecture: amd64\n",
    ],
)
def test_dpkg_deb_metadata_parser_accepts_realistic_bare_or_labelled_output(tmp_path, monkeypatch, stdout):
    captured = {}

    def fake_run(command, **kwargs):
        captured["command"] = list(command)
        captured["kwargs"] = kwargs
        return SimpleNamespace(returncode=0, stdout=stdout, stderr="")

    monkeypatch.setattr(installer.subprocess, "run", fake_run)
    package_path = tmp_path / "path with spaces; $HOME.deb"

    assert installer._run_metadata(package_path) == ("libreoffice25.8-calc", VERSION, "amd64")
    assert captured["command"] == [
        str(installer.DPKG_DEB_EXECUTABLE),
        "--field",
        str(package_path),
        "Package",
        "Version",
        "Architecture",
    ]
    assert captured["kwargs"].get("shell", False) is False


@pytest.mark.parametrize(
    "bad_name",
    ["../escape.deb", "/tmp/escape.deb", "root/DEBS/../../escape.deb", "C:/escape.deb"],
)
def test_archive_rejects_traversal_and_absolute_paths(tmp_path, bad_name):
    archive = _make_tar_gz(tmp_path / "unsafe.tar.gz", [(bad_name, b"payload")])
    with pytest.raises(installer._PreparationError, match="路径"):
        installer._stage_archive(archive, tmp_path / "stage", tmp_path)


@pytest.mark.parametrize("kind", [tarfile.SYMTYPE, tarfile.LNKTYPE, tarfile.CHRTYPE])
def test_archive_rejects_links_and_special_files(tmp_path, kind):
    link = tarfile.TarInfo("wrapper/DEBS/package.deb")
    link.type = kind
    link.linkname = "elsewhere"
    if kind == tarfile.CHRTYPE:
        link.devmajor = 1
        link.devminor = 3
    archive = _make_tar_gz(tmp_path / "unsafe.tar.gz", [(link.name, link)])
    with pytest.raises(installer._PreparationError, match="链接或特殊文件"):
        installer._stage_archive(archive, tmp_path / "stage", tmp_path)


def test_archive_rejects_duplicate_canonical_paths(tmp_path):
    archive = _make_tar_gz(
        tmp_path / "duplicate.tar.gz",
        [("root/DEBS/a.deb", b"one"), ("root/DEBS/a.deb", b"two")],
    )
    with pytest.raises(installer._PreparationError, match="重复文件路径"):
        installer._stage_archive(archive, tmp_path / "stage", tmp_path)


def test_archive_member_count_and_expanded_size_limits_are_enforced(tmp_path, monkeypatch):
    too_many = _make_tar_gz(
        tmp_path / "many.tar.gz",
        [("root/DEBS/{}.deb".format(index), b"x") for index in range(3)],
    )
    monkeypatch.setattr(installer, "MAX_ARCHIVE_MEMBERS", 2)
    with pytest.raises(installer._PreparationError, match="数量"):
        installer._stage_archive(too_many, tmp_path / "stage1", tmp_path)

    monkeypatch.setattr(installer, "MAX_ARCHIVE_MEMBERS", 20000)
    # The compressed archive remains small while its expanded member exceeds
    # the configured limit, so this exercises the expanded-size guard.
    monkeypatch.setattr(installer, "MAX_TOTAL_ARCHIVE_BYTES", 1000)
    too_large = _make_tar_gz(tmp_path / "large.tar.gz", [("root/DEBS/a.deb", b"x" * 10000)])
    with pytest.raises(installer._PreparationError, match="总大小"):
        installer._stage_archive(too_large, tmp_path / "stage2", tmp_path)


def test_folder_scan_limit_counts_non_deb_entries(tmp_path, monkeypatch):
    debs = _make_debs_folder(tmp_path / "source")
    for index in range(3):
        (debs / "notes-{}.txt".format(index)).write_text("ignored", encoding="utf-8")
    monkeypatch.setattr(installer, "MAX_ARCHIVE_MEMBERS", 2)

    with pytest.raises(installer._PreparationError, match="数量"):
        installer._folder_package_entries(debs)


def test_free_space_limit_is_actionable(tmp_path, monkeypatch):
    monkeypatch.setattr(installer.shutil, "disk_usage", lambda _path: SimpleNamespace(free=100))
    with pytest.raises(installer._PreparationError, match="磁盘空间不足"):
        installer._assert_free_space(tmp_path, 1000)


@pytest.mark.parametrize(
    "names,version,architecture,reason",
    [
        (REQUIRED_PACKAGES, VERSION, "arm64", "架构与本机不匹配"),
        (("gnome-shell",) + REQUIRED_PACKAGES, VERSION, "amd64", "非 LibreOffice"),
        (REQUIRED_PACKAGES[:-1], VERSION, "amd64", "不完整"),
        (REQUIRED_PACKAGES, "25.8.7.2-1", "amd64", "不一致"),
    ],
)
def test_package_set_rejects_wrong_architecture_unrelated_or_incomplete_packages(
    tmp_path, mocked_linux, monkeypatch, names, version, architecture, reason
):
    debs = _make_debs_folder(tmp_path / "source", names, version=version, architecture=architecture)
    if reason == "不一致":
        last = next(debs.glob("*libobasis25.8-calc*.deb"))
        last.write_text(_package_payload("libobasis25.8-calc", "26.8.0.0-1", architecture).decode("utf-8"), encoding="utf-8")

    service = installer.LibreOfficeInstaller(tmp_path / "app-data")
    service.prepare(str(debs))
    result = _wait_for(service, "failed")

    assert reason in result["message"]
    assert result["canInstall"] is False
    assert service._staging_dir is None


def test_only_official_debian_menu_package_may_use_matching_three_component_version(tmp_path, mocked_linux):
    names = REQUIRED_PACKAGES + ("libreoffice25.8-debian-menus",)
    debs = _make_debs_folder(tmp_path / "source", names, version=VERSION)
    menu_package = next(debs.glob("*debian-menus*.deb"))
    menu_package.write_text(
        _package_payload("libreoffice25.8-debian-menus", "25.8.7-2", "all").decode("utf-8"),
        encoding="utf-8",
    )
    service = installer.LibreOfficeInstaller(tmp_path / "app-data")

    result = _prepare_ready(service, debs)

    assert result["packageVersion"] == VERSION
    assert result["packageCount"] == 5


def test_unversioned_ure_must_match_primary_release_version(tmp_path, mocked_linux):
    debs = _make_debs_folder(tmp_path / "source", REQUIRED_PACKAGES + ("ure",))
    (debs / "package-4-ure.deb").write_text(_package_payload("ure", "25.8.7.1-1").decode("utf-8"), encoding="utf-8")
    service = installer.LibreOfficeInstaller(tmp_path / "app-data")

    service.prepare(str(debs))
    result = _wait_for(service, "failed")

    assert "版本不一致" in result["message"]
    assert result["canInstall"] is False
    assert service._staging_dir is None


def test_unversioned_versioned_release_mismatch_is_rejected(tmp_path):
    files = []
    for index, name in enumerate(REQUIRED_PACKAGES):
        path = tmp_path / "package-{}.deb".format(index)
        path.write_bytes(b"x")
        files.append((path, "digest", 1))
    original = installer._run_metadata
    metadata = {
        "package-0.deb": (REQUIRED_PACKAGES[0], "26.8.0.0-1", "amd64"),
        "package-1.deb": (REQUIRED_PACKAGES[1], "26.8.0.0-1", "amd64"),
        "package-2.deb": (REQUIRED_PACKAGES[2], "26.8.0.0-1", "amd64"),
        "package-3.deb": (REQUIRED_PACKAGES[3], "26.8.0.0-1", "amd64"),
    }
    try:
        installer._run_metadata = lambda path: metadata[Path(path).name]
        with pytest.raises(installer._PreparationError, match="主安装包"):
            installer._validate_package_set(files, "amd64")
    finally:
        installer._run_metadata = original


def test_windows_prepare_is_unavailable_and_never_runs_system_commands(tmp_path, monkeypatch):
    monkeypatch.setattr(installer, "_is_linux", lambda: False)

    def unexpected(*_args, **_kwargs):
        raise AssertionError("system command must not run on Windows")

    monkeypatch.setattr(installer.subprocess, "run", unexpected)
    monkeypatch.setattr(installer.subprocess, "Popen", unexpected)
    service = installer.LibreOfficeInstaller(tmp_path)

    result = service.prepare(str(tmp_path / "sample.tar.gz"))

    assert result["available"] is False
    assert result["supported"] is False
    assert result["phase"] == "failed"
    assert "仅适用于 Linux" in result["message"]


def test_missing_pkexec_and_dpkg_tools_are_reported_without_shell_fallback(monkeypatch):
    monkeypatch.setattr(installer, "_is_linux", lambda: True)
    monkeypatch.setattr(installer, "_is_root", lambda: False)
    monkeypatch.setattr(installer, "_is_executable", lambda path: Path(path) == installer.DPKG_EXECUTABLE)
    supported, message = installer._tool_support()
    assert supported is False
    assert "dpkg-deb" in message

    monkeypatch.setattr(
        installer,
        "_is_executable",
        lambda path: Path(path) in (installer.DPKG_EXECUTABLE, installer.DPKG_DEB_EXECUTABLE),
    )
    supported, message = installer._tool_support()
    assert supported is False
    assert "pkexec" in message


def test_installer_uses_safe_argv_and_reports_successful_target_version(tmp_path, mocked_linux, monkeypatch):
    debs = _make_debs_folder(tmp_path / "source")
    monkeypatch.setattr(installer, "_installed_libreoffice_inventory", lambda: ([], {"display": "", "path": "", "source": ""}))
    monkeypatch.setattr(installer, "_apt_get_path", lambda: None)
    malicious_name = debs / "calc ; touch injected $HOME.deb"
    existing_calc = next(debs.glob("*libreoffice25.8-calc*.deb"))
    existing_calc.rename(malicious_name)
    service = installer.LibreOfficeInstaller(tmp_path / "app data; $HOME")
    _prepare_ready(service, debs)
    monkeypatch.setattr(installer, "_is_root", lambda: False)
    monkeypatch.setattr(installer, "_probe_engine_version", lambda *_args: ("25.8.7.2", "/opt/libreoffice25.8/program/soffice"))
    captured = {}

    class FakeProcess:
        def wait(self):
            return 0

    def fake_popen(command, **kwargs):
        captured["command"] = list(command)
        captured["kwargs"] = kwargs
        return FakeProcess()

    monkeypatch.setattr(installer.subprocess, "Popen", fake_popen)
    service.start()
    result = _wait_for(service, "success")

    command = captured["command"]
    assert command[:4] == [
        str(installer.PKEXEC_EXECUTABLE),
        "--disable-internal-agent",
        str(installer.DPKG_EXECUTABLE),
        "--install",
    ]
    assert all(argument.startswith(str(service.data_dir / ".libreoffice-install-")) for argument in command[4:])
    assert len(command) == 8  # 4 fixed argv entries and 4 staged packages
    assert captured["kwargs"]["shell"] is False
    assert result["installed"] == "LibreOffice 25.8.7.2 (/opt/libreoffice25.8/program/soffice)"
    assert result["phase"] == "success"
    assert result["canInstall"] is False
    assert service._staging_dir is None


@pytest.mark.parametrize("return_code,expected_phase,expected_text", [(126, "cancelled", "取消了"), (127, "failed", "无法启动"), (2, "failed", "安装失败")])
def test_authorization_cancel_and_dpkg_failures_are_not_success(tmp_path, mocked_linux, monkeypatch, return_code, expected_phase, expected_text):
    debs = _make_debs_folder(tmp_path / "source")
    service = installer.LibreOfficeInstaller(tmp_path / "app-data")
    _prepare_ready(service, debs)
    monkeypatch.setattr(installer, "_is_root", lambda: False)

    class FakeProcess:
        def wait(self):
            return return_code

    monkeypatch.setattr(installer.subprocess, "Popen", lambda *_args, **_kwargs: FakeProcess())
    monkeypatch.setattr(installer, "_probe_engine_version", lambda *_args: pytest.fail("must not verify a failed install"))
    service.start()
    result = _wait_for(service, expected_phase)

    assert expected_text in result["message"]
    assert result["installed"] == ""
    assert Path(result["logPath"]).is_file()
    assert result["canInstall"] is False


def test_install_version_mismatch_is_a_failure_and_log_is_technical(tmp_path, mocked_linux, monkeypatch):
    debs = _make_debs_folder(tmp_path / "source")
    service = installer.LibreOfficeInstaller(tmp_path / "app-data")
    _prepare_ready(service, debs)
    monkeypatch.setattr(installer, "_is_root", lambda: True)
    monkeypatch.setattr(installer.subprocess, "Popen", lambda *_args, **_kwargs: SimpleNamespace(wait=lambda: 0))
    monkeypatch.setattr(installer, "_probe_engine_version", lambda *_args: ("7.6.7.2", "/usr/bin/soffice"))

    service.start()
    result = _wait_for(service, "failed")

    assert "无法确认" in result["message"]
    assert "expected" in Path(result["logPath"]).read_text(encoding="utf-8")


def test_post_install_probe_checks_target_binary_even_if_engine_refresh_finds_none(monkeypatch):
    import base_audit.engines.libreoffice as engine_module

    monkeypatch.setattr(engine_module, "find_calc_engine", lambda: None)
    monkeypatch.setattr(installer, "_is_executable", lambda path: Path(path).as_posix() == "/opt/libreoffice25.8/program/soffice")
    monkeypatch.setattr(
        installer.subprocess,
        "run",
        lambda command, **_kwargs: SimpleNamespace(returncode=0, stdout="LibreOffice 25.8.7.2\n"),
    )

    version, display = installer._probe_engine_version((25, 8, 7, 2))

    assert version == "25.8.7.2"
    assert Path(display).as_posix() == "/opt/libreoffice25.8/program/soffice"


def test_post_install_probe_does_not_fall_back_to_an_unrelated_existing_engine(monkeypatch):
    import base_audit.engines.libreoffice as engine_module

    monkeypatch.setattr(engine_module, "find_calc_engine", lambda: SimpleNamespace(command_prefix=("/usr/bin/soffice",), display="/usr/bin/soffice"))
    monkeypatch.setattr(engine_module, "_find_direct_engine", lambda: None)
    monkeypatch.setattr(installer, "_is_executable", lambda _path: False)
    monkeypatch.setattr(installer.subprocess, "run", lambda *_args, **_kwargs: pytest.fail("must not probe an unrelated engine"))

    version, detail = installer._probe_engine_version((25, 8, 7, 2))

    assert version == ""
    assert "expected target soffice was not found" in detail


def test_authorization_does_not_collect_credentials(tmp_path, mocked_linux, monkeypatch):
    debs = _make_debs_folder(tmp_path / "source")
    service = installer.LibreOfficeInstaller(tmp_path / "app-data")
    _prepare_ready(service, debs)
    monkeypatch.setattr(installer, "_is_root", lambda: False)
    monkeypatch.setattr(installer, "_probe_engine_version", lambda *_args: ("25.8.7.2", "/opt/libreoffice25.8/program/soffice"))

    class FakeProcess:
        def wait(self):
            return 126

    monkeypatch.setattr(installer.subprocess, "Popen", lambda *_args, **_kwargs: FakeProcess())
    service.start()
    result = _wait_for(service, "cancelled")
    log = Path(result["logPath"]).read_text(encoding="utf-8")

    assert "password" not in log.casefold()
    assert "sudo" not in log.casefold()
    assert "--disable-internal-agent" in log


def test_preparing_again_invalidates_old_ready_stage_and_concurrent_calls_are_serialized(tmp_path, mocked_linux, monkeypatch):
    first = _make_debs_folder(tmp_path / "first")
    second = tmp_path / "missing-debs"
    service = installer.LibreOfficeInstaller(tmp_path / "app-data")
    first_result = _prepare_ready(service, first)
    old_stage = service._staging_dir

    entered = threading.Event()
    release = threading.Event()
    real_validate = installer._run_metadata

    def slow_metadata(package_path):
        entered.set()
        release.wait(2)
        return real_validate(package_path)

    monkeypatch.setattr(installer, "_run_metadata", slow_metadata)
    service.prepare(str(first))
    assert entered.wait(2)
    second_prepare = service.prepare(str(second))
    install_attempt = service.start()
    assert second_prepare["phase"] == "checking"
    assert install_attempt["phase"] == "checking"
    assert install_attempt["canInstall"] is False
    release.set()
    result = _wait_for(service, "ready")

    assert result["source"] == str(first)
    assert result["packageCount"] == 4
    assert old_stage is not None and not old_stage.exists()


def test_invalid_replacement_clears_previous_ready_packages(tmp_path, mocked_linux):
    source = _make_debs_folder(tmp_path / "source")
    service = installer.LibreOfficeInstaller(tmp_path / "app-data")
    _prepare_ready(service, source)
    old_stage = service._staging_dir

    service.prepare(str(tmp_path / "missing.tar.gz"))
    result = _wait_for(service, "failed")

    assert result["canInstall"] is False
    assert result["packageCount"] == 0
    assert old_stage is not None and not old_stage.exists()


def test_ld_library_path_uses_original_host_value_and_unsets_bundled_value(monkeypatch):
    monkeypatch.setenv("LD_LIBRARY_PATH", "/frozen/app/lib")
    monkeypatch.delenv("LD_LIBRARY_PATH_ORIG", raising=False)
    assert "LD_LIBRARY_PATH" not in installer._system_command_env()

    monkeypatch.setenv("LD_LIBRARY_PATH_ORIG", "/usr/lib:/usr/local/lib")
    assert installer._system_command_env()["LD_LIBRARY_PATH"] == "/usr/lib:/usr/local/lib"


def test_staged_bytes_are_rechecked_before_install(tmp_path, mocked_linux, monkeypatch):
    source = _make_debs_folder(tmp_path / "source")
    service = installer.LibreOfficeInstaller(tmp_path / "app-data")
    _prepare_ready(service, source)
    staged_package = service._packages[0].path
    staged_package.chmod(0o600)
    staged_package.write_text("changed after validation", encoding="utf-8")
    monkeypatch.setattr(installer, "_is_root", lambda: True)
    monkeypatch.setattr(installer.subprocess, "Popen", lambda *_args, **_kwargs: pytest.fail("changed package must not install"))

    service.start()
    result = _wait_for(service, "failed")

    assert "发生变化" in result["message"]
