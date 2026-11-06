from __future__ import annotations

import os
from pathlib import Path

from base_audit.native.linux_file_picker import (
    _directory_bytes,
    _file_uri_to_path,
    _filters,
    portal_available,
)


def test_file_uri_is_decoded_once_and_preserves_unicode() -> None:
    selected = _file_uri_to_path("file:///tmp/%E6%8A%A5%E8%A1%A8%20%E7%94%B2.xlsx")
    assert Path(selected).name == "报表 甲.xlsx"


def test_file_uri_rejects_non_local_schemes() -> None:
    import pytest

    from base_audit.native.linux_file_picker import LinuxPortalSelectionError

    with pytest.raises(LinuxPortalSelectionError):
        _file_uri_to_path("smb://server/share/input.xlsx")


def test_directory_bytes_uses_existing_folder_and_nul_terminator(tmp_path: Path) -> None:
    raw = _directory_bytes(str(tmp_path))
    assert raw.endswith(b"\0")
    assert os.fsdecode(raw[:-1]) == str(tmp_path)


def test_filters_split_globs_for_portal() -> None:
    assert _filters((("Excel 文件", "*.xlsx *.xlsm"), ("空过滤器", ""))) == [
        ["Excel 文件", [[0, "*.xlsx"], [0, "*.xlsm"]]],
    ]


def test_portal_is_unavailable_without_session_bus(monkeypatch) -> None:
    monkeypatch.delenv("DBUS_SESSION_BUS_ADDRESS", raising=False)
    assert portal_available() is False
