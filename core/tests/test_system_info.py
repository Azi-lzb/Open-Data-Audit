# -*- coding: utf-8 -*-
"""本机运行环境摘要（高级设置展示 / S1F1 预检日志）的回归测试。"""

from __future__ import annotations

import re
import sys

import pytest

from base_audit.system_info import system_environment_summary, system_environment_text


@pytest.mark.skipif(sys.platform == "win32", reason="Linux/UOS 分支")
def test_summary_contains_required_fields():
    info = system_environment_summary()
    assert {"os", "glibc", "python", "libreoffice"} <= set(info)
    assert info["os"]
    # Linux 上 glibc 必须是可读版本号或明确的未知提示，不得误报"不适用（Windows）"。
    assert re.fullmatch(r"[0-9]+(\.[0-9]+)+", info["glibc"]) or info["glibc"].startswith("未知")
    assert re.search(r"[0-9]+\.[0-9]+", info["python"])


def test_text_format_is_single_line():
    text = system_environment_text()
    assert "\n" not in text
    assert "系统：" in text and "glibc：" in text and "Python：" in text


@pytest.mark.skipif(sys.platform == "win32", reason="Linux/UOS 分支")
def test_libreoffice_missing_reported(monkeypatch):
    import base_audit.system_info as si
    import base_audit.engines.libreoffice as libreoffice

    monkeypatch.setattr(libreoffice, "find_calc_engine", lambda: None)
    monkeypatch.setattr(si, "_SUMMARY", None)
    info = system_environment_summary()
    assert info["libreoffice"] == "未找到"
