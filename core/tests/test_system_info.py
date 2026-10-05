# -*- coding: utf-8 -*-
"""本机运行环境摘要（高级设置展示 / S1F1 预检日志）的回归测试。"""

from __future__ import annotations

import re
import sys

import pytest

from base_audit.system_info import system_environment_summary, system_environment_text

# 环境面板全部字段；Windows 专属项在 Linux 上报告"不适用"，openpyxl/硬件/产品/区域编码跨平台。
ALL_FIELDS = {
    "app", "os", "glibc", "python", "libreoffice", "locale",
    "excel", "wps", "com", "pywin32", "webview2", "openpyxl", "hardware",
}


@pytest.mark.skipif(sys.platform == "win32", reason="Linux/UOS 分支")
def test_summary_contains_required_fields():
    info = system_environment_summary()
    assert {"os", "glibc", "python", "libreoffice"} <= set(info)
    assert info["os"]
    # Linux 上 glibc 必须是可读版本号或明确的未知提示，不得误报"不适用（Windows）"。
    assert re.fullmatch(r"[0-9]+(\.[0-9]+)+", info["glibc"]) or info["glibc"].startswith("未知")
    assert re.search(r"[0-9]+\.[0-9]+", info["python"])


def test_summary_contains_all_fields():
    info = system_environment_summary()
    assert ALL_FIELDS <= set(info)


def test_text_format_is_single_line():
    text = system_environment_text()
    assert "\n" not in text
    assert "系统：" in text and "glibc：" in text and "Python：" in text


def test_text_covers_engine_and_shell_fields():
    text = system_environment_text()
    for label in (
        "产品：", "区域编码：", "LibreOffice：", "Excel：", "WPS：", "COM接管：",
        "pywin32：", "WebView2：", "openpyxl：", "硬件：",
    ):
        assert label in text


def test_app_field_reports_version_build_and_shell():
    info = system_environment_summary()
    # 版本来自 VERSION_MANIFEST（V年份.大版本.中版本.小版本）；清单不可用时明示。
    assert re.search(r"V\d+\.\d+\.\d+\.\d+", info["app"]) or "版本清单不可用" in info["app"]
    assert "源码运行" in info["app"] or "打包 exe" in info["app"]


def test_locale_field_reports_encoding():
    info = system_environment_summary()
    assert "Python偏好编码" in info["locale"]
    if sys.platform == "win32":
        assert "ANSI CP" in info["locale"] or "ANSI UTF-8" in info["locale"]


def test_openpyxl_reports_real_version():
    info = system_environment_summary()
    assert re.search(r"[0-9]+(\.[0-9]+)+", info["openpyxl"])


@pytest.mark.skipif(sys.platform != "win32", reason="Windows 分支")
def test_windows_engine_fields_report_real_values():
    info = system_environment_summary()
    # S1F1 的 Windows 引擎前提：Excel 与 WPS 至少装有其一。
    assert not (info["excel"].startswith("未找到") and info["wps"].startswith("未找到"))
    assert not info["excel"].startswith("不适用") and not info["wps"].startswith("不适用")
    # COM 接管首显注册表解析结果（实启结果由后台线程覆盖），不得留空。
    assert info["com"].startswith(("注册表解析到", "实启解析到", "COM 组件不可用", "实启失败", "未找到"))
    # pywin32 的版本号形如 312（无点号），WebView2 形如 154.0.4258.53。
    assert re.search(r"[0-9]+(\.[0-9]+)*", info["pywin32"]) or info["pywin32"].startswith("未安装")
    assert re.search(r"[0-9]+(\.[0-9]+)+", info["webview2"]) or info["webview2"].startswith("未找到")
    assert info["hardware"]


@pytest.mark.skipif(sys.platform != "win32", reason="Windows 分支")
def test_progid_resolution_reads_real_server():
    from base_audit.system_info import _resolve_progid_server
    server = _resolve_progid_server("Excel.Application")
    assert server
    upper = server.upper()
    assert "EXCEL.EXE" in upper or "KINGSOFT" in upper or "WPS" in upper


@pytest.mark.skipif(sys.platform == "win32", reason="Linux/UOS 分支")
def test_windows_only_fields_report_not_applicable_on_linux():
    info = system_environment_summary()
    for key in ("excel", "com", "pywin32", "webview2"):
        assert info[key].startswith("不适用")
    # WPS 在 Linux 上有对应版本，报真实状态而非不适用；openpyxl 跨平台报版本。
    assert info["wps"]
    assert re.search(r"[0-9]+(\.[0-9]+)+", info["openpyxl"])


@pytest.mark.skipif(sys.platform == "win32", reason="Linux/UOS 分支")
def test_libreoffice_missing_reported(monkeypatch):
    import base_audit.system_info as si
    import base_audit.engines.libreoffice as libreoffice

    monkeypatch.setattr(libreoffice, "find_calc_engine", lambda: None)
    monkeypatch.setattr(si, "_SUMMARY", None)
    info = system_environment_summary()
    assert info["libreoffice"] == "未找到"
