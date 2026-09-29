# -*- coding: utf-8 -*-
"""计算引擎预检（S1F1 前置探测）的回归测试。

覆盖三种现场：未安装 LibreOffice、已安装但与系统 glibc 不匹配（能找到、
起不动）、正常可用。
"""

from __future__ import annotations

import sys

import pytest

from base_audit.engines import calculation_engine_preflight


@pytest.mark.skipif(sys.platform == "win32", reason="Linux/UOS 分支探测")
def test_preflight_reports_missing_libreoffice(monkeypatch):
    import base_audit.engines.libreoffice as libreoffice

    monkeypatch.setattr(libreoffice, "find_calc_engine", lambda: None)
    ok, message = calculation_engine_preflight()
    assert ok is False
    assert "未找到 LibreOffice" in message
    assert "README" in message and "BASE_AUDIT_SOFFICE" in message


@pytest.mark.skipif(sys.platform == "win32", reason="Linux/UOS 分支探测")
def test_preflight_reports_launch_failure(monkeypatch):
    import base_audit.engines.libreoffice as libreoffice

    engine = libreoffice.CalcEngine(
        "LibreOffice Calc", "测试", ("/nonexistent/engine/probe",), "测试入口")
    monkeypatch.setattr(libreoffice, "find_calc_engine", lambda: engine)
    ok, message = calculation_engine_preflight()
    assert ok is False
    assert "启动失败" in message


@pytest.mark.skipif(sys.platform == "win32", reason="Linux/UOS 分支探测")
def test_preflight_reports_nonzero_exit(monkeypatch):
    import base_audit.engines.libreoffice as libreoffice

    engine = libreoffice.CalcEngine("LibreOffice Calc", "测试", ("false",), "测试入口")
    monkeypatch.setattr(libreoffice, "find_calc_engine", lambda: engine)
    ok, message = calculation_engine_preflight()
    assert ok is False
    assert "无法启动" in message
    assert "glibc" in message


@pytest.mark.skipif(sys.platform == "win32", reason="Linux/UOS 分支探测")
def test_preflight_passes_when_engine_runs(monkeypatch):
    import base_audit.engines.libreoffice as libreoffice

    engine = libreoffice.CalcEngine("LibreOffice Calc", "测试", ("echo",), "测试入口")
    monkeypatch.setattr(libreoffice, "find_calc_engine", lambda: engine)
    ok, message = calculation_engine_preflight()
    assert ok is True
    assert "计算引擎就绪" in message
