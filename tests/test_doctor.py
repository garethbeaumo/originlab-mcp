"""originlab_doctor diagnostics helpers and tool integration."""

from __future__ import annotations

import pytest

from originlab_mcp.tools.system import register_system_tools
from originlab_mcp.utils.doctor import build_doctor_report
from tests.fakes import DummyMCP


class TestBuildDoctorReport:
    def test_linux_without_originpro_is_degraded_not_fatal(self) -> None:
        report = build_doctor_report(None, ping_origin=False)
        assert report["overall"] in {"ok", "degraded", "error"}
        names = {c["name"]: c for c in report["checks"]}
        assert "platform" in names
        assert "python" in names
        assert "originpro" in names
        assert "autosave" in names
        assert "dispatch_timeout" in names
        assert "allowed_roots" in names
        assert report["policy"]["autosave_enabled"] is True

    def test_reports_allowed_roots(
        self, tmp_path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setenv("ORIGINLAB_MCP_ALLOWED_ROOTS", str(tmp_path))
        report = build_doctor_report(None)
        assert report["policy"]["allowed_roots"] == [str(tmp_path.resolve())]
        roots_check = next(c for c in report["checks"] if c["name"] == "allowed_roots")
        assert str(tmp_path.resolve()) in roots_check["detail"]

    def test_ping_origin_with_fake(
        self, fresh_manager, fake_origin, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_INTERVAL", "off")
        report = build_doctor_report(fresh_manager, ping_origin=True)
        assert report["connected"] is True
        assert report["origin_ping"] is not None
        assert report["origin_ping"]["ok"] is True
        ping = next(c for c in report["checks"] if c["name"] == "origin_ping")
        assert ping["status"] == "ok"


class TestDoctorTool:
    def test_tool_returns_standard_shape(self, fresh_manager, fake_origin) -> None:
        mcp = DummyMCP()
        register_system_tools(mcp, fresh_manager)
        result = mcp.tools["originlab_doctor"]()
        assert result["ok"] is True
        assert "overall" in result["data"]
        assert "checks" in result["data"]
        assert result["next_suggestions"]

    def test_tool_ping_origin(self, fresh_manager, fake_origin) -> None:
        mcp = DummyMCP()
        register_system_tools(mcp, fresh_manager)
        result = mcp.tools["originlab_doctor"](ping_origin=True)
        assert result["ok"] is True
        assert result["data"]["origin_ping"]["ok"] is True
