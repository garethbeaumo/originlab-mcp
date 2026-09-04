"""Soft COM dispatch timeout helpers and OriginManager integration."""

from __future__ import annotations

import time

import pytest

from originlab_mcp.exceptions import ToolError
from originlab_mcp.origin_manager import OriginManager
from originlab_mcp.tools.advanced import register_advanced_tools
from originlab_mcp.utils.dispatch import (
    parse_dispatch_timeout_seconds,
    resolve_dispatch_timeout,
    run_with_soft_timeout,
)
from tests.fakes import DummyMCP, FakeOrigin


class TestDispatchPolicy:
    def test_default_is_ninety_seconds(self) -> None:
        assert parse_dispatch_timeout_seconds({}) == 90.0

    def test_off_aliases(self) -> None:
        assert parse_dispatch_timeout_seconds({"ORIGINLAB_MCP_DISPATCH_TIMEOUT": "off"}) is None
        assert parse_dispatch_timeout_seconds({"ORIGINLAB_MCP_DISPATCH_TIMEOUT": "0"}) is None
        assert parse_dispatch_timeout_seconds({"ORIGINLAB_MCP_DISPATCH_TIMEOUT": "false"}) is None

    def test_custom_budget(self) -> None:
        assert parse_dispatch_timeout_seconds(
            {"ORIGINLAB_MCP_DISPATCH_TIMEOUT": "12.5"}
        ) == 12.5

    def test_resolve_override(self) -> None:
        assert resolve_dispatch_timeout(5.0, environ={}) == 5.0
        assert resolve_dispatch_timeout(0, environ={}) is None
        assert (
            resolve_dispatch_timeout(
                environ={"ORIGINLAB_MCP_DISPATCH_TIMEOUT": "off"}
            )
            is None
        )


class TestSoftTimeoutRunner:
    def test_completes_within_budget(self) -> None:
        assert run_with_soft_timeout(1.0, lambda: 42) == 42

    def test_raises_structured_timeout(self) -> None:
        with pytest.raises(ToolError) as exc_info:
            run_with_soft_timeout(0.05, lambda: time.sleep(1.0))
        err = exc_info.value
        assert err.error_type == "timeout"
        assert err.target == "dispatch_timeout"
        assert "get_origin_info" in err.suggested_alternatives


class TestManagerDispatchTimeout:
    def test_execute_soft_timeout(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("ORIGINLAB_MCP_DISPATCH_TIMEOUT", "0.1")
        manager = OriginManager(auto_recover_active=False, idle_timeout=0)
        origin = FakeOrigin()
        manager._op = origin
        manager._connected = True

        def _slow(op):
            time.sleep(1.0)
            return "done"

        with pytest.raises(ToolError, match="未响应"):
            manager.execute(_slow)

    def test_per_call_disable(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("ORIGINLAB_MCP_DISPATCH_TIMEOUT", "0.05")
        manager = OriginManager(auto_recover_active=False, idle_timeout=0)
        manager._op = FakeOrigin()
        manager._connected = True

        def _slow(op):
            time.sleep(0.15)
            return "ok"

        assert manager.execute(_slow, dispatch_timeout=0) == "ok"

    def test_get_info_reports_budget(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("ORIGINLAB_MCP_DISPATCH_TIMEOUT", "45")
        manager = OriginManager(auto_recover_active=False, idle_timeout=0)
        info = manager.get_info()
        assert info["dispatch_timeout"] == 45.0


class TestLabTalkTimeoutOverride:
    def test_tool_timeout_overrides_env_off(
        self, fresh_manager, fake_origin, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setenv("ORIGINLAB_MCP_DISPATCH_TIMEOUT", "off")
        mcp = DummyMCP()
        register_advanced_tools(mcp, fresh_manager)

        def _slow(command: str):
            time.sleep(1.0)
            return None

        fake_origin.lt_exec = _slow  # type: ignore[method-assign]
        result = mcp.tools["execute_labtalk"](
            command="window -a Graph1",
            timeout=0.1,
        )
        assert result["ok"] is False
        assert result["error"]["type"] == "timeout"
        assert result["error"]["target"] == "dispatch_timeout"
