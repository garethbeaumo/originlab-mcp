"""Autosave policy helpers and OriginManager preflight behavior."""

from __future__ import annotations

import pytest

from originlab_mcp.exceptions import ToolError
from originlab_mcp.origin_manager import OriginManager
from originlab_mcp.utils.autosave import (
    DESTRUCTIVE_TOOLS,
    AutosavePolicy,
    collect_autosave_warnings,
    should_autosave_labtalk,
    should_autosave_tool,
)
from tests.fakes import FakeOrigin, attach_fake_origin


@pytest.fixture
def manager_with_fake():
    manager = OriginManager(auto_recover_active=False, idle_timeout=0)
    origin = FakeOrigin()
    attach_fake_origin(manager, origin)
    return manager, origin


class TestPolicyHelpers:
    def test_policy_from_env_defaults(self) -> None:
        policy = AutosavePolicy.from_env({})
        assert policy.enabled is True
        assert policy.required is False

    def test_policy_aliases(self) -> None:
        assert AutosavePolicy.from_env({"ORIGINLAB_MCP_AUTOSAVE": "off"}).enabled is False
        assert AutosavePolicy.from_env({"ORIGINLAB_MCP_AUTOSAVE": "0"}).enabled is False
        assert AutosavePolicy.from_env({"ORIGINLAB_MCP_AUTOSAVE": "false"}).enabled is False
        assert AutosavePolicy.from_env({"ORIGINLAB_MCP_AUTOSAVE": "warn"}).enabled is True
        assert (
            AutosavePolicy.from_env({"ORIGINLAB_MCP_AUTOSAVE_REQUIRED": "1"}).required
            is True
        )
        assert (
            AutosavePolicy.from_env({"ORIGINLAB_MCP_AUTOSAVE_REQUIRED": "true"}).required
            is True
        )

    def test_labtalk_destructive_gate(self) -> None:
        assert should_autosave_labtalk("doc -n;", confirm=True)
        assert should_autosave_labtalk("del col(A);", confirm=True)
        assert should_autosave_labtalk("win -c Graph1;", confirm=True)
        assert not should_autosave_labtalk("doc -n;", confirm=False)
        assert not should_autosave_labtalk('type "hello";', confirm=True)
        assert not should_autosave_labtalk("doc -s;", confirm=True)

    def test_collect_warnings_from_status(self) -> None:
        assert collect_autosave_warnings(
            {
                "attempted": False,
                "saved": False,
                "path": None,
                "message": "no project path; autosave skipped",
            }
        ) == ["no project path; autosave skipped"]
        assert collect_autosave_warnings(
            {
                "attempted": True,
                "saved": True,
                "path": "/tmp/a.opju",
                "message": "preflight autosave (clear_worksheet) -> /tmp/a.opju",
            }
        ) == ["preflight autosave (clear_worksheet) -> /tmp/a.opju"]
        assert (
            collect_autosave_warnings(
                {
                    "attempted": False,
                    "saved": False,
                    "path": None,
                    "message": "autosave disabled",
                }
            )
            == []
        )

    def test_destructive_tool_registry(self) -> None:
        for name in (
            "new_project",
            "clear_worksheet",
            "delete_columns",
            "close_origin",
            "open_project",
            "remove_plot_from_graph",
            "remove_graph_label",
        ):
            assert should_autosave_tool(name)
            assert name in DESTRUCTIVE_TOOLS
        assert not should_autosave_tool("save_project")


class TestPreflightAutosave:
    def test_warn_without_path_returns_skip_status(
        self, monkeypatch: pytest.MonkeyPatch, manager_with_fake
    ) -> None:
        monkeypatch.delenv("ORIGINLAB_MCP_AUTOSAVE_REQUIRED", raising=False)
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE", "warn")
        manager, origin = manager_with_fake
        status = manager.preflight_autosave("clear_worksheet")
        assert status["attempted"] is False
        assert status["saved"] is False
        assert "no project path" in status["message"]
        assert origin.saved_paths == []

    def test_required_without_path_raises(
        self, monkeypatch: pytest.MonkeyPatch, manager_with_fake
    ) -> None:
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE", "1")
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_REQUIRED", "1")
        manager, _origin = manager_with_fake
        with pytest.raises(ToolError, match="项目路径"):
            manager.preflight_autosave("new_project")

    def test_saves_when_path_known(
        self, monkeypatch: pytest.MonkeyPatch, manager_with_fake, tmp_path
    ) -> None:
        monkeypatch.delenv("ORIGINLAB_MCP_AUTOSAVE_REQUIRED", raising=False)
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE", "1")
        path = str(tmp_path / "autosave.opju")
        manager, origin = manager_with_fake
        manager.project_path = path
        status = manager.preflight_autosave("delete_columns")
        assert status["saved"] is True
        assert status["path"] == path
        assert origin.saved_paths == [path]

    def test_off_skips_even_with_path(
        self, monkeypatch: pytest.MonkeyPatch, manager_with_fake, tmp_path
    ) -> None:
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE", "off")
        manager, origin = manager_with_fake
        manager.project_path = str(tmp_path / "x.opju")
        status = manager.preflight_autosave("close_origin")
        assert status["message"] == "autosave disabled"
        assert status["saved"] is False
        assert origin.saved_paths == []

    def test_required_env_alias_blocks_without_path(
        self, monkeypatch: pytest.MonkeyPatch, manager_with_fake
    ) -> None:
        monkeypatch.delenv("ORIGINLAB_MCP_AUTOSAVE", raising=False)
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_REQUIRED", "1")
        manager, _origin = manager_with_fake
        with pytest.raises(ToolError, match="项目路径"):
            manager.preflight_autosave("open_project")
