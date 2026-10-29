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
from tests.fakes import DummyMCP, FakeOrigin, attach_fake_origin

MIXED_SCRIPTS = [
    'python -s "print(1)"; doc -n;',
    'run.python("prepare.py"); win -c Graph1;',
    "%(cmd$); del col(A);",
]


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
        assert policy.interval_seconds == 300.0

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
        assert (
            AutosavePolicy.from_env({"ORIGINLAB_MCP_AUTOSAVE_INTERVAL": "off"}).interval_seconds
            is None
        )
        assert (
            AutosavePolicy.from_env({"ORIGINLAB_MCP_AUTOSAVE_INTERVAL": "60"}).interval_seconds
            == 60.0
        )
        # Master switch off also disables periodic interval.
        assert (
            AutosavePolicy.from_env(
                {
                    "ORIGINLAB_MCP_AUTOSAVE": "off",
                    "ORIGINLAB_MCP_AUTOSAVE_INTERVAL": "60",
                }
            ).interval_seconds
            is None
        )

    def test_labtalk_destructive_gate(self) -> None:
        assert should_autosave_labtalk("doc -n;", confirm=True)
        assert should_autosave_labtalk("del col(A);", confirm=True)
        assert should_autosave_labtalk("win -c Graph1;", confirm=True)
        assert not should_autosave_labtalk("doc -n;", confirm=False)
        assert not should_autosave_labtalk('type "hello";', confirm=True)
        assert not should_autosave_labtalk("doc -s;", confirm=True)

    @pytest.mark.parametrize("prefix", ['python -s "print(1)";', 'run.python("prepare.py");', "%(cmd$);"])
    @pytest.mark.parametrize("destructive", ["doc -n;", "win -c Graph1;", "del col(A);"])
    @pytest.mark.parametrize("destructive_first", [False, True])
    def test_all_script_risks_are_checked_for_autosave(self, prefix, destructive, destructive_first):
        script = destructive + prefix if destructive_first else prefix + destructive
        assert should_autosave_labtalk(script, confirm=True)
        assert not should_autosave_labtalk(script, confirm=False)

    @pytest.mark.parametrize("script", [
        'python -s "doc -n;";',
        'run.python("delete.py"); /* win -c Graph1; */',
        "%(cmd$); // del col(A);",
        "doc -s;",
        "",
    ])
    def test_nondestructive_scripts_do_not_autosave(self, script):
        assert not should_autosave_labtalk(script, confirm=True)

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
    @pytest.mark.parametrize("command", MIXED_SCRIPTS)
    @pytest.mark.parametrize("save_outcome", ["success", "false", "error"])
    def test_mixed_labtalk_requires_save_before_execution(
        self, monkeypatch, manager_with_fake, tmp_path, command, save_outcome
    ):
        from originlab_mcp.tools.advanced import register_advanced_tools

        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE", "1")
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_REQUIRED", "1")
        manager, origin = manager_with_fake
        path = str(tmp_path / "mixed.opju")
        origin.project_path = path
        events = []
        save_ok = save_outcome == "success"

        def save(file_path):
            events.append(("save", file_path))
            if save_outcome == "error":
                raise OSError("disk full")
            return save_ok

        def execute(script):
            events.append(("execute", script))

        monkeypatch.setattr(origin, "save", save)
        monkeypatch.setattr(origin, "lt_exec", execute)
        mcp = DummyMCP()
        register_advanced_tools(mcp, manager)
        blocked = mcp.tools["execute_labtalk"](command=command)
        assert blocked["ok"] is False
        assert blocked["error"]["target"] == "confirm"
        assert events == []
        result = mcp.tools["execute_labtalk"](command=command, confirm=True)
        assert result["ok"] is save_ok
        assert events == ([("save", path), ("execute", command)] if save_ok else [("save", path)])
        if save_ok:
            assert result["data"]["autosave"]["saved"] is True
        else:
            assert result["error"]["target"] == "autosave"

    @pytest.mark.parametrize("command", MIXED_SCRIPTS)
    def test_mixed_labtalk_is_blocked_without_required_save_path(
        self, monkeypatch, manager_with_fake, command
    ):
        from originlab_mcp.tools.advanced import register_advanced_tools

        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE", "1")
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_REQUIRED", "1")
        manager, origin = manager_with_fake
        mcp = DummyMCP()
        register_advanced_tools(mcp, manager)
        result = mcp.tools["execute_labtalk"](command=command, confirm=True)
        assert result["ok"] is False
        assert origin.lt_commands == []
        assert origin.saved_paths == []

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
        origin.project_path = path
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


class TestPeriodicAutosave:
    def test_tick_saves_when_path_known(
        self, monkeypatch: pytest.MonkeyPatch, manager_with_fake, tmp_path
    ) -> None:
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE", "1")
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_INTERVAL", "60")
        manager, origin = manager_with_fake
        path = str(tmp_path / "periodic.opju")
        manager.project_path = path
        origin.project_path = path
        status = manager.periodic_autosave_tick()
        assert status["saved"] is True
        assert status["path"] == path
        assert origin.saved_paths == [path]

    def test_tick_skips_without_path(
        self, monkeypatch: pytest.MonkeyPatch, manager_with_fake
    ) -> None:
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_INTERVAL", "60")
        manager, origin = manager_with_fake
        status = manager.periodic_autosave_tick()
        assert status["saved"] is False
        assert "no project path" in status["message"]
        assert origin.saved_paths == []

    def test_tick_disabled_by_interval_off(
        self, monkeypatch: pytest.MonkeyPatch, manager_with_fake, tmp_path
    ) -> None:
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_INTERVAL", "off")
        manager, origin = manager_with_fake
        manager.project_path = str(tmp_path / "x.opju")
        status = manager.periodic_autosave_tick()
        assert status["message"] == "periodic autosave disabled"
        assert origin.saved_paths == []

    def test_setting_path_schedules_timer(
        self, monkeypatch: pytest.MonkeyPatch, manager_with_fake, tmp_path
    ) -> None:
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE", "1")
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_INTERVAL", "3600")
        manager, _origin = manager_with_fake
        manager.project_path = str(tmp_path / "sched.opju")
        assert manager._autosave_timer is not None
        manager.project_path = None
        assert manager._autosave_timer is None

    def test_get_info_reports_interval(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_INTERVAL", "120")
        info = OriginManager(idle_timeout=0).get_info()
        assert info["autosave_interval"] == 120.0
        assert info["autosave_enabled"] is True
