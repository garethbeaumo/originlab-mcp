"""Real originpro save/open return-value contracts and allowlist enforcement."""

import pytest

from originlab_mcp.exceptions import ToolError
from originlab_mcp.tools.export import register_export_tools


@pytest.fixture
def project_tools(dummy_mcp, fresh_manager, fake_origin):
    register_export_tools(dummy_mcp, fresh_manager)
    return dummy_mcp.tools


@pytest.mark.parametrize("explicit_path", [True, False])
def test_failed_save_is_not_success(project_tools, fresh_manager, fake_origin, tmp_path, monkeypatch, explicit_path):
    old = str(tmp_path / "old.opju")
    fresh_manager.project_path = old
    fake_origin.project_path = old
    monkeypatch.setattr(fake_origin, "save", lambda *a, **k: False)
    result = project_tools["save_project"](str(tmp_path / "new.opju") if explicit_path else None)
    assert result["ok"] is False
    assert fresh_manager.project_path == old


def test_failed_open_keeps_context(project_tools, fresh_manager, fake_origin, tmp_path, monkeypatch):
    monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE", "off")
    path = tmp_path / "bad.opju"
    path.touch()
    fresh_manager.active_worksheet = "[Old]Sheet1"
    fresh_manager.active_graph = "OldGraph"
    fresh_manager.project_path = str(tmp_path / "old.opju")
    monkeypatch.setattr(fake_origin, "open", lambda **k: False)
    result = project_tools["open_project"](str(path))
    assert result["ok"] is False
    assert fresh_manager.active_worksheet == "[Old]Sheet1"
    assert fresh_manager.active_graph == "OldGraph"
    assert fresh_manager.project_path.endswith("old.opju")


def test_required_autosave_failure_blocks_destruction(project_tools, fresh_manager, fake_origin, tmp_path, monkeypatch):
    monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_REQUIRED", "1")
    fresh_manager.project_path = str(tmp_path / "work.opju")
    fake_origin.project_path = fresh_manager.project_path
    fake_origin.new_sheet()
    original_book = fake_origin.books[0]
    monkeypatch.setattr(fake_origin, "save", lambda *a, **k: False)
    result = project_tools["new_project"]()
    assert result["ok"] is False
    assert result["error"]["target"] == "autosave"
    assert fake_origin.books[0] is original_book


@pytest.mark.parametrize("periodic", [True, False])
def test_autosave_reports_false_return(fresh_manager, fake_origin, tmp_path, monkeypatch, periodic):
    monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_INTERVAL", "3600")
    fresh_manager.project_path = str(tmp_path / "work.opju")
    fake_origin.project_path = fresh_manager.project_path
    monkeypatch.setattr(fake_origin, "save", lambda *a, **k: False)
    try:
        result = (fresh_manager.periodic_autosave_tick() if periodic else fresh_manager.preflight_autosave("new_project"))
        assert result["saved"] is False
    finally:
        fresh_manager.shutdown()


@pytest.mark.parametrize("operation", ["save", "preflight", "periodic"])
def test_implicit_save_respects_allowed_roots(project_tools, fresh_manager, fake_origin, tmp_path, monkeypatch, operation):
    monkeypatch.setenv("ORIGINLAB_MCP_ALLOWED_ROOTS", str(tmp_path / "allowed"))
    monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_INTERVAL", "3600")
    monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_REQUIRED", "1")
    fresh_manager.project_path = str(tmp_path / "outside.opju")
    fake_origin.project_path = fresh_manager.project_path
    try:
        if operation == "save":
            assert project_tools["save_project"]()["ok"] is False
        elif operation == "preflight":
            with pytest.raises(ToolError):
                fresh_manager.preflight_autosave("new_project")
        else:
            assert fresh_manager.periodic_autosave_tick()["saved"] is False
        assert fake_origin.saved_paths == []
    finally:
        fresh_manager.shutdown()


def test_required_autosave_exception_blocks_destruction(project_tools, fresh_manager, fake_origin, tmp_path, monkeypatch):
    monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_REQUIRED", "1")
    fresh_manager.project_path = str(tmp_path / "work.opju")
    fake_origin.project_path = fresh_manager.project_path
    fake_origin.new_sheet()
    original_book = fake_origin.books[0]

    def fail(*args, **kwargs):
        raise OSError("disk full")

    monkeypatch.setattr(fake_origin, "save", fail)
    result = project_tools["new_project"]()
    assert result["ok"] is False
    assert result["error"]["target"] == "autosave"
    assert fake_origin.books[0] is original_book


def test_explicit_save_outside_roots_is_blocked(project_tools, fresh_manager, fake_origin, tmp_path, monkeypatch):
    monkeypatch.setenv("ORIGINLAB_MCP_ALLOWED_ROOTS", str(tmp_path / "allowed"))
    result = project_tools["save_project"](str(tmp_path / "outside.opju"))
    assert result["ok"] is False
    assert fake_origin.saved_paths == []
    assert fresh_manager.project_path is None


def test_explicit_and_implicit_save_inside_roots(project_tools, fresh_manager, fake_origin, tmp_path, monkeypatch):
    monkeypatch.setenv("ORIGINLAB_MCP_ALLOWED_ROOTS", str(tmp_path))
    path = str(tmp_path / "allowed.opju")
    assert project_tools["save_project"](path)["ok"] is True
    assert project_tools["save_project"]()["ok"] is True
    assert fake_origin.saved_paths == [path, path]
    assert fresh_manager.project_path == path


def test_autosave_uses_live_project_after_ui_switch(fresh_manager, fake_origin, tmp_path):
    fresh_manager.project_path = str(tmp_path / "old.opju")
    fake_origin.project_path = str(tmp_path / "new.opju")
    status = fresh_manager.preflight_autosave("new_project")
    assert status["saved"] is True
    assert fake_origin.saved_paths == [str(tmp_path / "new.opju")]
    assert fresh_manager.project_path == str(tmp_path / "new.opju")


def test_unknown_live_project_does_not_reuse_cached_destination(fresh_manager, fake_origin, tmp_path, monkeypatch):
    monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_REQUIRED", "1")
    fresh_manager.project_path = str(tmp_path / "old.opju")
    fake_origin.project_path = None
    with pytest.raises(ToolError):
        fresh_manager.preflight_autosave("new_project")
    assert fake_origin.saved_paths == []
    assert fresh_manager.project_path is None


def test_conflicting_live_paths_block_autosave(fresh_manager, fake_origin, tmp_path, monkeypatch):
    monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_REQUIRED", "1")
    fake_origin.project_path = str(tmp_path / "current.opju")
    fake_origin.lt_strings["filename$"] = str(tmp_path / "different.opju")
    with pytest.raises(ToolError):
        fresh_manager.preflight_autosave("new_project")
    assert fake_origin.saved_paths == []


def test_reconnect_invalidates_old_context(fresh_manager, fake_origin, tmp_path):
    fresh_manager.project_path = str(tmp_path / "old.opju")
    fresh_manager.active_worksheet = "OldSheet"
    fresh_manager.active_graph = "OldGraph"
    fresh_manager._connected = False
    fresh_manager.connect()
    assert fresh_manager.project_path is None
    assert fresh_manager.active_worksheet is None
    assert fresh_manager.active_graph is None


def test_release_invalidates_old_context(fresh_manager, fake_origin, tmp_path):
    fresh_manager.project_path = str(tmp_path / "old.opju")
    fresh_manager.active_worksheet = "OldSheet"
    fresh_manager.active_graph = "OldGraph"
    assert fresh_manager.release()
    assert fresh_manager.project_path is None
    assert fresh_manager.active_worksheet is None
    assert fresh_manager.active_graph is None
