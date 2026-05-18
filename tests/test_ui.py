"""Tests for the local diagnostics UI helpers."""

import json
import sys

from originlab_mcp import ui


def test_configure_client_creates_cursor_config(tmp_path, monkeypatch):
    monkeypatch.setattr(ui, "_repo_root", lambda: tmp_path)

    result = ui.configure_client("cursor")

    path = tmp_path / ".cursor" / "mcp.json"
    data = json.loads(path.read_text(encoding="utf-8"))

    assert result["ok"] is True
    assert result["created"] is True
    assert data["mcpServers"]["originlab"] == {
        "command": sys.executable,
        "args": ["-m", "originlab_mcp.server"],
    }


def test_configure_client_preserves_existing_servers(tmp_path, monkeypatch):
    monkeypatch.setattr(ui, "_repo_root", lambda: tmp_path)
    path = tmp_path / ".trae" / "mcp.json"
    path.parent.mkdir(parents=True)
    path.write_text(
        json.dumps(
            {
                "mcpServers": {
                    "other": {"command": "other-tool"},
                },
                "custom": True,
            },
        ),
        encoding="utf-8",
    )

    result = ui.configure_client("trae")
    data = json.loads(path.read_text(encoding="utf-8"))

    assert result["ok"] is True
    assert result["created"] is False
    assert result["changed"] is True
    assert data["custom"] is True
    assert data["mcpServers"]["other"] == {"command": "other-tool"}
    assert data["mcpServers"]["originlab"]["args"] == ["-m", "originlab_mcp.server"]
    assert list(path.parent.glob("mcp.json.bak-*"))


def test_detect_client_configs_reports_codex_and_trae(tmp_path, monkeypatch):
    monkeypatch.setattr(ui, "_repo_root", lambda: tmp_path)
    ui.configure_client("codex")
    ui.configure_client("trae")

    statuses = {item.id: item for item in ui.detect_client_configs()}

    assert statuses["codex"].configured is True
    assert statuses["codex"].path.endswith(".codex\\config.json") or statuses[
        "codex"
    ].path.endswith(".codex/config.json")
    assert statuses["trae"].configured is True
    assert statuses["trae"].path.endswith(".trae\\mcp.json") or statuses[
        "trae"
    ].path.endswith(".trae/mcp.json")


def test_configure_client_rejects_invalid_json(tmp_path, monkeypatch):
    monkeypatch.setattr(ui, "_repo_root", lambda: tmp_path)
    path = tmp_path / ".cursor" / "mcp.json"
    path.parent.mkdir(parents=True)
    path.write_text("{bad json", encoding="utf-8")

    result = ui.configure_client("cursor")

    assert result["ok"] is False
    assert "valid JSON" in result["message"]
    assert path.read_text(encoding="utf-8") == "{bad json"
