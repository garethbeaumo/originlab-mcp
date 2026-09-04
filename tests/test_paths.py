"""Allowed-roots path sandbox helpers and tool integration."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from originlab_mcp.origin_manager import OriginManager
from originlab_mcp.tools.data import register_data_tools
from originlab_mcp.tools.export import register_export_tools
from originlab_mcp.tools.plot import register_plot_tools
from originlab_mcp.utils.paths import (
    check_allowed_path,
    is_under_root,
    parse_allowed_roots,
    resolve_user_path,
)
from originlab_mcp.utils.validators import validate_file_path, validate_output_path
from tests.fakes import DummyMCP


class TestAllowedRootsPolicy:
    def test_unset_means_unrestricted(self) -> None:
        assert parse_allowed_roots({}) is None
        assert parse_allowed_roots({"ORIGINLAB_MCP_ALLOWED_ROOTS": ""}) is None
        assert check_allowed_path("/etc/passwd", environ={}) is None

    def test_parses_pathsep_and_comma(self, tmp_path: Path) -> None:
        a = tmp_path / "a"
        b = tmp_path / "b"
        a.mkdir()
        b.mkdir()
        roots = parse_allowed_roots(
            {"ORIGINLAB_MCP_ALLOWED_ROOTS": f"{a}{os.pathsep}{b}"}
        )
        assert roots is not None
        assert len(roots) == 2

        roots_comma = parse_allowed_roots(
            {"ORIGINLAB_MCP_ALLOWED_ROOTS": f"{a},{b}"}
        )
        assert roots_comma is not None
        assert len(roots_comma) == 2

    def test_blocks_outside_and_allows_inside(self, tmp_path: Path) -> None:
        allowed = tmp_path / "safe"
        allowed.mkdir()
        outside = tmp_path / "other" / "file.csv"
        outside.parent.mkdir()
        outside.write_text("x,y\n1,2\n", encoding="utf-8")
        inside = allowed / "data.csv"
        inside.write_text("x,y\n1,2\n", encoding="utf-8")

        env = {"ORIGINLAB_MCP_ALLOWED_ROOTS": str(allowed)}
        assert check_allowed_path(str(inside), environ=env) is None
        err = check_allowed_path(str(outside), environ=env)
        assert err is not None
        assert "允许的根" in err

    def test_blocks_traversal_escape(self, tmp_path: Path) -> None:
        allowed = tmp_path / "safe"
        allowed.mkdir()
        secret = tmp_path / "secret.txt"
        secret.write_text("nope", encoding="utf-8")
        sneaky = allowed / ".." / "secret.txt"
        env = {"ORIGINLAB_MCP_ALLOWED_ROOTS": str(allowed)}
        err = check_allowed_path(str(sneaky), environ=env)
        assert err is not None

    def test_is_under_root(self, tmp_path: Path) -> None:
        root = resolve_user_path(str(tmp_path))
        child = resolve_user_path(str(tmp_path / "a" / "b.csv"))
        assert is_under_root(child, root)
        assert is_under_root(root, root)
        assert not is_under_root(resolve_user_path(str(tmp_path.parent)), root)


class TestValidatorsAllowlist:
    def test_validate_file_path_respects_roots(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        allowed = tmp_path / "ok"
        allowed.mkdir()
        good = allowed / "a.csv"
        good.write_text("a\n1\n", encoding="utf-8")
        bad_dir = tmp_path / "bad"
        bad_dir.mkdir()
        bad = bad_dir / "a.csv"
        bad.write_text("a\n1\n", encoding="utf-8")

        monkeypatch.setenv("ORIGINLAB_MCP_ALLOWED_ROOTS", str(allowed))
        assert validate_file_path(str(good)) is None
        assert validate_file_path(str(bad)) is not None

    def test_validate_output_path_for_new_file(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        allowed = tmp_path / "ok"
        allowed.mkdir()
        monkeypatch.setenv("ORIGINLAB_MCP_ALLOWED_ROOTS", str(allowed))
        assert validate_output_path(str(allowed / "out.png")) is None
        assert validate_output_path(str(tmp_path / "nope" / "out.png")) is not None


class TestToolAllowlistIntegration:
    def test_import_csv_blocked_outside_roots(
        self, fresh_manager, fake_origin, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        allowed = tmp_path / "allowed"
        allowed.mkdir()
        blocked = tmp_path / "blocked" / "x.csv"
        blocked.parent.mkdir()
        blocked.write_text("A,B\n1,2\n", encoding="utf-8")
        monkeypatch.setenv("ORIGINLAB_MCP_ALLOWED_ROOTS", str(allowed))

        mcp = DummyMCP()
        register_data_tools(mcp, fresh_manager)
        result = mcp.tools["import_csv"](file_path=str(blocked))
        assert result["ok"] is False
        assert "允许的根" in result["message"]
        assert result["error"]["target"] == "file_path"

    def test_export_graph_blocked_outside_roots(
        self, fresh_manager, fake_origin, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        allowed = tmp_path / "allowed"
        allowed.mkdir()
        monkeypatch.setenv("ORIGINLAB_MCP_ALLOWED_ROOTS", str(allowed))

        mcp = DummyMCP()
        register_data_tools(mcp, fresh_manager)
        register_export_tools(mcp, fresh_manager)
        register_plot_tools(mcp, fresh_manager)
        mcp.tools["import_data_from_text"](data="X,Y\n1,2\n3,4\n")
        mcp.tools["create_plot"](x_col=0, y_cols=1)

        blocked_out = tmp_path / "blocked" / "chart.png"
        result = mcp.tools["export_graph"](output_path=str(blocked_out))
        assert result["ok"] is False
        assert result["error"]["target"] == "output_path"

        ok_out = allowed / "chart.png"
        exported = mcp.tools["export_graph"](output_path=str(ok_out))
        assert exported["ok"] is True
        assert ok_out.exists()

    def test_get_info_reports_roots(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        monkeypatch.setenv("ORIGINLAB_MCP_ALLOWED_ROOTS", str(tmp_path))
        info = OriginManager(idle_timeout=0).get_info()
        assert info["allowed_roots"] == [str(tmp_path.resolve())]
