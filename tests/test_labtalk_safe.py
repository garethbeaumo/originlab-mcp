"""Tests for LabTalk confirm gating and suggested_alternatives."""

from __future__ import annotations

from originlab_mcp.exceptions import NoActiveWorksheetError, WorksheetNotFoundError
from originlab_mcp.utils.labtalk_safe import (
    classify_labtalk_script,
    strip_labtalk_strings_and_comments,
)
from originlab_mcp.utils.validators import error_response, error_response_from_exception
from tests.fakes import DummyMCP


class TestLabTalkClassifier:
    def test_safe_commands_do_not_require_confirm(self):
        requires, reason, alts = classify_labtalk_script('type -a "hello"; window -a Graph1')
        assert requires is False
        assert reason == ""
        assert alts == ()

    def test_keywords_inside_strings_are_ignored(self):
        cleaned = strip_labtalk_strings_and_comments('type -a "please delete this";')
        assert "delete" not in cleaned
        requires, _, _ = classify_labtalk_script('type -a "please delete this";')
        assert requires is False

    def test_destructive_commands_require_confirm(self):
        requires, reason, alts = classify_labtalk_script("doc -n;")
        assert requires is True
        assert reason == "doc -n"
        assert "new_project" in alts

        requires, reason, alts = classify_labtalk_script("win -c Graph1;")
        assert requires is True
        assert "win -c" in reason
        assert alts


class TestSuggestedAlternatives:
    def test_error_response_includes_alternatives(self):
        result = error_response(
            message="missing",
            error_type="not_found",
            target="worksheet",
            suggested_alternatives=["list_worksheets"],
        )
        assert result["error"]["suggested_alternatives"] == ["list_worksheets"]

    def test_tool_error_preserves_alternatives(self):
        result = error_response_from_exception(WorksheetNotFoundError("Sheet9"))
        assert "list_worksheets" in result["error"]["suggested_alternatives"]

        result = error_response_from_exception(NoActiveWorksheetError())
        assert "import_csv" in result["error"]["suggested_alternatives"]


class TestExecuteLabtalkConfirmGate:
    def test_destructive_labtalk_is_blocked_without_confirm(
        self, fresh_manager, fake_origin
    ):
        from originlab_mcp.tools.advanced import register_advanced_tools

        mcp = DummyMCP()
        register_advanced_tools(mcp, fresh_manager)

        blocked = mcp.tools["execute_labtalk"](command="doc -s")
        assert blocked["ok"] is False
        assert blocked["error"]["target"] == "confirm"
        assert blocked["error"]["value"] == "doc -s"
        assert "save_project" in blocked["error"]["suggested_alternatives"]
        assert fake_origin.lt_commands == []

        allowed = mcp.tools["execute_labtalk"](command="doc -s", confirm=True)
        assert allowed["ok"] is True
        assert fake_origin.lt_commands == ["doc -s"]
        assert allowed["data"]["confirmed_risk"] == "doc -s"

    def test_safe_labtalk_runs_without_confirm(self, fresh_manager, fake_origin):
        from originlab_mcp.tools.advanced import register_advanced_tools

        mcp = DummyMCP()
        register_advanced_tools(mcp, fresh_manager)
        result = mcp.tools["execute_labtalk"](command="window -a Graph1")
        assert result["ok"] is True
        assert fake_origin.lt_commands == ["window -a Graph1"]
