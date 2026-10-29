"""Tests for LabTalk confirm gating and suggested_alternatives."""

from __future__ import annotations

import pytest

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

    def test_external_code_requires_confirm(self):
        for script, label in [
            ('python -s "import os"', "python"),
            ('run.python("x.py")', "run.<method>"),
            ("run.file(a.ogs);", "run.<method>"),
            ("run.section(a.ogs, main);", "run.section"),
        ]:
            requires, reason, _ = classify_labtalk_script(script)
            assert requires is True, script
            assert reason == label, script

    def test_dynamic_command_substitution_requires_confirm(self):
        for script in ['string s$="doc -n"; %(s$);', "%A;", "if (1) {%(cmd$)}"]:
            requires, reason, _ = classify_labtalk_script(script)
            assert requires is True, script
            assert reason == "%() dynamic command", script

    def test_ordinary_substitution_does_not_require_confirm(self):
        for script in [
            "%A = Book1;",
            "win -a %H;",
            "range rr = [%(page.name)]1!col(B);",
            'type -a "%(s$)";',
        ]:
            requires, _, _ = classify_labtalk_script(script)
            assert requires is False, script

    @pytest.mark.parametrize("script", [
        'string cmd$="doc -n"; if (1) %(cmd$);',
        "if (1) %A;",
        "if (max(1, (2 + 3))) %(cmd$);",
        "if (1) if (2) %A;",
        "IF /* condition */ ((1)) /* body */ %(cmd$);",
        "if\n((%(condition$))) %(cmd$);",
        'if (name$ == ")") %(cmd$);',
        "if (1 /* ) */) %A;",
        "if (0) type 1; else %(cmd$);",
        "if (0) {type 1;} else %A;",
        "if (0) type 1; else if ((1)) %(cmd$);",
        "if (1) (%(cmd$);)",
        "if (1) ((%A;))",
        "if (0) type 1; else (%(cmd$);)",
    ])
    def test_braceless_dynamic_commands_require_confirm(self, script):
        requires, reason, alts = classify_labtalk_script(script)
        assert requires is True
        assert reason == "%() dynamic command"
        assert alts == ("get_origin_info",)

    @pytest.mark.parametrize("script", [
        "if (1) %A = Book1;",
        "if ((1)) %A=Book1; else %B = Book2;",
        "if (1) win -a %H;",
        "if (1) range rr = [%(page.name)]1!col(B);",
        'if (1) type -a "%(cmd$)";',
        "if (len(%(name$))) type 1;",
        "if (1) type 1; // else %(cmd$);",
        'type -a "if (1) %A;"; /* if (1) %(cmd$); */',
        "if (1) (%A = Book1;)",
        "if ((%(condition$))) type 1;",
        "if (\n(%(condition$))) type 1;",
        "if\n(%(condition$)) type 1;",
        "if\r\n((%(condition$))) type 1;",
        "if /* header\ncomment */ ((%(condition$))) type 1;",
        "if // header comment\n((%(condition$))) type 1;",
        "value = (%(number$));",
        "if (1) value = (%(number$));",
        "type -a (%(name$));",
    ])
    def test_braceless_ordinary_substitutions_remain_safe(self, script):
        assert classify_labtalk_script(script) == (False, "", ())

    @pytest.mark.parametrize("script", [
        "for(ii=1; ii<=1; ii++) %(cmd$);",
        "for(; ii < max(1, (2)); ii++) %A;",
        "FOR /* header */ (ii=1; ii<=1; ii++) /* body */ %(cmd$);",
        "for(ii=1; ii<=1; ii++) for(jj=1; jj<=1; jj++) %A;",
        "switch(1) { case 1: %(cmd$); }",
        "switch(1) { default: %A; }",
        "switch(1) { CASE 1 2 5 to 7: /* body */ %A; }",
        'switch(%B) { case "a:b": case "c": %(cmd$); }',
        "switch(1) { case 1: for(ii=1; ii<=1; ii++) %(cmd$); }",
        "loop(ii, 1, 1) {%(cmd$);}",
        "repeat 2 {%(cmd$);}",
        "loop(ii, 1, 2) %(cmd$);",
        "loop(ii, min(1, 2), (2 + 3)) %A;",
        "loop\n(ii, 1, 2) %(cmd$);",
        "win -o Graph1 {%(cmd$);}",
        "window -oa Graph1 {%A;}",
        "layer -o [Graph1]2 {%(cmd$);}",
        "queue {%(cmd$);}",
        "define act {%1;}; act %(cmd$);",
        "define act {%5;}; act a b c d %(cmd$);",
    ])
    def test_loop_and_case_dynamic_commands_require_confirm(self, script):
        assert classify_labtalk_script(script) == (
            True, "%() dynamic command", ("get_origin_info",)
        )

    @pytest.mark.parametrize("script", [
        "for(ii=1; ii<=1; ii++) %A = Book1;",
        "for(ii=1; ii<=1; ii++) win -a %H;",
        "for(ii=1; ii<=1; ii++) range rr = [%(page.name)]1!col(B);",
        "for(ii=1; ii<=max(1, len(%(name$))); ii++) type 1;",
        "for(ii=1; (%(condition$)); ii++) type 1;",
        "for\n((%(initial$)); ii<1; ii++) type 1;",
        "switch(1) { case 1: %A = Book1; default: %B = Book2; }",
        "switch(1) { case 1: win -a %H; }",
        "switch(%(name$)) { case 1: type 1; }",
        "switch\n(%(value$)) { case 1: type 1; }",
        "switch(1) { case\n%A: type 1; }",
        "loop\n(ii, (%(start$)), 3) type 1;",
        "loop(ii, 1, 2) %A = Book1;",
        "loop(ii, 1, 2) type %(name$);",
        "win -o %(name$) {type %H;}",
        "layer -o [%(book$)]1 {type %H;}",
        "define act {type %1;}; act %(value$);",
        "range rr = col(A)[1:%(last$)];",
        "range rr = [case]1!col(A)[1:%(last$)];",
        "range rr = [Book1]case!col(A)[1:%(last$)];",
        "range case = col(A)[1:%(last$)];",
        "case = col(A)[1:%(last$)];",
        "range rr = [\ncase]1!col(A)[1:%(last$)];",
        'type "for(ii=1; ii<=1; ii++) %(cmd$);"; /* case 1: %A; */',
    ])
    def test_loop_and_case_ordinary_substitutions_remain_safe(self, script):
        assert classify_labtalk_script(script) == (False, "", ())

    @pytest.mark.parametrize("script", [
        "doc -e P %(cmd$);",
        "doc -e P %A;",
        'string cmd$="doc -n"; doc -e P %(cmd$);',
        "doc -ef W %A;",
        "DOCUMENT /* command */ -EF /* object */ LB /* body */ %(cmd$);",
        "document -e LR ANOVAOneWay %(cmd$);",
        "doc -ef LR DescStats_1 %A;",
        "doc -e LRW %(cmd$);",
        "if (1) doc -e P %(cmd$);",
        "doc -e W doc -e LW %A;",
        "doc -e P {%(cmd$);}",
    ])
    def test_document_iteration_dynamic_commands_require_confirm(self, script):
        assert classify_labtalk_script(script) == (
            True, "%() dynamic command", ("get_origin_info",)
        )

    @pytest.mark.parametrize("script", [
        "doc -e P print;",
        "doc -e W %A = %H;",
        "doc -ef LB win -a %H;",
        "document -e LR ANOVAOneWay type %(name$);",
        "doc -e P range rr = [%(page.name)]1!col(B);",
        'doc -e P type "%(cmd$)";',
        'type "doc -e P %(cmd$);"; /* doc -ef W %A; */',
        "doc -e P print; // doc -ef W %(cmd$);",
    ])
    def test_document_iteration_ordinary_substitutions_remain_safe(self, script):
        assert classify_labtalk_script(script) == (False, "", ())

    @pytest.mark.parametrize("script", [
        'queue -p "%(cmd$);";',
        'queue -ps "doc -n;";',
        "queue -p %(cmd$);",
    ])
    def test_queued_script_text_requires_confirm(self, script):
        assert classify_labtalk_script(script) == (
            True, "queue -p/-ps", ("get_origin_info",)
        )

    @pytest.mark.parametrize("script", [
        "doc -e %(obj$) %(cmd$);",
        "document -ef %A %B;",
        "doc -e %1 %(cmd$);",
        "doc -ef %(%(name$)) %(cmd$);",
        'string obj$="P"; string cmd$="doc -n"; doc -e %(obj$) %(cmd$);',
        "doc -e %(obj$) type %(name$);",
    ])
    def test_runtime_object_selectors_require_confirm(self, script):
        # A selector can expand to LR plus a class, or to text containing the
        # body itself. Its script boundary cannot be established statically.
        assert classify_labtalk_script(script)[0] is True

    @pytest.mark.parametrize("script", [
        "run test.ogs;",
        "run %(file$);",
        "run %A;",
        'run "folder with spaces/test.ogs";',
        "RUN /* file */ test.ogs;",
        "if (1) run test.ogs;",
        "for(ii=1; ii<2; ii++) run %(file$);",
        "doc -e P run test.ogs;",
    ])
    def test_legacy_script_file_execution_requires_confirm(self, script):
        assert classify_labtalk_script(script) == (
            True, "run script", ("get_origin_info",)
        )

    @pytest.mark.parametrize("script", [
        "run = 1;",
        "run$ = %(name$);",
        "range rr = [run]1!col(A);",
        "range run = col(A);",
        "type run %(name$);",
        'type "run test.ogs"; /* run other.ogs; */',
        "doc -e P type %(name$);",
    ])
    def test_data_substitutions_and_run_identifiers_remain_safe(self, script):
        assert classify_labtalk_script(script) == (False, "", ())

    @pytest.mark.parametrize("command", [
        "range rr = [case]1!col(A)[1:%(last$)];",
        "range rr = [Book1]case!col(A)[1:%(last$)];",
        "range case = col(A)[1:%(last$)];",
    ])
    def test_case_named_ranges_execute_without_confirmation(self, fresh_manager, fake_origin, command):
        from originlab_mcp.tools.advanced import register_advanced_tools

        mcp = DummyMCP()
        register_advanced_tools(mcp, fresh_manager)
        result = mcp.tools["execute_labtalk"](command=command)
        assert result["ok"] is True
        assert fake_origin.lt_commands == [command]

    def test_first_risk_and_rule_priority_are_preserved(self):
        assert classify_labtalk_script('python -s "print(1)"; doc -n;')[1] == "python"
        assert classify_labtalk_script("run.section(a.ogs, main); doc -n;")[1] == "run.section"


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
    @pytest.mark.parametrize("command", [
        "if (1) %(cmd$);",
        "if (max(1, (2))) %A;",
        "if (0) type 1; else %(cmd$);",
        "for(ii=1; ii<=1; ii++) %(cmd$);",
        "switch(1) { case 1: %(cmd$); }",
        "switch(1) { default: %A; }",
        "doc -e P %(cmd$);",
        "doc -ef W %A;",
        "document -e LR ANOVAOneWay %(cmd$);",
        "loop(ii, 1, 2) %(cmd$);",
    ])
    def test_braceless_dynamic_command_gate(self, fresh_manager, fake_origin, command):
        from originlab_mcp.tools.advanced import register_advanced_tools

        mcp = DummyMCP()
        register_advanced_tools(mcp, fresh_manager)
        blocked = mcp.tools["execute_labtalk"](command=command)
        assert blocked["ok"] is False
        assert blocked["error"]["target"] == "confirm"
        assert fake_origin.lt_commands == []

        allowed = mcp.tools["execute_labtalk"](command=command, confirm=True)
        assert allowed["ok"] is True
        assert allowed["data"]["confirmed_risk"] == "%() dynamic command"
        assert fake_origin.lt_commands == [command]

    @pytest.mark.parametrize("command", ["if (1) (%(cmd$);)", "if (0) type 1; else (%A;)"])
    def test_parenthesized_script_group_is_blocked(self, fresh_manager, fake_origin, command):
        from originlab_mcp.tools.advanced import register_advanced_tools

        mcp = DummyMCP()
        register_advanced_tools(mcp, fresh_manager)
        result = mcp.tools["execute_labtalk"](command=command)
        assert result["ok"] is False
        assert result["error"]["target"] == "confirm"
        assert fake_origin.lt_commands == []

    @pytest.mark.parametrize("command", [
        'queue -p "%(cmd$);";',
        'queue -ps "doc -n;";',
        "define act {%1;}; act %(cmd$);",
        "doc -e %(obj$) %(cmd$);",
        "run test.ogs;",
        "run %(file$);",
    ])
    def test_script_text_is_blocked_before_execution(self, fresh_manager, fake_origin, command):
        from originlab_mcp.tools.advanced import register_advanced_tools

        mcp = DummyMCP()
        register_advanced_tools(mcp, fresh_manager)
        result = mcp.tools["execute_labtalk"](command=command)
        assert result["ok"] is False
        assert result["error"]["target"] == "confirm"
        assert fake_origin.lt_commands == []

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
