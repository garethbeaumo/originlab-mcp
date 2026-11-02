"""Regression contracts from locked originpro 1.1.15 (no live Origin)."""
from types import SimpleNamespace

import pytest

from originlab_mcp.session import build_session_snapshot
from originlab_mcp.tools.data import register_data_tools
from tests.fakes import DummyMCP


@pytest.fixture
def data_tools(fresh_manager, fake_origin, monkeypatch):
    monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE", "0")
    mcp = DummyMCP()
    register_data_tools(mcp, fresh_manager)
    sheet = fake_origin.new_sheet(lname="Keys")
    for col, values in enumerate(([3, 1, 2], [20, 30, 10], [100, 300, 200])):
        sheet.from_list(col, values)
    fresh_manager.active_worksheet = "Keys"
    return mcp.tools, sheet


@pytest.mark.parametrize("col", [0, 1, 2])
@pytest.mark.parametrize("descending", [False, True])
def test_sort_keeps_rows_and_uses_requested_key(data_tools, col, descending):
    tools, sheet = data_tools
    original = list(zip(*(sheet.to_list(c) for c in range(3)), strict=True))
    result = tools["sort_worksheet"](col=col, descending=descending)
    assert result["ok"]
    assert list(zip(*(sheet.to_list(c) for c in range(3)), strict=True)) == sorted(
        original, key=lambda row: row[col], reverse=descending
    )


@pytest.mark.parametrize("tool,method,kwargs", [
    ("sort_worksheet", "sort", {"col": 0}),
    ("clear_worksheet", "clear", {}),
])
@pytest.mark.parametrize("outcome", [0, 7, None, "exception"])
def test_native_status_is_not_false_success(data_tools, tool, method, kwargs, outcome):
    tools, sheet = data_tools

    def operation(*args, **kw):
        if outcome == "exception":
            raise RuntimeError("native failure")
        return outcome

    setattr(sheet, method, operation)
    assert tools[tool](**kwargs)["ok"] is (outcome == 0)


@pytest.mark.parametrize("kwargs,cleared", [
    ({}, {0, 1, 2}), ({"start_col": 1}, {1, 2}),
    ({"end_col": 1}, {0, 1}), ({"start_col": 1, "end_col": 1}, {1}),
])
def test_clear_inclusive_range(data_tools, kwargs, cleared):
    tools, sheet = data_tools
    before = [sheet.to_list(c)[:] for c in range(3)]
    assert tools["clear_worksheet"](**kwargs)["ok"]
    for c in range(3):
        assert sheet.to_list(c) == ([] if c in cleared else before[c])


@pytest.mark.parametrize("kwargs", [
    {"start_col": -1}, {"end_col": -1}, {"start_col": 3},
    {"end_col": 3}, {"start_col": 2, "end_col": 1},
])
def test_invalid_clear_never_mutates(data_tools, kwargs):
    tools, sheet = data_tools
    before = [sheet.to_list(c)[:] for c in range(3)]
    assert not tools["clear_worksheet"](**kwargs)["ok"]
    assert [sheet.to_list(c) for c in range(3)] == before


@pytest.mark.parametrize("kind", ["w", "g", "m"])
@pytest.mark.parametrize("mode", ["raise", "none", "generator"])
def test_core_enumeration_failure_is_not_empty(kind, mode):
    def pages(k):
        if k != kind:
            return []
        if mode == "none":
            return None
        if mode == "raise":
            raise RuntimeError("COM unavailable")

        def failed():
            yield SimpleNamespace(name="partial")
            raise RuntimeError("COM iteration failed")
        return failed()

    with pytest.raises(RuntimeError, match=kind):
        build_session_snapshot(SimpleNamespace(pages=pages))


def test_empty_project_is_distinct_from_unsupported_optional_types():
    snapshot = build_session_snapshot(SimpleNamespace(pages=lambda k: []))
    assert snapshot["counts"]["worksheets"] == 0
    assert snapshot["counts"]["notes"] is None
    assert snapshot["enumeration"]["notes"]["status"] == "unsupported"
    assert snapshot["partial"] is True


def test_optional_failure_retains_core_data_and_unknown_count():
    def pages(kind):
        if kind == "Notes":
            raise RuntimeError("Notes failed")
        return []

    snapshot = build_session_snapshot(SimpleNamespace(
        pages=pages, session_page_types={"Notes", "Excel"}
    ))
    assert snapshot["counts"]["worksheets"] == 0
    assert snapshot["counts"]["notes"] is None
    assert snapshot["counts"]["excel_books"] == 0
    assert snapshot["enumeration"]["notes"]["status"] == "error"
    assert snapshot["warnings"]
