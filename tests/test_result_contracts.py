"""Regressions against the public originpro 1.1.15 return contracts."""

from __future__ import annotations

import json
import threading

import pytest

from originlab_mcp.session import collect_worksheet_preview, read_worksheet_detail
from originlab_mcp.tools.analysis import register_analysis_tools
from originlab_mcp.tools.data import register_data_tools
from originlab_mcp.tools.export import register_export_tools
from originlab_mcp.tools.plot import register_plot_tools
from originlab_mcp.tools.system import register_system_tools
from tests.fakes.origin import FakeGraph, FakeLinearFit, FakeNLFit


@pytest.fixture
def workflow(dummy_mcp, fresh_manager, fake_origin):
    for register in (register_analysis_tools, register_data_tools, register_export_tools,
                     register_plot_tools, register_system_tools):
        register(dummy_mcp, fresh_manager)
    sheet = fake_origin.new_sheet()
    sheet.from_list(0, [1, 2, 3], lname="X")
    sheet.from_list(1, [3, 5, 7], lname="Y")
    sheet.from_list(2, [4, 6, 8], lname="Y2")
    fresh_manager.active_worksheet = sheet.name
    return dummy_mcp.tools, fresh_manager, fake_origin, sheet


@pytest.mark.parametrize("fixed,slope,intercept", [
    ({}, 2, 1), ({"fix_intercept": 0}, 17 / 7, 0), ({"fix_slope": 0}, 0, 5),
])
def test_linear_numeric_acceptance(workflow, fixed, slope, intercept):
    tools, _, origin, _ = workflow
    result = tools["linear_fit"](0, 1, **fixed)
    assert result["ok"], result
    assert result["data"]["parameters"]["Slope"]["value"] == pytest.approx(slope)
    assert result["data"]["parameters"]["Intercept"]["value"] == pytest.approx(intercept)
    for method in fixed:
        assert method in origin.linear_fits[-1].calls


def test_fake_rejects_method_overwrite():
    with pytest.raises(AttributeError):
        FakeLinearFit().fix_slope = 0


@pytest.mark.parametrize("report", [False, True])
def test_nonlinear_flat_result_and_report_order(workflow, report):
    tools, _, origin, _ = workflow
    result = tools["nonlinear_fit"]("Gauss", 0, 1, initial_params={"xc": 2.5}, generate_report=report)
    assert result["ok"], result
    assert result["data"]["result"]["xc"] == 2.5
    assert result["data"]["result"]["cod"] == 0.98
    assert result["data"]["parameters"]["xc"] == {"value": 2.5, "error": 0.02}
    if report:
        assert origin.nl_fits[-1].calls[-2:] == ["report", "result"]
        assert result["data"]["report"]["report_range"]


@pytest.mark.parametrize("tool,cls,stage,kwargs", [
    ("linear_fit", FakeLinearFit, "result", {}),
    ("linear_fit", FakeLinearFit, "report", {"confidence_band": True}),
    ("nonlinear_fit", FakeNLFit, "result", {"function_name": "Gauss"}),
    ("nonlinear_fit", FakeNLFit, "report", {"function_name": "Gauss", "generate_report": True}),
])
@pytest.mark.parametrize("failure", ["exception", "empty"])
def test_fit_failure_is_not_success(workflow, monkeypatch, tool, cls, stage, kwargs, failure):
    def fail(*args):
        if failure == "exception":
            raise RuntimeError("controlled fit failure")
        return ("", "") if stage == "report" else {}
    monkeypatch.setattr(cls, stage, fail)
    result = workflow[0][tool](x_col=0, y_col=1, **kwargs)
    assert result["ok"] is False


@pytest.mark.parametrize("returned", ["", None, False])
def test_export_unconfirmed_return(workflow, monkeypatch, tmp_path, returned):
    tools, manager, origin, _ = workflow
    graph = origin.new_graph()
    manager.active_graph = graph.name
    monkeypatch.setattr(graph, "save_fig", lambda *a, **kw: returned)
    result = tools["export_graph"](str(tmp_path / "wanted.png"))
    assert result["ok"] is False


def test_export_uses_actual_path(workflow, monkeypatch, tmp_path):
    tools, manager, origin, _ = workflow
    graph = origin.new_graph()
    manager.active_graph = graph.name
    actual = str(tmp_path / "actual.png")
    monkeypatch.setattr(graph, "save_fig", lambda *a, **kw: actual)
    result = tools["export_graph"](str(tmp_path / "requested.png"))
    assert result["ok"]
    assert result["data"]["output_path"] == actual


@pytest.mark.parametrize("failure", ["empty", "exception"])
def test_batch_export_retains_partial_success(workflow, monkeypatch, tmp_path, failure):
    tools, _, origin, _ = workflow
    first, second, third = [origin.new_graph() for _ in range(3)]
    def fail(*a, **kw):
        if failure == "exception":
            raise RuntimeError("export disconnected")
        return ""
    monkeypatch.setattr(second, "save_fig", fail)
    result = tools["export_all_graphs"](str(tmp_path))
    assert result["ok"] is False
    assert result["data"]["count"] == 2
    assert [item["graph_name"] for item in result["data"]["files"]] == [first.name, third.name]
    assert result["data"]["failed"][0]["graph_name"] == second.name
    assert result["data"]["failed"][0]["output_state"] == "unconfirmed"


def test_batch_enumeration_failure_retains_confirmed_files(workflow, monkeypatch, tmp_path):
    tools, _, origin, _ = workflow
    graph = origin.new_graph()
    def pages(kind):
        yield graph
        raise RuntimeError("enumeration disconnected")
    monkeypatch.setattr(origin, "pages", pages)
    result = tools["export_all_graphs"](str(tmp_path))
    assert result["ok"] is False
    assert result["data"]["count"] == 1
    assert result["data"]["enumeration_complete"] is False
    assert result["data"]["failed"][0]["stage"] == "enumerate_graphs"


class LargeSheet:
    rows = 1_000_000
    cols = 10_000
    shape = (2_000_000, 10_000)

    def __init__(self):
        self.reads = []
        self.labels = []
        self.value = 1

    def get_label(self, index, kind):
        self.labels.append(index)
        return f"C{index}"

    def to_list(self, *args):
        raise AssertionError("Unbounded column read")

    def to_list2(self, r1, r2, c1, c2):
        assert min(r1, r2, c1, c2) >= 0
        self.reads.append((r1, r2, c1, c2))
        return [[self.value for _ in range(c1, c2 + 1)] for _ in range(r1, r2 + 1)]


def test_preview_reads_only_three_by_four():
    sheet = LargeSheet()
    result = collect_worksheet_preview(sheet, 3, 4)
    assert sheet.reads == [(0, 2, 0, 3)]
    assert sheet.labels == [0, 1, 2, 3]
    assert result["returned_rows"] == 3 and result["returned_cols"] == 4
    assert result["total_rows"] == 1_000_000
    assert result["allocated_rows"] == 2_000_000
    assert result["omitted_cols"] == 9996


@pytest.mark.parametrize("rows,cols", [(0, 4), (3, 0)])
def test_zero_preview_never_reads_cells(rows, cols):
    sheet = LargeSheet()
    result = collect_worksheet_preview(sheet, rows, cols)
    assert sheet.reads == [] and result["data"] == []


def test_preview_caps_rectangle_and_text():
    sheet = LargeSheet()
    sheet.value = "x" * 100_000
    result = collect_worksheet_preview(sheet, 100_000, 100_000)
    assert sheet.reads == [(0, 99, 0, 31)]
    assert result["shortened_text_values"] > 0 and result["omissions"]
    assert sum(len(v) for row in result["data"] for v in row.values()) <= 20_000
    assert max(len(v) for row in result["data"] for v in row.values()) <= 256


def test_preview_long_duplicate_names_are_bounded_and_unique(monkeypatch):
    sheet = LargeSheet()
    monkeypatch.setattr(sheet, "get_label", lambda *a: "x" * 100_000)
    result = collect_worksheet_preview(sheet, 3, 4)
    assert len(set(result["columns"])) == 4
    assert max(map(len, result["columns"])) <= 130
    assert result["shortened_text_values"] == 4


def test_resource_metadata_has_column_and_text_limits(workflow, monkeypatch):
    origin = workflow[2]
    sheet = LargeSheet()
    monkeypatch.setattr(sheet, "get_label", lambda *a: "x" * 100_000)
    monkeypatch.setattr(origin, "find_sheet", lambda *a: sheet)
    detail = read_worksheet_detail(origin, "Large", max_preview_rows=0)
    assert len(detail["columns"]) == 32
    assert detail["omitted_columns"] == 9968
    assert sheet.reads == []
    assert "designation" in detail["columns"][0]["truncated_fields"]
    assert len(detail["columns"][0]["designation"]) == 256


@pytest.mark.parametrize("max_rows,max_cols", [(-1, 4), (3, -1)])
def test_preview_negative_limits_rejected(workflow, max_rows, max_cols):
    result = workflow[0]["get_worksheet_data"](max_rows=max_rows, max_cols=max_cols)
    assert result["ok"] is False
    assert result["error"]["type"] == "invalid_input"


def test_preview_short_rectangle_is_error(workflow, monkeypatch):
    monkeypatch.setattr(workflow[3], "to_list2", lambda *a: [])
    assert workflow[0]["get_worksheet_data"]()["ok"] is False


def test_preview_tool_uses_same_bounded_reader(workflow, monkeypatch):
    tools, _, origin, _ = workflow
    sheet = LargeSheet()
    monkeypatch.setattr(origin, "find_sheet", lambda *a: sheet)
    assert tools["get_worksheet_data"](max_rows=3, max_cols=4)["ok"]
    assert sheet.reads == [(0, 2, 0, 3)]


@pytest.mark.parametrize("tool", ["get_worksheet_data", "read_origin_session"])
def test_preview_com_failure_is_explicit(workflow, monkeypatch, tool):
    tools, manager, _, sheet = workflow
    monkeypatch.setattr(manager, "peek_active_context", lambda: {"active_worksheet": sheet.name})
    def fail(*a):
        raise RuntimeError("preview COM failure")
    monkeypatch.setattr(sheet, "to_list2", fail)
    result = tools[tool](**({"include_preview": True} if tool == "read_origin_session" else {}))
    assert result["ok"] is False
    assert "preview COM failure" in result["message"]


def test_plot_second_curve_failure_preserves_identity(workflow, monkeypatch):
    tools, manager, origin, _ = workflow
    graph = FakeGraph("RetainedGraph")
    original = graph[0].add_plot
    def add(wks, **kwargs):
        if kwargs["coly"] == 2:
            raise RuntimeError("second curve failed")
        return original(wks, **kwargs)
    monkeypatch.setattr(graph[0], "add_plot", add)
    monkeypatch.setattr(origin, "new_graph", lambda **kw: graph)
    result = tools["create_plot"](0, [1, 2])
    assert result["ok"] is False
    data = result["data"]
    assert data["graph_name"] == manager.active_graph == "RetainedGraph"
    assert data["creation_state"] == "created" and data["failed_step"] == "add_plot:1"
    assert data["curve_count"] == graph[0].num_plots == 1
    assert "add_plot:0" in data["completed_steps"] and "rescale" not in data["completed_steps"]


def test_plot_preflight_failure_creates_nothing(workflow):
    tools, _, origin, _ = workflow
    result = tools["create_plot"](0, 999)
    assert result["ok"] is False and not origin.graphs
    assert result["data"]["creation_state"] == "not_attempted"
    assert result["data"]["failed_step"] == "preflight"


def test_plot_disconnected_during_creation_is_unknown(workflow, monkeypatch):
    tools, _, origin, _ = workflow
    def disconnected(**kw):
        raise RuntimeError("RPC disconnected; creation outcome unknown")
    monkeypatch.setattr(origin, "new_graph", disconnected)
    result = tools["create_plot"](0, 1)
    assert result["ok"] is False
    assert result["data"]["creation_state"] == "unknown"
    assert result["data"]["graph_name"] is None
    assert result["data"]["current_state_verified"] is False


@pytest.mark.parametrize("stage", ["rescale", "group"])
def test_plot_later_failure_preserves_completed_curves(workflow, monkeypatch, stage):
    tools, _, origin, _ = workflow
    graph = origin.new_graph()
    monkeypatch.setattr(origin, "new_graph", lambda **kw: graph)
    def fail():
        raise RuntimeError("disconnected after curve creation")
    monkeypatch.setattr(graph[0], stage, fail)
    result = tools["create_plot"](0, [1, 2])
    assert result["ok"] is False
    assert result["data"]["curve_count"] == graph[0].num_plots == 2
    assert result["data"]["graph_name"] == graph.name
    assert result["data"]["failed_step"] == stage


def test_plot_timeout_snapshot_does_not_change_later(workflow, monkeypatch):
    tools, _, origin, _ = workflow
    started, release, finished = threading.Event(), threading.Event(), threading.Event()
    graph = origin.new_graph()
    def blocked(wks, **kw):
        started.set()
        assert release.wait(5)
        finished.set()
        return object()
    monkeypatch.setattr(origin, "new_graph", lambda **kw: graph)
    monkeypatch.setattr(graph[0], "add_plot", blocked)
    monkeypatch.setenv("ORIGINLAB_MCP_DISPATCH_TIMEOUT", "0.1")
    try:
        result = tools["create_plot"](0, 1)
        assert started.is_set() and result["ok"] is False
        assert result["error"]["type"] == "timeout"
        snapshot = json.dumps(result["data"])
    finally:
        release.set()
    assert finished.wait(2)
    assert json.dumps(result["data"]) == snapshot
