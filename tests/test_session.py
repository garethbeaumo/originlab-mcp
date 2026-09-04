"""Tests for Origin session reading helpers, system tool, and MCP resources."""

from __future__ import annotations

import json
from types import MethodType

from originlab_mcp.origin_manager import OriginManager
from originlab_mcp.resources import register_session_resources
from originlab_mcp.session import (
    build_session_snapshot,
    collect_graphs,
    collect_matrix_sheets,
    collect_named_pages,
    collect_project_meta,
    collect_worksheets,
    decode_resource_segment,
    read_graph_detail,
    read_worksheet_detail,
)
from originlab_mcp.tools.system import register_system_tools


class DummyMCP:
    def __init__(self) -> None:
        self.tools: dict = {}
        self.resources: dict = {}

    def tool(self, **_kwargs):
        def decorator(fn):
            self.tools[fn.__name__] = fn
            return fn

        return decorator

    def resource(self, uri, **_kwargs):
        def decorator(fn):
            self.resources[uri] = fn
            return fn

        return decorator


class StubSheet:
    def __init__(self, name: str, rows: int, cols: int, columns: list[list], labels: dict):
        self.name = name
        self.rows = rows
        self.cols = cols
        self._columns = columns
        self._labels = labels

    def get_label(self, index: int, label_type: str) -> str:
        return self._labels[label_type][index]

    def get_labels(self, label_type: str) -> list[str]:
        return list(self._labels[label_type])

    def to_list(self, index: int) -> list:
        return self._columns[index]


class StubBook:
    def __init__(self, name: str, sheets: list[StubSheet]):
        self.name = name
        self._sheets = sheets

    def __iter__(self):
        return iter(self._sheets)


class StubPlot:
    def __init__(self, color: tuple[int, int, int], source: str):
        self.color = color
        self._source = source

    def lt_range(self) -> str:
        return self._source


class StubLayer:
    def __init__(self, plots: list[StubPlot]):
        self._plots = plots

    def plot_list(self) -> list[StubPlot]:
        return self._plots


class StubGraph:
    def __init__(self, name: str, layers: list[StubLayer], long_name: str = ""):
        self.name = name
        self.lname = long_name
        self._layers = layers

    def __len__(self) -> int:
        return len(self._layers)

    def __getitem__(self, index: int) -> StubLayer:
        return self._layers[index]


class StubNote:
    def __init__(self, name: str):
        self.name = name
        self.lname = f"{name} note"


class StubOp:
    def __init__(self) -> None:
        self.sheet = StubSheet(
            "Sheet1",
            rows=2,
            cols=2,
            columns=[[1, 2], [10]],
            labels={
                "G": ["A", "B"],
                "L": ["Value", "Value"],
                "U": ["", ""],
                "C": ["", ""],
                "D": ["X", "Y"],
            },
        )
        self.graph = StubGraph(
            "Graph1",
            [StubLayer([StubPlot((255, 0, 0), "[Book1]Sheet1!B")])],
            long_name="Demo",
        )
        self._pages = {
            "Book": [StubBook("Book1", [self.sheet])],
            "Graph": [self.graph],
            "Matrix": [StubBook("MBook1", [])],
            "Notes": [StubNote("Notes1")],
            "Excel": [],
        }
        self._lt = {
            "pe_path$": r"C:\data\demo.opju",
            "pe_name$": "demo.opju",
        }

    def pages(self, kind: str):
        if kind == "Excel":
            raise RuntimeError("excel pages unavailable")
        return self._pages[kind]

    def path(self, kind: str) -> str:
        return r"C:\Origin" if kind == "e" else r"C:\Users\Origin"

    def get_lt_str(self, name: str) -> str:
        return self._lt.get(name, "")

    def find_sheet(self, kind: str, name: str) -> StubSheet | None:
        assert kind == "w"
        if name in {"[Book1]Sheet1", "Sheet1"}:
            return self.sheet
        return None

    def find_graph(self, name: str) -> StubGraph | None:
        if name == "Graph1":
            return self.graph
        return None


def _patch_execute(manager: OriginManager, op: StubOp) -> None:
    manager.execute = MethodType(
        lambda self, func, *args, **kwargs: func(op, *args, **kwargs),
        manager,
    )
    manager.peek_active_context = MethodType(
        lambda self: {
            "active_worksheet": self.active_worksheet,
            "active_graph": self.active_graph,
        },
        manager,
    )


class TestSessionSnapshot:
    def test_collect_worksheets_and_graphs(self):
        op = StubOp()
        worksheets = collect_worksheets(op)
        graphs = collect_graphs(op)

        assert worksheets == [
            {
                "book_name": "Book1",
                "sheet_name": "Sheet1",
                "full_name": "[Book1]Sheet1",
                "rows": 2,
                "cols": 2,
            }
        ]
        assert graphs[0]["graph_name"] == "Graph1"
        assert graphs[0]["layers"] == 1
        assert graphs[0]["long_name"] == "Demo"

    def test_missing_page_types_are_skipped(self):
        op = StubOp()
        assert collect_named_pages(op, "Excel") == []
        matrices = collect_matrix_sheets(op)
        assert matrices[0]["book_name"] == "MBook1"
        assert matrices[0]["sheet_name"] is None

    def test_project_meta_and_snapshot_counts(self):
        op = StubOp()
        snapshot = build_session_snapshot(
            op,
            active_worksheet="[Book1]Sheet1",
            active_graph="Graph1",
        )

        assert collect_project_meta(op)["path"] == r"C:\data\demo.opju"
        assert snapshot["counts"] == {
            "worksheets": 1,
            "graphs": 1,
            "matrices": 1,
            "notes": 1,
            "excel_books": 0,
        }
        assert snapshot["active_worksheet"] == "[Book1]Sheet1"
        assert "active_worksheet_preview" not in snapshot

    def test_include_preview_reads_active_worksheet(self):
        op = StubOp()
        snapshot = build_session_snapshot(
            op,
            active_worksheet="[Book1]Sheet1",
            include_preview=True,
            max_preview_rows=1,
        )
        preview = snapshot["active_worksheet_preview"]["preview"]
        assert preview["truncated"] is True
        assert preview["returned_rows"] == 1
        assert preview["columns"] == ["Value", "Value_2"]
        assert preview["data"] == [{"Value": 1, "Value_2": 10}]

    def test_worksheet_detail_includes_designations(self):
        detail = read_worksheet_detail(StubOp(), "[Book1]Sheet1", include_preview=False)
        assert detail["columns"][0]["designation"] == "X"
        assert detail["columns"][1]["name"] == "B"
        assert "preview" not in detail

    def test_graph_detail_includes_plot_color(self):
        detail = read_graph_detail(StubOp(), "Graph1")
        assert detail["layer_count"] == 1
        plot = detail["layers"][0]["plots"][0]
        assert plot["color"] == "#ff0000"
        assert plot["range"] == "[Book1]Sheet1!B"

    def test_decode_resource_segment(self):
        assert decode_resource_segment("Book%201") == "Book 1"


class TestReadOriginSessionTool:
    def test_read_origin_session_returns_snapshot(self):
        mcp = DummyMCP()
        manager = OriginManager(auto_recover_active=False)
        register_system_tools(mcp, manager)
        manager.active_worksheet = "[Book1]Sheet1"
        manager.active_graph = "Graph1"
        _patch_execute(manager, StubOp())

        result = mcp.tools["read_origin_session"]()

        assert result["ok"] is True
        assert result["data"]["counts"]["worksheets"] == 1
        assert result["resource"]["active_graph"] == "Graph1"
        assert "已阅读当前 Origin 会话" in result["message"]

    def test_read_origin_session_rejects_negative_preview_rows(self):
        mcp = DummyMCP()
        manager = OriginManager(auto_recover_active=False)
        register_system_tools(mcp, manager)

        result = mcp.tools["read_origin_session"](max_preview_rows=-1)

        assert result["ok"] is False
        assert result["error"]["target"] == "max_preview_rows"


class TestSessionResources:
    def test_session_resource_returns_json(self):
        mcp = DummyMCP()
        manager = OriginManager(auto_recover_active=False)
        register_session_resources(mcp, manager)
        manager.active_worksheet = "[Book1]Sheet1"
        _patch_execute(manager, StubOp())

        payload = json.loads(mcp.resources["originlab://session"]())

        assert payload["ok"] is True
        assert payload["data"]["worksheets"][0]["full_name"] == "[Book1]Sheet1"
        assert "originlab://worksheet/{book}/{sheet}" in mcp.resources
        assert "originlab://graph/{name}" in mcp.resources

    def test_worksheet_and_graph_templates(self):
        mcp = DummyMCP()
        manager = OriginManager(auto_recover_active=False)
        register_session_resources(mcp, manager)
        _patch_execute(manager, StubOp())

        sheet = json.loads(
            mcp.resources["originlab://worksheet/{book}/{sheet}"]("Book1", "Sheet1")
        )
        graph = json.loads(mcp.resources["originlab://graph/{name}"]("Graph1"))

        assert sheet["ok"] is True
        assert sheet["data"]["preview"]["returned_rows"] == 2
        assert graph["data"]["graph_name"] == "Graph1"

    def test_resource_returns_error_json_when_origin_unavailable(self):
        mcp = DummyMCP()
        manager = OriginManager(auto_recover_active=False)
        register_session_resources(mcp, manager)

        def _boom(self, func, *args, **kwargs):
            raise RuntimeError("无法导入 originpro")

        manager.execute = MethodType(_boom, manager)

        payload = json.loads(mcp.resources["originlab://session"]())

        assert payload["ok"] is False
        assert "originpro" in payload["message"]
