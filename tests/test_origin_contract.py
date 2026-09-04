"""Origin COM contract tests using an in-memory FakeOrigin backend.

Inspired by Excel MCP's MockOpenpyxlBackend pattern: exercise multi-tool
workflows and response contracts without a live Origin COM session.
"""

from __future__ import annotations

from pathlib import Path

from originlab_mcp.tools.customize import register_customize_tools
from originlab_mcp.tools.data import register_data_tools
from originlab_mcp.tools.export import register_export_tools
from originlab_mcp.tools.plot import register_plot_tools
from originlab_mcp.tools.system import register_system_tools
from originlab_mcp.types import OriginProProtocol
from tests.fakes import DummyMCP, FakeOrigin, attach_fake_origin

REQUIRED_SUCCESS_KEYS = {"ok", "message", "data", "resource", "warnings", "next_suggestions"}


def _register_core(mcp: DummyMCP, manager) -> None:
    register_data_tools(mcp, manager)
    register_plot_tools(mcp, manager)
    register_customize_tools(mcp, manager)
    register_export_tools(mcp, manager)
    register_system_tools(mcp, manager)


class TestFakeOriginProtocol:
    def test_fake_origin_satisfies_protocol(self):
        assert isinstance(FakeOrigin(), OriginProProtocol)


class TestImportPlotExportContract:
    def test_end_to_end_import_plot_style_export(
        self, fresh_manager, fake_origin, tmp_path: Path
    ):
        mcp = DummyMCP()
        _register_core(mcp, fresh_manager)

        imported = mcp.tools["import_data_from_text"](
            data="Time,Voltage\n1,10\n2,20\n3,30\n4,25"
        )
        assert imported["ok"] is True
        assert imported.keys() >= REQUIRED_SUCCESS_KEYS
        assert imported["data"]["rows"] == 4
        assert imported["data"]["cols"] == 2
        assert fresh_manager.active_worksheet is not None
        assert fake_origin.books

        designations = mcp.tools["set_column_designations"](designations="XY")
        assert designations["ok"] is True

        plotted = mcp.tools["create_plot"](x_col=0, y_cols=1, plot_type="scatter")
        assert plotted["ok"] is True
        assert plotted["data"]["curve_count"] == 1
        assert fresh_manager.active_graph is not None
        graph = fake_origin.find_graph(fresh_manager.active_graph)
        assert graph is not None
        assert graph.template == "scatter"
        assert graph[0].num_plots == 1
        assert graph[0].rescaled is True

        titled = mcp.tools["set_axis_title"](axis="x", title="Time (s)")
        assert titled["ok"] is True
        assert graph[0].axis("x").title == "Time (s)"

        colored = mcp.tools["set_plot_color"](color="#ff0000")
        assert colored["ok"] is True
        assert graph[0].plot(0).color == "#ff0000"

        output = tmp_path / "chart.png"
        exported = mcp.tools["export_graph"](output_path=str(output))
        assert exported["ok"] is True
        assert output.exists()
        assert output.read_bytes() == b"fake-origin-export"
        assert graph.saved and graph.saved[0][1]["type"] == "png"

        listed = mcp.tools["list_worksheets"]()
        assert listed["ok"] is True
        assert any(
            item["full_name"] == imported["data"]["sheet_name"]
            for item in listed["data"]["worksheets"]
        )

        info = mcp.tools["get_worksheet_info"]()
        assert info["ok"] is True
        assert info["data"]["cols"] == 2

        session = mcp.tools["read_origin_session"]()
        assert session["ok"] is True
        assert session["data"]["counts"]["worksheets"] >= 1
        assert session["data"]["counts"]["graphs"] >= 1

    def test_import_csv_then_list_and_get_cell(
        self, fresh_manager, fake_origin, tmp_path: Path
    ):
        mcp = DummyMCP()
        _register_core(mcp, fresh_manager)

        csv_path = tmp_path / "sample.csv"
        csv_path.write_text("X,Y\n1,2\n3,4\n", encoding="utf-8")

        imported = mcp.tools["import_csv"](file_path=str(csv_path))
        assert imported["ok"] is True
        assert imported["data"]["rows"] == 2
        assert imported["data"]["cols"] == 2

        cell = mcp.tools["get_cell_value"](row=0, col=1)
        assert cell["ok"] is True
        assert cell["data"]["value"] == 2


class TestProjectLifecycleContract:
    def test_new_project_clears_fake_session(self, fresh_manager, fake_origin):
        mcp = DummyMCP()
        _register_core(mcp, fresh_manager)

        mcp.tools["import_data_from_text"](data="A,B\n1,2\n3,4")
        mcp.tools["create_plot"](x_col=0, y_cols=1)
        assert fake_origin.graphs

        result = mcp.tools["new_project"]()
        assert result["ok"] is True
        assert len(fake_origin.graphs) == 0
        assert len(fake_origin.books) == 1
        assert fresh_manager.active_graph is None


class TestResponseContract:
    def test_error_responses_keep_standard_shape(self, fresh_manager, fake_origin):
        mcp = DummyMCP()
        _register_core(mcp, fresh_manager)

        result = mcp.tools["create_plot"](x_col=0, y_cols=1)
        assert result["ok"] is False
        assert "error" in result
        assert result["error"]["target"] in {"sheet_name", "graph_name"}

    def test_attach_helper_avoids_real_connect(self, fresh_manager):
        origin = FakeOrigin()
        attach_fake_origin(fresh_manager, origin)
        called = {"n": 0}

        def _probe(op):
            called["n"] += 1
            assert op is origin
            return op.path("e")

        path = fresh_manager.execute(_probe)
        assert called["n"] == 1
        assert path.endswith("Origin.exe")
        assert fresh_manager.is_connected
