"""
Tool 基础测试

测试 validators 模块的返回结构和校验函数。
这些测试不依赖 Origin COM，可以在任何环境运行。
"""

from types import MethodType

import pytest

from originlab_mcp.origin_manager import OriginManager
from originlab_mcp.tools.customize import register_customize_tools
from originlab_mcp.tools.data import register_data_tools
from originlab_mcp.tools.export import register_export_tools
from originlab_mcp.utils.constants import (
    ColumnDesignation,
    ErrorType,
    ExportFormat,
    PlotType,
    ScaleType,
)
from originlab_mcp.utils.validators import (
    error_response,
    normalize_y_cols,
    success_response,
    validate_column_index,
    validate_column_indices,
    validate_designation,
    validate_export_format,
    validate_plot_type,
    validate_scale_type,
)


# ===================================================================
# success_response 测试
# ===================================================================


class TestSuccessResponse:
    def test_basic(self):
        r = success_response(message="done")
        assert r["ok"] is True
        assert r["message"] == "done"
        assert r["data"] == {}
        assert r["resource"] == {}
        assert r["warnings"] == []
        assert r["next_suggestions"] == []

    def test_with_data(self):
        r = success_response(
            message="ok",
            data={"key": "val"},
            resource={"active_worksheet": "Sheet1"},
            warnings=["warn1"],
            next_suggestions=["next_tool"],
        )
        assert r["data"]["key"] == "val"
        assert r["resource"]["active_worksheet"] == "Sheet1"
        assert len(r["warnings"]) == 1
        assert "next_tool" in r["next_suggestions"]


# ===================================================================
# error_response 测试
# ===================================================================


class TestErrorResponse:
    def test_basic(self):
        r = error_response(
            message="not found",
            error_type=ErrorType.NOT_FOUND,
            target="worksheet",
            value="Sheet9",
            hint="Call list_worksheets",
        )
        assert r["ok"] is False
        assert r["error"]["type"] == "not_found"
        assert r["error"]["target"] == "worksheet"
        assert r["error"]["value"] == "Sheet9"
        assert r["error"]["hint"] == "Call list_worksheets"

    def test_string_type(self):
        r = error_response(
            message="error",
            error_type="custom_error",
            target="x",
        )
        assert r["error"]["type"] == "custom_error"


# ===================================================================
# 验证函数测试
# ===================================================================


class TestValidateColumnIndex:
    def test_valid(self):
        assert validate_column_index(0, 3) is None
        assert validate_column_index(2, 3) is None

    def test_out_of_range(self):
        err = validate_column_index(5, 3)
        assert err is not None
        assert "5" in err
        assert "3" in err

    def test_negative(self):
        err = validate_column_index(-1, 3)
        assert err is not None


class TestValidateColumnIndices:
    def test_valid(self):
        assert validate_column_indices([0, 1, 2], 3) is None

    def test_one_bad(self):
        err = validate_column_indices([0, 1, 5], 3)
        assert err is not None


class TestValidatePlotType:
    def test_valid(self):
        for pt in PlotType:
            assert validate_plot_type(pt.value) is None

    def test_invalid(self):
        err = validate_plot_type("heatmap")
        assert err is not None
        assert "heatmap" in err


class TestValidateDesignation:
    def test_valid(self):
        for d in ColumnDesignation:
            assert validate_designation(d.value) is None
        assert validate_designation("YErr") is None

    def test_invalid(self):
        err = validate_designation("W")
        assert err is not None


class TestValidateScaleType:
    def test_valid(self):
        for s in ScaleType:
            assert validate_scale_type(s.value) is None

    def test_invalid(self):
        err = validate_scale_type("exp")
        assert err is not None


class TestValidateExportFormat:
    def test_valid(self):
        for f in ExportFormat:
            assert validate_export_format(f.value) is None

    def test_invalid(self):
        err = validate_export_format("bmp")
        assert err is not None


# ===================================================================
# normalize_y_cols 测试
# ===================================================================


class TestNormalizeYCols:
    def test_single_int(self):
        assert normalize_y_cols(1) == [1]

    def test_list(self):
        assert normalize_y_cols([1, 2, 3]) == [1, 2, 3]

    def test_empty_list(self):
        assert normalize_y_cols([]) == []

    def test_rejects_non_list_sequence(self):
        with pytest.raises(ValueError):
            normalize_y_cols((1, 2))

    def test_rejects_string(self):
        with pytest.raises(ValueError):
            normalize_y_cols("12")  # type: ignore[arg-type]

    def test_rejects_non_int_element(self):
        with pytest.raises(ValueError):
            normalize_y_cols([1, "2"])  # type: ignore[list-item]

    def test_rejects_bool(self):
        with pytest.raises(ValueError):
            normalize_y_cols(True)  # type: ignore[arg-type]


class DummyMCP:
    def __init__(self):
        self.tools = {}

    def tool(self):
        def decorator(fn):
            self.tools[fn.__name__] = fn
            return fn

        return decorator


@pytest.fixture
def fresh_manager():
    OriginManager.reset_for_testing()
    manager = OriginManager()
    yield manager
    OriginManager.reset_for_testing()


class TestToolRegressions:
    def test_get_worksheet_data_uses_sheet_lists_and_deduplicates_columns(
        self,
        fresh_manager,
    ):
        mcp = DummyMCP()
        register_data_tools(mcp)
        manager = OriginManager()
        manager.active_worksheet = "[Book1]Sheet1"

        class StubCol:
            def __init__(self, name: str, long_name: str = ""):
                self.name = name
                self._long_name = long_name

            def get_label(self, label_type: str) -> str:
                return self._long_name if label_type == "L" else ""

        class StubSheet:
            rows = 99
            cols = 2

            def __init__(self):
                self._cols = [
                    StubCol("A", "Value"),
                    StubCol("B", "Value"),
                ]
                self._data = [
                    [1, 2],
                    [10],
                ]

            def get_col(self, index: int) -> StubCol:
                return self._cols[index]

            def to_list(self, index: int) -> list[int]:
                return self._data[index]

        class StubOp:
            def __init__(self):
                self.sheet = StubSheet()

            def find_sheet(self, kind: str, name: str) -> StubSheet:
                assert kind == "w"
                assert name == "[Book1]Sheet1"
                return self.sheet

        op = StubOp()
        manager.execute = MethodType(
            lambda self, func, *args, **kwargs: func(op, *args, **kwargs),
            manager,
        )

        result = mcp.tools["get_worksheet_data"](max_rows=5)

        assert result["ok"] is True
        assert result["data"]["total_rows"] == 2
        assert result["data"]["columns"] == ["Value", "Value_2"]
        assert result["data"]["data"] == [
            {"Value": 1, "Value_2": 10},
            {"Value": 2, "Value_2": None},
        ]

    def test_set_column_designations_normalizes_legacy_yerr(self, fresh_manager):
        mcp = DummyMCP()
        register_data_tools(mcp)
        manager = OriginManager()
        manager.active_worksheet = "[Book1]Sheet1"

        class StubCol:
            def __init__(self, name: str):
                self.name = name

        class StubSheet:
            cols = 3

            def __init__(self):
                self.applied = ""

            def cols_axis(self, spec: str) -> None:
                self.applied = spec

            def get_labels(self, label_type: str) -> list[str]:
                assert label_type == "D"
                return list(self.applied)

            def get_col(self, index: int) -> StubCol:
                return StubCol(f"C{index + 1}")

        class StubOp:
            def __init__(self):
                self.sheet = StubSheet()

            def find_sheet(self, kind: str, name: str) -> StubSheet:
                assert kind == "w"
                assert name == "[Book1]Sheet1"
                return self.sheet

        op = StubOp()
        manager.execute = MethodType(
            lambda self, func, *args, **kwargs: func(op, *args, **kwargs),
            manager,
        )

        result = mcp.tools["set_column_designations"]("XYYErr")

        assert result["ok"] is True
        assert op.sheet.applied == "XYE"
        assert result["data"]["designations"] == "XYE"
        assert result["data"]["requested_designations"] == "XYYErr"

    def test_create_plot_invalid_y_cols_returns_invalid_input(self, fresh_manager):
        mcp = DummyMCP()
        register_data_tools(mcp)
        manager = OriginManager()
        manager.active_worksheet = "[Book1]Sheet1"

        from originlab_mcp.tools.plot import register_plot_tools

        register_plot_tools(mcp)

        result = mcp.tools["create_plot"](x_col=0, y_cols="1")

        assert result["ok"] is False
        assert result["error"]["type"] == "invalid_input"
        assert result["error"]["target"] == "y_cols"

    def test_set_axis_scale_uses_origin_scale_strings(self, fresh_manager):
        mcp = DummyMCP()
        register_customize_tools(mcp)
        manager = OriginManager()
        manager.active_graph = "Graph1"

        class StubLayer:
            def __init__(self):
                self.assigned = []

            @property
            def xscale(self):
                return None

            @xscale.setter
            def xscale(self, value):
                self.assigned.append(("x", value))

            @property
            def yscale(self):
                return None

            @yscale.setter
            def yscale(self, value):
                self.assigned.append(("y", value))

        class StubGraph:
            def __init__(self, layer: StubLayer):
                self.layer = layer

            def __getitem__(self, index: int) -> StubLayer:
                return self.layer

        class StubOp:
            def __init__(self, layer: StubLayer):
                self.layer = layer

            def find_graph(self, name: str) -> StubGraph:
                assert name == "Graph1"
                return StubGraph(self.layer)

        layer = StubLayer()
        manager.execute = MethodType(
            lambda self, func, *args, **kwargs: func(
                StubOp(layer),
                *args,
                **kwargs,
            ),
            manager,
        )

        result = mcp.tools["set_axis_scale"]("x", "log")

        assert result["ok"] is True
        assert layer.assigned == [("x", "log10")]

    def test_export_graph_passes_explicit_type(self, fresh_manager):
        mcp = DummyMCP()
        register_export_tools(mcp)
        manager = OriginManager()
        manager.active_graph = "Graph1"

        class StubGraph:
            def __init__(self):
                self.calls = []

            def save_fig(self, path: str, **kwargs) -> str:
                self.calls.append((path, kwargs))
                return path

        class StubOp:
            def __init__(self):
                self.graph = StubGraph()

            def find_graph(self, name: str) -> StubGraph:
                assert name == "Graph1"
                return self.graph

        op = StubOp()
        manager.execute = MethodType(
            lambda self, func, *args, **kwargs: func(op, *args, **kwargs),
            manager,
        )

        result = mcp.tools["export_graph"]("out/chart", output_format="svg")

        assert result["ok"] is True
        assert op.graph.calls == [("out/chart", {"type": "svg", "width": 800})]
