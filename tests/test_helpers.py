"""测试 utils/helpers.py 共享辅助函数和装饰器。"""

import pytest

from originlab_mcp.exceptions import (
    ColumnIndexError,
    GraphNotFoundError,
    InvalidAxisError,
    NoActiveGraphError,
    NoActiveWorksheetError,
    PlotIndexError,
    ToolError,
    WorksheetNotFoundError,
)
from originlab_mcp.origin_manager import OriginManager
from originlab_mcp.utils.helpers import (
    find_graph,
    find_worksheet,
    get_plot,
    resolve_graph_name,
    resolve_worksheet_name,
    sanitize_labtalk_name,
    tool_error_handler,
    validate_axis,
    validate_column_indices,
)


@pytest.fixture
def manager():
    """创建一个全新的 OriginManager 实例（不依赖单例）。"""
    return OriginManager()


# ===================================================================
# resolve_worksheet_name 测试
# ===================================================================


class TestResolveWorksheetName:
    def test_explicit_name(self, manager):
        result = resolve_worksheet_name("Sheet1", manager)
        assert result == "Sheet1"

    def test_active_worksheet(self, manager):
        manager.active_worksheet = "[Book1]Sheet1"
        result = resolve_worksheet_name(None, manager)
        assert result == "[Book1]Sheet1"

    def test_no_active_raises(self, manager):
        with pytest.raises(NoActiveWorksheetError):
            resolve_worksheet_name(None, manager)

    def test_explicit_overrides_active(self, manager):
        manager.active_worksheet = "[Book1]Sheet1"
        result = resolve_worksheet_name("Sheet2", manager)
        assert result == "Sheet2"


# ===================================================================
# resolve_graph_name 测试
# ===================================================================


class TestResolveGraphName:
    def test_explicit_name(self, manager):
        result = resolve_graph_name("Graph1", manager)
        assert result == "Graph1"

    def test_active_graph(self, manager):
        manager.active_graph = "Graph2"
        result = resolve_graph_name(None, manager)
        assert result == "Graph2"

    def test_no_active_raises(self, manager):
        with pytest.raises(NoActiveGraphError):
            resolve_graph_name(None, manager)


# ===================================================================
# find_worksheet 测试
# ===================================================================


class TestFindWorksheet:
    def test_found(self):
        class StubOp:
            def find_sheet(self, kind, name):
                return "found_sheet"

        result = find_worksheet(StubOp(), "Sheet1")
        assert result == "found_sheet"

    def test_not_found_raises(self):
        class StubOp:
            def find_sheet(self, kind, name):
                return None

        with pytest.raises(WorksheetNotFoundError):
            find_worksheet(StubOp(), "NoSheet")


# ===================================================================
# find_graph 测试
# ===================================================================


class TestFindGraph:
    def test_found(self):
        class StubOp:
            def find_graph(self, name):
                return "found_graph"

        result = find_graph(StubOp(), "Graph1")
        assert result == "found_graph"

    def test_not_found_raises(self):
        class StubOp:
            def find_graph(self, name):
                return None

        with pytest.raises(GraphNotFoundError):
            find_graph(StubOp(), "NoGraph")


# ===================================================================
# get_plot 测试
# ===================================================================


class TestGetPlot:
    def test_found(self):
        class StubLayer:
            def plot(self, idx):
                return f"plot_{idx}"

        result = get_plot(StubLayer(), 0)
        assert result == "plot_0"

    def test_not_found_raises(self):
        class StubLayer:
            def plot(self, idx):
                return None

        with pytest.raises(PlotIndexError):
            get_plot(StubLayer(), 5)

    def test_negative_raises(self):
        class StubLayer:
            def plot(self, idx):
                return f"plot_{idx}"

        with pytest.raises(PlotIndexError):
            get_plot(StubLayer(), -1)


# ===================================================================
# validate_axis 测试
# ===================================================================


class TestValidateAxis:
    def test_valid_x(self):
        assert validate_axis("x") == "x"

    def test_valid_y(self):
        assert validate_axis("y") == "y"

    def test_case_insensitive(self):
        assert validate_axis("X") == "x"
        assert validate_axis("Y") == "y"

    def test_invalid_raises(self):
        with pytest.raises(InvalidAxisError):
            validate_axis("z")

    def test_custom_supported(self):
        assert validate_axis("Z", ("x", "y", "z")) == "z"


# ===================================================================
# validate_column_indices 测试
# ===================================================================


class TestValidateColumnIndices:
    def test_valid(self):
        validate_column_indices([0, 1, 2], 3)

    def test_out_of_range_raises(self):
        with pytest.raises(ColumnIndexError):
            validate_column_indices([0, 5], 3)

    def test_negative_raises(self):
        with pytest.raises(ColumnIndexError):
            validate_column_indices([-1], 3)

    def test_empty_list_ok(self):
        validate_column_indices([], 0)


# ===================================================================
# sanitize_labtalk_name 测试
# ===================================================================


class TestSanitizeLabtalkName:
    def test_valid_name(self):
        assert sanitize_labtalk_name("Graph1") == "Graph1"

    def test_valid_underscore(self):
        assert sanitize_labtalk_name("my_sheet") == "my_sheet"

    def test_empty_raises(self):
        with pytest.raises(ToolError):
            sanitize_labtalk_name("")

    def test_injection_raises(self):
        with pytest.raises(ToolError):
            sanitize_labtalk_name('Graph1"; delete -all;')

    def test_space_raises(self):
        with pytest.raises(ToolError):
            sanitize_labtalk_name("Graph 1")

    def test_digit_start_raises(self):
        with pytest.raises(ToolError):
            sanitize_labtalk_name("1Graph")


# ===================================================================
# tool_error_handler 测试
# ===================================================================


class TestToolErrorHandler:
    def test_success_passthrough(self):
        @tool_error_handler("test_tool")
        def my_tool() -> dict:
            return {"ok": True, "message": "success"}

        result = my_tool()
        assert result["ok"] is True
        assert result["message"] == "success"

    def test_tool_error_caught(self):
        @tool_error_handler("test_tool")
        def my_tool() -> dict:
            raise WorksheetNotFoundError("Sheet99")

        result = my_tool()
        assert result["ok"] is False
        assert "Sheet99" in result["message"]
        assert result["error"]["type"] == "not_found"

    def test_generic_exception_caught(self):
        @tool_error_handler("test_tool", "自定义提示")
        def my_tool() -> dict:
            raise RuntimeError("unexpected")

        result = my_tool()
        assert result["ok"] is False
        assert "unexpected" in result["message"]
        assert result["error"]["type"] == "internal_error"
        assert result["error"]["hint"] == "自定义提示"

    def test_preserves_function_name(self):
        @tool_error_handler("test_tool")
        def my_special_tool() -> dict:
            return {"ok": True, "message": "ok"}

        assert my_special_tool.__name__ == "my_special_tool"

    def test_passes_args_and_kwargs(self):
        @tool_error_handler("test_tool")
        def my_tool(a: int, b: str = "default") -> dict:
            return {"ok": True, "message": f"{a}-{b}"}

        result = my_tool(42, b="custom")
        assert result["message"] == "42-custom"

    def test_nested_tool_error_subclass(self):
        @tool_error_handler("test_tool")
        def my_tool() -> dict:
            raise GraphNotFoundError("Graph99")

        result = my_tool()
        assert result["ok"] is False
        assert result["error"]["type"] == "not_found"
        assert result["error"]["value"] == "Graph99"


# ===================================================================
# get_graph_layer 测试
# ===================================================================


class TestGetGraphLayer:
    def test_default_layer_zero(self):
        from originlab_mcp.utils.helpers import get_graph_layer

        class StubGraph:
            def __getitem__(self, idx):
                return f"layer_{idx}"
            def __len__(self):
                return 3

        result = get_graph_layer(StubGraph(), 0)
        assert result == "layer_0"

    def test_non_zero_layer(self):
        from originlab_mcp.utils.helpers import get_graph_layer

        class StubGraph:
            def __getitem__(self, idx):
                return f"layer_{idx}"
            def __len__(self):
                return 3

        result = get_graph_layer(StubGraph(), 2)
        assert result == "layer_2"

    def test_out_of_range_raises(self):
        from originlab_mcp.exceptions import LayerIndexError
        from originlab_mcp.utils.helpers import get_graph_layer

        class StubGraph:
            def __getitem__(self, idx):
                return f"layer_{idx}"
            def __len__(self):
                return 1

        with pytest.raises(LayerIndexError) as exc_info:
            get_graph_layer(StubGraph(), 3)
        assert exc_info.value.error_type == "invalid_input"
        assert "3" in str(exc_info.value)

    def test_negative_raises(self):
        from originlab_mcp.exceptions import LayerIndexError
        from originlab_mcp.utils.helpers import get_graph_layer

        class StubGraph:
            def __getitem__(self, idx):
                return f"layer_{idx}"
            def __len__(self):
                return 1

        with pytest.raises(LayerIndexError):
            get_graph_layer(StubGraph(), -1)
