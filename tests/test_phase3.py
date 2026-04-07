"""测试 Phase 3 新增功能：sanitize_labtalk_name、统一 resolve 范式等。"""

from types import MethodType

import pytest

from originlab_mcp.exceptions import (
    NoActiveGraphError,
    NoActiveWorksheetError,
    ToolError,
)
from originlab_mcp.origin_manager import OriginManager
from originlab_mcp.utils.helpers import sanitize_labtalk_name


# ===================================================================
# DummyMCP 复用
# ===================================================================


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
    return OriginManager()


# ===================================================================
# sanitize_labtalk_name 测试
# ===================================================================


class TestSanitizeLabtalkName:
    """LabTalk 命令注入防护函数测试。"""

    def test_valid_simple_name(self):
        assert sanitize_labtalk_name("Graph1") == "Graph1"

    def test_valid_underscore(self):
        assert sanitize_labtalk_name("my_sheet_1") == "my_sheet_1"

    def test_valid_underscore_start(self):
        assert sanitize_labtalk_name("_private") == "_private"

    def test_empty_raises(self):
        with pytest.raises(ToolError) as exc_info:
            sanitize_labtalk_name("")
        assert exc_info.value.error_type == "invalid_input"

    def test_none_like_empty(self):
        """空字符串应抛出异常。"""
        with pytest.raises(ToolError):
            sanitize_labtalk_name("")

    def test_injection_semicolon(self):
        """分号注入应被拦截。"""
        with pytest.raises(ToolError) as exc_info:
            sanitize_labtalk_name('Graph1"; delete -all;')
        assert "非法字符" in str(exc_info.value)

    def test_injection_space(self):
        """空格注入应被拦截。"""
        with pytest.raises(ToolError):
            sanitize_labtalk_name("Graph 1")

    def test_injection_quote(self):
        """引号注入应被拦截。"""
        with pytest.raises(ToolError):
            sanitize_labtalk_name('Graph"1')

    def test_starts_with_digit(self):
        """不允许数字开头。"""
        with pytest.raises(ToolError):
            sanitize_labtalk_name("1Graph")

    def test_chinese_chars_rejected(self):
        """非 ASCII 字符应被拒绝。"""
        with pytest.raises(ToolError):
            sanitize_labtalk_name("图表1")

    def test_dash_rejected(self):
        """连字符应被拒绝。"""
        with pytest.raises(ToolError):
            sanitize_labtalk_name("my-graph")

    def test_custom_param_name(self):
        """param_name 应出现在错误提示中。"""
        with pytest.raises(ToolError) as exc_info:
            sanitize_labtalk_name("", param_name="graph_name")
        assert "graph_name" in str(exc_info.value)


# ===================================================================
# plot.py 统一 resolve 范式测试
# ===================================================================


class TestPlotResolveUnified:
    """验证 plot.py 函数使用统一的 resolve 范式。"""

    def test_remove_plot_no_active_graph(self, fresh_manager):
        """无活动图表时应返回错误。"""
        from originlab_mcp.tools.plot import register_plot_tools

        mcp = DummyMCP()
        register_plot_tools(mcp, fresh_manager)

        result = mcp.tools["remove_plot_from_graph"](plot_index=0)
        assert result["ok"] is False
        # 通过 @tool_error_handler 捕获 NoActiveGraphError
        assert result["error"]["target"] == "graph_name"

    def test_add_plot_no_active_graph(self, fresh_manager):
        """add_plot_to_graph 无活动图表时应返回错误。"""
        from originlab_mcp.tools.plot import register_plot_tools

        mcp = DummyMCP()
        register_plot_tools(mcp, fresh_manager)

        result = mcp.tools["add_plot_to_graph"](x_col=0, y_cols=1)
        assert result["ok"] is False

    def test_add_graph_layer_no_active(self, fresh_manager):
        """add_graph_layer 无活动图表时应返回错误。"""
        from originlab_mcp.tools.plot import register_plot_tools

        mcp = DummyMCP()
        register_plot_tools(mcp, fresh_manager)

        result = mcp.tools["add_graph_layer"]()
        assert result["ok"] is False

    def test_create_double_y_no_active_sheet(self, fresh_manager):
        """create_double_y_plot 无活动工作表时应返回错误。"""
        from originlab_mcp.tools.plot import register_plot_tools

        mcp = DummyMCP()
        register_plot_tools(mcp, fresh_manager)

        result = mcp.tools["create_double_y_plot"](x_col=0, y1_col=1, y2_col=2)
        assert result["ok"] is False


# ===================================================================
# advanced.py 装饰器和防护测试
# ===================================================================


class TestAdvancedTools:
    """验证 advanced.py 的 @tool_error_handler 和命令限制。"""

    def test_empty_command(self, fresh_manager):
        from originlab_mcp.tools.advanced import register_advanced_tools

        mcp = DummyMCP()
        register_advanced_tools(mcp, fresh_manager)

        result = mcp.tools["execute_labtalk"](command="")
        assert result["ok"] is False
        assert result["error"]["target"] == "command"

    def test_whitespace_command(self, fresh_manager):
        from originlab_mcp.tools.advanced import register_advanced_tools

        mcp = DummyMCP()
        register_advanced_tools(mcp, fresh_manager)

        result = mcp.tools["execute_labtalk"](command="   ")
        assert result["ok"] is False

    def test_too_long_command(self, fresh_manager):
        from originlab_mcp.tools.advanced import register_advanced_tools

        mcp = DummyMCP()
        register_advanced_tools(mcp, fresh_manager)

        result = mcp.tools["execute_labtalk"](command="x" * 2001)
        assert result["ok"] is False
        assert "过长" in result["message"]
