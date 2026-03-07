"""图表定制类 tools。

让 AI 能对图表执行高频、低歧义的样式调整。

包含:
    set_axis_range: 设置坐标轴范围
    set_axis_scale: 设置坐标轴缩放类型
    set_axis_title: 设置坐标轴标题
    set_plot_color: 设置曲线颜色
    set_plot_colormap: 设置颜色映射
    set_plot_symbols: 设置数据点符号
"""

from __future__ import annotations

from typing import Any

from originlab_mcp.exceptions import (
    GraphNotFoundError,
    InvalidAxisError,
    NoActiveGraphError,
    PlotIndexError,
    ToolError,
)
from originlab_mcp.origin_manager import OriginManager
from originlab_mcp.utils.constants import (
    SCALE_TYPE_TO_INT,
    ScaleType,
)
from originlab_mcp.utils.validators import (
    error_response,
    error_response_from_exception,
    success_response,
    validate_scale_type,
)


def _resolve_graph_name(
    graph_name: str | None,
    manager: OriginManager,
) -> str:
    """解析图表名称，未指定时使用活动图表。

    Raises:
        NoActiveGraphError: 未指定且无活动图表时。
    """
    name = graph_name or manager.active_graph
    if not name:
        raise NoActiveGraphError()
    return name


def _find_graph(op: Any, graph_name: str) -> Any:
    """查找图表对象，不存在时抛出异常。

    Raises:
        GraphNotFoundError: 图表不存在时。
    """
    gr = op.find_graph(graph_name)
    if gr is None:
        raise GraphNotFoundError(graph_name)
    return gr


def _get_plot(gl: Any, plot_index: int) -> Any:
    """获取指定索引的曲线，不存在时抛出异常。

    Raises:
        PlotIndexError: 曲线索引不存在时。
    """
    plot = gl.plot(plot_index)
    if plot is None:
        raise PlotIndexError(plot_index)
    return plot


def _validate_axis(axis: str, supported: tuple[str, ...] = ("x", "y")) -> str:
    """验证并归一化轴标识。

    Raises:
        InvalidAxisError: 不支持的轴标识时。
    """
    normalized = axis.lower()
    if normalized not in supported:
        raise InvalidAxisError(axis, supported)
    return normalized


def register_customize_tools(mcp: Any) -> None:
    """注册图表定制类 tools 到 MCP Server。"""

    manager = OriginManager()

    # =================================================================
    # set_axis_range
    # =================================================================

    @mcp.tool()
    def set_axis_range(
        axis: str,
        min_val: float,
        max_val: float,
        graph_name: str | None = None,
    ) -> dict:
        """设置图表坐标轴的范围。

        何时使用：需要手动指定坐标轴显示范围（不使用自动缩放）时使用。
        何时不用：使用自动缩放（默认）时无需调用。

        默认行为：
        - graph_name 省略时使用当前活动图表

        示例：
        - set_axis_range(axis="x", min_val=0, max_val=100)
        - set_axis_range(axis="y", min_val=-5, max_val=50, graph_name="Graph1")
        """
        try:
            target_name = _resolve_graph_name(graph_name, manager)
            normalized_axis = _validate_axis(axis)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                gl = gr[0]

                if normalized_axis == "x":
                    gl.set_xlim(min_val, max_val)
                else:
                    gl.set_ylim(min_val, max_val)

                return {
                    "graph_name": target_name,
                    "axis": axis,
                    "min": min_val,
                    "max": max_val,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"轴范围已设置: {axis} = [{min_val}, {max_val}]。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "set_axis_title", "set_axis_scale", "export_graph",
                ],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置轴范围失败: {e}",
                error_type="internal_error",
                target="axis",
                hint="请检查参数值。",
            )

    # =================================================================
    # set_axis_scale
    # =================================================================

    @mcp.tool()
    def set_axis_scale(
        axis: str,
        scale_type: str,
        graph_name: str | None = None,
    ) -> dict:
        """设置坐标轴的缩放类型。

        何时使用：需要将轴改为对数刻度（log/ln）或从对数改回线性时使用。
        何时不用：默认使用线性刻度时无需调用。

        默认行为：
        - graph_name 省略时使用当前活动图表

        示例：
        - set_axis_scale(axis="y", scale_type="log")
        - set_axis_scale(axis="x", scale_type="linear")
        """
        err = validate_scale_type(scale_type)
        if err:
            return error_response(
                message=err,
                error_type="unsupported",
                target="scale_type",
                value=scale_type,
                hint=f"Supported types: {[e.value for e in ScaleType]}",
            )

        try:
            target_name = _resolve_graph_name(graph_name, manager)
            normalized_axis = _validate_axis(axis)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                gl = gr[0]
                scale_val = SCALE_TYPE_TO_INT.get(scale_type, 0)

                if normalized_axis == "x":
                    gl.xscale(scale_val)
                else:
                    gl.yscale(scale_val)

                return {
                    "graph_name": target_name,
                    "axis": axis,
                    "scale_type": scale_type,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"轴缩放已设置: {axis} = {scale_type}。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "set_axis_range", "set_axis_title", "export_graph",
                ],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置轴缩放失败: {e}",
                error_type="internal_error",
                target="axis",
                hint="请检查参数值。",
            )

    # =================================================================
    # set_axis_title
    # =================================================================

    @mcp.tool()
    def set_axis_title(
        axis: str,
        title: str,
        graph_name: str | None = None,
    ) -> dict:
        """设置坐标轴标题。

        何时使用：需要修改或设置 X/Y 轴的文字标题时使用。
        何时不用：轴标题已正确时无需调用。

        默认行为：
        - graph_name 省略时使用当前活动图表

        示例：
        - set_axis_title(axis="x", title="Time (s)")
        - set_axis_title(axis="y", title="Voltage (mV)", graph_name="Graph1")
        """
        try:
            target_name = _resolve_graph_name(graph_name, manager)
            normalized_axis = _validate_axis(axis, ("x", "y", "z"))

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                gl = gr[0]
                ax = gl.axis(normalized_axis)
                ax.title = title

                return {
                    "graph_name": target_name,
                    "axis": axis,
                    "title": title,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"{axis.upper()} 轴标题已设置为 '{title}'。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "set_plot_color", "set_axis_range", "export_graph",
                ],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置轴标题失败: {e}",
                error_type="internal_error",
                target="axis",
                hint="请检查参数值。",
            )

    # =================================================================
    # set_plot_color
    # =================================================================

    @mcp.tool()
    def set_plot_color(
        color: str,
        plot_index: int = 0,
        graph_name: str | None = None,
    ) -> dict:
        """设置指定曲线的颜色。

        何时使用：需要修改图表中某条曲线的颜色时使用。
        何时不用：使用默认配色时无需调用。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - plot_index 默认为 0（第一条曲线）

        参数说明：
        - color: 十六进制颜色值，如 "#ff5833"、"#0000ff"

        示例：
        - set_plot_color(color="#ff0000")
        - set_plot_color(color="#0000ff", plot_index=1, graph_name="Graph1")
        """
        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                plot = _get_plot(gr[0], plot_index)
                plot.color = color

                return {
                    "graph_name": target_name,
                    "plot_index": plot_index,
                    "color": color,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"曲线 {plot_index} 的颜色已设置为 {color}。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "set_axis_title", "set_plot_symbols", "export_graph",
                ],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置颜色失败: {e}",
                error_type="internal_error",
                target="color",
                value=color,
                hint="请使用十六进制颜色值，如 '#ff5833'。",
            )

    # =================================================================
    # set_plot_colormap
    # =================================================================

    @mcp.tool()
    def set_plot_colormap(
        colormap: str,
        plot_index: int = 0,
        graph_name: str | None = None,
    ) -> dict:
        """设置指定曲线的颜色映射表。

        何时使用：需要为数据点应用渐变色或热力图配色方案时使用。
        何时不用：只需设置单一颜色时请用 set_plot_color。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - plot_index 默认为 0（第一条曲线）

        示例：
        - set_plot_colormap(colormap="Candy")
        - set_plot_colormap(colormap="Rainbow", plot_index=0)
        """
        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                plot = _get_plot(gr[0], plot_index)
                plot.colormap = colormap

                return {
                    "graph_name": target_name,
                    "plot_index": plot_index,
                    "colormap": colormap,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"曲线 {plot_index} 的颜色映射已设置为 '{colormap}'。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_plot_color", "export_graph"],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置颜色映射失败: {e}",
                error_type="internal_error",
                target="colormap",
                value=colormap,
                hint="请检查 colormap 名称是否正确。",
            )

    # =================================================================
    # set_plot_symbols
    # =================================================================

    @mcp.tool()
    def set_plot_symbols(
        shape_list: list[int],
        plot_index: int = 0,
        graph_name: str | None = None,
    ) -> dict:
        """设置指定曲线的数据点符号形状。

        何时使用：需要修改散点图或折线+符号图中数据点的形状时使用。
        何时不用：纯折线图或不需要自定义符号时无需调用。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - plot_index 默认为 0（第一条曲线）

        参数说明：
        - shape_list: 符号形状编号列表（Origin 内置编号）

        示例：
        - set_plot_symbols(shape_list=[3, 2, 1])
        """
        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                plot = _get_plot(gr[0], plot_index)
                plot.shapelist = shape_list

                return {
                    "graph_name": target_name,
                    "plot_index": plot_index,
                    "shape_list": shape_list,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"曲线 {plot_index} 的符号形状已更新。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_plot_color", "export_graph"],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置符号失败: {e}",
                error_type="internal_error",
                target="shape_list",
                value=shape_list,
                hint="请检查 shape_list 格式。",
            )
