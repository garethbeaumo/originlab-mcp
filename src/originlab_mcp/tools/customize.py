"""
图表定制类 tools

让 AI 能对图表执行高频、低歧义的样式调整。

包含：
- set_axis_range
- set_axis_scale
- set_axis_title
- set_plot_color
- set_plot_colormap
- set_plot_symbols
"""

from __future__ import annotations

from originlab_mcp.origin_manager import OriginManager
from originlab_mcp.utils.constants import (
    SCALE_TYPE_TO_INT,
    AxisId,
    ScaleType,
)
from originlab_mcp.utils.validators import (
    error_response,
    success_response,
    validate_scale_type,
)


def register_customize_tools(mcp) -> None:
    """注册图表定制类 tools 到 MCP Server。"""

    manager = OriginManager()

    def _find_graph(op, graph_name: str | None):
        """查找图表。返回 (graph_obj, resolved_name) 或 (None, name)。"""
        name = graph_name or manager.active_graph
        if not name:
            return None, None
        gr = op.find_graph(name)
        return gr, name

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
            def _set(op):
                gr, name = _find_graph(op, graph_name)
                if name is None:
                    return None, "no_graph"
                if gr is None:
                    return name, "not_found"

                gl = gr[0]
                if axis.lower() == "x":
                    gl.set_xlim(min_val, max_val)
                elif axis.lower() == "y":
                    gl.set_ylim(min_val, max_val)
                else:
                    return axis, "bad_axis"

                return {
                    "graph_name": name,
                    "axis": axis,
                    "min": min_val,
                    "max": max_val,
                }, "ok"

            result, status = manager.execute(_set)

            if status == "no_graph":
                return error_response(
                    message="未指定图表且无活动图表",
                    error_type="invalid_input",
                    target="graph_name",
                    hint="请指定 graph_name 或先创建图表。",
                )
            if status == "not_found":
                return error_response(
                    message=f"图表 '{result}' 不存在",
                    error_type="not_found",
                    target="graph",
                    value=result,
                    hint="Call list_graphs to inspect available graph names.",
                )
            if status == "bad_axis":
                return error_response(
                    message=f"不支持的轴标识 '{result}'",
                    error_type="invalid_input",
                    target="axis",
                    value=result,
                    hint="Supported axis values: 'x', 'y'.",
                )

            return success_response(
                message=f"轴范围已设置: {axis} = [{min_val}, {max_val}]。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_axis_title", "set_axis_scale", "export_graph"],
            )
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
            def _set(op):
                gr, name = _find_graph(op, graph_name)
                if name is None:
                    return None, "no_graph"
                if gr is None:
                    return name, "not_found"

                gl = gr[0]
                scale_val = SCALE_TYPE_TO_INT.get(scale_type, 0)

                if axis.lower() == "x":
                    gl.xscale(scale_val)
                elif axis.lower() == "y":
                    gl.yscale(scale_val)
                else:
                    return axis, "bad_axis"

                return {
                    "graph_name": name,
                    "axis": axis,
                    "scale_type": scale_type,
                }, "ok"

            result, status = manager.execute(_set)

            if status == "no_graph":
                return error_response(
                    message="未指定图表且无活动图表",
                    error_type="invalid_input",
                    target="graph_name",
                    hint="请指定 graph_name 或先创建图表。",
                )
            if status == "not_found":
                return error_response(
                    message=f"图表 '{result}' 不存在",
                    error_type="not_found",
                    target="graph",
                    value=result,
                    hint="Call list_graphs to inspect available graph names.",
                )
            if status == "bad_axis":
                return error_response(
                    message=f"不支持的轴标识 '{result}'",
                    error_type="invalid_input",
                    target="axis",
                    value=result,
                    hint="Supported axis values: 'x', 'y'.",
                )

            return success_response(
                message=f"轴缩放已设置: {axis} = {scale_type}。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_axis_range", "set_axis_title", "export_graph"],
            )
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
            def _set(op):
                gr, name = _find_graph(op, graph_name)
                if name is None:
                    return None, "no_graph"
                if gr is None:
                    return name, "not_found"

                gl = gr[0]

                if axis.lower() not in ("x", "y", "z"):
                    return axis, "bad_axis"

                ax = gl.axis(axis.lower())
                ax.title = title

                return {
                    "graph_name": name,
                    "axis": axis,
                    "title": title,
                }, "ok"

            result, status = manager.execute(_set)

            if status == "no_graph":
                return error_response(
                    message="未指定图表且无活动图表",
                    error_type="invalid_input",
                    target="graph_name",
                    hint="请指定 graph_name 或先创建图表。",
                )
            if status == "not_found":
                return error_response(
                    message=f"图表 '{result}' 不存在",
                    error_type="not_found",
                    target="graph",
                    value=result,
                    hint="Call list_graphs to inspect available graph names.",
                )
            if status == "bad_axis":
                return error_response(
                    message=f"不支持的轴标识 '{result}'",
                    error_type="invalid_input",
                    target="axis",
                    value=result,
                    hint="Supported axis values: 'x', 'y', 'z'.",
                )

            return success_response(
                message=f"{axis.upper()} 轴标题已设置为 '{title}'。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_plot_color", "set_axis_range", "export_graph"],
            )
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
            def _set(op):
                gr, name = _find_graph(op, graph_name)
                if name is None:
                    return None, "no_graph"
                if gr is None:
                    return name, "not_found"

                gl = gr[0]
                plot = gl.plot(plot_index)
                if plot is None:
                    return plot_index, "bad_index"

                plot.color = color

                return {
                    "graph_name": name,
                    "plot_index": plot_index,
                    "color": color,
                }, "ok"

            result, status = manager.execute(_set)

            if status == "no_graph":
                return error_response(
                    message="未指定图表且无活动图表",
                    error_type="invalid_input",
                    target="graph_name",
                    hint="请指定 graph_name 或先创建图表。",
                )
            if status == "not_found":
                return error_response(
                    message=f"图表 '{result}' 不存在",
                    error_type="not_found",
                    target="graph",
                    value=result,
                    hint="Call list_graphs to inspect available graph names.",
                )
            if status == "bad_index":
                return error_response(
                    message=f"曲线索引 {result} 不存在",
                    error_type="invalid_input",
                    target="plot_index",
                    value=result,
                    hint="请检查图表中的曲线数量。plot_index 从 0 开始。",
                )

            return success_response(
                message=f"曲线 {plot_index} 的颜色已设置为 {color}。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_axis_title", "set_plot_symbols", "export_graph"],
            )
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
            def _set(op):
                gr, name = _find_graph(op, graph_name)
                if name is None:
                    return None, "no_graph"
                if gr is None:
                    return name, "not_found"

                gl = gr[0]
                plot = gl.plot(plot_index)
                if plot is None:
                    return plot_index, "bad_index"

                plot.colormap = colormap

                return {
                    "graph_name": name,
                    "plot_index": plot_index,
                    "colormap": colormap,
                }, "ok"

            result, status = manager.execute(_set)

            if status == "no_graph":
                return error_response(
                    message="未指定图表且无活动图表",
                    error_type="invalid_input",
                    target="graph_name",
                    hint="请指定 graph_name 或先创建图表。",
                )
            if status == "not_found":
                return error_response(
                    message=f"图表 '{result}' 不存在",
                    error_type="not_found",
                    target="graph",
                    value=result,
                    hint="Call list_graphs to inspect available graph names.",
                )
            if status == "bad_index":
                return error_response(
                    message=f"曲线索引 {result} 不存在",
                    error_type="invalid_input",
                    target="plot_index",
                    value=result,
                    hint="请检查图表中的曲线数量。",
                )

            return success_response(
                message=f"曲线 {plot_index} 的颜色映射已设置为 '{colormap}'。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_plot_color", "export_graph"],
            )
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
            def _set(op):
                gr, name = _find_graph(op, graph_name)
                if name is None:
                    return None, "no_graph"
                if gr is None:
                    return name, "not_found"

                gl = gr[0]
                plot = gl.plot(plot_index)
                if plot is None:
                    return plot_index, "bad_index"

                plot.shapelist = shape_list

                return {
                    "graph_name": name,
                    "plot_index": plot_index,
                    "shape_list": shape_list,
                }, "ok"

            result, status = manager.execute(_set)

            if status == "no_graph":
                return error_response(
                    message="未指定图表且无活动图表",
                    error_type="invalid_input",
                    target="graph_name",
                    hint="请指定 graph_name 或先创建图表。",
                )
            if status == "not_found":
                return error_response(
                    message=f"图表 '{result}' 不存在",
                    error_type="not_found",
                    target="graph",
                    value=result,
                    hint="Call list_graphs to inspect available graph names.",
                )
            if status == "bad_index":
                return error_response(
                    message=f"曲线索引 {result} 不存在",
                    error_type="invalid_input",
                    target="plot_index",
                    value=result,
                    hint="请检查图表中的曲线数量。",
                )

            return success_response(
                message=f"曲线 {plot_index} 的符号形状已更新。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_plot_color", "export_graph"],
            )
        except Exception as e:
            return error_response(
                message=f"设置符号失败: {e}",
                error_type="internal_error",
                target="shape_list",
                value=shape_list,
                hint="请检查 shape_list 格式。",
            )
