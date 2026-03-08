"""图表定制类 tools。

让 AI 能对图表执行高频、低歧义的样式调整。

包含:
    set_axis_range: 设置坐标轴范围
    set_axis_scale: 设置坐标轴缩放类型
    set_axis_title: 设置坐标轴标题
    set_plot_color: 设置曲线颜色
    set_plot_colormap: 设置颜色映射
    set_plot_symbols: 设置数据点符号
    set_plot_line_width: 设置曲线线宽
    set_plot_line_style: 设置曲线线型（实线/虚线/点线等）
    set_error_bar_style: 设置误差棒样式（线宽/端帽/颜色/方向）
    set_legend: 控制图例显示、位置和字号
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
    SCALE_TYPE_TO_ORIGIN,
    ScaleType,
)
from originlab_mcp.utils.helpers import (
    find_graph as _find_graph,
    get_plot as _get_plot,
    resolve_graph_name as _resolve_graph_name,
    validate_axis as _validate_axis,
)
from originlab_mcp.utils.validators import (
    error_response,
    error_response_from_exception,
    success_response,
    validate_scale_type,
)


# 注: _resolve_graph_name, _find_graph, _get_plot, _validate_axis 从 utils.helpers 导入


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
                hint=f"支持的类型: {[e.value for e in ScaleType]}",
            )

        try:
            target_name = _resolve_graph_name(graph_name, manager)
            normalized_axis = _validate_axis(axis)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                gl = gr[0]
                scale_val = SCALE_TYPE_TO_ORIGIN[scale_type]

                if normalized_axis == "x":
                    gl.xscale = scale_val
                else:
                    gl.yscale = scale_val

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

    # =================================================================
    # set_plot_transparency
    # =================================================================

    @mcp.tool()
    def set_plot_transparency(
        transparency: int,
        plot_index: int = 0,
        graph_name: str | None = None,
    ) -> dict:
        """设置指定曲线的透明度。

        何时使用：需要调整曲线或数据点的透明度时使用，常用于多曲线重叠的场景。
        何时不用：不需要透明效果时无需调用。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - plot_index 默认为 0（第一条曲线）

        参数说明：
        - transparency: 透明度百分比，0=完全不透明，100=完全透明

        示例：
        - set_plot_transparency(transparency=50)
        - set_plot_transparency(transparency=30, plot_index=1)
        """
        if transparency < 0 or transparency > 100:
            return error_response(
                message="transparency 必须在 0-100 之间",
                error_type="invalid_input",
                target="transparency",
                value=transparency,
                hint="0=完全不透明，100=完全透明。",
            )

        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                plot = _get_plot(gr[0], plot_index)
                plot.transparency = transparency

                return {
                    "graph_name": target_name,
                    "plot_index": plot_index,
                    "transparency": transparency,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"曲线 {plot_index} 的透明度已设置为 {transparency}%。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_plot_color", "export_graph"],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置透明度失败: {e}",
                error_type="internal_error",
                target="transparency",
                value=transparency,
                hint="请检查参数值。",
            )

    # =================================================================
    # set_symbol_size
    # =================================================================

    @mcp.tool()
    def set_symbol_size(
        size: float,
        plot_index: int = 0,
        graph_name: str | None = None,
    ) -> dict:
        """设置指定曲线的数据点符号大小。

        何时使用：需要调整散点图或折线+符号图中数据点的大小时使用。
        何时不用：使用默认大小时无需调用。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - plot_index 默认为 0（第一条曲线）

        示例：
        - set_symbol_size(size=12)
        - set_symbol_size(size=8.5, plot_index=1)
        """
        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                plot = _get_plot(gr[0], plot_index)
                plot.symbol_size = size

                return {
                    "graph_name": target_name,
                    "plot_index": plot_index,
                    "symbol_size": size,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"曲线 {plot_index} 的符号大小已设置为 {size}。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_plot_color", "set_plot_symbols", "export_graph"],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置符号大小失败: {e}",
                error_type="internal_error",
                target="size",
                value=size,
                hint="请检查参数值。",
            )

    # =================================================================
    # set_fill_area
    # =================================================================

    @mcp.tool()
    def set_fill_area(
        above_color: int,
        fill_type: int = 9,
        below_color: int | None = None,
        plot_index: int = 0,
        graph_name: str | None = None,
    ) -> dict:
        """设置折线图曲线下方的填充区域。

        何时使用：需要在折线图的曲线下方填充颜色时使用（如展示面积、置信区间等）。
        何时不用：散点图或不需要填充效果时无需调用。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - plot_index 默认为 0（第一条曲线）
        - fill_type 默认为 9（填充到下一条曲线）

        参数说明：
        - above_color: 上方填充颜色编号
        - fill_type: 填充选项（9=填充到下一条曲线）
        - below_color: 下方填充颜色编号（可选）

        示例：
        - set_fill_area(above_color=2)
        - set_fill_area(above_color=2, fill_type=9, below_color=3)
        """
        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                plot = _get_plot(gr[0], plot_index)
                if below_color is not None:
                    plot.set_fill_area(above_color, fill_type, below_color)
                else:
                    plot.set_fill_area(above_color, fill_type)

                return {
                    "graph_name": target_name,
                    "plot_index": plot_index,
                    "above_color": above_color,
                    "fill_type": fill_type,
                    "below_color": below_color,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"曲线 {plot_index} 的填充区域已设置。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_plot_color", "export_graph"],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置填充区域失败: {e}",
                error_type="internal_error",
                target="fill_area",
                hint="请检查参数值。此功能仅适用于折线图。",
            )

    # =================================================================
    # get_graph_info
    # =================================================================

    @mcp.tool()
    def get_graph_info(
        graph_name: str | None = None,
    ) -> dict:
        """获取图表的详细信息，包括图层数、曲线列表和数据源。

        何时使用：需要了解图表中有哪些曲线、各曲线的数据源和样式信息时使用。
        何时不用：只需知道图表列表时请用 list_graphs。

        默认行为：
        - graph_name 省略时使用当前活动图表

        示例：
        - get_graph_info()
        - get_graph_info(graph_name="Graph1")
        """
        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _info(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)

                layers = []
                for li in range(len(gr)):
                    gl = gr[li]
                    plots = []
                    try:
                        plot_items = gl.plot_list()
                        for pi, plot in enumerate(plot_items):
                            plot_info: dict[str, Any] = {
                                "index": pi,
                                "range": plot.lt_range() if hasattr(plot, 'lt_range') else None,
                            }
                            try:
                                r, g, b = plot.color
                                plot_info["color"] = f"#{r:02x}{g:02x}{b:02x}"
                            except Exception:
                                pass
                            plots.append(plot_info)
                    except Exception:
                        pass

                    layer_info: dict[str, Any] = {
                        "index": li,
                        "plot_count": len(plots),
                        "plots": plots,
                    }
                    layers.append(layer_info)

                return {
                    "graph_name": target_name,
                    "layer_count": len(layers),
                    "layers": layers,
                }

            result = manager.execute(_info)

            total_plots = sum(l["plot_count"] for l in result["layers"])

            return success_response(
                message=(
                    f"图表 '{target_name}' 共 {result['layer_count']} 个图层，"
                    f"{total_plots} 条曲线。"
                ),
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "set_plot_color",
                    "set_axis_title",
                    "export_graph",
                ],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"获取图表信息失败: {e}",
                error_type="internal_error",
                target="graph_name",
                hint="请检查图表名称。调用 list_graphs 查看可用图表。",
            )

    # =================================================================
    # add_text_label
    # =================================================================

    @mcp.tool()
    def add_text_label(
        text: str,
        x: float | None = None,
        y: float | None = None,
        graph_name: str | None = None,
    ) -> dict:
        """在图表上添加文字标注。

        何时使用：需要在图表上标注峰位、拐点、说明文字等注释时使用。
        何时不用：设置坐标轴标题请用 set_axis_title。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - x, y 省略时由 Origin 自动放置

        参数说明：
        - text: 标注文字内容
        - x, y: 标注位置坐标（图层坐标系）

        示例：
        - add_text_label(text="Peak A")
        - add_text_label(text="Transition Point", x=3.5, y=100)
        """
        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _add(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                gl = gr[0]
                label_obj = gl.add_label(text, x, y)
                label_name = label_obj.name if hasattr(label_obj, 'name') else None

                return {
                    "graph_name": target_name,
                    "text": text,
                    "x": x,
                    "y": y,
                    "label_name": label_name,
                }

            result = manager.execute(_add)

            return success_response(
                message=f"已在图表 '{target_name}' 上添加文字标注 '{text}'。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["add_line_to_graph", "export_graph"],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"添加文字标注失败: {e}",
                error_type="internal_error",
                target="text",
                hint="请检查图表是否存在。",
            )

    # =================================================================
    # add_line_to_graph
    # =================================================================

    @mcp.tool()
    def add_line_to_graph(
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        line_width: float = 1.0,
        arrow: bool = False,
        graph_name: str | None = None,
    ) -> dict:
        """在图表上添加线条或箭头。

        何时使用：需要在图表上绘制参考线、指示线或箭头时使用。
        何时不用：设置坐标轴请用 set_axis_range。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - arrow 为 True 时在终点添加箭头

        参数说明：
        - x1, y1: 起点坐标
        - x2, y2: 终点坐标
        - line_width: 线宽
        - arrow: 是否在终点添加箭头

        示例：
        - add_line_to_graph(x1=0, y1=0, x2=10, y2=10)
        - add_line_to_graph(x1=5, y1=0, x2=5, y2=100, arrow=True)
        """
        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _add(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                gl = gr[0]
                line_obj = gl.add_line(x1, y1, x2, y2)
                if hasattr(line_obj, 'width'):
                    line_obj.width = line_width
                if arrow and hasattr(line_obj, 'set_int'):
                    line_obj.set_int('arrowendshape', 2)

                return {
                    "graph_name": target_name,
                    "start": [x1, y1],
                    "end": [x2, y2],
                    "line_width": line_width,
                    "arrow": arrow,
                }

            result = manager.execute(_add)

            desc = "箭头" if arrow else "线条"
            return success_response(
                message=f"已在图表 '{target_name}' 上添加{desc} ({x1},{y1})->({x2},{y2})。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["add_text_label", "export_graph"],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"添加线条失败: {e}",
                error_type="internal_error",
                target="line",
                hint="请检查坐标值和图表是否存在。",
            )

    # =================================================================
    # remove_graph_label
    # =================================================================

    @mcp.tool()
    def remove_graph_label(
        label_name: str,
        graph_name: str | None = None,
    ) -> dict:
        """移除图表上的标签/注释对象。

        何时使用：需要删除之前添加的文字标注或其他标签对象时使用。
        何时不用：不确定标签名称时先调用 get_graph_info 查看。

        默认行为：
        - graph_name 省略时使用当前活动图表

        参数说明：
        - label_name: 要移除的标签名称（如 "xb" 为X轴标题标签）

        示例：
        - remove_graph_label(label_name="Text1")
        - remove_graph_label(label_name="xb")
        """
        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _remove(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                gl = gr[0]
                gl.remove_label(label_name)

                return {
                    "graph_name": target_name,
                    "removed_label": label_name,
                }

            result = manager.execute(_remove)

            return success_response(
                message=f"已移除标签 '{label_name}'。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["add_text_label", "get_graph_info"],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"移除标签失败: {e}",
                error_type="internal_error",
                target="label_name",
                value=label_name,
                hint="请检查标签名称是否正确。",
            )

    # =================================================================
    # set_graph_title
    # =================================================================

    @mcp.tool()
    def set_graph_title(
        title: str,
        graph_name: str | None = None,
    ) -> dict:
        """设置图表的大标题（图层标题）。

        何时使用：需要为图表设置或修改主标题时使用。
        何时不用：设置坐标轴标题请用 set_axis_title。

        默认行为：
        - graph_name 省略时使用当前活动图表

        示例：
        - set_graph_title(title="Temperature vs Time")
        - set_graph_title(title="Fig. 1", graph_name="Graph1")
        """
        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _set_title(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                gr.lname = title

                return {
                    "graph_name": target_name,
                    "title": title,
                }

            result = manager.execute(_set_title)

            return success_response(
                message=f"图表标题已设置为 '{title}'。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_axis_title", "export_graph"],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置图表标题失败: {e}",
                error_type="internal_error",
                target="title",
                hint="请检查图表是否存在。",
            )

    # =================================================================
    # set_axis_step
    # =================================================================

    @mcp.tool()
    def set_axis_step(
        axis: str,
        step: float,
        graph_name: str | None = None,
    ) -> dict:
        """设置坐标轴的刻度步长（增量）。

        何时使用：需要精确控制坐标轴刻度间距时使用。
        何时不用：只需设置范围请用 set_axis_range。

        默认行为：
        - graph_name 省略时使用当前活动图表

        参数说明：
        - axis: 轴标识，"x" 或 "y"
        - step: 刻度步长值

        示例：
        - set_axis_step(axis="x", step=0.5)
        - set_axis_step(axis="y", step=10)
        """
        if axis not in ("x", "y", "z"):
            return error_response(
                message=f"不支持的轴标识 '{axis}'",
                error_type="invalid_input",
                target="axis",
                value=axis,
                hint="支持的轴: x, y, z",
            )

        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                gl = gr[0]

                if axis == "x":
                    gl.set_xlim(step=step)
                elif axis == "y":
                    gl.set_ylim(step=step)
                elif axis == "z":
                    gl.set_zlim(step=step)

                return {
                    "graph_name": target_name,
                    "axis": axis,
                    "step": step,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"{axis.upper()} 轴刻度步长已设置为 {step}。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_axis_range", "export_graph"],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置轴步长失败: {e}",
                error_type="internal_error",
                target="step",
                value=step,
                hint="请检查步长值是否合理。",
            )

    # =================================================================
    # set_symbol_interior
    # =================================================================

    @mcp.tool()
    def set_symbol_interior(
        interior: int,
        plot_index: int = 0,
        graph_name: str | None = None,
    ) -> dict:
        """设置数据点符号的内部填充类型。

        何时使用：需要区分多条曲线时使用（如实心 vs 空心符号）。
        何时不用：使用默认填充时无需调用。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - plot_index 默认为 0（第一条曲线）

        参数说明：
        - interior: 填充类型编号
            - 0 = 无符号
            - 1 = 实心
            - 2 = 空心
            - 3 = 圆点中心

        示例：
        - set_symbol_interior(interior=2)
        - set_symbol_interior(interior=1, plot_index=1)
        """
        if interior not in (0, 1, 2, 3):
            return error_response(
                message="interior 必须是 0, 1, 2, 3 之一",
                error_type="invalid_input",
                target="interior",
                value=interior,
                hint="0=无符号, 1=实心, 2=空心, 3=圆点中心",
            )

        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                plot = _get_plot(gr[0], plot_index)
                plot.symbol_interior = interior

                interiors = {0: "无符号", 1: "实心", 2: "空心", 3: "圆点中心"}
                return {
                    "graph_name": target_name,
                    "plot_index": plot_index,
                    "interior": interior,
                    "interior_name": interiors.get(interior, ""),
                }

            result = manager.execute(_set)

            return success_response(
                message=f"曲线 {plot_index} 的符号填充已设为 {result['interior_name']}。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_symbol_size", "set_plot_color"],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置符号填充失败: {e}",
                error_type="internal_error",
                target="interior",
                value=interior,
                hint="请检查参数值。",
            )

    # =================================================================
    # set_color_increment
    # =================================================================

    @mcp.tool()
    def set_color_increment(
        increment: int,
        plot_index: int = 0,
        graph_name: str | None = None,
    ) -> dict:
        """设置分组曲线的颜色增量。

        何时使用：多条分组曲线需要自动递增颜色时使用。
        何时不用：手动设置每条曲线颜色时请用 set_plot_color。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - plot_index 默认为 0（通常设在组头曲线上）

        参数说明：
        - increment: 颜色增量值（1=每条曲线递增一个颜色）

        示例：
        - set_color_increment(increment=1)
        """
        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                plot = _get_plot(gr[0], plot_index)
                plot.colorinc = increment

                return {
                    "graph_name": target_name,
                    "plot_index": plot_index,
                    "color_increment": increment,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"曲线 {plot_index} 的颜色增量已设置为 {increment}。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_symbol_increment", "set_plot_color"],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置颜色增量失败: {e}",
                error_type="internal_error",
                target="increment",
                value=increment,
                hint="请检查参数值。",
            )

    # =================================================================
    # set_symbol_increment
    # =================================================================

    @mcp.tool()
    def set_symbol_increment(
        increment: int,
        plot_index: int = 0,
        graph_name: str | None = None,
    ) -> dict:
        """设置分组曲线的符号形状增量。

        何时使用：多条分组曲线需要自动递增符号形状时使用。
        何时不用：手动设置每条曲线符号时请用 set_plot_symbols。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - plot_index 默认为 0（通常设在组头曲线上）

        参数说明：
        - increment: 符号形状增量值（1=每条曲线递增一个形状）

        示例：
        - set_symbol_increment(increment=1)
        """
        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                plot = _get_plot(gr[0], plot_index)
                plot.symbol_kindinc = increment

                return {
                    "graph_name": target_name,
                    "plot_index": plot_index,
                    "symbol_increment": increment,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"曲线 {plot_index} 的符号形状增量已设置为 {increment}。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=["set_color_increment", "set_plot_symbols"],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置符号增量失败: {e}",
                error_type="internal_error",
                target="increment",
                value=increment,
                hint="请检查参数值。",
            )

    # =================================================================
    # set_plot_line_width
    # =================================================================

    @mcp.tool()
    def set_plot_line_width(
        width: float,
        plot_index: int = 0,
        graph_name: str | None = None,
    ) -> dict:
        """设置指定曲线的线宽。

        何时使用：需要调整折线图或折线+符号图的线条粗细时使用。
        何时不用：使用默认线宽时无需调用。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - plot_index 默认为 0（第一条曲线）

        参数说明：
        - width: 线宽值（单位：点），常用值 0.5, 1, 1.5, 2, 3

        示例：
        - set_plot_line_width(width=2)
        - set_plot_line_width(width=1.5, plot_index=1)
        """
        if width <= 0:
            return error_response(
                message="width 必须大于 0",
                error_type="invalid_input",
                target="width",
                value=width,
                hint="常用线宽值: 0.5, 1, 1.5, 2, 3",
            )

        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                gl = gr[0]
                # 验证 plot_index 有效
                _get_plot(gl, plot_index)

                # 通过 LabTalk 设置线宽
                op.lt_exec(f'win -a {gr.name}')
                op.lt_exec(f'layer -s 1')
                op.lt_exec(f'layer.plot = {plot_index + 1}')
                op.lt_exec(f'set %C -w {width}')

                return {
                    "graph_name": target_name,
                    "plot_index": plot_index,
                    "width": width,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"曲线 {plot_index} 的线宽已设置为 {width}。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "set_plot_line_style", "set_plot_color", "export_graph",
                ],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置线宽失败: {e}",
                error_type="internal_error",
                target="width",
                value=width,
                hint="请检查参数值。",
            )

    # =================================================================
    # set_plot_line_style
    # =================================================================

    # 线型编号映射
    LINE_STYLE_MAP = {
        "solid": 1,
        "dash": 2,
        "dot": 3,
        "dashdot": 4,
        "dashdotdot": 5,
        "short_dash": 6,
        "short_dot": 7,
        "short_dashdot": 8,
    }

    @mcp.tool()
    def set_plot_line_style(
        style: str,
        plot_index: int = 0,
        graph_name: str | None = None,
    ) -> dict:
        """设置指定曲线的线型（实线、虚线、点线等）。

        何时使用：需要区分多条曲线或满足论文排版要求时使用。
        何时不用：使用默认实线时无需调用。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - plot_index 默认为 0（第一条曲线）

        参数说明：
        - style: 线型名称
            - "solid" = 实线
            - "dash" = 虚线
            - "dot" = 点线
            - "dashdot" = 点划线
            - "dashdotdot" = 双点划线
            - "short_dash" = 短虚线
            - "short_dot" = 短点线
            - "short_dashdot" = 短点划线

        示例：
        - set_plot_line_style(style="dash")
        - set_plot_line_style(style="dot", plot_index=1)
        """
        style_lower = style.lower()
        if style_lower not in LINE_STYLE_MAP:
            return error_response(
                message=f"不支持的线型 '{style}'",
                error_type="invalid_input",
                target="style",
                value=style,
                hint=f"支持的线型: {list(LINE_STYLE_MAP.keys())}",
            )

        try:
            target_name = _resolve_graph_name(graph_name, manager)
            style_code = LINE_STYLE_MAP[style_lower]

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                gl = gr[0]
                _get_plot(gl, plot_index)

                op.lt_exec(f'win -a {gr.name}')
                op.lt_exec(f'layer -s 1')
                op.lt_exec(f'layer.plot = {plot_index + 1}')
                op.lt_exec(f'set %C -d {style_code}')

                return {
                    "graph_name": target_name,
                    "plot_index": plot_index,
                    "style": style,
                    "style_code": style_code,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"曲线 {plot_index} 的线型已设置为 {style}。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "set_plot_line_width", "set_plot_color", "export_graph",
                ],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置线型失败: {e}",
                error_type="internal_error",
                target="style",
                value=style,
                hint=f"支持的线型: {list(LINE_STYLE_MAP.keys())}",
            )

    # =================================================================
    # set_error_bar_style
    # =================================================================

    @mcp.tool()
    def set_error_bar_style(
        plot_index: int = 0,
        line_width: float | None = None,
        cap_width: float | None = None,
        color: str | None = None,
        direction: str | None = None,
        graph_name: str | None = None,
    ) -> dict:
        """设置误差棒的样式（线宽、端帽、颜色、方向）。

        何时使用：需要调整误差棒的外观时使用（如论文排版、视觉区分等）。
        何时不用：使用默认误差棒样式时无需调用。前提是已通过 create_plot 或 add_plot_to_graph 的 yerr_col 参数创建了误差棒。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - plot_index 默认为 0（第一条曲线）
        - 未提供的参数不会被修改

        参数说明：
        - line_width: 误差棒线宽（单位：点），常用值 1, 1.5, 2
        - cap_width: 误差棒端帽宽度（单位：点），常用值 5, 8, 10, 15
        - color: 误差棒颜色，十六进制值如 "#ff0000"；设为 "auto" 则跟随曲线颜色
        - direction: 误差棒方向
            - "both" = 上下双向（默认）
            - "plus" = 仅向上
            - "minus" = 仅向下

        示例：
        - set_error_bar_style(line_width=2, cap_width=10)
        - set_error_bar_style(color="#000000", plot_index=1)
        - set_error_bar_style(direction="plus")
        """
        if line_width is not None and line_width <= 0:
            return error_response(
                message="line_width 必须大于 0",
                error_type="invalid_input",
                target="line_width",
                value=line_width,
            )
        if cap_width is not None and cap_width < 0:
            return error_response(
                message="cap_width 不能为负数",
                error_type="invalid_input",
                target="cap_width",
                value=cap_width,
            )

        direction_map = {"both": 0, "plus": 1, "minus": 2}
        if direction is not None and direction.lower() not in direction_map:
            return error_response(
                message=f"不支持的方向 '{direction}'",
                error_type="invalid_input",
                target="direction",
                value=direction,
                hint="支持: both, plus, minus",
            )

        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)
                gl = gr[0]
                _get_plot(gl, plot_index)

                # 激活目标图表和曲线
                op.lt_exec(f'win -a {gr.name}')
                op.lt_exec(f'layer -s 1')
                op.lt_exec(f'layer.plot = {plot_index + 1}')

                changes = {}

                if line_width is not None:
                    op.lt_exec(f'set %C -ew {line_width}')
                    changes["line_width"] = line_width

                if cap_width is not None:
                    op.lt_exec(f'set %C -ecw {cap_width}')
                    changes["cap_width"] = cap_width

                if color is not None:
                    if color.lower() == "auto":
                        # 恢复跟随曲线颜色
                        op.lt_exec('set %C -ecc 0')
                        changes["color"] = "auto"
                    else:
                        # 解析十六进制颜色
                        c = color.lstrip('#')
                        r = int(c[0:2], 16)
                        g = int(c[2:4], 16)
                        b = int(c[4:6], 16)
                        # 设置自定义颜色
                        op.lt_exec('set %C -ecc 1')
                        op.lt_exec(
                            f'set %C -ec color(rgb({r},{g},{b}))'
                        )
                        changes["color"] = color

                if direction is not None:
                    d_code = direction_map[direction.lower()]
                    op.lt_exec(f'set %C -ed {d_code}')
                    changes["direction"] = direction

                return {
                    "graph_name": target_name,
                    "plot_index": plot_index,
                    **changes,
                }

            result = manager.execute(_set)

            return success_response(
                message=f"曲线 {plot_index} 的误差棒样式已更新。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "set_plot_color", "set_plot_line_width", "export_graph",
                ],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置误差棒样式失败: {e}",
                error_type="internal_error",
                target="error_bar_style",
                hint="请确认该曲线已添加误差棒（通过 yerr_col 参数创建）。",
            )

    # =================================================================
    # set_legend
    # =================================================================

    @mcp.tool()
    def set_legend(
        visible: bool | None = None,
        position: str | None = None,
        font_size: float | None = None,
        graph_name: str | None = None,
    ) -> dict:
        """控制图例的显示、位置和字号。

        何时使用：需要显示/隐藏图例、调整图例位置或字号时使用。
        何时不用：使用默认图例设置时无需调用。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - 未提供的参数不会被修改

        参数说明：
        - visible: 是否显示图例（True=显示, False=隐藏）
        - position: 图例位置
            - "top_left" = 左上角
            - "top_right" = 右上角
            - "bottom_left" = 左下角
            - "bottom_right" = 右下角
            - "top_center" = 顶部居中
            - "bottom_center" = 底部居中
            - "right_center" = 右侧居中
        - font_size: 图例文字大小（单位：点），常用值 8, 10, 12, 14

        示例：
        - set_legend(visible=True, position="top_right")
        - set_legend(visible=False)
        - set_legend(font_size=10, position="bottom_left")
        """
        position_map = {
            "top_left": (15, 15),
            "top_right": (85, 15),
            "bottom_left": (15, 85),
            "bottom_right": (85, 85),
            "top_center": (50, 10),
            "bottom_center": (50, 90),
            "right_center": (85, 50),
        }

        if position is not None and position.lower() not in position_map:
            return error_response(
                message=f"不支持的图例位置 '{position}'",
                error_type="invalid_input",
                target="position",
                value=position,
                hint=f"支持的位置: {list(position_map.keys())}",
            )

        if font_size is not None and font_size <= 0:
            return error_response(
                message="font_size 必须大于 0",
                error_type="invalid_input",
                target="font_size",
                value=font_size,
            )

        try:
            target_name = _resolve_graph_name(graph_name, manager)

            def _set(op: Any) -> dict[str, Any]:
                gr = _find_graph(op, target_name)

                op.lt_exec(f'win -a {gr.name}')
                op.lt_exec(f'layer -s 1')

                changes = {}

                if visible is not None:
                    if visible:
                        # 重新生成图例
                        op.lt_exec('legend -r')
                        op.lt_exec('legend -s')
                        changes["visible"] = True
                    else:
                        op.lt_exec('legend -h')
                        changes["visible"] = False

                if position is not None:
                    pos_key = position.lower()
                    x_pct, y_pct = position_map[pos_key]
                    op.lt_exec(f'legend.x = {x_pct}')
                    op.lt_exec(f'legend.y = {y_pct}')
                    changes["position"] = position

                if font_size is not None:
                    op.lt_exec(f'legend.fsize = {font_size}')
                    changes["font_size"] = font_size

                return {
                    "graph_name": target_name,
                    **changes,
                }

            result = manager.execute(_set)

            parts = []
            if visible is not None:
                parts.append("已显示" if visible else "已隐藏")
            if position is not None:
                parts.append(f"位置={position}")
            if font_size is not None:
                parts.append(f"字号={font_size}")
            desc = "，".join(parts) if parts else "已更新"

            return success_response(
                message=f"图例{desc}。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "set_plot_color", "set_graph_title", "export_graph",
                ],
            )
        except ToolError as e:
            return error_response_from_exception(e)
        except Exception as e:
            return error_response(
                message=f"设置图例失败: {e}",
                error_type="internal_error",
                target="legend",
                hint="请检查图表是否存在。",
            )
