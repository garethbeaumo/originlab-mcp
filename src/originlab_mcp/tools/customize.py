"""图表定制类 tools。

让 AI 能对图表执行高频、低歧义的样式调整。

包含:
    set_axis_range: 设置坐标轴范围
    set_axis_scale: 设置坐标轴缩放类型
    set_axis_title: 设置坐标轴标题
    set_graph_font: 设置图表字体
    set_tick_style: 设置刻度样式
    apply_publication_style: 一键应用论文风格
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

from originlab_mcp.utils.constants import (
    SCALE_TYPE_TO_ORIGIN,
    ScaleType,
)
from originlab_mcp.utils.helpers import (
    find_graph as _find_graph,
)
from originlab_mcp.utils.helpers import (
    get_graph_layer as _get_layer,
)
from originlab_mcp.utils.helpers import (
    get_plot as _get_plot,
)
from originlab_mcp.utils.helpers import (
    resolve_graph_name as _resolve_graph_name,
)
from originlab_mcp.utils.helpers import (
    sanitize_labtalk_name as _sanitize_name,
)
from originlab_mcp.utils.helpers import (
    tool_error_handler,
)
from originlab_mcp.utils.helpers import (
    validate_axis as _validate_axis,
)
from originlab_mcp.utils.validators import (
    error_response,
    success_response,
    validate_scale_type,
)

# 注: _resolve_graph_name, _find_graph, _get_plot, _validate_axis 从 utils.helpers 导入


# =====================================================================
# 公共辅助函数：LabTalk 命令封装
# =====================================================================


def _activate_plot(
    op: Any,
    graph_name: str,
    plot_index: int,
    layer_index: int = 0,
) -> None:
    """激活指定图表的第一个图层和指定曲线。

    多处 customize tool 需要通过 LabTalk 设置曲线属性，
    都需要先激活图表窗口和目标曲线，此函数消除重复代码。
    """
    safe_name = _sanitize_name(graph_name, "graph_name")
    op.lt_exec(f'win -a {safe_name}')
    op.lt_exec(f"layer -s {layer_index + 1}")
    op.lt_exec(f'layer.plot = {plot_index + 1}')


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

# 色盲友好默认配色与符号序列，用于 publication-style 一键样式
PUBLICATION_COLORS = (
    "#0072B2",  # blue
    "#D55E00",  # vermillion
    "#009E73",  # green
    "#E69F00",  # orange
    "#CC79A7",  # magenta
    "#56B4E9",  # sky blue
)
PUBLICATION_SYMBOLS = (2, 3, 1, 4, 5, 6)
LEGEND_POSITION_MAP = {
    "top_left": (15, 15),
    "top_right": (85, 15),
    "bottom_left": (15, 85),
    "bottom_right": (85, 85),
    "top_center": (50, 10),
    "bottom_center": (50, 90),
    "right_center": (85, 50),
}


def _escape_labtalk_text(value: str) -> str:
    """转义 LabTalk 字符串字面量。"""
    return value.replace("\\", "\\\\").replace('"', '\\"')


def _activate_graph_layer(op: Any, graph_name: str, layer_index: int = 0) -> None:
    """激活指定图表窗口及图层。"""
    safe_name = _sanitize_name(graph_name, "graph_name")
    op.lt_exec(f"win -a {safe_name}")
    op.lt_exec(f"layer -s {layer_index + 1}")


def _set_tick_style_commands(
    op: Any,
    graph_name: str,
    layer_index: int,
    tick_direction: str,
    major_length: int,
    minor_count: int,
    show_minor: bool,
) -> None:
    """用统一命令设置刻度方向和长度。"""
    direction_map = {"in": 1, "out": 2, "both": 3}
    _activate_graph_layer(op, graph_name, layer_index)
    minor = minor_count if show_minor else 0
    direction_value = direction_map[tick_direction]
    op.lt_exec(f"layer.x.ticks = {direction_value}")
    op.lt_exec(f"layer.y.ticks = {direction_value}")
    op.lt_exec(f"layer.x.minor = {minor}")
    op.lt_exec(f"layer.y.minor = {minor}")
    op.lt_exec(f"layer.x.majorLen = {major_length}")
    op.lt_exec(f"layer.y.majorLen = {major_length}")


def _set_graph_font_commands(
    op: Any,
    graph_name: str,
    layer_index: int,
    font_name: str,
    font_size: int,
    target: str,
    tick_font_size: int | None = None,
    legend_font_size: int | None = None,
) -> dict[str, Any]:
    """用统一命令设置图表字体。"""
    _activate_graph_layer(op, graph_name, layer_index)
    changes: dict[str, Any] = {
        "font_name": font_name,
        "font_size": font_size,
        "target": target,
    }

    if target in ("all", "axes"):
        escaped_font = _escape_labtalk_text(font_name)
        op.lt_exec(f'xb.font$ = "{escaped_font}"')
        op.lt_exec(f"xb.fsize = {font_size}")
        op.lt_exec(f'yl.font$ = "{escaped_font}"')
        op.lt_exec(f"yl.fsize = {font_size}")

    if target in ("all", "tick"):
        tick_size = tick_font_size if tick_font_size is not None else max(font_size - 4, 8)
        op.lt_exec(f"layer.x.label.pt = {tick_size}")
        op.lt_exec(f"layer.y.label.pt = {tick_size}")
        op.lt_exec("layer.x.label.bold = 1")
        op.lt_exec("layer.y.label.bold = 1")
        changes["tick_font_size"] = tick_size

    if target in ("all", "legend"):
        escaped_font = _escape_labtalk_text(font_name)
        legend_size = (
            legend_font_size
            if legend_font_size is not None
            else max(font_size - 4, 8)
        )
        op.lt_exec(f'legend.font$ = "{escaped_font}"')
        op.lt_exec(f"legend.fsize = {legend_size}")
        changes["legend_font_size"] = legend_size

    if target == "title":
        escaped_font = _escape_labtalk_text(font_name)
        op.lt_exec(f'title.font$ = "{escaped_font}"')
        op.lt_exec(f"title.fsize = {font_size}")

    return changes


def register_customize_tools(mcp: Any, manager: Any) -> None:
    """注册图表定制类 tools 到 MCP Server。

    Args:
        mcp: FastMCP 实例。
        manager: OriginManager 实例（依赖注入）。
    """
    # =================================================================
    # set_axis_range
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置轴范围", "请检查参数值。")
    def set_axis_range(
        axis: str,
        min_val: float,
        max_val: float,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the display range of a graph axis.

        When to use: To manually specify axis display range instead of auto-scaling.
        When not to use: When using auto-scaling (default).

        Default behavior:
        - graph_name omitted: uses current active graph

        Examples:
        - set_axis_range(axis="x", min_val=0, max_val=100)
        - set_axis_range(axis="y", min_val=-5, max_val=50, graph_name="Graph1")
        """
        target_name = _resolve_graph_name(graph_name, manager)
        normalized_axis = _validate_axis(axis)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            gl = _get_layer(gr, layer_index)

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

    # =================================================================
    # set_axis_scale
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置轴缩放", "请检查参数值。")
    def set_axis_scale(
        axis: str,
        scale_type: str,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the scale type of a graph axis.

        When to use: To switch axis to logarithmic scale (log/ln) or back to linear.
        When not to use: When using linear scale (default).

        Default behavior:
        - graph_name omitted: uses current active graph

        Examples:
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

        target_name = _resolve_graph_name(graph_name, manager)
        normalized_axis = _validate_axis(axis)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            gl = _get_layer(gr, layer_index)
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

    # =================================================================
    # set_axis_title
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置轴标题", "请检查参数值。")
    def set_axis_title(
        axis: str,
        title: str,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the title of a graph axis.

        When to use: To modify or set X/Y axis text titles.
        When not to use: If axis titles are already correct.

        Default behavior:
        - graph_name omitted: uses current active graph

        Examples:
        - set_axis_title(axis="x", title="Time (s)")
        - set_axis_title(axis="y", title="Voltage (mV)", graph_name="Graph1")
        """
        target_name = _resolve_graph_name(graph_name, manager)
        normalized_axis = _validate_axis(axis, ("x", "y", "z"))

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            gl = _get_layer(gr, layer_index)
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

    # =================================================================
    # set_plot_color
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置颜色", "请使用十六进制颜色值，如 '#ff5833'。")
    def set_plot_color(
        color: str,
        plot_index: int = 0,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the color of a specified curve.

        When to use: To change the color of a curve in the graph.
        When not to use: When using default colors.

        Default behavior:
        - graph_name omitted: uses current active graph
        - plot_index defaults to 0 (first curve)

        Parameter notes:
        - color: hex color value, e.g. "#ff5833", "#0000ff"

        Examples:
        - set_plot_color(color="#ff0000")
        - set_plot_color(color="#0000ff", plot_index=1, graph_name="Graph1")
        """
        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            plot = _get_plot(_get_layer(gr, layer_index), plot_index)
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

    # =================================================================
    # set_plot_colormap
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置颜色映射", "请检查 colormap 名称是否正确。")
    def set_plot_colormap(
        colormap: str,
        plot_index: int = 0,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the colormap for a specified curve.

        When to use: To apply gradient colors or heatmap color schemes to data points.
        When not to use: For a single solid color, use set_plot_color.

        Default behavior:
        - graph_name omitted: uses current active graph
        - plot_index defaults to 0 (first curve)

        Examples:
        - set_plot_colormap(colormap="Candy")
        - set_plot_colormap(colormap="Rainbow", plot_index=0)
        """
        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            plot = _get_plot(_get_layer(gr, layer_index), plot_index)
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

    # =================================================================
    # set_plot_symbols
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置符号", "请检查 shape_list 格式。")
    def set_plot_symbols(
        shape_list: list[int],
        plot_index: int = 0,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the data point symbol shapes for a specified curve.

        When to use: To modify symbol shapes in scatter or line+symbol plots.
        When not to use: For pure line plots or when default symbols are fine.

        Default behavior:
        - graph_name omitted: uses current active graph
        - plot_index defaults to 0 (first curve)

        Parameter notes:
        - shape_list: list of symbol shape numbers (Origin built-in numbers)

        Examples:
        - set_plot_symbols(shape_list=[3, 2, 1])
        """
        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            plot = _get_plot(_get_layer(gr, layer_index), plot_index)
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

    # =================================================================
    # set_plot_transparency
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置透明度", "请检查参数值。")
    def set_plot_transparency(
        transparency: int,
        plot_index: int = 0,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the transparency of a specified curve.

        When to use: To adjust curve or data point transparency, commonly used for overlapping multi-curve scenarios.
        When not to use: When no transparency effect is needed.

        Default behavior:
        - graph_name omitted: uses current active graph
        - plot_index defaults to 0 (first curve)

        Parameter notes:
        - transparency: percentage, 0=fully opaque, 100=fully transparent

        Examples:
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

        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            plot = _get_plot(_get_layer(gr, layer_index), plot_index)
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

    # =================================================================
    # set_symbol_size
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置符号大小", "请检查参数值。")
    def set_symbol_size(
        size: float,
        plot_index: int = 0,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the data point symbol size for a specified curve.

        When to use: To adjust data point size in scatter or line+symbol plots.
        When not to use: When using default size.

        Default behavior:
        - graph_name omitted: uses current active graph
        - plot_index defaults to 0 (first curve)

        Examples:
        - set_symbol_size(size=12)
        - set_symbol_size(size=8.5, plot_index=1)
        """
        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            plot = _get_plot(_get_layer(gr, layer_index), plot_index)
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

    # =================================================================
    # set_fill_area
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置填充区域", "请检查参数值。此功能仅适用于折线图。")
    def set_fill_area(
        above_color: int,
        fill_type: int = 9,
        below_color: int | None = None,
        plot_index: int = 0,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the fill area under a line plot curve.

        When to use: To fill color under a line plot curve (e.g. showing area, confidence intervals).
        When not to use: For scatter plots or when no fill effect is needed.

        Default behavior:
        - graph_name omitted: uses current active graph
        - plot_index defaults to 0 (first curve)
        - fill_type defaults to 9 (fill to next curve)

        Parameter notes:
        - above_color: fill color number for above region
        - fill_type: fill option (9=fill to next curve)
        - below_color: fill color number for below region (optional)

        Examples:
        - set_fill_area(above_color=2)
        - set_fill_area(above_color=2, fill_type=9, below_color=3)
        """
        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            plot = _get_plot(_get_layer(gr, layer_index), plot_index)
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

    # =================================================================
    # get_graph_info
    # =================================================================

    @mcp.tool()
    @tool_error_handler("获取图表信息", "请检查图表名称。调用 list_graphs 查看可用图表。")
    def get_graph_info(
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Get detailed graph info including layers, curve list, and data sources.

        When to use: To understand which curves exist, their data sources and style info.
        When not to use: To only see the graph list, use list_graphs.

        Default behavior:
        - graph_name omitted: uses current active graph

        Examples:
        - get_graph_info()
        - get_graph_info(graph_name="Graph1")
        """
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

        total_plots = sum(
            layer_info["plot_count"] for layer_info in result["layers"]
        )

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

    # =================================================================
    # add_text_label
    # =================================================================

    @mcp.tool()
    @tool_error_handler("添加文字标注", "请检查图表是否存在。")
    def add_text_label(
        text: str,
        x: float | None = None,
        y: float | None = None,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Add a text annotation to a graph.

        When to use: To annotate peaks, inflection points, or add explanatory text on a graph.
        When not to use: To set axis titles, use set_axis_title.

        Default behavior:
        - graph_name omitted: uses current active graph
        - x, y omitted: Origin auto-places the label

        Parameter notes:
        - text: annotation text content
        - x, y: position coordinates (layer coordinate system)

        Examples:
        - add_text_label(text="Peak A")
        - add_text_label(text="Transition Point", x=3.5, y=100)
        """
        target_name = _resolve_graph_name(graph_name, manager)

        def _add(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            gl = _get_layer(gr, layer_index)
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

    # =================================================================
    # add_line_to_graph
    # =================================================================

    @mcp.tool()
    @tool_error_handler("添加线条", "请检查坐标值和图表是否存在。")
    def add_line_to_graph(
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        line_width: float = 1.0,
        arrow: bool = False,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Add a line or arrow to a graph.

        When to use: To draw reference lines, indicator lines, or arrows on a graph.
        When not to use: To set axis ranges, use set_axis_range.

        Default behavior:
        - graph_name omitted: uses current active graph
        - arrow=True adds an arrowhead at the endpoint

        Parameter notes:
        - x1, y1: start point coordinates
        - x2, y2: end point coordinates
        - line_width: line width
        - arrow: whether to add arrowhead at endpoint

        Examples:
        - add_line_to_graph(x1=0, y1=0, x2=10, y2=10)
        - add_line_to_graph(x1=5, y1=0, x2=5, y2=100, arrow=True)
        """
        target_name = _resolve_graph_name(graph_name, manager)

        def _add(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            gl = _get_layer(gr, layer_index)
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

    # =================================================================
    # remove_graph_label
    # =================================================================

    @mcp.tool()
    @tool_error_handler("移除标签", "请检查标签名称是否正确。")
    def remove_graph_label(
        label_name: str,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Remove a label/annotation object from a graph.

        When to use: To delete previously added text labels or other label objects.
        When not to use: If unsure about label names, call get_graph_info first.

        Default behavior:
        - graph_name omitted: uses current active graph

        Parameter notes:
        - label_name: name of the label to remove (e.g. "xb" for X axis title label)

        Examples:
        - remove_graph_label(label_name="Text1")
        - remove_graph_label(label_name="xb")
        """
        target_name = _resolve_graph_name(graph_name, manager)

        def _remove(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            gl = _get_layer(gr, layer_index)
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

    # =================================================================
    # set_graph_title
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置图表标题", "请检查图表是否存在。")
    def set_graph_title(
        title: str,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the main title of a graph (layer title).

        When to use: To set or modify the main title of a graph.
        When not to use: To set axis titles, use set_axis_title.

        Default behavior:
        - graph_name omitted: uses current active graph

        Examples:
        - set_graph_title(title="Temperature vs Time")
        - set_graph_title(title="Fig. 1", graph_name="Graph1")
        """
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

    # =================================================================
    # set_axis_step
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置轴步长", "请检查步长值是否合理。")
    def set_axis_step(
        axis: str,
        step: float,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the tick step (increment) for a graph axis.

        When to use: To precisely control axis tick spacing.
        When not to use: To only set range, use set_axis_range.

        Default behavior:
        - graph_name omitted: uses current active graph

        Parameter notes:
        - axis: axis identifier, "x" or "y"
        - step: tick step value

        Examples:
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

        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            gl = _get_layer(gr, layer_index)

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

    # =================================================================
    # set_symbol_interior
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置符号填充", "请检查参数值。")
    def set_symbol_interior(
        interior: int,
        plot_index: int = 0,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the interior fill type of data point symbols.

        When to use: To distinguish multiple curves (e.g. solid vs hollow symbols).
        When not to use: When using default fill.

        Default behavior:
        - graph_name omitted: uses current active graph
        - plot_index defaults to 0 (first curve)

        Parameter notes:
        - interior: fill type number
            - 0 = no symbol
            - 1 = solid
            - 2 = hollow
            - 3 = dot center

        Examples:
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

        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            plot = _get_plot(_get_layer(gr, layer_index), plot_index)
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

    # =================================================================
    # set_color_increment
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置颜色增量", "请检查参数值。")
    def set_color_increment(
        increment: int,
        plot_index: int = 0,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the color increment for grouped curves.

        When to use: When grouped curves need auto-incrementing colors.
        When not to use: To manually set each curve's color, use set_plot_color.

        Default behavior:
        - graph_name omitted: uses current active graph
        - plot_index defaults to 0 (usually set on the group leader curve)

        Parameter notes:
        - increment: color increment value (1=each curve increments by one color)

        Examples:
        - set_color_increment(increment=1)
        """
        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            plot = _get_plot(_get_layer(gr, layer_index), plot_index)
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

    # =================================================================
    # set_symbol_increment
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置符号增量", "请检查参数值。")
    def set_symbol_increment(
        increment: int,
        plot_index: int = 0,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the symbol shape increment for grouped curves.

        When to use: When grouped curves need auto-incrementing symbol shapes.
        When not to use: To manually set each curve's symbols, use set_plot_symbols.

        Default behavior:
        - graph_name omitted: uses current active graph
        - plot_index defaults to 0 (usually set on the group leader curve)

        Parameter notes:
        - increment: symbol shape increment value (1=each curve increments by one shape)

        Examples:
        - set_symbol_increment(increment=1)
        """
        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            plot = _get_plot(_get_layer(gr, layer_index), plot_index)
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

    # =================================================================
    # set_plot_line_width
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置线宽", "请检查参数值。")
    def set_plot_line_width(
        width: float,
        plot_index: int = 0,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the line width of a specified curve.

        When to use: To adjust line thickness in line or line+symbol plots.
        When not to use: When using default line width.

        Default behavior:
        - graph_name omitted: uses current active graph
        - plot_index defaults to 0 (first curve)

        Parameter notes:
        - width: line width in points, common values: 0.5, 1, 1.5, 2, 3

        Examples:
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

        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            gl = _get_layer(gr, layer_index)
            # 验证 plot_index 有效
            _get_plot(gl, plot_index)

            # 通过 LabTalk 设置线宽
            _activate_plot(op, gr.name, plot_index, layer_index)
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

    # =================================================================
    # set_plot_line_style
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置线型", "请检查线型名称。")
    def set_plot_line_style(
        style: str,
        plot_index: int = 0,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set the line style of a specified curve (solid, dash, dot, etc.).

        When to use: To differentiate multiple curves or meet publication formatting requirements.
        When not to use: When using default solid line.

        Default behavior:
        - graph_name omitted: uses current active graph
        - plot_index defaults to 0 (first curve)

        Parameter notes:
        - style: line style name
            - "solid" = solid line
            - "dash" = dashed line
            - "dot" = dotted line
            - "dashdot" = dash-dot line
            - "dashdotdot" = dash-dot-dot line
            - "short_dash" = short dashed line
            - "short_dot" = short dotted line
            - "short_dashdot" = short dash-dot line

        Examples:
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

        target_name = _resolve_graph_name(graph_name, manager)
        style_code = LINE_STYLE_MAP[style_lower]

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            gl = _get_layer(gr, layer_index)
            _get_plot(gl, plot_index)

            _activate_plot(op, gr.name, plot_index, layer_index)
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

    # =================================================================
    # set_error_bar_style
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置误差棒样式", "请确认该曲线已添加误差棒（通过 yerr_col 参数创建）。")
    def set_error_bar_style(
        plot_index: int = 0,
        line_width: float | None = None,
        cap_width: float | None = None,
        color: str | None = None,
        direction: str | None = None,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set error bar style (line width, cap, color, direction).

        When to use: To adjust error bar appearance (e.g. for publication formatting, visual distinction).
        When not to use: When using default error bar style. Requires error bars created via yerr_col parameter in create_plot or add_plot_to_graph.

        Default behavior:
        - graph_name omitted: uses current active graph
        - plot_index defaults to 0 (first curve)
        - Parameters not provided will not be modified

        Parameter notes:
        - line_width: error bar line width (points), common values 1, 1.5, 2
        - cap_width: error bar cap width (points), common values 5, 8, 10, 15
        - color: hex color value like "#ff0000"; set to "auto" to follow curve color
        - direction: error bar direction
            - "both" = both directions (default)
            - "plus" = upward only
            - "minus" = downward only

        Examples:
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

        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            gl = _get_layer(gr, layer_index)
            _get_plot(gl, plot_index)

            # 激活目标图表和曲线
            _activate_plot(op, gr.name, plot_index, layer_index)

            changes: dict[str, Any] = {}

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

    # =================================================================
    # set_legend
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置图例", "请检查图表是否存在。")
    def set_legend(
        visible: bool | None = None,
        position: str | None = None,
        font_size: float | None = None,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Control legend visibility, position, and font size.

        When to use: To show/hide legend, adjust legend position, or change font size.
        When not to use: When using default legend settings.

        Default behavior:
        - graph_name omitted: uses current active graph
        - Parameters not provided will not be modified

        Parameter notes:
        - visible: whether to show legend (True=show, False=hide)
        - position: legend position
            - "top_left" = top left corner
            - "top_right" = top right corner
            - "bottom_left" = bottom left corner
            - "bottom_right" = bottom right corner
            - "top_center" = top center
            - "bottom_center" = bottom center
            - "right_center" = right center
        - font_size: legend text size (points), common values 8, 10, 12, 14

        Examples:
        - set_legend(visible=True, position="top_right")
        - set_legend(visible=False)
        - set_legend(font_size=10, position="bottom_left")
        """
        position_map = LEGEND_POSITION_MAP

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

        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            _get_layer(gr, layer_index)

            _activate_graph_layer(op, gr.name, layer_index)

            changes: dict[str, Any] = {}

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

    # =================================================================
    # set_graph_font
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置图表字体", "请检查字体名称和字号。")
    def set_graph_font(
        font_name: str = "Arial",
        font_size: int = 24,
        target: str = "all",
        tick_font_size: int | None = None,
        legend_font_size: int | None = None,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set font family and size for graph elements.

        When to use: To unify graph typography for axes, ticks, title, or legend.
        When not to use: For single axis titles only, use set_axis_title.

        Default behavior:
        - graph_name omitted: uses current active graph
        - target defaults to "all"

        Parameter notes:
        - target: "all", "axes", "title", "legend", or "tick"
        - tick_font_size: overrides derived tick label size
        - legend_font_size: overrides derived legend size

        Examples:
        - set_graph_font(font_name="Arial", font_size=24)
        - set_graph_font(font_name="Times New Roman", font_size=18, target="legend")
        """
        if not font_name.strip():
            return error_response(
                message="font_name 不能为空",
                error_type="invalid_input",
                target="font_name",
                value=font_name,
            )
        if font_size <= 0:
            return error_response(
                message="font_size 必须大于 0",
                error_type="invalid_input",
                target="font_size",
                value=font_size,
            )
        if tick_font_size is not None and tick_font_size <= 0:
            return error_response(
                message="tick_font_size 必须大于 0",
                error_type="invalid_input",
                target="tick_font_size",
                value=tick_font_size,
            )
        if legend_font_size is not None and legend_font_size <= 0:
            return error_response(
                message="legend_font_size 必须大于 0",
                error_type="invalid_input",
                target="legend_font_size",
                value=legend_font_size,
            )

        normalized_target = target.lower()
        valid_targets = {"all", "axes", "title", "legend", "tick"}
        if normalized_target not in valid_targets:
            return error_response(
                message=f"不支持的 target '{target}'",
                error_type="invalid_input",
                target="target",
                value=target,
                hint=f"支持的取值: {sorted(valid_targets)}",
            )

        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            _get_layer(gr, layer_index)
            result = _set_graph_font_commands(
                op=op,
                graph_name=gr.name,
                layer_index=layer_index,
                font_name=font_name,
                font_size=font_size,
                target=normalized_target,
                tick_font_size=tick_font_size,
                legend_font_size=legend_font_size,
            )
            result["graph_name"] = target_name
            return result

        result = manager.execute(_set)

        return success_response(
            message=(
                f"图表字体已更新：{font_name} {font_size}pt "
                f"（target={normalized_target}）。"
            ),
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "set_tick_style",
                "apply_publication_style",
                "export_graph",
            ],
        )

    # =================================================================
    # set_tick_style
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置刻度样式", "请检查刻度参数。")
    def set_tick_style(
        tick_direction: str = "in",
        major_length: int = 8,
        minor_count: int = 4,
        show_minor: bool = True,
        graph_name: str | None = None,
        layer_index: int = 0,
    ) -> dict:
        """Set tick mark direction, length, and minor tick count.

        When to use: To standardize graph ticks for presentation or publication.
        When not to use: When the default tick style is acceptable.

        Default behavior:
        - graph_name omitted: uses current active graph
        - tick_direction defaults to "in"
        - show_minor defaults to true

        Examples:
        - set_tick_style()
        - set_tick_style(tick_direction="both", major_length=10, minor_count=2)
        """
        normalized_direction = tick_direction.lower()
        if normalized_direction not in {"in", "out", "both"}:
            return error_response(
                message=f"不支持的 tick_direction '{tick_direction}'",
                error_type="invalid_input",
                target="tick_direction",
                value=tick_direction,
                hint="支持的方向: in, out, both",
            )
        if major_length <= 0:
            return error_response(
                message="major_length 必须大于 0",
                error_type="invalid_input",
                target="major_length",
                value=major_length,
            )
        if minor_count < 0:
            return error_response(
                message="minor_count 不能为负数",
                error_type="invalid_input",
                target="minor_count",
                value=minor_count,
            )

        target_name = _resolve_graph_name(graph_name, manager)

        def _set(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            _get_layer(gr, layer_index)
            _set_tick_style_commands(
                op=op,
                graph_name=gr.name,
                layer_index=layer_index,
                tick_direction=normalized_direction,
                major_length=major_length,
                minor_count=minor_count,
                show_minor=show_minor,
            )
            return {
                "graph_name": target_name,
                "tick_direction": normalized_direction,
                "major_length": major_length,
                "minor_count": minor_count if show_minor else 0,
                "show_minor": show_minor,
            }

        result = manager.execute(_set)

        return success_response(
            message=(
                f"刻度样式已更新：方向={normalized_direction}，"
                f"主刻度长度={major_length}。"
            ),
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "set_graph_font",
                "apply_publication_style",
                "export_graph",
            ],
        )

    # =================================================================
    # apply_publication_style
    # =================================================================

    @mcp.tool()
    @tool_error_handler("应用论文风格", "请检查图表名称和样式参数。")
    def apply_publication_style(
        graph_name: str | None = None,
        x_label: str = "",
        y_label: str = "",
        x_min: float | None = None,
        x_max: float | None = None,
        y_min: float | None = None,
        y_max: float | None = None,
        legend_position: str = "top_right",
        legend_visible: bool = True,
        font_name: str = "Arial",
        layer_index: int = 0,
        axis_title_size: int = 28,
        tick_font_size: int = 20,
        legend_font_size: int = 20,
        line_width: float = 2,
        symbol_size: float = 10,
        tick_direction: str = "in",
        major_length: int = 8,
        minor_count: int = 4,
        show_minor: bool = True,
    ) -> dict:
        """Apply a publication-ready graph style in one call.

        When to use: To quickly standardize a graph for papers or reports.
        When not to use: When you need fine-grained control over every style detail.

        Default behavior:
        - graph_name omitted: uses current active graph
        - legend defaults to visible at top_right
        - font defaults to Arial
        - layer_index defaults to 0 (main layer)

        Examples:
        - apply_publication_style(x_label="Time (s)", y_label="Intensity (a.u.)")
        - apply_publication_style(graph_name="Graph1", x_min=0, x_max=10, y_min=0)
        - apply_publication_style(layer_index=1, line_width=3, symbol_size=12)
        """
        normalized_position = legend_position.lower()
        if normalized_position not in LEGEND_POSITION_MAP:
            return error_response(
                message=f"不支持的图例位置 '{legend_position}'",
                error_type="invalid_input",
                target="legend_position",
                value=legend_position,
                hint=f"支持的位置: {list(LEGEND_POSITION_MAP.keys())}",
            )
        if not font_name.strip():
            return error_response(
                message="font_name 不能为空",
                error_type="invalid_input",
                target="font_name",
                value=font_name,
            )
        normalized_direction = tick_direction.lower()
        if normalized_direction not in {"in", "out", "both"}:
            return error_response(
                message=f"不支持的 tick_direction '{tick_direction}'",
                error_type="invalid_input",
                target="tick_direction",
                value=tick_direction,
                hint="支持的方向: in, out, both",
            )
        if axis_title_size <= 0:
            return error_response(
                message="axis_title_size 必须大于 0",
                error_type="invalid_input",
                target="axis_title_size",
                value=axis_title_size,
            )
        if tick_font_size <= 0:
            return error_response(
                message="tick_font_size 必须大于 0",
                error_type="invalid_input",
                target="tick_font_size",
                value=tick_font_size,
            )
        if legend_font_size <= 0:
            return error_response(
                message="legend_font_size 必须大于 0",
                error_type="invalid_input",
                target="legend_font_size",
                value=legend_font_size,
            )
        if line_width <= 0:
            return error_response(
                message="line_width 必须大于 0",
                error_type="invalid_input",
                target="line_width",
                value=line_width,
            )
        if symbol_size <= 0:
            return error_response(
                message="symbol_size 必须大于 0",
                error_type="invalid_input",
                target="symbol_size",
                value=symbol_size,
            )
        if major_length <= 0:
            return error_response(
                message="major_length 必须大于 0",
                error_type="invalid_input",
                target="major_length",
                value=major_length,
            )
        if minor_count < 0:
            return error_response(
                message="minor_count 不能为负数",
                error_type="invalid_input",
                target="minor_count",
                value=minor_count,
            )

        target_name = _resolve_graph_name(graph_name, manager)

        def _apply(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_name)
            gl = _get_layer(gr, layer_index)
            plots = gl.plot_list()

            _activate_graph_layer(op, gr.name, layer_index)

            if x_label:
                escaped = _escape_labtalk_text(x_label)
                escaped_font = _escape_labtalk_text(font_name)
                op.lt_exec(f'xb.text$ = "\\b({escaped})"')
                op.lt_exec(f'xb.font$ = "{escaped_font}"')
                op.lt_exec(f"xb.fsize = {axis_title_size}")
            if y_label:
                escaped = _escape_labtalk_text(y_label)
                escaped_font = _escape_labtalk_text(font_name)
                op.lt_exec(f'yl.text$ = "\\b({escaped})"')
                op.lt_exec(f'yl.font$ = "{escaped_font}"')
                op.lt_exec(f"yl.fsize = {axis_title_size}")

            _set_graph_font_commands(
                op=op,
                graph_name=gr.name,
                layer_index=layer_index,
                font_name=font_name,
                font_size=max(axis_title_size, tick_font_size, legend_font_size),
                target="tick",
                tick_font_size=tick_font_size,
            )
            _set_tick_style_commands(
                op=op,
                graph_name=gr.name,
                layer_index=layer_index,
                tick_direction=normalized_direction,
                major_length=major_length,
                minor_count=minor_count,
                show_minor=show_minor,
            )

            if x_min is not None:
                op.lt_exec(f"layer.x.from = {x_min}")
            if x_max is not None:
                op.lt_exec(f"layer.x.to = {x_max}")
            if y_min is not None:
                op.lt_exec(f"layer.y.from = {y_min}")
            if y_max is not None:
                op.lt_exec(f"layer.y.to = {y_max}")

            op.lt_exec("layer.x.opposite = 1")
            op.lt_exec("layer.y.opposite = 1")
            op.lt_exec("layer.x.thickness = 2")
            op.lt_exec("layer.y.thickness = 2")
            op.lt_exec("layer.x.grid = 0")
            op.lt_exec("layer.y.grid = 0")
            op.lt_exec("layer.x.minorGrid = 0")
            op.lt_exec("layer.y.minorGrid = 0")

            styled_plots: list[dict[str, Any]] = []
            for plot_index, plot in enumerate(plots):
                color = PUBLICATION_COLORS[plot_index % len(PUBLICATION_COLORS)]
                symbol = PUBLICATION_SYMBOLS[plot_index % len(PUBLICATION_SYMBOLS)]

                if hasattr(plot, "color"):
                    plot.color = color
                if hasattr(plot, "symbol_size"):
                    plot.symbol_size = symbol_size
                if hasattr(plot, "shapelist"):
                    plot.shapelist = [symbol]

                _activate_plot(op, gr.name, plot_index, layer_index)
                op.lt_exec(f"set %C -w {line_width}")
                op.lt_exec("set %C -d 1")
                op.lt_exec(f"set %C -k {symbol}")
                op.lt_exec(f"set %C -z {symbol_size}")

                styled_plots.append(
                    {
                        "plot_index": plot_index,
                        "color": color,
                        "symbol": symbol,
                        "line_width": line_width,
                        "symbol_size": symbol_size,
                    }
                )

            _activate_graph_layer(op, gr.name, layer_index)
            if legend_visible:
                op.lt_exec("legend -r")
                op.lt_exec("legend -s")
                x_pct, y_pct = LEGEND_POSITION_MAP[normalized_position]
                op.lt_exec(f"legend.x = {x_pct}")
                op.lt_exec(f"legend.y = {y_pct}")
                escaped_font = _escape_labtalk_text(font_name)
                op.lt_exec(f'legend.font$ = "{escaped_font}"')
                op.lt_exec(f"legend.fsize = {legend_font_size}")
            else:
                op.lt_exec("legend -h")

            return {
                "graph_name": target_name,
                "layer_index": layer_index,
                "x_label": x_label,
                "y_label": y_label,
                "x_range": {"min": x_min, "max": x_max},
                "y_range": {"min": y_min, "max": y_max},
                "legend_visible": legend_visible,
                "legend_position": normalized_position if legend_visible else None,
                "font_name": font_name,
                "axis_title_size": axis_title_size,
                "tick_font_size": tick_font_size,
                "legend_font_size": legend_font_size,
                "tick_direction": normalized_direction,
                "major_length": major_length,
                "minor_count": minor_count if show_minor else 0,
                "show_minor": show_minor,
                "styled_plots": styled_plots,
            }

        result = manager.execute(_apply)

        return success_response(
            message=(
                f"已对图表 '{target_name}' 应用论文风格："
                f"{len(result['styled_plots'])} 条曲线已统一样式。"
            ),
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "export_graph",
                "export_all_graphs",
                "save_project",
            ],
        )
