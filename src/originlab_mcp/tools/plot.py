"""绘图类 tools。

让 AI 能把结构化数据稳定转换成图表。

包含:
    create_plot: 创建新图表
    add_plot_to_graph: 在已有图表上追加曲线
    create_double_y_plot: 创建双 Y 轴图
    list_graphs: 列出所有图表
    list_graph_templates: 列出支持的图表模板
"""

from __future__ import annotations

from contextlib import suppress
from typing import Any

from originlab_mcp.utils.constants import (
    DEFAULT_PLOT_TYPE,
    PLOT_TYPE_TO_TEMPLATE,
    PlotType,
)
from originlab_mcp.utils.helpers import (
    find_graph as _find_graph,
)
from originlab_mcp.utils.helpers import (
    find_worksheet as _find_worksheet,
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
    resolve_worksheet_name as _resolve_worksheet_name,
)
from originlab_mcp.utils.helpers import (
    tool_error_handler,
)
from originlab_mcp.utils.helpers import (
    validate_column_indices as _validate_cols,
)
from originlab_mcp.utils.validators import (
    error_response,
    normalize_y_cols,
    success_response,
    validate_plot_type,
)


def register_plot_tools(mcp: Any, manager: Any) -> None:
    """注册绘图类 tools 到 MCP Server。

    Args:
        mcp: FastMCP 实例。
        manager: OriginManager 实例（依赖注入）。
    """

    # =================================================================
    # create_plot
    # =================================================================

    @mcp.tool()
    @tool_error_handler("创建图表", "请检查工作表数据和列索引是否正确。")
    def create_plot(
        x_col: int,
        y_cols: int | list[int],
        sheet_name: str | None = None,
        plot_type: str = DEFAULT_PLOT_TYPE.value,
        yerr_col: int | None = None,
        xerr_col: int | None = None,
    ) -> dict:
        """Create a new graph (supports single or multiple curves).

        When to use: To create a graph from scratch.
        When not to use: To add curves to an existing graph,
        use add_plot_to_graph.

        Default behavior:
        - sheet_name omitted: uses current active worksheet
        - plot_type omitted: defaults to line
        - yerr_col omitted: no Y error bars
        - xerr_col omitted: no X error bars

        Examples:
        - create_plot(x_col=0, y_cols=1)
        - create_plot(x_col=0, y_cols=[1, 2], plot_type="scatter")
        - create_plot(x_col=0, y_cols=1, yerr_col=2)
        - create_plot(x_col=0, y_cols=1, yerr_col=2, xerr_col=3)
        """
        err = validate_plot_type(plot_type)
        if err:
            return error_response(
                message=err,
                error_type="unsupported",
                target="plot_type",
                value=plot_type,
                hint=f"支持的类型: {[e.value for e in PlotType]}",
            )

        target_name = _resolve_worksheet_name(sheet_name, manager)

        try:
            y_col_list = normalize_y_cols(y_cols)
        except ValueError as e:
            return error_response(
                message=str(e),
                error_type="invalid_input",
                target="y_cols",
                value=y_cols,
                hint="请传入单个整数列索引，或整数列表。",
            )

        def _plot(op: Any) -> dict[str, Any]:
            wks = _find_worksheet(op, target_name)
            total_cols = wks.cols

            cols_to_check = [x_col] + y_col_list
            if yerr_col is not None:
                cols_to_check.append(yerr_col)
            if xerr_col is not None:
                cols_to_check.append(xerr_col)
            _validate_cols(cols_to_check, total_cols)

            template = PLOT_TYPE_TO_TEMPLATE.get(plot_type, "line")
            gr = op.new_graph(template=template)
            gl = _get_layer(gr, 0)

            curves = []
            for i, yc in enumerate(y_col_list):
                plot_kwargs = {"coly": yc, "colx": x_col}
                if yerr_col is not None:
                    plot_kwargs["colyerr"] = yerr_col
                if xerr_col is not None:
                    plot_kwargs["colxerr"] = xerr_col
                gl.add_plot(wks, **plot_kwargs)
                curve_info = {
                    "y_col": yc, "x_col": x_col, "plot_index": i,
                }
                if yerr_col is not None:
                    curve_info["yerr_col"] = yerr_col
                if xerr_col is not None:
                    curve_info["xerr_col"] = xerr_col
                curves.append(curve_info)

            gl.rescale()
            with suppress(Exception):
                gl.group()

            graph_name = gr.name
            manager.active_graph = graph_name

            return {
                "graph_name": graph_name,
                "plot_type": plot_type,
                "sheet_name": target_name,
                "curves": curves,
                "curve_count": len(curves),
            }

        result = manager.execute(_plot)

        return success_response(
            message=(
                f"图表 '{result['graph_name']}' 已创建，"
                f"共 {result['curve_count']} 条曲线。"
            ),
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "set_axis_title",
                "set_plot_color",
                "add_plot_to_graph",
                "export_graph",
            ],
        )

    # =================================================================
    # add_plot_to_graph
    # =================================================================

    @mcp.tool()
    @tool_error_handler("追加曲线", "请检查图表和工作表是否存在，以及列索引是否正确。")
    def add_plot_to_graph(
        x_col: int,
        y_cols: int | list[int],
        graph_name: str | None = None,
        sheet_name: str | None = None,
        yerr_col: int | None = None,
        xerr_col: int | None = None,
    ) -> dict:
        """Add one or more curves to an existing graph.

        When to use: To overlay new data curves on an existing graph.
        When not to use: To create a new graph, use create_plot.

        Default behavior:
        - graph_name omitted: uses current active graph
        - sheet_name omitted: uses current active worksheet
        - yerr_col omitted: no Y error bars
        - xerr_col omitted: no X error bars

        Examples:
        - add_plot_to_graph(x_col=0, y_cols=2)
        - add_plot_to_graph(x_col=0, y_cols=[3, 4], graph_name="Graph1")
        - add_plot_to_graph(x_col=0, y_cols=1, yerr_col=2)
        - add_plot_to_graph(x_col=0, y_cols=1, yerr_col=2, xerr_col=3)
        """
        target_graph = _resolve_graph_name(graph_name, manager)
        target_sheet = _resolve_worksheet_name(sheet_name, manager)

        try:
            y_col_list = normalize_y_cols(y_cols)
        except ValueError as e:
            return error_response(
                message=str(e),
                error_type="invalid_input",
                target="y_cols",
                value=y_cols,
                hint="请传入单个整数列索引，或整数列表。",
            )

        def _add(op: Any) -> dict[str, Any]:
            wks = _find_worksheet(op, target_sheet)
            gr = _find_graph(op, target_graph)

            cols_to_check = [x_col] + y_col_list
            if yerr_col is not None:
                cols_to_check.append(yerr_col)
            if xerr_col is not None:
                cols_to_check.append(xerr_col)
            _validate_cols(cols_to_check, wks.cols)

            gl = _get_layer(gr, 0)
            existing_count = (
                gl.num_plots if hasattr(gl, "num_plots") else 0
            )

            new_curves = []
            for i, yc in enumerate(y_col_list):
                plot_kwargs = {"coly": yc, "colx": x_col}
                if yerr_col is not None:
                    plot_kwargs["colyerr"] = yerr_col
                if xerr_col is not None:
                    plot_kwargs["colxerr"] = xerr_col
                gl.add_plot(wks, **plot_kwargs)
                curve_info = {
                    "y_col": yc,
                    "x_col": x_col,
                    "plot_index": existing_count + i,
                }
                if yerr_col is not None:
                    curve_info["yerr_col"] = yerr_col
                if xerr_col is not None:
                    curve_info["xerr_col"] = xerr_col
                new_curves.append(curve_info)

            gl.rescale()
            with suppress(Exception):
                gl.group()

            return {
                "graph_name": target_graph,
                "sheet_name": target_sheet,
                "new_curves": new_curves,
                "new_curve_count": len(new_curves),
                "total_curve_count": existing_count + len(y_col_list),
            }

        result = manager.execute(_add)

        return success_response(
            message=(
                f"已向图表 '{target_graph}' 追加 "
                f"{result['new_curve_count']} 条曲线，"
                f"当前共 {result['total_curve_count']} 条。"
            ),
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "set_axis_title",
                "set_plot_color",
                "export_graph",
            ],
        )

    # =================================================================
    # create_double_y_plot
    # =================================================================

    @mcp.tool()
    @tool_error_handler("创建双Y轴图", "请检查列索引是否正确。")
    def create_double_y_plot(
        x_col: int,
        y1_col: int,
        y2_col: int,
        sheet_name: str | None = None,
    ) -> dict:
        """Create a double-Y-axis graph.

        When to use: To display data with different scales or units using
        left and right Y axes in a single graph.
        When not to use: If all Y columns have similar scales,
        use create_plot instead.

        Default behavior:
        - sheet_name omitted: uses current active worksheet
        - y1_col binds to left Y axis, y2_col binds to right Y axis

        Examples:
        - create_double_y_plot(x_col=0, y1_col=1, y2_col=2)
        """
        target_name = _resolve_worksheet_name(sheet_name, manager)

        def _dbl(op: Any) -> dict[str, Any]:
            wks = _find_worksheet(op, target_name)
            _validate_cols([x_col, y1_col, y2_col], wks.cols)

            gr = op.new_graph(template="doubley")
            _get_layer(gr, 0).add_plot(wks, coly=y1_col, colx=x_col)
            gr[1].add_plot(wks, coly=y2_col, colx=x_col)
            _get_layer(gr, 0).rescale()
            gr[1].rescale()

            graph_name = gr.name
            manager.active_graph = graph_name

            return {
                "graph_name": graph_name,
                "sheet_name": target_name,
                "left_y": {"y_col": y1_col, "x_col": x_col},
                "right_y": {"y_col": y2_col, "x_col": x_col},
            }

        result = manager.execute(_dbl)

        return success_response(
            message=f"双 Y 轴图 '{result['graph_name']}' 已创建。",
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "set_axis_title",
                "set_plot_color",
                "export_graph",
            ],
        )

    # =================================================================
    # list_graphs
    # =================================================================

    @mcp.tool()
    @tool_error_handler("列出图表", "请确认 Origin 已连接。")
    def list_graphs() -> dict:
        """List all graphs in the current project.

        When to use: To see available graphs or get graph names.
        When not to use: If graph name is already known.

        Examples:
        - list_graphs()
        """
        def _list(op: Any) -> list[dict[str, Any]]:
            graphs = []
            for g in op.pages("Graph"):
                graphs.append({
                    "graph_name": g.name,
                    "layers": (
                        len(g) if hasattr(g, "__len__") else 1
                    ),
                })
            return graphs

        result = manager.execute(_list)

        return success_response(
            message=f"当前项目共有 {len(result)} 个图表。",
            data={
                "graphs": result,
                "count": len(result),
                "active_graph": manager.active_graph,
            },
            resource=manager.get_resource_context(),
            next_suggestions=[
                "set_axis_title",
                "set_plot_color",
                "export_graph",
            ],
        )

    # =================================================================
    # list_graph_templates
    # =================================================================

    @mcp.tool()
    def list_graph_templates() -> dict:
        """List supported graph templates with recommended use cases.

        When to use: When unsure which graph type to use; browse available
        templates and recommendations.
        When not to use: If the desired plot_type is already known.

        Examples:
        - list_graph_templates()
        """
        templates = [
            {"type": "line", "template": "line",
             "description": "折线图，适用于连续数据趋势展示"},
            {"type": "scatter", "template": "scatter",
             "description": "散点图，适用于数据分布和相关性分析"},
            {"type": "line_symbol", "template": "linesymb",
             "description": "折线+符号图，同时显示数据点和趋势"},
            {"type": "column", "template": "column",
             "description": "柱状图，适用于分类数据对比"},
            {"type": "area", "template": "area",
             "description": "面积图，适用于累积量展示"},
            {"type": "auto", "template": "line",
             "description": "自动选择（默认折线图）"},
        ]

        return success_response(
            message=f"共有 {len(templates)} 种图表模板可用。",
            data={"templates": templates},
            resource=manager.get_resource_context(),
            next_suggestions=["create_plot"],
        )

    # =================================================================
    # remove_plot_from_graph
    # =================================================================

    @mcp.tool()
    @tool_error_handler("移除曲线", "请检查 plot_index 是否在范围内。调用 get_graph_info 查看曲线列表。")
    def remove_plot_from_graph(
        plot_index: int,
        graph_name: str | None = None,
    ) -> dict:
        """Remove a specified curve from a graph.

        When to use: To delete an unwanted curve from a graph.
        When not to use: To delete the entire graph, this tool is not needed.

        Default behavior:
        - graph_name omitted: uses current active graph

        Examples:
        - remove_plot_from_graph(plot_index=0)
        - remove_plot_from_graph(plot_index=2, graph_name="Graph1")
        """
        target_graph = _resolve_graph_name(graph_name, manager)

        def _remove(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_graph)
            gl = _get_layer(gr, 0)
            _get_plot(gl, plot_index)
            gl.remove_plot(plot_index)

            remaining = len(gl.plot_list()) if hasattr(gl, "plot_list") else 0

            return {
                "graph_name": target_graph,
                "removed_index": plot_index,
                "remaining_plots": remaining,
            }

        result = manager.execute(_remove)

        return success_response(
            message=(
                f"已从图表 '{target_graph}' 移除曲线 {plot_index}，"
                f"剩余 {result['remaining_plots']} 条曲线。"
            ),
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=["get_graph_info", "export_graph"],
        )

    # =================================================================
    # add_graph_layer
    # =================================================================

    @mcp.tool()
    @tool_error_handler("添加图层", "请检查图表是否存在和 layer_type 值是否有效。")
    def add_graph_layer(
        layer_type: int = 2,
        graph_name: str | None = None,
    ) -> dict:
        """Add a new layer to a graph (e.g. right Y axis, top X axis).

        When to use: To display data with different scales using multiple
        Y or X axes on the same graph.
        When not to use: For simple single-Y-axis graphs. For double Y
        axis, can also use create_double_y_plot directly.

        Default behavior:
        - graph_name omitted: uses current active graph
        - layer_type defaults to 2 (right Y axis)

        Parameter notes:
        - layer_type: layer type number
            - 2 = right Y axis (righty)
            - 3 = top X axis (topx)
            - 4 = right Y + top X (rightytopx)
            - See Origin layadd documentation for other values

        Examples:
        - add_graph_layer()
        - add_graph_layer(layer_type=3, graph_name="Graph1")
        """
        target_graph = _resolve_graph_name(graph_name, manager)

        def _add_layer(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_graph)
            gr.add_layer(layer_type)

            layer_types = {2: "右Y轴", 3: "顶X轴", 4: "右Y+顶X"}
            return {
                "graph_name": target_graph,
                "layer_type": layer_type,
                "layer_type_name": layer_types.get(layer_type, f"类型{layer_type}"),
                "total_layers": len(gr),
            }

        result = manager.execute(_add_layer)

        return success_response(
            message=(
                f"已为图表 '{target_graph}' 添加"
                f" {result['layer_type_name']} 图层，"
                f"当前共 {result['total_layers']} 个图层。"
            ),
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "add_plot_to_graph",
                "set_axis_title",
            ],
        )

    # =================================================================
    # change_plot_data
    # =================================================================

    @mcp.tool()
    @tool_error_handler("更换数据源", "请检查工作表和列索引是否正确。")
    def change_plot_data(
        x_col: int | str,
        y_col: int | str,
        plot_index: int = 0,
        sheet_name: str | None = None,
        graph_name: str | None = None,
    ) -> dict:
        """Change the data source of an existing curve in a graph.

        When to use: To replace the data columns bound to a curve without
        recreating the graph.
        When not to use: To add a new curve, use add_plot_to_graph.

        Default behavior:
        - graph_name omitted: uses current active graph
        - sheet_name omitted: uses current active worksheet
        - plot_index defaults to 0 (first curve)

        Examples:
        - change_plot_data(x_col=0, y_col=2)
        - change_plot_data(x_col="A", y_col="D", plot_index=1)
        """
        target_graph = _resolve_graph_name(graph_name, manager)
        target_sheet = _resolve_worksheet_name(sheet_name, manager)

        def _change(op: Any) -> dict[str, Any]:
            wks = _find_worksheet(op, target_sheet)
            gr = _find_graph(op, target_graph)
            gl = _get_layer(gr, 0)

            dp = _get_plot(gl, plot_index)
            dp.change_data(wks, x=x_col, y=y_col)
            gl.rescale()

            return {
                "graph_name": target_graph,
                "sheet_name": target_sheet,
                "plot_index": plot_index,
                "new_x_col": x_col,
                "new_y_col": y_col,
            }

        result = manager.execute(_change)

        return success_response(
            message=(
                f"曲线 {plot_index} 的数据源已更换为 "
                f"X={x_col}, Y={y_col}。"
            ),
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=["set_axis_title", "export_graph"],
        )

    # =================================================================
    # copy_graph_to_clipboard
    # =================================================================

    @mcp.tool()
    @tool_error_handler("复制图表", "请检查图表是否存在。")
    def copy_graph_to_clipboard(
        format: str = "png",
        dpi: int = 300,
        graph_name: str | None = None,
    ) -> dict:
        """Copy graph to system clipboard (paste directly into Word/PPT).

        When to use: To quickly paste a graph into a document or
        presentation.
        When not to use: To save as a file, use export_graph.

        Default behavior:
        - graph_name omitted: uses current active graph
        - format defaults to "png"
        - dpi defaults to 300

        Parameter notes:
        - format: image format ("png", "emf", "dib", "jpg")
        - dpi: resolution

        Examples:
        - copy_graph_to_clipboard()
        - copy_graph_to_clipboard(format="emf", dpi=600)
        """
        target_graph = _resolve_graph_name(graph_name, manager)

        def _copy(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_graph)
            transparent = (format.lower() == "png")
            gr.copy_page(format.upper(), dpi, 100, transparent)

            return {
                "graph_name": target_graph,
                "format": format,
                "dpi": dpi,
            }

        result = manager.execute(_copy)

        return success_response(
            message=(
                f"图表 '{target_graph}' 已复制到剪贴板 "
                f"({format.upper()}, {dpi} DPI)。"
            ),
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=["export_graph"],
        )

    # =================================================================
    # group_plots
    # =================================================================

    @mcp.tool()
    @tool_error_handler("分组曲线", "请检查曲线索引范围。")
    def group_plots(
        begin: int = 0,
        end: int = -1,
        graph_name: str | None = None,
    ) -> dict:
        """Group curves in a graph (enabling linked color/symbol increments).

        When to use: When multiple curves need a unified color/symbol
        increment scheme.
        When not to use: For independent curves that don't need grouping.

        Default behavior:
        - graph_name omitted: uses current active graph
        - begin=0, end=-1 groups all curves together

        Parameter notes:
        - begin: starting curve index (0-based)
        - end: ending curve index (-1 means last curve)

        Examples:
        - group_plots()
        - group_plots(begin=0, end=2)
        - group_plots(begin=3, end=5)
        """
        target_graph = _resolve_graph_name(graph_name, manager)

        def _group(op: Any) -> dict[str, Any]:
            gr = _find_graph(op, target_graph)
            gl = _get_layer(gr, 0)
            gl.group(True, begin, end)

            return {
                "graph_name": target_graph,
                "begin": begin,
                "end": end,
            }

        result = manager.execute(_group)

        end_str = "最后" if end == -1 else str(end)
        return success_response(
            message=f"曲线 {begin} 至 {end_str} 已分组。",
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "set_color_increment",
                "set_symbol_increment",
            ],
        )
