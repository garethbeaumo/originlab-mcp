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
        """创建新图表（支持单条或多条曲线）。

        何时使用：需要从零开始创建一个新图表时使用。
        何时不用：如果要在已有图表上追加曲线，请使用 add_plot_to_graph。

        默认行为：
        - sheet_name 省略时使用当前活动工作表
        - plot_type 省略时默认为 line
        - yerr_col 省略时不绘制 Y 方向误差棒
        - xerr_col 省略时不绘制 X 方向误差棒

        示例：
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
        """在已有图表上追加一条或多条曲线。

        何时使用：需要在已存在的图表上叠加新的数据曲线时使用。
        何时不用：如果需要创建全新图表，请使用 create_plot。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - sheet_name 省略时使用当前活动工作表
        - yerr_col 省略时不绘制 Y 方向误差棒
        - xerr_col 省略时不绘制 X 方向误差棒

        示例：
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
        """创建双 Y 轴图。

        何时使用：需要在同一图表中用左右两个 Y 轴展示不同量级或单位的数据时使用。
        何时不用：如果所有 Y 列量级相近，使用 create_plot 即可。

        默认行为：
        - sheet_name 省略时使用当前活动工作表
        - y1_col 绑定左 Y 轴，y2_col 绑定右 Y 轴

        示例：
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
        """列出当前项目中的所有图表。

        何时使用：需要查看有哪些图表、获取图表名称时使用。
        何时不用：已知图表名称时无需调用。

        示例：
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
        """列出支持的图表模板及推荐使用场景。

        何时使用：不确定应该使用什么图表类型时，查看可用模板和使用建议。
        何时不用：已知要使用的 plot_type 时无需调用。

        示例：
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
        """从图表中移除指定的曲线。

        何时使用：需要从图表中删除某条不需要的曲线时使用。
        何时不用：需要删除整个图表时无需调用此工具。

        默认行为：
        - graph_name 省略时使用当前活动图表

        示例：
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
        """为图表添加新的图层（如右 Y 轴、顶 X 轴等）。

        何时使用：需要在同一图表上使用多个 Y 轴或 X 轴展示不同量纲数据时使用。
        何时不用：只需单个 Y 轴的简单图表无需调用。如需双 Y 轴，也可直接用 create_double_y_plot。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - layer_type 默认为 2（右 Y 轴）

        参数说明：
        - layer_type: 图层类型编号
            - 2 = 右 Y 轴 (righty)
            - 3 = 顶 X 轴 (topx)
            - 4 = 右 Y + 顶 X (rightytopx)
            - 其他值参见 Origin layadd 文档

        示例：
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
        """更换图表中已有曲线的数据源。

        何时使用：需要替换曲线绑定的数据列（不重新创建图表）时使用。
        何时不用：需要添加新曲线请用 add_plot_to_graph。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - sheet_name 省略时使用当前活动工作表
        - plot_index 默认为 0（第一条曲线）

        示例：
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
        """将图表复制到系统剪贴板（可直接粘贴到 Word/PPT）。

        何时使用：需要将图表快速粘贴到文档或演示文稿中时使用。
        何时不用：需要保存为文件请用 export_graph。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - format 默认为 "png"
        - dpi 默认为 300

        参数说明：
        - format: 图片格式（"png", "emf", "dib", "jpg"）
        - dpi: 分辨率

        示例：
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
        """对图表中的曲线进行分组（使其联动颜色/符号递增）。

        何时使用：多条曲线需要统一的颜色/符号递增方案时使用。
        何时不用：不需要分组递增的独立曲线无需调用。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - begin=0, end=-1 表示将所有曲线分为一组

        参数说明：
        - begin: 起始曲线索引（0-offset）
        - end: 结束曲线索引（-1 表示到最后一条）

        示例：
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
