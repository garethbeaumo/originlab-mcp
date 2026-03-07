"""
绘图类 tools

让 AI 能把结构化数据稳定转换成图表。

包含：
- create_plot
- add_plot_to_graph
- create_double_y_plot
- list_graphs
- list_graph_templates
"""

from __future__ import annotations

from originlab_mcp.origin_manager import OriginManager
from originlab_mcp.utils.constants import (
    DEFAULT_PLOT_TYPE,
    PLOT_TYPE_TO_TEMPLATE,
    PlotType,
)
from originlab_mcp.utils.validators import (
    error_response,
    normalize_y_cols,
    success_response,
    validate_column_index,
    validate_column_indices,
    validate_plot_type,
)


def register_plot_tools(mcp) -> None:
    """注册绘图类 tools 到 MCP Server。"""

    manager = OriginManager()

    # =================================================================
    # create_plot
    # =================================================================

    @mcp.tool()
    def create_plot(
        x_col: int,
        y_cols: int | list[int],
        sheet_name: str | None = None,
        plot_type: str = DEFAULT_PLOT_TYPE.value,
        z_col: int | None = None,
    ) -> dict:
        """创建新图表（支持单条或多条曲线）。

        何时使用：需要从零开始创建一个新图表时使用。
        何时不用：如果要在已有图表上追加曲线，请使用 add_plot_to_graph。

        默认行为：
        - sheet_name 省略时使用当前活动工作表
        - plot_type 省略时默认为 line

        示例：
        - create_plot(x_col=0, y_cols=1)
        - create_plot(x_col=0, y_cols=[1, 2], plot_type="scatter")
        """
        # 验证 plot_type
        err = validate_plot_type(plot_type)
        if err:
            return error_response(
                message=err,
                error_type="unsupported",
                target="plot_type",
                value=plot_type,
                hint=f"Supported types: {[e.value for e in PlotType]}",
            )

        target_name = sheet_name or manager.active_worksheet
        if not target_name:
            return error_response(
                message="未指定工作表且无活动工作表",
                error_type="invalid_input",
                target="sheet_name",
                hint="请指定 sheet_name 或先导入数据。",
            )

        y_col_list = normalize_y_cols(y_cols)

        try:
            def _plot(op):
                wks = op.find_sheet("w", target_name)
                if wks is None:
                    return None, "sheet_not_found"

                total_cols = wks.cols
                # 验证列索引
                col_err = validate_column_index(x_col, total_cols)
                if col_err:
                    return col_err, "col_error"
                col_err = validate_column_indices(y_col_list, total_cols)
                if col_err:
                    return col_err, "col_error"

                template = PLOT_TYPE_TO_TEMPLATE.get(plot_type, "line")
                gr = op.new_graph(template=template)
                gl = gr[0]

                curves = []
                for i, yc in enumerate(y_col_list):
                    plot = gl.add_plot(wks, coly=yc, colx=x_col)
                    curves.append({"y_col": yc, "x_col": x_col, "plot_index": i})

                gl.rescale()
                try:
                    gl.group()
                except Exception:
                    pass  # group 可能在只有一条曲线时不需要

                graph_name = gr.name
                manager.active_graph = graph_name

                return {
                    "graph_name": graph_name,
                    "plot_type": plot_type,
                    "sheet_name": target_name,
                    "curves": curves,
                    "curve_count": len(curves),
                }, "ok"

            result, status = manager.execute(_plot)

            if status == "sheet_not_found":
                return error_response(
                    message=f"工作表 '{target_name}' 不存在",
                    error_type="not_found",
                    target="worksheet",
                    value=target_name,
                    hint="Call list_worksheets to inspect available worksheet names.",
                )

            if status == "col_error":
                return error_response(
                    message=result,
                    error_type="invalid_input",
                    target="y_cols",
                    value=y_col_list,
                    hint="Call get_worksheet_info to inspect column details.",
                )

            return success_response(
                message=f"图表 '{result['graph_name']}' 已创建，共 {result['curve_count']} 条曲线。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "set_axis_title",
                    "set_plot_color",
                    "add_plot_to_graph",
                    "export_graph",
                ],
            )
        except Exception as e:
            return error_response(
                message=f"创建图表失败: {e}",
                error_type="internal_error",
                target="create_plot",
                hint="请检查工作表数据和列索引是否正确。",
            )

    # =================================================================
    # add_plot_to_graph
    # =================================================================

    @mcp.tool()
    def add_plot_to_graph(
        x_col: int,
        y_cols: int | list[int],
        graph_name: str | None = None,
        sheet_name: str | None = None,
        plot_type: str | None = None,
    ) -> dict:
        """在已有图表上追加一条或多条曲线。

        何时使用：需要在已存在的图表上叠加新的数据曲线时使用。
        何时不用：如果需要创建全新图表，请使用 create_plot。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - sheet_name 省略时使用当前活动工作表

        示例：
        - add_plot_to_graph(x_col=0, y_cols=2)
        - add_plot_to_graph(x_col=0, y_cols=[3, 4], graph_name="Graph1")
        """
        target_graph = graph_name or manager.active_graph
        target_sheet = sheet_name or manager.active_worksheet

        if not target_graph:
            return error_response(
                message="未指定图表且无活动图表",
                error_type="invalid_input",
                target="graph_name",
                hint="请指定 graph_name 或先创建图表。",
            )
        if not target_sheet:
            return error_response(
                message="未指定工作表且无活动工作表",
                error_type="invalid_input",
                target="sheet_name",
                hint="请指定 sheet_name 或先导入数据。",
            )

        y_col_list = normalize_y_cols(y_cols)

        try:
            def _add(op):
                wks = op.find_sheet("w", target_sheet)
                if wks is None:
                    return None, "sheet_not_found"

                gr = op.find_graph(target_graph)
                if gr is None:
                    return None, "graph_not_found"

                total_cols = wks.cols
                col_err = validate_column_indices(y_col_list, total_cols)
                if col_err:
                    return col_err, "col_error"

                gl = gr[0]
                new_curves = []
                # 计算当前已有的曲线数作为起始 index
                existing_count = gl.num_plots if hasattr(gl, "num_plots") else 0

                for i, yc in enumerate(y_col_list):
                    gl.add_plot(wks, coly=yc, colx=x_col)
                    new_curves.append({
                        "y_col": yc,
                        "x_col": x_col,
                        "plot_index": existing_count + i,
                    })

                gl.rescale()
                try:
                    gl.group()
                except Exception:
                    pass

                total_count = existing_count + len(y_col_list)

                return {
                    "graph_name": target_graph,
                    "sheet_name": target_sheet,
                    "new_curves": new_curves,
                    "new_curve_count": len(new_curves),
                    "total_curve_count": total_count,
                }, "ok"

            result, status = manager.execute(_add)

            if status == "sheet_not_found":
                return error_response(
                    message=f"工作表 '{target_sheet}' 不存在",
                    error_type="not_found",
                    target="worksheet",
                    value=target_sheet,
                    hint="Call list_worksheets to inspect available worksheet names.",
                )
            if status == "graph_not_found":
                return error_response(
                    message=f"图表 '{target_graph}' 不存在",
                    error_type="not_found",
                    target="graph",
                    value=target_graph,
                    hint="Call list_graphs to inspect available graph names.",
                )
            if status == "col_error":
                return error_response(
                    message=result,
                    error_type="invalid_input",
                    target="y_cols",
                    value=y_col_list,
                    hint="Call get_worksheet_info to inspect column details.",
                )

            return success_response(
                message=f"已向图表 '{target_graph}' 追加 {result['new_curve_count']} 条曲线，当前共 {result['total_curve_count']} 条。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "set_axis_title",
                    "set_plot_color",
                    "export_graph",
                ],
            )
        except Exception as e:
            return error_response(
                message=f"追加曲线失败: {e}",
                error_type="internal_error",
                target="add_plot_to_graph",
                hint="请检查图表和工作表是否存在，以及列索引是否正确。",
            )

    # =================================================================
    # create_double_y_plot
    # =================================================================

    @mcp.tool()
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
        target_name = sheet_name or manager.active_worksheet
        if not target_name:
            return error_response(
                message="未指定工作表且无活动工作表",
                error_type="invalid_input",
                target="sheet_name",
                hint="请指定 sheet_name 或先导入数据。",
            )

        try:
            def _dbl(op):
                wks = op.find_sheet("w", target_name)
                if wks is None:
                    return None, "sheet_not_found"

                total_cols = wks.cols
                for c, name in [(x_col, "x_col"), (y1_col, "y1_col"), (y2_col, "y2_col")]:
                    err = validate_column_index(c, total_cols)
                    if err:
                        return (err, name), "col_error"

                gr = op.new_graph(template="doubley")
                # 左 Y 轴
                gr[0].add_plot(wks, coly=y1_col, colx=x_col)
                # 右 Y 轴
                gr[1].add_plot(wks, coly=y2_col, colx=x_col)

                gr[0].rescale()
                gr[1].rescale()

                graph_name = gr.name
                manager.active_graph = graph_name

                return {
                    "graph_name": graph_name,
                    "sheet_name": target_name,
                    "left_y": {"y_col": y1_col, "x_col": x_col},
                    "right_y": {"y_col": y2_col, "x_col": x_col},
                }, "ok"

            result, status = manager.execute(_dbl)

            if status == "sheet_not_found":
                return error_response(
                    message=f"工作表 '{target_name}' 不存在",
                    error_type="not_found",
                    target="worksheet",
                    value=target_name,
                    hint="Call list_worksheets to inspect available worksheet names.",
                )
            if status == "col_error":
                err_msg, col_name = result
                return error_response(
                    message=err_msg,
                    error_type="invalid_input",
                    target=col_name,
                    hint="Call get_worksheet_info to inspect column details.",
                )

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
        except Exception as e:
            return error_response(
                message=f"创建双 Y 轴图失败: {e}",
                error_type="internal_error",
                target="create_double_y_plot",
                hint="请检查列索引是否正确。",
            )

    # =================================================================
    # list_graphs
    # =================================================================

    @mcp.tool()
    def list_graphs() -> dict:
        """列出当前项目中的所有图表。

        何时使用：需要查看有哪些图表、获取图表名称时使用。
        何时不用：已知图表名称时无需调用。

        示例：
        - list_graphs()
        """
        try:
            def _list(op):
                graphs = []
                for g in op.pages("Graph"):
                    graphs.append({
                        "graph_name": g.name,
                        "layers": len(g) if hasattr(g, "__len__") else 1,
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
        except Exception as e:
            return error_response(
                message=f"列出图表失败: {e}",
                error_type="internal_error",
                target="graphs",
                hint="请确认 Origin 已连接。",
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
            {"type": "line", "template": "line", "description": "折线图，适用于连续数据趋势展示"},
            {"type": "scatter", "template": "scatter", "description": "散点图，适用于数据分布和相关性分析"},
            {"type": "line_symbol", "template": "linesymb", "description": "折线+符号图，同时显示数据点和趋势"},
            {"type": "column", "template": "column", "description": "柱状图，适用于分类数据对比"},
            {"type": "area", "template": "area", "description": "面积图，适用于累积量展示"},
            {"type": "auto", "template": "line", "description": "自动选择（默认折线图）"},
        ]

        return success_response(
            message=f"共有 {len(templates)} 种图表模板可用。",
            data={"templates": templates},
            resource=manager.get_resource_context(),
            next_suggestions=["create_plot"],
        )
