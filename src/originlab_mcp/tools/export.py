"""导出与项目管理类 tools。

让 AI 能完成结果交付和项目持久化。

包含:
    export_graph: 导出图表为图片
    export_worksheet_to_csv: 导出工作表为 CSV
    save_project: 保存项目
    open_project: 打开项目
    new_project: 新建空白项目
"""

from __future__ import annotations

import os
from typing import Any

from originlab_mcp.exceptions import (
    NoActiveGraphError,
    NoActiveWorksheetError,
    ToolError,
)
from originlab_mcp.utils.constants import (
    DEFAULT_EXPORT_FORMAT,
    DEFAULT_EXPORT_WIDTH,
    ExportFormat,
)
from originlab_mcp.utils.helpers import (
    find_graph,
    find_worksheet,
    resolve_graph_name,
    resolve_worksheet_name,
    tool_error_handler,
)
from originlab_mcp.utils.validators import (
    error_response,
    error_response_from_exception,
    success_response,
    validate_export_format,
    validate_file_path,
)


def register_export_tools(mcp: Any, manager: Any) -> None:
    """注册导出与项目管理类 tools 到 MCP Server。

    Args:
        mcp: FastMCP 实例。
        manager: OriginManager 实例（依赖注入）。
    """

    # =================================================================
    # export_graph
    # =================================================================

    @mcp.tool()
    @tool_error_handler("导出图表", "请检查输出路径和格式。")
    def export_graph(
        output_path: str,
        graph_name: str | None = None,
        output_format: str | None = None,
        width: int = DEFAULT_EXPORT_WIDTH,
    ) -> dict:
        """导出图表为图片文件。

        何时使用：需要将图表导出为 PNG/SVG/PDF 文件时使用。
        何时不用：只想在 Origin 中查看图表时无需调用。

        默认行为：
        - graph_name 省略时使用当前活动图表
        - output_format 省略时根据 output_path 的扩展名推断（默认 png）
        - width 默认 800 像素

        示例：
        - export_graph(output_path="C:\\\\output\\\\chart.png")
        - export_graph(output_path="C:\\\\output\\\\chart.svg", output_format="svg", width=1200)
        """
        # 推断格式
        fmt = output_format
        if fmt is None:
            ext = os.path.splitext(output_path)[1].lstrip(".").lower()
            fmt = ext if ext in [e.value for e in ExportFormat] else DEFAULT_EXPORT_FORMAT.value

        err = validate_export_format(fmt)
        if err:
            return error_response(
                message=err,
                error_type="unsupported",
                target="output_format",
                value=fmt,
                hint=f"支持的格式: {[e.value for e in ExportFormat]}",
            )

        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.isdir(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
            except OSError as e:
                return error_response(
                    message=f"无法创建输出目录: {e}",
                    error_type="invalid_input",
                    target="output_path",
                    value=output_path,
                    hint="请检查输出路径是否有效。",
                )

        target_graph = resolve_graph_name(graph_name, manager)

        def _export(op: Any) -> dict[str, Any]:
            gr = find_graph(op, target_graph)
            gr.save_fig(output_path, type=fmt, width=width)
            return {
                "graph_name": target_graph,
                "output_path": os.path.abspath(output_path),
                "format": fmt,
                "width": width,
            }

        result = manager.execute(_export)

        return success_response(
            message=f"图表已导出到 '{result['output_path']}'。",
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=["save_project"],
        )

    # =================================================================
    # export_worksheet_to_csv
    # =================================================================

    @mcp.tool()
    @tool_error_handler("导出工作表", "请检查工作表和输出路径。")
    def export_worksheet_to_csv(
        output_path: str,
        sheet_name: str | None = None,
    ) -> dict:
        """将工作表数据导出为 CSV 文件。

        何时使用：需要将 Origin 工作表的数据保存为 CSV 文件时使用。
        何时不用：只需在 Origin 内查看数据时请用 get_worksheet_data。

        默认行为：
        - sheet_name 省略时使用当前活动工作表

        示例：
        - export_worksheet_to_csv(output_path="C:\\\\output\\\\data.csv")
        """
        target_name = resolve_worksheet_name(sheet_name, manager)

        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.isdir(output_dir):
            os.makedirs(output_dir, exist_ok=True)

        def _export(op: Any) -> dict[str, Any]:
            wks = find_worksheet(op, target_name)

            # 收集数据
            rows = []
            headers = []
            for ci in range(wks.cols):
                col = wks.get_col(ci)
                long_name = col.get_label("L") if hasattr(col, "get_label") else ""
                headers.append(long_name or col.name)

            data_lists = [wks.to_list(ci) for ci in range(wks.cols)]
            max_rows = max((len(d) for d in data_lists), default=0)

            with open(output_path, "w", encoding="utf-8", newline="") as f:
                f.write(",".join(headers) + "\n")
                for ri in range(max_rows):
                    row_vals = []
                    for ci in range(wks.cols):
                        if ri < len(data_lists[ci]):
                            row_vals.append(str(data_lists[ci][ri]))
                        else:
                            row_vals.append("")
                    f.write(",".join(row_vals) + "\n")

            return {
                "sheet_name": target_name,
                "output_path": os.path.abspath(output_path),
                "rows": max_rows,
                "columns": len(headers),
            }

        result = manager.execute(_export)

        return success_response(
            message=f"工作表已导出到 '{result['output_path']}'（{result['rows']} 行 x {result['columns']} 列）。",
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=["save_project"],
        )

    # =================================================================
    # save_project
    # =================================================================

    @mcp.tool()
    @tool_error_handler("保存项目", "请检查文件路径和写入权限。")
    def save_project(file_path: str | None = None) -> dict:
        """保存当前 Origin 项目。

        何时使用：需要将当前工作保存到 .opju 文件时使用。
        何时不用：只是临时查看数据不需要持久化时无需调用。

        默认行为：
        - file_path 省略时保存到当前项目路径（如果有的话）

        示例：
        - save_project()
        - save_project(file_path="C:\\\\data\\\\analysis.opju")
        """
        def _save(op: Any) -> str:
            if file_path:
                save_dir = os.path.dirname(file_path)
                if save_dir and not os.path.isdir(save_dir):
                    os.makedirs(save_dir, exist_ok=True)
                op.save(file_path)
                return os.path.abspath(file_path)
            else:
                op.save()
                return "current project"

        saved_path = manager.execute(_save)

        return success_response(
            message=f"项目已保存到 '{saved_path}'。",
            data={"saved_path": saved_path},
            resource=manager.get_resource_context(),
        )

    # =================================================================
    # open_project
    # =================================================================

    @mcp.tool()
    @tool_error_handler("打开项目", "请检查文件是否为有效的 Origin 项目文件。")
    def open_project(file_path: str, readonly: bool = False) -> dict:
        """打开 Origin 项目文件。

        何时使用：需要打开一个已有的 .opju 项目文件时使用。
        何时不用：当前已在项目中工作且不需要切换项目时无需调用。

        默认行为：
        - readonly 默认为 false

        示例：
        - open_project(file_path="C:\\\\data\\\\analysis.opju")
        - open_project(file_path="C:\\\\data\\\\analysis.opju", readonly=True)
        """
        err = validate_file_path(file_path)
        if err:
            return error_response(
                message=err,
                error_type="invalid_input",
                target="file_path",
                value=file_path,
                hint="文件不存在，请检查文件路径。",
            )

        def _open(op: Any) -> str:
            op.open(file=file_path, readonly=readonly)

            manager.active_worksheet = None
            manager.active_graph = None

            # 尝试将第一个工作表设为活动
            try:
                for book in op.pages("Book"):
                    for sheet in book:
                        manager.active_worksheet = f"[{book.name}]{sheet.name}"
                        break
                    break
            except Exception:
                pass

            return os.path.abspath(file_path)

        opened_path = manager.execute(_open)

        return success_response(
            message=f"项目已打开: '{opened_path}'。",
            data={"file_path": opened_path, "readonly": readonly},
            resource=manager.get_resource_context(),
            next_suggestions=["list_worksheets", "list_graphs"],
        )

    # =================================================================
    # new_project
    # =================================================================

    @mcp.tool()
    @tool_error_handler("新建项目", "请检查 Origin 连接状态。")
    def new_project() -> dict:
        """新建空白 Origin 项目。

        何时使用：需要从零开始一个全新的项目时使用。注意这会清除当前项目中所有未保存的内容。
        何时不用：在当前项目继续工作时不要调用。请先用 save_project 保存。

        示例：
        - new_project()
        """
        def _new(op: Any) -> None:
            op.new()
            manager.active_worksheet = None
            manager.active_graph = None

        manager.execute(_new)

        return success_response(
            message="已创建新的空白项目。",
            data={"status": "reset"},
            resource=manager.get_resource_context(),
            warnings=["之前项目中未保存的内容已丢失。"],
            next_suggestions=[
                "import_csv",
                "import_excel",
                "import_data_from_text",
            ],
        )
