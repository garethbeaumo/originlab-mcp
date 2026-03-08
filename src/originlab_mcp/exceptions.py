"""OriginLab MCP Server 自定义异常。

定义了 tool 执行过程中的结构化异常，替代 (result, status) 元组模式。
每个异常携带足够的上下文信息，可直接转换为 error_response。
"""

from __future__ import annotations


class ToolError(Exception):
    """tool 执行过程中的基类异常。

    Attributes:
        error_type: 错误类别（对应 constants.ErrorType）。
        target: 出错的参数名或对象名。
        value: 导致错误的实际值。
        hint: 给 AI 的修复建议。
    """

    def __init__(
        self,
        message: str,
        *,
        error_type: str = "internal_error",
        target: str = "",
        value: object = None,
        hint: str = "",
    ) -> None:
        super().__init__(message)
        self.error_type = error_type
        self.target = target
        self.value = value
        self.hint = hint


class WorksheetNotFoundError(ToolError):
    """指定的工作表不存在。"""

    def __init__(self, sheet_name: str) -> None:
        super().__init__(
            f"工作表 '{sheet_name}' 不存在",
            error_type="not_found",
            target="worksheet",
            value=sheet_name,
            hint="Call list_worksheets to inspect available worksheet names.",
        )


class GraphNotFoundError(ToolError):
    """指定的图表不存在。"""

    def __init__(self, graph_name: str) -> None:
        super().__init__(
            f"图表 '{graph_name}' 不存在",
            error_type="not_found",
            target="graph",
            value=graph_name,
            hint="Call list_graphs to inspect available graph names.",
        )


class NoActiveWorksheetError(ToolError):
    """未指定工作表且无活动工作表。"""

    def __init__(self) -> None:
        super().__init__(
            "未指定工作表且无活动工作表",
            error_type="invalid_input",
            target="sheet_name",
            hint="请指定 sheet_name，或先调用 import_csv / list_worksheets。",
        )


class NoActiveGraphError(ToolError):
    """未指定图表且无活动图表。"""

    def __init__(self) -> None:
        super().__init__(
            "未指定图表且无活动图表",
            error_type="invalid_input",
            target="graph_name",
            hint="请指定 graph_name 或先创建图表。",
        )


class ColumnIndexError(ToolError):
    """列索引超出范围。"""

    def __init__(self, col_index: int, total_cols: int) -> None:
        super().__init__(
            f"Column index {col_index} out of range. "
            f"Current worksheet has {total_cols} columns (0-{total_cols - 1}).",
            error_type="invalid_input",
            target="col",
            value=col_index,
            hint="Call get_worksheet_info to inspect column details.",
        )


class InvalidAxisError(ToolError):
    """不支持的轴标识。"""

    def __init__(self, axis: str, supported: tuple[str, ...] = ("x", "y")) -> None:
        super().__init__(
            f"不支持的轴标识 '{axis}'",
            error_type="invalid_input",
            target="axis",
            value=axis,
            hint=f"Supported axis values: {list(supported)}.",
        )


class PlotIndexError(ToolError):
    """曲线索引不存在。"""

    def __init__(self, plot_index: int) -> None:
        super().__init__(
            f"曲线索引 {plot_index} 不存在",
            error_type="invalid_input",
            target="plot_index",
            value=plot_index,
            hint="请检查图表中的曲线数量。plot_index 从 0 开始。",
        )


class FitFunctionNotFoundError(ToolError):
    """拟合函数不存在。"""

    def __init__(self, function_name: str) -> None:
        super().__init__(
            f"拟合函数 '{function_name}' 不存在或不可用",
            error_type="not_found",
            target="function_name",
            value=function_name,
            hint="请调用 list_fit_functions 查看可用的拟合函数。",
        )


class FitConvergenceError(ToolError):
    """拟合未收敛。"""

    def __init__(self, function_name: str, detail: str = "") -> None:
        msg = f"使用函数 '{function_name}' 拟合未收敛"
        if detail:
            msg += f": {detail}"
        super().__init__(
            msg,
            error_type="internal_error",
            target="fit",
            value=function_name,
            hint="请尝试调整初始参数值或选择其他拟合函数。",
        )

