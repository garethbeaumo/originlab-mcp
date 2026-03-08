"""共享辅助函数。

提供跨 tool 模块复用的查找、解析函数，
以及统一的错误处理装饰器，消除重复代码。
"""

from __future__ import annotations

import functools
import logging
from typing import Any, Callable

from originlab_mcp.exceptions import (
    ColumnIndexError,
    GraphNotFoundError,
    InvalidAxisError,
    NoActiveGraphError,
    NoActiveWorksheetError,
    PlotIndexError,
    ToolError,
    WorksheetNotFoundError,
)
from originlab_mcp.origin_manager import OriginManager
from originlab_mcp.types import (
    GraphProtocol,
    OriginProProtocol,
    WorksheetProtocol,
)
from originlab_mcp.utils.validators import (
    error_response,
    error_response_from_exception,
)

logger = logging.getLogger(__name__)


def tool_error_handler(
    target: str,
    error_hint: str = "请检查参数值。",
) -> Callable:
    """通用 tool 错误处理装饰器。

    消除每个 tool 函数中重复的 try/except 样板代码。
    捕获 ToolError 和通用 Exception，转换为标准 error_response。

    Args:
        target: 出错时的目标标识（如 tool 名称或参数名）。
        error_hint: 通用异常时给 AI 的提示信息。

    Returns:
        装饰后的函数。

    用法::

        @tool_error_handler("import_csv", "请检查文件路径和格式。")
        def import_csv(...) -> dict:
            ...  # 无需 try/except，直接写业务逻辑
    """
    def decorator(fn: Callable[..., dict]) -> Callable[..., dict]:
        @functools.wraps(fn)
        def wrapper(*args: Any, **kwargs: Any) -> dict:
            try:
                return fn(*args, **kwargs)
            except ToolError as e:
                return error_response_from_exception(e)
            except Exception as e:
                logger.exception("tool %s 执行失败", target)
                return error_response(
                    message=f"{target} 执行失败: {e}",
                    error_type="internal_error",
                    target=target,
                    hint=error_hint,
                )
        return wrapper
    return decorator



def resolve_worksheet_name(
    sheet_name: str | None,
    manager: OriginManager,
) -> str:
    """解析工作表名称，未指定时使用活动工作表。

    Args:
        sheet_name: 用户指定的工作表名，可为 None。
        manager: OriginManager 实例。

    Returns:
        解析后的工作表名称。

    Raises:
        NoActiveWorksheetError: 未指定且无活动工作表时。
    """
    name = sheet_name or manager.active_worksheet
    if not name:
        raise NoActiveWorksheetError()
    return name


def find_worksheet(op: OriginProProtocol, sheet_name: str) -> WorksheetProtocol:
    """查找工作表，不存在时抛出异常。

    Args:
        op: originpro 操作接口。
        sheet_name: 工作表名称。

    Returns:
        工作表对象。

    Raises:
        WorksheetNotFoundError: 工作表不存在时。
    """
    wks = op.find_sheet("w", sheet_name)
    if wks is None:
        raise WorksheetNotFoundError(sheet_name)
    return wks


def resolve_graph_name(
    graph_name: str | None,
    manager: OriginManager,
) -> str:
    """解析图表名称，未指定时使用活动图表。

    Args:
        graph_name: 用户指定的图表名，可为 None。
        manager: OriginManager 实例。

    Returns:
        解析后的图表名称。

    Raises:
        NoActiveGraphError: 未指定且无活动图表时。
    """
    name = graph_name or manager.active_graph
    if not name:
        raise NoActiveGraphError()
    return name


def find_graph(op: OriginProProtocol, graph_name: str) -> GraphProtocol:
    """查找图表对象，不存在时抛出异常。

    Args:
        op: originpro 操作接口。
        graph_name: 图表名称。

    Returns:
        图表对象。

    Raises:
        GraphNotFoundError: 图表不存在时。
    """
    gr = op.find_graph(graph_name)
    if gr is None:
        raise GraphNotFoundError(graph_name)
    return gr


def get_plot(gl: Any, plot_index: int) -> Any:
    """获取指定索引的曲线，不存在时抛出异常。

    Args:
        gl: 图层对象。
        plot_index: 曲线索引（0-based）。

    Returns:
        曲线对象。

    Raises:
        PlotIndexError: 曲线索引不存在时。
    """
    plot = gl.plot(plot_index)
    if plot is None:
        raise PlotIndexError(plot_index)
    return plot


def validate_axis(
    axis: str,
    supported: tuple[str, ...] = ("x", "y"),
) -> str:
    """验证并归一化轴标识。

    Args:
        axis: 用户传入的轴标识。
        supported: 支持的轴标识元组。

    Returns:
        归一化后的轴标识（小写）。

    Raises:
        InvalidAxisError: 不支持的轴标识时。
    """
    normalized = axis.lower()
    if normalized not in supported:
        raise InvalidAxisError(axis, supported)
    return normalized


def validate_column_indices(col_indices: list[int], total_cols: int) -> None:
    """验证列索引范围，越界时抛出 ColumnIndexError。

    Args:
        col_indices: 要验证的列索引列表。
        total_cols: 工作表总列数。

    Raises:
        ColumnIndexError: 列索引超出范围时。
    """
    for idx in col_indices:
        if idx < 0 or idx >= total_cols:
            raise ColumnIndexError(idx, total_cols)
