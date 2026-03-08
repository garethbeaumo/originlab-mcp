"""共享辅助函数。

提供跨 tool 模块复用的查找、解析函数，消除重复代码。
"""

from __future__ import annotations

from typing import Any

from originlab_mcp.exceptions import (
    GraphNotFoundError,
    InvalidAxisError,
    NoActiveGraphError,
    NoActiveWorksheetError,
    PlotIndexError,
    WorksheetNotFoundError,
    ColumnIndexError,
)
from originlab_mcp.origin_manager import OriginManager


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


def find_worksheet(op: Any, sheet_name: str) -> Any:
    """查找工作表，不存在时抛出异常。

    Args:
        op: originpro 模块引用。
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


def find_graph(op: Any, graph_name: str) -> Any:
    """查找图表对象，不存在时抛出异常。

    Args:
        op: originpro 模块引用。
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
