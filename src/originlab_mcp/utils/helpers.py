"""共享辅助函数。

提供跨 tool 模块复用的查找、解析函数，
以及统一的错误处理装饰器，消除重复代码。
"""

from __future__ import annotations

import functools
import logging
import re
from collections.abc import Callable
from typing import Any

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


def get_graph_layer(gr: Any, layer_index: int = 0) -> Any:
    """获取图表的指定图层，越界时抛出异常。

    Args:
        gr: 图表对象。
        layer_index: 图层索引（0-based），默认 0 为主图层。

    Returns:
        图层对象。

    Raises:
        LayerIndexError: 图层索引越界时。
    """
    from originlab_mcp.exceptions import LayerIndexError

    total = len(gr)
    if layer_index < 0 or layer_index >= total:
        raise LayerIndexError(layer_index, total)
    return gr[layer_index]


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
    if plot_index < 0:
        raise PlotIndexError(plot_index)

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


# 合法的 Origin 对象名：字母、数字、下划线，首字符为字母或下划线
_SAFE_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def sanitize_labtalk_name(name: str, param_name: str = "name") -> str:
    """验证 Origin 对象名是否安全，防止 LabTalk 命令注入。

    Args:
        name: 要验证的对象名（图表名、工作表名等）。
        param_name: 参数名称，用于错误提示。

    Returns:
        通过验证的原始名称。

    Raises:
        ToolError: 名称包含非法字符时。
    """
    if not name:
        raise ToolError(
            f"{param_name} 不能为空",
            error_type="invalid_input",
            target=param_name,
            hint=f"请提供有效的 {param_name}。",
        )
    if not _SAFE_NAME_RE.match(name):
        raise ToolError(
            f"{param_name} '{name}' 包含非法字符，仅允许字母、数字和下划线",
            error_type="invalid_input",
            target=param_name,
            value=name,
            hint="Origin 对象名只能包含英文字母、数字和下划线，且不能以数字开头。",
        )
    return name

