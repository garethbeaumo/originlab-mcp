"""
统一返回结构与输入校验模块

提供 success_response / error_response 构建函数，
以及常用的参数验证工具函数。
"""

from __future__ import annotations

import os
from typing import Any

from .constants import (
    ColumnDesignation,
    ErrorType,
    ExportFormat,
    PlotType,
    ScaleType,
)
from originlab_mcp.exceptions import ToolError


# ===================================================================
# 统一返回结构
# ===================================================================


def success_response(
    message: str,
    data: dict[str, Any] | None = None,
    resource: dict[str, Any] | None = None,
    warnings: list[str] | None = None,
    next_suggestions: list[str] | None = None,
) -> dict[str, Any]:
    """构建统一的成功返回结构。

    Args:
        message: 简短结果摘要。
        data: 与本次操作直接相关的结构化数据。
        resource: 新建或更新后的核心对象标识。
        warnings: 非致命问题列表。
        next_suggestions: 建议 AI 下一步可调用的 tool。

    Returns:
        统一格式的成功返回字典。
    """
    return {
        "ok": True,
        "message": message,
        "data": data or {},
        "resource": resource or {},
        "warnings": warnings or [],
        "next_suggestions": next_suggestions or [],
    }


def error_response(
    message: str,
    error_type: str | ErrorType,
    target: str,
    value: Any = None,
    hint: str = "",
) -> dict[str, Any]:
    """构建统一的错误返回结构。

    Args:
        message: 简短的错误摘要。
        error_type: 错误类别（见 constants.ErrorType）。
        target: 出错的参数名或对象名。
        value: 导致错误的实际值。
        hint: 给 AI 的修复建议。

    Returns:
        统一格式的错误返回字典。
    """
    return {
        "ok": False,
        "message": message,
        "error": {
            "type": str(error_type.value if isinstance(error_type, ErrorType) else error_type),
            "target": target,
            "value": value,
            "hint": hint,
        },
    }


# ===================================================================
# 参数验证工具
# ===================================================================


def validate_file_path(file_path: str) -> str | None:
    """验证文件路径是否存在。

    Returns:
        如果路径无效，返回错误描述；否则返回 None。
    """
    if not file_path:
        return "file_path 不能为空"
    if not os.path.isfile(file_path):
        return f"文件不存在: {file_path}"
    return None


def validate_dir_path(dir_path: str) -> str | None:
    """验证目录路径是否存在。"""
    if not dir_path:
        return "目录路径不能为空"
    if not os.path.isdir(dir_path):
        return f"目录不存在: {dir_path}"
    return None


def validate_column_index(col_index: int, total_cols: int) -> str | None:
    """验证列索引是否在范围内（0-based）。

    Returns:
        如果索引越界返回错误描述；否则返回 None。
    """
    if col_index < 0 or col_index >= total_cols:
        return (
            f"列索引 {col_index} 超出范围，"
            f"当前工作表共 {total_cols} 列（0-{total_cols - 1}）。"
        )
    return None


def validate_column_indices(col_indices: list[int], total_cols: int) -> str | None:
    """验证一组列索引是否全部在范围内。"""
    for idx in col_indices:
        err = validate_column_index(idx, total_cols)
        if err:
            return err
    return None


def validate_plot_type(plot_type: str) -> str | None:
    """验证图表类型是否支持。"""
    valid = [e.value for e in PlotType]
    if plot_type not in valid:
        return f"不支持的图表类型 '{plot_type}'。支持的类型: {valid}"
    return None


def validate_designation(designation: str) -> str | None:
    """验证列角色是否支持。"""
    if designation == "YErr":
        return None
    valid = [e.value for e in ColumnDesignation]
    if designation not in valid:
        return f"不支持的列角色 '{designation}'。支持的值: {valid}"
    return None


def validate_scale_type(scale_type: str) -> str | None:
    """验证缩放类型是否支持。"""
    valid = [e.value for e in ScaleType]
    if scale_type not in valid:
        return f"不支持的缩放类型 '{scale_type}'。支持的类型: {valid}"
    return None


def validate_export_format(fmt: str) -> str | None:
    """验证导出格式是否支持。"""
    valid = [e.value for e in ExportFormat]
    if fmt not in valid:
        return f"不支持的导出格式 '{fmt}'。支持的格式: {valid}"
    return None


def normalize_y_cols(y_cols: int | list[int]) -> list[int]:
    """将 y_cols 统一为 list[int]，并执行显式类型校验。

    支持传入单个整数或整数列表；其他类型会抛出 ValueError。
    """
    if isinstance(y_cols, bool):
        raise ValueError("y_cols 必须是 int 或 list[int]，不能是布尔值")

    if isinstance(y_cols, int):
        return [y_cols]

    if not isinstance(y_cols, list):
        raise ValueError(f"y_cols 必须是 int 或 list[int]，当前类型: {type(y_cols).__name__}")

    for i, col in enumerate(y_cols):
        if isinstance(col, bool) or not isinstance(col, int):
            raise ValueError(
                f"y_cols[{i}] 必须是 int，当前值: {col!r}（{type(col).__name__}）"
            )

    return y_cols.copy()


def error_response_from_exception(exc: Exception) -> dict[str, Any]:
    """将 ToolError 异常转换为标准 error_response。

    Args:
        exc: ToolError 实例或其子类。

    Returns:
        统一格式的错误返回字典。
    """
    if isinstance(exc, ToolError):
        return error_response(
            message=str(exc),
            error_type=exc.error_type,
            target=exc.target,
            value=exc.value,
            hint=exc.hint,
        )
    return error_response(
        message=str(exc),
        error_type="internal_error",
        target="unknown",
        hint="发生未知错误，请检查 Origin 连接状态。",
    )

