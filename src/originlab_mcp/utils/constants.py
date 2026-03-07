"""
常量定义模块

维护图类型、错误类型、列角色、默认值等枚举和常量。
"""

from enum import Enum


# ---------------------------------------------------------------------------
# 图表类型
# ---------------------------------------------------------------------------

class PlotType(str, Enum):
    """支持的图表类型"""
    LINE = "line"
    SCATTER = "scatter"
    LINE_SYMBOL = "line_symbol"
    COLUMN = "column"
    AREA = "area"
    AUTO = "auto"


# PlotType -> originpro 模板名映射
PLOT_TYPE_TO_TEMPLATE: dict[str, str] = {
    PlotType.LINE: "line",
    PlotType.SCATTER: "scatter",
    PlotType.LINE_SYMBOL: "linesymb",
    PlotType.COLUMN: "column",
    PlotType.AREA: "area",
    PlotType.AUTO: "line",
}


# ---------------------------------------------------------------------------
# 列角色
# ---------------------------------------------------------------------------

class ColumnDesignation(str, Enum):
    """列角色类型"""
    X = "X"
    Y = "Y"
    Z = "Z"
    Y_ERROR = "YErr"
    LABEL = "L"
    DISREGARD = "N"


# ---------------------------------------------------------------------------
# 轴标识
# ---------------------------------------------------------------------------

class AxisId(str, Enum):
    """轴标识"""
    X = "x"
    Y = "y"
    Z = "z"


# ---------------------------------------------------------------------------
# 坐标轴缩放类型
# ---------------------------------------------------------------------------

class ScaleType(str, Enum):
    """坐标轴缩放类型"""
    LINEAR = "linear"
    LOG = "log"
    LN = "ln"


# originpro 中 scale type 的数值映射
SCALE_TYPE_TO_INT: dict[str, int] = {
    ScaleType.LINEAR: 0,
    ScaleType.LOG: 1,
    ScaleType.LN: 2,
}


# ---------------------------------------------------------------------------
# 导出格式
# ---------------------------------------------------------------------------

class ExportFormat(str, Enum):
    """支持的导出图片格式"""
    PNG = "png"
    SVG = "svg"
    PDF = "pdf"


# ---------------------------------------------------------------------------
# 错误类型
# ---------------------------------------------------------------------------

class ErrorType(str, Enum):
    """统一错误类别"""
    NOT_FOUND = "not_found"
    INVALID_INPUT = "invalid_input"
    UNSUPPORTED = "unsupported"
    CONFLICT = "conflict"
    ENVIRONMENT_ERROR = "environment_error"
    INTERNAL_ERROR = "internal_error"


# ---------------------------------------------------------------------------
# 默认值
# ---------------------------------------------------------------------------

DEFAULT_PLOT_TYPE = PlotType.LINE
DEFAULT_SEPARATOR = ","
DEFAULT_HAS_HEADER = True
DEFAULT_EXPORT_FORMAT = ExportFormat.PNG
DEFAULT_EXPORT_WIDTH = 800
DEFAULT_MAX_PREVIEW_ROWS = 20
