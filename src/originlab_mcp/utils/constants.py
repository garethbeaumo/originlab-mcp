"""
常量定义模块

维护图类型、错误类型、列角色、默认值等枚举和常量。
"""

from dataclasses import dataclass
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
    Y_ERROR = "E"
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
    """坐标轴缩放类型（完整覆盖 originpro API 支持的 9 种）"""
    LINEAR = "linear"
    LOG10 = "log10"
    LOG = "log"          # log10 的别名，保持向后兼容
    PROBABILITY = "probability"
    PROBIT = "probit"
    RECIPROCAL = "reciprocal"
    OFFSET_RECIPROCAL = "offset_reciprocal"
    LOGIT = "logit"
    LN = "ln"
    LOG2 = "log2"


# originpro 中 scale type 的字符串映射
SCALE_TYPE_TO_ORIGIN: dict[str, str] = {
    ScaleType.LINEAR.value: "linear",
    ScaleType.LOG10.value: "log10",
    ScaleType.LOG.value: "log10",
    ScaleType.PROBABILITY.value: "probability",
    ScaleType.PROBIT.value: "probit",
    ScaleType.RECIPROCAL.value: "reciprocal",
    ScaleType.OFFSET_RECIPROCAL.value: "offset_reciprocal",
    ScaleType.LOGIT.value: "logit",
    ScaleType.LN.value: "ln",
    ScaleType.LOG2.value: "log2",
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


# ---------------------------------------------------------------------------
# add_plot 子类型映射
# ---------------------------------------------------------------------------

class AddPlotType(str, Enum):
    """GLayer.add_plot() 的 type 参数"""
    LINE = "l"
    SCATTER = "s"
    LINE_SYMBOL = "y"
    COLUMN = "c"
    AUTO = "?"


# ---------------------------------------------------------------------------
# 常用非线性拟合函数
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class FitFunctionInfo:
    """拟合函数的元信息。"""

    params: tuple[str, ...]
    description: str


COMMON_FIT_FUNCTIONS: dict[str, FitFunctionInfo] = {
    "Gauss": FitFunctionInfo(
        params=("y0", "xc", "w", "A"),
        description="高斯函数 - 适用于峰形分布数据",
    ),
    "Lorentz": FitFunctionInfo(
        params=("y0", "xc", "w", "A"),
        description="洛伦兹函数 - 适用于共振峰拟合",
    ),
    "ExpDec1": FitFunctionInfo(
        params=("y0", "A1", "t1"),
        description="单指数衰减 - 适用于一阶反应动力学",
    ),
    "ExpDec2": FitFunctionInfo(
        params=("y0", "A1", "t1", "A2", "t2"),
        description="双指数衰减 - 适用于混合衰减过程",
    ),
    "ExpGrow1": FitFunctionInfo(
        params=("y0", "A1", "t1"),
        description="单指数增长 - 适用于增长过程",
    ),
    "Boltzmann": FitFunctionInfo(
        params=("A1", "A2", "x0", "dx"),
        description="玻尔兹曼函数 - 适用于 S 型曲线拟合",
    ),
    "Logistic": FitFunctionInfo(
        params=("A1", "A2", "x0", "p"),
        description="逻辑斯谛函数 - 适用于增长饱和曲线",
    ),
    "Hill": FitFunctionInfo(
        params=("Vmax", "k", "n"),
        description="Hill 方程 - 适用于剂量-反应曲线",
    ),
    "Polynomial": FitFunctionInfo(
        params=(),
        description="多项式拟合 - 通用拟合",
    ),
    "Line": FitFunctionInfo(
        params=("A", "B"),
        description="线性函数 y = A + B*x",
    ),
    "Plane": FitFunctionInfo(
        params=("z0", "a", "b"),
        description="平面拟合 z = z0 + a*x + b*y",
    ),
}
