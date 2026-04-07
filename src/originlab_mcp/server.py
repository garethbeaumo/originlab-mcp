"""
OriginLab MCP Server 主入口

负责：
- 创建 FastMCP 实例
- 创建唯一的 OriginManager 实例并注入到各 tool 模块
- 注册全部 tools（按能力域分组）
- 处理服务启动与关闭生命周期
"""

from __future__ import annotations

import atexit
import logging
from contextlib import asynccontextmanager

from mcp.server.fastmcp import FastMCP

from originlab_mcp.origin_manager import OriginManager
from originlab_mcp.tools.advanced import register_advanced_tools
from originlab_mcp.tools.analysis import register_analysis_tools
from originlab_mcp.tools.customize import register_customize_tools
from originlab_mcp.tools.data import register_data_tools
from originlab_mcp.tools.export import register_export_tools
from originlab_mcp.tools.plot import register_plot_tools
from originlab_mcp.tools.system import register_system_tools

# ---------------------------------------------------------------------------
# 日志配置
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
)
logger = logging.getLogger("originlab-mcp")

# ---------------------------------------------------------------------------
# 创建唯一的 OriginManager 实例（依赖注入源头）
# ---------------------------------------------------------------------------

_manager = OriginManager()

# ---------------------------------------------------------------------------
# 生命周期管理
# ---------------------------------------------------------------------------


@asynccontextmanager
async def _lifespan(server: FastMCP):
    """MCP Server 生命周期管理。"""
    logger.info("OriginLab MCP Server 已启动")
    try:
        yield {}
    finally:
        logger.info("MCP Server 正在关闭，清理资源...")
        _manager.shutdown()
        logger.info("资源清理完成")


# 同时注册 atexit 作为兜底，确保异常退出时也能清理
atexit.register(_manager.shutdown)

# ---------------------------------------------------------------------------
# 创建 MCP Server 实例
# ---------------------------------------------------------------------------

mcp = FastMCP(
    "originlab",
    instructions=(
        "OriginLab MCP Server - 通过 MCP 协议控制 OriginLab，"
        "支持数据导入、工作表管理、图表创建与定制、"
        "数据分析（线性/非线性拟合）、导出等功能。"
    ),
    lifespan=_lifespan,
)

# ---------------------------------------------------------------------------
# 注册所有 tools（通过依赖注入传递 manager）
# ---------------------------------------------------------------------------

register_system_tools(mcp, _manager)      # 系统状态类
register_data_tools(mcp, _manager)        # 数据类（导入、工作表操作、排序、公式、导出）
register_analysis_tools(mcp, _manager)    # 数据分析类（线性/非线性拟合）
register_plot_tools(mcp, _manager)        # 绘图类
register_customize_tools(mcp, _manager)   # 图表定制类（轴、颜色、符号、透明度、填充）
register_export_tools(mcp, _manager)      # 导出与项目管理类
register_advanced_tools(mcp, _manager)    # 高级逃生舱

logger.info("所有 tools 已注册完成")

# ---------------------------------------------------------------------------
# 入口函数
# ---------------------------------------------------------------------------


def main():
    """MCP Server 入口点，通过 pyproject.toml 的 [project.scripts] 调用。"""
    logger.info("OriginLab MCP Server 启动中...")
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
