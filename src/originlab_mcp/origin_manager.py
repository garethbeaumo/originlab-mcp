"""
Origin COM 连接管理器

负责：
- 管理 Origin COM 连接（懒连接 + 自动重连）
- 线程安全的 COM 调用（threading.Lock）
- 活动对象（工作表 / 图表）追踪
- Server 关闭时的资源清理
"""

from __future__ import annotations

import logging
import threading
from typing import Any, Callable, TypeVar

logger = logging.getLogger(__name__)

T = TypeVar("T")


class OriginManager:
    """Origin COM 连接的统一管理入口。

    通过依赖注入方式传递给各 tool 注册函数，
    保证整个 MCP Server 生命周期内只有一个管理器实例。
    所有 COM 操作通过 `execute` 方法串行执行。
    """

    def __init__(self) -> None:
        self._com_lock = threading.Lock()
        self._connected = False
        self._op = None  # originpro 模块引用

        # 活动对象追踪
        self._active_worksheet: str | None = None
        self._active_graph: str | None = None

        logger.info("OriginManager 初始化完成（尚未连接）")

    # -----------------------------------------------------------------
    # 连接管理
    # -----------------------------------------------------------------

    def connect(self) -> None:
        """建立与 Origin 的 COM 连接。

        首次调用时导入 originpro 并初始化连接。
        如果已连接则跳过。
        """
        if self._connected and self._op is not None:
            return

        try:
            import originpro as op

            self._op = op
            # 确保 Origin 窗口可见
            self._op.set_show(True)
            self._connected = True
            logger.info("已连接到 Origin")
        except ImportError as e:
            self._connected = False
            raise RuntimeError(
                "无法导入 originpro。请确认已安装 originpro 包且 OriginLab 已正确安装。"
            ) from e
        except Exception as e:
            self._connected = False
            raise RuntimeError(f"连接 Origin 失败: {e}") from e

    def disconnect(self) -> None:
        """释放 Origin COM 连接。"""
        if self._op is not None:
            try:
                self._op.exit()
                logger.info("已断开 Origin 连接")
            except Exception as e:
                logger.warning("断开连接时出错: %s", e)
            finally:
                self._connected = False
                self._op = None

    def ensure_connected(self) -> None:
        """检查连接有效性，必要时重连。"""
        if not self._connected or self._op is None:
            logger.info("连接无效，尝试重新连接...")
            self._connected = False
            self._op = None
            self.connect()
            return

        # 尝试一个轻量操作来验证连接是否仍然有效
        try:
            self._op.path("e")  # 获取 Origin 安装路径
        except Exception:
            logger.warning("Origin 连接已断开，尝试重新连接...")
            self._connected = False
            self._op = None
            self.connect()

    @property
    def is_connected(self) -> bool:
        """返回当前连接状态。"""
        return self._connected and self._op is not None

    @property
    def op(self):
        """获取 originpro 模块引用。

        调用前必须确保已连接。
        """
        if not self._connected or self._op is None:
            raise RuntimeError("Origin 尚未连接，请先调用 connect()")
        return self._op

    # -----------------------------------------------------------------
    # 线程安全执行
    # -----------------------------------------------------------------

    def execute(self, func: Callable[..., T], *args: Any, **kwargs: Any) -> T:
        """在 COM 锁保护下执行函数。

        自动确保连接有效，并在 COM 锁内串行执行。

        Parameters
        ----------
        func : Callable
            要执行的函数，第一个参数会接收 originpro 模块。
        *args, **kwargs
            传给 func 的额外参数。
        """
        with self._com_lock:
            self.ensure_connected()
            return func(self._op, *args, **kwargs)

    # -----------------------------------------------------------------
    # 活动对象追踪
    # -----------------------------------------------------------------

    @property
    def active_worksheet(self) -> str | None:
        """当前活动工作表名。"""
        return self._active_worksheet

    @active_worksheet.setter
    def active_worksheet(self, name: str | None) -> None:
        self._active_worksheet = name
        if name:
            logger.debug("活动工作表 -> %s", name)

    @property
    def active_graph(self) -> str | None:
        """当前活动图表名。"""
        return self._active_graph

    @active_graph.setter
    def active_graph(self, name: str | None) -> None:
        self._active_graph = name
        if name:
            logger.debug("活动图表 -> %s", name)

    def get_resource_context(self) -> dict[str, str | None]:
        """返回当前活动对象上下文，用于填充返回结构的 resource 字段。"""
        return {
            "active_worksheet": self._active_worksheet,
            "active_graph": self._active_graph,
        }

    # -----------------------------------------------------------------
    # 环境信息
    # -----------------------------------------------------------------

    def get_info(self) -> dict[str, Any]:
        """获取 Origin 环境信息。

        Returns
        -------
        dict
            包含连接状态、版本、路径等信息。
        """
        info: dict[str, Any] = {
            "connected": self.is_connected,
        }

        if self.is_connected:
            try:
                op = self._op
                info["exe_path"] = op.path("e")
                info["user_path"] = op.path("u")
                info["active_worksheet"] = self._active_worksheet
                info["active_graph"] = self._active_graph

                # 尝试获取当前项目中的工作表和图表数量
                try:
                    worksheets = []
                    for book in op.pages("Book"):
                        for sheet in book:
                            worksheets.append(f"[{book.name}]{sheet.name}")
                    info["worksheet_count"] = len(worksheets)
                except Exception:
                    info["worksheet_count"] = None

                try:
                    graphs = [g.name for g in op.pages("Graph")]
                    info["graph_count"] = len(graphs)
                except Exception:
                    info["graph_count"] = None

            except Exception as e:
                info["info_error"] = str(e)

        return info

    # -----------------------------------------------------------------
    # 生命周期
    # -----------------------------------------------------------------

    def shutdown(self) -> None:
        """MCP Server 关闭时调用，清理资源。"""
        logger.info("OriginManager 正在关闭...")
        self.disconnect()
        self._active_worksheet = None
        self._active_graph = None
        logger.info("OriginManager 已关闭")
