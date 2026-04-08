"""
Origin COM 连接管理器

负责：
- 管理 Origin COM 连接（懒连接 + 自动重连）
- 线程安全的 COM 调用（threading.Lock）
- 空闲超时自动释放 COM 控制权
- 活动对象（工作表 / 图表）追踪
- Server 关闭时的资源清理
"""

from __future__ import annotations

import importlib
import logging
import threading
from collections.abc import Callable
from typing import Any, TypeVar, cast

from originlab_mcp.types import OriginProProtocol

logger = logging.getLogger(__name__)

T = TypeVar("T")

# 默认空闲超时秒数（5 分钟）
DEFAULT_IDLE_TIMEOUT = 300


class OriginManager:
    """Origin COM 连接的统一管理入口。

    通过依赖注入方式传递给各 tool 注册函数，
    保证整个 MCP Server 生命周期内只有一个管理器实例。
    所有 COM 操作通过 `execute` 方法串行执行。

    生命周期策略（按需持有，自动释放）：
    - 首次 tool 调用时自动连接
    - 连续操作期间保持连接（不再每次 detach）
    - 空闲超时后自动 detach，释放 Origin 控制权
    - AI 可通过 release_origin 主动释放控制权
    - 释放后再有 tool 调用时自动重连
    """

    def __init__(self, idle_timeout: int = DEFAULT_IDLE_TIMEOUT) -> None:
        self._com_lock = threading.Lock()
        self._connected = False
        self._op: OriginProProtocol | None = None

        # 空闲超时机制
        self._idle_timeout = idle_timeout
        self._idle_timer: threading.Timer | None = None
        self._timer_lock = threading.Lock()  # 保护 timer 的创建/取消

        # 活动对象追踪
        self._active_worksheet: str | None = None
        self._active_graph: str | None = None

        logger.info(
            "OriginManager 初始化完成（空闲超时=%ds）", self._idle_timeout
        )

    # -----------------------------------------------------------------
    # 空闲超时管理
    # -----------------------------------------------------------------

    def _reset_idle_timer(self) -> None:
        """重置空闲计时器。每次 execute() 调用后重新开始计时。"""
        with self._timer_lock:
            # 取消现有的计时器
            if self._idle_timer is not None:
                self._idle_timer.cancel()
                self._idle_timer = None

            # 仅在已连接状态下启动新计时器
            if self._connected and self._idle_timeout > 0:
                self._idle_timer = threading.Timer(
                    self._idle_timeout, self._on_idle_timeout
                )
                self._idle_timer.daemon = True  # 不阻止进程退出
                self._idle_timer.start()
                logger.debug("空闲计时器已重置（%ds 后释放）", self._idle_timeout)

    def _cancel_idle_timer(self) -> None:
        """取消空闲计时器。"""
        with self._timer_lock:
            if self._idle_timer is not None:
                self._idle_timer.cancel()
                self._idle_timer = None

    def _on_idle_timeout(self) -> None:
        """空闲超时回调：自动 detach 释放 Origin 控制权。"""
        logger.info("空闲超时，自动释放 Origin COM 控制权...")
        with self._com_lock:
            self._do_detach()

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
            module = importlib.import_module("originpro")
            op = cast(OriginProProtocol, module)
            self._op = op
            # 确保 Origin 窗口可见
            op.set_show(True)
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

    def _do_detach(self) -> None:
        """内部方法：执行 COM detach 并更新状态。

        释放 COM 控制权但不关闭 Origin，允许用户自由操作。
        此方法不获取锁，调用者必须已持有 _com_lock。
        """
        op = self._op
        if op is not None:
            try:
                op.detach()
                logger.info("已释放 Origin COM 控制权（Origin 保持运行）")
            except Exception as e:
                logger.debug("detach 时出错（可忽略）: %s", e)
            finally:
                self._connected = False

    def release(self) -> bool:
        """显式释放 Origin COM 控制权。

        释放后 Origin 保持运行，用户可自由操作甚至关闭 Origin。
        下次 tool 调用时会自动重连。

        Returns
        -------
        bool
            是否成功释放（如果本来就未连接则返回 False）。
        """
        self._cancel_idle_timer()
        with self._com_lock:
            if not self._connected or self._op is None:
                return False
            self._do_detach()
            return True

    def disconnect(self) -> None:
        """释放 Origin COM 连接（不关闭 Origin）。

        与 release() 的区别：disconnect() 会清除 _op 引用，
        用于 shutdown 场景。release() 保留 _op 以便自动重连。
        """
        self._cancel_idle_timer()
        with self._com_lock:
            op = self._op
            if op is not None:
                try:
                    op.detach()
                    logger.info("已断开 Origin 连接（Origin 保持运行）")
                except Exception as e:
                    logger.warning("断开连接时出错: %s", e)
                finally:
                    self._connected = False
                    self._op = None

    def close_and_exit(self) -> None:
        """关闭 Origin 应用程序并释放连接。

        先确保 COM 连接有效，再调用 op.exit() 正常退出 Origin。
        """
        self._cancel_idle_timer()
        if self._op is not None:
            try:
                self.ensure_connected()
                self.op.exit()
                logger.info("已关闭 Origin")
            except Exception as e:
                logger.warning("关闭 Origin 时出错: %s", e)
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
    def op(self) -> OriginProProtocol:
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
        不再在每次操作后 detach，而是重置空闲计时器。
        当空闲超时后自动释放 COM 控制权。

        Parameters
        ----------
        func : Callable
            要执行的函数，第一个参数会接收 originpro 模块。
        *args, **kwargs
            传给 func 的额外参数。
        """
        # 先取消可能正在运行的空闲计时器（防止执行期间触发 detach）
        self._cancel_idle_timer()

        with self._com_lock:
            self.ensure_connected()
            op = self.op
            try:
                return func(op, *args, **kwargs)
            finally:
                # 操作完成后启动空闲计时器（而不是立即 detach）
                self._reset_idle_timer()

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
            "idle_timeout": self._idle_timeout,
        }

        if self.is_connected:
            try:
                op = self.op
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
        self._cancel_idle_timer()
        self.disconnect()
        self._active_worksheet = None
        self._active_graph = None
        logger.info("OriginManager 已关闭")
