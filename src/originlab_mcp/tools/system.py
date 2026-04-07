"""System tools

Manage Origin connection status, environment info,
and COM control lifecycle.
"""

from __future__ import annotations

from originlab_mcp.utils.validators import error_response, success_response


def register_system_tools(mcp, manager) -> None:
    """Register system tools to the MCP Server.

    Args:
        mcp: FastMCP instance.
        manager: OriginManager instance (dependency injection).
    """

    @mcp.tool()
    def get_origin_info() -> dict:
        """Return Origin connection status and environment info.

        When to use: To check if Origin is connected, view installation path,
        or get current project overview.
        When not to use: If Origin connection is known to be working and no
        environment info is needed.

        Default behavior:
        - Automatically connects if not yet connected

        Examples:
        - get_origin_info() -> Returns connection status, install path,
          current worksheet and graph count
        """
        try:
            manager.connect()
            info = manager.get_info()

            return success_response(
                message="Origin 已连接，环境信息获取成功。",
                data=info,
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "list_worksheets",
                    "import_csv",
                    "new_project",
                    "open_project",
                ],
            )
        except RuntimeError as e:
            return error_response(
                message=str(e),
                error_type="environment_error",
                target="origin_connection",
                hint=(
                    "请确认以下条件：\n"
                    "1. OriginLab 2021+ 已安装\n"
                    "2. 有效的 OriginLab 许可证\n"
                    "3. originpro 包已安装（pip install originpro）\n"
                    "4. 当前运行在 Windows 操作系统上"
                ),
            )
        except Exception as e:
            return error_response(
                message=f"获取 Origin 信息时发生未知错误: {e}",
                error_type="internal_error",
                target="origin_connection",
                hint="请检查 Origin 是否在运行，或尝试重启 MCP Server。",
            )

    @mcp.tool()
    def release_origin() -> dict:
        """Release Origin COM control so the user can freely operate or close Origin.

        When to use: After completing a batch of operations, call this to let
        the user interact with Origin manually. Origin stays running and the
        user can freely operate or even close it. Next tool call will
        automatically reconnect.

        When not to use: If there are more operations to follow, no need to
        call this. The connection auto-releases after 5 minutes of idle time.

        ⚠️ Difference from close_origin:
        - release_origin: Only releases COM control; Origin keeps running
        - close_origin: Releases control AND closes Origin application

        Examples:
        - release_origin() -> Releases control; user can freely operate Origin
        """
        released = manager.release()

        if released:
            return success_response(
                message=(
                    "已释放 Origin COM 控制权。\n"
                    "Origin 保持运行，你现在可以自由操作或关闭 Origin。\n"
                    "下次调用任何 tool 时会自动重新连接。"
                ),
                data={
                    "released": True,
                    "origin_still_running": True,
                },
                resource=manager.get_resource_context(),
                next_suggestions=["reconnect_origin", "get_origin_info"],
            )
        else:
            return success_response(
                message="Origin 当前未处于受控状态，无需释放。",
                data={"released": False},
                resource=manager.get_resource_context(),
            )

    @mcp.tool()
    def reconnect_origin() -> dict:
        """Reconnect to Origin COM after releasing control.

        When to use: After calling release_origin, use this to resume
        Origin operations. Note: most tools auto-reconnect, so explicit
        call is usually not needed.

        When not to use: If already connected.

        Examples:
        - reconnect_origin() -> Re-establishes COM connection
        """
        try:
            manager.connect()
            return success_response(
                message="已重新连接到 Origin。",
                data=manager.get_info(),
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "list_worksheets",
                    "list_graphs",
                ],
            )
        except RuntimeError as e:
            return error_response(
                message=str(e),
                error_type="environment_error",
                target="origin_connection",
                hint=(
                    "Origin 可能已被关闭。如果需要重新启动 Origin，"
                    "请手动打开 OriginLab，然后再调用此工具。"
                ),
            )
        except Exception as e:
            return error_response(
                message=f"重新连接时发生错误: {e}",
                error_type="internal_error",
                target="origin_connection",
                hint="请检查 Origin 是否在运行。",
            )

    @mcp.tool()
    def close_origin() -> dict:
        """Disconnect COM and close the Origin application.

        When to use: To shut down Origin completely. Disconnects COM first,
        then exits Origin.
        When not to use: To only release COM control without closing Origin,
        use release_origin instead.

        ⚠️ WARNING: This will close Origin and lose unsaved data.
        Call save_project first.

        Examples:
        - close_origin()
        """
        if not manager.is_connected:
            return success_response(
                message="Origin 未连接，无需关闭。",
                data={"was_connected": False},
                resource=manager.get_resource_context(),
            )

        try:
            manager.close_and_exit()
            return success_response(
                message="Origin 已正确关闭，COM 连接已释放。",
                data={"was_connected": True, "closed": True},
                resource=manager.get_resource_context(),
            )
        except Exception as e:
            return error_response(
                message=f"关闭 Origin 时出错: {e}",
                error_type="internal_error",
                target="origin_connection",
                hint="可尝试在任务管理器中手动结束 Origin64.exe 进程。",
            )
