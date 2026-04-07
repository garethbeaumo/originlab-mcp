"""系统状态类 tools

让 AI 能判断当前连接状态和运行环境，
并能主动管理 Origin COM 控制权的获取与释放。
"""

from __future__ import annotations

from originlab_mcp.utils.validators import error_response, success_response


def register_system_tools(mcp, manager) -> None:
    """注册系统状态类 tools 到 MCP Server。

    Args:
        mcp: FastMCP 实例。
        manager: OriginManager 实例（依赖注入）。
    """

    @mcp.tool()
    def get_origin_info() -> dict:
        """返回 Origin 连接状态和基础环境信息。

        何时使用：需要确认 Origin 是否已连接、查看安装路径或当前项目概况时使用。
        何时不用：已知 Origin 连接正常且不需要环境信息时无需调用。

        默认行为：
        - 自动建立连接（如尚未连接）

        示例：
        - get_origin_info() -> 返回连接状态、安装路径、当前工作表和图表数量
        """
        # 注意：此函数不使用 @tool_error_handler，
        # 因为它需要特殊处理 RuntimeError（连接失败）以给出安装提示。
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
        """释放 Origin COM 控制权，使用户可自由操作或关闭 Origin。

        何时使用：当一批操作完成后，用户希望手动操作 Origin 时调用此工具。
        释放后 Origin 保持运行，用户可自由操作甚至关闭 Origin。
        下次调用其他 tool 时会自动重新连接。

        何时不用：如果还有后续操作要执行，不需要调用此工具。
        连接空闲超过 5 分钟会自动释放。

        ⚠️ 与 close_origin 的区别：
        - release_origin：仅释放控制权，Origin 保持运行
        - close_origin：释放控制权并关闭 Origin 应用

        示例：
        - release_origin() -> 释放控制权，用户可自由操作 Origin
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
        """重新连接 Origin COM。

        何时使用：在调用 release_origin 释放控制权后，
        需要恢复对 Origin 的操作时使用。
        注意：大多数 tool 会自动重连，通常不需要显式调用此工具。

        何时不用：已连接状态下无需调用。

        示例：
        - reconnect_origin() -> 重新建立 COM 连接
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
        """正确断开 COM 连接并关闭 Origin。

        何时使用：需要关闭 Origin 应用程序时使用。会先释放 COM 连接，再退出 Origin。
        何时不用：仅需释放控制权而不关闭 Origin 时请使用 release_origin。

        ⚠️ 注意：此操作会关闭 Origin 并丢失未保存的数据，请先调用 save_project 保存。

        示例：
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
