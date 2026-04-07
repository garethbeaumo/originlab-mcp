"""高级逃生舱 tools

处理标准 tool 覆盖不到的少量特殊场景。

包含：
- execute_labtalk
"""

from __future__ import annotations

from originlab_mcp.utils.helpers import tool_error_handler
from originlab_mcp.utils.validators import error_response, success_response


def register_advanced_tools(mcp, manager) -> None:
    """注册高级逃生舱 tools 到 MCP Server。

    Args:
        mcp: FastMCP 实例。
        manager: OriginManager 实例（依赖注入）。
    """

    @mcp.tool()
    @tool_error_handler("LabTalk命令执行", "请检查 LabTalk 命令语法是否正确。")
    def execute_labtalk(command: str) -> dict:
        """Execute an arbitrary LabTalk command (high risk, use only when standard tools are insufficient).

        ⚠️ HIGH RISK TOOL: This tool directly executes LabTalk commands and may cause unpredictable effects on the Origin project.

        When to use: ONLY when all standard tools (import_csv, create_plot, set_axis_title, etc.) cannot fulfill the requirement, as a last resort.
        When not to use: Any operation achievable with standard tools should NOT use this tool.

        Examples:
        - execute_labtalk(command="window -a Graph1")
        - execute_labtalk(command="type -a \"hello\"")
        """
        if not command or not command.strip():
            return error_response(
                message="command 不能为空",
                error_type="invalid_input",
                target="command",
                hint="请提供要执行的 LabTalk 命令。",
            )

        # 基本安全检查：限制命令长度，防止超长注入
        if len(command) > 2000:
            return error_response(
                message="命令过长，最多允许 2000 个字符",
                error_type="invalid_input",
                target="command",
                hint="请缩短 LabTalk 命令。",
            )

        def _exec(op):
            result = op.lt_exec(command)
            return result

        result = manager.execute(_exec)

        return success_response(
            message="LabTalk 命令已执行。",
            data={
                "command": command,
                "result": str(result) if result is not None else None,
            },
            resource=manager.get_resource_context(),
            warnings=[
                "此操作通过逃生舱（execute_labtalk）执行，不属于标准调用路径。",
                "执行结果可能影响 Origin 内部状态，建议后续用 list_worksheets / list_graphs 确认当前状态。",
            ],
            next_suggestions=[
                "list_worksheets",
                "list_graphs",
                "get_origin_info",
            ],
        )
