"""
高级逃生舱 tools

处理标准 tool 覆盖不到的少量特殊场景。

包含：
- execute_labtalk
"""

from __future__ import annotations

from originlab_mcp.origin_manager import OriginManager
from originlab_mcp.utils.validators import error_response, success_response


def register_advanced_tools(mcp) -> None:
    """注册高级逃生舱 tools 到 MCP Server。"""

    manager = OriginManager()

    @mcp.tool()
    def execute_labtalk(command: str) -> dict:
        """执行任意 LabTalk 命令（高风险，仅在标准 tool 不足时使用）。

        ⚠️ 高风险工具：此工具直接执行 LabTalk 命令，可能对 Origin 项目产生不可预测的影响。

        何时使用：仅在所有标准 tool（import_csv, create_plot, set_axis_title 等）都无法满足需求时，作为最后手段使用。
        何时不用：任何可以用标准 tool 完成的操作，都不应使用此工具。

        示例：
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

        try:
            def _exec(op):
                result = op.lt_exec(command)
                return result

            result = manager.execute(_exec)

            return success_response(
                message=f"LabTalk 命令已执行。",
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
        except Exception as e:
            return error_response(
                message=f"LabTalk 命令执行失败: {e}",
                error_type="internal_error",
                target="command",
                value=command,
                hint="请检查 LabTalk 命令语法是否正确。",
            )
