"""高级逃生舱 tools

处理标准 tool 覆盖不到的少量特殊场景。

包含：
- execute_labtalk
- get_labtalk_variable
"""

from __future__ import annotations

import re

from originlab_mcp.utils.helpers import tool_error_handler
from originlab_mcp.utils.validators import error_response, success_response

_LABTALK_VAR_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*\$?$")


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

    @mcp.tool()
    @tool_error_handler("读取 LabTalk 变量", "请检查变量名是否存在。")
    def get_labtalk_variable(name: str) -> dict:
        """Read a LabTalk variable value without executing arbitrary commands.

        When to use: To inspect a LabTalk numeric or string variable.
        When not to use: To modify Origin state or execute scripts, use execute_labtalk.

        Parameter notes:
        - String variables must end with '$', e.g. "fname$"
        - Numeric variables should not include '$', e.g. "pi"

        Examples:
        - get_labtalk_variable(name="fname$")
        - get_labtalk_variable(name="pi")
        """
        normalized_name = name.strip()
        if not normalized_name:
            return error_response(
                message="name 不能为空",
                error_type="invalid_input",
                target="name",
                hint="请提供要读取的 LabTalk 变量名。",
            )
        if len(normalized_name) > 128:
            return error_response(
                message="变量名过长，最多允许 128 个字符",
                error_type="invalid_input",
                target="name",
                value=name,
                hint="请提供合法的 LabTalk 变量名。",
            )
        if not _LABTALK_VAR_RE.match(normalized_name):
            return error_response(
                message=f"变量名 '{name}' 包含非法字符",
                error_type="invalid_input",
                target="name",
                value=name,
                hint="变量名仅允许字母、数字、下划线，字符串变量可带结尾 '$'。",
            )

        def _read(op):
            if normalized_name.endswith("$"):
                return {
                    "name": normalized_name,
                    "value": op.get_lt_str(normalized_name),
                    "value_type": "string",
                }
            return {
                "name": normalized_name,
                "value": op.lt_float(normalized_name),
                "value_type": "numeric",
            }

        result = manager.execute(_read)

        return success_response(
            message=f"已读取 LabTalk 变量 '{normalized_name}'。",
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "get_origin_info",
                "execute_labtalk",
            ],
        )
