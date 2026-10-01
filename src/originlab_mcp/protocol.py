"""Map the project's JSON failure envelope to the MCP tool-error flag."""

from __future__ import annotations

import json
from typing import Any

from mcp.server.fastmcp import FastMCP
from mcp.types import CallToolResult, TextContent


class OriginFastMCP(FastMCP):
    """Keep legacy tool payloads while exposing their failures to MCP clients."""

    async def call_tool(self, name: str, arguments: dict[str, Any]) -> Any:
        # SDK 1.26 converts our unparameterized `dict` tool results to one JSON
        # TextContent block. Use the public handler hook; do not replace its
        # argument validation, error handling, or tool metadata.
        result = await super().call_tool(name, arguments)
        if (
            isinstance(result, list)
            and len(result) == 1
            and isinstance(result[0], TextContent)
        ):
            try:
                payload = json.loads(result[0].text)
            except json.JSONDecodeError:
                return result
            if isinstance(payload, dict) and payload.get("ok") is False:
                # Preserve the SDK-produced text byte for byte, including
                # error.type/target/value/hint/suggested_alternatives.
                return CallToolResult(content=result, isError=True)
        return result
