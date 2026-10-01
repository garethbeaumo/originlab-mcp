"""Wire-level contracts through the real stdio server and SDK ClientSession."""

from __future__ import annotations

import asyncio
import json
import os
import sys
from datetime import timedelta
from pathlib import Path

import pytest
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


@pytest.fixture(scope="module")
def wire_results():
    async def collect():
        root = Path(__file__).resolve().parents[1]
        env = dict(os.environ)
        # Exercise this checkout, not a stale installed wheel or a .pth file.
        env["PYTHONPATH"] = os.pathsep.join([str(root / "src"), str(root)])
        env["ORIGINLAB_MCP_AUTOSAVE_INTERVAL"] = "off"
        env["ORIGINLAB_MCP_DISPATCH_TIMEOUT"] = "1"
        params = StdioServerParameters(
            command=sys.executable,
            args=["-m", "tests.protocol_server"],
            cwd=str(root),
            env=env,
        )
        async with (
            stdio_client(params) as (read, write),
            ClientSession(read, write, read_timeout_seconds=timedelta(seconds=10)) as session,
        ):
            await session.initialize()
            results = {}
            results["business"] = await session.call_tool("execute_labtalk", {"command": ""})
            results["exception"] = await session.call_tool("list_worksheets", {})
            results["timeout"] = await session.call_tool("execute_labtalk", {"command": "type 1;", "timeout": 0.01})
            results["schema"] = await session.call_tool("create_plot", {"x_col": "not-an-int", "y_cols": 1})
            results["unknown"] = await session.call_tool("nonexistent_tool", {})
            results["success"] = await session.call_tool("list_fit_functions", {})
            results["tools"] = await session.list_tools()
            return results

    return asyncio.run(collect())


@pytest.mark.parametrize("case,error_type,target", [
    ("business", "invalid_input", "command"),
    ("exception", "internal_error", "列出工作表"),
    ("timeout", "timeout", "dispatch_timeout"),
])
def test_tool_failures_set_protocol_error_flag(wire_results, case, error_type, target):
    result = wire_results[case]
    payload = json.loads(result.content[0].text)
    assert result.structuredContent is None
    assert payload["ok"] is False
    assert payload["error"]["type"] == error_type
    assert payload["error"]["target"] == target
    assert payload["error"]["hint"]
    assert result.isError is True, f"{case}: failed payload was marked as MCP success"


def test_business_failure_preserves_recovery_details(wire_results):
    error = json.loads(wire_results["business"].content[0].text)["error"]
    assert error["value"] is None
    assert "list_worksheets" in error["suggested_alternatives"]


@pytest.mark.parametrize("case", ["schema", "unknown"])
def test_sdk_rejections_remain_protocol_errors(wire_results, case):
    assert wire_results[case].isError is True


def test_success_and_discovery_remain_compatible(wire_results):
    result = wire_results["success"]
    assert result.isError is False
    assert json.loads(result.content[0].text)["ok"] is True
    assert result.structuredContent is None
    tools = wire_results["tools"].tools
    assert len(tools) == 67
    for tool in tools:
        assert tool.annotations is not None
        assert tool.inputSchema["type"] == "object"
        assert tool.outputSchema is None
