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
def wire_results(tmp_path_factory):
    async def collect():
        output_dir = tmp_path_factory.mktemp("wire-export")
        root = Path(__file__).resolve().parents[1]
        env = dict(os.environ)
        # Exercise this checkout, not a stale installed wheel or a .pth file.
        env["PYTHONPATH"] = os.pathsep.join([str(root / "src"), str(root)])
        env["ORIGINLAB_MCP_AUTOSAVE_INTERVAL"] = "off"
        env["ORIGINLAB_MCP_AUTOSAVE"] = "1"
        env["ORIGINLAB_MCP_AUTOSAVE_REQUIRED"] = "1"
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
            results["dynamic_confirm"] = await session.call_tool("execute_labtalk", {"command": "if ((1)) %(cmd$);"})
            results["for_confirm"] = await session.call_tool("execute_labtalk", {"command": "for(ii=1; ii<=1; ii++) %(cmd$);"})
            results["case_confirm"] = await session.call_tool("execute_labtalk", {"command": "switch(1) { case 1: %(cmd$); }"})
            results["default_confirm"] = await session.call_tool("execute_labtalk", {"command": "switch(1) { default: %A; }"})
            results["group_confirm"] = await session.call_tool("execute_labtalk", {"command": "if (1) (%(cmd$);)"})
            results["mixed_autosave"] = await session.call_tool("execute_labtalk", {"command": 'python -s "print(1)"; doc -n;', "confirm": True})
            results["exception"] = await session.call_tool("list_worksheets", {})
            results["timeout"] = await session.call_tool("execute_labtalk", {"command": "type 1;", "timeout": 0.01})
            results["schema"] = await session.call_tool("create_plot", {"x_col": "not-an-int", "y_cols": 1})
            results["unknown"] = await session.call_tool("nonexistent_tool", {})
            results["success"] = await session.call_tool("list_fit_functions", {})
            for name, kwargs in (("free", {}), ("intercept", {"fix_intercept": 0}), ("slope", {"fix_slope": 0})):
                results[name] = await session.call_tool("linear_fit", {"sheet_name": "FitData", "x_col": 0, "y_col": 1, **kwargs})
            for tool in ("linear_fit", "nonlinear_fit"):
                for stage in ("result", "report"):
                    kwargs = {"sheet_name": "ResultFailure" if stage == "result" else "ReportFailure", "x_col": 0, "y_col": 1}
                    if tool == "linear_fit":
                        kwargs["confidence_band"] = stage == "report"
                    else:
                        kwargs.update(function_name="Gauss", generate_report=stage == "report")
                    results[f"{tool}_{stage}"] = await session.call_tool(tool, kwargs)
            results["plot_partial"] = await session.call_tool("create_plot", {"sheet_name": "PlotPartial", "x_col": 0, "y_cols": [1, 2]})
            results["export_failure"] = await session.call_tool("export_graph", {"graph_name": "ExportFailure", "output_path": str(output_dir / "single.png")})
            results["export_partial"] = await session.call_tool("export_all_graphs", {"output_dir": str(output_dir)})
            results["tools"] = await session.list_tools()
            return results

    return asyncio.run(collect())


@pytest.mark.parametrize("case,error_type,target", [
    ("business", "invalid_input", "command"),
    ("exception", "internal_error", "列出工作表"),
    ("timeout", "timeout", "dispatch_timeout"),
    ("dynamic_confirm", "invalid_input", "confirm"),
    ("for_confirm", "invalid_input", "confirm"),
    ("case_confirm", "invalid_input", "confirm"),
    ("default_confirm", "invalid_input", "confirm"),
    ("group_confirm", "invalid_input", "confirm"),
    ("mixed_autosave", "invalid_input", "autosave"),
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


@pytest.mark.parametrize("case,slope,intercept", [("free", 2, 1), ("intercept", 17 / 7, 0), ("slope", 0, 5)])
def test_linear_acceptance_over_stdio(wire_results, case, slope, intercept):
    result = wire_results[case]
    assert result.isError is False
    payload = json.loads(result.content[0].text)
    assert payload["ok"]
    params = payload["data"]["parameters"]
    assert params["Slope"]["value"] == pytest.approx(slope)
    assert params["Intercept"]["value"] == pytest.approx(intercept)


@pytest.mark.parametrize("case", ["linear_fit_result", "linear_fit_report", "nonlinear_fit_result", "nonlinear_fit_report", "plot_partial", "export_failure", "export_partial"])
def test_result_failures_over_stdio(wire_results, case):
    result = wire_results[case]
    assert result.isError is True
    payload = json.loads(result.content[0].text)
    assert payload["ok"] is False
    if case == "plot_partial":
        assert payload["data"]["curve_count"] == 1
        assert payload["data"]["graph_name"]
        assert payload["data"]["failed_step"] == "add_plot:1"


def test_export_partial_over_stdio(wire_results):
    payload = json.loads(wire_results["export_partial"].content[0].text)
    assert payload["data"]["count"] == 1
    assert payload["data"]["failed_count"] == 1
    assert payload["data"]["files"][0]["output_path"].endswith(".png")
