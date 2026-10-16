"""Verify real MCP stdio; opt in to a disposable Windows/Origin workflow.

Portable: uv run python scripts/verify_mcp.py
Windows:  uv run python scripts/verify_mcp.py --live --allow-project-reset --output-dir C:\\temp\\origin-smoke
"""

from __future__ import annotations

import argparse
import json
import math
import os
import platform
import sys
from pathlib import Path

import anyio
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


async def verify(args: argparse.Namespace) -> None:
    env = dict(os.environ)
    if args.live:
        # Only this child process receives the smoke-test policies.
        env["ORIGINLAB_MCP_ALLOWED_ROOTS"] = str(args.output_dir)
        env["ORIGINLAB_MCP_AUTOSAVE"] = "off"
        env["ORIGINLAB_MCP_AUTOSAVE_INTERVAL"] = "off"
        env["ORIGINLAB_MCP_DISPATCH_TIMEOUT"] = "30"
    params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "tests.protocol_server" if args.fake_workflow else "originlab_mcp.server"],
        env=env,
    )
    async with (
        stdio_client(params) as (read, write),
        ClientSession(read, write) as session,
    ):
        await session.initialize()
        tools = await session.list_tools()
        prompts = await session.list_prompts()
        resources = await session.list_resources()
        names = {tool.name for tool in tools.tools}
        if not {"originlab_doctor", "save_project", "list_worksheets", "list_graphs"} <= names:
            raise RuntimeError("Required tools missing from MCP discovery")
        if len(prompts.prompts) != 5 or len(resources.resources) != 3:
            raise RuntimeError("Unexpected prompt/resource discovery")
        print(json.dumps({"tools": len(names), "prompts": len(prompts.prompts), "resources": len(resources.resources)}))

        async def call(name: str, **kwargs):
            result = await session.call_tool(name, kwargs)
            text = next((item.text for item in result.content if item.type == "text"), "{}")
            payload = json.loads(text)
            if result.isError or not payload.get("ok"):
                raise RuntimeError(f"{name}: {payload}")
            print(json.dumps({"tool": name, "result": payload}, ensure_ascii=False))
            return payload["data"]

        await call("originlab_doctor", ping_origin=args.live)
        invalid = await session.call_tool("execute_labtalk", {"command": ""})
        failure = json.loads(invalid.content[0].text)
        if not invalid.isError or failure.get("ok") is not False:
            raise RuntimeError("Tool failure was not reported as an MCP error")
        print("PASS: business failure retains its JSON payload and sets isError=true.")
        if not args.live and not args.fake_workflow:
            print("PASS: portable MCP discovery and diagnostic call; no Origin operations run.")
            return

        if args.live:
            await call("get_origin_info")
            await call("new_project")
        imported = await call("import_data_from_text", data="X,Y\n1,3\n2,5\n3,7", has_header=True)
        for fixed, slope, intercept in (({}, 2, 1), ({"fix_intercept": 0}, 17 / 7, 0), ({"fix_slope": 0}, 0, 5)):
            fit = await call("linear_fit", x_col=0, y_col=1, **fixed)
            for name, expected in (("Slope", slope), ("Intercept", intercept)):
                actual = fit["parameters"][name]["value"]
                if not math.isclose(actual, expected, rel_tol=1e-7, abs_tol=1e-7):
                    raise RuntimeError(f"{fixed}: {name} expected {expected}, got {actual}")
        print("PASS: linear fit acceptance (free, zero intercept, zero slope).")
        if args.fake_workflow:
            for tool in ("linear_fit", "nonlinear_fit"):
                for stage in ("result", "report"):
                    kwargs = {"sheet_name": "ResultFailure" if stage == "result" else "ReportFailure", "x_col": 0, "y_col": 1}
                    if tool == "linear_fit":
                        kwargs["confidence_band"] = stage == "report"
                    else:
                        kwargs.update(function_name="Gauss", generate_report=stage == "report")
                    failed = await session.call_tool(tool, kwargs)
                    payload = json.loads(failed.content[0].text)
                    if not failed.isError or payload.get("ok") is not False:
                        raise RuntimeError(f"{tool} {stage} failure was reported as success")
            print("PASS: report/result failure injection over real stdio; fake Origin only.")
            return
        sheets = await call("list_worksheets")
        if sheets["count"] < 1:
            raise RuntimeError(f"Worksheet discovery failed after import: {imported}")
        await call("create_plot", x_col=0, y_cols=1, plot_type="scatter")
        graphs = await call("list_graphs")
        if graphs["count"] != 1:
            raise RuntimeError(f"Expected one graph, got {graphs}")
        snapshot = await call("read_origin_session", include_preview=True)
        if snapshot["counts"]["worksheets"] < 1 or snapshot["counts"]["graphs"] != 1:
            raise RuntimeError("Session snapshot lost the imported worksheet/graph")
        project = args.output_dir / "smoke.opju"
        await call("save_project", file_path=str(project))
        exported = await call("export_all_graphs", output_dir=str(args.output_dir), output_format="png")
        if exported["count"] != 1:
            raise RuntimeError("Batch export failed to discover the graph")
        await call("export_worksheet_to_csv", output_path=str(args.output_dir / "smoke.csv"))
        for path in [project, args.output_dir / "smoke.csv", Path(exported["files"][0]["output_path"])]:
            if not path.is_file() or path.stat().st_size == 0:
                raise RuntimeError(f"Missing or empty output: {path}")
        await call("open_project", file_path=str(project))
        reopened = await call("read_origin_session")
        if reopened["counts"]["worksheets"] < 1 or reopened["counts"]["graphs"] != 1:
            raise RuntimeError("Reopened project lost worksheet/graph")
        await call("release_origin")
        print("PASS: live Origin smoke workflow. Output retained for inspection.")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument("--live", action="store_true")
    modes.add_argument("--fake-workflow", action="store_true", help="Run numerical and failure acceptance over real stdio with the test backend (checkout required).")
    parser.add_argument("--allow-project-reset", action="store_true")
    parser.add_argument("--output-dir", type=Path)
    args = parser.parse_args()
    if args.live:
        if platform.system() != "Windows":
            parser.error("Live verification requires Windows and licensed OriginLab.")
        if not args.allow_project_reset or not args.output_dir:
            parser.error("Live verification resets Origin's current project: supply --allow-project-reset and --output-dir.")
        args.output_dir = args.output_dir.resolve()
        # New directory prevents accidental overwrite of earlier evidence.
        args.output_dir.mkdir(parents=True, exist_ok=False)
    anyio.run(verify, args)


if __name__ == "__main__":
    main()
