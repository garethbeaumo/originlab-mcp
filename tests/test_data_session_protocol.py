"""Data errors and lifecycle acceptance over actual MCP stdio; fake Origin only."""
import asyncio
import json
import os
import sys
from pathlib import Path

import pytest
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

ROOT = Path(__file__).resolve().parents[1]


def environment(mode):
    env = dict(os.environ)
    env.update(PYTHONPATH=os.pathsep.join([str(ROOT / "src"), str(ROOT)]),
               ORIGINLAB_MCP_AUTOSAVE="0", ORIGINLAB_MCP_AUTOSAVE_INTERVAL="off",
               ORIGINLAB_MCP_DISPATCH_TIMEOUT="off", TEST_SESSION_MODE=mode)
    return env


@pytest.mark.parametrize("mode", ["data", "empty", "core_failure", "partial"])
def test_data_session_results_over_stdio(mode):
    async def check():
        params = StdioServerParameters(command=sys.executable,
            args=["-m", "tests.data_session_server"], cwd=str(ROOT), env=environment(mode))
        async with stdio_client(params) as (read, write), ClientSession(read, write) as session:
            await session.initialize()
            if mode == "data":
                for tool in ("sort_worksheet", "clear_worksheet"):
                    for outcome in ("zero", "nonzero", "none", "error"):
                        kwargs = {"sheet_name": outcome}
                        if tool == "sort_worksheet":
                            kwargs["col"] = 0
                        result = await session.call_tool(tool, kwargs)
                        assert result.isError is (outcome != "zero")
                        assert json.loads(result.content[0].text)["ok"] is (outcome == "zero")
            else:
                result = await session.call_tool("read_origin_session", {})
                payload = json.loads(result.content[0].text)
                assert result.isError is (mode == "core_failure")
                assert payload["ok"] is (mode != "core_failure")
                if mode != "core_failure":
                    assert payload["data"]["counts"]["worksheets"] == 0
                    assert payload["data"]["partial"] is (mode == "partial")
                    if mode == "partial":
                        assert payload["data"]["counts"]["notes"] is None
                        assert payload["warnings"]
    asyncio.run(check())


def test_cancel_disconnect_restart_detaches_without_exit(tmp_path):
    async def run():
        log = tmp_path / "lifecycle.log"
        env = environment("empty")
        env["TEST_LIFECYCLE_LOG"] = str(log)
        for attempt in range(2):
            process = await asyncio.create_subprocess_exec(
                sys.executable, "-m", "tests.data_session_server", cwd=ROOT, env=env,
                stdin=asyncio.subprocess.PIPE, stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.DEVNULL, limit=1024 * 1024)

            async def send(message, process=process):
                process.stdin.write((json.dumps({"jsonrpc": "2.0", **message}) + "\n").encode())
                await process.stdin.drain()

            async def response(request_id, process=process):
                while True:
                    line = await asyncio.wait_for(process.stdout.readline(), 10)
                    assert line, "server disconnected before response"
                    value = json.loads(line)
                    if value.get("id") == request_id:
                        return value

            try:
                await send({"id": 1, "method": "initialize", "params": {
                    "protocolVersion": "2024-11-05", "capabilities": {},
                    "clientInfo": {"name": "lifecycle-test", "version": "1"}}})
                assert "result" in await response(1)
                await send({"method": "notifications/initialized"})
                await send({"id": 2, "method": "tools/call", "params": {
                    "name": "execute_labtalk", "arguments": {"command": "type 1;"}}})
                async def started(attempt=attempt):
                    while not log.exists() or log.read_text().count("started") < attempt + 1:
                        await asyncio.sleep(0.01)
                await asyncio.wait_for(started(), 10)
                await send({"method": "notifications/cancelled", "params": {"requestId": 2}})
                if attempt == 0:
                    await send({"id": 3, "method": "tools/list"})
                    assert len((await response(3))["result"]["tools"]) == 67
                # Second cycle closes the transport while the native call is active.
                process.stdin.close()
                assert await asyncio.wait_for(process.wait(), 10) == 0
            finally:
                if process.returncode is None:
                    process.kill()  # only the test subprocess, never Origin
                    await process.wait()
        events = log.read_text().splitlines()
        assert events.count("startup") == events.count("detach") == 2
        assert events.count("finished") == 2  # cancellation is not native abort
        assert "exit" not in events
    asyncio.run(run())
