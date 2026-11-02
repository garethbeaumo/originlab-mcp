"""Production MCP server with a fake backend for data/session/lifecycle tests."""
import os
import threading
from pathlib import Path

from originlab_mcp import server
from tests.fakes import FakeOrigin


def record(event):
    path = os.environ.get("TEST_LIFECYCLE_LOG")
    if path:
        with Path(path).open("a", encoding="utf-8") as stream:
            stream.write(event + "\n")


def native_result(outcome):
    def operation(*args, **kwargs):
        if outcome == "error":
            raise RuntimeError("native operation failed")
        return {"zero": 0, "nonzero": 7, "none": None}[outcome]
    return operation


class ContractOrigin(FakeOrigin):
    session_page_types = {"Notes", "Excel"}

    def __init__(self):
        super().__init__()
        if os.environ.get("TEST_SESSION_MODE") == "data":
            for outcome in ("zero", "nonzero", "none", "error"):
                sheet = self.new_sheet(lname=outcome)
                sheet.from_list(0, [3, 1, 2])
                sheet.sort = native_result(outcome)
                sheet.clear = native_result(outcome)

    def pages(self, kind):
        mode = os.environ.get("TEST_SESSION_MODE")
        if mode == "core_failure" and kind == "w":
            raise RuntimeError("core enumeration failed")
        if mode == "partial" and kind == "Notes":
            raise RuntimeError("optional enumeration failed")
        return super().pages(kind)

    def lt_exec(self, command):
        record("started")
        threading.Event().wait(0.3)
        record("finished")

    def detach(self):
        record("detach")
        return super().detach()

    def exit(self):
        record("exit")
        return super().exit()


if __name__ == "__main__":
    record("startup")
    server._manager._op = ContractOrigin()
    server._manager._connected = True
    server._manager._idle_timeout = 0
    server.main()
