"""Test-only stdio host: production registrations with a controlled backend.

Never imported by the application. No Origin installation or COM is used.
"""

import threading

from originlab_mcp import server
from tests.fakes import FakeOrigin


class ProtocolOrigin(FakeOrigin):
    def pages(self, kind: str):
        if kind == "w":
            raise RuntimeError("controlled backend failure")
        return super().pages(kind)

    def lt_exec(self, command: str):
        # A finite blocked call exercises the real manager's soft deadline.
        threading.Event().wait(0.2)
        return None


if __name__ == "__main__":
    server._manager._op = ProtocolOrigin()
    server._manager._connected = True
    server._manager._idle_timeout = 0
    server.main()
