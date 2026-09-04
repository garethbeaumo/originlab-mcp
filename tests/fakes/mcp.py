"""Shared test doubles for MCP registration without a live Origin session."""

from __future__ import annotations

from types import MethodType
from typing import Any

from originlab_mcp.origin_manager import OriginManager


class DummyMCP:
    """Minimal FastMCP stand-in that records registered tools/resources."""

    def __init__(self) -> None:
        self.tools: dict[str, Any] = {}
        self.resources: dict[str, Any] = {}

    def tool(self, **_kwargs):
        def decorator(fn):
            self.tools[fn.__name__] = fn
            return fn

        return decorator

    def resource(self, uri, **_kwargs):
        def decorator(fn):
            self.resources[uri] = fn
            return fn

        return decorator


def attach_fake_origin(manager: OriginManager, op: Any) -> Any:
    """Bind a fake originpro module to a manager without COM connect()."""

    manager._op = op
    manager._connected = True

    def _execute(self: OriginManager, func, *args, **kwargs):
        self._cancel_idle_timer()
        with self._com_lock:
            return func(op, *args, **kwargs)

    manager.execute = MethodType(_execute, manager)
    return op
