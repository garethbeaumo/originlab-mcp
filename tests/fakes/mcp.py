"""Shared test doubles for MCP registration without a live Origin session."""

from __future__ import annotations

from types import MethodType
from typing import Any

from originlab_mcp.origin_manager import OriginManager


class DummyMCP:
    """Minimal FastMCP stand-in that records registered tools/resources/prompts."""

    def __init__(self) -> None:
        self.tools: dict[str, Any] = {}
        self.resources: dict[str, Any] = {}
        self.prompts: dict[str, Any] = {}

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

    def prompt(self, **_kwargs):
        def decorator(fn):
            name = _kwargs.get("name") or fn.__name__
            self.prompts[name] = fn
            return fn

        return decorator


def attach_fake_origin(manager: OriginManager, op: Any) -> Any:
    """Bind a fake originpro module to a manager without COM connect()."""

    manager._op = op
    manager._connected = True

    def _execute(
        self: OriginManager,
        func,
        *args,
        dispatch_timeout=None,
        **kwargs,
    ):
        from originlab_mcp.utils.dispatch import (
            UNSET,
            resolve_dispatch_timeout,
            run_with_soft_timeout,
        )

        override = UNSET if dispatch_timeout is None else dispatch_timeout
        budget = resolve_dispatch_timeout(override)

        def _run():
            self._cancel_idle_timer()
            with self._com_lock:
                return func(op, *args, **kwargs)

        return run_with_soft_timeout(budget, _run)

    manager.execute = MethodType(_execute, manager)
    return op
