"""Soft dispatch timeout for Origin COM calls.

Inspired by Origin-Pro-MCP's ORIGIN_PRO_MCP_DISPATCH_TIMEOUT: give each
tool a soft wall-clock budget. When exceeded, raise a structured error so
the agent can recover — without force-killing Origin or cancelling the COM
call (the lock may stay held until Origin unblocks).
"""

from __future__ import annotations

import os
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import TimeoutError as FuturesTimeoutError
from typing import Final, TypeVar

from originlab_mcp.exceptions import ToolError

T = TypeVar("T")

# Sentinel: use ORIGINLAB_MCP_DISPATCH_TIMEOUT from the environment.
UNSET: Final[object] = object()

_DEFAULT_SECONDS: Final[float] = 90.0


def _falsey(raw: str | None) -> bool:
    return raw is None or raw.strip().lower() in {"off", "false", "no", "0", ""}


def parse_dispatch_timeout_seconds(
    environ: dict[str, str] | None = None,
) -> float | None:
    """Return soft budget in seconds, or ``None`` when disabled.

    ``ORIGINLAB_MCP_DISPATCH_TIMEOUT`` defaults to 90. Set ``off`` / ``false`` /
    ``no`` / ``0`` to disable. Any positive number is accepted.
    """
    env = environ if environ is not None else os.environ
    raw = env.get("ORIGINLAB_MCP_DISPATCH_TIMEOUT")
    if raw is None:
        return _DEFAULT_SECONDS
    if _falsey(raw):
        return None
    try:
        value = float(raw.strip())
    except ValueError:
        return _DEFAULT_SECONDS
    if value <= 0:
        return None
    return value


def resolve_dispatch_timeout(
    override: float | None | object = UNSET,
    *,
    environ: dict[str, str] | None = None,
) -> float | None:
    """Resolve per-call override or fall back to the env default.

    - ``UNSET`` → env policy
    - ``None`` or ``<= 0`` → disabled for this call
    - positive float → that many seconds
    """
    if override is UNSET:
        return parse_dispatch_timeout_seconds(environ)
    if override is None:
        return None
    try:
        value = float(override)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return parse_dispatch_timeout_seconds(environ)
    if value <= 0:
        return None
    return value


def run_with_soft_timeout(budget: float | None, fn: Callable[[], T]) -> T:
    """Run ``fn`` under a soft wall-clock budget.

    On timeout, raises ``ToolError`` and does **not** cancel ``fn`` — the COM
    call may still be blocked (e.g. modal dialog). Subsequent ``execute``
    calls will wait on the COM lock until it clears.
    """
    if budget is None or budget <= 0:
        return fn()

    executor = ThreadPoolExecutor(max_workers=1)
    try:
        future = executor.submit(fn)
        try:
            return future.result(timeout=budget)
        except FuturesTimeoutError as exc:
            raise ToolError(
                f"Origin 在 {budget:g}s 内未响应（软超时，未强制终止 COM）。",
                error_type="timeout",
                target="dispatch_timeout",
                hint=(
                    "常见原因是 Origin 弹出了模态对话框。"
                    "请切换到 Origin 窗口关闭对话框后重试；"
                    "或增大 ORIGINLAB_MCP_DISPATCH_TIMEOUT，"
                    "设为 off/0 可关闭软超时。"
                ),
                suggested_alternatives=["get_origin_info", "release_origin"],
            ) from exc
    finally:
        # Do not wait — a wedged COM call must not block tool teardown.
        executor.shutdown(wait=False, cancel_futures=False)
