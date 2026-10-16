"""Environment and Origin connectivity diagnostics.

Inspired by origin-mcp's ``origin_doctor``: give agents a single read-only
checklist before blaming tools for connection or policy misconfiguration.
"""

from __future__ import annotations

import importlib.util
import os
import platform
import sys
from typing import Any

from originlab_mcp import __version__
from originlab_mcp.utils.autosave import AutosavePolicy
from originlab_mcp.utils.dispatch import parse_dispatch_timeout_seconds
from originlab_mcp.utils.paths import parse_allowed_roots


def _check(
    name: str,
    *,
    status: str,
    detail: str,
    hint: str = "",
) -> dict[str, str]:
    item = {"name": name, "status": status, "detail": detail}
    if hint:
        item["hint"] = hint
    return item


def build_doctor_report(
    manager: Any | None = None,
    *,
    ping_origin: bool = False,
) -> dict[str, Any]:
    """Build a structured health report.

    Does not require Origin unless ``ping_origin`` is true. Status values:
    ``ok``, ``warn``, ``error``.
    """
    checks: list[dict[str, str]] = []

    system = platform.system()
    checks.append(
        _check(
            "platform",
            status="ok" if system == "Windows" else "warn",
            detail=f"{system} {platform.release()} ({platform.machine()})",
            hint=(
                ""
                if system == "Windows"
                else "Origin COM automation requires Windows + OriginLab."
            ),
        )
    )

    checks.append(
        _check(
            "python",
            status="ok",
            detail=f"{sys.version_info.major}.{sys.version_info.minor}."
            f"{sys.version_info.micro} ({sys.executable})",
        )
    )

    checks.append(
        _check(
            "package_version",
            status="ok",
            detail=__version__,
        )
    )

    originpro_spec = importlib.util.find_spec("originpro")
    if originpro_spec is not None:
        checks.append(
            _check(
                "originpro",
                status="ok",
                detail=f"importable ({originpro_spec.origin})",
            )
        )
    else:
        checks.append(
            _check(
                "originpro",
                status="error" if system == "Windows" else "warn",
                detail="not installed / not importable",
                hint=(
                    "On Windows: pip/uv install originpro with a licensed OriginLab."
                    if system == "Windows"
                    else "Expected on Linux CI; COM tools need Windows + originpro."
                ),
            )
        )

    policy = AutosavePolicy.from_env()
    checks.append(
        _check(
            "autosave",
            status="ok",
            detail=(
                f"enabled={policy.enabled} required={policy.required} "
                f"interval={policy.interval_seconds}"
            ),
        )
    )

    timeout = parse_dispatch_timeout_seconds()
    checks.append(
        _check(
            "dispatch_timeout",
            status="ok",
            detail="off" if timeout is None else f"{timeout:g}s",
        )
    )

    roots = parse_allowed_roots()
    if roots is None:
        checks.append(
            _check(
                "allowed_roots",
                status="ok",
                detail="unrestricted (ORIGINLAB_MCP_ALLOWED_ROOTS unset)",
            )
        )
    else:
        checks.append(
            _check(
                "allowed_roots",
                status="ok",
                detail=os.pathsep.join(str(r) for r in roots),
            )
        )

    connected = bool(manager is not None and getattr(manager, "is_connected", False))
    project_path = None
    if manager is not None:
        project_path = getattr(manager, "project_path", None)

    origin_ping: dict[str, Any] | None = None
    if ping_origin and manager is not None:
        try:
            manager.connect()
            info = manager.get_info()
            connected = True
            project_path = info.get("project_path") or project_path
            origin_ping = {
                "ok": True,
                "exe_path": info.get("exe_path"),
                "user_path": info.get("user_path"),
                "worksheet_count": info.get("worksheet_count"),
                "graph_count": info.get("graph_count"),
            }
            checks.append(
                _check(
                    "origin_ping",
                    status="ok",
                    detail=f"connected exe={info.get('exe_path')}",
                )
            )
        except Exception as exc:
            origin_ping = {"ok": False, "error": str(exc)}
            checks.append(
                _check(
                    "origin_ping",
                    status="error",
                    detail=str(exc),
                    hint="Start OriginLab, confirm license, then retry with ping_origin=true.",
                )
            )
    else:
        checks.append(
            _check(
                "origin_connection",
                status="ok" if connected else "warn",
                detail="connected" if connected else "not connected (ping skipped)",
                hint=(
                    ""
                    if connected
                    else "Pass ping_origin=true to attempt a live Origin connect."
                ),
            )
        )

    statuses = {c["status"] for c in checks}
    if "error" in statuses:
        overall = "error"
    elif "warn" in statuses:
        overall = "degraded"
    else:
        overall = "ok"

    return {
        "overall": overall,
        "checks": checks,
        "connected": connected,
        "project_path": project_path,
        "origin_ping": origin_ping,
        "policy": {
            "autosave_enabled": policy.enabled,
            "autosave_required": policy.required,
            "autosave_interval": policy.interval_seconds,
            "dispatch_timeout": timeout,
            "allowed_roots": [str(r) for r in roots] if roots is not None else None,
        },
    }
