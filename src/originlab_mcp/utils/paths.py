"""Filesystem path allowlist for MCP file I/O.

Inspired by Excel MCP ``ALLOWED_ROOTS`` / origin-mcp ``ORIGIN_MCP_ALLOWED_ROOTS``:
when configured, import/export/save/open paths must resolve under one of the
allowed roots. Unset or empty keeps the previous unrestricted behavior
(desktop Origin users often open files anywhere under their profile).
"""

from __future__ import annotations

import os
from pathlib import Path


def parse_allowed_roots(
    environ: dict[str, str] | None = None,
) -> tuple[Path, ...] | None:
    """Return normalized allowed roots, or ``None`` when unrestricted.

    ``ORIGINLAB_MCP_ALLOWED_ROOTS`` uses ``os.pathsep`` (``;`` on Windows,
    ``:`` on POSIX). Commas are also accepted as separators for JSON-friendly
    configs. Empty / unset → unrestricted.
    """
    env = environ if environ is not None else os.environ
    raw = env.get("ORIGINLAB_MCP_ALLOWED_ROOTS")
    if raw is None or not raw.strip():
        return None

    parts: list[str] = []
    for chunk in raw.replace(",", os.pathsep).split(os.pathsep):
        chunk = chunk.strip().strip('"').strip("'")
        if chunk:
            parts.append(chunk)
    if not parts:
        return None

    roots: list[Path] = []
    for part in parts:
        roots.append(Path(part).expanduser().resolve(strict=False))
    return tuple(roots)


def resolve_user_path(path: str) -> Path:
    """Expand ``~`` and resolve to an absolute path (may not exist yet)."""
    return Path(path).expanduser().resolve(strict=False)


def is_under_root(path: Path, root: Path) -> bool:
    """True when ``path`` is ``root`` or a descendant (after resolve)."""
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False


def check_allowed_path(
    path: str,
    *,
    environ: dict[str, str] | None = None,
) -> str | None:
    """Return an error message when ``path`` is outside allowed roots.

    Returns ``None`` when unrestricted or the path is allowed.
    """
    if not path or not str(path).strip():
        return "路径不能为空"

    roots = parse_allowed_roots(environ)
    if roots is None:
        return None

    try:
        resolved = resolve_user_path(path)
    except OSError as exc:
        return f"无法解析路径 '{path}': {exc}"

    if any(is_under_root(resolved, root) for root in roots):
        return None

    roots_display = os.pathsep.join(str(r) for r in roots)
    return (
        f"路径不在允许的根目录内: {resolved}。"
        f"允许的根: {roots_display}"
        "（由 ORIGINLAB_MCP_ALLOWED_ROOTS 配置）"
    )
