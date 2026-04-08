"""OriginLab MCP Server - AI 驱动的 OriginLab 自动化控制。"""

from __future__ import annotations

import re
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

__all__ = ["__version__"]

_DEFAULT_VERSION = "0.2.1"
_PYPROJECT = Path(__file__).resolve().parents[2] / "pyproject.toml"

if _PYPROJECT.is_file():
    content = _PYPROJECT.read_text(encoding="utf-8")
    match = re.search(r'^version = "([^"]+)"$', content, re.MULTILINE)
    __version__ = match.group(1) if match else _DEFAULT_VERSION
else:
    try:
        __version__ = version("originlab-mcp")
    except PackageNotFoundError:
        __version__ = _DEFAULT_VERSION
