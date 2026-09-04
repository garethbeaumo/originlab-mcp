"""Preflight autosave before destructive Origin mutations.

Inspired by Origin-Pro-MCP: save the open project in place before ops that
can wipe worksheets, graphs, or the whole project. This module is COM-light —
classification is pure Python; OriginManager performs the actual save.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Final

from originlab_mcp.utils.labtalk_safe import classify_labtalk_script

# Typed tools that can destroy or replace project content.
DESTRUCTIVE_TOOLS: Final[frozenset[str]] = frozenset(
    {
        "new_project",
        "open_project",
        "delete_columns",
        "clear_worksheet",
        "remove_plot_from_graph",
        "remove_graph_label",
        "close_origin",
    }
)

# LabTalk confirm-gate labels that destroy project data (not mere save/export).
_DESTRUCTIVE_LABTALK_LABELS: Final[frozenset[str]] = frozenset(
    {
        "del/delete",
        "win -c/-cd/-ct",
        "doc -n",
    }
)


def _falsey(raw: str | None) -> bool:
    return raw is None or raw.strip().lower() in {"off", "false", "no", "0", ""}


@dataclass(frozen=True)
class AutosavePolicy:
    """Env-driven autosave configuration.

    ``ORIGINLAB_MCP_AUTOSAVE``          default ON; set off/false/no/0 to disable
    ``ORIGINLAB_MCP_AUTOSAVE_REQUIRED`` default OFF; when ON, failed preflight
                                       blocks the destructive operation
    """

    enabled: bool = True
    required: bool = False

    @classmethod
    def from_env(cls, environ: dict[str, str] | None = None) -> AutosavePolicy:
        env = environ if environ is not None else os.environ
        raw = env.get("ORIGINLAB_MCP_AUTOSAVE")
        enabled = True if raw is None else not _falsey(raw)
        req_raw = env.get("ORIGINLAB_MCP_AUTOSAVE_REQUIRED")
        required = False if req_raw is None else not _falsey(req_raw)
        return cls(enabled=enabled, required=required)


def should_autosave_labtalk(command: str, *, confirm: bool) -> bool:
    """True when a confirmed LabTalk script is project-destructive."""
    if not confirm or not command:
        return False
    requires, reason, _alts = classify_labtalk_script(command)
    return requires and reason in _DESTRUCTIVE_LABTALK_LABELS


def should_autosave_tool(tool_name: str) -> bool:
    return tool_name in DESTRUCTIVE_TOOLS


def collect_autosave_warnings(status: dict) -> list[str]:
    """Convert a preflight_autosave status dict into response warnings."""
    message = str(status.get("message") or "").strip()
    if not message:
        return []
    if status.get("saved"):
        return [message]
    if status.get("attempted") and not status.get("saved"):
        return [message]
    if "no project path" in message:
        return [message]
    return []
