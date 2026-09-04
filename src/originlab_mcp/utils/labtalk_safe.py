"""LabTalk safety helpers for the execute_labtalk escape hatch.

Inspired by Origin-Pro-MCP's confirm-gated LabTalk scanner: allow ordinary
styling/analysis scripts by default, but require explicit confirmation for
commands that can delete windows, reset the project, or reach the OS.
"""

from __future__ import annotations

import re
from typing import Final

# (human label, pattern, suggested typed tools)
_LABTALK_CONFIRM_RULES: Final[tuple[tuple[str, re.Pattern[str], tuple[str, ...]], ...]] = (
    (
        "system.",
        re.compile(r"\bsystem\s*\.", re.IGNORECASE),
        ("get_origin_info",),
    ),
    (
        "system(",
        re.compile(r"\bsystem\s*\(", re.IGNORECASE),
        ("get_origin_info",),
    ),
    (
        "run.section",
        re.compile(r"\brun\s*\.\s*section\b", re.IGNORECASE),
        ("get_origin_info",),
    ),
    (
        "run -",
        re.compile(r"\brun\s*-\s*[A-Za-z]", re.IGNORECASE),
        ("get_origin_info",),
    ),
    (
        "dll",
        re.compile(r"\bdll\b", re.IGNORECASE),
        ("get_origin_info",),
    ),
    (
        "dde",
        re.compile(r"\bdde\b", re.IGNORECASE),
        ("get_origin_info",),
    ),
    (
        "getfilename",
        re.compile(r"\bgetfilename\b", re.IGNORECASE),
        ("import_csv", "import_excel", "open_project"),
    ),
    (
        "getsavename",
        re.compile(r"\bgetsavename\b", re.IGNORECASE),
        ("save_project", "export_graph"),
    ),
    (
        "doc -s",
        re.compile(r"\bdoc\s*-\s*s\b", re.IGNORECASE),
        ("save_project",),
    ),
    (
        "doc -n",
        re.compile(r"\bdoc\s*-\s*n\b", re.IGNORECASE),
        ("new_project",),
    ),
    (
        "del/delete",
        re.compile(r"\b(?:del|delete)\b", re.IGNORECASE),
        ("remove_plot_from_graph", "delete_columns", "clear_worksheet"),
    ),
    (
        "win -c/-cd/-ct",
        re.compile(r"\bwin\s*-\s*c[dt]?\b", re.IGNORECASE),
        ("remove_plot_from_graph", "new_project"),
    ),
    (
        "label -r",
        re.compile(r"\blabel\s*-\s*r\b", re.IGNORECASE),
        ("remove_graph_label",),
    ),
)


def strip_labtalk_strings_and_comments(script: str) -> str:
    """Blank string literals and comments so keywords inside them are ignored."""
    out: list[str] = []
    i = 0
    n = len(script)
    state = "normal"
    while i < n:
        ch = script[i]
        nxt = script[i + 1] if i + 1 < n else ""
        if state == "normal":
            if ch == '"':
                state = "string"
                out.append(" ")
                i += 1
            elif ch == "/" and nxt == "/":
                state = "line_comment"
                out.append(" ")
                i += 2
            elif ch == "/" and nxt == "*":
                state = "block_comment"
                out.append(" ")
                i += 2
            else:
                out.append(ch)
                i += 1
        elif state == "string":
            if ch == '"':
                state = "normal"
            out.append("\n" if ch == "\n" else " ")
            i += 1
        elif state == "line_comment":
            if ch == "\n":
                state = "normal"
                out.append("\n")
            else:
                out.append(" ")
            i += 1
        else:  # block_comment
            if ch == "*" and nxt == "/":
                state = "normal"
                out.append(" ")
                i += 2
            else:
                out.append("\n" if ch == "\n" else " ")
                i += 1
    return "".join(out)


def classify_labtalk_script(script: str) -> tuple[bool, str, tuple[str, ...]]:
    """Classify whether a LabTalk script needs explicit confirmation.

    Returns:
        (requires_confirm, reason_label, suggested_alternatives)
    """
    cleaned = strip_labtalk_strings_and_comments(script)
    best_start: int | None = None
    best_label = ""
    best_alts: tuple[str, ...] = ()
    for label, pattern, alts in _LABTALK_CONFIRM_RULES:
        match = pattern.search(cleaned)
        if match is not None and (best_start is None or match.start() < best_start):
            best_start = match.start()
            best_label = label
            best_alts = alts
    if best_start is None:
        return (False, "", ())
    return (True, best_label, best_alts)
