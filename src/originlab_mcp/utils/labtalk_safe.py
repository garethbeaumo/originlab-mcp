"""LabTalk safety helpers for the execute_labtalk escape hatch.

Inspired by Origin-Pro-MCP's confirm-gated LabTalk scanner: allow ordinary
styling/analysis scripts by default, but require explicit confirmation for
commands that can delete windows, reset the project, or reach the OS.

This is a best-effort keyword scan, not a sandbox: LabTalk can build command
text at runtime in ways a static scan cannot follow. The real safeguard is the
tool's destructive/open-world annotation, which makes clients ask the user.
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
        # run.file / run.python / run.loadOC etc. execute external code.
        "run.<method>",
        re.compile(r"\brun\s*\.\s*[A-Za-z]", re.IGNORECASE),
        ("get_origin_info",),
    ),
    (
        "python",
        re.compile(r"\bpython\b", re.IGNORECASE),
        ("get_origin_info",),
    ),
    (
        # A statement starting with %(...) / %A / %1 runs substituted text as a
        # command, e.g. string s$="doc -n"; %(s$); — the scan cannot see it.
        # Enter parenthesized script groups only at a statement boundary.
        # Plain string-register assignment (%A = ...) is left alone.
        "%() dynamic command",
        re.compile(r"(?:^|[;{}\n])\s*(?:\(\s*)*%(?:\(|[A-Za-z1-5](?!\s*=))", re.IGNORECASE),
        ("get_origin_info",),
    ),
    (
        # These options evaluate script text supplied as a string, which is
        # intentionally blanked before keyword scanning like python's input.
        "queue -p/-ps",
        re.compile(r"\bqueue\s*-\s*ps?\b", re.IGNORECASE),
        ("get_origin_info",),
    ),
    (
        "run -",
        re.compile(r"\brun\s*-\s*[A-Za-z]", re.IGNORECASE),
        ("get_origin_info",),
    ),
    (
        # Legacy no-option run executes an external script file. Match only
        # command positions and leave assignments/methods to their own rules.
        "run script",
        re.compile(
            r"(?:^|[;{}\n])\s*(?:\(\s*)*run(?=\s|[;}]|$)"
            r"(?!\s*(?:=|\.[A-Za-z]|-\s*[A-Za-z]))",
            re.IGNORECASE,
        ),
        ("get_origin_info",),
    ),
    (
        # A runtime selector may expand to LR plus a class or include the
        # script itself, so its body boundary cannot be established safely.
        "doc -e/-ef dynamic selector",
        re.compile(
            r"\bdoc(?:ument)?\s*-\s*ef?\s+%(?:\(|[A-Za-z1-5])",
            re.IGNORECASE,
        ),
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


def _with_statement_boundaries(cleaned: str) -> str:
    """Expose braceless control-flow and document-iteration script bodies.

    Balance condition parentheses after strings/comments have been blanked;
    a flat regex would stop at a nested function call or parenthesized term.
    Preserve offsets so risk ordering still refers to the original scan.
    """
    out = list(cleaned)
    for header in re.finditer(r"\b(?:if|for|loop|switch)\s*\(", cleaned, re.IGNORECASE):
        # Whitespace before the opening parenthesis is still part of the
        # header, including newlines left by blanked comments.
        for index in range(header.start(), header.end()):
            if cleaned[index] == "\n":
                out[index] = " "
        depth = 1
        for index in range(header.end(), len(cleaned)):
            if cleaned[index] == "(":
                depth += 1
            elif cleaned[index] == ")":
                depth -= 1
                if depth == 0:
                    out[index] = ";"
                    break
            elif cleaned[index] in ";\n":
                # Control-header expressions are not script statements.
                out[index] = " "
    for header in re.finditer(r"\belse\b", cleaned, re.IGNORECASE):
        out[header.end() - 1] = ";"
    # doc -e/-ef accepts a single statement after the object selector.
    # LR additionally takes a report class name; LRW is a single selector.
    # https://docs.originlab.com/labtalk/ref/document_options_others/
    for header in re.finditer(
        r"\bdoc(?:ument)?\s*-\s*ef?\s+(?:lr\s+)?[A-Za-z_]\w*\b",
        cleaned,
        re.IGNORECASE,
    ):
        out[header.end() - 1] = ";"
    # Labels start statements, unlike a book/sheet/variable named "case".
    # Skip nested expression/range colons and reject assignments or unmatched
    # closing delimiters (including a multiline [case] range reference).
    for label in re.finditer(
        r"(?:^|(?<=[;{}\n:]))\s*(?:case|default)\b", cleaned, re.IGNORECASE
    ):
        brackets: list[str] = []
        for index in range(label.end(), len(cleaned)):
            ch = cleaned[index]
            if ch in "([":
                brackets.append(ch)
            elif ch in ")]":
                if not brackets or brackets.pop() != {")": "(", "]": "["}[ch]:
                    break
            elif ch in ";{}=" and not brackets:
                break
            elif ch == ":" and not brackets:
                for header_index in range(label.start(), index):
                    if cleaned[header_index] == "\n":
                        out[header_index] = " "
                out[index] = ";"
                break
    return "".join(out)


def find_labtalk_risks(script: str) -> tuple[tuple[str, tuple[str, ...]], ...]:
    """Return every matching risk rule in script order, with rule-order ties.

    Consumers such as autosave must inspect all risks: the first confirmation
    reason can be unrelated to a later project-destructive operation.
    """
    cleaned = strip_labtalk_strings_and_comments(script)
    statement_boundaries = _with_statement_boundaries(cleaned)
    matches: list[tuple[int, str, tuple[str, ...]]] = []
    for label, pattern, alts in _LABTALK_CONFIRM_RULES:
        scan = (
            statement_boundaries
            if label in {"%() dynamic command", "run script"}
            else cleaned
        )
        match = pattern.search(scan)
        if match is not None:
            matches.append((match.start(), label, alts))
    matches.sort(key=lambda match: match[0])
    return tuple((label, alts) for _start, label, alts in matches)


def classify_labtalk_script(script: str) -> tuple[bool, str, tuple[str, ...]]:
    """Return (requires_confirm, first_reason_label, suggested_alternatives)."""
    risks = find_labtalk_risks(script)
    if not risks:
        return (False, "", ())
    label, alternatives = risks[0]
    return (True, label, alternatives)
