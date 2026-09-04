"""MCP ToolAnnotations presets for OriginLab tools.

Hints follow the MCP ToolAnnotations contract so clients can prefer
read-only tools, avoid accidental destructive calls, and understand
filesystem / LabTalk open-world side effects.
"""

from __future__ import annotations

from mcp.types import ToolAnnotations

# Pure inspection — no Origin or filesystem mutation.
READ_ONLY = ToolAnnotations(
    readOnlyHint=True,
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=False,
)

# Additive or reversible Origin edits (create/set/add/fit).
MUTATING = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=False,
    idempotentHint=False,
    openWorldHint=False,
)

# Absolute setters that are safe to retry with the same arguments.
IDEMPOTENT_MUTATING = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=False,
)

# Deletes, clears, or replaces the live Origin session.
DESTRUCTIVE = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=True,
    idempotentHint=False,
    openWorldHint=False,
)

# Reads or writes paths outside the closed Origin session.
OPEN_WORLD = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=False,
    idempotentHint=False,
    openWorldHint=True,
)

# Open-world operations that may destroy or overwrite user data.
DESTRUCTIVE_OPEN_WORLD = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=True,
    idempotentHint=False,
    openWorldHint=True,
)

# Arbitrary LabTalk — open world and potentially destructive.
LABTALK = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=True,
    idempotentHint=False,
    openWorldHint=True,
)
