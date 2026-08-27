#!/usr/bin/env bash
set -euo pipefail

# Install uv (project package/dependency manager) if it is not already present.
if ! command -v uv >/dev/null 2>&1; then
    curl -LsSf https://astral.sh/uv/install.sh | sh
fi
export PATH="$HOME/.local/bin:$PATH"

# `originpro`/`originext` only ship Windows wheels and cannot install on Linux.
# They are imported lazily and are only needed for a live OriginLab (Windows)
# connection, so the MCP server, UI panel, and test suite all run without them.
uv sync --no-install-package originpro --no-install-package originext
