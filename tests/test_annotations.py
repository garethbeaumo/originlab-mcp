"""Regression tests for MCP ToolAnnotations hints."""

from __future__ import annotations

import asyncio

from originlab_mcp.server import mcp

READ_ONLY_TOOLS = {
    "get_origin_info",
    "originlab_doctor",
    "read_origin_session",
    "list_worksheets",
    "get_worksheet_info",
    "get_worksheet_data",
    "get_cell_value",
    "list_graphs",
    "list_graph_templates",
    "get_graph_info",
    "list_fit_functions",
    "get_labtalk_variable",
}

DESTRUCTIVE_TOOLS = {
    "close_origin",
    "clear_worksheet",
    "delete_columns",
    "remove_plot_from_graph",
    "remove_graph_label",
    "new_project",
    "open_project",
    "execute_labtalk",
}

OPEN_WORLD_TOOLS = {
    "import_csv",
    "import_excel",
    "copy_graph_to_clipboard",
    "export_graph",
    "export_all_graphs",
    "export_worksheet_to_csv",
    "save_project",
    "open_project",
    "execute_labtalk",
}


def _tools_by_name():
    tools = asyncio.run(mcp.list_tools())
    return {tool.name: tool for tool in tools}


def test_all_tools_have_annotations():
    by_name = _tools_by_name()
    assert len(by_name) == 67
    missing = [name for name, tool in by_name.items() if tool.annotations is None]
    assert missing == [], f"tools missing annotations: {missing}"


def test_read_only_hints():
    by_name = _tools_by_name()
    for name in READ_ONLY_TOOLS:
        ann = by_name[name].annotations
        assert ann is not None
        assert ann.readOnlyHint is True
        assert ann.destructiveHint is False
        assert ann.idempotentHint is True


def test_destructive_and_open_world_hints():
    by_name = _tools_by_name()
    for name in DESTRUCTIVE_TOOLS:
        ann = by_name[name].annotations
        assert ann is not None
        assert ann.destructiveHint is True
        assert ann.readOnlyHint is False
    for name in OPEN_WORLD_TOOLS:
        ann = by_name[name].annotations
        assert ann is not None
        assert ann.openWorldHint is True
