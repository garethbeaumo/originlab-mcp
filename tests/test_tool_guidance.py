"""Agent-guidance contracts for MCP tool descriptions and next_suggestions.

Inspired by 2026 MCP guidance research: selection quality hinges on
explicit When-to-use / When-not-to-use bounds and valid next-step hints.
"""

from __future__ import annotations

import ast
import asyncio
from pathlib import Path

from originlab_mcp.server import mcp

TOOLS_DIR = Path(__file__).resolve().parents[1] / "src" / "originlab_mcp" / "tools"


def _registered_tools():
    return asyncio.run(mcp.list_tools())


def _iter_success_suggestion_lists(tree: ast.AST):
    """Yield (tool_function_name, suggestion_list) from success_response calls."""

    class Visitor(ast.NodeVisitor):
        def __init__(self):
            self.current_tool: str | None = None
            self.found: list[tuple[str, list[str]]] = []

        def visit_FunctionDef(self, node: ast.FunctionDef):
            previous = self.current_tool
            # Nested tool functions inside register_* are the ones we care about.
            if (
                (previous is None or not node.name.startswith("register_"))
                and not node.name.startswith("register_")
                and not node.name.startswith("_")
            ):
                self.current_tool = node.name
            self.generic_visit(node)
            self.current_tool = previous

        def visit_Call(self, node: ast.Call):
            func = node.func
            name = None
            if isinstance(func, ast.Name):
                name = func.id
            elif isinstance(func, ast.Attribute):
                name = func.attr
            if name == "success_response" and self.current_tool:
                for kw in node.keywords:
                    if kw.arg != "next_suggestions":
                        continue
                    if isinstance(kw.value, ast.List):
                        values = []
                        for elt in kw.value.elts:
                            if isinstance(elt, ast.Constant) and isinstance(elt.value, str):
                                values.append(elt.value)
                            else:
                                raise AssertionError(
                                    f"{self.current_tool}: non-literal suggestion {ast.dump(elt)}"
                                )
                        self.found.append((self.current_tool, values))
            self.generic_visit(node)

    visitor = Visitor()
    visitor.visit(tree)
    return visitor.found


def test_all_tools_have_selection_bounds():
    tools = _registered_tools()
    missing = []
    for tool in tools:
        doc = tool.description or ""
        problems = []
        if "When to use" not in doc:
            problems.append("missing When to use")
        if "When not to use" not in doc:
            problems.append("missing When not to use")
        if problems:
            missing.append((tool.name, problems))
    assert missing == [], f"tool description gaps: {missing}"


def test_next_suggestions_reference_registered_tools():
    registered = {tool.name for tool in _registered_tools()}
    unknown: list[tuple[str, str]] = []
    missing_on_success: list[str] = []

    for path in sorted(TOOLS_DIR.glob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        pairs = _iter_success_suggestion_lists(tree)
        tools_with_success = {
            node.name
            for node in ast.walk(tree)
            if isinstance(node, ast.FunctionDef)
            and not node.name.startswith("register_")
            and not node.name.startswith("_")
            and any(
                (
                    isinstance(call, ast.Call)
                    and (
                        (isinstance(call.func, ast.Name) and call.func.id == "success_response")
                        or (
                            isinstance(call.func, ast.Attribute)
                            and call.func.attr == "success_response"
                        )
                    )
                )
                for call in ast.walk(node)
            )
        }
        suggested_tools = {name for name, _ in pairs}
        for tool_name in sorted(tools_with_success - suggested_tools):
            # Tools that only return success_response without next_suggestions.
            # Detect precisely by scanning each function body.
            for node in ast.walk(tree):
                if not isinstance(node, ast.FunctionDef) or node.name != tool_name:
                    continue
                for call in ast.walk(node):
                    if not isinstance(call, ast.Call):
                        continue
                    func = call.func
                    is_success = (
                        isinstance(func, ast.Name) and func.id == "success_response"
                    ) or (isinstance(func, ast.Attribute) and func.attr == "success_response")
                    if not is_success:
                        continue
                    keys = {kw.arg for kw in call.keywords}
                    if "next_suggestions" not in keys:
                        missing_on_success.append(f"{path.name}:{tool_name}")

        for tool_name, suggestions in pairs:
            for suggestion in suggestions:
                if suggestion not in registered:
                    unknown.append((tool_name, suggestion))

    assert unknown == [], f"unknown next_suggestions: {unknown}"
    assert missing_on_success == [], (
        "success_response missing next_suggestions: "
        f"{sorted(set(missing_on_success))}"
    )


def test_plot_creation_tools_cross_reference_each_other():
    tools = {tool.name: tool for tool in _registered_tools()}
    create_doc = tools["create_plot"].description or ""
    add_doc = tools["add_plot_to_graph"].description or ""
    change_doc = tools["change_plot_type"].description or ""

    assert "change_plot_type" in create_doc
    assert "add_plot_to_graph" in create_doc
    assert "create_plot" in add_doc
    assert "change_plot_type" in add_doc
    assert "create_plot" in change_doc
    assert "add_plot_to_graph" in change_doc
