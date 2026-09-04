"""MCP prompt templates for common OriginLab agent workflows.

Prompts are reusable message templates (MCP prompts/list + prompts/get).
They guide hosts toward typed tools and safe sequencing without executing
Origin COM themselves.
"""

from __future__ import annotations

from typing import Any

SERVER_INSTRUCTIONS = """\
OriginLab MCP Server — control OriginLab via MCP.

Preferred workflow:
1. Read first: use resource originlab://session or tool read_origin_session.
2. Mutate with typed tools (import_*, create_plot, set_*, fit, export_*).
3. Prefer standard tools over execute_labtalk; LabTalk is a last resort and
   destructive/system commands require confirm=true.
4. Call save_project(file_path=...) early so preflight autosave can protect
   destructive ops (new_project, clear_worksheet, delete_columns, etc.).
5. Respect ORIGINLAB_MCP_ALLOWED_ROOTS when configured; paths outside the
   allowlist are rejected.
6. If a tool returns error.type=timeout, check Origin for a modal dialog,
   then retry (ORIGINLAB_MCP_DISPATCH_TIMEOUT soft budget).

Never invent Origin COM or LabTalk when a typed tool exists.
"""


def register_prompts(mcp: Any) -> None:
    """Register workflow prompt templates on a FastMCP (or DummyMCP) instance."""

    @mcp.prompt(
        name="originlab_inspect_session",
        title="Inspect Origin session",
        description="Read-only first look at the current Origin project.",
    )
    def originlab_inspect_session() -> str:
        """Guide the agent to inspect the live Origin session safely."""
        return (
            "Inspect the current OriginLab session before making changes.\n"
            "1. Prefer resource originlab://session, or call read_origin_session.\n"
            "2. Optionally list worksheets/graphs via originlab://worksheets and "
            "originlab://graphs.\n"
            "3. Summarize active worksheet/graph, counts, and whether a project "
            "path is known.\n"
            "4. Do not mutate data until the user asks."
        )

    @mcp.prompt(
        name="originlab_csv_to_plot",
        title="CSV import to plot",
        description="Import a CSV, set XY designations, create a plot, and export.",
    )
    def originlab_csv_to_plot(
        file_path: str = "",
        plot_type: str = "scatter",
    ) -> str:
        """Guide import → designations → plot → export."""
        path_line = (
            f"Import file: {file_path}"
            if file_path.strip()
            else "Ask the user for the CSV path (must be under ALLOWED_ROOTS if set)."
        )
        return (
            "Create a figure from a CSV using typed OriginLab MCP tools.\n"
            f"{path_line}\n"
            f"Requested plot type: {plot_type}\n"
            "Steps:\n"
            "1. import_csv(file_path=...)\n"
            "2. get_worksheet_info → set_column_designations (typically XY)\n"
            "3. create_plot(x_col=..., y_cols=..., plot_type=...)\n"
            "4. Optionally set_axis_title / set_plot_color\n"
            "5. export_graph(output_path=...) then save_project(file_path=...)\n"
            "Do not use execute_labtalk for this workflow."
        )

    @mcp.prompt(
        name="originlab_publication_style",
        title="Publication figure styling",
        description="Apply publication-ready styling to the active (or named) graph.",
    )
    def originlab_publication_style(graph_name: str = "") -> str:
        """Guide publication styling with apply_publication_style and related tools."""
        target = (
            f"Target graph: {graph_name}"
            if graph_name.strip()
            else "Target: current active graph (list_graphs / get_graph_info if unsure)."
        )
        return (
            "Style an Origin graph for publication.\n"
            f"{target}\n"
            "Steps:\n"
            "1. read_origin_session or get_graph_info to confirm the graph/layer.\n"
            "2. apply_publication_style (set layer_index if needed).\n"
            "3. Tune with set_axis_title, set_plot_line_width, set_symbol_size, "
            "set_legend as required.\n"
            "4. export_graph to PNG/SVG/PDF, then save_project.\n"
            "Prefer typed customize tools over LabTalk."
        )

    @mcp.prompt(
        name="originlab_fit_curve",
        title="Curve fitting workflow",
        description="Run linear or nonlinear fitting on the active worksheet.",
    )
    def originlab_fit_curve(function_name: str = "Gauss") -> str:
        """Guide linear_fit / nonlinear_fit usage."""
        return (
            "Fit data in the active Origin worksheet.\n"
            f"Preferred function: {function_name}\n"
            "Steps:\n"
            "1. get_worksheet_info / get_worksheet_data to confirm X/Y columns.\n"
            "2. list_fit_functions if unsure which models exist.\n"
            "3. Use linear_fit for straight lines; otherwise nonlinear_fit with "
            f"function_name='{function_name}' and sensible initial_params.\n"
            "4. Report parameters, errors, and R² from the tool response.\n"
            "5. Optionally plot residuals or export results; save_project when done."
        )

    @mcp.prompt(
        name="originlab_safe_destructive",
        title="Safe destructive operation",
        description="Checklist before new_project / clear / delete / close Origin.",
    )
    def originlab_safe_destructive(operation: str = "new_project") -> str:
        """Guide save-first sequencing for destructive tools."""
        return (
            "Prepare a destructive Origin operation safely.\n"
            f"Requested operation: {operation}\n"
            "Checklist:\n"
            "1. read_origin_session — confirm what will be lost.\n"
            "2. save_project(file_path=...) so project_path is known "
            "(enables preflight autosave).\n"
            "3. If ORIGINLAB_MCP_AUTOSAVE_REQUIRED is on, ensure save succeeds.\n"
            "4. Call the typed destructive tool "
            "(new_project / clear_worksheet / delete_columns / "
            "remove_plot_from_graph / close_origin).\n"
            "5. For LabTalk wipe/reset, use execute_labtalk(..., confirm=true) "
            "only when no typed tool exists.\n"
            "Never skip saving when the user still needs the current project."
        )
