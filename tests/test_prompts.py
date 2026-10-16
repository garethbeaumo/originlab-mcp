"""MCP prompt registration and workflow template contracts."""

from __future__ import annotations

import asyncio

from originlab_mcp.prompts import SERVER_INSTRUCTIONS, register_prompts
from originlab_mcp.server import mcp
from tests.fakes import DummyMCP

EXPECTED_PROMPTS = {
    "originlab_inspect_session",
    "originlab_csv_to_plot",
    "originlab_publication_style",
    "originlab_fit_curve",
    "originlab_safe_destructive",
}


class TestServerInstructions:
    def test_instructions_cover_safety_and_read_first(self) -> None:
        assert "read_origin_session" in SERVER_INSTRUCTIONS
        assert "execute_labtalk" in SERVER_INSTRUCTIONS
        assert "ALLOWED_ROOTS" in SERVER_INSTRUCTIONS
        assert "save_project" in SERVER_INSTRUCTIONS
        assert mcp.instructions == SERVER_INSTRUCTIONS


class TestPromptRegistration:
    def test_dummy_mcp_registers_expected_prompts(self) -> None:
        dummy = DummyMCP()
        register_prompts(dummy)
        assert set(dummy.prompts) == EXPECTED_PROMPTS

    def test_prompt_bodies_mention_typed_tools(self) -> None:
        dummy = DummyMCP()
        register_prompts(dummy)

        inspect_text = dummy.prompts["originlab_inspect_session"]()
        assert "originlab://session" in inspect_text
        assert "read_origin_session" in inspect_text

        csv_text = dummy.prompts["originlab_csv_to_plot"](
            file_path=r"C:\data\a.csv",
            plot_type="line",
        )
        assert "import_csv" in csv_text
        assert "create_plot" in csv_text
        assert "C:\\data\\a.csv" in csv_text
        assert "execute_labtalk" in csv_text

        pub_text = dummy.prompts["originlab_publication_style"](graph_name="Graph1")
        assert "apply_publication_style" in pub_text
        assert "Graph1" in pub_text

        fit_text = dummy.prompts["originlab_fit_curve"](function_name="Lorentz")
        assert "nonlinear_fit" in fit_text
        assert "Lorentz" in fit_text

        safe_text = dummy.prompts["originlab_safe_destructive"](
            operation="clear_worksheet"
        )
        assert "save_project" in safe_text
        assert "clear_worksheet" in safe_text

    def test_fastmcp_lists_all_prompts(self) -> None:
        prompts = asyncio.run(mcp.list_prompts())
        names = {p.name for p in prompts}
        assert names >= EXPECTED_PROMPTS
