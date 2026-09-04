"""Fake Origin backends and MCP doubles for Origin-less tests."""

from tests.fakes.mcp import DummyMCP, attach_fake_origin
from tests.fakes.origin import FakeOrigin

__all__ = ["DummyMCP", "FakeOrigin", "attach_fake_origin"]
