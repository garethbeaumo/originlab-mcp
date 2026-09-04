"""Shared pytest fixtures for OriginLab MCP tests."""

from __future__ import annotations

import pytest

from originlab_mcp.origin_manager import OriginManager
from tests.fakes import DummyMCP, FakeOrigin, attach_fake_origin


@pytest.fixture
def fresh_manager():
    return OriginManager(auto_recover_active=False, idle_timeout=0)


@pytest.fixture
def dummy_mcp():
    return DummyMCP()


@pytest.fixture
def fake_origin(fresh_manager):
    origin = FakeOrigin()
    attach_fake_origin(fresh_manager, origin)
    return origin
