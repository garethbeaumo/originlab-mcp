"""Shared pytest fixtures for OriginLab MCP tests."""

from __future__ import annotations

import importlib

import pytest

from originlab_mcp.origin_manager import OriginManager
from tests.fakes import DummyMCP, FakeOrigin, attach_fake_origin


@pytest.fixture(autouse=True)
def _never_connect_to_real_origin(monkeypatch: pytest.MonkeyPatch):
    """Portable tests must never attach to the user's desktop Origin session.

    A validation-only tool can run an autosave preflight before failing. On
    Windows the installed bridge would otherwise be imported in such a test,
    and OriginExt can abort at exit when no Origin application is installed.
    Keep fake reconnect paths and production manager logic intact.
    """
    original = importlib.import_module

    def import_without_origin(name, package=None):
        if name == "originpro":
            raise ImportError("Live Origin is disabled in portable tests; attach FakeOrigin")
        return original(name, package)

    monkeypatch.setattr(importlib, "import_module", import_without_origin)


@pytest.fixture(autouse=True)
def _disable_dispatch_timeout_by_default(monkeypatch: pytest.MonkeyPatch):
    """Keep the suite fast; timeout-specific tests re-enable the budget."""
    monkeypatch.setenv("ORIGINLAB_MCP_DISPATCH_TIMEOUT", "off")
    monkeypatch.setenv("ORIGINLAB_MCP_AUTOSAVE_INTERVAL", "off")


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
