"""Shared test fixtures for the Workbook Copilot backend.

Provides fixtures used across unit, integration, and eval tests:
- ``reset_settings_cache`` (autouse): clears the ``get_settings`` LRU cache
  between tests so settings mutations do not leak.
- ``tmp_mcp_store``: an ``MCPServerStore`` backed by a temporary file.
- ``tmp_mcp_service``: an ``MCPServerService`` backed by a temporary directory.
- ``test_client``: a FastAPI ``TestClient`` wired to a clean app instance with
  the mock provider enabled and MCP storage isolated to ``tmp_path``.
- ``minimal_chat_request``: a ready-to-use ``ChatRequest`` with all required
  fields populated.
"""

from __future__ import annotations

import os
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app.config import Settings, get_settings
from app.mcp import MCPServerService, MCPServerStore
from app.schemas import ChatRequest, ChatMessage, MessageRole, MessageKind


@pytest.fixture(autouse=True)
def reset_settings_cache():
    """Clear the ``get_settings`` LRU cache before and after each test."""
    get_settings.cache_clear()
    yield
    get_settings.cache_clear()


@pytest.fixture
def tmp_mcp_store(tmp_path: Path) -> MCPServerStore:
    """Return an ``MCPServerStore`` backed by a temporary JSON file."""
    return MCPServerStore(tmp_path / "mcp_servers.json")


@pytest.fixture
def tmp_mcp_service(tmp_path: Path) -> MCPServerService:
    """Return an ``MCPServerService`` with storage isolated to ``tmp_path``."""
    return MCPServerService(
        storage_path=tmp_path / "mcp_servers.json",
        request_timeout_seconds=5,
    )


@pytest.fixture
def test_client(tmp_path: Path, monkeypatch) -> TestClient:
    """Create a FastAPI ``TestClient`` with the mock provider enabled.

    MCP server storage is redirected to ``tmp_path`` so tests do not
    interfere with each other or with real data.
    """
    monkeypatch.setenv("COPILOT_MOCK_PROVIDER_ENABLED", "true")
    monkeypatch.setenv("COPILOT_MCP_CONFIG_PATH", str(tmp_path / "mcp_servers.json"))
    # Ensure no real API keys leak
    monkeypatch.delenv("COPILOT_OPENAI_API_KEY", raising=False)
    monkeypatch.delenv("COPILOT_ANTHROPIC_API_KEY", raising=False)

    get_settings.cache_clear()

    from app.main import create_app

    settings = get_settings()
    app = create_app(settings)
    return TestClient(app)


@pytest.fixture
def minimal_chat_request() -> ChatRequest:
    """Return a minimal ``ChatRequest`` suitable for mock provider calls."""
    return ChatRequest(
        prompt="Hello",
        provider="mock",
        messages=[
            ChatMessage(
                id="msg-1",
                role=MessageRole.USER,
                kind=MessageKind.MESSAGE,
                content="Hello",
                created_at="2024-01-01T00:00:00Z",
            )
        ],
        selection=[],
    )
