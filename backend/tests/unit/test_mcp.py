"""Unit tests for the MCP persistence layer and JSON-RPC client helpers.

Tests ``MCPServerStore`` (file I/O), ``MCPServerService`` (CRUD), and
``MCPJsonRpcClient`` helper methods (``_parse_tools``, ``_candidate_rpc_urls``,
``_build_headers``).  No real HTTP requests are made.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from fastapi import HTTPException

from app.mcp import (
    JsonRpcSession,
    MCPJsonRpcClient,
    MCPServerRecord,
    MCPServerService,
    MCPServerStore,
)
from app.schemas import MCPServerCreateRequest, MCPServerUpdateRequest


# ---- MCPServerStore ----

class TestMCPServerStore:
    """Thread-safe JSON store for MCP server records."""

    def test_load_empty_file(self, tmp_path: Path):
        store = MCPServerStore(tmp_path / "servers.json")
        assert store.load() == []

    def test_save_and_load_roundtrip(self, tmp_path: Path):
        store = MCPServerStore(tmp_path / "servers.json")
        record = MCPServerRecord(name="Test", base_url="https://example.com")
        store.save([record])
        loaded = store.load()
        assert len(loaded) == 1
        assert loaded[0].name == "Test"
        assert str(loaded[0].base_url) == "https://example.com/"

    def test_invalid_json_returns_empty(self, tmp_path: Path):
        path = tmp_path / "servers.json"
        path.write_text("not valid json", encoding="utf-8")
        store = MCPServerStore(path)
        assert store.load() == []

    def test_atomic_write(self, tmp_path: Path):
        """The store writes to a .tmp file then renames atomically."""
        store = MCPServerStore(tmp_path / "servers.json")
        record = MCPServerRecord(name="A", base_url="https://a.com")
        store.save([record])
        # After save, the .tmp file should be gone
        assert not (tmp_path / "servers.tmp").exists()
        assert (tmp_path / "servers.json").exists()

    def test_multiple_records(self, tmp_path: Path):
        store = MCPServerStore(tmp_path / "servers.json")
        records = [
            MCPServerRecord(name="A", base_url="https://a.com"),
            MCPServerRecord(name="B", base_url="https://b.com"),
        ]
        store.save(records)
        loaded = store.load()
        assert len(loaded) == 2

    def test_creates_parent_dir(self, tmp_path: Path):
        nested = tmp_path / "sub" / "dir" / "servers.json"
        store = MCPServerStore(nested)
        assert nested.parent.exists()

    def test_empty_file_returns_empty(self, tmp_path: Path):
        path = tmp_path / "servers.json"
        path.write_text("", encoding="utf-8")
        store = MCPServerStore(path)
        assert store.load() == []

    def test_skip_invalid_records(self, tmp_path: Path):
        path = tmp_path / "servers.json"
        # Write a list with one valid and one invalid record
        data = [
            {"name": "Valid", "base_url": "https://valid.com"},
            {"invalid_field_only": True},
        ]
        path.write_text(json.dumps(data), encoding="utf-8")
        store = MCPServerStore(path)
        records = store.load()
        # The valid record should be loaded
        assert any(r.name == "Valid" for r in records)


# ---- MCPServerService CRUD ----

class TestMCPServerServiceCRUD:
    """CRUD operations on MCPServerService."""

    def test_create(self, tmp_mcp_service: MCPServerService):
        payload = MCPServerCreateRequest(
            name="Test Server",
            base_url="https://example.com",
            auto_refresh=False,
        )
        record = tmp_mcp_service.create_server(payload)
        assert record.name == "Test Server"
        assert record.id

    def test_list_empty(self, tmp_mcp_service: MCPServerService):
        assert tmp_mcp_service.list_servers() == []

    def test_list_enabled_filters_disabled(self, tmp_mcp_service: MCPServerService):
        p1 = MCPServerCreateRequest(
            name="Enabled", base_url="https://a.com", enabled=True, auto_refresh=False,
        )
        p2 = MCPServerCreateRequest(
            name="Disabled", base_url="https://b.com", enabled=False, auto_refresh=False,
        )
        tmp_mcp_service.create_server(p1)
        tmp_mcp_service.create_server(p2)
        enabled = tmp_mcp_service.list_enabled_records()
        assert len(enabled) == 1
        assert enabled[0].name == "Enabled"

    def test_update_name(self, tmp_mcp_service: MCPServerService):
        payload = MCPServerCreateRequest(
            name="Original", base_url="https://x.com", auto_refresh=False,
        )
        record = tmp_mcp_service.create_server(payload)
        updated = tmp_mcp_service.update_server(
            record.id, MCPServerUpdateRequest(name="Renamed")
        )
        assert updated.name == "Renamed"

    def test_update_not_found(self, tmp_mcp_service: MCPServerService):
        with pytest.raises(HTTPException) as exc_info:
            tmp_mcp_service.update_server(
                "nonexistent", MCPServerUpdateRequest(name="X")
            )
        assert exc_info.value.status_code == 404

    def test_delete(self, tmp_mcp_service: MCPServerService):
        payload = MCPServerCreateRequest(
            name="ToDelete", base_url="https://d.com", auto_refresh=False,
        )
        record = tmp_mcp_service.create_server(payload)
        tmp_mcp_service.delete_server(record.id)
        assert len(tmp_mcp_service.list_servers()) == 0

    def test_delete_not_found(self, tmp_mcp_service: MCPServerService):
        with pytest.raises(HTTPException) as exc_info:
            tmp_mcp_service.delete_server("nonexistent")
        assert exc_info.value.status_code == 404

    def test_to_public_masks_api_key(self, tmp_mcp_service: MCPServerService):
        record = MCPServerRecord(
            name="Keyed", base_url="https://k.com", api_key="secret-key-123"
        )
        public = tmp_mcp_service.to_public(record)
        # MCPServerPublic does not have an api_key field
        assert not hasattr(public, "api_key")


# ---- MCPJsonRpcClient._parse_tools ----

class TestParseTools:
    """Tool list parsing from JSON-RPC responses."""

    def _client(self) -> MCPJsonRpcClient:
        return MCPJsonRpcClient(timeout_seconds=5)

    def test_valid_list(self):
        payload = [
            {"name": "find", "description": "Find items", "inputSchema": {"type": "object"}},
        ]
        tools = self._client()._parse_tools(payload)
        assert len(tools) == 1
        assert tools[0].name == "find"
        assert tools[0].description == "Find items"

    def test_skips_missing_name(self):
        payload = [{"description": "No name"}]
        tools = self._client()._parse_tools(payload)
        assert len(tools) == 0

    def test_annotation_description_fallback(self):
        payload = [
            {"name": "t", "annotations": {"description": "From annotations"}},
        ]
        tools = self._client()._parse_tools(payload)
        assert tools[0].description == "From annotations"

    def test_non_dict_schema_defaults(self):
        payload = [{"name": "t", "inputSchema": "not a dict"}]
        tools = self._client()._parse_tools(payload)
        assert tools[0].input_schema == {"type": "object"}


# ---- _candidate_rpc_urls ----

class TestCandidateRpcUrls:
    """URL candidate generation for MCP JSON-RPC transport."""

    def _client(self) -> MCPJsonRpcClient:
        return MCPJsonRpcClient(timeout_seconds=5)

    def test_bare_base_adds_mcp(self):
        server = MCPServerRecord(name="s", base_url="https://example.com")
        urls = self._client()._candidate_rpc_urls(server)
        assert "https://example.com" in urls
        assert "https://example.com/mcp" in urls

    def test_with_path_no_extra(self):
        server = MCPServerRecord(name="s", base_url="https://example.com/api/v1")
        urls = self._client()._candidate_rpc_urls(server)
        assert len(urls) == 1
        assert "https://example.com/api/v1" in urls

    def test_no_duplicates(self):
        server = MCPServerRecord(name="s", base_url="https://example.com/mcp")
        urls = self._client()._candidate_rpc_urls(server)
        assert len(urls) == len(set(urls))

    def test_trailing_slash_stripped(self):
        server = MCPServerRecord(name="s", base_url="https://example.com/")
        urls = self._client()._candidate_rpc_urls(server)
        assert all(not u.endswith("/") for u in urls)


# ---- _build_headers ----

class TestBuildHeaders:
    """HTTP header construction for MCP requests."""

    def _client(self) -> MCPJsonRpcClient:
        return MCPJsonRpcClient(timeout_seconds=5)

    def test_no_auth_no_session(self):
        server = MCPServerRecord(name="s", base_url="https://x.com")
        headers = self._client()._build_headers(server, None)
        assert "Authorization" not in headers
        assert "Mcp-Session-Id" not in headers

    def test_with_api_key(self):
        server = MCPServerRecord(name="s", base_url="https://x.com", api_key="key123")
        headers = self._client()._build_headers(server, None)
        assert headers["Authorization"] == "Bearer key123"

    def test_with_session(self):
        server = MCPServerRecord(name="s", base_url="https://x.com")
        session = JsonRpcSession(
            session_id="sess1",
            protocol_version="2025-03-26",
            endpoint="https://x.com/mcp",
        )
        headers = self._client()._build_headers(server, session)
        assert headers["Mcp-Session-Id"] == "sess1"
        assert headers["Mcp-Protocol-Version"] == "2025-03-26"
