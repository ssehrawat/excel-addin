"""Integration tests for the MCP server management API endpoints.

Tests CRUD operations via ``/mcp/servers`` routes.  All tests use an
isolated ``tmp_path`` MCP store so they do not affect each other.
"""

from __future__ import annotations

import pytest


def _create_payload(**overrides):
    base = {
        "name": "Test MCP",
        "baseUrl": "https://mcp.example.com",
        "autoRefresh": False,
    }
    base.update(overrides)
    return base


class TestMCPServerList:
    """GET /mcp/servers returns the server list."""

    def test_list_empty(self, test_client):
        response = test_client.get("/mcp/servers")
        assert response.status_code == 200
        assert response.json()["servers"] == []


class TestMCPServerCreate:
    """POST /mcp/servers registers a new server."""

    def test_create(self, test_client):
        response = test_client.post("/mcp/servers", json=_create_payload())
        assert response.status_code == 201
        server = response.json()["server"]
        assert server["name"] == "Test MCP"

    def test_create_persists(self, test_client):
        test_client.post("/mcp/servers", json=_create_payload())
        response = test_client.get("/mcp/servers")
        assert len(response.json()["servers"]) == 1

    def test_invalid_url_422(self, test_client):
        response = test_client.post(
            "/mcp/servers",
            json=_create_payload(baseUrl="not-a-url"),
        )
        assert response.status_code == 422

    def test_initial_status_unknown(self, test_client):
        response = test_client.post(
            "/mcp/servers", json=_create_payload()
        )
        server = response.json()["server"]
        assert server["status"] in ("unknown", "error")

    def test_response_shape(self, test_client):
        response = test_client.post("/mcp/servers", json=_create_payload())
        server = response.json()["server"]
        assert "id" in server
        assert "name" in server
        assert "baseUrl" in server
        assert "enabled" in server
        assert "tools" in server


class TestMCPServerUpdate:
    """PATCH /mcp/servers/{server_id} updates server fields."""

    def test_update_name(self, test_client):
        create_resp = test_client.post("/mcp/servers", json=_create_payload())
        server_id = create_resp.json()["server"]["id"]
        response = test_client.patch(
            f"/mcp/servers/{server_id}",
            json={"name": "Updated"},
        )
        assert response.status_code == 200
        assert response.json()["server"]["name"] == "Updated"

    def test_update_not_found(self, test_client):
        response = test_client.patch(
            "/mcp/servers/nonexistent",
            json={"name": "X"},
        )
        assert response.status_code == 404


class TestMCPServerDelete:
    """DELETE /mcp/servers/{server_id} removes a server."""

    def test_delete(self, test_client):
        create_resp = test_client.post("/mcp/servers", json=_create_payload())
        server_id = create_resp.json()["server"]["id"]
        response = test_client.delete(f"/mcp/servers/{server_id}")
        assert response.status_code == 204
        # Verify gone
        list_resp = test_client.get("/mcp/servers")
        assert len(list_resp.json()["servers"]) == 0

    def test_delete_not_found(self, test_client):
        response = test_client.delete("/mcp/servers/nonexistent")
        assert response.status_code == 404


class TestMCPServerListAll:
    """Verify list returns all created servers."""

    def test_list_returns_all(self, test_client):
        test_client.post("/mcp/servers", json=_create_payload(name="A"))
        test_client.post("/mcp/servers", json=_create_payload(name="B"))
        response = test_client.get("/mcp/servers")
        assert len(response.json()["servers"]) == 2
