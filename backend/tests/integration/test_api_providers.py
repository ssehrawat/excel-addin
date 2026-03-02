"""Integration tests for the /health and /providers API endpoints.

Uses the FastAPI ``TestClient`` fixture to exercise route handlers end-to-end
with the mock provider enabled.
"""

from __future__ import annotations


class TestHealthEndpoint:
    """GET /health returns a simple status response."""

    def test_health_200(self, test_client):
        response = test_client.get("/health")
        assert response.status_code == 200
        assert response.json()["status"] == "ok"


class TestProvidersEndpoint:
    """GET /providers returns the list of available providers."""

    def test_returns_list(self, test_client):
        response = test_client.get("/providers")
        assert response.status_code == 200
        data = response.json()
        assert "providers" in data
        assert isinstance(data["providers"], list)

    def test_includes_mock_when_enabled(self, test_client):
        response = test_client.get("/providers")
        providers = response.json()["providers"]
        ids = [p["id"] for p in providers]
        assert "mock" in ids

    def test_required_fields_present(self, test_client):
        response = test_client.get("/providers")
        for p in response.json()["providers"]:
            assert "id" in p
            assert "label" in p
            assert "description" in p
            assert "requiresKey" in p
