"""Integration tests for the POST /chat API endpoint.

Exercises both the non-streaming (JSON) and streaming (NDJSON) paths using
the mock provider.
"""

from __future__ import annotations

import json

import pytest


def _chat_payload(**overrides):
    """Build a minimal chat request payload."""
    base = {
        "prompt": "Hello",
        "provider": "mock",
        "messages": [
            {
                "id": "msg-1",
                "role": "user",
                "kind": "message",
                "content": "Hello",
                "createdAt": "2024-01-01T00:00:00Z",
            }
        ],
        "selection": [],
    }
    base.update(overrides)
    return base


class TestChatNonStreaming:
    """POST /chat without NDJSON accept header returns a JSON response."""

    def test_returns_json(self, test_client):
        response = test_client.post(
            "/chat",
            json=_chat_payload(),
            headers={"Accept": "application/json"},
        )
        assert response.status_code == 200
        data = response.json()
        assert "messages" in data


class TestChatStreaming:
    """POST /chat with NDJSON accept header returns an event stream."""

    def test_returns_ndjson(self, test_client):
        response = test_client.post(
            "/chat",
            json=_chat_payload(),
            headers={"Accept": "application/x-ndjson"},
        )
        assert response.status_code == 200
        assert "application/x-ndjson" in response.headers.get("content-type", "")

    def test_events_parseable(self, test_client):
        response = test_client.post(
            "/chat",
            json=_chat_payload(),
            headers={"Accept": "application/x-ndjson"},
        )
        lines = [line for line in response.text.strip().split("\n") if line.strip()]
        for line in lines:
            event = json.loads(line)
            assert "type" in event

    def test_contains_done(self, test_client):
        response = test_client.post(
            "/chat",
            json=_chat_payload(),
            headers={"Accept": "application/x-ndjson"},
        )
        lines = [line for line in response.text.strip().split("\n") if line.strip()]
        types = [json.loads(line)["type"] for line in lines]
        assert "done" in types


class TestChatValidation:
    """Request validation for POST /chat."""

    def test_missing_provider_422(self, test_client):
        response = test_client.post(
            "/chat",
            json={"prompt": "hi", "messages": [], "selection": []},
        )
        assert response.status_code == 422


class TestChatWithContext:
    """POST /chat with workbook metadata and tool results."""

    def test_with_workbook_metadata(self, test_client):
        payload = _chat_payload(
            workbookMetadata={
                "success": True,
                "fileName": "test.xlsx",
                "sheetsMetadata": [],
                "totalSheets": 0,
            }
        )
        response = test_client.post(
            "/chat",
            json=payload,
            headers={"Accept": "application/x-ndjson"},
        )
        assert response.status_code == 200

    def test_with_tool_results(self, test_client):
        payload = _chat_payload(
            toolResults=[
                {
                    "id": "tr-1",
                    "tool": "get_xl_range_as_csv",
                    "result": "a,b\n1,2",
                }
            ]
        )
        response = test_client.post(
            "/chat",
            json=payload,
            headers={"Accept": "application/x-ndjson"},
        )
        assert response.status_code == 200
