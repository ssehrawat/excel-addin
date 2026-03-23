"""Integration tests for the POST /transcribe endpoint.

Exercises file upload validation, API-key gating, and successful
transcription via a mocked OpenAI Whisper client.
"""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from fastapi.testclient import TestClient
from pathlib import Path

from app.config import get_settings


def _make_client(tmp_path: Path, monkeypatch, *, api_key: str | None = None) -> TestClient:
    """Create a TestClient with optional COPILOT_OPENAI_API_KEY override."""
    monkeypatch.setenv("COPILOT_MOCK_PROVIDER_ENABLED", "true")
    monkeypatch.setenv("COPILOT_MCP_CONFIG_PATH", str(tmp_path / "mcp_servers.json"))
    if api_key is not None:
        monkeypatch.setenv("COPILOT_OPENAI_API_KEY", api_key)
    else:
        # Set to empty string rather than deleting — pydantic-settings
        # would still read from .env if the env var is absent.
        monkeypatch.setenv("COPILOT_OPENAI_API_KEY", "")
    monkeypatch.delenv("COPILOT_ANTHROPIC_API_KEY", raising=False)

    get_settings.cache_clear()
    from app.main import create_app

    settings = get_settings()
    app = create_app(settings)
    return TestClient(app)


class TestTranscribeValidation:
    """POST /transcribe input validation."""

    def test_transcribe_no_file_422(self, test_client):
        """Missing audio file returns 422."""
        response = test_client.post("/transcribe")
        assert response.status_code == 422

    def test_transcribe_no_api_key_400(self, tmp_path, monkeypatch):
        """POST with file but no COPILOT_OPENAI_API_KEY returns 400."""
        client = _make_client(tmp_path, monkeypatch, api_key=None)
        response = client.post(
            "/transcribe",
            files={"audio": ("test.webm", b"fake-audio", "audio/webm")},
        )
        assert response.status_code == 400
        assert "OpenAI API key required" in response.json()["detail"]


class TestTranscribeSuccess:
    """POST /transcribe with mocked OpenAI Whisper."""

    @patch("app.main.AsyncOpenAI" if False else "openai.AsyncOpenAI")
    def test_transcribe_success(self, mock_openai_cls, tmp_path, monkeypatch):
        """Successful transcription returns {"text": "..."}."""
        # Build a mock that AsyncOpenAI() returns.
        mock_transcript = MagicMock()
        mock_transcript.text = "hello world"

        mock_client = MagicMock()
        mock_client.audio.transcriptions.create = AsyncMock(return_value=mock_transcript)
        mock_openai_cls.return_value = mock_client

        client = _make_client(tmp_path, monkeypatch, api_key="sk-test-key")
        response = client.post(
            "/transcribe",
            files={"audio": ("recording.webm", b"fake-audio-bytes", "audio/webm")},
        )
        assert response.status_code == 200
        data = response.json()
        assert data == {"text": "hello world"}

    @patch("openai.AsyncOpenAI")
    def test_transcribe_accepts_webm(self, mock_openai_cls, tmp_path, monkeypatch):
        """POST with audio/webm content type processes successfully."""
        mock_transcript = MagicMock()
        mock_transcript.text = "webm transcription"

        mock_client = MagicMock()
        mock_client.audio.transcriptions.create = AsyncMock(return_value=mock_transcript)
        mock_openai_cls.return_value = mock_client

        client = _make_client(tmp_path, monkeypatch, api_key="sk-test-key")
        response = client.post(
            "/transcribe",
            files={"audio": ("audio.webm", b"\x1a\x45\xdf\xa3", "audio/webm")},
        )
        assert response.status_code == 200
        assert response.json()["text"] == "webm transcription"

        # Verify the create call was made with the right model.
        mock_client.audio.transcriptions.create.assert_called_once()
        call_kwargs = mock_client.audio.transcriptions.create.call_args
        assert call_kwargs.kwargs["model"] == "whisper-1"
