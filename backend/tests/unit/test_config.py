"""Unit tests for ``app.config.Settings`` and the ``get_settings`` singleton.

Validates default values, environment variable overrides, the ``extra="ignore"``
behaviour that prevents crashes from leftover env vars, and the LRU cache.
"""

from __future__ import annotations

import os

import pytest

from app.config import Settings, get_settings


class TestSettings:
    """Settings defaults and environment overrides."""

    def test_defaults(self, monkeypatch, tmp_path):
        """When no env vars or .env file are present, defaults apply."""
        # Clear all COPILOT_ env vars so the .env file doesn't leak
        for key in list(os.environ):
            if key.startswith("COPILOT_"):
                monkeypatch.delenv(key, raising=False)
        # Point env_file to a non-existent file so it doesn't read .env
        monkeypatch.chdir(tmp_path)
        s = Settings()
        assert s.mock_provider_enabled is True
        assert s.openai_model == "gpt-4o-mini"
        assert s.anthropic_model == "claude-3-5-sonnet-20240620"
        assert s.request_timeout_seconds == 120

    def test_env_override(self, monkeypatch):
        monkeypatch.setenv("COPILOT_OPENAI_MODEL", "gpt-5")
        s = Settings()
        assert s.openai_model == "gpt-5"

    def test_extra_vars_ignored(self, monkeypatch):
        """Lingering env vars from removed features should not crash startup."""
        monkeypatch.setenv("COPILOT_REMOVED_FEATURE", "true")
        s = Settings()
        assert s.app_name  # No crash

    def test_lru_cache_singleton(self):
        a = get_settings()
        b = get_settings()
        assert a is b
