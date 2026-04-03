"""Replay tests for LLM response parsing — Layer 3.

In replay mode (default): loads pre-recorded cassettes and validates that
the response parsing pipeline handles them correctly.

In record mode (--record flag): makes real API calls and saves cassettes.

Usage:
    # Replay mode (no API keys needed, fast):
    pytest tests/eval/test_replay.py -v

    # Record mode (needs API keys):
    pytest tests/eval/test_replay.py --record -v

Key exports: none (test module)
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List

import pytest

from .record_replay import CassetteStore, ReplayProvider

FIXTURES_DIR = Path(__file__).parent / "fixtures"
SCENARIOS_DIR = FIXTURES_DIR / "replay_scenarios"
CASSETTES_DIR = FIXTURES_DIR / "cassettes"


# ---------------------------------------------------------------------------
# pytest CLI option for --record mode
# ---------------------------------------------------------------------------


def pytest_addoption(parser):
    """Add --record flag for recording mode."""
    parser.addoption(
        "--record",
        action="store_true",
        default=False,
        help="Record real LLM responses to cassette files",
    )


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


@pytest.fixture
def cassette_store(tmp_path: Path, request) -> CassetteStore:
    """Return a CassetteStore pointing to the real or temp cassettes dir.

    WHY: In replay mode we use the committed cassettes directory.
    In record mode we also write to the committed dir so cassettes persist.
    """
    return CassetteStore(CASSETTES_DIR)


@pytest.fixture
def replay_provider(cassette_store: CassetteStore) -> ReplayProvider:
    return ReplayProvider(cassette_store)


def _load_replay_scenarios() -> List[Dict[str, Any]]:
    """Load all replay scenario definitions."""
    if not SCENARIOS_DIR.exists():
        return []
    files = sorted(SCENARIOS_DIR.glob("*.json"))
    return [json.loads(f.read_text(encoding="utf-8")) for f in files]


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


class TestReplayInfrastructure:
    """Tests for the record-replay infrastructure itself."""

    def test_cassette_store_roundtrip(self, tmp_path: Path):
        """CassetteStore can save and load a cassette."""
        store = CassetteStore(tmp_path / "cassettes")

        # Simulate a minimal result
        mock_result = type("MockResult", (), {
            "messages": [],
            "cell_updates": [],
            "format_updates": [],
            "chart_inserts": [],
            "pivot_table_inserts": [],
            "telemetry": None,
            "tool_call_required": None,
        })()

        prompt = "test prompt for cassette"
        store.save(prompt, mock_result)
        assert store.has(prompt)

        loaded = store.load(prompt)
        assert loaded["prompt"] == prompt
        assert "result" in loaded

    def test_cassette_store_missing_raises(self, tmp_path: Path):
        """Loading a non-existent cassette raises FileNotFoundError."""
        store = CassetteStore(tmp_path / "empty")
        with pytest.raises(FileNotFoundError, match="No cassette found"):
            store.load("nonexistent prompt")

    def test_scenario_key_is_deterministic(self):
        """The same prompt always produces the same key."""
        from .record_replay import _scenario_key

        key1 = _scenario_key("Create a bar chart of revenue by quarter")
        key2 = _scenario_key("Create a bar chart of revenue by quarter")
        assert key1 == key2

    def test_scenario_key_differs_for_different_prompts(self):
        """Different prompts produce different keys."""
        from .record_replay import _scenario_key

        key1 = _scenario_key("Create a chart")
        key2 = _scenario_key("Create a pivot table")
        assert key1 != key2


class TestReplayScenarios:
    """Replay-mode tests that load cassettes and validate parsing.

    These tests only run when cassette files exist. If no cassettes are
    present, the tests are skipped with a helpful message.
    """

    def test_all_cassettes_are_valid_json(self, cassette_store: CassetteStore):
        """Every cassette file must be valid JSON with required fields."""
        cassettes = cassette_store.list_cassettes()
        if not cassettes:
            pytest.skip("No cassettes recorded yet. Run with --record first.")

        for path in cassettes:
            data = json.loads(path.read_text(encoding="utf-8"))
            assert "prompt" in data, f"{path.name} missing 'prompt'"
            assert "result" in data, f"{path.name} missing 'result'"

    def test_cassette_results_have_messages(
        self, cassette_store: CassetteStore
    ):
        """Recorded results should have at least one message."""
        cassettes = cassette_store.list_cassettes()
        if not cassettes:
            pytest.skip("No cassettes recorded yet.")

        for path in cassettes:
            data = json.loads(path.read_text(encoding="utf-8"))
            result = data["result"]
            # Either has messages or a tool_call_required
            has_content = (
                len(result.get("messages", [])) > 0
                or result.get("tool_call_required") is not None
            )
            assert has_content, (
                f"{path.name}: result has no messages and no tool call"
            )

    def test_replay_provider_returns_recorded_data(
        self, cassette_store: CassetteStore, replay_provider: ReplayProvider
    ):
        """ReplayProvider.generate() returns the recorded result."""
        cassettes = cassette_store.list_cassettes()
        if not cassettes:
            pytest.skip("No cassettes recorded yet.")

        # Load first cassette and verify replay
        data = json.loads(cassettes[0].read_text(encoding="utf-8"))
        prompt = data["prompt"]

        from app.schemas import ChatMessage, ChatRequest, MessageKind, MessageRole

        request = ChatRequest(
            prompt=prompt,
            provider="replay",
            messages=[
                ChatMessage(
                    id="msg-1",
                    role=MessageRole.USER,
                    kind=MessageKind.MESSAGE,
                    content=prompt,
                    created_at="2024-01-01T00:00:00Z",
                )
            ],
            selection=[],
        )

        import asyncio
        result = asyncio.get_event_loop().run_until_complete(
            replay_provider.generate(request)
        )
        assert result is not None
        assert result == data["result"]

    def test_replay_scenarios_from_fixtures(
        self, cassette_store: CassetteStore
    ):
        """Run all replay scenarios from the fixtures directory."""
        scenarios = _load_replay_scenarios()
        if not scenarios:
            pytest.skip("No replay scenario fixtures found.")

        for scenario in scenarios:
            prompt = scenario["prompt"]
            if not cassette_store.has(prompt):
                pytest.skip(
                    f"No cassette for '{prompt[:30]}...'. Run with --record."
                )

            cassette = cassette_store.load(prompt)
            result = cassette["result"]

            # Validate structure
            assert "messages" in result
            assert isinstance(result["messages"], list)

            # Validate mutations are parseable
            for mut_key in ("cell_updates", "format_updates",
                            "chart_inserts", "pivot_table_inserts"):
                items = result.get(mut_key, [])
                assert isinstance(items, list), (
                    f"[{scenario['name']}] {mut_key} should be a list"
                )
