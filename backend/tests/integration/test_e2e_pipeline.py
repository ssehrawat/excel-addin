"""E2E pipeline integration tests — Layer 2.

Simulates a headless frontend: constructs ChatRequest payloads, POSTs to the
real backend (TestClient + mock provider), parses NDJSON event streams, and
exercises tool call round-trips. Validates event sequences, mutation payloads,
and serialization correctness without requiring Excel or LLM API keys.

Key exports: none (test module)
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Optional

import pytest

from .workbook_fixture import WorkbookFixture, build_selection

FIXTURES_DIR = Path(__file__).parent / "fixtures" / "e2e_scenarios"

MAX_TOOL_ROUNDS = 3


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _load_scenarios() -> List[Dict[str, Any]]:
    """Load all scenario JSON files from the fixtures directory."""
    files = sorted(FIXTURES_DIR.glob("*.json"))
    scenarios = []
    for f in files:
        scenarios.append(json.loads(f.read_text(encoding="utf-8")))
    return scenarios


def _build_fixture(scenario: Dict[str, Any]) -> WorkbookFixture:
    """Create a WorkbookFixture from a scenario's workbook definition."""
    wb = scenario.get("workbook", {})
    return WorkbookFixture(
        sheets=wb.get("sheets", {}),
        active_sheet=wb.get("activeSheet", "Sheet1"),
        selected_range=wb.get("selectedRange", "A1"),
    )


def _build_payload(
    scenario: Dict[str, Any],
    fixture: WorkbookFixture,
    tool_results: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """Build a ChatRequest payload from a scenario and fixture."""
    selection = build_selection(fixture)
    payload: Dict[str, Any] = {
        "prompt": scenario["prompt"],
        "provider": scenario.get("provider", "mock"),
        "messages": [
            {
                "id": "msg-1",
                "role": "user",
                "kind": "message",
                "content": scenario["prompt"],
                "createdAt": "2024-01-01T00:00:00Z",
            }
        ],
        "selection": selection,
    }

    wb_def = scenario.get("workbook", {})
    if wb_def.get("sheets"):
        sheets_meta = []
        for name, cells in wb_def["sheets"].items():
            sheets_meta.append({
                "id": f"sheet-{name}",
                "name": name,
                "index": len(sheets_meta),
                "maxRows": len(cells),
                "maxColumns": len(cells),
            })
        payload["workbookMetadata"] = {
            "success": True,
            "fileName": "test.xlsx",
            "sheetsMetadata": sheets_meta,
            "totalSheets": len(sheets_meta),
        }

    if tool_results:
        payload["toolResults"] = tool_results

    return payload


def _stream_events(test_client, payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    """POST /chat with NDJSON accept, return parsed events."""
    response = test_client.post(
        "/chat",
        json=payload,
        headers={"Accept": "application/x-ndjson"},
    )
    assert response.status_code == 200, (
        f"Expected 200, got {response.status_code}: {response.text[:200]}"
    )
    lines = [line for line in response.text.strip().split("\n") if line.strip()]
    events = []
    for line in lines:
        events.append(json.loads(line))
    return events


def _run_scenario_with_tool_loop(
    test_client,
    scenario: Dict[str, Any],
    fixture: WorkbookFixture,
) -> tuple[List[Dict[str, Any]], int]:
    """Run a scenario through the full tool-call loop.

    Returns:
        Tuple of (all_events, tool_round_count).
    """
    payload = _build_payload(scenario, fixture)
    all_events: List[Dict[str, Any]] = []
    tool_rounds = 0

    for _round in range(MAX_TOOL_ROUNDS + 1):
        events = _stream_events(test_client, payload)
        all_events.extend(events)

        # Check for tool_call_required
        tool_call_events = [e for e in events if e["type"] == "tool_call_required"]
        if not tool_call_events:
            break  # Final answer received

        tool_rounds += 1
        # Execute tools against fixture
        tool_results = []
        for tc in tool_call_events[0]["payload"]:
            result = fixture.execute_tool(tc["tool"], tc.get("args", {}))
            tool_results.append(result)

        # Re-POST with tool results
        payload = _build_payload(scenario, fixture, tool_results=tool_results)

    return all_events, tool_rounds


# ---------------------------------------------------------------------------
# Parametrized test suite
# ---------------------------------------------------------------------------


class TestE2EPipeline:
    """Data-driven E2E pipeline tests."""

    @pytest.fixture(autouse=True)
    def _scenarios(self):
        self.scenarios = _load_scenarios()

    def test_scenarios_loaded(self):
        """Sanity check: at least one scenario file exists."""
        assert len(self.scenarios) > 0, "No scenario files found"

    def test_all_scenarios_have_required_fields(self):
        """Validate scenario fixture structure."""
        for scenario in self.scenarios:
            assert "name" in scenario, f"Missing 'name' in {scenario}"
            assert "prompt" in scenario, f"Missing 'prompt' in {scenario}"

    def test_basic_chat(self, test_client):
        """Simple prompt with no selection returns a message and done."""
        payload = {
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
        events = _stream_events(test_client, payload)
        types = [e["type"] for e in events]
        assert "done" in types, f"Missing 'done' event. Got: {types}"
        assert any(
            t.startswith("message") for t in types
        ), f"No message events. Got: {types}"

    def test_chart_with_selection(self, test_client):
        """Prompt with 'chart' keyword and selection produces chart_inserts."""
        payload = {
            "prompt": "Create a chart",
            "provider": "mock",
            "messages": [
                {
                    "id": "msg-1",
                    "role": "user",
                    "kind": "message",
                    "content": "Create a chart",
                    "createdAt": "2024-01-01T00:00:00Z",
                }
            ],
            "selection": [
                {
                    "address": "Sheet1!A1:B3",
                    "values": [["X", "Y"], [1, 2], [3, 4]],
                    "worksheet": "Sheet1",
                }
            ],
        }
        events = _stream_events(test_client, payload)
        types = [e["type"] for e in events]
        assert "chart_inserts" in types, f"Expected chart_inserts. Got: {types}"

    def test_pivot_with_selection(self, test_client):
        """Prompt with 'pivot' keyword and selection produces pivot_table_inserts."""
        payload = {
            "prompt": "Create a pivot table",
            "provider": "mock",
            "messages": [
                {
                    "id": "msg-1",
                    "role": "user",
                    "kind": "message",
                    "content": "Create a pivot table",
                    "createdAt": "2024-01-01T00:00:00Z",
                }
            ],
            "selection": [
                {
                    "address": "Sheet1!A1:B3",
                    "values": [["Cat", "Amt"], ["A", 100], ["B", 200]],
                    "worksheet": "Sheet1",
                }
            ],
        }
        events = _stream_events(test_client, payload)
        types = [e["type"] for e in events]
        assert "pivot_table_inserts" in types, (
            f"Expected pivot_table_inserts. Got: {types}"
        )

    def test_format_with_color_keyword(self, test_client):
        """Prompt with 'color' keyword and selection produces format_updates."""
        payload = {
            "prompt": "Add color formatting",
            "provider": "mock",
            "messages": [
                {
                    "id": "msg-1",
                    "role": "user",
                    "kind": "message",
                    "content": "Add color formatting",
                    "createdAt": "2024-01-01T00:00:00Z",
                }
            ],
            "selection": [
                {
                    "address": "A1:B2",
                    "values": [[1, 2], [3, 4]],
                    "worksheet": "Sheet1",
                }
            ],
        }
        events = _stream_events(test_client, payload)
        types = [e["type"] for e in events]
        assert "format_updates" in types, (
            f"Expected format_updates. Got: {types}"
        )

    def test_tool_call_round_trip(self, test_client):
        """Prompt with 'lookup' triggers tool_call_required -> re-POST -> final answer."""
        fixture = WorkbookFixture(
            sheets={"Sheet1": {"A1": "Name", "B1": "Age", "A2": "Alice", "B2": 30}},
            active_sheet="Sheet1",
            selected_range="A1:B2",
        )
        scenario = {
            "name": "tool_call_round_trip",
            "prompt": "lookup the data in the sheet",
            "provider": "mock",
            "workbook": {
                "sheets": fixture.sheets,
                "activeSheet": fixture.active_sheet,
                "selectedRange": fixture.selected_range,
            },
        }
        all_events, tool_rounds = _run_scenario_with_tool_loop(
            test_client, scenario, fixture
        )
        types = [e["type"] for e in all_events]

        # WHY: The mock provider returns tool_call_required on first call,
        # then a final answer on the re-POST with tool results.
        assert tool_rounds >= 1, f"Expected at least 1 tool round, got {tool_rounds}"
        assert "tool_call_required" in types
        assert "done" in types

    def test_event_stream_ends_with_done(self, test_client):
        """Every NDJSON stream must end with a 'done' event."""
        payload = {
            "prompt": "Hello world",
            "provider": "mock",
            "messages": [
                {
                    "id": "msg-1",
                    "role": "user",
                    "kind": "message",
                    "content": "Hello world",
                    "createdAt": "2024-01-01T00:00:00Z",
                }
            ],
            "selection": [],
        }
        events = _stream_events(test_client, payload)
        assert events[-1]["type"] == "done", (
            f"Last event should be 'done', got '{events[-1]['type']}'"
        )

    def test_all_events_have_type(self, test_client):
        """Every NDJSON event must have a 'type' field."""
        payload = {
            "prompt": "Check event structure",
            "provider": "mock",
            "messages": [
                {
                    "id": "msg-1",
                    "role": "user",
                    "kind": "message",
                    "content": "Check event structure",
                    "createdAt": "2024-01-01T00:00:00Z",
                }
            ],
            "selection": [
                {
                    "address": "A1",
                    "values": [["test"]],
                    "worksheet": "Sheet1",
                }
            ],
        }
        events = _stream_events(test_client, payload)
        for event in events:
            assert "type" in event, f"Event missing 'type': {event}"

    def test_telemetry_present(self, test_client):
        """Telemetry event is emitted in the stream."""
        payload = {
            "prompt": "Check telemetry",
            "provider": "mock",
            "messages": [
                {
                    "id": "msg-1",
                    "role": "user",
                    "kind": "message",
                    "content": "Check telemetry",
                    "createdAt": "2024-01-01T00:00:00Z",
                }
            ],
            "selection": [],
        }
        events = _stream_events(test_client, payload)
        types = [e["type"] for e in events]
        assert "telemetry" in types, f"Expected telemetry event. Got: {types}"

    def test_message_sequence_order(self, test_client):
        """Message events follow start -> delta(s) -> done order."""
        payload = {
            "prompt": "Test ordering",
            "provider": "mock",
            "messages": [
                {
                    "id": "msg-1",
                    "role": "user",
                    "kind": "message",
                    "content": "Test ordering",
                    "createdAt": "2024-01-01T00:00:00Z",
                }
            ],
            "selection": [],
        }
        events = _stream_events(test_client, payload)
        msg_events = [
            e for e in events
            if e["type"] in ("message_start", "message_delta", "message_done")
        ]
        if msg_events:
            assert msg_events[0]["type"] == "message_start"
            assert msg_events[-1]["type"] == "message_done"

    def test_scenario_fixtures(self, test_client):
        """Run all scenario fixtures from the e2e_scenarios directory."""
        for scenario in self.scenarios:
            fixture = _build_fixture(scenario)
            all_events, tool_rounds = _run_scenario_with_tool_loop(
                test_client, scenario, fixture
            )
            types = [e["type"] for e in all_events]

            # Common assertions for all scenarios
            assert "done" in types, (
                f"[{scenario['name']}] Missing 'done'. Got: {types}"
            )

            # Check expected event types if specified
            expected_types = scenario.get("expectedEventTypes")
            if expected_types:
                for et in expected_types:
                    assert et in types, (
                        f"[{scenario['name']}] Missing '{et}'. Got: {types}"
                    )

            # Check expected tool rounds if specified
            expected_rounds = scenario.get("expectedToolRounds")
            if expected_rounds is not None:
                assert tool_rounds == expected_rounds, (
                    f"[{scenario['name']}] Expected {expected_rounds} tool "
                    f"rounds, got {tool_rounds}"
                )

            # Check expected mutations if specified
            expected_mutations = scenario.get("expectedMutations", {})
            for mut_type, expected in expected_mutations.items():
                # Map camelCase mutation names to event types
                event_type_map = {
                    "chartInserts": "chart_inserts",
                    "cellUpdates": "cell_updates",
                    "formatUpdates": "format_updates",
                    "pivotTableInserts": "pivot_table_inserts",
                }
                etype = event_type_map.get(mut_type, mut_type)
                mut_events = [e for e in all_events if e["type"] == etype]
                assert len(mut_events) > 0, (
                    f"[{scenario['name']}] Expected {etype} events. Got: {types}"
                )
