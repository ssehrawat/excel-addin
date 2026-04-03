"""Layer 4 — Real LLM E2E tests with NDJSON cassette recording.

In live mode (--live flag): sends real prompts to a real LLM provider with
full workbook context from Financial_Market_Demo.xlsx, parses the NDJSON
stream, handles tool call round-trips, records the event stream to cassette
files, and runs structural assertions.

In replay mode (default): loads recorded cassettes and runs the same
structural assertions without API keys.

Usage:
    # Replay mode (no API keys, fast):
    pytest tests/e2e/test_live_llm.py -v

    # Live mode (needs API keys in .env):
    pytest tests/e2e/test_live_llm.py --live -v

    # Live with specific provider:
    pytest tests/e2e/test_live_llm.py --live --provider=anthropic -v

Key exports: none (test module)
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

import pytest

from tests.integration.workbook_fixture import WorkbookFixture, build_full_context, build_selection
from tests.e2e.ndjson_cassette import (
    CassetteStore,
    NDJSONCassette,
    ToolRound,
    extract_event_types,
    extract_final_answer,
    extract_mutations,
    has_error_events,
)

FIXTURES_DIR = Path(__file__).parent / "fixtures"
SCENARIOS_DIR = FIXTURES_DIR / "live_scenarios"
CASSETTES_DIR = FIXTURES_DIR / "cassettes"
WORKBOOK_DATA = FIXTURES_DIR / "workbook_data.json"

MAX_TOOL_ROUNDS = 3


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _load_workbook_data() -> Dict[str, Dict[str, Any]]:
    """Load the Financial_Market_Demo.xlsx data from the JSON fixture."""
    raw = json.loads(WORKBOOK_DATA.read_text(encoding="utf-8"))
    # WHY: openpyxl exports numeric values as int/float but the JSON
    # serialization may stringify some. Convert back where possible.
    result = {}
    for sheet_name, cells in raw.items():
        converted = {}
        for addr, val in cells.items():
            if isinstance(val, str):
                try:
                    val = int(val)
                except ValueError:
                    try:
                        val = float(val)
                    except ValueError:
                        pass
            converted[addr] = val
        result[sheet_name] = converted
    return result


def _load_scenarios() -> List[Dict[str, Any]]:
    """Load all scenario JSON files from the fixtures directory."""
    files = sorted(SCENARIOS_DIR.glob("*.json"))
    return [json.loads(f.read_text(encoding="utf-8")) for f in files]


def _build_fixture(scenario: Dict[str, Any], wb_data: Dict) -> WorkbookFixture:
    """Create a WorkbookFixture for a scenario using real workbook data.

    Supports an optional ``extraCells`` field in the scenario to inject
    additional cell data per sheet (e.g., a Rating column that would have
    been created by a prior scenario).
    """
    # WHY: Deep-copy so per-scenario overrides don't leak across scenarios
    import copy
    sheets = copy.deepcopy(wb_data)

    extra = scenario.get("extraCells", {})
    for sheet_name, cells in extra.items():
        if sheet_name not in sheets:
            sheets[sheet_name] = {}
        sheets[sheet_name].update(cells)

    return WorkbookFixture(
        sheets=sheets,
        active_sheet=scenario.get("activeSheet", "Portfolio"),
        selected_range=scenario.get("selectedRange", "A1"),
    )


def _build_payload(
    scenario: Dict[str, Any],
    fixture: WorkbookFixture,
    provider: str,
    tool_results: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """Build a full ChatRequest payload with real workbook context."""
    context = build_full_context(fixture)

    payload: Dict[str, Any] = {
        "prompt": scenario["prompt"],
        "provider": provider,
        "messages": [
            {
                "id": "msg-1",
                "role": "user",
                "kind": "message",
                "content": scenario["prompt"],
                "createdAt": "2024-01-01T00:00:00Z",
            }
        ],
        "selection": context["selection"],
        "workbookMetadata": context["workbookMetadata"],
        "userContext": context["userContext"],
        "activeSheetPreview": context["activeSheetPreview"],
    }

    if tool_results:
        payload["toolResults"] = tool_results

    return payload


def _stream_events(client, payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    """POST /chat with NDJSON accept, return parsed events."""
    response = client.post(
        "/chat",
        json=payload,
        headers={"Accept": "application/x-ndjson"},
    )
    assert response.status_code == 200, (
        f"Expected 200, got {response.status_code}: {response.text[:300]}"
    )
    lines = [line for line in response.text.strip().split("\n") if line.strip()]
    return [json.loads(line) for line in lines]


def _run_live_scenario(
    client,
    scenario: Dict[str, Any],
    fixture: WorkbookFixture,
    provider: str,
) -> NDJSONCassette:
    """Run a live scenario with tool call loop, return a cassette."""
    payload = _build_payload(scenario, fixture, provider)
    cassette = NDJSONCassette(
        scenario=scenario["name"],
        prompt=scenario["prompt"],
        provider=provider,
        recorded_at=datetime.now(timezone.utc).isoformat(),
    )

    for round_num in range(MAX_TOOL_ROUNDS + 1):
        events = _stream_events(client, payload)

        if round_num == 0:
            cassette.events = events
        else:
            # Events from tool round re-POST
            cassette.tool_rounds[-1].events = events

        # Check for tool_call_required
        tc_events = [e for e in events if e["type"] == "tool_call_required"]
        if not tc_events:
            break  # Final answer received

        # Execute tools against fixture
        tool_calls = tc_events[0]["payload"]
        tool_results = []
        for tc in tool_calls:
            result = fixture.execute_tool(tc["tool"], tc.get("args", {}))
            tool_results.append(result)

        cassette.tool_rounds.append(ToolRound(
            round=round_num + 1,
            tool_calls=tool_calls,
            tool_results=tool_results,
            events=[],  # filled on next iteration
        ))

        # Re-POST with tool results
        payload = _build_payload(scenario, fixture, provider, tool_results)

    return cassette


# ---------------------------------------------------------------------------
# Structural assertion evaluator
# ---------------------------------------------------------------------------


def _evaluate_assertions(
    scenario: Dict[str, Any],
    all_events: List[Dict[str, Any]],
) -> None:
    """Run structural assertions from a scenario against the event stream."""
    asserts = scenario.get("assertions", {})
    name = scenario["name"]
    mutations = extract_mutations(all_events)
    answer = extract_final_answer(all_events)
    event_types = extract_event_types(all_events)

    # no_errors
    if asserts.get("no_errors"):
        assert not has_error_events(all_events), (
            f"[{name}] Error events found in stream"
        )

    # has_answer
    if asserts.get("has_answer"):
        assert len(answer) > 0, f"[{name}] No answer in message_done events"

    # answer_mentions_any
    keywords = asserts.get("answer_mentions_any", [])
    if keywords:
        answer_lower = answer.lower()
        found = [kw for kw in keywords if kw.lower() in answer_lower]
        assert len(found) > 0, (
            f"[{name}] Answer does not mention any of {keywords}. "
            f"Answer: {answer[:200]}..."
        )

    # mutation_types_present
    expected_types = asserts.get("mutation_types_present", [])
    for mt in expected_types:
        assert mt in event_types, (
            f"[{name}] Expected '{mt}' event. Got types: {event_types}"
        )

    # cell_updates_min_count
    min_cu = asserts.get("cell_updates_min_count")
    if min_cu is not None:
        assert len(mutations["cell_updates"]) >= min_cu, (
            f"[{name}] Expected >= {min_cu} cell_updates, "
            f"got {len(mutations['cell_updates'])}"
        )

    # format_updates_min_count
    min_fu = asserts.get("format_updates_min_count")
    if min_fu is not None:
        assert len(mutations["format_updates"]) >= min_fu, (
            f"[{name}] Expected >= {min_fu} format_updates, "
            f"got {len(mutations['format_updates'])}"
        )

    # chart_inserts_min_count
    min_ci = asserts.get("chart_inserts_min_count")
    if min_ci is not None:
        assert len(mutations["chart_inserts"]) >= min_ci, (
            f"[{name}] Expected >= {min_ci} chart_inserts, "
            f"got {len(mutations['chart_inserts'])}"
        )

    # chart_has_title
    if asserts.get("chart_has_title"):
        charts = mutations["chart_inserts"]
        assert len(charts) > 0, f"[{name}] No charts to check title on"
        has_title = any(
            c.get("title") or c.get("chartTitle")
            for c in charts
        )
        assert has_title, f"[{name}] No chart has a title. Charts: {charts}"

    # pivot_inserts_min_count
    min_pi = asserts.get("pivot_inserts_min_count")
    if min_pi is not None:
        assert len(mutations["pivot_table_inserts"]) >= min_pi, (
            f"[{name}] Expected >= {min_pi} pivot_table_inserts, "
            f"got {len(mutations['pivot_table_inserts'])}"
        )


# ---------------------------------------------------------------------------
# Test suite
# ---------------------------------------------------------------------------


class TestLiveLLME2E:
    """Layer 4 E2E tests — live LLM calls or cassette replay."""

    @pytest.fixture(autouse=True)
    def _setup(self, request):
        self.is_live = request.config.getoption("--live", default=False)
        self.provider = request.config.getoption("--provider", default="openai")
        self.wb_data = _load_workbook_data()
        self.scenarios = _load_scenarios()
        self.cassette_store = CassetteStore(CASSETTES_DIR)

    def test_scenarios_loaded(self):
        """Sanity check: scenario files exist."""
        assert len(self.scenarios) > 0, "No scenario files found"

    def test_workbook_data_loaded(self):
        """Sanity check: workbook fixture data has all 5 sheets."""
        assert "Portfolio" in self.wb_data
        assert "Trade Log" in self.wb_data
        assert "Monthly Returns" in self.wb_data
        assert "Sector Allocation" in self.wb_data
        assert "Risk Metrics" in self.wb_data

    def _run_one_scenario(self, live_client, scenario: Dict[str, Any]):
        """Run a single scenario in live or replay mode and assert."""
        name = scenario["name"]

        if self.is_live:
            fixture = _build_fixture(scenario, self.wb_data)
            cassette = _run_live_scenario(
                live_client, scenario, fixture, self.provider
            )
            self.cassette_store.save(cassette)
            all_events = cassette.all_events()
        else:
            if not self.cassette_store.has(name):
                pytest.skip(
                    f"No cassette for '{name}'. Run with --live to record."
                )
            cassette = self.cassette_store.load(name)
            all_events = cassette.all_events()

        _evaluate_assertions(scenario, all_events)

    # WHY: Individual test methods per scenario so each runs independently.
    # A failure in one scenario doesn't block the others from executing
    # and recording their cassettes.

    def test_demo_s2_basic_qa(self, live_client):
        s = next((s for s in self.scenarios if s["name"] == "demo_s2_basic_qa"), None)
        if s:
            self._run_one_scenario(live_client, s)
        else:
            pytest.skip("Scenario not found")

    def test_demo_s4_rating_column(self, live_client):
        s = next((s for s in self.scenarios if s["name"] == "demo_s4_rating_column"), None)
        if s:
            self._run_one_scenario(live_client, s)
        else:
            pytest.skip("Scenario not found")

    def test_demo_s5_format_ratings(self, live_client):
        s = next((s for s in self.scenarios if s["name"] == "demo_s5_format_ratings"), None)
        if s:
            self._run_one_scenario(live_client, s)
        else:
            pytest.skip("Scenario not found")

    def test_demo_s6_bar_chart(self, live_client):
        s = next((s for s in self.scenarios if s["name"] == "demo_s6_bar_chart"), None)
        if s:
            self._run_one_scenario(live_client, s)
        else:
            pytest.skip("Scenario not found")

    def test_demo_s7_pivot_trade(self, live_client):
        s = next((s for s in self.scenarios if s["name"] == "demo_s7_pivot_trade"), None)
        if s:
            self._run_one_scenario(live_client, s)
        else:
            pytest.skip("Scenario not found")

    def test_demo_s8_cross_sheet(self, live_client):
        s = next((s for s in self.scenarios if s["name"] == "demo_s8_cross_sheet"), None)
        if s:
            self._run_one_scenario(live_client, s)
        else:
            pytest.skip("Scenario not found")

    def test_demo_s10_line_chart(self, live_client):
        s = next((s for s in self.scenarios if s["name"] == "demo_s10_line_chart"), None)
        if s:
            self._run_one_scenario(live_client, s)
        else:
            pytest.skip("Scenario not found")

    def test_cassette_integrity(self):
        """All saved cassettes are valid JSON with required fields."""
        cassettes = self.cassette_store.list_cassettes()
        if not cassettes:
            pytest.skip("No cassettes recorded yet.")

        for path in cassettes:
            data = json.loads(path.read_text(encoding="utf-8"))
            assert "scenario" in data, f"{path.name} missing 'scenario'"
            assert "prompt" in data, f"{path.name} missing 'prompt'"
            assert "events" in data, f"{path.name} missing 'events'"
            assert isinstance(data["events"], list)
            # Every cassette must have at least a done event
            types = [e.get("type") for e in data["events"]]
            assert "done" in types, (
                f"{path.name} missing 'done' event. Types: {types}"
            )

    def test_no_error_events_in_cassettes(self):
        """No recorded cassette should contain error events."""
        cassettes = self.cassette_store.list_cassettes()
        if not cassettes:
            pytest.skip("No cassettes recorded yet.")

        for path in cassettes:
            data = json.loads(path.read_text(encoding="utf-8"))
            cassette = NDJSONCassette.from_dict(data)
            all_events = cassette.all_events()
            assert not has_error_events(all_events), (
                f"{path.name} contains error events"
            )
