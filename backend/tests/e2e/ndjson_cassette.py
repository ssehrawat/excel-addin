"""NDJSON cassette recording and replay for Layer 4 live LLM tests.

Records the full NDJSON event stream from a /chat response — every event
as emitted by the backend, including tool call round-trips. Replays the
saved stream deterministically so tests run without API keys.

Key exports:
- NDJSONCassette: dataclass holding a recorded scenario
- CassetteStore: save/load cassettes to disk
- extract_mutations: pull mutation payloads from an event list
- extract_final_answer: get the final message text from events
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional


@dataclass
class ToolRound:
    """One tool-call round-trip within a scenario."""

    round: int
    tool_calls: List[Dict[str, Any]]
    tool_results: List[Dict[str, Any]]
    events: List[Dict[str, Any]]


@dataclass
class NDJSONCassette:
    """A recorded NDJSON event stream for a single scenario.

    Stores the complete event history including any tool-call round-trips,
    so replay can exercise the full parsing pipeline.
    """

    scenario: str
    prompt: str
    provider: str
    recorded_at: str
    events: List[Dict[str, Any]] = field(default_factory=list)
    tool_rounds: List[ToolRound] = field(default_factory=list)

    def all_events(self) -> List[Dict[str, Any]]:
        """Return all events in order: initial + tool rounds + final."""
        result = list(self.events)
        for tr in self.tool_rounds:
            result.extend(tr.events)
        return result

    def to_dict(self) -> Dict[str, Any]:
        return {
            "scenario": self.scenario,
            "prompt": self.prompt,
            "provider": self.provider,
            "recorded_at": self.recorded_at,
            "events": self.events,
            "tool_rounds": [
                {
                    "round": tr.round,
                    "tool_calls": tr.tool_calls,
                    "tool_results": tr.tool_results,
                    "events": tr.events,
                }
                for tr in self.tool_rounds
            ],
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "NDJSONCassette":
        tool_rounds = [
            ToolRound(
                round=tr["round"],
                tool_calls=tr["tool_calls"],
                tool_results=tr["tool_results"],
                events=tr["events"],
            )
            for tr in data.get("tool_rounds", [])
        ]
        return cls(
            scenario=data["scenario"],
            prompt=data["prompt"],
            provider=data["provider"],
            recorded_at=data["recorded_at"],
            events=data["events"],
            tool_rounds=tool_rounds,
        )


def _safe_filename(scenario_name: str) -> str:
    """Convert a scenario name to a filesystem-safe filename."""
    slug = re.sub(r"[^a-z0-9_]", "_", scenario_name.lower())
    return re.sub(r"_+", "_", slug).strip("_")


class CassetteStore:
    """Manages reading and writing NDJSON cassette files.

    Args:
        cassette_dir: Directory where cassette JSON files are stored.
    """

    def __init__(self, cassette_dir: Path):
        self.cassette_dir = cassette_dir
        self.cassette_dir.mkdir(parents=True, exist_ok=True)

    def _path_for(self, scenario_name: str) -> Path:
        return self.cassette_dir / f"{_safe_filename(scenario_name)}.json"

    def has(self, scenario_name: str) -> bool:
        return self._path_for(scenario_name).exists()

    def save(self, cassette: NDJSONCassette) -> Path:
        """Save a cassette to disk. Returns the file path."""
        path = self._path_for(cassette.scenario)
        path.write_text(
            json.dumps(cassette.to_dict(), indent=2, default=str),
            encoding="utf-8",
        )
        return path

    def load(self, scenario_name: str) -> NDJSONCassette:
        """Load a cassette from disk.

        Raises:
            FileNotFoundError: If no cassette exists for this scenario.
        """
        path = self._path_for(scenario_name)
        if not path.exists():
            raise FileNotFoundError(
                f"No cassette for scenario '{scenario_name}'. "
                f"Run with --live to record it. Looked for: {path}"
            )
        data = json.loads(path.read_text(encoding="utf-8"))
        return NDJSONCassette.from_dict(data)

    def list_cassettes(self) -> List[Path]:
        return sorted(self.cassette_dir.glob("*.json"))


# ---------------------------------------------------------------------------
# Event extraction helpers
# ---------------------------------------------------------------------------


def extract_mutations(events: List[Dict[str, Any]]) -> Dict[str, List[Any]]:
    """Extract all mutation payloads from an event list.

    Returns:
        Dict with keys: cell_updates, format_updates, chart_inserts,
        pivot_table_inserts — each a flat list of mutation dicts.
    """
    mutations: Dict[str, List[Any]] = {
        "cell_updates": [],
        "format_updates": [],
        "chart_inserts": [],
        "pivot_table_inserts": [],
    }
    for event in events:
        etype = event.get("type", "")
        if etype in mutations:
            payload = event.get("payload", [])
            if isinstance(payload, list):
                mutations[etype].extend(payload)
    return mutations


def extract_final_answer(events: List[Dict[str, Any]]) -> str:
    """Extract the final message text from message_done events.

    Concatenates all message_done contents (there may be multiple messages).
    """
    parts = []
    for event in events:
        if event.get("type") == "message_done":
            content = event.get("payload", {}).get("content", "")
            if content:
                parts.append(content)
    return "\n".join(parts)


def extract_event_types(events: List[Dict[str, Any]]) -> List[str]:
    """Return the ordered list of event types from an event stream."""
    return [e.get("type", "unknown") for e in events]


def has_error_events(events: List[Dict[str, Any]]) -> bool:
    """Check if any error events are present in the stream."""
    return any(e.get("type") == "error" for e in events)
