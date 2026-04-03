"""Record-replay infrastructure for LLM response testing — Layer 3.

Captures real LLM responses at the ProviderResult level and replays them
deterministically. This avoids the brittleness of raw HTTP recording (vcrpy)
while still testing real response parsing and mutation extraction.

Key exports:
- RecordingProvider: wraps a real provider, saves responses to cassette files
- ReplayProvider: loads saved cassettes, returns recorded responses
- CassetteStore: manages reading/writing cassette files

WHY: Record at ProviderResult level instead of HTTP level because:
1. API endpoints change across versions
2. Auth tokens expire
3. Response headers shift
4. We only care about the LLM's structured response, not transport details
"""

from __future__ import annotations

import hashlib
import json
import logging
from dataclasses import asdict, fields, is_dataclass
from pathlib import Path
from typing import Any, AsyncIterator, Dict, List, Optional

from app.schemas import (
    ChatMessage,
    ChatRequest,
    CellUpdate,
    ChartInsert,
    FormatUpdate,
    PivotTableInsert,
    Telemetry,
)

logger = logging.getLogger(__name__)


def _scenario_key(prompt: str) -> str:
    """Generate a stable key for a scenario based on its prompt.

    WHY: Use a hash of the prompt to create filenames that are safe for
    all filesystems while still being deterministic.
    """
    h = hashlib.sha256(prompt.encode("utf-8")).hexdigest()[:12]
    # Also create a readable slug from the first few words
    slug = "_".join(prompt.lower().split()[:5])
    slug = "".join(c if c.isalnum() or c == "_" else "" for c in slug)[:40]
    return f"{slug}_{h}"


def _serialize_result(result: Any) -> Dict[str, Any]:
    """Serialize a ProviderResult-like object to a JSON-safe dict.

    WHY: ProviderResult uses dataclasses and Pydantic models which need
    special handling for JSON serialization.
    """
    if result is None:
        return None

    data: Dict[str, Any] = {}

    # Messages
    messages = getattr(result, "messages", [])
    data["messages"] = [
        m.model_dump(by_alias=True) if hasattr(m, "model_dump") else str(m)
        for m in messages
    ]

    # Mutations
    for field_name in ("cell_updates", "format_updates", "chart_inserts",
                       "pivot_table_inserts"):
        items = getattr(result, field_name, [])
        data[field_name] = [
            item.model_dump(by_alias=True) if hasattr(item, "model_dump")
            else str(item)
            for item in items
        ]

    # Telemetry
    telemetry = getattr(result, "telemetry", None)
    if telemetry and hasattr(telemetry, "model_dump"):
        data["telemetry"] = telemetry.model_dump()
    else:
        data["telemetry"] = None

    # Tool call
    tool_call = getattr(result, "tool_call_required", None)
    if tool_call and hasattr(tool_call, "model_dump"):
        data["tool_call_required"] = tool_call.model_dump()
    else:
        data["tool_call_required"] = None

    return data


class CassetteStore:
    """Manages reading and writing cassette files.

    Each cassette is a JSON file containing the serialized ProviderResult
    for a given scenario prompt.

    Args:
        cassette_dir: Directory where cassette files are stored.
    """

    def __init__(self, cassette_dir: Path):
        self.cassette_dir = cassette_dir
        self.cassette_dir.mkdir(parents=True, exist_ok=True)

    def _path_for(self, prompt: str) -> Path:
        key = _scenario_key(prompt)
        return self.cassette_dir / f"{key}.json"

    def has(self, prompt: str) -> bool:
        """Check if a cassette exists for this prompt."""
        return self._path_for(prompt).exists()

    def save(self, prompt: str, result: Any) -> Path:
        """Save a ProviderResult to a cassette file.

        Args:
            prompt: The scenario prompt (used as key).
            result: The ProviderResult to save.

        Returns:
            Path to the saved cassette file.
        """
        path = self._path_for(prompt)
        data = {
            "prompt": prompt,
            "result": _serialize_result(result),
        }
        path.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")
        logger.info("Saved cassette: %s", path.name)
        return path

    def load(self, prompt: str) -> Dict[str, Any]:
        """Load a cassette for a given prompt.

        Args:
            prompt: The scenario prompt.

        Returns:
            Dict with 'prompt' and 'result' keys.

        Raises:
            FileNotFoundError: If no cassette exists for this prompt.
        """
        path = self._path_for(prompt)
        if not path.exists():
            raise FileNotFoundError(
                f"No cassette found for prompt: '{prompt[:50]}...' "
                f"(looked for {path.name}). Run with --record to create it."
            )
        data = json.loads(path.read_text(encoding="utf-8"))
        return data

    def list_cassettes(self) -> List[Path]:
        """List all cassette files."""
        return sorted(self.cassette_dir.glob("*.json"))


class RecordingProvider:
    """Wraps a real provider and saves its responses to cassette files.

    Used during recording mode to capture real LLM responses that can
    later be replayed deterministically.

    Args:
        real_provider: The actual LLM provider to delegate to.
        store: CassetteStore for saving responses.
    """

    def __init__(self, real_provider: Any, store: CassetteStore):
        self.real_provider = real_provider
        self.store = store
        self.id = real_provider.id
        self.label = f"Recording({real_provider.label})"

    async def generate(self, request: ChatRequest) -> Any:
        """Generate via the real provider and save the result."""
        result = await self.real_provider.generate(request)
        self.store.save(request.prompt, result)
        return result


class ReplayProvider:
    """Replays recorded ProviderResults from cassette files.

    Used during replay mode to test response parsing and mutation extraction
    without making real API calls.

    Args:
        store: CassetteStore for loading recorded responses.
    """

    id = "replay"
    label = "Replay provider"
    description = "Replays recorded LLM responses from cassette files."
    requires_key = False

    def __init__(self, store: CassetteStore):
        self.store = store

    def _reconstruct_messages(
        self, raw_messages: List[Dict[str, Any]]
    ) -> List[ChatMessage]:
        """Reconstruct ChatMessage objects from serialized dicts."""
        messages = []
        for m in raw_messages:
            messages.append(ChatMessage(**m))
        return messages

    def _reconstruct_mutations(
        self, data: Dict[str, Any]
    ) -> Dict[str, List[Any]]:
        """Reconstruct mutation lists from serialized dicts."""
        result = {}
        if data.get("cell_updates"):
            result["cell_updates"] = [
                CellUpdate(**cu) for cu in data["cell_updates"]
            ]
        if data.get("format_updates"):
            result["format_updates"] = [
                FormatUpdate(**fu) for fu in data["format_updates"]
            ]
        if data.get("chart_inserts"):
            result["chart_inserts"] = [
                ChartInsert(**ci) for ci in data["chart_inserts"]
            ]
        if data.get("pivot_table_inserts"):
            result["pivot_table_inserts"] = [
                PivotTableInsert(**pi) for pi in data["pivot_table_inserts"]
            ]
        return result

    async def generate(self, request: ChatRequest) -> Dict[str, Any]:
        """Load and return a recorded response for this prompt.

        Returns:
            A dict matching ProviderResult structure with messages and mutations.
        """
        cassette = self.store.load(request.prompt)
        return cassette["result"]
