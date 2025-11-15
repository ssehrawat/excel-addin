from __future__ import annotations

import logging
import time
from typing import Any, AsyncIterator, Dict, TypedDict

from langgraph.graph import END, StateGraph

from .providers import (
    ProviderResult,
    ProviderStreamEvent,
    available_providers,
    create_provider,
)
from .schemas import ChatRequest, ChatResponse

logger = logging.getLogger(__name__)


class OrchestratorState(TypedDict, total=False):
    request: ChatRequest
    provider_id: str
    provider_result: ProviderResult
    response: ChatResponse
    metadata: Dict[str, Any]


class LangGraphOrchestrator:
    def __init__(self) -> None:
        self.graph = self._build_graph()
        self._compiled = self.graph.compile()

    def _build_graph(self) -> StateGraph:
        graph: StateGraph = StateGraph(OrchestratorState)

        async def attach_request(state: OrchestratorState) -> OrchestratorState:
            if "request" not in state:
                raise ValueError("Request missing from orchestration state.")
            logger.debug(
                "Received chat request for provider '%s'", state["request"].provider
            )
            return {
                "provider_id": state["request"].provider.lower(),
                "metadata": {"started_at": time.time()},
            }

        async def invoke_provider(state: OrchestratorState) -> OrchestratorState:
            provider_id = state["provider_id"]
            provider = create_provider(provider_id)
            result = await provider.generate(state["request"])
            metadata = state.get("metadata", {})
            metadata["provider_label"] = getattr(provider, "label", provider_id)
            metadata["available_providers"] = [
                item["id"] for item in available_providers()
            ]
            return {"provider_result": result, "metadata": metadata}

        def finalize(state: OrchestratorState) -> OrchestratorState:
            result = state["provider_result"]
            response = result.to_response()
            metadata = state.get("metadata", {})
            started_at = metadata.get("started_at")
            if started_at:
                elapsed = (time.time() - started_at) * 1000
                if response.telemetry:
                    response.telemetry.latency_ms = elapsed
            return {"response": response}

        graph.add_node("attach_request", attach_request)
        graph.add_node("invoke_provider", invoke_provider)
        graph.add_node("finalize", finalize)

        graph.set_entry_point("attach_request")
        graph.add_edge("attach_request", "invoke_provider")
        graph.add_edge("invoke_provider", "finalize")
        graph.add_edge("finalize", END)

        return graph

    async def run(self, request: ChatRequest) -> ChatResponse:
        result_state = await self._compiled.ainvoke({"request": request})
        return result_state["response"]

    async def stream(self, request: ChatRequest) -> AsyncIterator[ProviderStreamEvent]:
        provider_id = request.provider.lower()
        started_at = time.time()
        telemetry_payload: Dict[str, Any] = {}

        try:
            provider = create_provider(provider_id)
        except Exception as exc:  # pragma: no cover
            logger.exception("Failed to initialize provider '%s'", provider_id)
            yield {
                "type": "error",
                "payload": {"message": str(exc)},
            }
            return

        try:
            async for event in provider.stream(request):
                if event.get("type") == "telemetry":
                    payload = event.get("payload") or {}
                    if isinstance(payload, dict):
                        telemetry_payload.update(payload)
                    else:
                        telemetry_payload.update(dict(payload))
                    continue
                yield event
        except Exception as exc:  # pragma: no cover
            logger.exception(
                "Error while streaming response from provider '%s'", provider_id
            )
            yield {
                "type": "error",
                "payload": {"message": str(exc)},
            }
            return

        latency_ms = (time.time() - started_at) * 1000
        telemetry = dict(telemetry_payload)
        telemetry.setdefault("provider", getattr(provider, "id", provider_id))
        if "model" not in telemetry and hasattr(provider, "model"):
            telemetry["model"] = getattr(provider, "model")
        telemetry["latency_ms"] = latency_ms

        yield {"type": "telemetry", "payload": telemetry}
        yield {"type": "done"}
