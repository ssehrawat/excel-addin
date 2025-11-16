from __future__ import annotations

import json
import logging
import re
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, AsyncIterator, Dict, List, Optional, Sequence, TypedDict
from uuid import uuid4

from fastapi import HTTPException
from langgraph.graph import END, StateGraph

from .mcp import MCPServerRecord, MCPServerService, ToolInvocationResult
from .router import MCPRouter, RouterSelection
from .providers import (
    ProviderResult,
    ProviderStreamEvent,
    available_providers,
    create_provider,
)
from .schemas import (
    ChatMessage,
    ChatRequest,
    ChatResponse,
    MCPTool,
    MessageKind,
    MessageRole,
    Telemetry,
)

logger = logging.getLogger(__name__)


@dataclass
class PlannedToolCall:
    server: MCPServerRecord
    tool_name: str
    rationale: str


@dataclass
class PlanResult:
    steps: List[str] = field(default_factory=list)
    tool_calls: List[PlannedToolCall] = field(default_factory=list)
    clarification: Optional[str] = None


class OrchestratorState(TypedDict, total=False):
    request: ChatRequest
    provider_id: str
    provider_result: ProviderResult
    response: ChatResponse
    metadata: Dict[str, Any]


class LangGraphOrchestrator:
    MAX_TOOL_CALLS = 3

    def __init__(
        self,
        mcp_service: Optional[MCPServerService] = None,
        router: Optional[MCPRouter] = None,
    ) -> None:
        self._mcp_service = mcp_service
        self._router = router
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
        plan = await self._build_plan(request)
        pre_messages: List[ChatMessage] = []
        context_messages: List[ChatMessage] = []
        telemetry_updates: Dict[str, Any] = {}

        if plan.clarification:
            clarification_message = self._message(
                MessageKind.MESSAGE,
                plan.clarification,
            )
            pre_messages.append(clarification_message)
            return ChatResponse(messages=pre_messages)

        if plan.steps:
            pre_messages.append(
                self._message(
                    MessageKind.STEP,
                    self._format_plan(plan.steps),
                )
            )

        if plan.tool_calls:
            (
                context_messages,
                user_messages,
                telemetry_updates,
            ) = await self._execute_tool_calls(plan.tool_calls, request)
            pre_messages.extend(user_messages)
            telemetry_updates.setdefault(
                "mcpServersUsed",
                sorted({call.server.id for call in plan.tool_calls}),
            )

        augmented_request = self._augment_request(request, context_messages)
        result_state = await self._compiled.ainvoke({"request": augmented_request})
        response = result_state["response"]

        if pre_messages:
            response.messages = pre_messages + response.messages

        if telemetry_updates:
            response.telemetry = self._merge_telemetry(
                response.telemetry, telemetry_updates
            )

        return response

    async def stream(self, request: ChatRequest) -> AsyncIterator[ProviderStreamEvent]:
        provider_id = request.provider.lower()
        started_at = time.time()
        telemetry_payload: Dict[str, Any] = {}
        plan = await self._build_plan(request)
        context_messages: List[ChatMessage] = []
        mcp_telemetry: Dict[str, Any] = {}

        placeholder = self._message(
            MessageKind.THOUGHT,
            "Analyzing your prompt…",
        )
        yield {"type": "message", "payload": placeholder.model_dump(by_alias=True)}

        if plan.clarification:
            clarification = self._message(
                MessageKind.MESSAGE,
                plan.clarification,
            )
            yield {"type": "message", "payload": clarification.model_dump(by_alias=True)}
            yield {"type": "done"}
            return

        if plan.steps:
            plan_message = self._message(
                MessageKind.STEP, self._format_plan(plan.steps)
            )
            yield {"type": "message", "payload": plan_message.model_dump(by_alias=True)}

        if plan.tool_calls:
            if self._mcp_service is None:
                logger.warning("Plan requested MCP tools but no service is configured.")
            else:
                mcp_telemetry["mcpServersUsed"] = sorted(
                    {call.server.id for call in plan.tool_calls}
                )
                for call in plan.tool_calls:
                    start_msg = self._message(
                        MessageKind.THOUGHT,
                        (
                            f"Calling MCP tool '{call.tool_name}' "
                            f"via {call.server.name}: {call.rationale}"
                        ),
                    )
                    yield {
                        "type": "message",
                        "payload": start_msg.model_dump(by_alias=True),
                    }

                    try:
                        invocation = await self._invoke_tool(call, request)
                    except HTTPException as exc:  # pragma: no cover - surface error
                        error_msg = self._message(
                            MessageKind.STEP,
                            f"Failed to use {call.tool_name}: {exc.detail}",
                        )
                        yield {
                            "type": "message",
                            "payload": error_msg.model_dump(by_alias=True),
                        }
                        continue

                    summary_text = self._summarize_tool_result(invocation)
                    user_summary = self._message(MessageKind.CONTEXT, summary_text)
                    yield {
                        "type": "message",
                        "payload": user_summary.model_dump(by_alias=True),
                    }
                    context_messages.append(
                        self._message(
                            MessageKind.CONTEXT,
                            summary_text,
                            role=MessageRole.SYSTEM,
                        )
                    )
                    mcp_telemetry.setdefault("mcpTools", []).append(
                        {
                            "server": call.server.name,
                            "tool": call.tool_name,
                        }
                    )

        augmented_request = self._augment_request(request, context_messages)

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
            async for event in provider.stream(augmented_request):
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
        telemetry.update(mcp_telemetry)
        telemetry.setdefault("provider", getattr(provider, "id", provider_id))
        if "model" not in telemetry and hasattr(provider, "model"):
            telemetry["model"] = getattr(provider, "model")
        telemetry["latency_ms"] = latency_ms

        yield {"type": "telemetry", "payload": telemetry}
        yield {"type": "done"}

    async def _execute_tool_calls(
        self,
        tool_calls: Sequence[PlannedToolCall],
        request: ChatRequest,
    ) -> tuple[List[ChatMessage], List[ChatMessage], Dict[str, Any]]:
        context_messages: List[ChatMessage] = []
        user_messages: List[ChatMessage] = []
        telemetry_updates: Dict[str, Any] = {}
        if not self._mcp_service:
            return context_messages, user_messages, telemetry_updates

        telemetry_updates["mcpTools"] = []

        for call in tool_calls[: self.MAX_TOOL_CALLS]:
            start_msg = self._message(
                MessageKind.THOUGHT,
                (
                    f"Calling MCP tool '{call.tool_name}' "
                    f"via {call.server.name}: {call.rationale}"
                ),
            )
            user_messages.append(start_msg)

            try:
                invocation = await self._invoke_tool(call, request)
            except HTTPException as exc:
                failure = self._message(
                    MessageKind.STEP,
                    f"Failed to use {call.tool_name}: {exc.detail}",
                )
                user_messages.append(failure)
                continue

            summary = self._summarize_tool_result(invocation)
            user_summary = self._message(MessageKind.CONTEXT, summary)
            user_messages.append(user_summary)
            context_messages.append(
                self._message(
                    MessageKind.CONTEXT,
                    summary,
                    role=MessageRole.SYSTEM,
                )
            )
            telemetry_updates["mcpTools"].append(
                {"server": call.server.name, "tool": call.tool_name}
            )

        return context_messages, user_messages, telemetry_updates

    def _augment_request(
        self, request: ChatRequest, context_messages: Sequence[ChatMessage]
    ) -> ChatRequest:
        if not context_messages:
            return request
        augmented_messages = [*request.messages, *context_messages]
        return request.model_copy(update={"messages": augmented_messages})

    async def _build_plan(self, request: ChatRequest) -> PlanResult:
        if not self._mcp_service:
            return PlanResult(
                steps=[
                    "Review the latest user prompt and workbook context.",
                    "Use the selected LLM provider directly.",
                ]
            )
        servers = self._mcp_service.list_enabled_records()
        if not servers:
            return PlanResult(
                steps=[
                    "Review the latest user prompt and workbook context.",
                    "Use the selected LLM provider directly.",
                ]
            )
        clarification = self._needs_clarification(request)
        if clarification:
            return PlanResult(clarification=clarification)

        prompt = request.prompt.strip()
        steps = ["Review the latest user prompt and workbook context."]
        tool_calls = await self._select_tools(prompt, servers)

        if tool_calls:
            for call in tool_calls:
                steps.append(
                    f"Use {call.tool_name} via {call.server.name} to {call.rationale}."
                )
            steps.append("Combine results with the selected LLM provider.")
        else:
            steps.append("Use the selected LLM provider directly.")

        return PlanResult(steps=steps, tool_calls=tool_calls)

    async def _select_tools(
        self, prompt: str, servers: Sequence[MCPServerRecord]
    ) -> List[PlannedToolCall]:
        if not prompt or not servers:
            return []
        router_calls: List[PlannedToolCall] = []
        if self._router:
            try:
                selections = await self._router.route(prompt, servers)
                router_calls = self._convert_router_selections(selections, servers)
            except Exception as exc:  # pragma: no cover
                logger.warning("Router failed, falling back to heuristics: %s", exc)
        if router_calls:
            return router_calls[: self.MAX_TOOL_CALLS]
        return self._rank_tool_candidates(prompt, servers)

    def _convert_router_selections(
        self,
        selections: Sequence[RouterSelection],
        servers: Sequence[MCPServerRecord],
    ) -> List[PlannedToolCall]:
        if not selections:
            return []
        server_map = {server.id: server for server in servers}
        planned: List[PlannedToolCall] = []
        for selection in selections:
            server = server_map.get(selection.server_id)
            if not server:
                continue
            tool = next(
                (tool for tool in server.tools if tool.name == selection.tool_name),
                None,
            )
            if not tool:
                continue
            planned.append(
                PlannedToolCall(
                    server=server,
                    tool_name=tool.name,
                    rationale=selection.rationale or tool.description or "gather data",
                )
            )
        return planned

    def _rank_tool_candidates(
        self, prompt: str, servers: Sequence[MCPServerRecord]
    ) -> List[PlannedToolCall]:
        if not prompt:
            return []
        candidates: List[tuple[int, PlannedToolCall]] = []
        normalized_prompt = prompt.lower()
        for server in servers:
            if not server.tools:
                continue
            for tool in server.tools:
                score = self._score_tool(normalized_prompt, tool)
                if score <= 0:
                    continue
                rationale = (
                    tool.description.strip()
                    if tool.description
                    else "gather supporting context"
                )
                candidates.append(
                    (
                        score,
                        PlannedToolCall(
                            server=server,
                            tool_name=tool.name,
                            rationale=rationale,
                        ),
                    )
                )
        candidates.sort(key=lambda item: item[0], reverse=True)
        return [item[1] for item in candidates[: self.MAX_TOOL_CALLS]]

    def _format_plan(self, steps: Sequence[str]) -> str:
        return "\n".join(f"{idx + 1}. {step}" for idx, step in enumerate(steps))

    def _needs_clarification(self, request: ChatRequest) -> Optional[str]:
        prompt = request.prompt.strip()
        tokens = [token for token in re.split(r"\W+", prompt) if token]
        if len(tokens) < 4 and not request.selection:
            return (
                "Could you provide a bit more detail so I can pick the right tools?"
            )
        return None

    def _score_tool(self, normalized_prompt: str, tool: MCPTool) -> int:
        score = 0
        tool_name = tool.name.lower()
        if tool_name in normalized_prompt:
            score += 3
        description = (tool.description or "").lower()
        for token in set(normalized_prompt.split()):
            if len(token) < 4:
                continue
            if token in description:
                score += 1
        return score

    def _message(
        self,
        kind: MessageKind,
        content: str,
        role: MessageRole = MessageRole.ASSISTANT,
    ) -> ChatMessage:
        return ChatMessage(
            id=str(uuid4()),
            role=role,
            kind=kind,
            content=content,
            created_at=datetime.now(timezone.utc).isoformat(),
        )

    async def _invoke_tool(
        self, call: PlannedToolCall, request: ChatRequest
    ) -> ToolInvocationResult:
        if not self._mcp_service:
            raise HTTPException(status_code=503, detail="MCP service unavailable.")
        payload = {
            "prompt": request.prompt,
            "selection": [sel.model_dump(by_alias=True) for sel in request.selection],
            "metadata": request.metadata,
            "history": [
                {
                    "role": message.role.value,
                    "kind": message.kind.value,
                    "content": message.content,
                }
                for message in request.messages[-6:]
            ],
            "tool": call.tool_name,
            "requestedAt": datetime.now(timezone.utc).isoformat(),
            "rationale": call.rationale,
        }
        return await self._mcp_service.invoke_tool(
            call.server.id,
            call.tool_name,
            payload,
        )

    def _summarize_tool_result(self, invocation: ToolInvocationResult) -> str:
        try:
            serialized = json.dumps(invocation.response, ensure_ascii=False)
        except (TypeError, ValueError):
            serialized = str(invocation.response)
        if len(serialized) > 600:
            serialized = serialized[:600] + "…"
        return (
            f"Tool '{invocation.tool_name}' from {invocation.server_name} returned:\n"
            f"{serialized}"
        )

    def _merge_telemetry(
        self, telemetry: Optional[Telemetry], updates: Dict[str, Any]
    ) -> Telemetry:
        if telemetry:
            raw = telemetry.raw or {}
            raw.update(updates)
            telemetry.raw = raw
            return telemetry
        return Telemetry(raw=updates)
