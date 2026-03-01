from __future__ import annotations

import asyncio
import json
import logging
import re
import time
from datetime import datetime, timezone
from typing import Any, AsyncIterator, Dict, List, Optional, Sequence, TypedDict
from uuid import uuid4

from fastapi import HTTPException
from langgraph.graph import END, StateGraph

from .mcp import MCPServerRecord, MCPServerService, ToolInvocationResult
from .providers import (
    MCPToolEntry,
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
    WorkbookToolResult,
)

logger = logging.getLogger(__name__)

MAX_REACT_ITERATIONS = 8


class OrchestratorState(TypedDict, total=False):
    request: ChatRequest
    provider_id: str
    mcp_tools: List[MCPToolEntry]
    provider_result: ProviderResult
    response: ChatResponse
    metadata: Dict[str, Any]


class LangGraphOrchestrator:
    """Central orchestrator that manages MCP tools and LLM calls via a ReAct loop.

    The streaming path (``stream()``) is the primary execution path used by
    the frontend and implements the full ReAct (Reason + Act) agent loop.
    The LangGraph graph handles the non-streaming fallback (``run()``).
    """

    def __init__(self, mcp_service: Optional[MCPServerService] = None) -> None:
        """Initialise the orchestrator.

        Args:
            mcp_service: Optional MCP server service for tool invocation and
                health checking. When ``None``, all MCP tool functionality is
                disabled and the ReAct loop operates with Excel tools only.
        """
        self._mcp_service = mcp_service
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
            mcp_tools = state.get("mcp_tools", [])
            provider = create_provider(provider_id, mcp_tools=mcp_tools)
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
        """Non-streaming execution path.

        Performs a per-request MCP health check to inject live tool definitions
        into the system prompt, then invokes the LLM once via the LangGraph graph.

        Args:
            request: The chat request.

        Returns:
            A complete ChatResponse.
        """
        _, mcp_tools = await self._get_live_mcp_tools()
        context_messages = (
            self._build_tool_result_context(request) if request.tool_results else []
        )
        augmented = self._augment_request(request, context_messages)
        result_state = await self._compiled.ainvoke(
            {"request": augmented, "mcp_tools": mcp_tools}
        )
        return result_state["response"]

    async def stream(self, request: ChatRequest) -> AsyncIterator[ProviderStreamEvent]:
        """Stream a ReAct agent response for the given chat request.

        Implements the Reason + Act loop:

        1. Health-check all enabled MCP servers in parallel to determine which
           are reachable. Only live servers contribute tools to the system
           prompt, so the LLM never attempts to call an offline server.
        2. Inject MCP and Excel tool definitions into the system prompt via
           ``build_system_prompt()``. The LLM decides which tools to call
           based on the query — no heuristics or pre-selection.
        3. Loop up to ``MAX_REACT_ITERATIONS`` times:
           - If the LLM returns a direct answer, stream it and finish.
           - If the LLM calls an MCP tool (``mcp__<id>__<name>`` prefix),
             execute it server-side, inject the result as a CONTEXT message,
             and continue the loop.
           - If the LLM calls an Excel tool, emit ``tool_call_required`` so
             the frontend can execute it via Office.js, then end the stream.
             The frontend will re-POST with ``tool_results``.
        4. On re-POST (``request.tool_results`` is non-empty): inject results
           as CONTEXT, re-run the health check, and continue the loop.

        The stream emits the following event types (in order of appearance):
        - ``step``: progress step with ``{id, text, status}`` where status is
          ``"active"`` or ``"done"`` (replaces the old ``status`` event type)
        - ``tool_call_required``: emitted instead of an answer when an Excel
          tool must run in the browser; always followed by ``done``
        - ``message_start`` / ``message_delta`` / ``message_done``: streamed
          final answer text
        - ``cell_updates`` / ``format_updates`` / ``chart_inserts``: Excel
          mutations to apply after the stream completes
        - ``telemetry``: performance metadata (hidden from UI)
        - ``done``: stream complete
        - ``error``: unrecoverable error (provider failure, max iterations)

        Args:
            request: The incoming chat request including prompt, conversation
                history, workbook context, and optional tool results.

        Yields:
            ProviderStreamEvent dicts compatible with the frontend NDJSON
            streaming contract.
        """
        provider_id = request.provider.lower()
        started_at = time.time()
        telemetry_payload: Dict[str, Any] = {}
        mcp_telemetry: Dict[str, Any] = {}
        context_messages: List[ChatMessage] = []
        step_counter = 0
        current_step_text = ""

        def _next_step_id() -> str:
            nonlocal step_counter
            step_counter += 1
            return f"step-{step_counter}"

        def _emit_step(step_id: str, text: str, status: str) -> Dict[str, Any]:
            nonlocal current_step_text
            current_step_text = text
            return {
                "type": "step",
                "payload": {"id": step_id, "text": text, "status": status},
            }

        # Emit initial thinking step
        current_step_id = _next_step_id()
        yield _emit_step(current_step_id, "Thinking\u2026", "active")

        # --- Determine context and MCP tools ---
        if request.tool_results:
            # Re-POST: frontend returned Excel tool results
            context_messages = self._build_tool_result_context(request)

            # Mark thinking done, emit reading step
            yield _emit_step(current_step_id, current_step_text, "done")
            current_step_id = _next_step_id()
            yield _emit_step(current_step_id, "Reading your Excel data\u2026", "active")

            _, mcp_tools = await self._get_live_mcp_tools()
        else:
            # Fresh request: health-check MCP servers before building system prompt
            enabled_count = (
                len(self._mcp_service.list_enabled_records())
                if self._mcp_service
                else 0
            )
            if enabled_count > 0:
                yield _emit_step(current_step_id, current_step_text, "done")
                current_step_id = _next_step_id()
                yield _emit_step(current_step_id, "Checking available tools\u2026", "active")
            _, mcp_tools = await self._get_live_mcp_tools()

        augmented = self._augment_request(request, context_messages)

        try:
            provider = create_provider(provider_id, mcp_tools=mcp_tools)
        except Exception as exc:
            logger.exception("Failed to initialize provider '%s'", provider_id)
            yield {"type": "error", "payload": {"message": str(exc)}}
            return

        # --- ReAct loop ---
        for _iteration in range(MAX_REACT_ITERATIONS):
            # Emit analyzing step before provider call
            yield _emit_step(current_step_id, current_step_text, "done")
            current_step_id = _next_step_id()
            analyze_text = (
                "Reading your Excel data\u2026"
                if request.tool_results and _iteration == 0
                else "Analyzing your data\u2026"
            )
            yield _emit_step(current_step_id, analyze_text, "active")

            try:
                result = await provider.generate(augmented)
            except Exception as exc:
                logger.exception(
                    "Error while calling provider '%s'", provider_id
                )
                yield {"type": "error", "payload": {"message": str(exc)}}
                return

            if result.tool_call_required is None:
                # Mark last step done before streaming answer
                yield _emit_step(current_step_id, current_step_text, "done")
                # LLM produced a final answer — stream it
                try:
                    async for event in provider.stream_result(result):
                        if event.get("type") == "telemetry":
                            payload = event.get("payload") or {}
                            if isinstance(payload, dict):
                                telemetry_payload.update(payload)
                            continue
                        yield event
                except Exception as exc:
                    logger.exception(
                        "Error while streaming result from provider '%s'", provider_id
                    )
                    yield {"type": "error", "payload": {"message": str(exc)}}
                    return
                break  # Final answer streamed — exit ReAct loop

            tc = result.tool_call_required

            if tc.tool.startswith("mcp__"):
                # MCP tool — execute server-side and continue loop
                if not self._mcp_service:
                    failure_msg = (
                        f"MCP tool '{tc.tool}' cannot be called: "
                        "MCP service is not configured."
                    )
                    logger.warning(failure_msg)
                    context_messages.append(
                        self._message(
                            MessageKind.CONTEXT, failure_msg, role=MessageRole.SYSTEM
                        )
                    )
                    augmented = self._augment_request(request, context_messages)
                    continue

                try:
                    server_id, tool_name = self._parse_mcp_tool_name(tc.tool)
                except ValueError as exc:
                    failure_msg = f"Malformed MCP tool name '{tc.tool}': {exc}"
                    logger.warning(failure_msg)
                    context_messages.append(
                        self._message(
                            MessageKind.CONTEXT, failure_msg, role=MessageRole.SYSTEM
                        )
                    )
                    augmented = self._augment_request(request, context_messages)
                    continue

                # Emit MCP tool call step
                yield _emit_step(current_step_id, current_step_text, "done")
                current_step_id = _next_step_id()
                yield _emit_step(current_step_id, f"Calling {tool_name}\u2026", "active")
                try:
                    invocation = await self._mcp_service.invoke_tool(
                        server_id, tool_name, {"arguments": tc.args}
                    )
                    summary = self._summarize_tool_result(invocation)
                    # Track MCP usage for telemetry
                    mcp_telemetry.setdefault("mcpServersUsed", [])
                    if server_id not in mcp_telemetry["mcpServersUsed"]:
                        mcp_telemetry["mcpServersUsed"].append(server_id)
                    mcp_telemetry.setdefault("mcpTools", []).append(
                        {"server": server_id, "tool": tool_name}
                    )
                except HTTPException as exc:
                    summary = (
                        f"MCP tool '{tool_name}' failed: {exc.detail}"
                    )
                    logger.warning(
                        "MCP tool '%s' on server '%s' failed: %s",
                        tool_name,
                        server_id,
                        exc.detail,
                    )

                context_messages.append(
                    self._message(
                        MessageKind.CONTEXT, summary, role=MessageRole.SYSTEM
                    )
                )
                augmented = self._augment_request(request, context_messages)
                # Continue the ReAct loop with the tool result injected

            else:
                # Excel tool — requires browser round-trip
                yield {
                    "type": "tool_call_required",
                    "payload": [tc.model_dump()],
                }
                yield {"type": "done"}
                return

        else:
            # Loop exhausted without a final answer
            yield {
                "type": "error",
                "payload": {"message": "Max reasoning iterations reached."},
            }
            return

        # --- Emit merged telemetry and done ---
        latency_ms = (time.time() - started_at) * 1000
        telemetry = dict(telemetry_payload)
        telemetry.update(mcp_telemetry)
        telemetry.setdefault("provider", getattr(provider, "id", provider_id))
        if "model" not in telemetry and hasattr(provider, "model"):
            telemetry["model"] = getattr(provider, "model")
        telemetry["latency_ms"] = latency_ms

        yield {"type": "telemetry", "payload": telemetry}
        yield {"type": "done"}

    async def _get_live_mcp_tools(
        self,
    ) -> tuple[List[MCPServerRecord], List[MCPToolEntry]]:
        """Health-check all enabled MCP servers in parallel and return live tools.

        Probes each enabled server by attempting a real ``fetch_tools`` RPC call.
        Servers that fail (unreachable, timeout, auth error) are silently excluded
        and logged at WARNING level. The result is ephemeral — it is NOT persisted
        to ``mcp_servers.json`` and does NOT update the UI status badge. The badge
        reflects only explicit user-triggered refreshes.

        Returns:
            Tuple of ``(live_servers, mcp_tool_entries)`` where:
            - ``live_servers`` is the subset of enabled records that responded.
            - ``mcp_tool_entries`` is the flattened list of ``MCPToolEntry``
              objects ready for injection into ``build_system_prompt()``.
            Both lists are empty when ``mcp_service`` is ``None`` or no servers
            are enabled.
        """
        if not self._mcp_service:
            return [], []
        servers = self._mcp_service.list_enabled_records()
        if not servers:
            return [], []

        async def _probe(
            server: MCPServerRecord,
        ) -> Optional[tuple[MCPServerRecord, List[MCPTool]]]:
            try:
                tools = await self._mcp_service.fetch_tools_live(server)
                return server, tools
            except Exception as exc:
                logger.warning(
                    "MCP server '%s' is unreachable; excluding from this request: %s",
                    server.name,
                    exc,
                )
                return None

        results = await asyncio.gather(*(_probe(s) for s in servers))

        live_servers: List[MCPServerRecord] = []
        mcp_tool_entries: List[MCPToolEntry] = []
        for result in results:
            if result is None:
                continue
            server, tools = result
            live_servers.append(server)
            for tool in tools:
                mcp_tool_entries.append(
                    MCPToolEntry(
                        namespaced_name=f"mcp__{server.id}__{tool.name}",
                        server_id=server.id,
                        server_name=server.name,
                        description=tool.description or "",
                        input_schema=tool.input_schema,
                    )
                )

        return live_servers, mcp_tool_entries

    def _parse_mcp_tool_name(self, name: str) -> tuple[str, str]:
        """Parse a namespaced MCP tool name into its server ID and tool name.

        Args:
            name: Namespaced tool name, e.g. ``mcp__3c8a1e9f__find``.

        Returns:
            Tuple of ``(server_id, tool_name)``.

        Raises:
            ValueError: If ``name`` does not follow the ``mcp__<id>__<tool>``
                format (e.g. it is an Excel tool name or malformed string).
        """
        parts = name.split("__", 2)
        if len(parts) != 3 or parts[0] != "mcp" or not parts[1] or not parts[2]:
            raise ValueError(
                f"Expected format 'mcp__<server_id>__<tool_name>', got '{name}'"
            )
        return parts[1], parts[2]

    def _build_tool_result_context(self, request: ChatRequest) -> List[ChatMessage]:
        """Convert frontend Excel tool results into LLM CONTEXT messages.

        Called at the start of a re-POST request (when ``request.tool_results``
        is non-empty). Formats each result via
        ``_summarize_workbook_tool_result()`` and wraps it in a SYSTEM-role
        CONTEXT ``ChatMessage`` so the LLM sees it as prior observations.

        Args:
            request: Incoming chat request that may contain ``tool_results``
                populated by the frontend after executing Office.js tools.

        Returns:
            List of ``ChatMessage`` instances with ``role=SYSTEM`` and
            ``kind=CONTEXT``. Returns an empty list when
            ``request.tool_results`` is empty.
        """
        context: List[ChatMessage] = []
        for tool_result in request.tool_results:
            summary = self._summarize_workbook_tool_result(tool_result)
            context.append(
                self._message(MessageKind.CONTEXT, summary, role=MessageRole.SYSTEM)
            )
        return context

    def _summarize_workbook_tool_result(self, result: WorkbookToolResult) -> str:
        """Format a WorkbookToolResult for injection into the LLM context.

        Args:
            result: A tool result returned by the frontend.

        Returns:
            Markdown-formatted summary string.
        """
        if result.error:
            return f"Excel tool '{result.tool}' returned an error: {result.error}"
        try:
            serialized = json.dumps(result.result, ensure_ascii=False)
        except (TypeError, ValueError):
            serialized = str(result.result)
        if len(serialized) > 50000:
            serialized = serialized[:50000] + "\u2026"
        return f"Excel tool '{result.tool}' returned:\n{serialized}"

    def _augment_request(
        self, request: ChatRequest, context_messages: Sequence[ChatMessage]
    ) -> ChatRequest:
        if not context_messages:
            return request
        augmented_messages = [*request.messages, *context_messages]
        return request.model_copy(update={"messages": augmented_messages})

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

    def _summarize_tool_result(self, invocation: ToolInvocationResult) -> str:
        rows = self._extract_rows(invocation.response)
        if rows:
            table = self._render_table(rows[:10])
            total = len(rows)
            return (
                f"Tool '{invocation.tool_name}' from {invocation.server_name} "
                f"returned {total} document(s).\n\n{table}"
            )
        try:
            serialized = json.dumps(invocation.response, ensure_ascii=False)
        except (TypeError, ValueError):
            serialized = str(invocation.response)
        if len(serialized) > 50000:
            serialized = serialized[:50000] + "\u2026"
        return (
            f"Tool '{invocation.tool_name}' from {invocation.server_name} returned:\n"
            f"{serialized}"
        )

    def _extract_rows(self, response: Dict[str, Any]) -> List[Dict[str, Any]]:
        rows: List[Dict[str, Any]] = []
        structured = response.get("structuredContent")
        if isinstance(structured, list):
            logger.debug("Found structuredContent with %d items", len(structured))
            for item in structured:
                if isinstance(item, dict):
                    rows.append(self._flatten_document(item))
        content = response.get("content")
        if isinstance(content, list):
            logger.debug("Found content array with %d items", len(content))
            for item in content:
                text = item.get("text") if isinstance(item, dict) else None
                if not text:
                    continue
                json_blob = self._extract_json_from_text(text)
                if not json_blob:
                    logger.debug("No JSON found in text item of length %d", len(text))
                    continue
                logger.debug("Extracted JSON blob of length %d", len(json_blob))
                try:
                    data = json.loads(json_blob)
                except json.JSONDecodeError as e:
                    logger.warning("Failed to parse JSON from text: %s", e)
                    continue
                if isinstance(data, list):
                    logger.debug("Parsed JSON list with %d items", len(data))
                    for doc in data:
                        if isinstance(doc, dict):
                            rows.append(self._flatten_document(doc))
                elif isinstance(data, dict):
                    logger.debug("Parsed single JSON object")
                    rows.append(self._flatten_document(data))
        logger.info("Extracted %d rows from tool response", len(rows))
        return rows

    def _extract_json_from_text(self, text: str) -> Optional[str]:
        match = re.search(
            r"<untrusted-user-data[^>]*>(.*?)</untrusted-user-data[^>]*>",
            text,
            re.DOTALL,
        )
        if match:
            candidate = match.group(1).strip()
            logger.debug(
                "Found content within untrusted-user-data tags, length: %d",
                len(candidate),
            )
        else:
            candidate = text.strip()
            logger.debug("No untrusted-user-data tags found, using full text")

        for opening, closing in (("[", "]"), ("{", "}")):
            blob = self._find_balanced_segment(candidate, opening, closing)
            if blob:
                logger.debug("Found balanced %s...%s segment", opening, closing)
                return blob

        result = self._decode_json_fragment(candidate)
        if result:
            logger.debug("JSON decoder found fragment")
        return result

    def _find_balanced_segment(
        self, text: str, opening: str, closing: str
    ) -> Optional[str]:
        start = text.find(opening)
        while start != -1:
            depth = 0
            in_string = False
            escape = False
            for index in range(start, len(text)):
                char = text[index]
                if in_string:
                    if escape:
                        escape = False
                    elif char == "\\":
                        escape = True
                    elif char == '"':
                        in_string = False
                    continue
                if char == '"':
                    in_string = True
                    continue
                if char == opening:
                    depth += 1
                elif char == closing:
                    if depth == 0:
                        break
                    depth -= 1
                    if depth == 0:
                        return text[start : index + 1]
            start = text.find(opening, start + 1)
        return None

    def _decode_json_fragment(self, text: str) -> Optional[str]:
        import json as _json

        decoder = _json.JSONDecoder()
        for opening in ("[", "{"):
            index = text.find(opening)
            while index != -1:
                try:
                    _, end = decoder.raw_decode(text[index:])
                except _json.JSONDecodeError:
                    index = text.find(opening, index + 1)
                    continue
                return text[index : index + end]
        return None

    def _flatten_document(self, doc: Dict[str, Any]) -> Dict[str, Any]:
        flat: Dict[str, Any] = {}
        for key, value in doc.items():
            flat[key] = self._stringify_value(value)
        return flat

    def _stringify_value(self, value: Any) -> str:
        if isinstance(value, dict):
            if "$date" in value:
                return str(value["$date"])
            if "$numberDouble" in value:
                return value["$numberDouble"]
            if "$oid" in value:
                return value["$oid"]
            return json.dumps(value, ensure_ascii=False)
        if isinstance(value, list):
            return ", ".join(self._stringify_value(item) for item in value)
        return str(value)

    def _render_table(self, rows: Sequence[Dict[str, Any]]) -> str:
        headers: List[str] = []
        for row in rows:
            for key in row.keys():
                if key not in headers:
                    headers.append(key)
        if not headers:
            return ""
        header_row = "| " + " | ".join(headers) + " |"
        divider = "| " + " | ".join("---" for _ in headers) + " |"
        lines = [header_row, divider]
        for row in rows:
            values = [str(row.get(header, ""))[:200] for header in headers]
            lines.append("| " + " | ".join(values) + " |")
        logger.debug(
            "Rendered table with %d rows and %d columns", len(rows), len(headers)
        )
        return "\n".join(lines)

    def _merge_telemetry(
        self, telemetry: Optional[Telemetry], updates: Dict[str, Any]
    ) -> Telemetry:
        if telemetry:
            raw = telemetry.raw or {}
            raw.update(updates)
            telemetry.raw = raw
            return telemetry
        return Telemetry(raw=updates)
