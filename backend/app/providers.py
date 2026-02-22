from __future__ import annotations

import asyncio
import json
import logging
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from uuid import uuid4
from typing import (
    Any,
    AsyncIterator,
    Dict,
    Iterable,
    List,
    Optional,
    Protocol,
    TypedDict,
)

from fastapi import HTTPException
from pydantic import ValidationError

from .config import get_settings
from .schemas import (
    ChartInsert,
    ChartSeriesBy,
    ChatMessage,
    ChatRequest,
    ChatResponse,
    CellUpdate,
    FormatUpdate,
    MessageKind,
    MessageRole,
    Telemetry,
)

logger = logging.getLogger(__name__)


STREAM_MESSAGE_DELAY_SECONDS = 0.35
STREAM_MESSAGE_CHUNK_DELAY_SECONDS = 0.08
STREAM_MESSAGE_CHUNK_SIZE = 48

CHART_TYPE_ALIASES: Dict[str, str] = {
    "scatter": "XYScatter",
    "scatterplot": "XYScatter",
    "scatterchart": "XYScatter",
    "scattermarkers": "XYScatter",
    "xyscatter": "XYScatter",
    "xyscattermarkers": "XYScatter",
    "scatterlines": "XYScatterLines",
    "scatterline": "XYScatterLines",
    "xyscatterlines": "XYScatterLines",
    "scatterlinenomarkers": "XYScatterLinesNoMarkers",
    "scatterlinesnomarkers": "XYScatterLinesNoMarkers",
    "xyscatterlinesnomarkers": "XYScatterLinesNoMarkers",
    "bubble": "Bubble",
    "line": "LineMarkers",
    "column": "ColumnClustered",
    "bar": "BarClustered",
}


def _timestamp() -> str:
    return datetime.now(timezone.utc).isoformat()


def _message_id() -> str:
    return uuid4().hex


def _ensure_iterable(value: Any) -> Iterable[Any]:
    if value is None:
        return []
    if isinstance(value, (list, tuple, set)):
        return value
    return [value]


def _is_response_format_error(error: Exception) -> bool:
    text = str(error).lower()
    if "response_format" not in text:
        return False
    return any(
        keyword in text
        for keyword in ["not supported", "unsupported", "only available"]
    )


@dataclass
class ProviderResult:
    messages: List[ChatMessage]
    cell_updates: List[CellUpdate] = field(default_factory=list)
    format_updates: List[FormatUpdate] = field(default_factory=list)
    chart_inserts: List[ChartInsert] = field(default_factory=list)
    telemetry: Optional[Telemetry] = None

    def to_response(self) -> ChatResponse:
        return ChatResponse(
            messages=self.messages,
            cell_updates=self.cell_updates,
            format_updates=self.format_updates,
            chart_inserts=self.chart_inserts,
            telemetry=self.telemetry,
        )


class ProviderStreamEvent(TypedDict, total=False):
    type: str
    payload: Any


def _chunk_text(text: str, size: int = STREAM_MESSAGE_CHUNK_SIZE) -> Iterable[str]:
    if not text:
        return []
    return [text[i : i + size] for i in range(0, len(text), size)]


async def _stream_result(result: ProviderResult) -> AsyncIterator[ProviderStreamEvent]:
    for index, message in enumerate(result.messages):
        if index > 0:
            await asyncio.sleep(STREAM_MESSAGE_DELAY_SECONDS)

        payload = message.model_dump(by_alias=True)
        payload["content"] = ""
        yield {"type": "message_start", "payload": payload}

        for chunk in _chunk_text(message.content):
            await asyncio.sleep(STREAM_MESSAGE_CHUNK_DELAY_SECONDS)
            yield {
                "type": "message_delta",
                "payload": {"id": message.id, "delta": chunk},
            }

        await asyncio.sleep(STREAM_MESSAGE_CHUNK_DELAY_SECONDS)
        yield {"type": "message_done", "payload": message.model_dump(by_alias=True)}

    if result.cell_updates:
        await asyncio.sleep(STREAM_MESSAGE_DELAY_SECONDS)
        yield {
            "type": "cell_updates",
            "payload": [item.model_dump(by_alias=True) for item in result.cell_updates],
        }
    if result.format_updates:
        await asyncio.sleep(STREAM_MESSAGE_DELAY_SECONDS)
        yield {
            "type": "format_updates",
            "payload": [
                item.model_dump(by_alias=True) for item in result.format_updates
            ],
        }
    if result.chart_inserts:
        await asyncio.sleep(STREAM_MESSAGE_DELAY_SECONDS)
        yield {
            "type": "chart_inserts",
            "payload": [
                item.model_dump(by_alias=True) for item in result.chart_inserts
            ],
        }
    if result.telemetry:
        await asyncio.sleep(STREAM_MESSAGE_DELAY_SECONDS)
        yield {
            "type": "telemetry",
            "payload": result.telemetry.model_dump(by_alias=True),
        }


class LLMProvider(Protocol):
    id: str
    label: str
    description: str
    requires_key: bool

    async def generate(self, request: ChatRequest) -> ProviderResult: ...

    async def stream(
        self, request: ChatRequest
    ) -> AsyncIterator[ProviderStreamEvent]: ...


class MockProvider:
    id = "mock"
    label = "Mock provider"
    description = (
        "Deterministic, local responses for development without external API calls."
    )
    requires_key = False

    async def generate(self, request: ChatRequest) -> ProviderResult:
        selection_preview = [
            f"{sel.address} = {sel.values}" for sel in request.selection[:3]
        ]
        messages = [
            ChatMessage(
                id=_message_id(),
                role=MessageRole.ASSISTANT,
                kind=MessageKind.THOUGHT,
                content="Reviewing your prompt and any selected cells.",
                created_at=_timestamp(),
            ),
            ChatMessage(
                id=_message_id(),
                role=MessageRole.ASSISTANT,
                kind=MessageKind.STEP,
                content="Using built-in logic to craft a helpful answer.",
                created_at=_timestamp(),
            ),
        ]

        if selection_preview:
            messages.append(
                ChatMessage(
                    id=_message_id(),
                    role=MessageRole.ASSISTANT,
                    kind=MessageKind.CONTEXT,
                    content="Selected data:\n" + "\n".join(selection_preview),
                    created_at=_timestamp(),
                )
            )

        messages.append(
            ChatMessage(
                id=_message_id(),
                role=MessageRole.ASSISTANT,
                kind=MessageKind.FINAL,
                content=(
                    f"(mock) Answer to '{request.prompt}'. "
                    "Replace me by configuring a real provider in the backend environment."
                ),
                created_at=_timestamp(),
            )
        )
        messages.append(
            ChatMessage(
                id=_message_id(),
                role=MessageRole.ASSISTANT,
                kind=MessageKind.SUGGESTION,
                content="Would you like me to summarize a table, build a formula, or draft insights?",
                created_at=_timestamp(),
            )
        )

        cell_updates: List[CellUpdate] = []
        format_updates: List[FormatUpdate] = []
        if request.selection:
            first_address = request.selection[0].address
            cell_updates.append(
                CellUpdate(
                    address=first_address,
                    values=[["(mock) Result written back to Excel"]],
                )
            )
            if "color" in request.prompt.lower():
                format_updates.append(
                    FormatUpdate(
                        address=first_address,
                        worksheet=request.selection[0].worksheet,
                        fill_color="#FFF2CC",
                    )
                )

        chart_inserts: List[ChartInsert] = []
        if request.selection and any(
            keyword in request.prompt.lower() for keyword in ("chart", "graph")
        ):
            chart_inserts.append(
                ChartInsert(
                    chart_type="columnClustered",
                    source_address=request.selection[0].address,
                    source_worksheet=request.selection[0].worksheet,
                    title="(mock) Selection overview",
                    series_by=ChartSeriesBy.AUTO,
                )
            )

        telemetry = Telemetry(
            provider=self.id,
            model="mock-local",
            latency_ms=5,
        )

        return ProviderResult(
            messages=messages,
            cell_updates=cell_updates,
            format_updates=format_updates,
            chart_inserts=chart_inserts,
            telemetry=telemetry,
        )

    async def stream(self, request: ChatRequest) -> AsyncIterator[ProviderStreamEvent]:
        result = await self.generate(request)
        async for event in _stream_result(result):
            yield event


class OpenAIProvider:
    id = "openai"
    label = "OpenAI"
    description = "Leverage OpenAI models such as GPT-4o."
    requires_key = True

    def __init__(self) -> None:
        settings = get_settings()
        self.api_key = settings.openai_api_key
        self.model = settings.openai_model
        self.temperature = settings.openai_temperature
        if not self.api_key:
            raise HTTPException(
                status_code=503,
                detail="OpenAI API key is not configured on the server.",
            )
        try:
            from langchain_openai import ChatOpenAI  # type: ignore
        except ImportError as exc:  # pragma: no cover
            raise HTTPException(
                status_code=500,
                detail="langchain-openai is not installed on the server.",
            ) from exc

        self._settings = settings
        self._json_mode_enabled = True
        self._client = self._instantiate_openai_client(json_mode=True)

    def _instantiate_openai_client(self, json_mode: bool):  # type: ignore[return-type]
        from langchain_openai import ChatOpenAI  # type: ignore

        client_kwargs = {
            "model": self.model,
            "api_key": self.api_key,
            "timeout": self._settings.request_timeout_seconds,
        }
        if self.temperature is not None:
            client_kwargs["temperature"] = self.temperature
        if json_mode:
            client_kwargs["model_kwargs"] = {"response_format": {"type": "json_object"}}
        client = ChatOpenAI(**client_kwargs)
        self._json_mode_enabled = json_mode
        return client

    async def generate(self, request: ChatRequest) -> ProviderResult:
        system_prompt = (
            "You are MyExcelCompanion, an Excel-focused assistant. "
            "Respond with JSON containing exact keys: "
            "`thoughts` (array of strings), `steps` (array of strings), `answer` (string), "
            "`suggestion` (single string - ONE best optional follow-up action), `cell_updates` (array of objects), "
            "`format_updates` (array of objects for formatting changes), and "
            "`chart_inserts` (array of objects describing charts to create). "
            "CRITICAL: Only include chart_inserts or format_updates when the user EXPLICITLY requests them. "
            "If the user asks to get/fetch/show/retrieve data WITHOUT explicitly asking for charts or formatting, "
            "use ONLY cell_updates to provide the data. Then suggest a chart or formatting in the suggestion field. "
            "Every cell update object MUST include: "
            "`address` (A1 notation, include worksheet like 'Sheet1!E7:E9' when known), "
            "`mode` ('replace' or 'append'), and `values` (a two-dimensional array of rows; "
            "wrap single rows like [['Header'], [123]]). "
            "Only use modes 'replace' or 'append'. Provide cell updates whenever data in Excel should change. "
            "Every format update MUST include `address`, optional `worksheet`, and formatting fields such as "
            "`fill_color`, `font_color`, `bold`, `italic`. Only include format_updates if explicitly requested. "
            "Every chart insert MUST include `chartType` (matching Excel.ChartType), `sourceAddress` "
            "(data range, include worksheet when known). Optional fields: `sourceWorksheet`, `destinationWorksheet`, "
            "`name`, `title`, `topLeftCell`, `bottomRightCell`, `seriesBy` ('auto' | 'rows' | 'columns'). "
            "Use official Excel.ChartType identifiers such as 'XYScatter', 'ColumnClustered', 'LineMarkers'; "
            "do not return aliases like 'xlXYScatter' or 'Scatter'. "
            "The suggestion field should contain ONLY ONE suggestion - the most relevant and useful next action "
            "the user might want to take. Use this for charts/formatting when not explicitly requested. "
            "Phrase it as an optional action that will NOT be executed automatically. "
            "Examples: 'Would you like me to create a line chart from this data?', 'I could also format these cells "
            "with conditional formatting if needed.' Do NOT act on suggestions automatically. "
            "Return strictly valid JSON with fully expanded arrays—never include formulas, code, or list comprehensions."
        )

        messages = [
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": build_prompt_payload(request),
            },
        ]

        try:
            result = await asyncio.to_thread(self._client.invoke, messages)
        except Exception as exc:
            if self._json_mode_enabled and _is_response_format_error(exc):
                logger.warning(
                    "OpenAI model '%s' does not support JSON mode. Falling back to lax parsing.",
                    self.model,
                )
                self._client = self._instantiate_openai_client(json_mode=False)
                result = await asyncio.to_thread(self._client.invoke, messages)
            else:
                raise

        parsed = parse_structured_response(result.content)
        provider_messages = assemble_messages_from_payload(parsed, request.prompt)
        cell_updates = build_cell_updates(parsed.get("cell_updates", []))
        format_updates = build_format_updates(parsed.get("format_updates", []))
        chart_inserts = build_chart_inserts(parsed.get("chart_inserts", []))

        telemetry = Telemetry(
            provider=self.id,
            model=self.model,
            raw={"openai_response": getattr(result, "response_metadata", {})},
        )

        return ProviderResult(
            messages=provider_messages,
            cell_updates=cell_updates,
            format_updates=format_updates,
            chart_inserts=chart_inserts,
            telemetry=telemetry,
        )

    async def stream(self, request: ChatRequest) -> AsyncIterator[ProviderStreamEvent]:
        result = await self.generate(request)
        async for event in _stream_result(result):
            yield event


class AnthropicProvider:
    id = "anthropic"
    label = "Anthropic"
    description = "Access Claude models via Anthropic API."
    requires_key = True

    def __init__(self) -> None:
        settings = get_settings()
        self.api_key = settings.anthropic_api_key
        self.model = settings.anthropic_model
        self.temperature = settings.anthropic_temperature
        if not self.api_key:
            raise HTTPException(
                status_code=503,
                detail="Anthropic API key is not configured on the server.",
            )
        try:
            from langchain_anthropic import ChatAnthropic  # type: ignore
        except ImportError as exc:  # pragma: no cover
            raise HTTPException(
                status_code=500,
                detail="langchain-anthropic is not installed on the server.",
            ) from exc

        self._settings = settings
        self._json_mode_enabled = True
        self._client = self._instantiate_anthropic_client(json_mode=True)

    def _instantiate_anthropic_client(self, json_mode: bool):  # type: ignore[return-type]
        from langchain_anthropic import ChatAnthropic  # type: ignore

        client_kwargs = {
            "model": self.model,
            "anthropic_api_key": self.api_key,
            "timeout": self._settings.request_timeout_seconds,
        }
        if self.temperature is not None:
            client_kwargs["temperature"] = self.temperature
        if json_mode:
            client_kwargs.setdefault("model_kwargs", {})["response_format"] = {
                "type": "json_object"
            }
        client = ChatAnthropic(**client_kwargs)
        self._json_mode_enabled = json_mode
        return client

    async def generate(self, request: ChatRequest) -> ProviderResult:
        system_prompt = (
            "You are MyExcelCompanion, an Excel-focused assistant. "
            "Respond with JSON containing keys thoughts, steps, answer, suggestion, cell_updates, format_updates. "
            "CRITICAL: Only include chart_inserts or format_updates when the user EXPLICITLY requests them. "
            "If the user asks to get/fetch/show/retrieve data WITHOUT explicitly asking for charts or formatting, "
            "use ONLY cell_updates to provide the data. Then suggest a chart or formatting in the suggestion field. "
            "Every cell update object MUST include: "
            "`address` (A1 notation, include worksheet such as 'Sheet1!E7:E9'), "
            "`mode` ('replace' or 'append'), and `values` (two-dimensional array; wrap rows like "
            "[['Header'], [123]]). "
            "Only use modes 'replace' or 'append'. Provide cell updates whenever the workbook should change. "
            "Every format update MUST include `address`, optional `worksheet`, and formatting fields "
            "like `fill_color`, `font_color`, `bold`, `italic`. Only include format_updates if explicitly requested. "
            "Include `chart_inserts` as an array ONLY when the user explicitly requests a chart. Each chart insert object must include "
            "`chartType` (Excel.ChartType string) and `sourceAddress` (include worksheet when known). "
            "Optional fields: `sourceWorksheet`, `destinationWorksheet`, `name`, `title`, `topLeftCell`, `bottomRightCell`, "
            "`seriesBy` ('auto' | 'rows' | 'columns'). "
            "Use official Excel.ChartType identifiers such as 'XYScatter', 'ColumnClustered', 'LineMarkers'; "
            "avoid aliases like 'xlXYScatter' or 'Scatter'. "
            "The suggestion field should contain ONLY ONE suggestion - the most relevant and useful next action "
            "the user might want to take. Use this for charts/formatting when not explicitly requested. "
            "Phrase it as an optional action that will NOT be executed automatically. "
            "Examples: 'Would you like me to create a line chart from this data?', 'I could also format these cells "
            "with conditional formatting if needed.' Do NOT act on suggestions automatically. "
            "Return strictly valid JSON with fully enumerated arrays—no formulas, code, or list comprehensions."
        )
        messages = [
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": build_prompt_payload(request),
            },
        ]

        try:
            result = await asyncio.to_thread(self._client.invoke, messages)
        except Exception as exc:
            if self._json_mode_enabled and _is_response_format_error(exc):
                logger.warning(
                    "Anthropic model '%s' does not support JSON mode. Falling back to lax parsing.",
                    self.model,
                )
                self._client = self._instantiate_anthropic_client(json_mode=False)
                result = await asyncio.to_thread(self._client.invoke, messages)
            else:
                raise

        parsed = parse_structured_response(result.content)
        provider_messages = assemble_messages_from_payload(parsed, request.prompt)
        cell_updates = build_cell_updates(parsed.get("cell_updates", []))
        format_updates = build_format_updates(parsed.get("format_updates", []))
        chart_inserts = build_chart_inserts(parsed.get("chart_inserts", []))

        telemetry = Telemetry(
            provider=self.id,
            model=self.model,
            raw={"anthropic_response": getattr(result, "response_metadata", {})},
        )

        return ProviderResult(
            messages=provider_messages,
            cell_updates=cell_updates,
            format_updates=format_updates,
            chart_inserts=chart_inserts,
            telemetry=telemetry,
        )

    async def stream(self, request: ChatRequest) -> AsyncIterator[ProviderStreamEvent]:
        result = await self.generate(request)
        async for event in _stream_result(result):
            yield event


def build_prompt_payload(request: ChatRequest) -> str:
    selection_text = "\n".join(
        f"{sel.address}: {sel.values}" for sel in request.selection
    )
    history_messages: List[str] = []
    tool_context_segments: List[str] = []
    for msg in request.messages:
        entry = f"[{msg.role}] {msg.content}"
        if msg.kind == MessageKind.CONTEXT:
            tool_context_segments.append(entry)
        else:
            history_messages.append(entry)

    trimmed_history = "\n".join(history_messages[-6:])
    context_block = "\n".join(tool_context_segments)
    if trimmed_history and context_block:
        history = f"{trimmed_history}\n\n[tool_context]\n{context_block}"
    elif context_block:
        history = context_block
    else:
        history = trimmed_history
    payload = {
        "conversation_history": history,
        "user_prompt": request.prompt,
        "selection": selection_text,
        "instructions": (
            "Provide JSON with keys: thoughts (array), steps (array), answer (string), "
            "suggestion (single string - ONE best optional follow-up action), cell_updates (array of {address, mode, values}), "
            "format_updates (array of {address, worksheet, fillColor, fontColor, bold, italic}), "
            "chart_inserts (array of {chartType, sourceAddress, sourceWorksheet?, "
            "destinationWorksheet?, name?, title?, topLeftCell?, bottomRightCell?, seriesBy?}). "
            "CRITICAL: Only include chart_inserts or format_updates when the user EXPLICITLY requests them. "
            "If the user asks to get/fetch/show/retrieve data WITHOUT explicitly asking for charts or formatting, "
            "provide ONLY cell_updates with the data, then suggest creating a chart in the suggestion field. "
            "Conversation history often contains context messages generated from MCP tools (e.g., MongoDB query results). "
            "Treat those context messages as authoritative data retrieved on your behalf and use them to build the answer "
            "and populate cell_updates. Do not claim you lack database access when such context is present. "
            "The suggestion should be the most relevant next action (e.g., creating a chart from the data), phrased as an optional request. "
            "It will NOT be executed automatically - the user must explicitly request it. "
            "Use concise actionable language."
        ),
    }
    return json.dumps(payload)


def _normalize_identifier(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", value.strip().lower().lstrip("xl"))


def _normalize_chart_type(value: Any) -> Optional[str]:
    if not isinstance(value, str):
        return None
    normalized = _normalize_identifier(value)
    if not normalized:
        return None
    if normalized in CHART_TYPE_ALIASES:
        return CHART_TYPE_ALIASES[normalized]
    return value


def parse_structured_response(content: Any) -> Dict[str, Any]:
    if isinstance(content, dict):
        return content
    if isinstance(content, str):
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            logger.warning("LLM returned non-JSON output: %s", content)
    return {
        "thoughts": ["Unable to parse model response."],
        "steps": [],
        "answer": content if isinstance(content, str) else "No answer produced.",
        "suggestion": "",
        "cell_updates": [],
        "format_updates": [],
        "chart_inserts": [],
    }


def assemble_messages_from_payload(
    payload: Dict[str, Any], prompt: str
) -> List[ChatMessage]:
    messages: List[ChatMessage] = []

    # Process thoughts, steps, and answer first
    for key, kind in [
        ("thoughts", MessageKind.THOUGHT),
        ("steps", MessageKind.STEP),
        ("answer", MessageKind.FINAL),
    ]:
        for item in _ensure_iterable(payload.get(key)):
            if not item:
                continue
            messages.append(
                ChatMessage(
                    id=_message_id(),
                    role=MessageRole.ASSISTANT,
                    kind=kind,
                    content=str(item),
                    created_at=_timestamp(),
                )
            )

    # Ensure we have a final answer
    if not any(msg.kind == MessageKind.FINAL for msg in messages):
        messages.append(
            ChatMessage(
                id=_message_id(),
                role=MessageRole.ASSISTANT,
                kind=MessageKind.FINAL,
                content=f"Here is my best effort answer to: {prompt}",
                created_at=_timestamp(),
            )
        )

    # Add single suggestion at the end if present
    suggestion = payload.get("suggestion")
    if suggestion and str(suggestion).strip():
        messages.append(
            ChatMessage(
                id=_message_id(),
                role=MessageRole.ASSISTANT,
                kind=MessageKind.SUGGESTION,
                content=str(suggestion),
                created_at=_timestamp(),
            )
        )

    return messages


def build_cell_updates(raw_updates: Any) -> List[CellUpdate]:
    updates: List[CellUpdate] = []
    for candidate in _ensure_iterable(raw_updates):
        if not isinstance(candidate, dict):
            continue
        address = candidate.get("address")
        values = candidate.get("values")
        if not address or not isinstance(values, list):
            continue
        mode = candidate.get("mode", "replace")
        if isinstance(mode, str):
            normalized_mode = mode.lower()
            if normalized_mode not in {"replace", "append"}:
                normalized_mode = "replace"
            mode = normalized_mode
        worksheet = candidate.get("worksheet")

        try:
            normalized_values = values
            if isinstance(values, list) and values and not isinstance(values[0], list):
                normalized_values = [values]

            updates.append(
                CellUpdate(
                    address=address,
                    values=normalized_values,
                    mode=mode,
                    worksheet=worksheet,
                )
            )
        except ValidationError as error:
            logger.warning(
                "Skipping invalid cell update %s due to validation error: %s",
                candidate,
                error,
            )
    return updates


def build_format_updates(raw_updates: Any) -> List[FormatUpdate]:
    updates: List[FormatUpdate] = []
    for candidate in _ensure_iterable(raw_updates):
        if not isinstance(candidate, dict):
            continue
        address = candidate.get("address")
        if not address:
            continue
        worksheet = candidate.get("worksheet")
        fill_color = candidate.get("fill_color", candidate.get("fillColor"))
        font_color = candidate.get("font_color", candidate.get("fontColor"))
        bold = candidate.get("bold")
        italic = candidate.get("italic")
        try:
            updates.append(
                FormatUpdate(
                    address=address,
                    worksheet=worksheet,
                    fill_color=fill_color,
                    font_color=font_color,
                    bold=bold if isinstance(bold, bool) else None,
                    italic=italic if isinstance(italic, bool) else None,
                )
            )
        except ValidationError as error:
            logger.warning(
                "Skipping invalid format update %s due to validation error: %s",
                candidate,
                error,
            )
    return updates


def build_chart_inserts(raw_inserts: Any) -> List[ChartInsert]:
    inserts: List[ChartInsert] = []
    for candidate in _ensure_iterable(raw_inserts):
        if not isinstance(candidate, dict):
            continue
        chart_type_raw = candidate.get("chart_type") or candidate.get("chartType")
        source_address = candidate.get("source_address") or candidate.get(
            "sourceAddress"
        )
        chart_type = _normalize_chart_type(chart_type_raw)
        if not chart_type or not source_address:
            if not chart_type:
                logger.warning(
                    "Skipping chart insert due to unsupported chart type: %s",
                    chart_type_raw,
                )
            continue
        source_worksheet = candidate.get("source_worksheet") or candidate.get(
            "sourceWorksheet"
        )
        destination_worksheet = candidate.get("destination_worksheet") or candidate.get(
            "destinationWorksheet"
        )
        top_left_cell = candidate.get("top_left_cell") or candidate.get("topLeftCell")
        bottom_right_cell = candidate.get("bottom_right_cell") or candidate.get(
            "bottomRightCell"
        )
        name = candidate.get("name")
        title = candidate.get("title")
        series_by_raw = candidate.get("series_by") or candidate.get("seriesBy")
        series_by = ChartSeriesBy.AUTO
        if isinstance(series_by_raw, str):
            normalized_series_by = series_by_raw.lower()
            if normalized_series_by in {item.value for item in ChartSeriesBy}:
                series_by = ChartSeriesBy(normalized_series_by)
        try:
            inserts.append(
                ChartInsert(
                    chart_type=chart_type,
                    source_address=source_address,
                    source_worksheet=source_worksheet,
                    destination_worksheet=destination_worksheet,
                    top_left_cell=top_left_cell,
                    bottom_right_cell=bottom_right_cell,
                    name=name,
                    title=title,
                    series_by=series_by,
                )
            )
        except ValidationError as error:
            logger.warning(
                "Skipping invalid chart insert %s due to validation error: %s",
                candidate,
                error,
            )
    return inserts


def available_providers() -> List[Dict[str, Any]]:
    settings = get_settings()
    providers: List[Dict[str, Any]] = []
    if settings.mock_provider_enabled:
        providers.append(
            {
                "id": MockProvider.id,
                "label": MockProvider.label,
                "description": MockProvider.description,
                "requiresKey": MockProvider.requires_key,
            }
        )
    providers.append(
        {
            "id": OpenAIProvider.id,
            "label": OpenAIProvider.label,
            "description": OpenAIProvider.description,
            "requiresKey": OpenAIProvider.requires_key,
        }
    )
    providers.append(
        {
            "id": AnthropicProvider.id,
            "label": AnthropicProvider.label,
            "description": AnthropicProvider.description,
            "requiresKey": AnthropicProvider.requires_key,
        }
    )
    return providers


def create_provider(provider_id: str) -> LLMProvider:
    provider_id = provider_id.lower()
    if provider_id == MockProvider.id:
        if not get_settings().mock_provider_enabled:
            raise HTTPException(
                status_code=400, detail="Mock provider is disabled on the server."
            )
        return MockProvider()
    if provider_id == OpenAIProvider.id:
        return OpenAIProvider()
    if provider_id == AnthropicProvider.id:
        return AnthropicProvider()
    raise HTTPException(status_code=400, detail=f"Unknown provider '{provider_id}'.")
