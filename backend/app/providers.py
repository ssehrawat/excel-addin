from __future__ import annotations

import asyncio
import json
import logging
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


@dataclass
class ProviderResult:
    messages: List[ChatMessage]
    cell_updates: List[CellUpdate] = field(default_factory=list)
    format_updates: List[FormatUpdate] = field(default_factory=list)
    telemetry: Optional[Telemetry] = None

    def to_response(self) -> ChatResponse:
        return ChatResponse(
            messages=self.messages,
            cell_updates=self.cell_updates,
            format_updates=self.format_updates,
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
                content="Ask me to summarize a table, build a formula, or draft insights.",
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

        telemetry = Telemetry(
            provider=self.id,
            model="mock-local",
            latency_ms=5,
        )

        return ProviderResult(
            messages=messages,
            cell_updates=cell_updates,
            format_updates=format_updates,
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
        client_kwargs = {
            "model": self.model,
            "api_key": self.api_key,
            "timeout": settings.request_timeout_seconds,
        }
        if self.temperature is not None:
            client_kwargs["temperature"] = self.temperature
        self._client = ChatOpenAI(**client_kwargs)

    async def generate(self, request: ChatRequest) -> ProviderResult:
        system_prompt = (
            "You are Workbook Copilot, an Excel-focused assistant. "
            "Respond with JSON containing exact keys: "
            "`thoughts` (array of strings), `steps` (array of strings), `answer` (string), "
            "`suggestions` (array of strings), `cell_updates` (array of objects), and "
            "`format_updates` (array of objects for formatting changes). "
            "Every cell update object MUST include: "
            "`address` (A1 notation, include worksheet like 'Sheet1!E7:E9' when known), "
            "`mode` ('replace' or 'append'), and `values` (a two-dimensional array of rows; "
            "wrap single rows like [['Header'], [123]]). "
            "Only use modes 'replace' or 'append'. Provide cell updates whenever data in Excel should change. "
            "Every format update MUST include `address`, optional `worksheet`, and formatting fields such as "
            "`fill_color`, `font_color`, `bold`, `italic`."
        )

        messages = [
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": build_prompt_payload(request),
            },
        ]

        result = await asyncio.to_thread(self._client.invoke, messages)

        parsed = parse_structured_response(result.content)
        provider_messages = assemble_messages_from_payload(parsed, request.prompt)
        cell_updates = build_cell_updates(parsed.get("cell_updates", []))
        format_updates = build_format_updates(parsed.get("format_updates", []))

        telemetry = Telemetry(
            provider=self.id,
            model=self.model,
            raw={"openai_response": getattr(result, "response_metadata", {})},
        )

        return ProviderResult(
            messages=provider_messages,
            cell_updates=cell_updates,
            format_updates=format_updates,
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
        client_kwargs = {
            "model": self.model,
            "anthropic_api_key": self.api_key,
            "timeout": settings.request_timeout_seconds,
        }
        if self.temperature is not None:
            client_kwargs["temperature"] = self.temperature
        self._client = ChatAnthropic(**client_kwargs)

    async def generate(self, request: ChatRequest) -> ProviderResult:
        system_prompt = (
            "You are Workbook Copilot, an Excel-focused assistant. "
            "Respond with JSON containing keys thoughts, steps, answer, suggestions, cell_updates, format_updates. "
            "Every cell update object MUST include: "
            "`address` (A1 notation, include worksheet such as 'Sheet1!E7:E9'), "
            "`mode` ('replace' or 'append'), and `values` (two-dimensional array; wrap rows like "
            "[['Header'], [123]]). "
            "Only use modes 'replace' or 'append'. Provide cell updates whenever the workbook should change. "
            "Every format update MUST include `address`, optional `worksheet`, and formatting fields "
            "like `fill_color`, `font_color`, `bold`, `italic`."
        )
        messages = [
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": build_prompt_payload(request),
            },
        ]

        result = await asyncio.to_thread(self._client.invoke, messages)

        parsed = parse_structured_response(result.content)
        provider_messages = assemble_messages_from_payload(parsed, request.prompt)
        cell_updates = build_cell_updates(parsed.get("cell_updates", []))
        format_updates = build_format_updates(parsed.get("format_updates", []))

        telemetry = Telemetry(
            provider=self.id,
            model=self.model,
            raw={"anthropic_response": getattr(result, "response_metadata", {})},
        )

        return ProviderResult(
            messages=provider_messages,
            cell_updates=cell_updates,
            format_updates=format_updates,
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
    history = "\n".join(f"[{msg.role}] {msg.content}" for msg in request.messages[-6:])
    payload = {
        "conversation_history": history,
        "user_prompt": request.prompt,
        "selection": selection_text,
        "instructions": (
            "Provide JSON with keys: thoughts (array), steps (array), answer (string), "
            "suggestions (array), cell_updates (array of {address, mode, values}). "
            "Use concise actionable language."
        ),
    }
    return json.dumps(payload)


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
        "suggestions": [],
        "cell_updates": [],
        "format_updates": [],
    }


def assemble_messages_from_payload(
    payload: Dict[str, Any], prompt: str
) -> List[ChatMessage]:
    messages: List[ChatMessage] = []
    for key, kind in [
        ("thoughts", MessageKind.THOUGHT),
        ("steps", MessageKind.STEP),
        ("answer", MessageKind.FINAL),
        ("suggestions", MessageKind.SUGGESTION),
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
