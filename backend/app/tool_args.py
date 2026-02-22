from __future__ import annotations

import asyncio
import json
import logging
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Sequence

from fastapi import HTTPException

from .config import Settings
from .schemas import ChatRequest, MCPTool

logger = logging.getLogger(__name__)


class ToolArgumentError(Exception):
    """Raised when arguments for a tool cannot be produced."""


class ToolArgumentBuilder:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._llm = _ArgumentLLM.from_settings(settings)

    async def build(
        self,
        call,
        request: ChatRequest,
    ) -> Dict[str, Any]:
        schema = call.tool.input_schema or {}
        properties = schema.get("properties") or {}
        required = list(schema.get("required") or [])

        heuristics = self._build_from_context(call.tool, request, properties)

        missing = [field for field in required if not heuristics.get(field)]
        if missing and not self._llm:
            raise ToolArgumentError(
                f"Tool '{call.tool_name}' requires {missing}. "
                "Please specify them in your prompt."
            )

        llm_args: Dict[str, Any] = {}
        if self._llm:
            try:
                prompt = self._render_prompt(call, request, schema, heuristics)
                llm_args = await self._llm.generate(prompt)
            except Exception as exc:
                logger.warning("Failed to build LLM arguments: %s", exc)

        arguments = {**heuristics, **llm_args}
        missing_after = [field for field in required if not arguments.get(field)]
        if missing_after:
            raise ToolArgumentError(
                f"Tool '{call.tool_name}' still needs {missing_after}. "
                "Please include them in your request."
            )
        return arguments

    def _build_from_context(
        self,
        tool: MCPTool,
        request: ChatRequest,
        properties: Dict[str, Any],
    ) -> Dict[str, Any]:
        arguments: Dict[str, Any] = {}
        metadata = request.metadata or {}
        selection = request.selection[0] if request.selection else None

        def set_if_absent(key: str, value: Optional[str]) -> None:
            if value and key not in arguments:
                arguments[key] = value

        # database inference
        if "database" in properties:
            db = metadata.get("database") or metadata.get("db")
            if not db:
                db = self._extract_from_prompt(request.prompt, r"(?:db|database)\s+([\w\-]+)")
            set_if_absent("database", db)

        # collection/table inference
        for prop in ("collection", "table", "dataset"):
            if prop in properties:
                collection = metadata.get(prop)
                if not collection and selection and selection.worksheet:
                    collection = selection.worksheet
                if not collection:
                    collection = self._extract_from_prompt(
                        request.prompt,
                        r"(?:collection|table|dataset)\s+([\w\-]+)",
                    )
                set_if_absent(prop, collection)

        # range inference
        if "range" in properties and selection:
            set_if_absent("range", selection.address)

        return arguments

    def _extract_from_prompt(self, prompt: str, pattern: str) -> Optional[str]:
        match = re.search(pattern, prompt, re.IGNORECASE)
        if match:
            return match.group(1)
        return None

    def _render_prompt(
        self,
        call,
        request: ChatRequest,
        schema: Dict[str, Any],
        heuristics: Dict[str, Any],
    ) -> str:
        selection_summary = self._selection_summary(request.selection)
        history = "\n".join(f"{msg.role}: {msg.content}" for msg in request.messages[-5:])
        schema_json = json.dumps(schema, indent=2)
        heuristics_json = json.dumps(heuristics, indent=2)

        instructions = (
            "You build JSON arguments for MCP tool calls.\n"
            "Return strictly valid JSON with keys that appear in the schema.\n"
            "Use the heuristics as defaults when they make sense, but you may overwrite them if the prompt clearly implies a better value.\n"
            "If you must invent a value, choose the most reasonable guess based on the prompt and selection data.\n"
            "Do not include extra commentary."
        )

        return (
            f"{instructions}\n\n"
            f"Tool name: {call.tool_name}\n"
            f"Tool description: {call.tool.description or 'N/A'}\n"
            f"Tool schema:\n{schema_json}\n\n"
            f"Heuristic defaults:\n{heuristics_json}\n\n"
            f"User prompt:\n{request.prompt}\n\n"
            f"Selection summary:\n{selection_summary}\n\n"
            f"Recent history:\n{history}\n\n"
            "Respond with JSON for the arguments."
        )

    def _selection_summary(self, selection: Sequence[Any]) -> str:
        lines: List[str] = []
        for sel in selection[:3]:
            sample = ""
            if sel.values:
                sample_values = []
                for row in sel.values[:2]:
                    sample_values.append(", ".join(str(value) for value in row[:3]))
                sample = " | ".join(sample_values)
            lines.append(f"{sel.address} ({sel.worksheet or 'Sheet'}): {sample}")
        if not lines:
            return "No selection provided."
        return "\n".join(lines)


class _ArgumentLLM:
    @staticmethod
    def from_settings(settings: Settings) -> Optional["_ArgumentLLM"]:
        if settings.openai_api_key:
            try:
                return _OpenAIArgumentLLM(settings)
            except Exception as exc:  # pragma: no cover
                logger.warning("Failed to init OpenAI argument LLM: %s", exc)
        if settings.anthropic_api_key:
            try:
                return _AnthropicArgumentLLM(settings)
            except Exception as exc:  # pragma: no cover
                logger.warning("Failed to init Anthropic argument LLM: %s", exc)
        return None

    async def generate(self, prompt: str) -> Dict[str, Any]:  # pragma: no cover - interface
        raise NotImplementedError


class _OpenAIArgumentLLM(_ArgumentLLM):
    def __init__(self, settings: Settings) -> None:
        from langchain_openai import ChatOpenAI  # type: ignore

        self._client = ChatOpenAI(
            model=settings.openai_model,
            api_key=settings.openai_api_key,
            temperature=settings.openai_temperature or 0,
            timeout=settings.request_timeout_seconds,
        )

    async def generate(self, prompt: str) -> Dict[str, Any]:
        messages = [
            {"role": "system", "content": "Respond with valid JSON only."},
            {"role": "user", "content": prompt},
        ]
        result = await asyncio.to_thread(self._client.invoke, messages)
        return _safe_json(result.content)


class _AnthropicArgumentLLM(_ArgumentLLM):
    def __init__(self, settings: Settings) -> None:
        from langchain_anthropic import ChatAnthropic  # type: ignore

        self._client = ChatAnthropic(
            model=settings.anthropic_model,
            anthropic_api_key=settings.anthropic_api_key,
            temperature=settings.anthropic_temperature or 0,
            timeout=settings.request_timeout_seconds,
        )

    async def generate(self, prompt: str) -> Dict[str, Any]:
        messages = [
            {"role": "system", "content": "Respond with valid JSON only."},
            {"role": "user", "content": prompt},
        ]
        result = await asyncio.to_thread(self._client.invoke, messages)
        return _safe_json(result.content)


def _safe_json(payload: Any) -> Dict[str, Any]:
    if isinstance(payload, dict):
        return payload
    if isinstance(payload, str):
        payload = payload.strip()
        if payload.startswith("```"):
            payload = payload.strip("`")
            payload = payload.replace("json", "", 1).strip()
        return json.loads(payload or "{}")
    raise ValueError("Argument builder returned unsupported payload.")

