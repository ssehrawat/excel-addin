from __future__ import annotations

import asyncio
import json
import logging
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Sequence

from .config import Settings
from .mcp import MCPServerRecord

logger = logging.getLogger(__name__)


@dataclass
class RouterSelection:
    server_id: str
    tool_name: str
    rationale: str


class MCPRouter:
    MAX_CANDIDATES = 18
    MAX_SELECTED = 3

    def __init__(self, settings: Settings) -> None:
        self._provider = settings.mcp_router_provider.lower()
        self._temperature = settings.mcp_router_temperature
        self._model = settings.mcp_router_model
        self._timeout = settings.request_timeout_seconds
        self._client = self._build_client(settings)

    def _build_client(self, settings: Settings):
        if self._provider == "openai":
            api_key = settings.openai_api_key
            if not api_key:
                raise ValueError("OpenAI API key is required for MCP router.")
            try:
                from langchain_openai import ChatOpenAI  # type: ignore
            except ImportError as exc:  # pragma: no cover
                raise RuntimeError("langchain-openai package not available.") from exc
            return ChatOpenAI(
                model=self._model,
                api_key=api_key,
                temperature=self._temperature,
                timeout=self._timeout,
            )
        if self._provider == "anthropic":
            api_key = settings.anthropic_api_key
            if not api_key:
                raise ValueError("Anthropic API key is required for MCP router.")
            try:
                from langchain_anthropic import ChatAnthropic  # type: ignore
            except ImportError as exc:  # pragma: no cover
                raise RuntimeError(
                    "langchain-anthropic package not available."
                ) from exc
            return ChatAnthropic(
                model=self._model,
                anthropic_api_key=api_key,
                temperature=self._temperature,
                timeout=self._timeout,
            )
        raise ValueError(f"Unsupported MCP router provider '{self._provider}'.")

    async def route(
        self, prompt: str, servers: Sequence[MCPServerRecord]
    ) -> List[RouterSelection]:
        prompt = prompt.strip()
        if not prompt or not servers:
            return []
        tools_context = self._build_tools_context(servers)
        if not tools_context:
            return []

        payload = {
            "user_prompt": prompt,
            "tools": tools_context,
            "instructions": (
                "Choose up to three tools that materially help answer the prompt. "
                "Only reference the provided tool names. "
                "If no tool is needed, return an empty list."
            ),
        }
        system_prompt = (
            "You are an MCP tool routing assistant. "
            "Return STRICT JSON with shape "
            '{"selections":[{"serverId":"...","toolName":"...",'
            '"useTool":true,"reason":"..."}]} and nothing else.'
            "Only mark useTool true for tools you actually need."
        )
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": json.dumps(payload)},
        ]
        try:
            result = await asyncio.to_thread(self._client.invoke, messages)
        except Exception as exc:
            logger.warning("MCP router call failed: %s", exc)
            return []

        content = getattr(result, "content", result)
        selections = self._parse_response(content)
        valid: List[RouterSelection] = []
        for item in selections:
            server_id = item.get("serverId")
            tool_name = item.get("toolName")
            use_tool = item.get("useTool", True)
            reason = item.get("reason") or "Router suggestion"
            if not (server_id and tool_name and use_tool):
                continue
            valid.append(
                RouterSelection(
                    server_id=str(server_id),
                    tool_name=str(tool_name),
                    rationale=str(reason),
                )
            )
            if len(valid) >= self.MAX_SELECTED:
                break
        return valid

    def _build_tools_context(
        self, servers: Sequence[MCPServerRecord]
    ) -> List[Dict[str, Any]]:
        context: List[Dict[str, Any]] = []
        for server in servers:
            for tool in server.tools:
                schema = tool.input_schema or {}
                schema_keys = list(schema.keys())
                context.append(
                    {
                        "serverId": server.id,
                        "serverName": server.name,
                        "toolName": tool.name,
                        "toolDescription": tool.description or "",
                        "schemaKeys": schema_keys[:6],
                    }
                )
                if len(context) >= self.MAX_CANDIDATES:
                    return context
        return context

    def _parse_response(self, content: Any) -> List[Dict[str, Any]]:
        if isinstance(content, list):
            try:
                return [json.loads(item) for item in content]
            except Exception:  # pragma: no cover
                content = content[-1]
        if isinstance(content, dict):
            data = content
        else:
            try:
                data = json.loads(content)
            except Exception:
                logger.warning("Router returned non-JSON payload: %s", content)
                return []
        selections = data.get("selections")
        if isinstance(selections, list):
            return [item for item in selections if isinstance(item, dict)]
        return []

