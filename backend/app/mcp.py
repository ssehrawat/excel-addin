from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from threading import Lock
from typing import Any, Dict, List, Optional, Sequence
from urllib.parse import urljoin
from uuid import uuid4

import httpx
from fastapi import HTTPException
from pydantic import AnyHttpUrl, BaseModel, Field

from .schemas import MCPServerCreateRequest, MCPServerPublic, MCPServerUpdateRequest, MCPTool

logger = logging.getLogger(__name__)


def _timestamp() -> str:
    return datetime.now(timezone.utc).isoformat()


class MCPServerRecord(BaseModel):
    id: str = Field(default_factory=lambda: uuid4().hex)
    name: str
    base_url: AnyHttpUrl
    description: Optional[str] = None
    enabled: bool = True
    api_key: Optional[str] = None
    status: str = "unknown"
    last_refreshed_at: Optional[str] = None
    tools: List[MCPTool] = Field(default_factory=list)
    created_at: str = Field(default_factory=_timestamp)
    updated_at: str = Field(default_factory=_timestamp)

    class Config:
        arbitrary_types_allowed = True


class MCPServerStore:
    def __init__(self, path: Path) -> None:
        self._path = path
        self._path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = Lock()

    def load(self) -> List[MCPServerRecord]:
        with self._lock:
            if not self._path.exists():
                return []
            content = self._path.read_text(encoding="utf-8")
        try:
            data = json.loads(content) if content else []
        except json.JSONDecodeError:
            logger.warning("Failed to parse MCP server store. Resetting file.")
            data = []
        records: List[MCPServerRecord] = []
        for item in data:
            try:
                records.append(MCPServerRecord(**item))
            except Exception as exc:  # pragma: no cover - defensive
                logger.error("Skipping invalid MCP server record: %s", exc)
        return records

    def save(self, records: Sequence[MCPServerRecord]) -> None:
        serialized = [record.model_dump(mode="json") for record in records]
        payload = json.dumps(serialized, indent=2)
        tmp_path = self._path.with_suffix(".tmp")
        with self._lock:
            tmp_path.write_text(payload, encoding="utf-8")
            tmp_path.replace(self._path)


class MCPHttpClient:
    def __init__(self, timeout_seconds: int) -> None:
        self._timeout = timeout_seconds

    async def fetch_tools(self, server: MCPServerRecord) -> List[MCPTool]:
        url = _join(server.base_url, "/tools")
        response = await self._request("GET", url, server)
        raw = response.json()
        if isinstance(raw, dict):
            tools_payload = raw.get("tools", [])
        else:
            tools_payload = raw
        tools: List[MCPTool] = []
        for item in tools_payload:
            if not isinstance(item, dict):
                continue
            name = item.get("name")
            if not name:
                continue
            tools.append(
                MCPTool(
                    name=str(name),
                    description=item.get("description"),
                    input_schema=item.get("input_schema")
                    or item.get("inputSchema")
                    or {},
                )
            )
        return tools

    async def invoke_tool(
        self, server: MCPServerRecord, tool_name: str, payload: Dict[str, Any]
    ) -> Dict[str, Any]:
        url = _join(server.base_url, f"/tools/{tool_name}/invoke")
        response = await self._request("POST", url, server, json={"input": payload})
        data = response.json()
        if isinstance(data, dict):
            return data
        return {"result": data}

    async def _request(
        self,
        method: str,
        url: str,
        server: MCPServerRecord,
        json: Optional[Dict[str, Any]] = None,
    ) -> httpx.Response:
        headers = {}
        if server.api_key:
            headers["Authorization"] = f"Bearer {server.api_key}"
        try:
            async with httpx.AsyncClient(timeout=self._timeout) as client:
                response = await client.request(
                    method, url, headers=headers, json=json
                )
                response.raise_for_status()
                return response
        except httpx.HTTPStatusError as exc:
            logger.warning(
                "MCP server %s responded with %s on %s %s",
                server.name,
                exc.response.status_code,
                method,
                url,
            )
            raise
        except httpx.RequestError as exc:
            logger.warning(
                "Failed to reach MCP server %s at %s: %s", server.name, url, exc
            )
            raise


def _join(base: AnyHttpUrl, path: str) -> str:
    base_url = str(base)
    if not base_url.endswith("/"):
        base_url = f"{base_url}/"
    return urljoin(base_url, path.lstrip("/"))


@dataclass
class ToolInvocationResult:
    server_id: str
    server_name: str
    tool_name: str
    response: Dict[str, Any]


class MCPServerService:
    def __init__(self, storage_path: Path, request_timeout_seconds: int = 15) -> None:
        self._store = MCPServerStore(storage_path)
        self._http = MCPHttpClient(timeout_seconds=request_timeout_seconds)

    def list_servers(self) -> List[MCPServerPublic]:
        return [self.to_public(record) for record in self._store.load()]

    def list_enabled_records(self) -> List[MCPServerRecord]:
        return [record for record in self._store.load() if record.enabled]

    def create_server(self, payload: MCPServerCreateRequest) -> MCPServerRecord:
        records = self._store.load()
        record = MCPServerRecord(
            name=payload.name.strip(),
            base_url=payload.base_url,
            description=payload.description,
            enabled=payload.enabled,
            api_key=payload.api_key,
        )
        records.append(record)
        self._store.save(records)
        logger.info("Registered MCP server '%s' (%s)", record.name, record.id)
        return record

    def update_server(
        self, server_id: str, payload: MCPServerUpdateRequest
    ) -> MCPServerRecord:
        records = self._store.load()
        for index, record in enumerate(records):
            if record.id != server_id:
                continue
            data = payload.model_dump(exclude_unset=True)
            updated = record.model_copy(
                update={
                    **data,
                    "updated_at": _timestamp(),
                }
            )
            records[index] = updated
            self._store.save(records)
            logger.info("Updated MCP server '%s'", updated.name)
            return updated
        raise HTTPException(status_code=404, detail="MCP server not found.")

    def delete_server(self, server_id: str) -> None:
        records = self._store.load()
        next_records = [record for record in records if record.id != server_id]
        if len(records) == len(next_records):
            raise HTTPException(status_code=404, detail="MCP server not found.")
        self._store.save(next_records)
        logger.info("Deleted MCP server %s", server_id)

    async def refresh_server(self, server_id: str) -> MCPServerRecord:
        record = self._get_record(server_id)
        try:
            tools = await self._http.fetch_tools(record)
            updated = record.model_copy(
                update={
                    "tools": tools,
                    "status": "online",
                    "last_refreshed_at": _timestamp(),
                    "updated_at": _timestamp(),
                }
            )
        except Exception as exc:
            logger.exception("Failed to refresh MCP server '%s'", record.name)
            updated = record.model_copy(
                update={
                    "status": "error",
                    "last_refreshed_at": _timestamp(),
                    "updated_at": _timestamp(),
                }
            )
            self._replace_record(updated)
            raise HTTPException(
                status_code=502,
                detail=f"Failed to refresh server '{record.name}': {exc}",
            )
        self._replace_record(updated)
        return updated

    async def invoke_tool(
        self,
        server_id: str,
        tool_name: str,
        payload: Dict[str, Any],
    ) -> ToolInvocationResult:
        record = self._get_record(server_id)
        try:
            response = await self._http.invoke_tool(record, tool_name, payload)
            return ToolInvocationResult(
                server_id=record.id,
                server_name=record.name,
                tool_name=tool_name,
                response=response,
            )
        except Exception as exc:
            logger.exception(
                "Failed to invoke tool '%s' on server '%s'", tool_name, record.name
            )
            raise HTTPException(
                status_code=502,
                detail=f"Tool '{tool_name}' failed on server '{record.name}': {exc}",
            )

    def _get_record(self, server_id: str) -> MCPServerRecord:
        for record in self._store.load():
            if record.id == server_id:
                return record
        raise HTTPException(status_code=404, detail="MCP server not found.")

    def _replace_record(self, updated: MCPServerRecord) -> None:
        records = self._store.load()
        replaced = False
        for index, record in enumerate(records):
            if record.id == updated.id:
                records[index] = updated
                replaced = True
                break
        if not replaced:
            records.append(updated)
        self._store.save(records)

    def to_public(self, record: MCPServerRecord) -> MCPServerPublic:
        return MCPServerPublic(
            id=record.id,
            name=record.name,
            base_url=str(record.base_url),
            description=record.description,
            enabled=record.enabled,
            status=record.status,
            last_refreshed_at=record.last_refreshed_at,
            tools=record.tools,
            created_at=record.created_at,
            updated_at=record.updated_at,
        )

