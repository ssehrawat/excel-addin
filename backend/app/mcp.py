from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from threading import Lock
from typing import Any, Dict, List, Literal, Optional, Sequence
from urllib.parse import urljoin, urlparse
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
    protocol: Literal["auto", "rest", "mcp"] = "auto"
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


@dataclass
class JsonRpcSession:
    session_id: str
    protocol_version: str
    endpoint: str


JSONRPC_VERSION = "2.0"
DEFAULT_PROTOCOL_VERSION = "2025-03-26"
CLIENT_NAME = "MyExcelCompanion"
CLIENT_VERSION = "0.1.0"


class MCPJsonRpcClient:
    def __init__(self, timeout_seconds: int) -> None:
        self._timeout = timeout_seconds

    async def fetch_tools(self, server: MCPServerRecord) -> List[MCPTool]:
        async with httpx.AsyncClient(timeout=self._timeout) as client:
            session = await self._initialize_session(client, server)
            try:
                tools: List[MCPTool] = []
                cursor: Optional[str] = None
                while True:
                    params: Dict[str, Any] = {}
                    if cursor:
                        params["cursor"] = cursor
                    result = await self._rpc_request(
                        client, server, session, "tools/list", params or None
                    )
                    tools.extend(self._parse_tools(result.get("tools", [])))
                    cursor = result.get("nextCursor")
                    if not cursor:
                        break
                return tools
            finally:
                await self._terminate_session(client, server, session)

    async def invoke_tool(
        self,
        server: MCPServerRecord,
        tool_name: str,
        payload: Dict[str, Any],
    ) -> Dict[str, Any]:
        async with httpx.AsyncClient(timeout=self._timeout) as client:
            session = await self._initialize_session(client, server)
            try:
                result = await self._rpc_request(
                    client,
                    server,
                    session,
                    "tools/call",
                    {"name": tool_name, "arguments": payload or {}},
                )
                normalized = dict(result)
                if "structured_content" in normalized and "structuredContent" not in normalized:
                    normalized["structuredContent"] = normalized.pop("structured_content")
                return normalized
            finally:
                await self._terminate_session(client, server, session)

    async def _initialize_session(
        self, client: httpx.AsyncClient, server: MCPServerRecord
    ) -> JsonRpcSession:
        request_id = uuid4().hex
        payload = {
            "jsonrpc": JSONRPC_VERSION,
            "id": request_id,
            "method": "initialize",
            "params": {
                "protocolVersion": DEFAULT_PROTOCOL_VERSION,
                "capabilities": {},
                "clientInfo": {
                    "name": CLIENT_NAME,
                    "version": CLIENT_VERSION,
                },
            },
        }
        response, endpoint = await self._post(client, server, payload)
        data = await self._parse_response(response)
        if data.get("id") != request_id:
            logger.warning("MCP server %s returned mismatched initialize id", server.name)
        result = data.get("result") or {}
        protocol_version = str(
            result.get("protocolVersion") or DEFAULT_PROTOCOL_VERSION
        )
        session_id = (
            response.headers.get("mcp-session-id")
            or result.get("sessionId")
            or result.get("session_id")
        )
        if not session_id:
            raise RuntimeError("MCP server did not provide a session id.")
        session = JsonRpcSession(
            session_id=str(session_id),
            protocol_version=protocol_version,
            endpoint=endpoint,
        )
        await self._post(
            client,
            server,
            {"jsonrpc": JSONRPC_VERSION, "method": "notifications/initialized"},
            session=session,
        )
        return session

    async def _rpc_request(
        self,
        client: httpx.AsyncClient,
        server: MCPServerRecord,
        session: JsonRpcSession,
        method: str,
        params: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        request_id = uuid4().hex
        payload: Dict[str, Any] = {
            "jsonrpc": JSONRPC_VERSION,
            "id": request_id,
            "method": method,
        }
        if params:
            payload["params"] = params
        response, _ = await self._post(client, server, payload, session=session)
        data = await self._parse_response(response)
        if "error" in data:
            raise RuntimeError(data["error"].get("message", "Unknown MCP error"))
        if data.get("id") != request_id:
            logger.debug(
                "MCP server %s returned mismatched id for method %s",
                server.name,
                method,
            )
        return data.get("result") or {}

    async def _terminate_session(
        self,
        client: httpx.AsyncClient,
        server: MCPServerRecord,
        session: JsonRpcSession,
    ) -> None:
        headers = self._build_headers(server, session)
        try:
            response = await client.delete(session.endpoint, headers=headers)
            if response.status_code not in {200, 202, 204, 405}:
                response.raise_for_status()
        except httpx.HTTPStatusError as exc:
            logger.debug(
                "Failed to terminate MCP session for %s: %s",
                server.name,
                exc.response.text,
            )
        except httpx.RequestError as exc:
            logger.debug(
                "Failed to terminate MCP session for %s: %s", server.name, exc
            )

    async def _post(
        self,
        client: httpx.AsyncClient,
        server: MCPServerRecord,
        payload: Dict[str, Any],
        session: Optional[JsonRpcSession] = None,
    ) -> httpx.Response:
        headers = self._build_headers(server, session)
        if session:
            response = await client.post(
                session.endpoint,
                headers=headers,
                json=payload,
            )
            response.raise_for_status()
            return response, session.endpoint

        last_error: Optional[Exception] = None
        candidates = self._candidate_rpc_urls(server)

        for index, url in enumerate(candidates):
            try:
                response = await client.post(
                    url,
                    headers=headers,
                    json=payload,
                )
                response.raise_for_status()
                return response, url
            except httpx.HTTPStatusError as exc:
                last_error = exc
                if exc.response.status_code == 404 and index < len(candidates) - 1:
                    continue
                raise
            except Exception as exc:
                last_error = exc
                if index < len(candidates) - 1:
                    continue
                raise

        if last_error:
            raise last_error
        raise RuntimeError("Unable to contact MCP server.")

    def _build_headers(
        self, server: MCPServerRecord, session: Optional[JsonRpcSession]
    ) -> Dict[str, str]:
        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream",
        }
        if server.api_key:
            headers["Authorization"] = f"Bearer {server.api_key}"
        if session:
            headers["Mcp-Session-Id"] = session.session_id
            headers["Mcp-Protocol-Version"] = session.protocol_version
        return headers

    async def _parse_response(self, response: httpx.Response) -> Dict[str, Any]:
        content_type = response.headers.get("content-type", "")
        if "text/event-stream" in content_type:
            return await self._parse_sse_response(response)
        try:
            data = response.json()
        except ValueError as exc:  # pragma: no cover
            raise RuntimeError("MCP server returned non-JSON response") from exc
        if isinstance(data, dict):
            return data
        raise RuntimeError("MCP server returned unexpected payload.")

    async def _parse_sse_response(self, response: httpx.Response) -> Dict[str, Any]:
        data_lines: List[str] = []
        async for line in response.aiter_lines():
            if line is None:
                break
            stripped = line.strip()
            if not stripped:
                if data_lines:
                    payload = "\n".join(data_lines).strip()
                    data_lines.clear()
                    if not payload:
                        continue
                    try:
                        data = json.loads(payload)
                    except json.JSONDecodeError as exc:
                        raise RuntimeError("Invalid SSE payload from MCP server") from exc
                    if isinstance(data, dict):
                        return data
                    raise RuntimeError("Unexpected SSE payload from MCP server.")
                continue
            if stripped.startswith("data:"):
                data_lines.append(stripped[5:].lstrip())
        raise RuntimeError("MCP server did not send any SSE data.")

    def _parse_tools(self, payload: Sequence[Dict[str, Any]]) -> List[MCPTool]:
        tools: List[MCPTool] = []
        for item in payload:
            name = item.get("name")
            if not name:
                continue
            description = item.get("description") or item.get("annotations", {}).get(
                "description"
            )
            input_schema = (
                item.get("inputSchema")
                or item.get("input_schema")
                or {"type": "object"}
            )
            if not isinstance(input_schema, dict):
                input_schema = {"type": "object"}
            tools.append(
                MCPTool(
                    name=str(name),
                    description=description,
                    input_schema=input_schema,
                )
            )
        return tools

    def _candidate_rpc_urls(self, server: MCPServerRecord) -> List[str]:
        base = str(server.base_url).rstrip("/")
        urls = [base]
        parsed = urlparse(base)
        path = parsed.path or ""
        if not path or path == "/":
            urls.append(_join(server.base_url, "/mcp"))
        # Remove duplicates while preserving order
        deduped: List[str] = []
        for url in urls:
            if url not in deduped:
                deduped.append(url)
        return deduped
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
        self._rest = MCPHttpClient(timeout_seconds=request_timeout_seconds)
        self._jsonrpc = MCPJsonRpcClient(timeout_seconds=request_timeout_seconds)

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
            protocol=payload.protocol,
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
            tools, protocol = await self._fetch_tools(record)
            updated = record.model_copy(
                update={
                    "tools": tools,
                    "status": "online",
                    "last_refreshed_at": _timestamp(),
                    "updated_at": _timestamp(),
                    "protocol": protocol,
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
            response, protocol = await self._invoke_with_strategy(
                record, tool_name, payload
            )
            if protocol != record.protocol:
                record = self._update_protocol(record, protocol)
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

    def _strategies_for(self, protocol: str) -> List[str]:
        if protocol == "mcp":
            return ["mcp"]
        if protocol == "rest":
            return ["rest"]
        return ["rest", "mcp"]

    async def _fetch_tools(
        self, record: MCPServerRecord
    ) -> tuple[List[MCPTool], str]:
        last_error: Optional[Exception] = None
        for strategy in self._strategies_for(record.protocol):
            try:
                if strategy == "rest":
                    return await self._rest.fetch_tools(record), "rest"
                if strategy == "mcp":
                    return await self._jsonrpc.fetch_tools(record), "mcp"
            except Exception as exc:
                last_error = exc
                logger.debug(
                    "Strategy %s failed for server '%s': %s",
                    strategy,
                    record.name,
                    exc,
                )
        if last_error:
            raise last_error
        raise RuntimeError("No MCP transport strategies available.")

    async def _invoke_with_strategy(
        self,
        record: MCPServerRecord,
        tool_name: str,
        payload: Dict[str, Any],
    ) -> tuple[Dict[str, Any], str]:
        last_error: Optional[Exception] = None
        for strategy in self._strategies_for(record.protocol):
            try:
                if strategy == "rest":
                    result = await self._rest.invoke_tool(record, tool_name, payload)
                else:
                    arguments = payload.get("arguments", payload)
                    result = await self._jsonrpc.invoke_tool(
                        record, tool_name, arguments
                    )
                return result, strategy
            except Exception as exc:
                last_error = exc
                logger.debug(
                    "Strategy %s failed for tool '%s' on server '%s': %s",
                    strategy,
                    tool_name,
                    record.name,
                    exc,
                )
        if last_error:
            raise last_error
        raise RuntimeError("No MCP transport strategies available.")

    def _update_protocol(
        self, record: MCPServerRecord, protocol: str
    ) -> MCPServerRecord:
        if protocol == record.protocol:
            return record
        updated = record.model_copy(
            update={"protocol": protocol, "updated_at": _timestamp()}
        )
        self._replace_record(updated)
        return updated

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
            protocol=record.protocol,
            created_at=record.created_at,
            updated_at=record.updated_at,
        )

