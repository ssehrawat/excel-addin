from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from threading import Lock
from typing import Any, Dict, List, Literal, Optional, Sequence
from urllib.parse import urlparse
from uuid import uuid4

import httpx
from fastapi import HTTPException
from pydantic import AnyHttpUrl, BaseModel, Field

from .schemas import MCPServerCreateRequest, MCPServerPublic, MCPServerUpdateRequest, MCPTool

logger = logging.getLogger(__name__)


def _timestamp() -> str:
    return datetime.now(timezone.utc).isoformat()


class MCPServerRecord(BaseModel):
    """Persisted MCP server configuration.

    The ``protocol`` field is kept for backwards compatibility with existing
    ``mcp_servers.json`` files.  All values ("auto", "rest", "mcp") are now
    treated identically — the JSON-RPC 2.0 transport is always used.
    """

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
    """Thread-safe JSON store for MCP server records.

    Args:
        path: Path to the JSON file used for persistence.
    """

    def __init__(self, path: Path) -> None:
        self._path = path
        self._path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = Lock()

    def load(self) -> List[MCPServerRecord]:
        """Load all server records from disk.

        Returns:
            List of MCPServerRecord instances.  Returns an empty list if the
            file does not exist or contains invalid JSON.
        """
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
        """Atomically persist records to disk.

        Args:
            records: Sequence of MCPServerRecord instances to save.
        """
        serialized = [record.model_dump(mode="json") for record in records]
        payload = json.dumps(serialized, indent=2)
        tmp_path = self._path.with_suffix(".tmp")
        with self._lock:
            tmp_path.write_text(payload, encoding="utf-8")
            tmp_path.replace(self._path)


@dataclass
class JsonRpcSession:
    session_id: str
    protocol_version: str
    endpoint: str


JSONRPC_VERSION = "2.0"
DEFAULT_PROTOCOL_VERSION = "2025-03-26"
CLIENT_NAME = "WorkbookCopilot"
CLIENT_VERSION = "0.2.0"


class MCPJsonRpcClient:
    """MCP JSON-RPC 2.0 transport client.

    Args:
        timeout_seconds: Request timeout in seconds.
    """

    def __init__(self, timeout_seconds: int) -> None:
        self._timeout = timeout_seconds

    async def fetch_tools(self, server: MCPServerRecord) -> List[MCPTool]:
        """Fetch the tool list from an MCP server.

        Args:
            server: The server record to query.

        Returns:
            List of MCPTool instances available on the server.
        """
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
        """Invoke a tool on an MCP server.

        Args:
            server: The server record hosting the tool.
            tool_name: Name of the tool to invoke.
            payload: Arguments to pass to the tool.

        Returns:
            The tool's response as a dict.
        """
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
    ) -> tuple[httpx.Response, str]:
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
            urls.append(f"{base}/mcp")
        # Remove duplicates while preserving order
        deduped: List[str] = []
        for url in urls:
            if url not in deduped:
                deduped.append(url)
        return deduped


@dataclass
class ToolInvocationResult:
    server_id: str
    server_name: str
    tool_name: str
    response: Dict[str, Any]


class MCPServerService:
    """High-level service for managing and invoking MCP servers.

    Uses JSON-RPC 2.0 transport exclusively.

    Args:
        storage_path: Path to the JSON persistence file.
        request_timeout_seconds: Per-request timeout.
    """

    def __init__(self, storage_path: Path, request_timeout_seconds: int = 15) -> None:
        self._store = MCPServerStore(storage_path)
        self._jsonrpc = MCPJsonRpcClient(timeout_seconds=request_timeout_seconds)

    def list_servers(self) -> List[MCPServerPublic]:
        """Return all servers as public-facing models."""
        return [self.to_public(record) for record in self._store.load()]

    def list_enabled_records(self) -> List[MCPServerRecord]:
        """Return only enabled server records."""
        return [record for record in self._store.load() if record.enabled]

    def create_server(self, payload: MCPServerCreateRequest) -> MCPServerRecord:
        """Register a new MCP server.

        Args:
            payload: Creation request from the API.

        Returns:
            Newly created MCPServerRecord.
        """
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
        """Update an existing MCP server record.

        Args:
            server_id: ID of the server to update.
            payload: Fields to update.

        Returns:
            Updated MCPServerRecord.

        Raises:
            HTTPException: 404 if server not found.
        """
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
        """Delete an MCP server record.

        Args:
            server_id: ID of the server to delete.

        Raises:
            HTTPException: 404 if server not found.
        """
        records = self._store.load()
        next_records = [record for record in records if record.id != server_id]
        if len(records) == len(next_records):
            raise HTTPException(status_code=404, detail="MCP server not found.")
        self._store.save(next_records)
        logger.info("Deleted MCP server %s", server_id)

    async def refresh_server(self, server_id: str) -> MCPServerRecord:
        """Refresh a server's tool list via JSON-RPC.

        Args:
            server_id: ID of the server to refresh.

        Returns:
            Updated MCPServerRecord with fresh tool list.

        Raises:
            HTTPException: 502 if the server cannot be reached.
        """
        record = self._get_record(server_id)
        try:
            tools = await self._jsonrpc.fetch_tools(record)
            updated = record.model_copy(
                update={
                    "tools": tools,
                    "status": "online",
                    "last_refreshed_at": _timestamp(),
                    "updated_at": _timestamp(),
                    "protocol": "mcp",
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
        """Invoke a tool on an MCP server via JSON-RPC.

        Args:
            server_id: ID of the server hosting the tool.
            tool_name: Name of the tool to invoke.
            payload: Arguments dict (may include an "arguments" key).

        Returns:
            ToolInvocationResult with the server response.

        Raises:
            HTTPException: 502 if the tool call fails.
        """
        record = self._get_record(server_id)
        try:
            # Unwrap nested "arguments" key if present (legacy callers)
            arguments = payload.get("arguments", payload)
            response = await self._jsonrpc.invoke_tool(record, tool_name, arguments)
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
        """Convert an internal record to the public API model.

        Args:
            record: Internal MCPServerRecord.

        Returns:
            MCPServerPublic suitable for API responses.
        """
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
