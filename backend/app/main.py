from __future__ import annotations

import logging
import json

import logging
from pathlib import Path

from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse

from .config import Settings, get_settings
from .mcp import MCPServerService
from .orchestrator import LangGraphOrchestrator
from .providers import available_providers
from .schemas import (
    ChatRequest,
    ChatResponse,
    HealthResponse,
    MCPServerCreateRequest,
    MCPServerListResponse,
    MCPServerResponse,
    MCPServerUpdateRequest,
    ProvidersResponse,
)

logger = logging.getLogger(__name__)


def _configure_logging(settings: Settings) -> None:
    level = getattr(logging, settings.log_level.upper(), logging.INFO)
    logging.basicConfig(level=level, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    for name in ("excel-addin", __name__):
        logging.getLogger(name).setLevel(level)
    logging.getLogger("uvicorn.access").setLevel(logging.WARNING)


def create_app(settings: Settings) -> FastAPI:
    _configure_logging(settings)
    app = FastAPI(
        title=settings.app_name,
        version="0.1.0",
    )

    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=False,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    storage_path = Path(settings.mcp_config_path)
    if not storage_path.is_absolute():
        storage_path = Path(__file__).resolve().parent / storage_path
    mcp_service = MCPServerService(
        storage_path=storage_path,
        request_timeout_seconds=settings.mcp_request_timeout_seconds,
    )

    orchestrator = LangGraphOrchestrator(mcp_service=mcp_service)

    @app.get("/health", response_model=HealthResponse, tags=["system"])
    async def health() -> HealthResponse:
        return HealthResponse()

    @app.get("/providers", response_model=ProvidersResponse, tags=["system"])
    async def providers() -> ProvidersResponse:
        return ProvidersResponse(providers=available_providers())

    @app.get(
        "/mcp/servers", response_model=MCPServerListResponse, tags=["mcp"]
    )
    async def list_mcp_servers() -> MCPServerListResponse:
        return MCPServerListResponse(servers=mcp_service.list_servers())

    @app.post(
        "/mcp/servers",
        response_model=MCPServerResponse,
        status_code=201,
        tags=["mcp"],
    )
    async def create_mcp_server(
        payload: MCPServerCreateRequest,
    ) -> MCPServerResponse:
        record = mcp_service.create_server(payload)
        if payload.auto_refresh:
            try:
                record = await mcp_service.refresh_server(record.id)
            except HTTPException as exc:
                logger.warning(
                    "Auto-refresh failed for MCP server '%s': %s",
                    record.name,
                    exc.detail,
                )
        return MCPServerResponse(server=mcp_service.to_public(record))

    @app.patch(
        "/mcp/servers/{server_id}",
        response_model=MCPServerResponse,
        tags=["mcp"],
    )
    async def update_mcp_server(
        server_id: str, payload: MCPServerUpdateRequest
    ) -> MCPServerResponse:
        record = mcp_service.update_server(server_id, payload)
        if payload.enabled is not None and payload.enabled and not record.tools:
            try:
                record = await mcp_service.refresh_server(server_id)
            except HTTPException as exc:
                logger.warning(
                    "Failed to refresh MCP server '%s' after enabling: %s",
                    record.name,
                    exc.detail,
                )
        return MCPServerResponse(server=mcp_service.to_public(record))

    @app.delete(
        "/mcp/servers/{server_id}",
        tags=["mcp"],
        status_code=204,
        response_class=Response,
    )
    async def delete_mcp_server(server_id: str) -> Response:
        mcp_service.delete_server(server_id)
        return Response(status_code=204)

    @app.post(
        "/mcp/servers/{server_id}/refresh",
        response_model=MCPServerResponse,
        tags=["mcp"],
    )
    async def refresh_mcp_server(server_id: str) -> MCPServerResponse:
        record = await mcp_service.refresh_server(server_id)
        return MCPServerResponse(server=mcp_service.to_public(record))

    @app.post("/chat", tags=["chat"])
    async def chat_endpoint(chat_request: ChatRequest, http_request: Request):
        logger.info(
            "Handling chat request with provider %s", chat_request.provider
        )

        accept_header = http_request.headers.get("accept", "").lower()
        wants_stream = "application/x-ndjson" in accept_header

        if wants_stream:
            async def stream_response():
                async for event in orchestrator.stream(chat_request):
                    yield json.dumps(event) + "\n"

            return StreamingResponse(
                stream_response(),
                media_type="application/x-ndjson",
                headers={"Cache-Control": "no-cache"},
            )

        response = await orchestrator.run(chat_request)
        return JSONResponse(content=response.model_dump(by_alias=True))

    return app


app = create_app(get_settings())
