from __future__ import annotations

import logging
import json

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse

from .config import Settings, get_settings
from .orchestrator import LangGraphOrchestrator
from .providers import available_providers
from .schemas import ChatRequest, ChatResponse, HealthResponse, ProvidersResponse

logger = logging.getLogger(__name__)


def create_app(settings: Settings) -> FastAPI:
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

    orchestrator = LangGraphOrchestrator()

    @app.get("/health", response_model=HealthResponse, tags=["system"])
    async def health() -> HealthResponse:
        return HealthResponse()

    @app.get("/providers", response_model=ProvidersResponse, tags=["system"])
    async def providers() -> ProvidersResponse:
        return ProvidersResponse(providers=available_providers())

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
