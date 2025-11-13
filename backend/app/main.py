from __future__ import annotations

import logging

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

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

    @app.post("/chat", response_model=ChatResponse, tags=["chat"])
    async def chat_endpoint(request: ChatRequest) -> ChatResponse:
        logger.info("Handling chat request with provider %s", request.provider)
        return await orchestrator.run(request)

    return app


app = create_app(get_settings())
