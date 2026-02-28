from functools import lru_cache
from typing import Optional

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_prefix="COPILOT_",
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    app_name: str = "MyExcelCompanion Backend"
    api_prefix: str = "/api"

    mock_provider_enabled: bool = True

    openai_api_key: Optional[str] = None
    openai_model: str = "gpt-4o-mini"
    openai_temperature: Optional[float] = 0.2

    anthropic_api_key: Optional[str] = None
    anthropic_model: str = "claude-3-5-sonnet-20240620"
    anthropic_temperature: Optional[float] = 0.2

    request_timeout_seconds: int = 120
    log_level: str = "INFO"
    mcp_config_path: str = "data/mcp_servers.json"
    mcp_request_timeout_seconds: int = 15


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    return Settings()
