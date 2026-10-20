from functools import lru_cache
from pathlib import Path
from typing import Literal

from pydantic_settings import BaseSettings, SettingsConfigDict

APP_DIR = Path(__file__).resolve().parent      # backend/app
BACKEND_DIR = APP_DIR.parent                   # backend


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=BACKEND_DIR / ".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    # App
    app_name: str = "AI Resume Builder"
    debug: bool = False
    cors_origins: list[str] = ["http://localhost:5173"]

    # Database (required: startup fails with a clear error if missing)
    database_url: str

    # LLM gateway
    llm_provider: Literal["gemini", "groq", "ollama"] = "gemini"
    llm_fallback_order: list[str] = ["gemini", "groq", "ollama"]
    llm_max_retries: int = 1

    gemini_api_key: str = ""
    gemini_model: str = ""      # set in .env, copy the name from the provider's docs
    groq_api_key: str = ""
    groq_model: str = ""
    ollama_base_url: str = "http://host.docker.internal:11434"
    ollama_model: str = ""

    # Paths (match your folder structure)
    templates_dir: Path = APP_DIR / "templates"
    prompts_dir: Path = APP_DIR / "prompts"
    export_dir: Path = BACKEND_DIR / "exports"


@lru_cache
def get_settings() -> Settings:
    return Settings()


settings = get_settings()