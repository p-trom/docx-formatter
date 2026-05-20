"""Application settings."""

from pathlib import Path
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application settings loaded from environment variables."""

    app_name: str = "DOCX Formatter API"
    debug: bool = False
    port: int = 8000
    max_file_size_mb: int = 16
    temp_dir: Path = Path("/tmp/docx_formatter")
    allowed_extensions: set[str] = {".docx"}
    cors_origins: list[str] = ["*"]

    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8")


settings = Settings()
