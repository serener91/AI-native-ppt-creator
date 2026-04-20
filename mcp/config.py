from pathlib import Path
from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    output_dir: Path = Path("outputs")
    host: str = "0.0.0.0"
    port: int = 8001          # 8000 is taken by the database MCP server
    on_overflow: str = "warn" # "warn" | "raise" | "ignore"


settings = Settings()
