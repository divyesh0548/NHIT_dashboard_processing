import os
from pathlib import Path


ENV_FILE = Path(__file__).resolve().with_name(".env")
REQUIRED_DB_VARS = ("DB_HOST", "DB_PORT", "DB_NAME", "DB_USER", "DB_PASSWORD")


def load_env_file(env_file: Path = ENV_FILE) -> None:
    """Load simple KEY=VALUE pairs from the local .env file into os.environ."""
    if not env_file.exists():
        return

    for raw_line in env_file.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key:
            os.environ.setdefault(key, value)


def get_db_connection_kwargs() -> dict:
    """Return PostgreSQL connection kwargs from environment variables."""
    load_env_file()

    missing = [name for name in REQUIRED_DB_VARS if not os.environ.get(name)]
    if missing:
        missing_vars = ", ".join(missing)
        raise RuntimeError(
            f"Missing required database environment variables in {ENV_FILE}: {missing_vars}"
        )

    return {
        "host": os.environ["DB_HOST"],
        "port": int(os.environ["DB_PORT"]),
        "database": os.environ["DB_NAME"],
        "user": os.environ["DB_USER"],
        "password": os.environ["DB_PASSWORD"],
    }
