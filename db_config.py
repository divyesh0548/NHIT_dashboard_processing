import os
from datetime import datetime
from pathlib import Path


def resolve_env_file() -> Path:
    """Find .env next to this file, then one directory up (project root)."""
    here = Path(__file__).resolve().parent
    for candidate in (here / ".env", here.parent / ".env"):
        if candidate.exists():
            return candidate
    return here / ".env"


ENV_FILE = resolve_env_file()
REQUIRED_DB_VARS = ("DB_HOST", "DB_PORT", "DB_NAME", "DB_USER", "DB_PASSWORD")
PROCESS_FROM_DATE_VAR = "PROCESS_FROM_DATE"
PROCESS_TO_DATE_VAR = "PROCESS_TO_DATE"


def load_env_file(env_file: Path = ENV_FILE) -> None:
    """Load simple KEY=VALUE pairs from the project .env file into os.environ."""
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


def parse_env_date(date_value: str) -> datetime:
    """Parse a process date from .env (YYYY-MM-DD preferred; also dd-mm-yyyy / dd/mm/yyyy)."""
    date_str = str(date_value).strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise ValueError(
        f"Could not parse date '{date_value}'. Use YYYY-MM-DD (preferred)."
    )


def get_process_date_range_strings() -> tuple[str, str]:
    """
    Return shared process dates from .env as YYYY-MM-DD strings.

    Used by ICICI_bank_bot and YES_bank_bot.
    """
    load_env_file()
    from_raw = os.environ.get(PROCESS_FROM_DATE_VAR, "").strip()
    to_raw = os.environ.get(PROCESS_TO_DATE_VAR, "").strip()
    if not from_raw or not to_raw:
        raise RuntimeError(
            f"Missing {PROCESS_FROM_DATE_VAR} and/or {PROCESS_TO_DATE_VAR} in {ENV_FILE}. "
            "Add both as YYYY-MM-DD."
        )

    from_dt = parse_env_date(from_raw)
    to_dt = parse_env_date(to_raw)
    if to_dt < from_dt:
        raise ValueError(
            f"{PROCESS_TO_DATE_VAR} must be on or after {PROCESS_FROM_DATE_VAR}."
        )
    return from_dt.strftime("%Y-%m-%d"), to_dt.strftime("%Y-%m-%d")


def get_process_date_range_datetimes() -> tuple[datetime, datetime]:
    """
    Return shared process dates from .env as datetime objects.

    Used by IDFC_downloader.
    """
    from_str, to_str = get_process_date_range_strings()
    return parse_env_date(from_str), parse_env_date(to_str)


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
