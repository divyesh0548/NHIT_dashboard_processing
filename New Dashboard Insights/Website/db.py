"""Database access for the analytics dashboard."""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pandas as pd
import psycopg2

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT_DIR))

from db_config import get_db_connection_kwargs, load_env_file


def get_analytics_db_connection_kwargs() -> dict:
    load_env_file()
    kwargs = get_db_connection_kwargs()
    nhit_db = os.environ.get("NHIT_DB", "").strip()
    if not nhit_db:
        raise RuntimeError("Missing NHIT_DB in project-root .env (e.g. NHIT_DB=nhit).")
    kwargs["database"] = nhit_db
    return kwargs


def get_analytics_table_name() -> str:
    load_env_file()
    table_name = os.environ.get("Analytical_DB_Table", "").strip()
    if not table_name:
        raise RuntimeError(
            "Missing Analytical_DB_Table in project-root .env "
            "(e.g. Analytical_DB_Table=nhit_analytics)."
        )
    return table_name


def fetch_analytics(
    plaza_name: str | None = None,
    start_date=None,
    end_date=None,
    ) -> pd.DataFrame:
    table_name = get_analytics_table_name()
    query = f"SELECT * FROM {table_name} WHERE 1=1"
    params: list = []

    if plaza_name:
        query += " AND plaza_name = %s"
        params.append(plaza_name)

    if start_date is not None:
        query += " AND date >= %s"
        params.append(start_date)

    if end_date is not None:
        query += " AND date <= %s"
        params.append(end_date)

    query += " ORDER BY date, hour"

    with psycopg2.connect(**get_analytics_db_connection_kwargs()) as conn:
        df = pd.read_sql(query, conn, params=params or None)

    if df.empty:
        return df

    df["date"] = pd.to_datetime(df["date"]).dt.date
    return df


def fetch_plaza_names() -> list[str]:
    table_name = get_analytics_table_name()
    query = f"SELECT DISTINCT plaza_name FROM {table_name} ORDER BY plaza_name"

    with psycopg2.connect(**get_analytics_db_connection_kwargs()) as conn:
        df = pd.read_sql(query, conn)

    return df["plaza_name"].tolist()
