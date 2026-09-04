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

DEFAULT_GAP_DISTRIBUTION_TABLE = "gap_distribution_per_lane"
LANE_DB_COLUMNS = {f"L{i:02d}": f"l{i:02d}" for i in range(1, 13)}


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


def get_gap_distribution_table_name() -> str:
    load_env_file()
    return (
        os.environ.get("Gap_Distribution_DB_Table", "").strip()
        or DEFAULT_GAP_DISTRIBUTION_TABLE
    )


def melt_gap_wide_to_long(df: pd.DataFrame) -> pd.DataFrame:
    """Convert wide l01…l12 avg-gap columns into long rows expected by gap charts."""
    if df.empty:
        return df

    value_cols = [col for col in LANE_DB_COLUMNS.values() if col in df.columns]
    if not value_cols:
        return df

    id_vars = [col for col in ("plaza_name", "date", "hour", "gap_lt_2s_count", "gap_lt_2s_pct") if col in df.columns]
    long_df = df.melt(
        id_vars=id_vars,
        value_vars=value_cols,
        var_name="_lane_col",
        value_name="avg_gap_sec",
    )
    col_to_lane = {col: lane for lane, col in LANE_DB_COLUMNS.items()}
    long_df["lane_no"] = long_df["_lane_col"].map(col_to_lane)
    long_df = long_df.drop(columns=["_lane_col"])

    # Keep lanes with a measured average; charts treat missing avg as no sample.
    long_df = long_df[long_df["avg_gap_sec"].notna()].copy()
    long_df["gap_count"] = 1
    long_df["vehicle_count"] = 1
    long_df["median_gap_sec"] = long_df["avg_gap_sec"]
    long_df["p10_gap_sec"] = long_df["avg_gap_sec"]
    long_df["p90_gap_sec"] = long_df["avg_gap_sec"]
    long_df["min_gap_sec"] = long_df["avg_gap_sec"]
    long_df["max_gap_sec"] = long_df["avg_gap_sec"]
    return long_df


def fetch_gap_distribution(
    plaza_name: str | None = None,
    start_date=None,
    end_date=None,
) -> pd.DataFrame:
    table_name = get_gap_distribution_table_name()
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
    df["hour"] = df["hour"].astype(str)
    long_df = melt_gap_wide_to_long(df)
    if long_df.empty:
        return long_df
    long_df["lane_no"] = long_df["lane_no"].astype(str)
    return long_df
