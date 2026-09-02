"""Aggregation helpers for dashboard time buckets."""

from __future__ import annotations

import calendar
from datetime import date

import pandas as pd

from config import DEFAULT_TOP_LANES, LANES, MOP_TYPES, VEHICLE_CLASSES


def hour_sort_key(hour_label: str) -> int:
    return int(str(hour_label).split("-")[0])


def hour_label(start_hour: int) -> str:
    return f"{start_hour}-{start_hour + 1}"


def hour_labels_for_window(start_hour: int, count: int = 6) -> list[str]:
    return [hour_label(start_hour + offset) for offset in range(count)]


def format_hour_window(hour_labels: list[str]) -> str:
    if not hour_labels:
        return ""
    if len(hour_labels) == 1:
        return hour_labels[0]
    start = str(hour_labels[0]).split("-")[0]
    end = str(hour_labels[-1]).split("-")[1]
    return f"{start}-{end}"


def sort_by_hour(df: pd.DataFrame, hour_col: str = "hour") -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    out["_hour_sort"] = out[hour_col].map(hour_sort_key)
    out = out.sort_values("_hour_sort").drop(columns="_hour_sort")
    return out


def melt_metric_columns(
    df: pd.DataFrame,
    x_col: str,
    column_map: dict[str, str],
    x_order: list | None = None,
    series_order: list | None = None,
) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=[x_col, "series", "count"])

    value_vars = [col for col in column_map.values() if col in df.columns]
    label_lookup = {col: label for label, col in column_map.items()}

    long_df = df.melt(
        id_vars=[x_col],
        value_vars=value_vars,
        var_name="_col",
        value_name="count",
    )
    long_df["series"] = long_df["_col"].map(label_lookup)
    long_df = long_df.drop(columns="_col")

    if x_order is not None:
        categories = [str(value) for value in x_order]
        long_df[x_col] = pd.Categorical(long_df[x_col], categories=categories, ordered=True)
    else:
        long_df[x_col] = long_df[x_col].astype(str)

    if series_order is not None:
        series_categories = [str(value) for value in series_order]
        long_df["series"] = pd.Categorical(long_df["series"], categories=series_categories, ordered=True)

    long_df = long_df.sort_values([x_col, "series"])
    return long_df


def melt_time_buckets_by_category(
    chart_df: pd.DataFrame,
    column_map: dict[str, str],
    bucket_col: str,
    bucket_order: list[str],
    x_col: str = "category",
) -> tuple[pd.DataFrame, str, list[str], list[str]]:
    """Pivot time-bucket rows so categories are on the x-axis and each bucket is a line."""
    categories = [label for label, col in column_map.items() if col in chart_df.columns]
    buckets_sorted = [str(label) for label in bucket_order]
    rows: list[dict] = []

    normalized_df = chart_df.copy()
    if bucket_col in normalized_df.columns:
        normalized_df[bucket_col] = normalized_df[bucket_col].astype(str)

    for bucket_text in buckets_sorted:
        bucket_rows = normalized_df[normalized_df[bucket_col] == bucket_text]
        source = bucket_rows.iloc[0] if not bucket_rows.empty else None
        for label, col in column_map.items():
            if col not in chart_df.columns:
                continue
            if source is not None and col in source.index and pd.notna(source[col]):
                count = int(source[col])
            else:
                count = 0
            rows.append({x_col: label, "series": bucket_text, "count": count})

    long_df = pd.DataFrame(rows, columns=[x_col, "series", "count"])
    if long_df.empty:
        return long_df, x_col, categories, buckets_sorted

    long_df[x_col] = pd.Categorical(long_df[x_col], categories=categories, ordered=True)
    long_df["series"] = pd.Categorical(long_df["series"], categories=buckets_sorted, ordered=True)
    return long_df.sort_values([x_col, "series"]), x_col, categories, buckets_sorted


def melt_hours_by_category(
    chart_df: pd.DataFrame,
    column_map: dict[str, str],
    hour_labels: list[str],
    x_col: str = "category",
) -> tuple[pd.DataFrame, str, list[str], list[str]]:
    return melt_time_buckets_by_category(chart_df, column_map, "hour", hour_labels, x_col)


def top_lanes_by_volume(df: pd.DataFrame, top_n: int = DEFAULT_TOP_LANES) -> list[str]:
    if df.empty:
        return []
    lane_totals = {lane: int(df[col].sum()) for lane, col in LANES.items() if col in df.columns}
    ranked = sorted(lane_totals.items(), key=lambda item: item[1], reverse=True)
    return [lane for lane, _count in ranked[:top_n]]


def aggregate_hourly_window(df: pd.DataFrame, selected_date: date, hour_labels: list[str]) -> pd.DataFrame:
    subset = df[(df["date"] == selected_date) & (df["hour"].isin(hour_labels))].copy()

    metric_columns = [
        column
        for column in df.columns
        if column not in {"id", "plaza_name", "hour", "date"}
    ]

    if not subset.empty:
        grouped = subset.groupby("hour", as_index=False).sum(numeric_only=True)
        grouped["hour"] = grouped["hour"].astype(str)
    else:
        grouped = pd.DataFrame()

    rows: list[dict] = []
    for hour in sorted(hour_labels, key=hour_sort_key):
        hour_text = str(hour)
        if not grouped.empty and hour_text in grouped["hour"].values:
            row = grouped[grouped["hour"] == hour_text].iloc[0].to_dict()
        else:
            row = {"hour": hour_text}
            for column in metric_columns:
                row[column] = 0
        rows.append(row)

    return pd.DataFrame(rows)


def aggregate_daily_range(df: pd.DataFrame, start: date, end: date) -> pd.DataFrame:
    subset = df[(df["date"] >= start) & (df["date"] <= end)].copy()
    if subset.empty:
        return pd.DataFrame()

    grouped = subset.groupby("date", as_index=False).sum(numeric_only=True)
    grouped["date"] = pd.to_datetime(grouped["date"])
    return grouped.sort_values("date")


def aggregate_monthly_daily(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    """One x-axis node per calendar day in the selected month."""
    last_day = calendar.monthrange(year, month)[1]
    metric_columns = list({**VEHICLE_CLASSES, **LANES, **MOP_TYPES}.values())
    if "total_transaction" in df.columns:
        metric_columns = ["total_transaction", *metric_columns]

    rows: list[dict] = []
    for day in range(1, last_day + 1):
        record_date = date(year, month, day)
        subset = df[df["date"] == record_date]
        row: dict = {"day": day, "day_label": f"{day:02d}"}

        for column in metric_columns:
            if column in df.columns:
                row[column] = int(subset[column].sum()) if not subset.empty else 0

        rows.append(row)

    return pd.DataFrame(rows)


def aggregate_monthly_parts(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    return aggregate_monthly_daily(df, year, month)


def total_volume(df: pd.DataFrame) -> int:
    if df.empty:
        return 0
    if "total_transaction" in df.columns:
        return int(df["total_transaction"].sum())
    class_cols = [col for col in VEHICLE_CLASSES.values() if col in df.columns]
    if not class_cols:
        return 0
    return int(df[class_cols].sum().sum())


def format_total_kpi(df: pd.DataFrame) -> str:
    return f"{total_volume(df):,}"
