"""
NHIT Analytics Dashboard — Streamlit + Plotly.

Run from project root:
    streamlit run "New Dashboard Insights/Website/main.py"
"""

from __future__ import annotations

import calendar
from datetime import date

import pandas as pd
import streamlit as st

from charts import class_bar_chart, line_series_chart
from config import (
    LANES,
    MAX_DAY_WINDOW,
    MAX_HOUR_WINDOW,
    MIN_DAY_WINDOW,
    MIN_HOUR_WINDOW,
    MOP_TYPES,
    VEHICLE_CLASSES,
)
from data_utils import (
    aggregate_daily_range,
    aggregate_hourly_window,
    aggregate_monthly_daily,
    format_hour_window,
    format_total_kpi,
    hour_labels_for_window,
    hour_sort_key,
    melt_metric_columns,
    melt_time_buckets_by_category,
    top_lanes_by_volume,
)
from db import fetch_analytics, fetch_plaza_names

st.set_page_config(
    page_title="NHIT Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

TIME_MODES = ("Hours", "Day on day", "Month")

AXIS_CONFIG = {
    "Hours": {"x_axis_title": "Hour", "hover_bucket_label": "Hour"},
    "Day on day": {"x_axis_title": "Date", "hover_bucket_label": "Date"},
    "Month": {"x_axis_title": "Date", "hover_bucket_label": "Date"},
}


VOLUME_X_AXIS_OPTIONS = ("Time", "Vehicle class")


def chart_axis_config(mode: str) -> dict[str, str]:
    return AXIS_CONFIG.get(mode, {"x_axis_title": "Time", "hover_bucket_label": "Time"})


def volume_chart_title(chart_kind: str, time_mode: str, x_axis_dim: str) -> str:
    metric = "Vehicle" if chart_kind == "class" else "Lane"
    if x_axis_dim == "Vehicle class":
        x_label = "vehicle class" if chart_kind == "class" else "lane"
        return f"{metric} count by {x_label}"

    x_label_by_mode = {
        "Hours": "hours",
        "Day on day": "day on day",
        "Month": "month",
    }
    return f"{metric} count by {x_label_by_mode.get(time_mode, 'time')}"


def volume_x_axis_title(chart_kind: str, time_mode: str, x_axis_dim: str) -> str:
    if x_axis_dim == "Vehicle class":
        return "Vehicle class" if chart_kind == "class" else "Lane"
    return chart_axis_config(time_mode)["x_axis_title"]


def mop_chart_title(time_mode: str) -> str:
    x_label_by_mode = {
        "Hours": "hours",
        "Day on day": "day on day",
        "Month": "month",
    }
    return f"MOP count by {x_label_by_mode.get(time_mode, 'time')}"


@st.cache_data(ttl=300)
def load_data(plaza_name: str) -> pd.DataFrame:
    return fetch_analytics(plaza_name=plaza_name)


def available_dates(df: pd.DataFrame) -> list[date]:
    if df.empty:
        return []
    return sorted(df["date"].unique())


def earliest_date(dates: list[date]) -> date:
    return dates[0]


def default_day_on_day_range(dates: list[date]) -> tuple[date, date]:
    """Start from the earliest available date with the minimum required window."""
    start = earliest_date(dates)
    end = dates[-1]
    for candidate in dates:
        if (candidate - start).days + 1 >= MIN_DAY_WINDOW:
            end = candidate
            break
    return start, end


def default_year_month(dates: list[date]) -> tuple[int, int]:
    earliest = earliest_date(dates)
    return earliest.year, earliest.month


def validate_day_range(start_date: date, end_date: date) -> tuple[bool, str | None]:
    if end_date < start_date:
        return False, "End date cannot be before start date."
    day_count = (end_date - start_date).days + 1
    if day_count < MIN_DAY_WINDOW:
        return False, f"Select at least {MIN_DAY_WINDOW} days (currently {day_count})."
    if day_count > MAX_DAY_WINDOW:
        return False, f"Select at most {MAX_DAY_WINDOW} days (currently {day_count})."
    return True, None


def render_time_mode_controls(
    section_key: str,
    flexible_hour_window: bool = False,
) -> tuple[str, dict]:
    """Shared Hours / Day on day / Month controls for line charts."""
    mode = st.radio(
        "Time granularity",
        TIME_MODES,
        horizontal=True,
        key=f"{section_key}_mode",
    )
    dates = available_dates(st.session_state.get("df", pd.DataFrame()))
    params: dict = {"mode": mode}

    if not dates:
        st.warning("No data available for the selected plaza.")
        return mode, params

    if mode == "Hours":
        if flexible_hour_window:
            col1, col2, col3 = st.columns(3)
            with col1:
                params["selected_date"] = st.date_input(
                    "Date",
                    value=earliest_date(dates),
                    min_value=dates[0],
                    max_value=dates[-1],
                    key=f"{section_key}_hour_date",
                )
            with col2:
                params["start_hour"] = st.selectbox(
                    "From",
                    options=list(range(0, 24 - MIN_HOUR_WINDOW + 1)),
                    format_func=str,
                    index=min(8, 24 - MIN_HOUR_WINDOW),
                    key=f"{section_key}_start_hour",
                )
            with col3:
                max_hours = min(MAX_HOUR_WINDOW, 24 - params["start_hour"])
                params["hour_count"] = st.slider(
                    "Hour range",
                    min_value=MIN_HOUR_WINDOW,
                    max_value=max(MIN_HOUR_WINDOW, max_hours),
                    value=min(6, max(MIN_HOUR_WINDOW, max_hours)),
                    key=f"{section_key}_hour_count",
                )
            params["hour_labels"] = hour_labels_for_window(
                params["start_hour"],
                params["hour_count"],
            )
            st.caption(f"Selected hour range: {format_hour_window(params['hour_labels'])}")
        else:
            col1, col2 = st.columns(2)
            with col1:
                params["selected_date"] = st.date_input(
                    "Date",
                    value=earliest_date(dates),
                    min_value=dates[0],
                    max_value=dates[-1],
                    key=f"{section_key}_hour_date",
                )
            with col2:
                params["start_hour"] = st.selectbox(
                    "From",
                    options=list(range(0, 24 - MIN_HOUR_WINDOW + 1)),
                    format_func=str,
                    index=min(8, 24 - MIN_HOUR_WINDOW),
                    key=f"{section_key}_start_hour",
                )
            params["hour_labels"] = hour_labels_for_window(params["start_hour"], MIN_HOUR_WINDOW)
            st.caption(f"Selected hour range: {format_hour_window(params['hour_labels'])}")

    elif mode == "Day on day":
        col1, col2 = st.columns(2)
        default_start, default_end = default_day_on_day_range(dates)
        with col1:
            params["start_date"] = st.date_input(
                "Start date",
                value=default_start,
                min_value=dates[0],
                max_value=dates[-1],
                key=f"{section_key}_daily_start",
            )
        with col2:
            params["end_date"] = st.date_input(
                "End date",
                value=default_end,
                min_value=max(params["start_date"], dates[0]),
                max_value=dates[-1],
                key=f"{section_key}_daily_end",
            )
        is_valid, error_message = validate_day_range(params["start_date"], params["end_date"])
        if not is_valid:
            st.error(error_message)

    else:
        default_year, default_month = default_year_month(dates)
        year_options = sorted({d.year for d in dates}, reverse=True)
        params["year"] = st.selectbox(
            "Year",
            options=year_options,
            index=year_options.index(default_year),
            key=f"{section_key}_year",
        )
        months_in_year = sorted({d.month for d in dates if d.year == params["year"]})
        month_default = default_month if default_year == params["year"] else months_in_year[0]
        params["month"] = st.selectbox(
            "Month",
            options=months_in_year,
            format_func=lambda m: calendar.month_name[m],
            index=months_in_year.index(month_default),
            key=f"{section_key}_month",
        )

    return mode, params


def build_time_series_df(df: pd.DataFrame, params: dict) -> tuple[pd.DataFrame, str, list | None, bool, str]:
    mode = params["mode"]

    if mode == "Hours":
        chart_df = aggregate_hourly_window(df, params["selected_date"], params["hour_labels"])
        x_col = "hour"
        x_order = [str(label) for label in sorted(params["hour_labels"], key=hour_sort_key)]
        return chart_df, x_col, x_order, False, mode

    if mode == "Day on day":
        chart_df = aggregate_daily_range(df, params["start_date"], params["end_date"])
        if chart_df.empty:
            return chart_df, "date", None, False, mode
        chart_df["date"] = chart_df["date"].dt.strftime("%Y-%m-%d")
        return chart_df, "date", chart_df["date"].tolist(), False, mode

    chart_df = aggregate_monthly_daily(df, params["year"], params["month"])
    x_order = chart_df["day_label"].astype(str).tolist() if not chart_df.empty else None
    return chart_df, "day_label", x_order, True, mode


def volume_time_bucket_config(
    time_mode: str,
    params: dict,
    x_col: str,
    x_order: list | None,
) -> tuple[str, list[str]]:
    if time_mode == "Hours":
        bucket_col = "hour"
        bucket_order = [str(label) for label in sorted(params["hour_labels"], key=hour_sort_key)]
        return bucket_col, bucket_order

    if not x_order:
        return x_col, []

    return x_col, [str(value) for value in x_order]


def build_volume_chart_data(
    chart_df: pd.DataFrame,
    column_map: dict[str, str],
    *,
    use_class_on_x_axis: bool,
    time_mode: str,
    params: dict,
    x_col: str,
    x_order: list | None,
) -> tuple[pd.DataFrame, str, list | None, list | None]:
    if use_class_on_x_axis:
        bucket_col, bucket_order = volume_time_bucket_config(time_mode, params, x_col, x_order)
        long_df, pivot_x_col, pivot_x_order, series_order = melt_time_buckets_by_category(
            chart_df,
            column_map,
            bucket_col,
            bucket_order,
        )
        return long_df, pivot_x_col, pivot_x_order, series_order

    long_df = melt_metric_columns(chart_df, x_col, column_map, x_order)
    return long_df, x_col, x_order, None


def render_volume_line_chart(
    long_df: pd.DataFrame,
    x_col: str,
    title: str,
    *,
    x_axis_title: str,
    hover_bucket_label: str,
    x_order: list | None,
    series_order: list | None = None,
    highlight_days: bool = False,
) -> None:
    st.plotly_chart(
        line_series_chart(
            long_df,
            x_col,
            title,
            x_axis_title=x_axis_title,
            hover_bucket_label=hover_bucket_label,
            x_order=x_order,
            series_order=series_order,
            highlight_every_nth_day=5 if highlight_days else None,
        ),
        use_container_width=True,
    )


def render_volume_section(df: pd.DataFrame) -> None:
    st.subheader("1. Traffic volume by class and lane")
    mode, params = render_time_mode_controls("volume", flexible_hour_window=True)

    if df.empty:
        return

    if mode == "Day on day":
        is_valid, _ = validate_day_range(params["start_date"], params["end_date"])
        if not is_valid:
            return

    x_axis_dim = st.radio(
        "X-axis",
        VOLUME_X_AXIS_OPTIONS,
        horizontal=True,
        key="volume_x_axis",
    )

    chart_df, x_col, x_order, highlight_days, time_mode = build_time_series_df(df, params)
    axis = chart_axis_config(time_mode)
    st.metric("Total vehicles (filtered)", format_total_kpi(chart_df))

    use_class_on_x_axis = x_axis_dim == "Vehicle class"

    class_long, class_x_col, class_x_order, class_series_order = build_volume_chart_data(
        chart_df,
        VEHICLE_CLASSES,
        use_class_on_x_axis=use_class_on_x_axis,
        time_mode=time_mode,
        params=params,
        x_col=x_col,
        x_order=x_order,
    )
    class_title = volume_chart_title("class", time_mode, x_axis_dim)
    class_x_title = volume_x_axis_title("class", time_mode, x_axis_dim) if use_class_on_x_axis else axis["x_axis_title"]
    class_hover_label = class_x_title if use_class_on_x_axis else axis["hover_bucket_label"]
    render_volume_line_chart(
        class_long,
        class_x_col,
        class_title,
        x_axis_title=class_x_title,
        hover_bucket_label=class_hover_label,
        x_order=class_x_order,
        series_order=class_series_order,
        highlight_days=highlight_days and not use_class_on_x_axis,
    )

    show_all_lanes = st.toggle("Show all lanes", value=False, key="show_all_lanes")
    lane_map = LANES
    if not show_all_lanes:
        top = top_lanes_by_volume(chart_df)
        lane_map = {lane: col for lane, col in LANES.items() if lane in top}

    lane_scope = "all lanes" if show_all_lanes else "top 5"
    lane_long, lane_x_col, lane_x_order, lane_series_order = build_volume_chart_data(
        chart_df,
        lane_map,
        use_class_on_x_axis=use_class_on_x_axis,
        time_mode=time_mode,
        params=params,
        x_col=x_col,
        x_order=x_order,
    )
    lane_title = f"{volume_chart_title('lane', time_mode, x_axis_dim)} ({lane_scope})"
    lane_x_title = volume_x_axis_title("lane", time_mode, x_axis_dim) if use_class_on_x_axis else axis["x_axis_title"]
    lane_hover_label = lane_x_title if use_class_on_x_axis else axis["hover_bucket_label"]
    render_volume_line_chart(
        lane_long,
        lane_x_col,
        lane_title,
        x_axis_title=lane_x_title,
        hover_bucket_label=lane_hover_label,
        x_order=lane_x_order,
        series_order=lane_series_order,
        highlight_days=highlight_days and not use_class_on_x_axis,
    )


def render_class_mix_section(df: pd.DataFrame) -> None:
    st.subheader("2. Vehicle class mix")
    mode = st.radio(
        "Aggregation",
        TIME_MODES,
        horizontal=True,
        key="class_mix_mode",
    )
    dates = available_dates(df)
    if not dates:
        st.warning("No data available.")
        return

    filtered = pd.DataFrame()

    if mode == "Hours":
        col1, col2 = st.columns(2)
        with col1:
            selected_date = st.date_input(
                "Date",
                value=earliest_date(dates),
                min_value=dates[0],
                max_value=dates[-1],
                key="class_mix_hour_date",
            )
        with col2:
            start_hour = st.selectbox(
                "From",
                options=list(range(0, 24 - MIN_HOUR_WINDOW + 1)),
                format_func=str,
                index=min(8, 24 - MIN_HOUR_WINDOW),
                key="class_mix_start_hour",
            )
        hour_labels = hour_labels_for_window(start_hour, MIN_HOUR_WINDOW)
        st.caption(f"Selected hour range: {format_hour_window(hour_labels)}")
        filtered = df[(df["date"] == selected_date) & (df["hour"].isin(hour_labels))]

    elif mode == "Day on day":
        col1, col2 = st.columns(2)
        default_start, default_end = default_day_on_day_range(dates)
        with col1:
            start_date = st.date_input(
                "Start date",
                value=default_start,
                min_value=dates[0],
                max_value=dates[-1],
                key="class_mix_daily_start",
            )
        with col2:
            end_date = st.date_input(
                "End date",
                value=default_end,
                min_value=max(start_date, dates[0]),
                max_value=dates[-1],
                key="class_mix_daily_end",
            )
        is_valid, error_message = validate_day_range(start_date, end_date)
        if not is_valid:
            st.error(error_message)
            return
        filtered = df[(df["date"] >= start_date) & (df["date"] <= end_date)]

    else:
        default_year, default_month = default_year_month(dates)
        year_options = sorted({d.year for d in dates}, reverse=True)
        year = st.selectbox(
            "Year",
            options=year_options,
            index=year_options.index(default_year),
            key="class_mix_year",
        )
        months_in_year = sorted({d.month for d in dates if d.year == year})
        month_default = default_month if default_year == year else months_in_year[0]
        month = st.selectbox(
            "Month",
            options=months_in_year,
            format_func=lambda m: calendar.month_name[m],
            index=months_in_year.index(month_default),
            key="class_mix_month",
        )
        month_start = date(year, month, 1)
        month_end = date(year, month, calendar.monthrange(year, month)[1])
        filtered = df[(df["date"] >= month_start) & (df["date"] <= month_end)]

    st.metric("Total vehicles (filtered)", format_total_kpi(filtered))
    st.plotly_chart(
        class_bar_chart(filtered, VEHICLE_CLASSES, f"Class distribution — {mode}"),
        use_container_width=True,
    )


def render_mop_section(df: pd.DataFrame) -> None:
    st.subheader("3. FASTag / Cash / Exempt over time")
    mode, params = render_time_mode_controls("mop")

    if df.empty:
        return

    if mode == "Day on day":
        is_valid, _ = validate_day_range(params["start_date"], params["end_date"])
        if not is_valid:
            return

    chart_df, x_col, x_order, highlight_days, time_mode = build_time_series_df(df, params)
    axis = chart_axis_config(time_mode)
    st.metric("Total transactions (filtered)", format_total_kpi(chart_df))

    mop_long = melt_metric_columns(chart_df, x_col, MOP_TYPES, x_order)
    st.plotly_chart(
        line_series_chart(
            mop_long,
            x_col,
            mop_chart_title(time_mode),
            y_title="Transaction count",
            x_axis_title=axis["x_axis_title"],
            hover_bucket_label=axis["hover_bucket_label"],
            x_order=x_order,
            highlight_every_nth_day=5 if highlight_days else None,
        ),
        use_container_width=True,
    )


def main() -> None:
    st.title("NHIT Toll Analytics Dashboard")
    st.caption("Traffic volume, vehicle class mix, and mode-of-payment trends from nhit_analytics.")

    try:
        plazas = fetch_plaza_names()
    except Exception as exc:
        st.error(f"Database connection failed: {exc}")
        st.stop()

    if not plazas:
        st.warning("No plaza data found in the analytics table. Run module1.py first.")
        st.stop()

    with st.sidebar:
        selected_plaza = st.selectbox("Plaza", plazas)
        if st.button("Refresh data"):
            load_data.clear()

    df = load_data(selected_plaza)
    st.session_state["df"] = df

    if df.empty:
        st.warning(f"No analytics rows for plaza '{selected_plaza}'.")
        st.stop()

    tab_volume, tab_class, tab_mop = st.tabs(
        ["Volume & Flow", "Class Mix", "MOP Distribution"]
    )

    with tab_volume:
        render_volume_section(df)

    with tab_class:
        render_class_mix_section(df)

    with tab_mop:
        render_mop_section(df)


if __name__ == "__main__":
    main()
