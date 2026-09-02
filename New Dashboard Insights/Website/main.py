"""
NHIT Analytics Dashboard — Streamlit + Plotly.

Run from project root:
    streamlit run "New Dashboard Insights/Website/main.py"
"""

from __future__ import annotations

import calendar
from datetime import date, timedelta

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
    active_lane_map,
    aggregate_daily_range,
    aggregate_hourly_window,
    aggregate_monthly_daily,
    format_hour_window,
    format_total_kpi,
    hour_labels_from_range,
    hour_sort_key,
    melt_metric_columns,
    melt_time_buckets_by_category,
    series_order_by_total,
    series_order_from_labels,
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
VIEW_VOLUME = "Vehicle class & Lane distribution"
VIEW_CLASS_MIX = "Class Mix"
VIEW_MOP = "MOP Distribution"
DASHBOARD_VIEWS = (VIEW_VOLUME, VIEW_CLASS_MIX, VIEW_MOP)

AXIS_CONFIG = {
    "Hours": {"x_axis_title": "Hour", "hover_bucket_label": "Hour"},
    "Day on day": {"x_axis_title": "Date", "hover_bucket_label": "Date"},
    "Month": {"x_axis_title": "Date", "hover_bucket_label": "Date"},
}


VOLUME_CLASS_X_AXIS_OPTIONS = ("Time", "Vehicle class")
VOLUME_LANE_X_AXIS_OPTIONS = ("Time", "Lanes")
DEFAULT_HOUR_RANGE = (8, 14)
CONTROL_MIN_WIDTH = "180px"
SLIDER_MIN_WIDTH = "260px"
DAY_TICK_DENSITY_LIMIT = 25
COMPACT_DAY_TICK_DENSITY_LIMIT = 16
PLOTLY_CHART_CONFIG = {
    "scrollZoom": False,
    "displayModeBar": False,
    "responsive": True,
    "doubleClick": False,
    "modeBarButtonsToRemove": [
        "zoom2d",
        "pan2d",
        "select2d",
        "lasso2d",
        "autoScale2d",
        "resetScale2d",
    ],
}


def render_plotly_chart(figure) -> None:
    st.plotly_chart(
        figure,
        width="stretch",
        height="content",
        config=PLOTLY_CHART_CONFIG,
    )


def apply_compact_control_styles() -> None:
    st.markdown(
        f"""
        <style>
        .block-container {{
            padding-top: 1.6rem;
        }}
        h1 {{
            letter-spacing: 0;
            margin-bottom: 0.15rem;
        }}
        h3 {{
            margin-top: 1.35rem;
            margin-bottom: 0.35rem;
            letter-spacing: 0;
        }}
        div[data-testid="stTabs"] {{
            margin-top: 0;
        }}
        div[data-testid="stTabs"] [data-baseweb="tab-border"] {{
            display: none !important;
        }}
        div[data-testid="stTabs"] button {{
            padding-top: 0.8rem;
            padding-bottom: 0.8rem;
        }}
        div[class*="lane_chart_controls"] {{
            width: 100% !important;
            max-width: 100% !important;
        }}
        div[class*="lane_chart_controls"] div[data-testid="stHorizontalBlock"] {{
            width: fit-content !important;
            max-width: 100% !important;
            align-items: flex-end !important;
            column-gap: clamp(0.85rem, 2vw, 1.5rem) !important;
            flex-wrap: nowrap !important;
        }}
        div[class*="lane_chart_controls"] div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {{
            flex: 0 1 auto !important;
            width: auto !important;
            min-width: 0 !important;
            padding: 0 !important;
        }}
        div[class*="lane_chart_controls"] div[data-testid="stSelectbox"] > div {{
            width: clamp(12rem, 22vw, 18rem) !important;
            min-width: 0 !important;
            max-width: 100% !important;
        }}
        div[class*="class_chart_controls"] {{
            width: 100% !important;
            max-width: 100% !important;
        }}
        div[class*="class_chart_controls"] div[data-testid="stSelectbox"] > div {{
            width: clamp(12rem, 22vw, 18rem) !important;
            min-width: 0 !important;
            max-width: 100% !important;
        }}
        div[class*="lane_chart_controls"] div[data-testid="stCheckbox"],
        div[class*="lane_chart_controls"] div[data-testid="stToggle"] {{
            min-width: max-content !important;
        }}
        div[class*="_time_controls"] {{
            border-bottom: 1px solid rgba(128, 128, 128, 0.35);
            padding-bottom: 0.85rem;
            margin-bottom: 1.25rem;
        }}
        div[data-testid="stPlotlyChart"],
        div[data-testid="stPlotlyChart"] > div,
        div[data-testid="stPlotlyChart"] iframe,
        div[data-testid="stElementContainer"]:has([data-testid="stPlotlyChart"]),
        [data-testid="stVerticalBlockBorderWrapper"]:has([data-testid="stPlotlyChart"]) {{
            overflow: hidden !important;
            overflow-x: hidden !important;
            overflow-y: hidden !important;
            scrollbar-width: none !important;
        }}
        div[data-testid="stPlotlyChart"]::-webkit-scrollbar,
        div[data-testid="stPlotlyChart"] > div::-webkit-scrollbar,
        div[data-testid="stElementContainer"]:has([data-testid="stPlotlyChart"])::-webkit-scrollbar {{
            display: none !important;
            width: 0 !important;
            height: 0 !important;
        }}
        div[data-testid="stPlotlyChart"] {{
            overscroll-behavior: none !important;
            max-height: none !important;
            touch-action: pan-y !important;
        }}
        div[data-testid="stSelectbox"] > div,
        div[data-testid="stDateInput"] > div {{
            width: 100%;
            min-width: {CONTROL_MIN_WIDTH};
            max-width: 100%;
        }}
        div[data-testid="stSlider"] {{
            width: 100%;
            min-width: {SLIDER_MIN_WIDTH};
            max-width: 100%;
            padding-top: 0.15rem;
        }}
        div[data-testid="stWidgetLabel"] {{
            margin-bottom: 0.2rem;
        }}
        div[data-testid="stWidgetLabel"] label {{
            font-weight: 600;
            line-height: 1.2;
        }}
        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div {{
            border-radius: 6px;
            min-height: 40px;
            box-shadow: none;
        }}
        div[data-testid="stPlotlyChart"] {{
            margin-top: 0.15rem;
        }}
        div[class*="_category_selector"] div[data-testid="stCheckbox"] {{
            min-width: max-content !important;
        }}
        div[class*="_category_selector"] div[data-testid="stCheckbox"] label p {{
            white-space: nowrap;
            font-size: 0.95rem;
        }}
        div[class*="_category_selector"] div[data-testid="stHorizontalBlock"] {{
            width: 100% !important;
            max-width: 100% !important;
            column-gap: 0.5rem !important;
            row-gap: 0.45rem !important;
            flex-wrap: wrap !important;
            justify-content: flex-start !important;
            align-items: center !important;
            overflow: visible !important;
        }}
        div[class*="_category_selector"] div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {{
            flex: 0 0 auto !important;
            width: auto !important;
            min-width: max-content !important;
            padding: 0.25rem 0.55rem !important;
            border: 1px solid rgba(128, 128, 128, 0.35) !important;
            border-radius: 6px;
            box-sizing: border-box;
        }}
        @media (max-width: 768px) {{
            div[data-testid="stSelectbox"] > div,
            div[data-testid="stDateInput"] > div,
            div[data-testid="stSlider"] {{
                max-width: 100%;
                min-width: 0;
            }}
            div[class*="_category_selector"] div[data-testid="stHorizontalBlock"] {{
                width: 100% !important;
            }}
            div[class*="lane_chart_controls"] div[data-testid="stHorizontalBlock"] {{
                width: 100% !important;
                flex-wrap: wrap !important;
            }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


SECTION_KEY_TO_VIEW = {
    "volume": VIEW_VOLUME,
    "class_mix": VIEW_CLASS_MIX,
    "mop": VIEW_MOP,
}

def mark_dashboard_view(view: str) -> None:
    st.session_state["dashboard_kpi_source"] = view


def mark_dashboard_view_for_section(section_key: str) -> None:
    view = SECTION_KEY_TO_VIEW.get(section_key)
    if view:
        mark_dashboard_view(view)


def section_date_input(section_key: str, label: str, *, key: str, **kwargs) -> date:
    return st.date_input(
        label,
        key=key,
        on_change=mark_dashboard_view_for_section,
        args=(section_key,),
        **kwargs,
    )


def section_selectbox(section_key: str, label: str, *, key: str, **kwargs):
    return st.selectbox(
        label,
        key=key,
        on_change=mark_dashboard_view_for_section,
        args=(section_key,),
        **kwargs,
    )


def filter_df_by_params(df: pd.DataFrame, params: dict) -> pd.DataFrame:
    mode = params["mode"]
    if mode == "Hours":
        return df[(df["date"] == params["selected_date"]) & (df["hour"].isin(params["hour_labels"]))]
    if mode == "Day on day":
        return df[(df["date"] >= params["start_date"]) & (df["date"] <= params["end_date"])]
    month_start = date(params["year"], params["month"], 1)
    month_end = date(params["year"], params["month"], calendar.monthrange(params["year"], params["month"])[1])
    return df[(df["date"] >= month_start) & (df["date"] <= month_end)]


def build_params_from_session(section_key: str, dates: list[date]) -> dict | None:
    mode_key = f"{section_key}_mode"
    if mode_key not in st.session_state:
        return None

    mode = st.session_state[mode_key]
    params: dict = {"mode": mode}
    flexible_hours = section_key == "volume"

    if mode == "Hours":
        date_key = f"{section_key}_hour_date"
        range_key = f"{section_key}_hour_range"
        if date_key not in st.session_state or range_key not in st.session_state:
            return None
        start_hour, end_hour = st.session_state[range_key]
        hour_labels = hour_labels_from_range(start_hour, end_hour)
        hour_valid, _ = validate_hour_range(start_hour, end_hour, enforce_max=not flexible_hours)
        if not hour_valid:
            return None
        params["selected_date"] = st.session_state[date_key]
        params["start_hour"] = start_hour
        params["hour_labels"] = hour_labels
        return params

    if mode == "Day on day":
        start_key = f"{section_key}_daily_start"
        end_key = f"{section_key}_daily_end"
        if start_key not in st.session_state or end_key not in st.session_state:
            return None
        params["start_date"] = st.session_state[start_key]
        params["end_date"] = st.session_state[end_key]
        day_valid, _ = validate_day_range(params["start_date"], params["end_date"])
        if not day_valid:
            return None
        return params

    year_key = f"{section_key}_year"
    month_key = f"{section_key}_month"
    if year_key not in st.session_state or month_key not in st.session_state:
        return None
    params["year"] = st.session_state[year_key]
    params["month"] = st.session_state[month_key]
    return params


def resolve_active_dashboard_view() -> str:
    """Sync KPI source from Streamlit tab state when available."""
    views = list(DASHBOARD_VIEWS)
    for key, value in st.session_state.items():
        if not isinstance(key, str) or "tab" not in key.lower():
            continue
        if isinstance(value, int) and 0 <= value < len(views):
            st.session_state["dashboard_kpi_source"] = views[value]
            return views[value]
        if isinstance(value, str) and value in views:
            st.session_state["dashboard_kpi_source"] = value
            return value
    return st.session_state.get("dashboard_kpi_source", VIEW_VOLUME)


def compute_section_kpi(df: pd.DataFrame, view: str) -> tuple[str, str] | None:
    if df.empty:
        return None

    section_key = {VIEW_VOLUME: "volume", VIEW_CLASS_MIX: "class_mix", VIEW_MOP: "mop"}.get(view)
    if not section_key:
        return None

    dates = available_dates(df)
    if not dates:
        return None

    params = build_params_from_session(section_key, dates)
    if not params:
        return None

    if view == VIEW_CLASS_MIX:
        filtered = filter_df_by_params(df, params)
        return ("Total vehicles", format_total_kpi(filtered))

    chart_df, _, _, _, _ = build_time_series_df(df, params)
    label = "Total transactions" if view == VIEW_MOP else "Total vehicles"
    return (label, format_total_kpi(chart_df))


def render_dashboard_title() -> st.delta_generator.DeltaGenerator:
    title_col, total_col = st.columns([4, 2], vertical_alignment="center")
    with title_col:
        st.title("NHIT Toll Analytics Dashboard")
        st.caption("Traffic volume, vehicle class mix, and mode-of-payment trends from nhit_analytics.")
    return total_col.empty()


def render_dashboard_kpi(kpi_slot: st.delta_generator.DeltaGenerator, label: str, total: str) -> None:
    with kpi_slot.container():
        st.markdown(
            f"""
            <div class="dashboard-kpi" style="text-align: right; white-space: nowrap; line-height: 1.3;">
                <span style="color: #6b7280; font-size: 0.95rem;">{label}:</span>
                <span style="font-size: 1.35rem; font-weight: 600; margin-left: 0.35rem;">{total}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_chart_section_header(section_title: str) -> None:
    st.subheader(section_title)


def category_checkbox_key(chart_key: str, category: str) -> str:
    return f"{chart_key}_cat_{category}"


def render_category_selector(chart_key: str, category_options: list[str]) -> list[str]:
    options = [str(category) for category in category_options]
    if not options:
        return []

    options_key = f"{chart_key}_category_options"
    if st.session_state.get(options_key) != options:
        st.session_state[options_key] = options
        for category in options:
            st.session_state[category_checkbox_key(chart_key, category)] = True

    for category in options:
        checkbox_key = category_checkbox_key(chart_key, category)
        if checkbox_key not in st.session_state:
            st.session_state[checkbox_key] = True

    with st.container(key=f"{chart_key}_category_selector"):
        st.caption("Visible categories")
        cols = st.columns(len(options), gap="small")
        selected: list[str] = []
        for index, category in enumerate(options):
            with cols[index]:
                if st.checkbox(category, key=category_checkbox_key(chart_key, category)):
                    selected.append(category)

    return selected


def format_period_label(params: dict) -> str:
    mode = params["mode"]
    if mode == "Hours":
        return f"{params['selected_date']} · {format_hour_window(params['hour_labels'])}"
    if mode == "Day on day":
        return f"{params['start_date']} – {params['end_date']}"
    return f"{calendar.month_name[params['month']]} {params['year']}"


def default_comparison_date(dates: list[date], primary_date: date) -> date:
    alternatives = [value for value in dates if value > primary_date]
    return alternatives[0] if alternatives else dates[0]


def default_comparison_day_range(
    dates: list[date],
    primary_start: date,
    primary_end: date,
) -> tuple[date, date]:
    primary_days = (primary_end - primary_start).days + 1
    window = min(max(primary_days, MIN_DAY_WINDOW), MAX_DAY_WINDOW)
    next_start = primary_end + timedelta(days=1)
    next_end = next_start + timedelta(days=window - 1)
    if next_end <= dates[-1]:
        return next_start, next_end

    fallback_end = dates[-1]
    fallback_start = max(dates[0], fallback_end - timedelta(days=window - 1))
    return fallback_start, fallback_end


def default_comparison_year_month(
    dates: list[date],
    primary_year: int,
    primary_month: int,
) -> tuple[int, int]:
    year_months = sorted({(value.year, value.month) for value in dates})
    primary_pair = (primary_year, primary_month)
    alternatives = [pair for pair in year_months if pair > primary_pair]
    if alternatives:
        return alternatives[0]
    return year_months[0]


def initialize_comparison_period(
    section_key: str,
    dates: list[date],
) -> None:
    """Seed comparison widgets with the next period when comparison is enabled."""
    mark_dashboard_view_for_section(section_key)
    if not st.session_state.get(f"{section_key}_compare"):
        return

    primary_params = build_params_from_session(section_key, dates)
    if not primary_params:
        return

    mode = primary_params["mode"]
    if mode == "Hours":
        start_hour = primary_params["start_hour"]
        end_hour = start_hour + len(primary_params["hour_labels"])
        st.session_state[f"{section_key}_cmp_hour_date"] = default_comparison_date(
            dates,
            primary_params["selected_date"],
        )
        st.session_state[f"{section_key}_cmp_hour_range"] = (start_hour, end_hour)
        return

    if mode == "Day on day":
        start_date, end_date = default_comparison_day_range(
            dates,
            primary_params["start_date"],
            primary_params["end_date"],
        )
        st.session_state[f"{section_key}_cmp_daily_start"] = start_date
        st.session_state[f"{section_key}_cmp_daily_end"] = end_date
        return

    year, month = default_comparison_year_month(
        dates,
        primary_params["year"],
        primary_params["month"],
    )
    st.session_state[f"{section_key}_cmp_year"] = year
    st.session_state[f"{section_key}_cmp_month"] = month


def filter_long_df_series(long_df: pd.DataFrame, selected_categories: list[str]) -> pd.DataFrame:
    if long_df.empty:
        return long_df
    if not selected_categories:
        return long_df.iloc[0:0].copy()
    selected = {str(value) for value in selected_categories}
    return long_df[long_df["series"].astype(str).isin(selected)].copy()


def filter_long_df_x_values(
    long_df: pd.DataFrame,
    x_col: str,
    selected_categories: list[str],
) -> pd.DataFrame:
    if long_df.empty:
        return long_df
    if not selected_categories:
        return long_df.iloc[0:0].copy()
    selected = {str(value) for value in selected_categories}
    return long_df[long_df[x_col].astype(str).isin(selected)].copy()


def category_totals(chart_df: pd.DataFrame, column_map: dict[str, str]) -> dict[str, int]:
    totals: dict[str, int] = {}
    for label, col in column_map.items():
        if col in chart_df.columns:
            totals[str(label)] = int(pd.to_numeric(chart_df[col], errors="coerce").fillna(0).sum())
    return totals


def format_chart_total(chart_df: pd.DataFrame, column_map: dict[str, str]) -> str:
    return f"{sum(category_totals(chart_df, column_map).values()):,}"


def day_x_tick_stride(time_mode: str, x_order: list | None, *, compact: bool) -> int:
    if time_mode != "Day on day" or not x_order:
        return 1
    density_limit = COMPACT_DAY_TICK_DENSITY_LIMIT if compact else DAY_TICK_DENSITY_LIMIT
    return 2 if len(x_order) > density_limit else 1


def build_line_chart_figure(
    chart_df: pd.DataFrame,
    column_map: dict[str, str],
    *,
    params: dict,
    time_mode: str,
    x_col: str,
    x_order: list | None,
    highlight_days: bool,
    categories_on_x_axis: bool,
    legend_mode: str,
    fixed_legend_order: list[str] | None,
    selected_categories: list[str],
    title: str,
    x_axis_title: str,
    hover_bucket_label: str,
    period_label: str,
    y_title: str = "Vehicle count",
    compact: bool = False,
):
    long_df, plot_x_col, plot_x_order, series_order = build_volume_chart_data(
        chart_df,
        column_map,
        categories_on_x_axis=categories_on_x_axis,
        time_mode=time_mode,
        params=params,
        x_col=x_col,
        x_order=x_order,
        legend_mode=legend_mode,
        fixed_legend_order=fixed_legend_order,
    )
    totals = category_totals(chart_df, column_map)
    series_totals = totals
    x_value_totals = None

    if categories_on_x_axis:
        long_df = filter_long_df_x_values(long_df, plot_x_col, selected_categories)
        if plot_x_order:
            plot_x_order = [value for value in plot_x_order if str(value) in selected_categories]
        series_totals = None
        x_value_totals = totals
    else:
        long_df = filter_long_df_series(long_df, selected_categories)
        if series_order:
            series_order = [value for value in series_order if value in selected_categories]

    x_tick_stride = day_x_tick_stride(
        time_mode,
        plot_x_order,
        compact=compact,
    ) if not categories_on_x_axis else 1

    return line_series_chart(
        long_df,
        plot_x_col,
        f"{title} — {period_label}",
        y_title=y_title,
        x_axis_title=x_axis_title,
        hover_bucket_label=hover_bucket_label,
        x_order=plot_x_order,
        series_order=series_order,
        highlight_every_nth_day=5 if highlight_days and not categories_on_x_axis else None,
        compact=compact,
        series_totals=series_totals,
        x_value_totals=x_value_totals,
        x_tick_stride=x_tick_stride,
    )


def render_comparable_line_chart(
    chart_key: str,
    df: pd.DataFrame,
    column_map: dict[str, str],
    primary_params: dict,
    *,
    time_mode: str,
    chart_df: pd.DataFrame,
    x_col: str,
    x_order: list | None,
    highlight_days: bool,
    categories_on_x_axis: bool,
    legend_mode: str,
    fixed_legend_order: list[str] | None,
    title: str,
    x_axis_title: str,
    hover_bucket_label: str,
    compare_enabled: bool,
    comparison_params: dict | None,
    y_title: str = "Vehicle count",
    total_label: str = "Total vehicles",
) -> None:
    category_options = list(fixed_legend_order or column_map.keys())
    selected_categories = render_category_selector(chart_key, category_options)

    primary_label = format_period_label(primary_params)

    if not compare_enabled:
        render_plotly_chart(
            build_line_chart_figure(
                chart_df,
                column_map,
                params=primary_params,
                time_mode=time_mode,
                x_col=x_col,
                x_order=x_order,
                highlight_days=highlight_days,
                categories_on_x_axis=categories_on_x_axis,
                legend_mode=legend_mode,
                fixed_legend_order=fixed_legend_order,
                selected_categories=selected_categories,
                title=title,
                x_axis_title=x_axis_title,
                hover_bucket_label=hover_bucket_label,
                period_label=primary_label,
                y_title=y_title,
                compact=False,
            ),
        )
        return

    if comparison_params is None:
        return

    left_col, right_col = st.columns(2, gap="small")
    with left_col:
        st.metric(total_label, format_chart_total(chart_df, column_map))
        render_plotly_chart(
            build_line_chart_figure(
                chart_df,
                column_map,
                params=primary_params,
                time_mode=time_mode,
                x_col=x_col,
                x_order=x_order,
                highlight_days=highlight_days,
                categories_on_x_axis=categories_on_x_axis,
                legend_mode=legend_mode,
                fixed_legend_order=fixed_legend_order,
                selected_categories=selected_categories,
                title=title,
                x_axis_title=x_axis_title,
                hover_bucket_label=hover_bucket_label,
                period_label=primary_label,
                y_title=y_title,
                compact=True,
            ),
        )

    with right_col:
        compare_chart_df, compare_x_col, compare_x_order, compare_highlight_days, _ = build_time_series_df(
            df,
            comparison_params,
        )
        compare_label = format_period_label(comparison_params)
        st.metric(total_label, format_chart_total(compare_chart_df, column_map))
        render_plotly_chart(
            build_line_chart_figure(
                compare_chart_df,
                column_map,
                params=comparison_params,
                time_mode=time_mode,
                x_col=compare_x_col,
                x_order=compare_x_order,
                highlight_days=compare_highlight_days,
                categories_on_x_axis=categories_on_x_axis,
                legend_mode=legend_mode,
                fixed_legend_order=fixed_legend_order,
                selected_categories=selected_categories,
                title=title,
                x_axis_title=x_axis_title,
                hover_bucket_label=hover_bucket_label,
                period_label=compare_label,
                y_title=y_title,
                compact=True,
            ),
        )


def chart_axis_config(mode: str) -> dict[str, str]:
    return AXIS_CONFIG.get(mode, {"x_axis_title": "Time", "hover_bucket_label": "Time"})


def volume_chart_title(chart_kind: str, time_mode: str, x_axis_dim: str) -> str:
    metric = "Vehicle" if chart_kind == "class" else "Lane"
    if chart_kind == "class" and x_axis_dim == "Vehicle class":
        return f"{metric} count by vehicle class"
    if chart_kind == "lane" and x_axis_dim == "Lanes":
        return f"{metric} count by lanes"

    x_label_by_mode = {
        "Hours": "hours",
        "Day on day": "day on day",
        "Month": "month",
    }
    return f"{metric} count by {x_label_by_mode.get(time_mode, 'time')}"


def volume_x_axis_title(chart_kind: str, time_mode: str, x_axis_dim: str) -> str:
    if chart_kind == "class" and x_axis_dim == "Vehicle class":
        return "Vehicle class"
    if chart_kind == "lane" and x_axis_dim == "Lanes":
        return "Lane"
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


def validate_hour_range(start_hour: int, end_hour: int, *, enforce_max: bool = True) -> tuple[bool, str | None]:
    hour_count = end_hour - start_hour
    if hour_count < MIN_HOUR_WINDOW:
        return False, f"Select at least {MIN_HOUR_WINDOW} hours (currently {hour_count})."
    if enforce_max and hour_count > MAX_HOUR_WINDOW:
        return False, f"Select at most {MAX_HOUR_WINDOW} hours (currently {hour_count})."
    return True, None


def render_hour_range_slider(
    section_key: str,
    *,
    enforce_max: bool = True,
    key_prefix: str | None = None,
    default_range: tuple[int, int] = DEFAULT_HOUR_RANGE,
) -> tuple[int, list[str], bool]:
    widget_key = key_prefix or section_key
    hour_range = st.slider(
        "Hour range",
        min_value=0,
        max_value=24,
        value=default_range,
        key=f"{widget_key}_hour_range",
        on_change=mark_dashboard_view_for_section,
        args=(section_key,),
    )
    start_hour, end_hour = hour_range
    hour_labels = hour_labels_from_range(start_hour, end_hour)
    if hour_labels:
        st.caption(f"Selected: {format_hour_window(hour_labels)}")
    is_valid, error_message = validate_hour_range(start_hour, end_hour, enforce_max=enforce_max)
    if not is_valid:
        st.error(error_message)
    return start_hour, hour_labels, is_valid


def render_time_mode_radio(section_key: str) -> str:
    return st.radio(
        "Time granularity",
        TIME_MODES,
        horizontal=True,
        key=f"{section_key}_mode",
        on_change=mark_dashboard_view_for_section,
        args=(section_key,),
    )


def render_period_controls(
    section_key: str,
    dates: list[date],
    mode: str,
    *,
    state_key: str | None = None,
    flexible_hour_window: bool = False,
    comparison: bool = False,
    primary_params: dict | None = None,
) -> tuple[dict, bool]:
    """Render one primary or comparison period using the same control layout."""
    state_key = state_key or section_key
    params: dict = {"mode": mode}

    if mode == "Hours":
        default_date = earliest_date(dates)
        default_range = DEFAULT_HOUR_RANGE
        if comparison and primary_params:
            default_date = default_comparison_date(dates, primary_params["selected_date"])
            primary_start = primary_params["start_hour"]
            primary_end = primary_start + len(primary_params["hour_labels"])
            default_range = (primary_start, primary_end)
        date_key = f"{state_key}_hour_date"
        range_key = f"{state_key}_hour_range"
        st.session_state.setdefault(date_key, default_date)
        st.session_state.setdefault(range_key, default_range)
        col1, col2 = st.columns([1, 1.65], gap="small")
        with col1:
            params["selected_date"] = section_date_input(
                section_key,
                "Date",
                value=default_date,
                min_value=dates[0],
                max_value=dates[-1],
                key=date_key,
            )
        with col2:
            start_hour, hour_labels, hour_valid = render_hour_range_slider(
                section_key,
                key_prefix=state_key,
                enforce_max=flexible_hour_window,
                default_range=default_range,
            )
            params["start_hour"] = start_hour
            params["hour_labels"] = hour_labels
            params["hour_range_valid"] = hour_valid
        return params, hour_valid

    if mode == "Day on day":
        if comparison and primary_params:
            default_start, default_end = default_comparison_day_range(
                dates,
                primary_params["start_date"],
                primary_params["end_date"],
            )
        else:
            default_start, default_end = default_day_on_day_range(dates)
        start_key = f"{state_key}_daily_start"
        end_key = f"{state_key}_daily_end"
        st.session_state.setdefault(start_key, default_start)
        st.session_state.setdefault(end_key, default_end)
        col1, col2 = st.columns(2, gap="small")
        with col1:
            params["start_date"] = section_date_input(
                section_key,
                "Start date",
                value=default_start,
                min_value=dates[0],
                max_value=dates[-1],
                key=start_key,
            )
        with col2:
            params["end_date"] = section_date_input(
                section_key,
                "End date",
                value=default_end,
                min_value=max(params["start_date"], dates[0]),
                max_value=dates[-1],
                key=end_key,
            )
        is_valid, error_message = validate_day_range(params["start_date"], params["end_date"])
        if not is_valid:
            st.error(error_message)
        return params, is_valid

    if comparison and primary_params:
        default_year, default_month = default_comparison_year_month(
            dates,
            primary_params["year"],
            primary_params["month"],
        )
    else:
        default_year, default_month = default_year_month(dates)
    year_options = sorted({value.year for value in dates}, reverse=True)
    year_key = f"{state_key}_year"
    month_key = f"{state_key}_month"
    st.session_state.setdefault(year_key, default_year)
    col1, col2 = st.columns(2, gap="small")
    with col1:
        params["year"] = section_selectbox(
            section_key,
            "Year",
            options=year_options,
            index=year_options.index(default_year),
            key=year_key,
        )
    with col2:
        months_in_year = sorted({value.month for value in dates if value.year == params["year"]})
        month_default = default_month if default_year == params["year"] else months_in_year[0]
        st.session_state.setdefault(month_key, month_default)
        params["month"] = section_selectbox(
            section_key,
            "Month",
            options=months_in_year,
            format_func=lambda month: calendar.month_name[month],
            index=months_in_year.index(month_default),
            key=month_key,
        )
    return params, True


def render_shared_time_controls(
    section_key: str,
    dates: list[date],
    *,
    flexible_hour_window: bool = False,
    allow_compare: bool = True,
) -> tuple[dict, dict | None, bool, bool, bool]:
    """Render shared granularity and primary/comparison period controls."""
    with st.container(key=f"{section_key}_time_controls"):
        mode_col, compare_col = st.columns([3, 1], gap="small", vertical_alignment="bottom")
        with mode_col:
            mode = render_time_mode_radio(section_key)
        with compare_col:
            if allow_compare:
                compare_enabled = st.toggle(
                    "Compare periods",
                    value=False,
                    key=f"{section_key}_compare",
                    on_change=initialize_comparison_period,
                    args=(section_key, dates),
                )
            else:
                compare_enabled = False

        if compare_enabled:
            primary_col, comparison_col = st.columns(2, gap="small")
            with primary_col:
                st.markdown("**Primary period**")
                primary_params, primary_valid = render_period_controls(
                    section_key,
                    dates,
                    mode,
                    flexible_hour_window=flexible_hour_window,
                )
            with comparison_col:
                st.markdown("**Comparison period**")
                comparison_params, comparison_valid = render_period_controls(
                    section_key,
                    dates,
                    mode,
                    state_key=f"{section_key}_cmp",
                    flexible_hour_window=flexible_hour_window,
                    comparison=True,
                    primary_params=primary_params,
                )
        else:
            primary_params, primary_valid = render_period_controls(
                section_key,
                dates,
                mode,
                flexible_hour_window=flexible_hour_window,
            )
            comparison_params = None
            comparison_valid = True

    return primary_params, comparison_params, compare_enabled, primary_valid, comparison_valid


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
    categories_on_x_axis: bool,
    time_mode: str,
    params: dict,
    x_col: str,
    x_order: list | None,
    legend_mode: str = "default",
    fixed_legend_order: list[str] | None = None,
) -> tuple[pd.DataFrame, str, list | None, list | None]:
    if categories_on_x_axis:
        bucket_col, bucket_order = volume_time_bucket_config(time_mode, params, x_col, x_order)
        long_df, pivot_x_col, pivot_x_order, series_order = melt_time_buckets_by_category(
            chart_df,
            column_map,
            bucket_col,
            bucket_order,
        )
        return long_df, pivot_x_col, pivot_x_order, series_order

    long_df = melt_metric_columns(chart_df, x_col, column_map, x_order)
    series_order = None
    if legend_mode == "total":
        series_order = series_order_by_total(long_df)
    elif legend_mode == "fixed" and fixed_legend_order:
        series_order = series_order_from_labels(long_df, fixed_legend_order)

    if series_order:
        long_df = melt_metric_columns(chart_df, x_col, column_map, x_order, series_order)
    return long_df, x_col, x_order, series_order


def render_volume_section(df: pd.DataFrame) -> tuple[str, str] | None:
    if df.empty:
        return None

    dates = available_dates(df)
    if not dates:
        st.warning("No data available for the selected plaza.")
        return None

    params, comparison_params, compare_enabled, primary_valid, comparison_valid = render_shared_time_controls(
        "volume",
        dates,
        flexible_hour_window=True,
        allow_compare=True,
    )
    if not primary_valid or (compare_enabled and not comparison_valid):
        return None

    mode = params["mode"]

    chart_df, x_col, x_order, highlight_days, time_mode = build_time_series_df(df, params)
    axis = chart_axis_config(time_mode)
    section_kpi = ("Total vehicles", format_total_kpi(chart_df))

    render_chart_section_header("1. Vehicle class Chart")
    with st.container(key="class_chart_controls"):
        class_x_axis_dim = section_selectbox(
            "volume",
            "X-axis (vehicle class chart)",
            options=VOLUME_CLASS_X_AXIS_OPTIONS,
            key="volume_class_x_axis",
        )
    use_class_categories_on_x = class_x_axis_dim == "Vehicle class"
    class_title = volume_chart_title("class", time_mode, class_x_axis_dim)
    class_x_title = (
        volume_x_axis_title("class", time_mode, class_x_axis_dim)
        if use_class_categories_on_x
        else axis["x_axis_title"]
    )
    class_hover_label = class_x_title if use_class_categories_on_x else axis["hover_bucket_label"]
    render_comparable_line_chart(
        "volume_class",
        df,
        VEHICLE_CLASSES,
        params,
        time_mode=time_mode,
        chart_df=chart_df,
        x_col=x_col,
        x_order=x_order,
        highlight_days=highlight_days,
        categories_on_x_axis=use_class_categories_on_x,
        legend_mode="total" if not use_class_categories_on_x else "default",
        fixed_legend_order=None,
        title=class_title,
        x_axis_title=class_x_title,
        hover_bucket_label=class_hover_label,
        compare_enabled=compare_enabled,
        comparison_params=comparison_params,
    )

    render_chart_section_header("2. Lane Chart")
    with st.container(key="lane_chart_controls"):
        lane_col1, lane_col2 = st.columns(2, gap="small", vertical_alignment="bottom")
        with lane_col1:
            lane_x_axis_dim = section_selectbox(
                "volume",
                "X-axis (lane chart)",
                options=VOLUME_LANE_X_AXIS_OPTIONS,
                key="volume_lane_x_axis",
            )
        with lane_col2:
            show_all_lanes = st.toggle(
                "Show all lanes",
                value=False,
                key="show_all_lanes",
                on_change=mark_dashboard_view_for_section,
                args=("volume",),
            )

    if show_all_lanes:
        lane_map = active_lane_map(chart_df)
        lane_scope = "all active lanes"
    else:
        top = top_lanes_by_volume(chart_df)
        lane_map = {lane: col for lane, col in LANES.items() if lane in top}
        lane_scope = "top 5"

    use_lane_categories_on_x = lane_x_axis_dim == "Lanes"
    lane_title = f"{volume_chart_title('lane', time_mode, lane_x_axis_dim)} ({lane_scope})"
    lane_x_title = (
        volume_x_axis_title("lane", time_mode, lane_x_axis_dim)
        if use_lane_categories_on_x
        else axis["x_axis_title"]
    )
    lane_hover_label = lane_x_title if use_lane_categories_on_x else axis["hover_bucket_label"]
    render_comparable_line_chart(
        "volume_lane",
        df,
        lane_map,
        params,
        time_mode=time_mode,
        chart_df=chart_df,
        x_col=x_col,
        x_order=x_order,
        highlight_days=highlight_days,
        categories_on_x_axis=use_lane_categories_on_x,
        legend_mode="fixed" if not use_lane_categories_on_x else "default",
        fixed_legend_order=list(lane_map.keys()),
        title=lane_title,
        x_axis_title=lane_x_title,
        hover_bucket_label=lane_hover_label,
        compare_enabled=compare_enabled,
        comparison_params=comparison_params,
    )

    return section_kpi


def render_class_mix_section(df: pd.DataFrame) -> tuple[str, str] | None:
    dates = available_dates(df)
    if not dates:
        st.warning("No data available.")
        return None

    params, _, _, params_valid, _ = render_shared_time_controls(
        "class_mix",
        dates,
        flexible_hour_window=False,
        allow_compare=False,
    )
    if not params_valid:
        return None

    mode = params["mode"]
    filtered = filter_df_by_params(df, params)

    render_chart_section_header("Vehicle class mix")
    render_plotly_chart(
        class_bar_chart(filtered, VEHICLE_CLASSES, f"Class distribution — {mode}"),
    )
    return ("Total vehicles", format_total_kpi(filtered))


def render_mop_section(df: pd.DataFrame) -> tuple[str, str] | None:
    dates = available_dates(df)
    if not dates:
        st.warning("No data available.")
        return None

    params, comparison_params, compare_enabled, primary_valid, comparison_valid = render_shared_time_controls(
        "mop",
        dates,
        flexible_hour_window=False,
        allow_compare=True,
    )
    if not primary_valid or (compare_enabled and not comparison_valid):
        return None

    mode = params["mode"]
    chart_df, x_col, x_order, highlight_days, time_mode = build_time_series_df(df, params)
    axis = chart_axis_config(time_mode)
    section_kpi = ("Total transactions", format_total_kpi(chart_df))

    render_chart_section_header("FASTag / Cash / Exempt over time")
    render_comparable_line_chart(
        "mop",
        df,
        MOP_TYPES,
        params,
        time_mode=time_mode,
        chart_df=chart_df,
        x_col=x_col,
        x_order=x_order,
        highlight_days=highlight_days,
        categories_on_x_axis=False,
        legend_mode="total",
        fixed_legend_order=None,
        title=mop_chart_title(time_mode),
        x_axis_title=axis["x_axis_title"],
        hover_bucket_label=axis["hover_bucket_label"],
        compare_enabled=compare_enabled,
        comparison_params=comparison_params,
        y_title="Transaction count",
        total_label="Total transactions",
    )

    return section_kpi


def main() -> None:
    apply_compact_control_styles()

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

    kpi_slot = render_dashboard_title()

    if df.empty:
        st.warning(f"No analytics rows for plaza '{selected_plaza}'.")
        render_dashboard_kpi(kpi_slot, "Total transactions", "0")
        st.stop()

    st.session_state.setdefault("dashboard_kpi_source", VIEW_VOLUME)

    tab_volume, tab_class_mix, tab_mop = st.tabs(list(DASHBOARD_VIEWS))

    with tab_volume:
        render_volume_section(df)
    with tab_class_mix:
        render_class_mix_section(df)
    with tab_mop:
        render_mop_section(df)

    active_view = resolve_active_dashboard_view()
    section_kpi = compute_section_kpi(df, active_view)
    active_section_key = {
        VIEW_VOLUME: "volume",
        VIEW_CLASS_MIX: "class_mix",
        VIEW_MOP: "mop",
    }.get(active_view)
    comparison_active = bool(active_section_key and st.session_state.get(f"{active_section_key}_compare"))
    if section_kpi and not comparison_active:
        kpi_label, kpi_value = section_kpi
        render_dashboard_kpi(kpi_slot, kpi_label, kpi_value)


if __name__ == "__main__":
    main()
