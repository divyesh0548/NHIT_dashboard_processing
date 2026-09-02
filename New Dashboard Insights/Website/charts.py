"""Plotly chart builders."""

from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from config import CHART_HEIGHT

PLOTLY_TEMPLATE = "plotly_white"
HOVER_LABEL_STYLE = dict(
    bgcolor="#1f2937",
    bordercolor="#9ca3af",
    font=dict(color="#ffffff", size=12),
    namelength=-1,
)

def _base_layout(
    title: str,
    y_title: str,
    *,
    show_legend: bool = True,
    hovermode: str = "x unified",
) -> dict:
    layout = dict(
        template=PLOTLY_TEMPLATE,
        height=CHART_HEIGHT,
        autosize=True,
        title=dict(text=title, x=0, xanchor="left", y=0.98, yanchor="top"),
        margin=dict(l=60, r=40, t=70, b=140 if show_legend else 90),
        yaxis_title=y_title,
        hovermode=hovermode,
        hoverlabel=HOVER_LABEL_STYLE,
    )
    if show_legend:
        layout["legend"] = dict(
            orientation="h",
            yanchor="top",
            y=-0.28,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(0,0,0,0)",
            borderwidth=0,
            font=dict(size=11),
            tracegroupgap=10,
            itemclick="toggleothers",
            itemdoubleclick="toggle",
        )
    return layout


def _configure_category_xaxis(
    fig: go.Figure,
    x_order: list[str] | None,
    x_axis_title: str,
) -> None:
    categories = [str(value) for value in x_order] if x_order else None
    fig.update_xaxes(
        title=x_axis_title,
        type="category",
        categoryorder="array" if categories else "category ascending",
        categoryarray=categories,
        tickmode="array" if categories else "auto",
        tickvals=categories,
        ticktext=categories,
    )


def _apply_monthly_highlights(fig: go.Figure, categories: list[str], every_n: int = 5) -> None:
    ticktext: list[str] = []
    for index, label in enumerate(categories, start=1):
        day_num = int(label) if str(label).isdigit() else index
        ticktext.append(f"<b>{label}</b>" if day_num % every_n == 0 else str(label))

    fig.update_xaxes(tickmode="array", tickvals=categories, ticktext=ticktext)


def line_series_chart(
    long_df: pd.DataFrame,
    x_col: str,
    title: str,
    y_title: str = "Vehicle count",
    x_axis_title: str = "Time",
    hover_bucket_label: str = "Time",
    x_order: list[str] | None = None,
    series_order: list[str] | None = None,
    highlight_every_nth_day: int | None = None,
) -> go.Figure:
    if long_df.empty:
        fig = go.Figure()
        fig.update_layout(**_base_layout(title, y_title))
        fig.add_annotation(text="No data for the selected filters", showarrow=False, y=0.5, x=0.5)
        return fig

    plot_df = long_df.copy()
    plot_df[x_col] = plot_df[x_col].astype(str)
    plot_df["series"] = plot_df["series"].astype(str)

    category_orders: dict[str, list[str]] = {}
    if x_order:
        category_orders[x_col] = [str(value) for value in x_order]
    if series_order:
        category_orders["series"] = [str(value) for value in series_order]

    fig = px.line(
        plot_df,
        x=x_col,
        y="count",
        color="series",
        markers=True,
        category_orders=category_orders or None,
        labels={x_col: x_axis_title, "count": y_title, "series": ""},
    )
    fig.update_layout(**_base_layout(title, y_title))

    categories = [str(value) for value in x_order] if x_order else list(dict.fromkeys(plot_df[x_col]))
    _configure_category_xaxis(fig, categories, x_axis_title)
    fig.update_xaxes(hoverformat=None)

    if highlight_every_nth_day and categories:
        _apply_monthly_highlights(fig, categories, every_n=highlight_every_nth_day)

    fig.update_traces(
        hovertemplate="%{y:,}<extra>%{fullData.name}</extra>",
        hoverlabel=HOVER_LABEL_STYLE,
    )
    return fig


def class_bar_chart(df: pd.DataFrame, column_map: dict[str, str], title: str) -> go.Figure:
    if df.empty:
        fig = go.Figure()
        fig.update_layout(**_base_layout(title, "Count", show_legend=False, hovermode="closest"))
        fig.add_annotation(text="No data for the selected filters", showarrow=False, y=0.5, x=0.5)
        return fig

    totals = {label: int(df[col].sum()) for label, col in column_map.items() if col in df.columns}
    chart_df = pd.DataFrame(
        {"class": list(totals.keys()), "count": list(totals.values())}
    )
    grand_total = chart_df["count"].sum()
    chart_df["pct"] = chart_df["count"].apply(
        lambda value: (value / grand_total * 100) if grand_total else 0
    )
    chart_df["pct_label"] = chart_df["pct"].map(lambda value: f"{value:.1f}%")

    fig = px.bar(
        chart_df,
        x="class",
        y="count",
        text="pct_label",
        labels={"class": "Vehicle class", "count": "Count"},
    )
    fig.update_layout(
        **_base_layout(title, "Count", show_legend=False, hovermode="closest"),
        showlegend=False,
        xaxis=dict(title="Vehicle class", type="category", tickangle=-30),
    )
    fig.update_traces(
        textposition="outside",
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Count: %{y:,}<br>"
            "Share: %{text}<extra></extra>"
        ),
        hoverlabel=HOVER_LABEL_STYLE,
    )
    return fig
