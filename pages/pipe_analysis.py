"""
Pipe-Level Analysis page: browse per-pipe repair data parsed from project
sheets during the daily Excel import (see project_parser.py).
"""

from __future__ import annotations

import dash
import plotly.express as px
from dash import Input, Output, callback, dash_table, dcc, html

from database import load_pipe_repair_details
from pipe_analysis import summarize_pipe_totals_by_sheet, worst_pipes

dash.register_page(__name__, path="/pipe-analysis", name="Pipe Analysis")


def layout():
    return html.Div(
        [
            html.Section(
                [
                    html.H2("Pipe-Level Analysis"),
                    html.P(
                        "Per-pipe repair data parsed from each project sheet during "
                        "Excel import. Pick a report date and a project sheet to inspect.",
                        className="help-text",
                    ),
                    html.Div(
                        [
                            html.Div(
                                [html.Label("Report Date"), dcc.Dropdown(id="pipe-date-dropdown")],
                                className="filter-field",
                            ),
                            html.Div(
                                [html.Label("Project Sheet"), dcc.Dropdown(id="pipe-sheet-dropdown")],
                                className="filter-field",
                            ),
                        ],
                        className="filter-row",
                    ),
                ],
                className="card",
            ),
            html.Section(dcc.Loading(html.Div(id="pipe-analysis-content")), className="card"),
        ]
    )


@callback(
    Output("pipe-date-dropdown", "options"),
    Output("pipe-date-dropdown", "value"),
    Input("pipe-date-dropdown", "id"),  # fires once, on page load
)
def load_dates(_):
    df = load_pipe_repair_details()
    if df.empty:
        return [], None
    dates = sorted(df["date"].dt.date.unique(), reverse=True)
    options = [{"label": d.strftime("%d.%m.%Y"), "value": d.isoformat()} for d in dates]
    return options, options[0]["value"]


@callback(
    Output("pipe-sheet-dropdown", "options"),
    Output("pipe-sheet-dropdown", "value"),
    Input("pipe-date-dropdown", "value"),
)
def load_sheets(selected_date):
    if not selected_date:
        return [], None
    df = load_pipe_repair_details()
    if df.empty:
        return [], None
    day_df = df[df["date"].dt.date.astype(str) == selected_date]
    sheets = sorted(day_df["project_sheet"].unique())
    options = [{"label": s, "value": s} for s in sheets]
    return options, options[0]["value"] if options else None


@callback(
    Output("pipe-analysis-content", "children"),
    Input("pipe-date-dropdown", "value"),
    Input("pipe-sheet-dropdown", "value"),
)
def render_pipe_analysis(selected_date, selected_sheet):
    if not selected_date:
        return html.P("No pipe-level data in the database yet. Import an Excel file first.")

    df = load_pipe_repair_details()
    day_df = df[df["date"].dt.date.astype(str) == selected_date]

    summary_table = summarize_pipe_totals_by_sheet(day_df)
    summary_section = html.Div(
        [
            html.H3("All Project Sheets — Summary"),
            dash_table.DataTable(
                data=summary_table.round(4).to_dict("records"),
                columns=[{"name": c, "id": c} for c in summary_table.columns],
                page_size=10,
                sort_action="native",
                style_table={"overflowX": "auto"},
                style_cell={"fontFamily": "inherit", "fontSize": "13px", "padding": "6px"},
            ),
        ]
    )

    if not selected_sheet:
        return summary_section

    sheet_df = day_df[day_df["project_sheet"] == selected_sheet].sort_values("pipe_no")

    worst_fig = px.bar(
        worst_pipes(sheet_df, top_n=15),
        x="pipe_no",
        y="repair_ratio",
        labels={"pipe_no": "Pipe No.", "repair_ratio": "Repair Ratio"},
        title=f"Top 15 Pipes by Repair Ratio — {selected_sheet}",
    )
    worst_fig.update_layout(template="plotly_white", margin=dict(l=40, r=20, t=50, b=40))
    worst_fig.update_xaxes(type="category")

    detail_columns = [c for c in sheet_df.columns if c != "date"]
    detail_table = dash_table.DataTable(
        data=sheet_df[detail_columns].round(4).to_dict("records"),
        columns=[{"name": c, "id": c} for c in detail_columns],
        page_size=20,
        sort_action="native",
        filter_action="native",
        style_table={"overflowX": "auto"},
        style_cell={"fontFamily": "inherit", "fontSize": "13px", "padding": "6px"},
    )

    return html.Div(
        [
            summary_section,
            html.H3(f"{selected_sheet} — Pipe Details"),
            dcc.Graph(figure=worst_fig),
            detail_table,
        ]
    )
