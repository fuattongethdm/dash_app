"""
Dashboard page: daily Excel import, the main dashboard, pipe-level
drill-down, and multi-project comparison — organized as tabs
so the page doesn't grow into one long endless scroll as more sections
are added.

Flow:
  1) User uploads the Excel file (dcc.Upload)
  2) Read via parser.parse_daily_repair_rate(), validated with validators
  3) Validation result and preview are shown to the user
  4) Clicking "Confirm Import" writes the data to the database (upsert)
  5) The dashboard tab refreshes from ALL data currently in the database
"""

from __future__ import annotations

import base64
import io
import logging
import re
import time

import dash
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import Input, Output, State, callback, clientside_callback, dash_table, dcc, html
from dash.dash_table.Format import Format, Scheme, Sign
from openpyxl import load_workbook

from calculations import (
    METERS_PER_FOOT,
    amount_in_display_unit,
    apply_meter_based_repair_ratios,
    daily_weighted_repair_ratios,
    daily_weighted_repair_ratios_for_dimension,
    daily_weighted_repair_ratios_for_type,
    length_in_display_unit,
    repair_amount_trend_data,
    unit_label,
)
from database import (
    get_existing_keys,
    load_historical_baselines,
    load_master_data,
    load_pipe_repair_details,
    load_project_sheet_links,
    upsert_historical_baselines,
    upsert_pipe_repair_details,
    upsert_repair_rates,
)
from parser import parse_daily_repair_rate, parse_repair_rate_archive
from pdf_report import build_pdf_report, build_pipe_analysis_pdf_report
from pipe_analysis import worst_pipes
from project_parser import parse_project_pipe_repairs
from validators import ValidationReport, mark_duplicate_counts

dash.register_page(__name__, path="/", name="Dashboard")


# ---------------------------------------------------------------------------
# Display labels — table columns show these instead of raw field names
# ---------------------------------------------------------------------------

COLUMN_LABELS = {
    "date": "Date",
    "production_type": "Type",
    "project_no": "Project",
    "dimensions": "Dimensions",
    "qty": "Qty",
    "project_total_pipe_length": "Total Pipe Length (ft)",
    "repaired_pipes_total_length": "Repaired Pipes Length (ft)",
    "repaired_spiral_length": "Repaired Spiral Length (ft)",
    "project_status": "Status",
    "repair_ratio": "Repair Ratio",
    "repair_ratio_incl_skelp": "Repair Ratio (incl. Skelp)",
    "total_repair_amount": "Total Repair Amount (m)",
    "total_repair_amount_incl_skelp": "Total Repair Amount (incl. Skelp, m)",
    "project_sheet": "Project Sheet",
    "pipe_count": "Pipe Count",
    "avg_repair_ratio": "Avg Repair Ratio",
    "max_repair_ratio": "Max Repair Ratio",
    "block_cell": "Cell Ref.",
    "pipe_no": "Pipe No.",
    "pipe_length_ft": "Pipe Length (ft)",
    "repair_amount": "Repair Amount (m)",
    "repair_count": "B.E Count",
    "repair_category": "Category",
    "surface_state": "Surface State",
    "status": "Status",
    "repaired_date": "Repaired Date",
}

# Display-only relabeling for the pipe status value: the underlying `status`
# column keeps storing "Produced" (matches project_parser.py/database.py's
# upsert logic — renaming the stored value would need a live migration for
# no real benefit). "Produced" alone read as "already finished" to users;
# this is what's actually shown in the pie chart / detail table instead.
STATUS_DISPLAY_LABELS = {"Produced": "Awaiting Repair"}

# Ratio columns are displayed as percentages instead of raw decimals. The
# underlying value is kept numeric (via dash_table's own Format, not a
# formatted string) so native sort/filter still compare numerically instead
# of alphabetically — a string "10.24%" sorts/filters *before* "5.00%".
PERCENT_COLUMNS = {"repair_ratio", "repair_ratio_incl_skelp", "avg_repair_ratio", "max_repair_ratio"}
PERCENT_FORMAT = Format(scheme=Scheme.percentage, precision=2)
NUMERIC_FORMAT = Format(scheme=Scheme.fixed, precision=2)

TABLE_HEADER_STYLE = {
    "backgroundColor": "#f4f6f8",
    "fontWeight": "700",
    "color": "var(--color-text)",
    "borderBottom": "2px solid #cbd5e1",
    "textAlign": "left",
}
TABLE_CELL_STYLE = {"fontFamily": "inherit", "fontSize": "13px", "padding": "8px 10px"}
TABLE_CONDITIONAL_STYLE = [{"if": {"row_index": "odd"}, "backgroundColor": "#f8fafc"}]

# Consistent hover-box styling across all charts: white background with a
# light border, instead of Plotly's default (a box the same color as the
# trace, which is hard to read against a same-colored bar/line).
HOVER_STYLE = dict(
    bgcolor="white",
    bordercolor="#e2e8f0",
    font=dict(size=12, family="Inter, Segoe UI, system-ui, sans-serif", color="#1e293b"),
)

# One consistent meaning per color across the whole dashboard.
COLOR_COIL = "#2563eb"
COLOR_PLATE = "#7c3aed"
COLOR_SECONDARY = "#f97316"
COLOR_MUTED = "#94a3b8"
COLOR_DANGER = "#dc2626"

# How many of the most recent days the trend charts (Overall/Type/Project/
# Dimension Detail) show, and the window the fitted trend line is drawn
# over — one constant so both stay in sync and are easy to change later.
TREND_WINDOW_DAYS = 20

# How many of the "worst" (highest repair-ratio) bubbles get their project
# name shown on the chart instead of just the ratio value.
BUBBLE_LABEL_TOP_N = 5

# Window for "Daily Repair Amount" and "Backlog Trend" specifically —
# separate from TREND_WINDOW_DAYS (used by the ratio trend charts) since
# these two are meant to show a fixed, shorter, always-14-day span rather
# than "however much history exists."
DAILY_CHARTS_WINDOW_DAYS = 14

# Two one-time bulk-load lumps, not real day-to-day activity, that would
# otherwise dwarf real daily counts on the "Daily Production vs. Repair,
# and Backlog Trend" chart: the initial bulk import (and pipes recovered
# later by parser fixes, backdated to blend into that same original batch)
# landed on 2026-07-24; separately, activating produced-pipe tracking
# itself backfilled ~376 already-existing-but-never-before-tracked
# "Produced" pipes in a single day (2026-08-18), showing up as a fake
# same-day spike to hundreds of "Produced" + a backlog cliff, not a real
# jump in shop activity. The chart hides anything before this floor.
PIPE_TREND_FLOOR_DATE = pd.Timestamp("2026-08-19")

# Which projects the Comparison tab pre-selects on page load — edit this
# list to change them. Values are project_sheet names (the same ones
# pipe-sheet-dropdown uses), not project numbers; check
# load_project_sheet_links()'s project_sheet column for exact names.
COMPARISON_DEFAULT_SHEETS = ["01-116 (42)", "06-131"]


TEXT_COLUMNS = {
    "date",
    "project_no",
    "project",
    "label",
    "dimensions",
    "project_status",
    "production_type",
    "project_sheet",
    "block_cell",
    "repair_category",
    "surface_state",
    "status",
    "repaired_date",
}
# Whole-number counts: no decimal formatting.
INTEGER_COLUMNS = {"id", "qty", "pipe_no", "repair_count", "pipe_count", "project_count"}


def _table_columns(columns, label_overrides: dict[str, str] | None = None) -> list[dict]:
    label_overrides = label_overrides or {}
    cols = []
    for c in columns:
        name = label_overrides.get(c) or COLUMN_LABELS.get(c, c.replace("_", " ").title())
        if c in PERCENT_COLUMNS:
            cols.append({"name": name, "id": c, "type": "numeric", "format": PERCENT_FORMAT})
        elif c in TEXT_COLUMNS:
            cols.append({"name": name, "id": c})
        elif c in INTEGER_COLUMNS:
            cols.append({"name": name, "id": c, "type": "numeric"})
        else:
            cols.append({"name": name, "id": c, "type": "numeric", "format": NUMERIC_FORMAT})
    return cols


def _table_records(df: pd.DataFrame, columns: list[str]) -> list[dict]:
    """Round numeric columns for display. Percent columns keep their raw
    0-1 fraction — the column's Format (see _table_columns) renders that as
    a percentage, so filtering/sorting stays numeric rather than string.
    NaN is converted to None (e.g. a Produced pipe's still-null
    repair_amount/repair_ratio) so DataTable renders a blank cell — a raw
    float NaN survives to JSON as the bare token NaN, which the frontend's
    numeric Format renders as the literal text "NaN" instead of blank."""
    out = df[columns].copy()
    for col in columns:
        if col in PERCENT_COLUMNS:
            continue
        if pd.api.types.is_numeric_dtype(out[col]):
            out[col] = out[col].round(4)
    out = out.astype(object).where(pd.notna(out), None)
    return out.to_dict("records")


# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------

_TAB_STYLE = {
    "padding": "12px 20px",
    "fontWeight": "600",
    "color": "#64748b",
    "border": "none",
    "borderBottom": "3px solid transparent",
    "backgroundColor": "#f4f6f8",
}
_TAB_SELECTED_STYLE = {
    **_TAB_STYLE,
    "color": "#2563eb",
    "borderBottom": "3px solid #2563eb",
    "backgroundColor": "#ffffff",
}


def layout():
    return html.Div(
        [
            # Keep the parsed (not-yet-written) data in the browser.
            dcc.Store(id="parsed-data-store"),
            dcc.Store(id="parsed-pipe-data-store"),
            dcc.Store(id="parsed-archive-baseline-store"),
            # Global display-unit toggle — sits above the tabs so it stays
            # visible/selected while switching tabs, and any callback can
            # read it directly (its value doubles as shared state, no
            # separate Store needed).
            html.Div(
                [
                    html.Span("Units:", className="unit-toggle-label"),
                    dcc.RadioItems(
                        id="unit-toggle",
                        options=[
                            {"label": "Meters (m)", "value": "m"},
                            {"label": "Feet (ft)", "value": "ft"},
                        ],
                        value="ft",
                        inline=True,
                        className="unit-toggle",
                    ),
                ],
                className="unit-toggle-row",
            ),
            dcc.Tabs(
                id="dashboard-tabs",
                value="tab-dashboard",
                style={"borderBottom": "1px solid #e2e8f0"},
                children=[
                    dcc.Tab(
                        label="Import",
                        value="tab-import",
                        style=_TAB_STYLE,
                        selected_style=_TAB_SELECTED_STYLE,
                        children=[
                            html.Section(
                                [
                                    html.H2("Upload Daily Excel"),
                                    dcc.Upload(
                                        id="excel-upload",
                                        children=html.Div(
                                            ["Drag and drop the Excel file here or ", html.A("select a file")]
                                        ),
                                        className="upload-box",
                                        multiple=False,
                                    ),
                                    dcc.Loading(
                                        html.Div(
                                            [
                                                html.Div(id="upload-validation-result"),
                                                html.Div(id="upload-preview-table"),
                                                html.Div(id="pipe-parse-summary"),
                                            ]
                                        ),
                                        type="circle",
                                    ),
                                    html.Button(
                                        "Confirm Import",
                                        id="confirm-import-btn",
                                        className="primary-btn",
                                        style={"display": "none"},
                                    ),
                                    dcc.Loading(html.Div(id="import-confirm-result"), type="circle"),
                                ],
                                className="card",
                            ),
                        ],
                    ),
                    dcc.Tab(
                        label="Dashboard",
                        value="tab-dashboard",
                        style=_TAB_STYLE,
                        selected_style=_TAB_SELECTED_STYLE,
                        children=[
                            html.Section(
                                [
                                    html.Div(
                                        [
                                            html.H2("Dashboard"),
                                            html.Div(
                                                [
                                                    html.Button(
                                                        "Download PDF Report",
                                                        id="download-pdf-report-btn",
                                                        className="secondary-btn",
                                                    ),
                                                    html.Button(
                                                        "Refresh Data",
                                                        id="refresh-dashboard-btn",
                                                        className="secondary-btn",
                                                    ),
                                                ],
                                                className="button-group",
                                            ),
                                        ],
                                        className="section-header-row",
                                    ),
                                    dcc.Download(id="pdf-report-download"),
                                    dcc.Loading(html.Div(id="dashboard-content")),
                                    # Lightweight "stage ready" signals the
                                    # sections below chain off of instead of
                                    # each other's full (heavy, Plotly-JSON-
                                    # laden) children — see render_dashboard /
                                    # render_pipe_overview.
                                    dcc.Store(id="dashboard-stage-1"),
                                    dcc.Store(id="dashboard-stage-2"),
                                ],
                                className="card",
                            ),
                            html.Section(
                                [
                                    html.H2("Dimension Detail"),
                                    html.H3("All Dimensions — Overview"),
                                    html.P(
                                        "Every dimension at once (latest day) — not affected by the picker below.",
                                        className="help-text",
                                    ),
                                    dcc.Loading(html.Div(id="dimension-overview-content")),
                                    html.H3("Drill Down by Dimension"),
                                    html.P(
                                        "Select a dimension to compare all projects sharing it (latest day).",
                                        className="help-text",
                                    ),
                                    html.Div(
                                        [
                                            html.Div(
                                                [
                                                    html.Label("Dimension"),
                                                    dcc.Dropdown(id="dimension-detail-dropdown"),
                                                ],
                                                className="filter-field",
                                            ),
                                        ],
                                        className="filter-row",
                                    ),
                                    dcc.Loading(html.Div(id="dimension-detail-content")),
                                    # This table depends on the dropdown
                                    # above (changes per selected dimension),
                                    # so it stays attached to this section
                                    # instead of moving to the static Tables
                                    # section below — still below its own
                                    # charts, just not detached from the
                                    # control that drives it.
                                    dcc.Loading(html.Div(id="dimension-detail-table-content")),
                                ],
                                className="card",
                            ),
                            # Tables that don't depend on a picker inside
                            # their own section, grouped here at the very
                            # bottom of the page instead of breaking up the
                            # charts above.
                            html.Section(
                                [
                                    html.H2("Tables"),
                                    dcc.Loading(html.Div(id="dashboard-table-content")),
                                    dcc.Loading(html.Div(id="pipe-overview-content")),
                                ],
                                className="card",
                            ),
                        ],
                    ),
                    dcc.Tab(
                        label="Pipe Analysis",
                        value="tab-pipe",
                        style=_TAB_STYLE,
                        selected_style=_TAB_SELECTED_STYLE,
                        children=[
                            html.Section(
                                [
                                    html.H2("Project Trend"),
                                    html.P(
                                        "Pick a project to see its repair ratio over all report dates, and "
                                        "everything about its pipes below.",
                                        className="help-text",
                                    ),
                                    html.Div(
                                        [
                                            html.Div(
                                                [html.Label("Project"), dcc.Dropdown(id="pipe-sheet-dropdown")],
                                                className="filter-field",
                                            ),
                                            html.Div(
                                                [
                                                    html.Label("Series"),
                                                    dcc.Checklist(
                                                        id="project-trend-checklist",
                                                        options=[
                                                            {"label": "Excl. Skelp", "value": "Excl"},
                                                            {"label": "Incl. Skelp", "value": "Incl"},
                                                        ],
                                                        value=["Excl", "Incl"],
                                                        inline=True,
                                                    ),
                                                ],
                                                className="filter-field",
                                            ),
                                            html.Div(
                                                [
                                                    html.Label("Timeline"),
                                                    dcc.Checklist(
                                                        id="project-trend-compact-toggle",
                                                        options=[
                                                            {
                                                                "label": "Compact Timeline (collapse empty days)",
                                                                "value": "compact",
                                                            }
                                                        ],
                                                        value=["compact"],
                                                        inline=True,
                                                    ),
                                                ],
                                                className="filter-field",
                                            ),
                                        ],
                                        className="filter-row",
                                    ),
                                    dcc.Loading(html.Div(id="project-trend-content")),
                                ],
                                className="card",
                            ),
                            html.Section(
                                [
                                    html.Div(
                                        [
                                            html.H2("Pipe-Level Analysis"),
                                            html.Button(
                                                "Download PDF Report",
                                                id="download-pipe-pdf-btn",
                                                className="secondary-btn",
                                            ),
                                        ],
                                        className="section-header-row",
                                    ),
                                    html.P(
                                        "Per-pipe repair data parsed from each project sheet during Excel import "
                                        "for whichever project is selected above.",
                                        className="help-text",
                                    ),
                                    dcc.Download(id="pipe-pdf-download"),
                                    dcc.Loading(html.Div(id="pipe-analysis-content")),
                                ],
                                className="card",
                            ),
                        ],
                    ),
                    dcc.Tab(
                        label="Comparison",
                        value="tab-comparison",
                        style=_TAB_STYLE,
                        selected_style=_TAB_SELECTED_STYLE,
                        children=[
                            html.Section(
                                [
                                    html.H2("Comparison"),
                                    html.P(
                                        "Select two or more projects to compare their repair ratio, trend, "
                                        "and pipe-level spread side by side.",
                                        className="help-text",
                                    ),
                                    html.Div(
                                        [
                                            html.Div(
                                                [
                                                    html.Label("Projects"),
                                                    dcc.Dropdown(id="comparison-project-dropdown", multi=True),
                                                ],
                                                className="filter-field",
                                            ),
                                            html.Div(
                                                [
                                                    html.Label("Series"),
                                                    dcc.Checklist(
                                                        id="comparison-series-checklist",
                                                        options=[
                                                            {"label": "Excl. Skelp", "value": "Excl"},
                                                            {"label": "Incl. Skelp", "value": "Incl"},
                                                        ],
                                                        value=["Excl"],
                                                        inline=True,
                                                    ),
                                                ],
                                                className="filter-field",
                                            ),
                                            html.Div(
                                                [
                                                    html.Label("Timeline"),
                                                    dcc.Checklist(
                                                        id="comparison-compact-toggle",
                                                        options=[
                                                            {
                                                                "label": "Compact Timeline (collapse empty days)",
                                                                "value": "compact",
                                                            }
                                                        ],
                                                        value=["compact"],
                                                        inline=True,
                                                    ),
                                                ],
                                                className="filter-field",
                                            ),
                                        ],
                                        className="filter-row",
                                    ),
                                    dcc.Loading(html.Div(id="comparison-content")),
                                ],
                                className="card",
                            ),
                        ],
                    ),
                ],
            ),
        ]
    )


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def _decode_upload(contents: str) -> io.BytesIO:
    _, content_string = contents.split(",", 1)
    decoded = base64.b64decode(content_string)
    return io.BytesIO(decoded)


def _validation_summary(report) -> html.Div:
    if not report.ok:
        # Only show the full checklist when something actually failed, so
        # the user can see which check tripped and why.
        check_rows = [
            html.Li(f"{'✅' if ok else '❌'} {name}", className="check-ok" if ok else "check-fail")
            for name, ok in report.checks.items()
        ]
        error_rows = [html.Li(err) for err in report.errors]
        children = [
            html.Ul(check_rows),
            html.H4(f"Errors ({len(report.errors)})", className="error-heading"),
            html.Ul(error_rows, className="error-list"),
        ]
        return html.Div(children, className="validation-box fail")

    summary_text = f"✅ {report.import_rows} rows read, no errors."
    if report.insert_rows or report.update_rows:
        summary_text += (
            f" ({report.insert_rows} new, {report.update_rows} will update existing records for this date.)"
        )
    return html.Div(html.P(summary_text, className="success-text"), className="validation-box ok")


def _empty_state(message: str) -> html.Div:
    return html.Div(message, className="empty-state")


def _add_bar_value_annotations(
    fig: go.Figure,
    categories,
    values,
    display_values,
    orientation: str = "v",
    min_fraction: float = 0.15,
) -> None:
    """Write ``display_values`` (e.g. pipe quantity) inside each bar, at the
    bar's mid-point — skipped for any bar whose value is under
    ``min_fraction`` of the chart's max, a plain stand-in for "won't fit
    without overflowing" (there's no real text-vs-bar pixel measurement
    available at figure-build time)."""
    values = list(values)
    max_value = max(values) if values else 0
    if not max_value:
        return
    for category, value, display_value in zip(categories, values, display_values):
        if value < max_value * min_fraction:
            continue
        if orientation == "h":
            fig.add_annotation(
                x=value / 2,
                y=category,
                text=str(display_value),
                showarrow=False,
                font=dict(size=10, color="white"),
                xanchor="center",
            )
        else:
            fig.add_annotation(
                x=category,
                y=value / 2,
                text=str(display_value),
                showarrow=False,
                font=dict(size=10, color="white"),
                yanchor="middle",
            )


def _latest_day_frame(master_df: pd.DataFrame) -> tuple[pd.Timestamp, pd.DataFrame]:
    """Latest report date + that day's rows, with meter-based ratios and
    both project-label variants attached. Shared by render_dashboard and
    the bubble-chart callback so the two can never drift out of sync."""
    latest_date = master_df["date"].max()
    latest_df = master_df[master_df["date"] == latest_date].copy()
    latest_df = apply_meter_based_repair_ratios(latest_df)
    # project_no alone repeats within a day (same project, different
    # dimensions) — charts that plot per-row project data need a unique
    # label per row, or a categorical axis shared by two traces (e.g. the
    # pareto bar+line) will collapse repeated labels onto one x position
    # and the line will zigzag/fold back on itself.
    latest_df["project_label"] = latest_df["project_no"].astype(str) + " (" + latest_df["dimensions"].astype(str) + ")"
    stripped_project_no = latest_df["project_no"].astype(str).str.replace(r"^\d{4}Q-", "", regex=True)
    latest_df["project_label_clean"] = stripped_project_no + " (" + latest_df["dimensions"].astype(str) + ")"
    return latest_date, latest_df


# ---------------------------------------------------------------------------
# Callback 1: Upload Excel -> parse -> validate -> show preview
# ---------------------------------------------------------------------------

@callback(
    Output("upload-validation-result", "children"),
    Output("upload-preview-table", "children"),
    Output("pipe-parse-summary", "children"),
    Output("confirm-import-btn", "style"),
    Output("parsed-data-store", "data"),
    Output("parsed-pipe-data-store", "data"),
    Output("parsed-archive-baseline-store", "data"),
    Input("excel-upload", "contents"),
    State("excel-upload", "filename"),
    prevent_initial_call=True,
)
def handle_upload(contents, filename):
    if contents is None:
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update

    # Load the workbook once and hand the same one to all three parsers
    # below, instead of each re-decoding + re-opening the same file from
    # scratch — measured 24s for one 4.4MB, 60+ sheet upload before this,
    # most of it from opening the whole workbook three separate times.
    try:
        file_obj = _decode_upload(contents)
        wb = load_workbook(file_obj, read_only=True, data_only=True)
    except Exception:
        # The raw exception (e.g. "File is not a zip file") means nothing
        # to a non-technical user — just say what to check instead.
        report = ValidationReport()
        report.add_check("Sheet found?", False)
        report.add_error("Could not open this file — make sure it's a valid Excel (.xlsx) file.")
        return _validation_summary(report), None, None, {"display": "none"}, None, None, None

    try:
        df, report = parse_daily_repair_rate(wb)
    except Exception as exc:  # don't crash the app on an unexpected error
        error_box = html.Div(
            f"An unexpected error occurred while reading the file: {exc}",
            className="validation-box fail",
        )
        return error_box, None, None, {"display": "none"}, None, None, None

    if not df.empty and report.ok:
        # Tell the user upfront whether this import will overwrite existing
        # rows for this date, rather than silently upserting on confirm.
        try:
            existing_keys = get_existing_keys(df)
            report = mark_duplicate_counts(existing_keys, df, report)
        except Exception:
            pass  # best-effort; missing counts just won't be shown

    validation_box = _validation_summary(report)

    if df.empty or not report.ok:
        return validation_box, None, None, {"display": "none"}, None, None, None

    preview_columns = [c for c in df.columns if c != "excel_row"]
    preview = dash_table.DataTable(
        data=_table_records(df, preview_columns),
        columns=_table_columns(preview_columns),
        page_size=15,
        style_table={"overflowX": "auto"},
        style_cell=TABLE_CELL_STYLE,
        style_header=TABLE_HEADER_STYLE,
        style_data_conditional=TABLE_CONDITIONAL_STYLE,
    )

    summary_lines = []
    pipe_data_json = None
    try:
        pipe_df, pipe_report = parse_project_pipe_repairs(wb, df["date"].iloc[0])
        if not pipe_df.empty:
            summary_lines.append(
                html.P(
                    f"Pipe-level: {pipe_report.parsed_rows} rows parsed from "
                    f"{pipe_report.parsed_sheets} project sheets.",
                    className="help-text",
                )
            )
            pipe_data_json = pipe_df.to_json(date_format="iso", orient="split")
    except Exception as exc:  # pipe-level parsing is best-effort, never blocks the main import
        summary_lines.append(
            html.P(
                f"Pipe-level parsing failed (main import is unaffected): {exc}",
                className="help-text",
            )
        )

    pipe_summary = html.Div(summary_lines) if summary_lines else None

    archive_data_json = None
    try:
        archive_df = parse_repair_rate_archive(wb)
        if not archive_df.empty:
            archive_data_json = archive_df.to_json(orient="split")
    except Exception:
        pass  # best-effort; completed-project archive totals just won't be refreshed

    return (
        validation_box,
        html.Div([html.H4(f"Preview — {filename}"), preview], className="preview-box"),
        pipe_summary,
        {"display": "inline-block"},
        df.to_json(date_format="iso", orient="split"),
        pipe_data_json,
        archive_data_json,
    )


# ---------------------------------------------------------------------------
# Callback 2: "Confirm Import" -> write to database
# ---------------------------------------------------------------------------

@callback(
    Output("import-confirm-result", "children"),
    Output("parsed-data-store", "data", allow_duplicate=True),
    Output("parsed-pipe-data-store", "data", allow_duplicate=True),
    Output("parsed-archive-baseline-store", "data", allow_duplicate=True),
    Input("confirm-import-btn", "n_clicks"),
    State("parsed-data-store", "data"),
    State("parsed-pipe-data-store", "data"),
    State("parsed-archive-baseline-store", "data"),
    prevent_initial_call=True,
)
def confirm_import(n_clicks, stored_json, stored_pipe_json, stored_archive_json):
    if not n_clicks or not stored_json:
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update

    df = pd.read_json(io.StringIO(stored_json), orient="split")
    written = upsert_repair_rates(df)

    pipe_written = 0
    if stored_pipe_json:
        pipe_df = pd.read_json(io.StringIO(stored_pipe_json), orient="split")
        pipe_written = upsert_pipe_repair_details(pipe_df)

    archive_written = 0
    if stored_archive_json:
        try:
            archive_df = pd.read_json(io.StringIO(stored_archive_json), orient="split")
            archive_written = upsert_historical_baselines(archive_df)
        except Exception:
            pass  # best-effort; completed-project archive totals just won't be refreshed

    message = f"✅ {written} rows saved to the database."
    if pipe_written:
        message += f" ({pipe_written} pipe-level rows saved.)"
    if archive_written:
        message += f" ({archive_written} completed-project archive totals refreshed.)"

    return (
        html.Div(message, className="success-text"),
        None,
        None,
        None,
    )


# ---------------------------------------------------------------------------
# Callback 3: Download the A3 PDF report for the latest report date
# ---------------------------------------------------------------------------

@callback(
    Output("pdf-report-download", "data"),
    Input("download-pdf-report-btn", "n_clicks"),
    State("unit-toggle", "value"),
    prevent_initial_call=True,
)
def download_pdf_report(n_clicks, selected_unit):
    if not n_clicks:
        return dash.no_update

    master_df = load_master_data()
    if master_df.empty:
        return dash.no_update

    baseline_df = load_historical_baselines()
    pipe_df = load_pipe_repair_details()
    links_df = load_project_sheet_links()
    latest_date = master_df["date"].max()
    pdf_bytes = build_pdf_report(
        master_df, baseline_df, latest_date, display_unit=selected_unit, pipe_df=pipe_df, links_df=links_df
    )
    filename = f"repair_rate_report_{latest_date.date().isoformat()}.pdf"

    return dcc.send_bytes(pdf_bytes, filename)


@callback(
    Output("pipe-pdf-download", "data"),
    Input("download-pipe-pdf-btn", "n_clicks"),
    State("pipe-sheet-dropdown", "value"),
    State("unit-toggle", "value"),
    prevent_initial_call=True,
)
def download_pipe_analysis_pdf(n_clicks, selected_sheet, selected_unit):
    if not n_clicks or not selected_sheet:
        return dash.no_update

    df = load_pipe_repair_details()
    if df.empty:
        return dash.no_update

    links_df = load_project_sheet_links()
    label_map = _pipe_sheet_label_map(links_df)
    sheet_df = df[df["project_sheet"] == selected_sheet].sort_values("pipe_no")
    if sheet_df.empty:
        return dash.no_update

    sheet_label = label_map.get(selected_sheet, selected_sheet)
    pdf_bytes = build_pipe_analysis_pdf_report(sheet_label, sheet_df, display_unit=selected_unit)
    filename = f"pipe_report_{selected_sheet}.pdf".replace(" ", "_")

    return dcc.send_bytes(pdf_bytes, filename)


# ---------------------------------------------------------------------------
# Callback 4: Load / refresh the dashboard
# ---------------------------------------------------------------------------

def _render_dashboard_inner(selected_unit):
    """Returns (main_content, table_content) — split so the callback can
    place the "Latest Day" table in the dedicated Tables section at the
    bottom of the page instead of in the middle of the charts."""
    master_df = load_master_data()

    if master_df.empty:
        empty = _empty_state("No data in the database yet. Upload and import an Excel file first.")
        return empty, None

    baseline_df = load_historical_baselines()
    u = unit_label(selected_unit)

    # --- Top summary cards ---
    latest_date, latest_df = _latest_day_frame(master_df)

    overall_ratio = daily_weighted_repair_ratios(master_df, baseline_df)
    current_overall_ratio = (
        overall_ratio.loc[overall_ratio["date"] == latest_date, "weighted_repair_ratio"].iloc[0]
        if not overall_ratio.empty
        else 0
    )

    summary_cards = html.Div(
        [
            _summary_card("Last Report Date", latest_date.strftime("%d.%m.%Y")),
            _summary_card(
                "Active Project Count",
                str((latest_df["project_status"] == "In Progress").sum()),
            ),
            _summary_card("Coil and Plate Repair Rate", f"{current_overall_ratio * 100:.2f}%"),
        ],
        className="summary-cards",
    )

    # --- Daily repair amount (bar) ---
    daily_amount = repair_amount_trend_data(master_df, display_unit=selected_unit)
    if not daily_amount.empty:
        daily_amount = daily_amount[
            daily_amount["date"] >= daily_amount["date"].max() - pd.Timedelta(days=DAILY_CHARTS_WINDOW_DAYS - 1)
        ]
    bar_fig = px.bar(
        daily_amount,
        x="date",
        y="daily_repair_amount_display",
        labels={"date": "Date", "daily_repair_amount_display": f"Daily Repair Amount ({u})"},
        title=f"Daily Repair Amount ({u})",
    )
    bar_fig.update_traces(
        marker_color=COLOR_COIL,
        texttemplate="%{y:.2f}",
        textposition="outside",
        textfont=dict(size=11),
        hovertemplate="%{x|%d.%m.%Y}<br>Daily Repair Amount: <b>%{y:.2f} " + u + "</b><extra></extra>",
    )
    bar_fig.update_layout(template="plotly_white", margin=dict(l=40, r=20, t=50, b=40), hoverlabel=HOVER_STYLE)
    bar_fig.update_xaxes(tickformat="%d.%m.%y", dtick="D1", tickangle=-45)

    # --- Daily Produced vs. Repaired (clustered bars) + a secondary-axis
    # line for the running "cut but not yet repaired" backlog across all
    # mapped projects — makes it visible at a glance whether the shop is
    # catching up or falling further behind, not just how much happened on
    # any one day. Produced/Repaired bars use the same TREND_WINDOW_DAYS/
    # PIPE_TREND_FLOOR_DATE window as the rest of this dashboard's daily
    # charts; the backlog line needs each pipe's *entire* history (a pipe
    # cut long before the window and still unrepaired still counts toward
    # every window day's backlog), not just the window.
    daily_pipe_fig = None
    daily_pipe_table = None
    pipe_df_for_daily = load_pipe_repair_details()
    if not pipe_df_for_daily.empty:
        pipe_label_map = _pipe_sheet_label_map(load_project_sheet_links())
        mapped_pipe_df = pipe_df_for_daily[pipe_df_for_daily["project_sheet"].isin(pipe_label_map)]
        if not mapped_pipe_df.empty:
            repaired_pipe_df = mapped_pipe_df[mapped_pipe_df["status"] == "Repaired"]
            latest_activity = mapped_pipe_df[["first_seen_date", "repaired_date"]].max().max()
            daily_pipe_window_start = max(
                latest_activity - pd.Timedelta(days=DAILY_CHARTS_WINDOW_DAYS - 1),
                PIPE_TREND_FLOOR_DATE,
            )
            # Nothing on/after PIPE_TREND_FLOOR_DATE yet (e.g. the floor was
            # just moved past the most recent upload to hide a one-time
            # bulk-load lump) -- skip the chart entirely rather than render
            # empty axes; it reappears on its own once a new upload lands
            # past the floor.
            if daily_pipe_window_start <= latest_activity:
                # A categorical axis divides the plot width evenly across
                # however many categories exist — with very few real days
                # (e.g. just today, right after the floor date moves
                # forward), 1-2 categories would each claim a huge share of
                # the width and render as grotesquely fat bars. Padding the
                # display out to a minimum number of days (with blank
                # future dates, zero bars, stock flat-lined at its last
                # known value) keeps bar proportions sane regardless of how
                # little real data exists yet; the padding shrinks to
                # nothing once enough real days accumulate. Padded out to
                # the same DAILY_CHARTS_WINDOW_DAYS as the window itself, so
                # the chart always shows a fixed 14 day-slots.
                display_end = max(
                    latest_activity, daily_pipe_window_start + pd.Timedelta(days=DAILY_CHARTS_WINDOW_DAYS - 1)
                )
                date_range = pd.date_range(daily_pipe_window_start, display_end, freq="D")

                produced_daily = (
                    mapped_pipe_df[mapped_pipe_df["first_seen_date"] >= daily_pipe_window_start]
                    .groupby(mapped_pipe_df["first_seen_date"].dt.normalize())
                    .size()
                    .reindex(date_range, fill_value=0)
                )
                repaired_daily = (
                    repaired_pipe_df[repaired_pipe_df["repaired_date"] >= daily_pipe_window_start]
                    .groupby(repaired_pipe_df["repaired_date"].dt.normalize())
                    .size()
                    .reindex(date_range, fill_value=0)
                )

                # Backlog: +1 the day a pipe is first seen, -1 the day it's
                # repaired (if ever), summed per day over the pipe's
                # *entire* history, then cumulatively summed — so a pipe
                # cut before the display window still correctly counts
                # toward every window day's backlog until (and unless)
                # it's repaired. Forward-fill within the real range so a
                # day with no events keeps the prior day's level instead of
                # dropping to 0.
                opens_all = mapped_pipe_df.groupby(mapped_pipe_df["first_seen_date"].dt.normalize()).size()
                closes_all = repaired_pipe_df.groupby(repaired_pipe_df["repaired_date"].dt.normalize()).size()
                net_all = opens_all.sub(closes_all, fill_value=0)
                full_history_range = pd.date_range(net_all.index.min(), latest_activity, freq="D")
                backlog_series = net_all.reindex(full_history_range, fill_value=0).cumsum()
                real_date_range = pd.date_range(daily_pipe_window_start, latest_activity, freq="D")
                backlog_window_real = backlog_series.reindex(real_date_range, method="ffill")
                # Reindexed onto the (possibly padded) display range
                # *without* ffill — the padding dates beyond latest_activity
                # are blank future days, not known data, so the stock line
                # should stop at the last real point instead of drawing a
                # fake flat continuation across them.
                backlog_window = backlog_window_real.reindex(date_range)

                # A continuous date axis infers bar width/spacing from the
                # gap between neighboring x-values and the overall axis
                # range — with very few days in the window (right after the
                # floor date moves forward, or early in a fresh deployment)
                # there's nothing to infer from, and both the bar width and
                # the grouped-bar offset come out wrong (bars stretch to
                # fill the whole plot and/or overlap each other instead of
                # sitting side by side). A categorical axis with pre-
                # formatted date labels sidesteps this entirely — same
                # approach already used for "Compact Timeline" elsewhere on
                # this dashboard — and grouped bars space correctly via
                # Plotly's ordinary category math regardless of how sparse
                # the dates are.
                x_labels = [d.strftime("%d.%m.%y") for d in date_range]

                daily_pipe_fig = go.Figure()
                daily_pipe_fig.add_trace(
                    go.Bar(
                        x=x_labels,
                        y=produced_daily.values,
                        name="Produced",
                        marker_color=COLOR_SECONDARY,
                        texttemplate="%{y}",
                        textposition="outside",
                        hovertemplate="%{x}<br>Produced: <b>%{y}</b><extra></extra>",
                    )
                )
                daily_pipe_fig.add_trace(
                    go.Bar(
                        x=x_labels,
                        y=repaired_daily.values,
                        name="Repaired",
                        marker_color=COLOR_COIL,
                        texttemplate="%{y}",
                        textposition="outside",
                        hovertemplate="%{x}<br>Repaired: <b>%{y}</b><extra></extra>",
                    )
                )
                # Single shared axis (no secondary y) — bars and the stock
                # line are drawn on the same "Pipe Count" scale, per
                # request. The line is added last so it draws on top of
                # the bars.
                daily_pipe_fig.add_trace(
                    go.Scatter(
                        x=x_labels,
                        y=backlog_window.values,
                        name="Total Stock / Remaining Pipes",
                        mode="lines+markers+text",
                        line=dict(color=COLOR_DANGER, width=3),
                        marker=dict(size=8, color=COLOR_DANGER, line=dict(color="white", width=1)),
                        text=backlog_window.values,
                        texttemplate="%{text}",
                        textposition="top center",
                        textfont=dict(color=COLOR_DANGER),
                        hovertemplate="%{x}<br>Total Stock: <b>%{y}</b> pipes<extra></extra>",
                    )
                )
                daily_pipe_fig.update_layout(
                    barmode="group",
                    title="Backlog Trend",
                    xaxis_title="Date",
                    yaxis=dict(title="Pipe Count", rangemode="tozero"),
                    template="plotly_white",
                    margin=dict(l=40, r=40, t=70, b=40),
                    hoverlabel=HOVER_STYLE,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                )
                daily_pipe_fig.update_xaxes(type="category", tickangle=-45)

                # Table alongside Daily Repair Amount below — same numbers
                # as the chart, scannable without reading bar heights off
                # the axis. Must slice from the real portion only
                # (date_range/x_labels were padded with blank future dates
                # above so the chart doesn't render absurdly fat bars with
                # too few real days — the table has no such rendering
                # constraint, so it should never show those placeholder
                # days). Every real day is included (not truncated) — native
                # pagination (page_size below) splits it 10-per-page, same
                # as every other table in this app; newest day first.
                real_day_count = (latest_activity - daily_pipe_window_start).days + 1
                daily_pipe_table_df = pd.DataFrame(
                    {
                        "date": x_labels[:real_day_count],
                        "produced": produced_daily.values[:real_day_count],
                        "repaired": repaired_daily.values[:real_day_count],
                        "stock": backlog_window.values[:real_day_count],
                    }
                ).iloc[::-1]
                # Net = Produced - Repaired = literal day-over-day change in
                # Stock (matches Stock exactly: negative means the backlog
                # shrank that day, positive means it grew).
                daily_pipe_table_df["net"] = (
                    daily_pipe_table_df["produced"] - daily_pipe_table_df["repaired"]
                )
                daily_pipe_table = dash_table.DataTable(
                    data=daily_pipe_table_df.to_dict("records"),
                    columns=[
                        {"name": "Date", "id": "date"},
                        {"name": "Produced", "id": "produced", "type": "numeric"},
                        {"name": "Repaired", "id": "repaired", "type": "numeric"},
                        {"name": "Net", "id": "net", "type": "numeric", "format": Format(sign=Sign.positive)},
                        {"name": "Stock", "id": "stock", "type": "numeric"},
                    ],
                    page_size=10,
                    sort_action="native",
                    style_table={"overflowX": "auto"},
                    style_cell=TABLE_CELL_STYLE,
                    style_header=TABLE_HEADER_STYLE,
                    style_data_conditional=TABLE_CONDITIONAL_STYLE
                    + [
                        {"if": {"column_id": "produced"}, "color": COLOR_SECONDARY, "fontWeight": "600"},
                        {"if": {"column_id": "repaired"}, "color": COLOR_COIL, "fontWeight": "600"},
                        {"if": {"column_id": "stock"}, "fontWeight": "700"},
                        # Net < 0 (Produced fewer than Repaired) shrinks the
                        # backlog -- good. Net > 0 grows it -- flag it.
                        {"if": {"column_id": "net", "filter_query": "{net} < 0"}, "color": "#16a34a"},
                        {"if": {"column_id": "net", "filter_query": "{net} > 0"}, "color": COLOR_DANGER},
                    ],
                )

    # --- Worst-performing projects (latest day) ---
    worst = latest_df.sort_values("repair_ratio", ascending=False).head(10)
    worst_fig = px.bar(
        worst,
        x="repair_ratio",
        y="project_label_clean",
        orientation="h",
        labels={"repair_ratio": "Repair Ratio", "project_label_clean": "Project"},
        title="Top 10 Projects by Repair Ratio (Latest Day)",
    )
    worst_fig.update_traces(
        marker_color=COLOR_COIL,
        texttemplate="%{x:.2%}",
        textposition="outside",
        textfont=dict(size=11),
        hovertemplate="%{y}<br>Repair Ratio: <b>%{x:.2%}</b><extra></extra>",
    )
    worst_fig.update_layout(
        template="plotly_white",
        yaxis=dict(autorange="reversed"),
        margin=dict(l=120, r=20, t=50, b=40),
        hoverlabel=HOVER_STYLE,
    )
    worst_fig.update_xaxes(tickformat=".2%")
    _add_bar_value_annotations(
        worst_fig, worst["project_label_clean"], worst["repair_ratio"], worst["qty"], orientation="h"
    )

    # --- Skelp impact: how much "incl. skelp" adds on top of the base ratio ---
    skelp_df = latest_df[["project_label", "project_label_clean", "qty", "repair_ratio", "repair_ratio_incl_skelp"]].copy()
    skelp_df["skelp_impact"] = (skelp_df["repair_ratio_incl_skelp"] - skelp_df["repair_ratio"]).round(6)
    skelp_top = skelp_df.sort_values("repair_ratio", ascending=False).head(5)
    # A stacked segment that's small relative to the tallest bar can't fit
    # its two-line "value / Qty N" label without overlapping the segment
    # next to it — below a min-fraction-of-max threshold (same heuristic as
    # _add_bar_value_annotations / the PDF version of this chart) the label
    # moves outside the bar (an annotation above the stack) instead of
    # being drawn inside it.
    ratio_max_stack = (skelp_top["repair_ratio"] + skelp_top["skelp_impact"]).max() if len(skelp_top) else 0
    skelp_base_text = [
        f"{r:.2%}<br>Qty {q}" if ratio_max_stack and r >= ratio_max_stack * 0.1 else ""
        for r, q in zip(skelp_top["repair_ratio"], skelp_top["qty"])
    ]
    skelp_impact_text = [
        f"+{imp:.2%}" if ratio_max_stack and imp >= ratio_max_stack * 0.08 else ""
        for imp in skelp_top["skelp_impact"]
    ]
    skelp_fig = go.Figure()
    skelp_fig.add_trace(
        go.Bar(
            x=skelp_top["project_label_clean"],
            y=skelp_top["repair_ratio"],
            name="Repair Ratio",
            marker_color=COLOR_COIL,
            text=skelp_base_text,
            textposition="inside",
            textfont=dict(size=11, color="white"),
            hovertemplate="%{x}<br>Repair Ratio: <b>%{y:.2%}</b><extra></extra>",
        )
    )
    skelp_fig.add_trace(
        go.Bar(
            x=skelp_top["project_label_clean"],
            y=skelp_top["skelp_impact"],
            name="Skelp Impact",
            marker_color=COLOR_SECONDARY,
            text=skelp_impact_text,
            textposition="inside",
            textfont=dict(size=11, color="white"),
            hovertemplate="%{x}<br>Skelp Impact: <b>+%{y:.2%}</b><extra></extra>",
        )
    )
    for label, r, imp, q, base_t, impact_t in zip(
        skelp_top["project_label_clean"],
        skelp_top["repair_ratio"],
        skelp_top["skelp_impact"],
        skelp_top["qty"],
        skelp_base_text,
        skelp_impact_text,
    ):
        # Impact line first, repair-ratio line last, so the text block reads
        # top-to-bottom in the same order the bar stacks bottom-to-top:
        # Repair Ratio (base, bottom of stack) ends up closest to the bar.
        lines = []
        if not impact_t:
            lines.append(f"+{imp:.2%}")
        if not base_t:
            lines.append(f"{r:.2%} (Qty {q})")
        if lines:
            skelp_fig.add_annotation(
                x=label,
                y=r + imp,
                text="<br>".join(lines),
                showarrow=False,
                yshift=14,
                font=dict(size=9, color="#1e293b"),
            )
    skelp_fig.update_layout(
        barmode="stack",
        title="Skelp-End Weld Impact on Repair Ratio (Top 5, Latest Day)",
        yaxis_title="Repair Ratio",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=80),
        hoverlabel=HOVER_STYLE,
        xaxis=dict(categoryorder="array", categoryarray=skelp_top["project_label_clean"].tolist()),
    )
    skelp_fig.update_xaxes(tickangle=-25)
    skelp_fig.update_yaxes(tickformat=".2%")

    # --- Skelp impact: same as above, but on the repair amount itself
    # rather than the ratio, and ranked by its own metric (total_repair_amount)
    # instead of reusing the ratio chart's top 5 — the two charts can show
    # different projects.
    skelp_amount_df = latest_df[
        ["project_label_clean", "qty", "total_repair_amount", "total_repair_amount_incl_skelp"]
    ].copy()
    skelp_amount_df["total_repair_amount"] = amount_in_display_unit(skelp_amount_df["total_repair_amount"], selected_unit)
    skelp_amount_df["total_repair_amount_incl_skelp"] = amount_in_display_unit(
        skelp_amount_df["total_repair_amount_incl_skelp"], selected_unit
    )
    skelp_amount_df["skelp_impact_amount"] = (
        skelp_amount_df["total_repair_amount_incl_skelp"] - skelp_amount_df["total_repair_amount"]
    ).round(4)
    skelp_amount_top = skelp_amount_df.sort_values("total_repair_amount", ascending=False).head(5)
    skelp_amount_top["pct_increase"] = (
        skelp_amount_top["skelp_impact_amount"]
        / skelp_amount_top["total_repair_amount"].where(skelp_amount_top["total_repair_amount"] != 0)
    ).fillna(0) * 100
    # Same "move the label outside the bar if the segment's too small to
    # fit it" rule as the ratio chart above.
    amount_max_stack = (
        (skelp_amount_top["total_repair_amount"] + skelp_amount_top["skelp_impact_amount"]).max()
        if len(skelp_amount_top)
        else 0
    )
    amount_base_text = [
        f"{v:.2f}<br>Qty {q}" if amount_max_stack and v >= amount_max_stack * 0.1 else ""
        for v, q in zip(skelp_amount_top["total_repair_amount"], skelp_amount_top["qty"])
    ]
    amount_impact_text = [
        f"+{imp:.2f}" if amount_max_stack and imp >= amount_max_stack * 0.08 else ""
        for imp in skelp_amount_top["skelp_impact_amount"]
    ]
    skelp_amount_fig = go.Figure()
    skelp_amount_fig.add_trace(
        go.Bar(
            x=skelp_amount_top["project_label_clean"],
            y=skelp_amount_top["total_repair_amount"],
            name=f"Repair Amount ({u})",
            marker_color=COLOR_COIL,
            text=amount_base_text,
            textposition="inside",
            textfont=dict(size=11, color="white"),
            hovertemplate="%{x}<br>Repair Amount: <b>%{y:.2f} " + u + "</b><extra></extra>",
        )
    )
    skelp_amount_fig.add_trace(
        go.Bar(
            x=skelp_amount_top["project_label_clean"],
            y=skelp_amount_top["skelp_impact_amount"],
            name=f"Skelp Impact ({u})",
            marker_color=COLOR_SECONDARY,
            customdata=skelp_amount_top["pct_increase"],
            text=amount_impact_text,
            textposition="inside",
            textfont=dict(size=11, color="white"),
            hovertemplate="%{x}<br>Skelp Impact: <b>+%{y:.2f} " + u + "</b> (+%{customdata:.2f}%)<extra></extra>",
        )
    )
    for label, v, imp, q, base_t, impact_t in zip(
        skelp_amount_top["project_label_clean"],
        skelp_amount_top["total_repair_amount"],
        skelp_amount_top["skelp_impact_amount"],
        skelp_amount_top["qty"],
        amount_base_text,
        amount_impact_text,
    ):
        # Impact line first, repair-amount line last — see skelp_fig above.
        lines = []
        if not impact_t:
            lines.append(f"+{imp:.2f}")
        if not base_t:
            lines.append(f"{v:.2f} (Qty {q})")
        if lines:
            skelp_amount_fig.add_annotation(
                x=label,
                y=v + imp,
                text="<br>".join(lines),
                showarrow=False,
                yshift=14,
                font=dict(size=9, color="#1e293b"),
            )
    skelp_amount_fig.update_layout(
        barmode="stack",
        title=f"Skelp-End Weld Impact on Repair Amount ({u}, Top 5, Latest Day)",
        yaxis_title=f"Repair Amount ({u})",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=80),
        hoverlabel=HOVER_STYLE,
        xaxis=dict(categoryorder="array", categoryarray=skelp_amount_top["project_label_clean"].tolist()),
    )
    skelp_amount_fig.update_xaxes(tickangle=-25)

    # --- Repair amount pareto: which projects drive most of the total ---
    pareto_df = latest_df[["project_label_clean", "qty", "total_repair_amount"]].sort_values(
        "total_repair_amount", ascending=False
    ).reset_index(drop=True)
    pareto_df["total_repair_amount"] = amount_in_display_unit(pareto_df["total_repair_amount"], selected_unit)
    pareto_df["cumulative_pct"] = (
        pareto_df["total_repair_amount"].cumsum() / pareto_df["total_repair_amount"].sum() * 100
    )
    pareto_fig = go.Figure()
    pareto_fig.add_trace(
        go.Bar(
            x=pareto_df["project_label_clean"],
            y=pareto_df["total_repair_amount"],
            name=f"Repair Amount ({u})",
            marker_color=COLOR_COIL,
            texttemplate="%{y:.2f}",
            textposition="outside",
            textfont=dict(size=10),
            hovertemplate="%{x}<br>Repair Amount: <b>%{y:.2f} " + u + "</b><extra></extra>",
        )
    )
    pareto_fig.add_trace(
        go.Scatter(
            x=pareto_df["project_label_clean"],
            y=pareto_df["cumulative_pct"],
            name="Cumulative %",
            yaxis="y2",
            line=dict(color=COLOR_SECONDARY, width=2),
            mode="lines+markers+text",
            text=pareto_df["cumulative_pct"].map("{:.2f}%".format),
            textposition="top center",
            textfont=dict(size=10, color=COLOR_SECONDARY),
            hovertemplate="%{x}<br>Cumulative: <b>%{y:.2f}%</b><extra></extra>",
        )
    )
    pareto_fig.update_layout(
        title=f"Repair Amount Pareto — Cumulative Contribution ({u}, Latest Day)",
        template="plotly_white",
        yaxis=dict(title=f"Repair Amount ({u})"),
        yaxis2=dict(title="Cumulative %", overlaying="y", side="right", range=[0, 105]),
        margin=dict(l=40, r=40, t=50, b=80),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hoverlabel=HOVER_STYLE,
        # Duplicate project numbers (same project, different dimensions)
        # mean the x labels aren't unique — pin the category order explicitly
        # so the bar and the cumulative line can't be reshuffled onto the
        # same tick and zigzag.
        xaxis=dict(categoryorder="array", categoryarray=pareto_df["project_label_clean"].tolist()),
    )
    pareto_fig.update_xaxes(tickangle=-45)
    _add_bar_value_annotations(
        pareto_fig, pareto_df["project_label_clean"], pareto_df["total_repair_amount"], pareto_df["qty"]
    )

    # --- Repair ratio pareto: same idea, but ranked by repair ratio instead
    # of absolute amount, so a small pipe with a very high ratio still shows
    # up as a top contributor even if its absolute meters are small ---
    ratio_pareto_df = latest_df[["project_label_clean", "qty", "repair_ratio"]].sort_values(
        "repair_ratio", ascending=False
    ).reset_index(drop=True)
    ratio_pareto_df["cumulative_pct"] = (
        ratio_pareto_df["repair_ratio"].cumsum() / ratio_pareto_df["repair_ratio"].sum() * 100
    )
    ratio_pareto_fig = go.Figure()
    ratio_pareto_fig.add_trace(
        go.Bar(
            x=ratio_pareto_df["project_label_clean"],
            y=ratio_pareto_df["repair_ratio"],
            name="Repair Ratio",
            marker_color=COLOR_COIL,
            texttemplate="%{y:.2%}",
            textposition="outside",
            textfont=dict(size=10),
            hovertemplate="%{x}<br>Repair Ratio: <b>%{y:.2%}</b><extra></extra>",
        )
    )
    ratio_pareto_fig.add_trace(
        go.Scatter(
            x=ratio_pareto_df["project_label_clean"],
            y=ratio_pareto_df["cumulative_pct"],
            name="Cumulative %",
            yaxis="y2",
            line=dict(color=COLOR_SECONDARY, width=2),
            mode="lines+markers+text",
            text=ratio_pareto_df["cumulative_pct"].map("{:.2f}%".format),
            textposition="top center",
            textfont=dict(size=10, color=COLOR_SECONDARY),
            hovertemplate="%{x}<br>Cumulative: <b>%{y:.2f}%</b><extra></extra>",
        )
    )
    ratio_pareto_fig.update_layout(
        title="Repair Ratio Pareto — Cumulative Contribution (Latest Day)",
        template="plotly_white",
        yaxis=dict(title="Repair Ratio", tickformat=".2%"),
        yaxis2=dict(title="Cumulative %", overlaying="y", side="right", range=[0, 105]),
        margin=dict(l=40, r=40, t=50, b=80),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hoverlabel=HOVER_STYLE,
        xaxis=dict(categoryorder="array", categoryarray=ratio_pareto_df["project_label_clean"].tolist()),
    )
    ratio_pareto_fig.update_xaxes(tickangle=-45)
    _add_bar_value_annotations(
        ratio_pareto_fig, ratio_pareto_df["project_label_clean"], ratio_pareto_df["repair_ratio"], ratio_pareto_df["qty"]
    )

    # Production quality/volume matrix (bubble plot) now lives in its own
    # callback (render_bubble_chart) with a view switcher — see below.

    # --- Latest day detail table ---
    table_columns = [
        "project_no",
        "dimensions",
        "production_type",
        "qty",
        "project_status",
        "repair_ratio",
    ]
    detail_table = dash_table.DataTable(
        data=_table_records(latest_df, table_columns),
        columns=_table_columns(table_columns),
        page_size=15,
        sort_action="native",
        filter_action="native",
        style_table={"overflowX": "auto"},
        style_cell=TABLE_CELL_STYLE,
        style_header=TABLE_HEADER_STYLE,
        style_data_conditional=TABLE_CONDITIONAL_STYLE,
    )

    return html.Div(
        [
            summary_cards,
            html.Div(
                [
                    html.Div(
                        [
                            dcc.Checklist(
                                id="overall-trend-checklist",
                                options=[
                                    {"label": "Excl. Skelp", "value": "Excl"},
                                    {"label": "Incl. Skelp", "value": "Incl"},
                                ],
                                value=["Excl", "Incl"],
                                inline=True,
                            ),
                            dcc.Loading(html.Div(id="overall-trend-graph-content")),
                        ],
                        className="chart-half",
                    ),
                    html.Div(
                        [
                            dcc.Checklist(
                                id="type-trend-checklist",
                                options=[
                                    {"label": "Coil", "value": "Coil"},
                                    {"label": "Plate", "value": "Plate"},
                                    {"label": "Mix (Coil + Plate)", "value": "Mix"},
                                ],
                                value=["Coil", "Mix"],
                                inline=True,
                            ),
                            dcc.Loading(html.Div(id="type-trend-graph-content")),
                        ],
                        className="chart-half",
                    ),
                ],
                className="chart-row chart-row-stacked",
            ),
            dcc.Graph(figure=worst_fig),
            html.P("The number inside each bar is the pipe quantity (Qty) for that project.", className="help-text"),
            html.Div(
                [
                    dcc.Graph(figure=skelp_fig, className="chart-half"),
                    dcc.Graph(figure=skelp_amount_fig, className="chart-half"),
                ],
                className="chart-row",
            ),
            dcc.Graph(id="pareto-ratio-graph", figure=ratio_pareto_fig),
            dcc.Graph(id="pareto-amount-graph", figure=pareto_fig),
            html.Div(id="pareto-charts-hover-sync", style={"display": "none"}),
            html.Div(
                [
                    dcc.RadioItems(
                        id="bubble-view-toggle",
                        options=[
                            {"label": "By Project — Spiral Length", "value": "project_spiral"},
                            {"label": "By Project — Pipe Length", "value": "project_pipe"},
                        ],
                        value="project_spiral",
                        inline=True,
                    ),
                    dcc.Loading(html.Div(id="bubble-chart-content")),
                ],
                className="bubble-view-row",
            ),
            *(
                [dcc.Graph(id="daily-pipe-count-graph", figure=daily_pipe_fig)]
                if daily_pipe_fig is not None
                else []
            ),
            html.Div(
                [dcc.Graph(id="daily-amount-graph", figure=bar_fig, className="chart-half")]
                + ([html.Div(daily_pipe_table, className="chart-half")] if daily_pipe_table is not None else []),
                className="chart-row",
            ),
            # Dummy target for the clientside hover-sync callback below —
            # Plotly.restyle does the actual work directly on the graph divs
            # via native plotly_hover/plotly_unhover listeners; this div's
            # own content is never used.
            html.Div(id="daily-charts-hover-sync", style={"display": "none"}),
        ]
    ), html.Div([html.H3("Latest Day — Project Details"), detail_table])


@callback(
    Output("dashboard-content", "children"),
    Output("dashboard-table-content", "children"),
    Output("dashboard-stage-1", "data"),
    Input("refresh-dashboard-btn", "n_clicks"),
    Input("import-confirm-result", "children"),  # auto-refresh after a successful import
    Input("unit-toggle", "value"),
    Input("dashboard-tabs", "value"),
)
def render_dashboard(_n_clicks, _import_result, selected_unit, active_tab):
    # Only the Dashboard tab needs this (heavy: several charts + a full
    # master_df load) — skip the work entirely while another tab is active,
    # instead of rebuilding content nobody can see on every unit-toggle/
    # refresh/import. Fires again once active_tab switches back here.
    if active_tab != "tab-dashboard":
        return dash.no_update, dash.no_update, dash.no_update
    # dashboard-stage-1 is what pipe-overview-content / the sub-chart
    # callbacks below chain off of (a lightweight signal, not this div's own
    # heavy children) — always write a fresh value here, even on error, or
    # a stuck signal would leave every downstream spinner spinning forever.
    try:
        content, table_content = _render_dashboard_inner(selected_unit)
    except Exception:
        logging.exception("render_dashboard failed")
        content = _empty_state("Something went wrong loading the dashboard. Try refreshing.")
        table_content = None
    return content, table_content, time.time()


# Hovering a bar in either of the two daily charts lightens the matching
# date's bar in the other (see assets/daily_charts_hover_sync.js). Wired to
# native Plotly events directly rather than Dash's hoverData prop — that
# prop didn't reliably reset to null on mouse-leave, which left the
# highlight stuck. Re-fires (idempotently) whenever the Dashboard tab's
# content is rebuilt, since that's when the graph divs are (re)created.
clientside_callback(
    "window.dash_clientside.clientside.wireDailyChartsSync",
    Output("daily-charts-hover-sync", "children"),
    Input("dashboard-content", "children"),
)

# Same idea, for the two Pareto charts (bar trace only — the cumulative-%
# line trace is left alone). Matched by exact project label, not date
# truncation, since the two Paretos are sorted by different metrics and so
# don't share the same x-axis order.
clientside_callback(
    "window.dash_clientside.clientside.wireParetoChartsSync",
    Output("pareto-charts-hover-sync", "children"),
    Input("dashboard-content", "children"),
)


# ---------------------------------------------------------------------------
# Callback: Pipe Activity Overview (Dashboard tab) — cross-project "what's
# newest" pipe list. The daily pipe-count chart itself lives in
# render_dashboard now, grouped next to Daily Repair Amount. Kept separate
# from the single-project Pipe Analysis tab, which only ever shows one
# project's own charts at a time (see render_pipe_analysis).
# ---------------------------------------------------------------------------

def _render_pipe_overview_inner(selected_unit):
    u = unit_label(selected_unit)
    df = load_pipe_repair_details()
    if df.empty:
        return _empty_state("No pipe-level data in the database yet. Import an Excel file first.")

    links_df = load_project_sheet_links()
    label_map = _pipe_sheet_label_map(links_df)
    mapped_df = df[df["project_sheet"].isin(label_map)]
    if mapped_df.empty:
        return _empty_state(
            "No mapped projects yet. Check out the "
            "feature/project-sheet-mapping-tool branch to link sheets first."
        )

    # Two independent "what's newest, right now" views — no date picker here
    # (that's what Project Trend and the per-project pages are for).
    # first_seen_date freezes at each pipe's original production/cut date, so
    # filtering on it surfaces fresh production activity; repaired_date is
    # the separate field for fresh repair activity, and the two dates often
    # differ (a pipe can be cut today and repaired days later, or repaired
    # today after being cut earlier) — hence two tables, not one.
    confirmed_links = links_df[links_df["status"] == "confirmed"][["project_sheet", "project_no", "dimensions"]]
    newest_columns = ["project_no", "dimensions", "pipe_no", "repair_amount", "repair_ratio"]

    def _newest_table(source_df: pd.DataFrame, title: str) -> html.Div:
        table_df = source_df.merge(confirmed_links, on="project_sheet", how="left").copy()
        if "repair_amount" in table_df.columns:
            table_df["repair_amount"] = amount_in_display_unit(table_df["repair_amount"], selected_unit)
        table_df = table_df[newest_columns].sort_values(["project_no", "pipe_no"])
        return html.Div(
            [
                html.H3(title),
                dash_table.DataTable(
                    data=_table_records(table_df, newest_columns),
                    columns=_table_columns(
                        newest_columns,
                        label_overrides={
                            "project_no": "Project",
                            "pipe_no": "Pipe No.",
                            "repair_amount": f"Repair Amount ({u})",
                        },
                    ),
                    page_size=15,
                    sort_action="native",
                    filter_action="native",
                    style_table={"overflowX": "auto"},
                    style_cell=TABLE_CELL_STYLE,
                    style_header=TABLE_HEADER_STYLE,
                    style_data_conditional=TABLE_CONDITIONAL_STYLE,
                ),
            ],
            className="chart-half",
        )

    tables = []

    produced_latest_date = mapped_df["first_seen_date"].max()
    produced_df = mapped_df[mapped_df["first_seen_date"] == produced_latest_date]
    tables.append(
        _newest_table(
            produced_df,
            f"Newest Produced Pipes — {pd.Timestamp(produced_latest_date).strftime('%d.%m.%Y')}",
        )
    )

    repaired_only = mapped_df[mapped_df["repaired_date"].notna()]
    if not repaired_only.empty:
        repaired_latest_date = repaired_only["repaired_date"].max()
        repaired_df = repaired_only[repaired_only["repaired_date"] == repaired_latest_date]
        tables.append(
            _newest_table(
                repaired_df,
                f"Newest Repaired Pipes — {pd.Timestamp(repaired_latest_date).strftime('%d.%m.%Y')}",
            )
        )

    return html.Div([html.Div(tables, className="chart-row")])


@callback(
    Output("pipe-overview-content", "children"),
    Output("dashboard-stage-2", "data"),
    Input("dashboard-stage-1", "data"),  # chained after dashboard-content, not parallel to it
    State("unit-toggle", "value"),
    prevent_initial_call=True,
)
def render_pipe_overview(_stage_1, selected_unit):
    try:
        content = _render_pipe_overview_inner(selected_unit)
    except Exception:
        logging.exception("render_pipe_overview failed")
        content = _empty_state("Something went wrong loading pipe activity. Try refreshing.")
    return content, time.time()


def _summary_card(label: str, value: str) -> html.Div:
    return html.Div(
        [html.Div(value, className="summary-value"), html.Div(label, className="summary-label")],
        className="summary-card",
    )


def _build_bubble_fig(
    df: pd.DataFrame,
    x_col: str,
    x_label: str,
    title: str,
    unit: str,
    label_col: str = "project_label_clean",
) -> go.Figure:
    """Bubble plot: X = ``x_col``, Y = repair_ratio_pct, size = repair
    amount, color = repair_ratio_pct on a yellow-to-red scale. ``df``'s
    ``x_col``/``total_repair_amount`` are expected to already be converted
    to ``unit`` by the caller — this only formats labels/hover text.
    ``label_col`` is "project_label_clean" (no "20XXQ-" prefix) for the
    by-project views, or "dimensions" for the by-dimension view. The worst
    (highest-ratio)
    ``BUBBLE_LABEL_TOP_N`` points also get their name on-chart, not just
    hover — the rest just show the ratio, to keep the chart readable."""
    worst_idx = set(df.nlargest(BUBBLE_LABEL_TOP_N, "repair_ratio_pct").index)
    ratio_text = df["repair_ratio_pct"].map("{:.2f}%".format)
    text_values = (df[label_col].astype(str) + "<br>" + ratio_text).where(df.index.isin(worst_idx), ratio_text)

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=df[x_col],
            y=df["repair_ratio_pct"],
            mode="markers+text",
            text=text_values,
            textposition="top center",
            textfont=dict(size=10),
            customdata=df[label_col],
            marker=dict(
                size=df["total_repair_amount"],
                sizemode="area",
                sizeref=2.0 * df["total_repair_amount"].max() / (40.0**2),
                sizemin=4,
                color=df["repair_ratio_pct"],
                colorscale="YlOrRd",
                showscale=True,
                colorbar=dict(title="Repair Ratio (%)"),
                opacity=0.55,
                line=dict(color="black", width=1.5),
            ),
            hovertemplate=(
                "%{customdata}<br>"
                f"{x_label}: <b>%{{x:.2f}} {unit}</b><br>"
                "Spiral Repair Ratio: <b>%{y:.2f}%</b><br>"
                f"Total Repair Amount: <b>%{{marker.size:.2f}} {unit}</b>"
                "<extra></extra>"
            ),
        )
    )
    fig.update_layout(
        title=title,
        xaxis_title=x_label,
        yaxis_title="Spiral Repair Ratio (%)",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=40),
        hoverlabel=HOVER_STYLE,
    )
    return fig


def _build_box_fig(
    df: pd.DataFrame, x_col: str, x_label: str, title: str, category_order: list[str] | None = None
) -> go.Figure:
    """Box plot of repair_ratio_pct grouped by x_col — shows the spread
    (median, quartiles, outliers) within each group instead of collapsing
    it to a single average, so a group that's "fine on average but wildly
    inconsistent" doesn't look the same as one that's genuinely uniform.
    Kept deliberately plain: only genuine outliers are drawn as points (not
    every pipe — that turns into a wall of dots), and a dashed mean line
    gives a second, more familiar reference next to the median."""
    # One uniform color for every box (categories are already distinguished
    # by their x-axis position/label) — matches how every other multi-
    # category chart on this dashboard (Top 15, Pareto) uses a single color
    # rather than a rainbow per bar. Outlier points use the dashboard's
    # existing orange accent, its usual "this is the noteworthy part" color.
    # hoveron="points" is the key fix for a real Plotly box-plot problem:
    # by default hovering the box itself pops up a *separate* tooltip for
    # each of its 7 internal stats (min/q1/median/mean/q3/max/fences),
    # stacked on top of each other and burying the chart. Restricting hover
    # to just the outlier dots (with the short template below) keeps the
    # box itself simple to look at without a wall of redundant tooltips.
    fig = px.box(df, x=x_col, y="repair_ratio_pct", points="outliers")
    fig.update_traces(
        boxmean=True,
        fillcolor="rgba(37, 99, 235, 0.15)",
        line=dict(color=COLOR_COIL),
        marker=dict(size=6, color=COLOR_SECONDARY, opacity=0.85),
        width=0.4,
        hoveron="points",
        hovertemplate="%{x}<br>Repair Ratio: <b>%{y:.2f}%</b><extra></extra>",
    )
    fig.update_layout(
        title=title,
        xaxis_title=x_label,
        yaxis_title="Repair Ratio (%)",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=80),
        hoverlabel=HOVER_STYLE,
        showlegend=False,
    )
    fig.update_xaxes(tickangle=-25)
    if category_order is not None:
        fig.update_xaxes(categoryorder="array", categoryarray=category_order)
    return fig


def _box_plot_section(df: pd.DataFrame, x_col: str, x_label: str, title: str) -> html.Div:
    """Box plot plus a plain-language stat line underneath, so the takeaway
    doesn't depend on knowing how to read a box-and-whisker shape. Width is
    capped so a chart with only a handful of categories doesn't stretch
    edge-to-edge and look mostly empty."""
    if df.empty:
        # e.g. a project whose pipes are all still Awaiting Repair — no
        # repair_ratio values exist yet at all. An empty px.box still
        # renders (blank axes), but the caption's median()/min()/max() on
        # an empty Series would show "nan%" three times, which reads as
        # broken rather than "no data yet."
        return html.Div(
            html.P("No repaired pipes yet for this selection.", className="help-text"),
            style={"maxWidth": "700px"},
        )
    fig = _build_box_fig(df, x_col, x_label, title)
    ratios = df["repair_ratio_pct"]
    caption = html.P(
        f"Median: {ratios.median():.2f}%  •  Lowest: {ratios.min():.2f}%  •  Highest: {ratios.max():.2f}%"
        " — the box covers where most values fall, the dashed line is the average, "
        "orange dots are the individual exceptions.",
        className="help-text",
    )
    return html.Div([dcc.Graph(figure=fig), caption], style={"maxWidth": "700px"})


# ---------------------------------------------------------------------------
# Callback: Production quality/volume matrix (bubble plot) — a view switcher
# picks between two per-project variants (pipe length / spiral length on the
# X axis) and one per-dimension variant, all sharing the same builder above.
# ---------------------------------------------------------------------------

@callback(
    Output("bubble-chart-content", "children"),
    Input("bubble-view-toggle", "value"),
    Input("dashboard-stage-1", "data"),  # chained after dashboard-content, not parallel to it
    State("unit-toggle", "value"),
    prevent_initial_call=True,
)
def render_bubble_chart(selected_view, _stage_1, selected_unit):
    master_df = load_master_data()
    if master_df.empty:
        return _empty_state("No data available. Import an Excel file first.")

    _latest_date, latest_df = _latest_day_frame(master_df)
    u = unit_label(selected_unit)

    bubble_df = latest_df[
        [
            "project_label_clean",
            "project_total_pipe_length",
            "repaired_spiral_length",
            "repair_ratio",
            "total_repair_amount",
        ]
    ].copy()
    bubble_df["repair_ratio_pct"] = (bubble_df["repair_ratio"] * 100).round(4)
    bubble_df["project_total_pipe_length"] = length_in_display_unit(bubble_df["project_total_pipe_length"], selected_unit)
    bubble_df["repaired_spiral_length"] = length_in_display_unit(bubble_df["repaired_spiral_length"], selected_unit)
    bubble_df["total_repair_amount"] = amount_in_display_unit(bubble_df["total_repair_amount"], selected_unit)

    if selected_view == "project_pipe":
        fig = _build_bubble_fig(
            bubble_df,
            x_col="project_total_pipe_length",
            x_label=f"Total Production Pipe Length ({u})",
            title=f"Repair Ratio vs. Production Volume by Project ({u}, Pipe Length, Latest Day)",
            unit=u,
        )
    else:
        fig = _build_bubble_fig(
            bubble_df,
            x_col="repaired_spiral_length",
            x_label=f"Total Spiral Length ({u})",
            title=f"Repair Ratio vs. Production Volume by Project ({u}, Spiral Length, Latest Day)",
            unit=u,
        )

    return dcc.Graph(figure=fig)



def _build_trend_trace(series_df: pd.DataFrame, value_col: str = "weighted_repair_ratio"):
    """Straight-line (least-squares) trend over at most the last
    ``TREND_WINDOW_DAYS`` days of ``series_df`` (date, ``value_col``) — the
    same window the chart itself is already limited to by the time this is
    called, so this is mostly a safety net for callers that pass unwindowed
    data."""
    windowed = series_df.dropna(subset=["date", value_col]).sort_values("date")
    if windowed.empty:
        return None
    window_start = windowed["date"].max() - pd.Timedelta(days=TREND_WINDOW_DAYS)
    windowed = windowed[windowed["date"] >= window_start]
    if len(windowed) < 2:
        return None

    day_offsets = (windowed["date"] - windowed["date"].min()).dt.days
    slope, intercept = np.polyfit(day_offsets, windowed[value_col] * 100, 1)
    trend_y = (slope * day_offsets + intercept).round(4)
    # Guard against floating-point noise (e.g. slope ~1e-16 on genuinely flat
    # data) being reported as "Rising"/"Falling".
    if slope > 1e-6:
        direction = "Rising"
    elif slope < -1e-6:
        direction = "Falling"
    else:
        direction = "Flat"
    return go.Scatter(
        x=windowed["date"],
        y=trend_y,
        mode="lines",
        name=f"Trend ({direction})",
        line=dict(width=2, color=COLOR_MUTED, dash="dot"),
        hovertemplate="%{x|%d.%m.%Y}<br>Trend: <b>%{y:.2f}%</b><extra></extra>",
    )


# ---------------------------------------------------------------------------
# Callback: Coil and Plate Repair Rate Trend — same idea as the production-type
# trend below: a checklist isolates Excl./Incl. Skelp, and a linear trend
# line is drawn (last TREND_WINDOW_DAYS days, or the full range if shorter)
# when only one of the two is selected.
# ---------------------------------------------------------------------------

@callback(
    Output("overall-trend-graph-content", "children"),
    Input("overall-trend-checklist", "value"),
    Input("dashboard-stage-1", "data"),  # chained after dashboard-content, not parallel to it
    prevent_initial_call=True,
)
def render_overall_trend_chart(selected_series, _stage_1):
    master_df = load_master_data()
    if master_df.empty:
        return _empty_state("No data available. Import an Excel file first.")
    if not selected_series:
        return _empty_state("Select at least one series to display.")

    baseline_df = load_historical_baselines()
    overall_ratio = daily_weighted_repair_ratios(master_df, baseline_df).copy()
    overall_ratio["weighted_repair_ratio"] = overall_ratio["weighted_repair_ratio"].round(4)
    overall_ratio["weighted_repair_ratio_incl_skelp"] = overall_ratio["weighted_repair_ratio_incl_skelp"].round(4)
    window_start = overall_ratio["date"].max() - pd.Timedelta(days=TREND_WINDOW_DAYS)
    overall_ratio = overall_ratio[overall_ratio["date"] >= window_start]

    fig = go.Figure()
    if "Excl" in selected_series:
        fig.add_trace(
            go.Scatter(
                x=overall_ratio["date"],
                y=overall_ratio["weighted_repair_ratio"] * 100,
                mode="lines+markers+text",
                name="Excl. Skelp",
                line=dict(color=COLOR_COIL, width=3),
                text=(overall_ratio["weighted_repair_ratio"] * 100).map("{:.2f}%".format),
                textposition="top center",
                textfont=dict(size=10, color=COLOR_COIL),
                hovertemplate="%{x|%d.%m.%Y}<br>Excl. Skelp: <b>%{y:.2f}%</b><extra></extra>",
            )
        )
    if "Incl" in selected_series:
        fig.add_trace(
            go.Scatter(
                x=overall_ratio["date"],
                y=overall_ratio["weighted_repair_ratio_incl_skelp"] * 100,
                mode="lines+markers+text",
                name="Incl. Skelp",
                line=dict(color=COLOR_SECONDARY, width=3, dash="dash"),
                text=(overall_ratio["weighted_repair_ratio_incl_skelp"] * 100).map("{:.2f}%".format),
                textposition="bottom center",
                textfont=dict(size=10, color=COLOR_SECONDARY),
                hovertemplate="%{x|%d.%m.%Y}<br>Incl. Skelp: <b>%{y:.2f}%</b><extra></extra>",
            )
        )

    # Same rule as the production-type trend: only draw a trend line when a
    # single series is isolated.
    if len(selected_series) == 1:
        value_col = "weighted_repair_ratio" if selected_series[0] == "Excl" else "weighted_repair_ratio_incl_skelp"
        trend_trace = _build_trend_trace(overall_ratio, value_col)
        if trend_trace is not None:
            fig.add_trace(trend_trace)

    if len(selected_series) == 2:
        title = "Coil and Plate Repair Rate Trend (Excl. vs Incl. Skelp)"
    elif selected_series == ["Excl"]:
        title = "Coil and Plate Repair Rate Trend (Excl. Skelp)"
    else:
        title = "Coil and Plate Repair Rate Trend — Skelp-Inclusive Trend"

    fig.update_layout(
        title=title,
        yaxis_title="Repair Rate (%)",
        xaxis_title="Date",
        template="plotly_white",
        height=580,
        margin=dict(l=40, r=20, t=50, b=40),
        hoverlabel=HOVER_STYLE,
    )
    fig.update_xaxes(tickformat="%d.%m.%y", dtick="D1", tickangle=-45)

    return dcc.Graph(figure=fig)


# ---------------------------------------------------------------------------
# Callback: Repair Rate Trend by Production Type — redraw for the lines
# selected in the checklist, so a small (e.g. 0.01%) change in one line isn't
# masked by the other lines sharing the same y-axis.
# ---------------------------------------------------------------------------

@callback(
    Output("type-trend-graph-content", "children"),
    Input("type-trend-checklist", "value"),
    Input("dashboard-stage-1", "data"),  # chained after dashboard-content, not parallel to it
    prevent_initial_call=True,
)
def render_type_trend_chart(selected_types, _stage_1):
    master_df = load_master_data()
    if master_df.empty:
        return _empty_state("No data available. Import an Excel file first.")
    if not selected_types:
        return _empty_state("Select at least one production type to display.")

    baseline_df = load_historical_baselines()
    window_start = master_df["date"].max() - pd.Timedelta(days=TREND_WINDOW_DAYS)
    _TYPE_COLORS = {"Coil": COLOR_COIL, "Plate": COLOR_PLATE}
    series_by_type = {}
    fig = go.Figure()
    for p_type in ("Coil", "Plate"):
        if p_type not in selected_types:
            continue
        type_trend = daily_weighted_repair_ratios_for_type(master_df, p_type, baseline_df).copy()
        # Round to the same 2-decimal-percent precision as the on-chart label
        # (round(ratio, 4) == round(ratio * 100, 2) / 100) so points that
        # display the same value also plot at the same height — otherwise
        # sub-label float noise renders as a jittery line.
        type_trend["weighted_repair_ratio"] = type_trend["weighted_repair_ratio"].round(4)
        type_trend = type_trend[type_trend["date"] >= window_start]
        series_by_type[p_type] = type_trend
        fig.add_trace(
            go.Scatter(
                x=type_trend["date"],
                y=type_trend["weighted_repair_ratio"] * 100,
                mode="lines+markers+text",
                name=p_type,
                line=dict(width=3, color=_TYPE_COLORS.get(p_type)),
                text=(type_trend["weighted_repair_ratio"] * 100).map("{:.2f}%".format),
                textposition="top center",
                textfont=dict(size=10, color=_TYPE_COLORS.get(p_type)),
                hovertemplate=f"%{{x|%d.%m.%Y}}<br>{p_type}: <b>%{{y:.2f}}%</b><extra></extra>",
            )
        )
    if "Mix" in selected_types:
        overall_ratio = daily_weighted_repair_ratios(master_df, baseline_df).copy()
        overall_ratio["weighted_repair_ratio"] = overall_ratio["weighted_repair_ratio"].round(4)
        overall_ratio = overall_ratio[overall_ratio["date"] >= window_start]
        series_by_type["Mix"] = overall_ratio
        fig.add_trace(
            go.Scatter(
                x=overall_ratio["date"],
                y=overall_ratio["weighted_repair_ratio"] * 100,
                mode="lines+markers+text",
                name="Mix (Coil + Plate)",
                line=dict(width=3, color=COLOR_SECONDARY, dash="dash"),
                text=(overall_ratio["weighted_repair_ratio"] * 100).map("{:.2f}%".format),
                textposition="bottom center",
                textfont=dict(size=10, color=COLOR_SECONDARY),
                hovertemplate="%{x|%d.%m.%Y}<br>Mix: <b>%{y:.2f}%</b><extra></extra>",
            )
        )

    # A trend line only makes sense when a single series is isolated —
    # overlaying it on top of 2-3 lines would add clutter, not remove it.
    if len(selected_types) == 1:
        trend_trace = _build_trend_trace(series_by_type[selected_types[0]])
        if trend_trace is not None:
            fig.add_trace(trend_trace)

    _TYPE_DISPLAY = {"Coil": "Coil", "Plate": "Plate", "Mix": "Mix"}
    ordered_selected = [t for t in ("Coil", "Plate", "Mix") if t in selected_types]
    if len(ordered_selected) == 1:
        title = f"Repair Rate Trend by Production Type — {_TYPE_DISPLAY[ordered_selected[0]]}"
    elif len(ordered_selected) == 2:
        title = (
            "Repair Rate Trend by Production Type — "
            f"{_TYPE_DISPLAY[ordered_selected[0]]} vs {_TYPE_DISPLAY[ordered_selected[1]]}"
        )
    else:
        title = "Repair Rate Trend by Production Type"

    fig.update_layout(
        title=title,
        yaxis_title="Repair Rate (%)",
        xaxis_title="Date",
        template="plotly_white",
        height=580,
        margin=dict(l=40, r=20, t=50, b=40),
        hoverlabel=HOVER_STYLE,
    )
    fig.update_xaxes(tickformat="%d.%m.%y", dtick="D1", tickangle=-45)

    return dcc.Graph(figure=fig)


def _collapse_unchanged_days(df: pd.DataFrame, value_col: str) -> pd.DataFrame:
    """Keep a row only where value_col changed from the previous day, plus
    always the first and last. A project can go untouched for days —
    repair_rates stores cumulative totals, so an untouched day's row repeats
    the previous report's value byte-for-byte — and plotting every one of
    those as a "new" point makes the line look busier than the underlying
    activity actually was."""
    if df.empty:
        return df
    changed = df[value_col].ne(df[value_col].shift())
    changed.iloc[0] = True
    changed.iloc[-1] = True  # always show through to the most recent date
    return df[changed]


def _gap_annotations(df: pd.DataFrame, value_col: str, date_index: dict, color: str, yshift: int) -> list[dict]:
    """'+N days' labels wherever _collapse_unchanged_days skipped more than
    one calendar day between two consecutive kept points — otherwise, once
    the axis is categorical (only kept dates get a tick), a two-day gap and
    a three-week gap look identical, since neither draws any empty axis
    space anymore. Positioned at the fractional category-index midpoint
    between the two points (a plain number, not a category label, is how
    Plotly places something *between* two ticks on a category axis)."""
    annotations = []
    dates = df["date"].tolist()
    values = (df[value_col] * 100).tolist()
    for i in range(1, len(dates)):
        gap_days = (dates[i] - dates[i - 1]).days
        if gap_days > 1:
            annotations.append(
                dict(
                    x=(date_index[dates[i - 1]] + date_index[dates[i]]) / 2,
                    y=(values[i - 1] + values[i]) / 2,
                    xref="x",
                    yref="y",
                    text=f"+{gap_days} days",
                    showarrow=False,
                    font=dict(size=9, color=color),
                    yshift=yshift,
                )
            )
    return annotations


# ---------------------------------------------------------------------------
# Callback: Project Trend — render the trend chart for the selected project
# (shares pipe-sheet-dropdown with Pipe-Level Analysis below — one project
# picker, not two, since the pipe-level tables were already the narrower/
# more relevant list).
# ---------------------------------------------------------------------------

@callback(
    Output("project-trend-content", "children"),
    Input("pipe-sheet-dropdown", "value"),
    Input("project-trend-checklist", "value"),
    Input("project-trend-compact-toggle", "value"),
    Input("dashboard-tabs", "value"),
)
def render_project_trend(selected_sheet, selected_series, compact_toggle, active_tab):
    if active_tab != "tab-pipe":
        return dash.no_update
    if not selected_sheet:
        return _empty_state("No data available. Import an Excel file first.")
    if not selected_series:
        return _empty_state("Select at least one series to display.")

    links_df = load_project_sheet_links()
    match = links_df[(links_df["project_sheet"] == selected_sheet) & (links_df["status"] == "confirmed")]
    if match.empty:
        return _empty_state("No confirmed project mapping for this sheet yet.")
    project_no = match.iloc[0]["project_no"]
    dimensions = match.iloc[0]["dimensions"]

    master_df = load_master_data()
    project_df = master_df[
        (master_df["project_no"] == project_no) & (master_df["dimensions"] == dimensions)
    ].copy()
    if project_df.empty:
        return _empty_state("No data available for this project.")

    project_df = apply_meter_based_repair_ratios(project_df).sort_values("date")
    project_df["repair_ratio"] = project_df["repair_ratio"].round(4)
    project_df["repair_ratio_incl_skelp"] = project_df["repair_ratio_incl_skelp"].round(4)
    window_start = project_df["date"].max() - pd.Timedelta(days=TREND_WINDOW_DAYS)
    project_df = project_df[project_df["date"] >= window_start]

    # Compact: drop days where the ratio didn't change (a project can go
    # untouched for days — see _collapse_unchanged_days) and use a
    # categorical axis so those empty days don't stretch the line across
    # axis space that has no data in it; a "+N days" annotation marks any
    # gap over a day so that missing time isn't just silently invisible.
    # Full: the plain, unmodified calendar-date view — every report date,
    # real date-scale spacing, no collapsing.
    compact = "compact" in (compact_toggle or [])
    excl_df = None
    incl_df = None
    if "Excl" in selected_series:
        excl_df = _collapse_unchanged_days(project_df, "repair_ratio") if compact else project_df
    if "Incl" in selected_series:
        incl_df = _collapse_unchanged_days(project_df, "repair_ratio_incl_skelp") if compact else project_df

    date_labels: dict = {}
    if compact:
        all_dates = sorted(
            set(excl_df["date"].tolist() if excl_df is not None else [])
            | set(incl_df["date"].tolist() if incl_df is not None else [])
        )
        date_labels = {d: d.strftime("%d.%m.%y") for d in all_dates}
        date_index = {d: i for i, d in enumerate(all_dates)}

    fig = go.Figure()
    annotations = []
    hover_x = "%{x}" if compact else "%{x|%d.%m.%Y}"

    if excl_df is not None:
        excl_x = [date_labels[d] for d in excl_df["date"]] if compact else excl_df["date"]
        fig.add_trace(
            go.Scatter(
                x=excl_x,
                y=excl_df["repair_ratio"] * 100,
                mode="lines+markers+text",
                name="Excl. Skelp",
                line=dict(color=COLOR_COIL, width=3),
                text=(excl_df["repair_ratio"] * 100).map("{:.2f}%".format),
                textposition="top center",
                textfont=dict(size=10, color=COLOR_COIL),
                hovertemplate=f"{hover_x}<br>Excl. Skelp: <b>%{{y:.2f}}%</b><extra></extra>",
            )
        )
        if compact:
            annotations += _gap_annotations(excl_df, "repair_ratio", date_index, COLOR_MUTED, yshift=22)
    if incl_df is not None:
        incl_x = [date_labels[d] for d in incl_df["date"]] if compact else incl_df["date"]
        fig.add_trace(
            go.Scatter(
                x=incl_x,
                y=incl_df["repair_ratio_incl_skelp"] * 100,
                mode="lines+markers+text",
                name="Incl. Skelp",
                line=dict(color=COLOR_SECONDARY, width=3, dash="dash"),
                text=(incl_df["repair_ratio_incl_skelp"] * 100).map("{:.2f}%".format),
                textposition="bottom center",
                textfont=dict(size=10, color=COLOR_SECONDARY),
                hovertemplate=f"{hover_x}<br>Incl. Skelp: <b>%{{y:.2f}}%</b><extra></extra>",
            )
        )
        if compact:
            annotations += _gap_annotations(incl_df, "repair_ratio_incl_skelp", date_index, COLOR_MUTED, yshift=-22)

    # Trend line only in Compact Timeline mode, only when a single series
    # is isolated. Outside compact mode the fit would have to use the full
    # uncollapsed history (every report row, including a dozen repeats of
    # the same value on days nothing changed) — that misrepresents the
    # trend by letting untouched days out-weigh the days something
    # actually happened, so it's not shown at all rather than shown wrong.
    # In compact mode the fit uses the collapsed (changed-days-only) frame
    # instead, which doesn't have that problem.
    if len(selected_series) == 1 and compact:
        value_col = "repair_ratio" if selected_series[0] == "Excl" else "repair_ratio_incl_skelp"
        trend_input_df = excl_df if selected_series[0] == "Excl" else incl_df
        trend_trace = _build_trend_trace(trend_input_df, value_col)
        if trend_trace is not None:
            start_date, end_date = trend_input_df["date"].min(), trend_input_df["date"].max()
            trend_trace.x = [date_labels[start_date], date_labels[end_date]]
            trend_trace.y = [trend_trace.y[0], trend_trace.y[-1]]
            # x is now a pre-formatted category label, not a real date —
            # _build_trend_trace's own hovertemplate expects a date and
            # formats it with %{x|%d.%m.%Y}, which on a plain string
            # renders as garbage ("NaN.NaN.0NaN").
            trend_trace.hovertemplate = "%{x}<br>Trend: <b>%{y:.2f}%</b><extra></extra>"
            fig.add_trace(trend_trace)

    if len(selected_series) == 2:
        skelp_suffix = " (Excl. vs Incl. Skelp)"
    elif selected_series == ["Excl"]:
        skelp_suffix = " (Excl. Skelp)"
    else:
        skelp_suffix = " — Skelp-Inclusive Trend"

    fig.update_layout(
        title=f"Repair Ratio Trend — {project_no} ({dimensions}){skelp_suffix}",
        yaxis_title="Repair Ratio (%)",
        xaxis_title="Date",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=40),
        hoverlabel=HOVER_STYLE,
        annotations=annotations,
    )
    if compact:
        all_date_labels = [date_labels[d] for d in sorted(date_labels)]
        fig.update_xaxes(type="category", categoryorder="array", categoryarray=all_date_labels, tickangle=-45)
    else:
        fig.update_xaxes(type="date", tickformat="%d.%m.%y", dtick="D1", tickangle=-45)

    return dcc.Graph(figure=fig)


# ---------------------------------------------------------------------------
# Callback 5: Pipe-Level Analysis — populate the project dropdown
# ---------------------------------------------------------------------------

def _pipe_sheet_label_map(links_df: pd.DataFrame) -> dict[str, str]:
    """project_sheet -> "project_no (dimensions)" for sheets confirmed via the
    "Project Mapping" review tool (feature/project-sheet-mapping-tool branch).
    Sheets with no confirmed mapping are simply absent from this map."""
    if links_df.empty:
        return {}
    confirmed = links_df[links_df["status"] == "confirmed"]
    labels = confirmed["project_no"].astype(str) + " (" + confirmed["dimensions"].astype(str) + ")"
    return dict(zip(confirmed["project_sheet"], labels))


@callback(
    Output("pipe-sheet-dropdown", "options"),
    Output("pipe-sheet-dropdown", "value"),
    Input("pipe-sheet-dropdown", "id"),  # fires once, on page load
    Input("import-confirm-result", "children"),  # refresh after a new import
)
def load_pipe_sheets(_id, _import_result):
    """pipe_repair_details is a current-state table (one row per physical
    pipe — see upsert_pipe_repair_details), so the sheet list is no longer
    scoped to a date; every mapped sheet with any current pipe shows up
    regardless of which date is picked above."""
    df = load_pipe_repair_details()
    if df.empty:
        return [], None
    sheets = set(df["project_sheet"].unique())
    # Only sheets with a confirmed project mapping show up here — the raw
    # Excel sheet name isn't meaningful to look at directly, and mixing
    # mapped/unmapped entries in one dropdown was confusing.
    label_map = _pipe_sheet_label_map(load_project_sheet_links())
    mapped_sheets = sorted(sheets & label_map.keys(), key=lambda s: label_map[s])
    options = [{"label": label_map[s], "value": s} for s in mapped_sheets]
    return options, options[0]["value"] if options else None


# ---------------------------------------------------------------------------
# Callback 7: Pipe-Level Analysis — render the summary + detail view
# ---------------------------------------------------------------------------

@callback(
    Output("pipe-analysis-content", "children"),
    Input("pipe-sheet-dropdown", "value"),
    Input("unit-toggle", "value"),
    Input("dashboard-tabs", "value"),
)
def render_pipe_analysis(selected_sheet, selected_unit, active_tab):
    """Everything here is scoped to one selected project — cross-project
    overview material (daily repair pace, all-projects summary) lives on
    the Dashboard tab instead (see render_pipe_overview), so this tab stays
    a focused single-project drill-down."""
    # Heavy (full pipe-details load + chart build) — skip while this tab
    # isn't visible instead of rebuilding on every unit-toggle change made
    # from another tab. Fires again once active_tab switches back here.
    if active_tab != "tab-pipe":
        return dash.no_update
    u = unit_label(selected_unit)
    if not selected_sheet:
        return _empty_state("No pipe-level data in the database yet. Import an Excel file first.")

    df = load_pipe_repair_details()
    if df.empty:
        return _empty_state("No pipe-level data in the database yet. Import an Excel file first.")

    links_df = load_project_sheet_links()
    master_df = load_master_data()
    label_map = _pipe_sheet_label_map(links_df)
    sheet_df = df[df["project_sheet"] == selected_sheet].sort_values("pipe_no")
    if sheet_df.empty:
        return _empty_state("No pipe-level data for this project yet.")

    sheet_label = label_map.get(selected_sheet, selected_sheet)

    # A still-Awaiting-Repair pipe has no real repair_ratio at all (NaN,
    # not 0) — feeding those into a repair-ratio chart doesn't just leave a
    # gap, it actively misleads: Plotly's line traces skip NaN points
    # entirely (so the "sequence" line looks broken/sparse rather than
    # showing real spacing), and a rolling average computed with min_periods=1
    # over mostly-NaN data collapses to whatever handful of real values
    # happen to fall in each window, producing spurious spikes/"islands"
    # that don't represent an actual 8-pipe average. The three charts below
    # that plot repair_ratio directly use only already-repaired pipes; the
    # detail table further down still shows every pipe, both statuses.
    repaired_sheet_df = sheet_df[sheet_df["status"] == "Repaired"]

    worst_fig = px.bar(
        worst_pipes(repaired_sheet_df, top_n=15),
        x="pipe_no",
        y="repair_ratio",
        labels={"pipe_no": "Pipe No.", "repair_ratio": "Repair Ratio"},
        title=f"Top 15 Pipes by Repair Ratio — {sheet_label}",
    )
    worst_fig.update_traces(
        marker_color=COLOR_COIL,
        texttemplate="%{y:.2%}",
        textposition="outside",
        textfont=dict(size=10),
        hovertemplate="Pipe %{x}<br>Repair Ratio: <b>%{y:.2%}</b><extra></extra>",
    )
    worst_fig.update_layout(template="plotly_white", margin=dict(l=40, r=20, t=50, b=40), hoverlabel=HOVER_STYLE)
    worst_fig.update_xaxes(type="category")
    worst_fig.update_yaxes(tickformat=".2%")

    # Pipes are produced in pipe_no order, so plotting every pipe against
    # that order (numeric axis, auto-ranging over the project's real
    # pipe_no span) shows whether bad pipes cluster at the start or end of
    # the run, with real gaps from never-repaired pipe numbers showing up
    # as real spacing — a categorical axis was tried here and rejected: a
    # production day isn't pipe_no-contiguous (e.g. one day can be pipes 1,
    # 13 and 42), so per-day line segments made no sense once the axis
    # needs to preserve numeric distance. One continuous line connects
    # every point in true pipe_no order instead; only marker color carries
    # the day.
    #
    # Marker color = production day, EXCEPT for pipes recorded before this
    # app started tracking cut date separately from repair date
    # (PIPE_TREND_FLOOR_DATE) — for those, first_seen_date is really just
    # "day first uploaded as already repaired" (the old parser only ever
    # captured pipes with repair data), not a real production date, so
    # color-coding it as one would be misleading. Those get one flat
    # neutral color instead of a day color.
    sequence_df = repaired_sheet_df.sort_values("pipe_no")
    is_tracked = sequence_df["first_seen_date"] >= PIPE_TREND_FLOOR_DATE

    # Okabe-Ito colorblind-safe set — engineered so adjacent hues stay
    # distinguishable under every common form of color vision deficiency.
    # Assigned by each tracked day's chronological rank, not by x-position,
    # so day N and day N+1 are never the same color regardless of how the
    # pipe numbers happen to fall.
    day_color_cycle = ["#E69F00", "#56B4E9", "#009E73", "#CC79A7", "#0072B2"]
    tracked_dates = sorted(sequence_df.loc[is_tracked, "first_seen_date"].unique())
    date_to_color = {d: day_color_cycle[i % len(day_color_cycle)] for i, d in enumerate(tracked_dates)}

    marker_colors = [
        date_to_color[d] if tracked else COLOR_MUTED
        for d, tracked in zip(sequence_df["first_seen_date"], is_tracked)
    ]
    hover_extra = [
        f"Produced: {d.strftime('%d.%m.%Y')}" if tracked else "Pre-tracking pipe — production date not recorded"
        for d, tracked in zip(sequence_df["first_seen_date"], is_tracked)
    ]

    sequence_fig = go.Figure()
    sequence_fig.add_trace(
        go.Scatter(
            x=sequence_df["pipe_no"],
            y=sequence_df["repair_ratio"],
            mode="lines+markers",
            name="Repair Ratio",
            line=dict(color=COLOR_COIL, width=1.5),
            marker=dict(size=6, color=marker_colors),
            customdata=hover_extra,
            hovertemplate="Pipe %{x}<br>%{customdata}<br>Repair Ratio: <b>%{y:.2%}</b><extra></extra>",
            showlegend=False,
        )
    )
    if not is_tracked.all():
        # A visible legend entry for the neutral color specifically (the
        # rotating day colors don't get one — with a 5-color rotation the
        # same color can mean different days, so a color->date legend
        # would be actively wrong, not just cluttered).
        sequence_fig.add_trace(
            go.Scatter(
                x=[None],
                y=[None],
                mode="markers",
                marker=dict(size=6, color=COLOR_MUTED),
                name="Pre-tracking (date unknown)",
            )
        )

    # One date label per tracked production day, above its lowest pipe_no
    # occurrence — only for days we actually know (pre-tracking pipes have
    # no reliable date at all, so they get no label, matching how their
    # hover text already omits a date). Capped the same way the old
    # per-day-trace version was, so a project with many tracked days
    # doesn't crowd the top of the chart.
    if tracked_dates:
        max_day_labels = 20
        label_every = max(1, -(-len(tracked_dates) // max_day_labels))
        tracked_only = sequence_df[is_tracked]
        for i, day in enumerate(tracked_dates):
            if i % label_every != 0:
                continue
            first_pipe_no = tracked_only.loc[tracked_only["first_seen_date"] == day, "pipe_no"].min()
            sequence_fig.add_annotation(
                x=first_pipe_no,
                y=1.0,
                yref="paper",
                yshift=6,
                text=pd.Timestamp(day).strftime("%d.%m"),
                showarrow=False,
                xanchor="left",
                font=dict(size=9, color=date_to_color[day]),
            )

    # Smoothed trend on top of the noisy per-pipe points — makes "is it
    # trending worse toward the start or end of the run" readable without
    # eyeballing every scattered point. min_periods=1 keeps this
    # well-behaved even on a 1-pipe sheet.
    rolling_window = 8 # at least 3 pipes, or 10% of the run
    rolling_avg = sequence_df["repair_ratio"].rolling(rolling_window, center=True, min_periods=1).mean()
    sequence_fig.add_trace(
        go.Scatter(
            x=sequence_df["pipe_no"],
            y=rolling_avg,
            mode="lines",
            name=f"{rolling_window}-Pipe Rolling Avg",
            line=dict(color=COLOR_DANGER, width=2, dash="dot"),
            hovertemplate="Pipe %{x}<br>Rolling Avg: <b>%{y:.2%}</b><extra></extra>",
        )
    )
    sequence_fig.update_layout(
        title=f"Repair Ratio by Production Order — {sheet_label}",
        xaxis_title="Pipe No.",
        yaxis_title="Repair Ratio",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=60),
        hoverlabel=HOVER_STYLE,
        # Without an explicit position, Plotly reserves a vertical sidebar
        # for the legend on the right, squeezing the actual plot into a
        # narrow left portion of the figure — placed below the plot
        # instead, matching the other charts in this file.
        legend=dict(orientation="h", yanchor="top", y=-0.25, xanchor="center", x=0.5),
    )
    # Numeric axis (not categorical) so it auto-ranges over the project's
    # real pipe_no span and gaps between repaired pipe numbers show up as
    # real, proportional spacing instead of being compressed away. Plotly's
    # own auto ticks for a numeric axis are sparse (e.g. every 10 units) —
    # labeling every actual pipe instead (thinned only if there are too
    # many to stay legible) keeps the true numeric spacing but shows far
    # more of the real pipe numbers along the axis.
    max_tick_labels = 40
    pipe_no_values = sequence_df["pipe_no"].tolist()
    tick_step = max(1, -(-len(pipe_no_values) // max_tick_labels))
    tick_values = pipe_no_values[::tick_step]
    sequence_fig.update_xaxes(
        type="linear",
        tickmode="array",
        tickvals=tick_values,
        ticktext=[str(v) for v in tick_values],
        tickangle=-45,
    )
    # rangemode="tozero": with very few pipes (a 1-pipe project in the
    # extreme), Plotly's autorange has no real spread to work from and can
    # center the only point in a huge, meaningless range (e.g. -50%..+100%
    # for a single ~1% value) — anchoring at zero keeps the scale sane.
    sequence_fig.update_yaxes(tickformat=".2%", rangemode="tozero")

    # How many of this project's pipes are still waiting on a repair,
    # already produced, or not even cut yet. The first two come straight
    # from sheet_df; "Planned" (not yet produced at all) has no signature
    # to count directly — a not-yet-produced pipe has nothing written in
    # its block at all, and raw "empty Excel block" counting was tried and
    # found unreliable (leftover template artifacts are indistinguishable
    # from real reserved slots — see the plan doc). Instead it's derived
    # from the project's own target quantity (repair_rates.qty, the Daily
    # Repair Rate sheet's planned/contract count) minus how many pipes
    # already exist here.
    status_counts = sheet_df["status"].value_counts().to_dict()
    produced_count = len(sheet_df)
    target_qty = None
    remaining = None
    confirmed_links = links_df[links_df["status"] == "confirmed"]
    link_match = confirmed_links[confirmed_links["project_sheet"] == selected_sheet]
    if not link_match.empty:
        project_no = link_match.iloc[0]["project_no"]
        dimensions = link_match.iloc[0]["dimensions"]
        project_master_df = master_df[
            (master_df["project_no"] == project_no) & (master_df["dimensions"] == dimensions)
        ]
        if not project_master_df.empty:
            qty = project_master_df.loc[project_master_df["date"].idxmax(), "qty"]
            if pd.notna(qty):
                target_qty = int(qty)
                remaining = max(target_qty - produced_count, 0)

    awaiting_repair_label = STATUS_DISPLAY_LABELS["Produced"]
    pie_entries = [
        ("Repaired", status_counts.get("Repaired", 0), COLOR_COIL),
        (awaiting_repair_label, status_counts.get("Produced", 0), COLOR_SECONDARY),
    ]
    if remaining is not None:
        pie_entries.append(("Planned", remaining, COLOR_MUTED))
    # Drop zero-value slices — Plotly still places a text label for a 0%
    # wedge (invisible slice, but not an invisible label), which collided
    # with its neighbor's label instead of just not being there.
    pie_entries = [entry for entry in pie_entries if entry[1] > 0]
    pie_names = [name for name, _, _ in pie_entries]
    pie_values = [value for _, value, _ in pie_entries]
    pie_colors = {name: color for name, _, color in pie_entries}

    status_fig = px.pie(
        names=pie_names,
        values=pie_values,
        title=f"Pipe Status — {sheet_label}",
        color=pie_names,
        color_discrete_map=pie_colors,
        hole=0.45,
    )
    status_fig.update_traces(
        textinfo="value+percent",
        textposition="inside",
        insidetextorientation="radial",
        hovertemplate="%{label}: <b>%{value}</b> pipes (%{percent})<extra></extra>",
    )
    status_fig.update_layout(template="plotly_white", margin=dict(l=20, r=20, t=50, b=20), hoverlabel=HOVER_STYLE)
    status_caption = None
    if target_qty is not None:
        # Mirrors the pie's own labels exactly (Repaired / Awaiting Repair /
        # Planned) instead of a lumped "produced so far" figure — that
        # phrasing was easy to misread against the Awaiting Repair count
        # right next to it, the same ambiguity the pie relabel just fixed.
        status_caption = html.P(
            f"Target Qty: {target_qty} · Repaired: {status_counts.get('Repaired', 0)} · "
            f"{awaiting_repair_label}: {status_counts.get('Produced', 0)} · Planned: {remaining}",
            className="help-text",
        )

    # Distribution of this one project's pipe ratios — median/quartiles/
    # outliers in a single box, complementing the production-order line
    # above (which shows *where* in the run things went wrong; this shows
    # *how consistent* the whole run was).
    box_df = repaired_sheet_df[["repair_ratio"]].copy()
    box_df["project_label"] = sheet_label
    box_df["repair_ratio_pct"] = (box_df["repair_ratio"] * 100).round(4)
    box_section = _box_plot_section(
        box_df,
        x_col="project_label",
        x_label="Project",
        title=f"Repair Ratio Distribution — {sheet_label}",
    )

    sheet_df = sheet_df.copy()
    if "pipe_length_ft" in sheet_df.columns:
        sheet_df["pipe_length_ft"] = length_in_display_unit(sheet_df["pipe_length_ft"], selected_unit)
    if "repair_amount" in sheet_df.columns:
        sheet_df["repair_amount"] = amount_in_display_unit(sheet_df["repair_amount"], selected_unit)
    if "repaired_date" in sheet_df.columns:
        # NaT (a still-Produced pipe) formats to None, not a crash — renders
        # as a blank cell same as the null repair_amount/repair_ratio above.
        sheet_df["repaired_date"] = sheet_df["repaired_date"].dt.strftime("%d.%m.%Y")
    if "status" in sheet_df.columns:
        sheet_df["status"] = sheet_df["status"].replace(STATUS_DISPLAY_LABELS)

    detail_columns = [
        c
        for c in sheet_df.columns
        if c not in ("id", "first_seen_date", "last_updated_date", "surface_state", "repair_category", "project_sheet")
    ]
    detail_table = dash_table.DataTable(
        data=_table_records(sheet_df, detail_columns),
        columns=_table_columns(
            detail_columns,
            label_overrides={"pipe_length_ft": f"Pipe Length ({u})", "repair_amount": f"Repair Amount ({u})"},
        ),
        page_size=20,
        sort_action="native",
        filter_action="native",
        style_table={"overflowX": "auto"},
        style_cell=TABLE_CELL_STYLE,
        style_header=TABLE_HEADER_STYLE,
        style_data_conditional=TABLE_CONDITIONAL_STYLE,
    )

    children = [
        html.H3(f"{sheet_label} — Pipe Details"),
        dcc.Graph(figure=status_fig, style={"maxWidth": "500px"}),
    ]
    if status_caption is not None:
        children.append(status_caption)
    children.extend(
        [
            dcc.Graph(figure=worst_fig),
            dcc.Graph(figure=sequence_fig),
            box_section,
            detail_table,
        ]
    )
    return html.Div(children)


# ---------------------------------------------------------------------------
# Callback: Comparison — populate the multi-select project dropdown
# ---------------------------------------------------------------------------

@callback(
    Output("comparison-project-dropdown", "options"),
    Output("comparison-project-dropdown", "value"),
    Input("comparison-project-dropdown", "id"),  # fires once, on page load
    Input("import-confirm-result", "children"),  # refresh after a new import
)
def load_comparison_options(_id, _import_result):
    links_df = load_project_sheet_links()
    label_map = _pipe_sheet_label_map(links_df)
    pipe_df = load_pipe_repair_details()
    if pipe_df.empty or not label_map:
        return [], []
    sheets = set(pipe_df["project_sheet"].unique())
    mapped_sheets = sorted(sheets & label_map.keys(), key=lambda s: label_map[s])
    options = [{"label": label_map[s], "value": s} for s in mapped_sheets]
    # Pre-select these so the Comparison tab shows a working example right
    # away instead of an empty "pick something" state — edit the sheet
    # names in COMPARISON_DEFAULT_SHEETS above to change which ones.
    default = [s for s in COMPARISON_DEFAULT_SHEETS if s in mapped_sheets] or [o["value"] for o in options[:2]]
    return options, default


# ---------------------------------------------------------------------------
# Callback: Comparison — trend/ratio/spread charts + table for the selected
# projects. Every chart here reuses an existing builder (the skelp-impact
# stacked bar, the per-project trend line, _build_box_fig) rather than
# inventing new plotting logic — see the plan doc for why each was picked.
# ---------------------------------------------------------------------------

@callback(
    Output("comparison-content", "children"),
    Input("comparison-project-dropdown", "value"),
    Input("comparison-series-checklist", "value"),
    Input("comparison-compact-toggle", "value"),
    Input("unit-toggle", "value"),
    Input("dashboard-tabs", "value"),
)
def render_comparison(selected_sheets, selected_series, compact_toggle, selected_unit, active_tab):
    if active_tab != "tab-comparison":
        return dash.no_update
    selected_sheets = selected_sheets or []
    if len(selected_sheets) < 2:
        return _empty_state("Select at least two projects to compare.")

    links_df = load_project_sheet_links()
    confirmed = links_df[links_df["status"] == "confirmed"]
    master_df = load_master_data()
    pipe_df = load_pipe_repair_details()
    u = unit_label(selected_unit)

    projects = []
    for sheet in selected_sheets:
        match = confirmed[confirmed["project_sheet"] == sheet]
        if match.empty:
            continue
        project_no = match.iloc[0]["project_no"]
        dimensions = match.iloc[0]["dimensions"]
        stripped_project_no = re.sub(r"^\d{4}Q-", "", project_no)
        label = f"{stripped_project_no} ({dimensions})"
        projects.append({"sheet": sheet, "project_no": project_no, "dimensions": dimensions, "label": label})
    if len(projects) < 2:
        return _empty_state("Select at least two projects to compare.")
    projects_df = pd.DataFrame(projects)

    mask = pd.Series(False, index=master_df.index)
    for p in projects:
        mask |= (master_df["project_no"] == p["project_no"]) & (master_df["dimensions"] == p["dimensions"])
    selected_master_df = apply_meter_based_repair_ratios(master_df[mask].copy())
    selected_master_df = selected_master_df.merge(
        projects_df[["project_no", "dimensions", "label", "sheet"]], on=["project_no", "dimensions"], how="inner"
    )
    if selected_master_df.empty:
        return _empty_state("No data available for the selected projects.")

    # --- Repair Ratio Trend: one line per project (color) x per selected
    # series (line style) — the same feature set as Pipe Analysis's Project
    # Trend chart (render_project_trend), generalized to several projects
    # at once: flat/untouched-day runs collapsed per line, an optional
    # Compact Timeline (categorical axis + "+N days" gap annotations),
    # always-on point labels, and a single-series trend-line fit. ---
    SERIES_META = {
        "Excl": {"col": "repair_ratio", "dash": "solid", "text_pos": "top center", "label": "Excl. Skelp"},
        "Incl": {"col": "repair_ratio_incl_skelp", "dash": "dash", "text_pos": "bottom center", "label": "Incl. Skelp"},
    }
    selected_series = [s for s in (selected_series or []) if s in SERIES_META]
    compact = "compact" in (compact_toggle or [])
    window_start = selected_master_df["date"].max() - pd.Timedelta(days=TREND_WINDOW_DAYS)
    palette = px.colors.qualitative.Plotly

    per_project_series: dict = {}
    for i, p in enumerate(projects):
        proj_df = selected_master_df[selected_master_df["sheet"] == p["sheet"]].sort_values("date")
        proj_df = proj_df[proj_df["date"] >= window_start].copy()
        if proj_df.empty:
            continue
        proj_df["repair_ratio"] = proj_df["repair_ratio"].round(4)
        proj_df["repair_ratio_incl_skelp"] = proj_df["repair_ratio_incl_skelp"].round(4)
        for series in selected_series:
            value_col = SERIES_META[series]["col"]
            per_project_series[(i, series)] = _collapse_unchanged_days(proj_df, value_col) if compact else proj_df

    date_labels: dict = {}
    if compact:
        all_dates = sorted({d for df in per_project_series.values() for d in df["date"]})
        date_labels = {d: d.strftime("%d.%m.%y") for d in all_dates}
        date_index = {d: i for i, d in enumerate(all_dates)}

    trend_fig = go.Figure()
    trend_annotations = []
    hover_x = "%{x}" if compact else "%{x|%d.%m.%Y}"
    for (i, series), series_df in per_project_series.items():
        p = projects[i]
        meta = SERIES_META[series]
        color = palette[i % len(palette)]
        x = [date_labels[d] for d in series_df["date"]] if compact else series_df["date"]
        trend_fig.add_trace(
            go.Scatter(
                x=x,
                y=series_df[meta["col"]] * 100,
                mode="lines+markers+text",
                name=f"{p['label']} — {meta['label']}",
                line=dict(color=color, width=2.5, dash=meta["dash"]),
                marker=dict(size=6, color=color),
                text=(series_df[meta["col"]] * 100).map("{:.2f}%".format),
                textposition=meta["text_pos"],
                textfont=dict(size=9, color=color),
                hovertemplate=f"{hover_x}<br>{p['label']} ({meta['label']}): <b>%{{y:.2f}}%</b><extra></extra>",
            )
        )
        if compact:
            trend_annotations += _gap_annotations(
                series_df, meta["col"], date_index, color, yshift=22 if series == "Excl" else -22
            )

    trend_fig.update_layout(
        title="Repair Ratio Trend (Selected Projects)",
        xaxis_title="Date",
        yaxis_title="Repair Ratio (%)",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=40),
        hoverlabel=HOVER_STYLE,
        annotations=trend_annotations,
    )
    if compact:
        category_order = [date_labels[d] for d in sorted(date_labels)]
        trend_fig.update_xaxes(type="category", categoryorder="array", categoryarray=category_order, tickangle=-45)
    else:
        trend_fig.update_xaxes(type="date", tickformat="%d.%m.%y", dtick="D1", tickangle=-45)

    # --- Latest-Day Repair Ratio: same stacked base+skelp-impact bar the
    # dashboard uses for its "top 5" chart, filtered to the selected
    # projects instead — "where does each stand today." ---
    snapshot_df = (
        selected_master_df.sort_values("date")
        .groupby(["project_no", "dimensions"], as_index=False)
        .tail(1)
        .sort_values("repair_ratio", ascending=False)
    )
    snapshot_df["skelp_impact"] = (snapshot_df["repair_ratio_incl_skelp"] - snapshot_df["repair_ratio"]).round(6)
    max_stack = (snapshot_df["repair_ratio"] + snapshot_df["skelp_impact"]).max() if len(snapshot_df) else 0
    base_text = [
        f"{r:.2%}<br>Qty {q:.0f}" if max_stack and r >= max_stack * 0.1 else ""
        for r, q in zip(snapshot_df["repair_ratio"], snapshot_df["qty"])
    ]
    impact_text = [
        f"+{imp:.2%}" if max_stack and imp >= max_stack * 0.08 else "" for imp in snapshot_df["skelp_impact"]
    ]
    ratio_fig = go.Figure()
    ratio_fig.add_trace(
        go.Bar(
            x=snapshot_df["label"],
            y=snapshot_df["repair_ratio"],
            name="Repair Ratio",
            marker_color=COLOR_COIL,
            text=base_text,
            textposition="inside",
            textfont=dict(size=11, color="white"),
            hovertemplate="%{x}<br>Repair Ratio: <b>%{y:.2%}</b><extra></extra>",
        )
    )
    ratio_fig.add_trace(
        go.Bar(
            x=snapshot_df["label"],
            y=snapshot_df["skelp_impact"],
            name="Skelp Impact",
            marker_color=COLOR_SECONDARY,
            text=impact_text,
            textposition="inside",
            textfont=dict(size=11, color="white"),
            hovertemplate="%{x}<br>Skelp Impact: <b>+%{y:.2%}</b><extra></extra>",
        )
    )
    for label, r, imp, q, base_t, impact_t in zip(
        snapshot_df["label"], snapshot_df["repair_ratio"], snapshot_df["skelp_impact"], snapshot_df["qty"],
        base_text, impact_text,
    ):
        lines = []
        if not impact_t:
            lines.append(f"+{imp:.2%}")
        if not base_t:
            lines.append(f"{r:.2%} (Qty {q:.0f})")
        if lines:
            ratio_fig.add_annotation(
                x=label, y=r + imp, text="<br>".join(lines), showarrow=False, yshift=14,
                font=dict(size=9, color="#1e293b"),
            )
    ratio_fig.update_layout(
        barmode="stack",
        title="Latest-Day Repair Ratio (Selected Projects)",
        yaxis_title="Repair Ratio",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=80),
        hoverlabel=HOVER_STYLE,
        xaxis=dict(categoryorder="array", categoryarray=snapshot_df["label"].tolist()),
    )
    ratio_fig.update_xaxes(tickangle=-25)
    ratio_fig.update_yaxes(tickformat=".2%")

    # --- Repair Ratio Distribution: one box per project, pipe-level spread
    # comparison — _build_box_fig already groups by whatever x_col you give
    # it, so handing it just the selected projects' pipe rows is a direct
    # reuse, not a new chart. ---
    box_section = None
    box_df = pipe_df[pipe_df["project_sheet"].isin([p["sheet"] for p in projects])].copy()
    if not box_df.empty:
        label_by_sheet = dict(zip(projects_df["sheet"], projects_df["label"]))
        box_df["project_label"] = box_df["project_sheet"].map(label_by_sheet)
        box_df["repair_ratio_pct"] = (box_df["repair_ratio"] * 100).round(4)
        box_fig = _build_box_fig(
            box_df,
            "project_label",
            "Project",
            "Repair Ratio Distribution (Selected Projects)",
            category_order=snapshot_df["label"].tolist(),
        )
        box_section = dcc.Graph(figure=box_fig)

    # --- Summary table: exact numbers, since every chart above rounds for
    # readability. ---
    table_df = snapshot_df[
        ["label", "dimensions", "qty", "repair_ratio", "repair_ratio_incl_skelp", "total_repair_amount", "project_status", "date"]
    ].copy()
    table_df["total_repair_amount"] = amount_in_display_unit(table_df["total_repair_amount"], selected_unit)
    table_df["date"] = table_df["date"].dt.strftime("%d.%m.%Y")
    table_columns = [
        "label", "dimensions", "qty", "repair_ratio", "repair_ratio_incl_skelp", "total_repair_amount",
        "project_status", "date",
    ]
    comparison_table = dash_table.DataTable(
        data=_table_records(table_df, table_columns),
        columns=_table_columns(
            table_columns,
            label_overrides={
                "label": "Project",
                "total_repair_amount": f"Repair Amount ({u})",
                "date": "Latest Report",
            },
        ),
        page_size=10,
        sort_action="native",
        style_table={"overflowX": "auto"},
        style_cell=TABLE_CELL_STYLE,
        style_header=TABLE_HEADER_STYLE,
        style_data_conditional=TABLE_CONDITIONAL_STYLE,
    )

    trend_section = (
        dcc.Graph(figure=trend_fig)
        if selected_series
        else _empty_state("Select at least one series to display the trend chart.")
    )
    children = [
        html.H3(f"Comparing {len(projects)} Projects"),
        trend_section,
        dcc.Graph(figure=ratio_fig),
    ]
    if box_section is not None:
        children.append(box_section)
    children.append(html.H4("Summary"))
    children.append(comparison_table)
    return html.Div(children)


# ---------------------------------------------------------------------------
# Callback 11: Dimension Detail — populate the dimension dropdown
# ---------------------------------------------------------------------------

@callback(
    Output("dimension-detail-dropdown", "options"),
    Output("dimension-detail-dropdown", "value"),
    Input("dimension-detail-dropdown", "id"),  # fires once, on page load
    Input("import-confirm-result", "children"),  # refresh after a new import
)
def load_dimension_options(_id, _import_result):
    master_df = load_master_data()
    if master_df.empty:
        return [], None
    latest_date = master_df["date"].max()
    latest_df = master_df[master_df["date"] == latest_date]
    dims = sorted(latest_df["dimensions"].dropna().unique())
    options = [{"label": d, "value": d} for d in dims]
    return options, options[0]["value"] if options else None


# ---------------------------------------------------------------------------
# Callback 12: Dimension Detail — compare projects sharing the selected dimension
# ---------------------------------------------------------------------------

@callback(
    Output("dimension-overview-content", "children"),
    Input("dashboard-stage-2", "data"),  # chained after pipe-overview-content, not parallel to it
    State("unit-toggle", "value"),
    prevent_initial_call=True,
)
def render_dimension_overview(_stage_2, selected_unit):
    """The two "all dimensions" charts — independent of which single
    dimension is selected below, so this is its own callback rather than
    part of render_dimension_detail: that function used to rebuild these
    two on every dropdown change even though neither one depends on the
    dropdown's value."""
    master_df = load_master_data()
    if master_df.empty:
        return _empty_state("No data available. Import an Excel file first.")
    latest_date, latest_df = _latest_day_frame(master_df)
    u = unit_label(selected_unit)

    # Ranked by weighted_ratio (not total_repair_amount) since that's the
    # metric the chart actually displays — capping by volume first could
    # silently drop a high-ratio, low-volume dimension the chart's own
    # title implies would be shown.
    overview_dim_df = latest_df.groupby("dimensions", as_index=False).agg(
        total_repair_amount=("total_repair_amount", "sum"),
        repaired_spiral_length=("repaired_spiral_length", "sum"),
    )
    overview_denom = overview_dim_df["repaired_spiral_length"] * METERS_PER_FOOT
    overview_dim_df["weighted_ratio"] = (
        overview_dim_df["total_repair_amount"] / overview_denom.where(overview_denom != 0)
    ).fillna(0)
    overview_dim_df = overview_dim_df.sort_values("weighted_ratio", ascending=False).head(15)
    dimension_ratio_fig = px.bar(
        overview_dim_df,
        x="dimensions",
        y="weighted_ratio",
        labels={"dimensions": "Dimensions", "weighted_ratio": "Weighted Repair Ratio"},
        title="Weighted Repair Ratio by Dimension (Latest Day)",
    )
    dimension_ratio_fig.update_traces(
        marker_color=COLOR_COIL,
        texttemplate="%{y:.2%}",
        textposition="outside",
        textfont=dict(size=11),
        hovertemplate="%{x}<br>Weighted Repair Ratio: <b>%{y:.2%}</b><extra></extra>",
    )
    dimension_ratio_fig.update_layout(template="plotly_white", margin=dict(l=40, r=20, t=50, b=80), hoverlabel=HOVER_STYLE)
    dimension_ratio_fig.update_xaxes(tickangle=-45)
    dimension_ratio_fig.update_yaxes(tickformat=".2%")

    overview_bubble_df = latest_df.groupby("dimensions", as_index=False).agg(
        repaired_spiral_length=("repaired_spiral_length", "sum"),
        total_repair_amount=("total_repair_amount", "sum"),
    )
    overview_bubble_denom = overview_bubble_df["repaired_spiral_length"] * METERS_PER_FOOT
    overview_bubble_df["repair_ratio_pct"] = (
        (overview_bubble_df["total_repair_amount"] / overview_bubble_denom.where(overview_bubble_denom != 0)).fillna(0)
        * 100
    ).round(4)
    overview_bubble_df["repaired_spiral_length"] = length_in_display_unit(
        overview_bubble_df["repaired_spiral_length"], selected_unit
    )
    overview_bubble_df["total_repair_amount"] = amount_in_display_unit(
        overview_bubble_df["total_repair_amount"], selected_unit
    )
    dimension_bubble_fig = _build_bubble_fig(
        overview_bubble_df,
        x_col="repaired_spiral_length",
        x_label=f"Total Spiral Length ({u})",
        title=f"Repair Ratio vs. Production Volume by Dimension ({u}, Spiral Length, Latest Day)",
        unit=u,
        label_col="dimensions",
    )

    return html.Div([dcc.Graph(figure=dimension_ratio_fig), dcc.Graph(figure=dimension_bubble_fig)])


# ---------------------------------------------------------------------------
# Callback 12b: Dimension Detail — the per-selected-dimension drill-down
# ---------------------------------------------------------------------------

@callback(
    Output("dimension-detail-content", "children"),
    Output("dimension-detail-table-content", "children"),
    Input("dimension-detail-dropdown", "value"),
)
def render_dimension_detail(selected_dimension):
    """Returns (chart_content, table_content) — split so the table lands in
    the dedicated Tables section at the bottom of the page instead of in
    the middle of this dimension's charts."""
    if not selected_dimension:
        empty = _empty_state("No data available. Import an Excel file first.")
        return empty, None

    master_df = load_master_data()
    _latest_date, latest_df = _latest_day_frame(master_df)

    dim_df = latest_df[latest_df["dimensions"] == selected_dimension].sort_values(
        "repair_ratio", ascending=False
    ).copy()
    if dim_df.empty:
        empty = _empty_state("No projects found for this dimension on the latest day.")
        return empty, None

    # Repair-ratio chart — same clean (no 20XXQ- prefix) label as the
    # dashboard's ratio charts, with qty shown inside the bar instead.
    # project_label_clean already came along from latest_df (_latest_day_frame).

    dim_trend = daily_weighted_repair_ratios_for_dimension(master_df, selected_dimension).copy()
    dim_trend["weighted_repair_ratio"] = dim_trend["weighted_repair_ratio"].round(4)
    dim_trend["weighted_repair_ratio_incl_skelp"] = dim_trend["weighted_repair_ratio_incl_skelp"].round(4)
    dim_window_start = dim_trend["date"].max() - pd.Timedelta(days=TREND_WINDOW_DAYS)
    dim_trend = dim_trend[dim_trend["date"] >= dim_window_start]
    trend_fig = go.Figure()
    trend_fig.add_trace(
        go.Scatter(
            x=dim_trend["date"],
            y=dim_trend["weighted_repair_ratio"] * 100,
            mode="lines+markers+text",
            name="Excl. Skelp",
            line=dict(color=COLOR_COIL, width=3),
            text=(dim_trend["weighted_repair_ratio"] * 100).map("{:.2f}%".format),
            textposition="top center",
            textfont=dict(size=10, color=COLOR_COIL),
            hovertemplate="%{x|%d.%m.%Y}<br>Excl. Skelp: <b>%{y:.2f}%</b><extra></extra>",
        )
    )
    trend_fig.add_trace(
        go.Scatter(
            x=dim_trend["date"],
            y=dim_trend["weighted_repair_ratio_incl_skelp"] * 100,
            mode="lines+markers+text",
            name="Incl. Skelp",
            line=dict(color=COLOR_SECONDARY, width=3, dash="dash"),
            text=(dim_trend["weighted_repair_ratio_incl_skelp"] * 100).map("{:.2f}%".format),
            textposition="bottom center",
            textfont=dict(size=10, color=COLOR_SECONDARY),
            hovertemplate="%{x|%d.%m.%Y}<br>Incl. Skelp: <b>%{y:.2f}%</b><extra></extra>",
        )
    )
    trend_fig.update_layout(
        title=f"Repair Rate Trend — Dimension {selected_dimension}",
        yaxis_title="Repair Rate (%)",
        xaxis_title="Date",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=40),
        hoverlabel=HOVER_STYLE,
    )
    trend_fig.update_xaxes(tickformat="%d.%m.%y", dtick="D1", tickangle=-45)

    fig = px.bar(
        dim_df,
        x="project_label_clean",
        y="repair_ratio",
        labels={"project_label_clean": "Project", "repair_ratio": "Repair Ratio"},
        title=f"Projects with Dimension {selected_dimension} (Latest Day)",
    )
    fig.update_traces(
        marker_color=COLOR_COIL,
        texttemplate="%{y:.2%}",
        textposition="outside",
        textfont=dict(size=11),
        hovertemplate="%{x}<br>Repair Ratio: <b>%{y:.2%}</b><extra></extra>",
    )
    fig.update_layout(template="plotly_white", margin=dict(l=40, r=20, t=50, b=80), hoverlabel=HOVER_STYLE)
    fig.update_xaxes(tickangle=-45)
    fig.update_yaxes(tickformat=".2%")
    _add_bar_value_annotations(fig, dim_df["project_label_clean"], dim_df["repair_ratio"], dim_df["qty"])

    table_columns = ["project_no", "production_type", "qty", "project_status", "repair_ratio"]
    table = dash_table.DataTable(
        data=_table_records(dim_df, table_columns),
        columns=_table_columns(table_columns),
        page_size=10,
        sort_action="native",
        style_table={"overflowX": "auto"},
        style_cell=TABLE_CELL_STYLE,
        style_header=TABLE_HEADER_STYLE,
        style_data_conditional=TABLE_CONDITIONAL_STYLE,
    )

    children = [
        html.H3(f"Dimension {selected_dimension} — Detail"),
        dcc.Graph(figure=trend_fig),
        dcc.Graph(figure=fig),
    ]

    # Pipe-level distribution, one box per project sharing this dimension —
    # only the handful of projects with this exact dimension, so it stays
    # readable (unlike a single chart trying to compare every dimension or
    # every project on the dashboard at once). Pipe-level data is optional
    # (feature/project-sheet-mapping-tool branch), so this is skipped
    # gracefully if none of these projects have a confirmed sheet mapping
    # or any pipe-level rows yet.
    links_df = load_project_sheet_links()
    dim_sheets = links_df[(links_df["status"] == "confirmed") & (links_df["dimensions"] == selected_dimension)]
    if not dim_sheets.empty:
        pipe_df = load_pipe_repair_details()
        box_df = pipe_df[pipe_df["project_sheet"].isin(dim_sheets["project_sheet"])].copy()
        if not box_df.empty:
            label_map = dict(zip(dim_sheets["project_sheet"], dim_sheets["project_no"]))
            box_df["project_label"] = box_df["project_sheet"].map(label_map)
            box_df["repair_ratio_pct"] = (box_df["repair_ratio"] * 100).round(4)
            box_section = _box_plot_section(
                box_df,
                x_col="project_label",
                x_label="Project",
                title=f"Repair Ratio Distribution — Pipe-Level Detail, Dimension {selected_dimension}",
            )
            children.append(box_section)

    table_content = html.Div([html.H3(f"Dimension {selected_dimension} — Detail Table"), table])
    return html.Div(children), table_content
