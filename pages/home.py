"""
Dashboard page: daily Excel import, historical baseline import, the main
dashboard, pipe-level drill-down, and project grouping — organized as tabs
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

import dash
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import Input, Output, State, callback, dash_table, dcc, html
from dash.dash_table.Format import Format, Scheme

from baseline import baseline_template_csv, parse_historical_baseline_csv
from calculations import (
    METERS_PER_FOOT,
    apply_meter_based_repair_ratios,
    daily_weighted_repair_ratios,
    daily_weighted_repair_ratios_for_type,
    repair_amount_trend_data,
)
from database import (
    load_historical_baselines,
    load_master_data,
    load_pipe_repair_details,
    load_project_group_config,
    upsert_historical_baselines,
    upsert_pipe_repair_details,
    upsert_project_group_config,
    upsert_repair_rates,
)
from parser import parse_daily_repair_rate
from pdf_report import build_pdf_report
from pipe_analysis import summarize_pipe_totals_by_sheet, worst_pipes
from project_parser import parse_project_pipe_repairs

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
}

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


TEXT_COLUMNS = {
    "date",
    "project_no",
    "dimensions",
    "project_status",
    "production_type",
    "project_sheet",
    "block_cell",
    "repair_category",
    "surface_state",
}
# Whole-number counts: no decimal formatting.
INTEGER_COLUMNS = {"qty", "pipe_no", "repair_count", "pipe_count", "project_count"}


def _table_columns(columns) -> list[dict]:
    cols = []
    for c in columns:
        name = COLUMN_LABELS.get(c, c.replace("_", " ").title())
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
    a percentage, so filtering/sorting stays numeric rather than string."""
    out = df[columns].copy()
    for col in columns:
        if col in PERCENT_COLUMNS:
            continue
        if pd.api.types.is_numeric_dtype(out[col]):
            out[col] = out[col].round(4)
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
                                    html.Div(id="upload-validation-result"),
                                    html.Div(id="upload-preview-table"),
                                    html.Div(id="pipe-parse-summary"),
                                    html.Button(
                                        "Confirm Import",
                                        id="confirm-import-btn",
                                        className="primary-btn",
                                        style={"display": "none"},
                                    ),
                                    html.Div(id="import-confirm-result"),
                                ],
                                className="card",
                            ),
                            html.Section(
                                [
                                    html.H2("Historical Baseline"),
                                    html.P(
                                        "Upload a CSV to include the total repair amount of projects "
                                        "carried over from previous years in the overall repair rate calculation.",
                                        className="help-text",
                                    ),
                                    html.Button(
                                        "Download Template CSV",
                                        id="download-baseline-template-btn",
                                        className="secondary-btn",
                                    ),
                                    dcc.Download(id="baseline-template-download"),
                                    dcc.Store(id="parsed-baseline-store"),
                                    dcc.Upload(
                                        id="baseline-upload",
                                        children=html.Div(
                                            ["Drag and drop the CSV file here or ", html.A("select a file")]
                                        ),
                                        className="upload-box",
                                        multiple=False,
                                    ),
                                    html.Div(id="baseline-validation-result"),
                                    html.Div(id="baseline-preview-table"),
                                    html.Button(
                                        "Confirm Baseline Import",
                                        id="confirm-baseline-import-btn",
                                        className="primary-btn",
                                        style={"display": "none"},
                                    ),
                                    html.Div(id="baseline-import-confirm-result"),
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
                                ],
                                className="card",
                            ),
                            html.Section(
                                [
                                    html.H2("Project Trend"),
                                    html.P(
                                        "Select a project to see its repair ratio over all report dates.",
                                        className="help-text",
                                    ),
                                    html.Div(
                                        [
                                            html.Div(
                                                [html.Label("Project"), dcc.Dropdown(id="project-trend-dropdown")],
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
                                    html.H2("Dimension Detail"),
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
                                    dcc.Loading(html.Div(id="pipe-analysis-content")),
                                ],
                                className="card",
                            ),
                        ],
                    ),
                    dcc.Tab(
                        label="Project Grouping",
                        value="tab-groups",
                        style=_TAB_STYLE,
                        selected_style=_TAB_SELECTED_STYLE,
                        children=[
                            html.Section(
                                [
                                    html.H2("Project Grouping"),
                                    html.P(
                                        "Define named pipe-number ranges for a project sheet, for use in "
                                        "group-level analysis. Format: \"1-20:Group A; 21-45:Group B\".",
                                        className="help-text",
                                    ),
                                    html.Div(
                                        [
                                            html.Div(
                                                [html.Label("Project Sheet"), dcc.Dropdown(id="group-sheet-dropdown")],
                                                className="filter-field",
                                            ),
                                        ],
                                        className="filter-row",
                                    ),
                                    html.Div(
                                        [
                                            html.Div(
                                                [
                                                    html.Label("Pipe Groups"),
                                                    dcc.Textarea(
                                                        id="pipe-groups-input",
                                                        placeholder="e.g. 1-20:Team A; 21-45:Team B",
                                                        className="group-textarea",
                                                    ),
                                                ],
                                                className="filter-field",
                                            ),
                                            html.Div(
                                                [
                                                    html.Label("Machine Groups"),
                                                    dcc.Textarea(
                                                        id="machine-groups-input",
                                                        placeholder="e.g. 1-30:Machine 1; 31-60:Machine 2",
                                                        className="group-textarea",
                                                    ),
                                                ],
                                                className="filter-field",
                                            ),
                                        ],
                                        className="filter-row",
                                    ),
                                    html.Button("Save Groups", id="save-groups-btn", className="primary-btn"),
                                    html.Div(id="save-groups-result"),
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
    check_rows = [
        html.Li(f"{'✅' if ok else '❌'} {name}", className="check-ok" if ok else "check-fail")
        for name, ok in report.checks.items()
    ]
    error_rows = [html.Li(err) for err in report.errors]

    children = [html.Ul(check_rows)]
    if report.errors:
        children.append(html.H4(f"Errors ({len(report.errors)})", className="error-heading"))
        children.append(html.Ul(error_rows, className="error-list"))
    else:
        children.append(html.P(f"{report.import_rows} rows read, no errors.", className="success-text"))

    return html.Div(children, className="validation-box ok" if report.ok else "validation-box fail")


def _empty_state(message: str) -> html.Div:
    return html.Div(message, className="empty-state")


def _baseline_validation_box(errors: list[str], row_count: int) -> html.Div:
    if errors:
        return html.Div(
            [
                html.H4(f"Errors ({len(errors)})", className="error-heading"),
                html.Ul([html.Li(err) for err in errors], className="error-list"),
            ],
            className="validation-box fail",
        )
    return html.Div(
        html.P(f"{row_count} rows read, no errors.", className="success-text"),
        className="validation-box ok",
    )


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
    Input("excel-upload", "contents"),
    State("excel-upload", "filename"),
    prevent_initial_call=True,
)
def handle_upload(contents, filename):
    if contents is None:
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update

    try:
        file_obj = _decode_upload(contents)
        df, report = parse_daily_repair_rate(file_obj)
    except Exception as exc:  # don't crash the app on an unexpected error
        error_box = html.Div(
            f"An unexpected error occurred while reading the file: {exc}",
            className="validation-box fail",
        )
        return error_box, None, None, {"display": "none"}, None, None

    validation_box = _validation_summary(report)

    if df.empty or not report.ok:
        return validation_box, None, None, {"display": "none"}, None, None

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

    pipe_summary = None
    pipe_data_json = None
    try:
        pipe_file_obj = _decode_upload(contents)
        pipe_df, pipe_report = parse_project_pipe_repairs(pipe_file_obj, df["date"].iloc[0])
        if not pipe_df.empty:
            pipe_summary = html.P(
                f"Pipe-level: {pipe_report.parsed_rows} rows parsed from "
                f"{pipe_report.parsed_sheets} project sheets.",
                className="help-text",
            )
            pipe_data_json = pipe_df.to_json(date_format="iso", orient="split")
    except Exception as exc:  # pipe-level parsing is best-effort, never blocks the main import
        pipe_summary = html.P(
            f"Pipe-level parsing failed (main import is unaffected): {exc}",
            className="help-text",
        )

    return (
        validation_box,
        html.Div([html.H4(f"Preview — {filename}"), preview], className="preview-box"),
        pipe_summary,
        {"display": "inline-block"},
        df.to_json(date_format="iso", orient="split"),
        pipe_data_json,
    )


# ---------------------------------------------------------------------------
# Callback 2: "Confirm Import" -> write to database
# ---------------------------------------------------------------------------

@callback(
    Output("import-confirm-result", "children"),
    Output("parsed-data-store", "data", allow_duplicate=True),
    Output("parsed-pipe-data-store", "data", allow_duplicate=True),
    Input("confirm-import-btn", "n_clicks"),
    State("parsed-data-store", "data"),
    State("parsed-pipe-data-store", "data"),
    prevent_initial_call=True,
)
def confirm_import(n_clicks, stored_json, stored_pipe_json):
    if not n_clicks or not stored_json:
        return dash.no_update, dash.no_update, dash.no_update

    df = pd.read_json(io.StringIO(stored_json), orient="split")
    written = upsert_repair_rates(df)

    pipe_written = 0
    if stored_pipe_json:
        pipe_df = pd.read_json(io.StringIO(stored_pipe_json), orient="split")
        pipe_written = upsert_pipe_repair_details(pipe_df)

    message = f"✅ {written} rows saved to the database."
    if pipe_written:
        message += f" ({pipe_written} pipe-level rows saved.)"

    return (
        html.Div(message, className="success-text"),
        None,
        None,
    )


# ---------------------------------------------------------------------------
# Callback 3: Download baseline template CSV
# ---------------------------------------------------------------------------

@callback(
    Output("baseline-template-download", "data"),
    Input("download-baseline-template-btn", "n_clicks"),
    prevent_initial_call=True,
)
def download_baseline_template(n_clicks):
    if not n_clicks:
        return dash.no_update
    return dcc.send_string(baseline_template_csv(), "historical_baseline_template.csv")


# ---------------------------------------------------------------------------
# Callback 4: Upload baseline CSV -> parse -> validate -> show preview
# ---------------------------------------------------------------------------

@callback(
    Output("baseline-validation-result", "children"),
    Output("baseline-preview-table", "children"),
    Output("confirm-baseline-import-btn", "style"),
    Output("parsed-baseline-store", "data"),
    Input("baseline-upload", "contents"),
    State("baseline-upload", "filename"),
    prevent_initial_call=True,
)
def handle_baseline_upload(contents, filename):
    if contents is None:
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update

    try:
        file_obj = _decode_upload(contents)
        df, errors = parse_historical_baseline_csv(file_obj)
    except Exception as exc:  # don't crash the app on an unexpected error
        error_box = html.Div(
            f"An unexpected error occurred while reading the file: {exc}",
            className="validation-box fail",
        )
        return error_box, None, {"display": "none"}, None

    validation_box = _baseline_validation_box(errors, len(df))

    if df.empty or errors:
        return validation_box, None, {"display": "none"}, None

    preview = dash_table.DataTable(
        data=_table_records(df, list(df.columns)),
        columns=_table_columns(df.columns),
        page_size=15,
        style_table={"overflowX": "auto"},
        style_cell=TABLE_CELL_STYLE,
        style_header=TABLE_HEADER_STYLE,
        style_data_conditional=TABLE_CONDITIONAL_STYLE,
    )

    return (
        validation_box,
        html.Div([html.H4(f"Preview — {filename}"), preview], className="preview-box"),
        {"display": "inline-block"},
        df.to_json(date_format="iso", orient="split"),
    )


# ---------------------------------------------------------------------------
# Callback 5: "Confirm Baseline Import" -> write to database
# ---------------------------------------------------------------------------

@callback(
    Output("baseline-import-confirm-result", "children"),
    Output("parsed-baseline-store", "data", allow_duplicate=True),
    Input("confirm-baseline-import-btn", "n_clicks"),
    State("parsed-baseline-store", "data"),
    prevent_initial_call=True,
)
def confirm_baseline_import(n_clicks, stored_json):
    if not n_clicks or not stored_json:
        return dash.no_update, dash.no_update

    df = pd.read_json(io.StringIO(stored_json), orient="split")
    written = upsert_historical_baselines(df)
    return (
        html.Div(f"✅ {written} rows saved as historical baseline.", className="success-text"),
        None,
    )


# ---------------------------------------------------------------------------
# Callback 6: Download the A3 PDF report for the latest report date
# ---------------------------------------------------------------------------

@callback(
    Output("pdf-report-download", "data"),
    Input("download-pdf-report-btn", "n_clicks"),
    prevent_initial_call=True,
)
def download_pdf_report(n_clicks):
    if not n_clicks:
        return dash.no_update

    master_df = load_master_data()
    if master_df.empty:
        return dash.no_update

    baseline_df = load_historical_baselines()
    baseline_df = baseline_df[baseline_df.get("include_in_dashboard", True)] if not baseline_df.empty else baseline_df

    latest_date = master_df["date"].max()
    pdf_bytes = build_pdf_report(master_df, baseline_df, latest_date)
    filename = f"repair_rate_report_{latest_date.date().isoformat()}.pdf"

    return dcc.send_bytes(pdf_bytes, filename)


# ---------------------------------------------------------------------------
# Callback 7: Load / refresh the dashboard
# ---------------------------------------------------------------------------

@callback(
    Output("dashboard-content", "children"),
    Input("refresh-dashboard-btn", "n_clicks"),
    Input("import-confirm-result", "children"),  # auto-refresh after a successful import
    Input("baseline-import-confirm-result", "children"),  # auto-refresh after a baseline import
)
def render_dashboard(_n_clicks, _import_result, _baseline_import_result):
    master_df = load_master_data()

    if master_df.empty:
        return _empty_state("No data in the database yet. Upload and import an Excel file first.")

    baseline_df = load_historical_baselines()
    baseline_df = baseline_df[baseline_df.get("include_in_dashboard", True)] if not baseline_df.empty else baseline_df

    # --- Top summary cards ---
    latest_date = master_df["date"].max()
    latest_df = master_df[master_df["date"] == latest_date]
    latest_df = apply_meter_based_repair_ratios(latest_df)

    overall_ratio = daily_weighted_repair_ratios(master_df, baseline_df)
    current_overall_ratio = (
        overall_ratio.loc[overall_ratio["date"] == latest_date, "weighted_repair_ratio"].iloc[0]
        if not overall_ratio.empty
        else 0
    )

    summary_cards = html.Div(
        [
            _summary_card("Last Report Date", latest_date.strftime("%d.%m.%Y")),
            _summary_card("Active Project Count", str(latest_df["project_no"].nunique())),
            _summary_card("Overall Repair Rate", f"%{current_overall_ratio * 100:.2f}"),
        ],
        className="summary-cards",
    )

    # --- Trend chart: overall repair rate over time ---
    trend_fig = go.Figure()
    trend_fig.add_trace(
        go.Scatter(
            x=overall_ratio["date"],
            y=overall_ratio["weighted_repair_ratio"] * 100,
            mode="lines+markers",
            name="Repair Rate (%)",
            line=dict(color=COLOR_COIL, width=3),
            hovertemplate="%{x|%d.%m.%Y}<br>Repair Rate: <b>%{y:.2f}%</b><extra></extra>",
        )
    )
    trend_fig.update_layout(
        title="Overall Repair Rate Trend",
        yaxis_title="Repair Rate (%)",
        xaxis_title="Date",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=40),
        hoverlabel=HOVER_STYLE,
    )
    trend_fig.update_xaxes(tickformat="%d.%m.%Y", dtick="D1")

    # --- Daily repair amount (bar) ---
    daily_amount = repair_amount_trend_data(master_df, display_unit="m")
    bar_fig = px.bar(
        daily_amount,
        x="date",
        y="daily_repair_amount_display",
        labels={"date": "Date", "daily_repair_amount_display": "Daily Repair Amount (m)"},
        title="Daily Repair Amount",
    )
    bar_fig.update_traces(
        marker_color=COLOR_COIL,
        hovertemplate="%{x|%d.%m.%Y}<br>Daily Repair Amount: <b>%{y:.2f} m</b><extra></extra>",
    )
    bar_fig.update_layout(template="plotly_white", margin=dict(l=40, r=20, t=50, b=40), hoverlabel=HOVER_STYLE)
    bar_fig.update_xaxes(tickformat="%d.%m.%Y", dtick="D1")

    # --- Worst-performing projects (latest day) ---
    worst = latest_df.sort_values("repair_ratio", ascending=False).head(10)
    worst_fig = px.bar(
        worst,
        x="repair_ratio",
        y="project_no",
        orientation="h",
        labels={"repair_ratio": "Repair Ratio", "project_no": "Project"},
        title="Top 10 Projects by Repair Ratio (Latest Day)",
    )
    worst_fig.update_traces(
        marker_color=COLOR_COIL,
        hovertemplate="%{y}<br>Repair Ratio: <b>%{x:.2%}</b><extra></extra>",
    )
    worst_fig.update_layout(
        template="plotly_white",
        yaxis=dict(autorange="reversed"),
        margin=dict(l=120, r=20, t=50, b=40),
        hoverlabel=HOVER_STYLE,
    )
    worst_fig.update_xaxes(tickformat=".0%")

    # --- Skelp impact: how much "incl. skelp" adds on top of the base ratio ---
    skelp_df = latest_df[["project_no", "repair_ratio", "repair_ratio_incl_skelp"]].copy()
    skelp_df["skelp_impact"] = skelp_df["repair_ratio_incl_skelp"] - skelp_df["repair_ratio"]
    skelp_top = skelp_df.sort_values("skelp_impact", ascending=False).head(10)
    skelp_fig = go.Figure()
    skelp_fig.add_trace(
        go.Bar(
            x=skelp_top["project_no"],
            y=skelp_top["repair_ratio"],
            name="Repair Ratio",
            marker_color=COLOR_COIL,
            hovertemplate="%{x}<br>Repair Ratio: <b>%{y:.2%}</b><extra></extra>",
        )
    )
    skelp_fig.add_trace(
        go.Bar(
            x=skelp_top["project_no"],
            y=skelp_top["skelp_impact"],
            name="Skelp Impact",
            marker_color=COLOR_SECONDARY,
            hovertemplate="%{x}<br>Skelp Impact: <b>+%{y:.2%}</b><extra></extra>",
        )
    )
    skelp_fig.update_layout(
        barmode="stack",
        title="Skelp Impact on Repair Ratio (Top 10, Latest Day)",
        yaxis_title="Repair Ratio",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=80),
        hoverlabel=HOVER_STYLE,
    )
    skelp_fig.update_xaxes(tickangle=-45)
    skelp_fig.update_yaxes(tickformat=".0%")

    # --- Repair amount by dimension ---
    dimension_df = (
        latest_df.groupby("dimensions", as_index=False)
        .agg(
            total_repair_amount=("total_repair_amount", "sum"),
            repaired_spiral_length=("repaired_spiral_length", "sum"),
            project_count=("project_no", "nunique"),
        )
        .sort_values("total_repair_amount", ascending=False)
        .head(15)
    )
    dimension_fig = px.bar(
        dimension_df,
        x="dimensions",
        y="total_repair_amount",
        labels={"dimensions": "Dimensions", "total_repair_amount": "Total Repair Amount (m)"},
        title="Repair Amount by Dimension (Latest Day)",
    )
    dimension_fig.update_traces(
        marker_color=COLOR_COIL,
        hovertemplate="%{x}<br>Total Repair Amount: <b>%{y:.2f} m</b><extra></extra>",
    )
    dimension_fig.update_layout(template="plotly_white", margin=dict(l=40, r=20, t=50, b=80), hoverlabel=HOVER_STYLE)
    dimension_fig.update_xaxes(tickangle=-45)

    # --- Weighted repair ratio by dimension ---
    dimension_ratio_df = dimension_df.copy()
    denom = dimension_ratio_df["repaired_spiral_length"] * METERS_PER_FOOT
    dimension_ratio_df["weighted_ratio"] = (dimension_ratio_df["total_repair_amount"] / denom.where(denom != 0)).fillna(0)
    dimension_ratio_df = dimension_ratio_df.sort_values("weighted_ratio", ascending=False)
    dimension_ratio_fig = px.bar(
        dimension_ratio_df,
        x="dimensions",
        y="weighted_ratio",
        labels={"dimensions": "Dimensions", "weighted_ratio": "Weighted Repair Ratio"},
        title="Weighted Repair Ratio by Dimension (Latest Day)",
    )
    dimension_ratio_fig.update_traces(
        marker_color=COLOR_COIL,
        hovertemplate="%{x}<br>Weighted Repair Ratio: <b>%{y:.2%}</b><extra></extra>",
    )
    dimension_ratio_fig.update_layout(template="plotly_white", margin=dict(l=40, r=20, t=50, b=80), hoverlabel=HOVER_STYLE)
    dimension_ratio_fig.update_xaxes(tickangle=-45)
    dimension_ratio_fig.update_yaxes(tickformat=".1%")

    # --- Repair amount pareto: which projects drive most of the total ---
    pareto_df = latest_df[["project_no", "total_repair_amount"]].sort_values(
        "total_repair_amount", ascending=False
    ).reset_index(drop=True)
    pareto_df["cumulative_pct"] = (
        pareto_df["total_repair_amount"].cumsum() / pareto_df["total_repair_amount"].sum() * 100
    )
    pareto_fig = go.Figure()
    pareto_fig.add_trace(
        go.Bar(
            x=pareto_df["project_no"],
            y=pareto_df["total_repair_amount"],
            name="Repair Amount (m)",
            marker_color=COLOR_COIL,
            hovertemplate="%{x}<br>Repair Amount: <b>%{y:.2f} m</b><extra></extra>",
        )
    )
    pareto_fig.add_trace(
        go.Scatter(
            x=pareto_df["project_no"],
            y=pareto_df["cumulative_pct"],
            name="Cumulative %",
            yaxis="y2",
            line=dict(color=COLOR_SECONDARY, width=2),
            mode="lines+markers",
            hovertemplate="%{x}<br>Cumulative: <b>%{y:.2f}%</b><extra></extra>",
        )
    )
    pareto_fig.update_layout(
        title="Repair Amount Pareto (Latest Day)",
        template="plotly_white",
        yaxis=dict(title="Repair Amount (m)"),
        yaxis2=dict(title="Cumulative %", overlaying="y", side="right", range=[0, 105]),
        margin=dict(l=40, r=40, t=50, b=80),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hoverlabel=HOVER_STYLE,
    )
    pareto_fig.update_xaxes(tickangle=-45)

    production_types = sorted(master_df["production_type"].dropna().unique())
    _TYPE_COLORS = {"Coil": COLOR_COIL, "Plate": COLOR_PLATE}

    # --- Repair rate trend by production type ---
    type_trend_fig = go.Figure()
    for p_type in production_types:
        type_trend = daily_weighted_repair_ratios_for_type(master_df, p_type, baseline_df)
        type_trend_fig.add_trace(
            go.Scatter(
                x=type_trend["date"],
                y=type_trend["weighted_repair_ratio"] * 100,
                mode="lines+markers",
                name=p_type,
                line=dict(width=3, color=_TYPE_COLORS.get(p_type)),
                hovertemplate=f"%{{x|%d.%m.%Y}}<br>{p_type}: <b>%{{y:.2f}}%</b><extra></extra>",
            )
        )
    type_trend_fig.update_layout(
        title="Repair Rate Trend by Production Type",
        yaxis_title="Repair Rate (%)",
        xaxis_title="Date",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=40),
        hoverlabel=HOVER_STYLE,
    )
    type_trend_fig.update_xaxes(tickformat="%d.%m.%Y", dtick="D1")

    # --- Weighted repair ratio by production type (latest day) ---
    type_rows = []
    for p_type in production_types:
        type_latest = latest_df[latest_df["production_type"] == p_type]
        denom = type_latest["repaired_spiral_length"].sum() * METERS_PER_FOOT
        ratio = type_latest["total_repair_amount"].sum() / denom if denom else 0
        type_rows.append({"production_type": p_type, "weighted_ratio": ratio})
    type_analysis_fig = px.bar(
        pd.DataFrame(type_rows),
        x="production_type",
        y="weighted_ratio",
        color="production_type",
        color_discrete_map=_TYPE_COLORS,
        labels={"production_type": "Production Type", "weighted_ratio": "Weighted Repair Ratio"},
        title="Weighted Repair Ratio by Production Type (Latest Day)",
    )
    type_analysis_fig.update_traces(hovertemplate="%{x}<br>Weighted Repair Ratio: <b>%{y:.2%}</b><extra></extra>")
    type_analysis_fig.update_layout(
        template="plotly_white", margin=dict(l=40, r=20, t=50, b=40), showlegend=False, hoverlabel=HOVER_STYLE
    )
    type_analysis_fig.update_yaxes(tickformat=".1%")

    # --- Weighted repair ratio by status (latest day) ---
    status_rows = []
    for status in sorted(latest_df["project_status"].dropna().unique()):
        status_latest = latest_df[latest_df["project_status"] == status]
        denom = status_latest["repaired_spiral_length"].sum() * METERS_PER_FOOT
        ratio = status_latest["total_repair_amount"].sum() / denom if denom else 0
        status_rows.append({"project_status": status, "weighted_ratio": ratio})
    status_fig = px.bar(
        pd.DataFrame(status_rows),
        x="project_status",
        y="weighted_ratio",
        labels={"project_status": "Status", "weighted_ratio": "Weighted Repair Ratio"},
        title="Weighted Repair Ratio by Status (Latest Day)",
    )
    status_fig.update_traces(
        marker_color=COLOR_COIL,
        hovertemplate="%{x}<br>Weighted Repair Ratio: <b>%{y:.2%}</b><extra></extra>",
    )
    status_fig.update_layout(template="plotly_white", margin=dict(l=40, r=20, t=50, b=40), hoverlabel=HOVER_STYLE)
    status_fig.update_yaxes(tickformat=".1%")

    # --- Current period vs historical baseline ---
    current_denom = latest_df["repaired_spiral_length"].sum() * METERS_PER_FOOT
    current_ratio_value = latest_df["total_repair_amount"].sum() / current_denom if current_denom else 0
    baseline_ratio_value = 0
    if not baseline_df.empty:
        baseline_denom = baseline_df["repaired_spiral_length"].sum() * METERS_PER_FOOT
        baseline_ratio_value = baseline_df["total_repair_amount"].sum() / baseline_denom if baseline_denom else 0
    benchmark_fig = px.bar(
        pd.DataFrame(
            {
                "period": ["Current (Latest Day)", "Historical Baseline"],
                "ratio": [current_ratio_value, baseline_ratio_value],
            }
        ),
        x="period",
        y="ratio",
        labels={"period": "", "ratio": "Weighted Repair Ratio"},
        title="Current vs Historical Baseline Repair Ratio",
    )
    benchmark_fig.update_traces(
        marker_color=[COLOR_COIL, COLOR_MUTED],
        hovertemplate="%{x}<br>Weighted Repair Ratio: <b>%{y:.2%}</b><extra></extra>",
    )
    benchmark_fig.update_layout(template="plotly_white", margin=dict(l=40, r=20, t=50, b=40), hoverlabel=HOVER_STYLE)
    benchmark_fig.update_yaxes(tickformat=".1%")

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
                    dcc.Graph(figure=trend_fig, className="chart-half"),
                    dcc.Graph(figure=bar_fig, className="chart-half"),
                ],
                className="chart-row",
            ),
            dcc.Graph(figure=worst_fig),
            html.Div(
                [
                    dcc.Graph(figure=skelp_fig, className="chart-half"),
                    dcc.Graph(figure=dimension_fig, className="chart-half"),
                ],
                className="chart-row",
            ),
            dcc.Graph(figure=dimension_ratio_fig),
            dcc.Graph(figure=pareto_fig),
            dcc.Graph(figure=type_trend_fig),
            html.Div(
                [
                    dcc.Graph(figure=type_analysis_fig, className="chart-half"),
                    dcc.Graph(figure=status_fig, className="chart-half"),
                ],
                className="chart-row",
            ),
            dcc.Graph(figure=benchmark_fig),
            html.H3("Latest Day — Project Details"),
            detail_table,
        ]
    )


def _summary_card(label: str, value: str) -> html.Div:
    return html.Div(
        [html.Div(value, className="summary-value"), html.Div(label, className="summary-label")],
        className="summary-card",
    )


# ---------------------------------------------------------------------------
# Callback X1: Project Trend — populate the project dropdown
# ---------------------------------------------------------------------------

@callback(
    Output("project-trend-dropdown", "options"),
    Output("project-trend-dropdown", "value"),
    Input("project-trend-dropdown", "id"),  # fires once, on page load
    Input("import-confirm-result", "children"),  # refresh after a new import
)
def load_project_trend_options(_id, _import_result):
    master_df = load_master_data()
    if master_df.empty:
        return [], None
    projects = sorted(master_df["project_no"].dropna().unique())
    options = [{"label": p, "value": p} for p in projects]
    return options, options[0]["value"] if options else None


# ---------------------------------------------------------------------------
# Callback X2: Project Trend — render the trend chart for the selected project
# ---------------------------------------------------------------------------

@callback(
    Output("project-trend-content", "children"),
    Input("project-trend-dropdown", "value"),
)
def render_project_trend(selected_project):
    if not selected_project:
        return _empty_state("No data available. Import an Excel file first.")

    master_df = load_master_data()
    project_df = master_df[master_df["project_no"] == selected_project].copy()
    if project_df.empty:
        return _empty_state("No data available for this project.")

    project_df = apply_meter_based_repair_ratios(project_df).sort_values("date")

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=project_df["date"],
            y=project_df["repair_ratio"] * 100,
            mode="lines+markers",
            name="Repair Ratio (%)",
            line=dict(color=COLOR_COIL, width=3),
            hovertemplate="%{x|%d.%m.%Y}<br>Repair Ratio: <b>%{y:.2f}%</b><extra></extra>",
        )
    )
    fig.update_layout(
        title=f"Repair Ratio Trend — {selected_project}",
        yaxis_title="Repair Ratio (%)",
        xaxis_title="Date",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=40),
        hoverlabel=HOVER_STYLE,
    )
    fig.update_xaxes(tickformat="%d.%m.%Y", dtick="D1")

    return dcc.Graph(figure=fig)


# ---------------------------------------------------------------------------
# Callback 8: Pipe-Level Analysis — populate the report date dropdown
# ---------------------------------------------------------------------------

@callback(
    Output("pipe-date-dropdown", "options"),
    Output("pipe-date-dropdown", "value"),
    Input("pipe-date-dropdown", "id"),  # fires once, on page load
    Input("import-confirm-result", "children"),  # refresh after a new import
)
def load_pipe_dates(_id, _import_result):
    df = load_pipe_repair_details()
    if df.empty:
        return [], None
    dates = sorted(df["date"].dt.date.unique(), reverse=True)
    options = [{"label": d.strftime("%d.%m.%Y"), "value": d.isoformat()} for d in dates]
    return options, options[0]["value"]


# ---------------------------------------------------------------------------
# Callback 9: Pipe-Level Analysis — populate the project sheet dropdown
# ---------------------------------------------------------------------------

@callback(
    Output("pipe-sheet-dropdown", "options"),
    Output("pipe-sheet-dropdown", "value"),
    Input("pipe-date-dropdown", "value"),
)
def load_pipe_sheets(selected_date):
    if not selected_date:
        return [], None
    df = load_pipe_repair_details()
    if df.empty:
        return [], None
    day_df = df[df["date"].dt.date.astype(str) == selected_date]
    sheets = sorted(day_df["project_sheet"].unique())
    options = [{"label": s, "value": s} for s in sheets]
    return options, options[0]["value"] if options else None


# ---------------------------------------------------------------------------
# Callback 10: Pipe-Level Analysis — render the summary + detail view
# ---------------------------------------------------------------------------

@callback(
    Output("pipe-analysis-content", "children"),
    Input("pipe-date-dropdown", "value"),
    Input("pipe-sheet-dropdown", "value"),
)
def render_pipe_analysis(selected_date, selected_sheet):
    if not selected_date:
        return _empty_state("No pipe-level data in the database yet. Import an Excel file first.")

    df = load_pipe_repair_details()
    day_df = df[df["date"].dt.date.astype(str) == selected_date]

    summary_table = summarize_pipe_totals_by_sheet(day_df)
    summary_columns = list(summary_table.columns)
    summary_section = html.Div(
        [
            html.H3("All Project Sheets — Summary"),
            dash_table.DataTable(
                data=_table_records(summary_table, summary_columns),
                columns=_table_columns(summary_columns),
                page_size=10,
                sort_action="native",
                style_table={"overflowX": "auto"},
                style_cell=TABLE_CELL_STYLE,
                style_header=TABLE_HEADER_STYLE,
                style_data_conditional=TABLE_CONDITIONAL_STYLE,
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
    worst_fig.update_traces(
        marker_color=COLOR_COIL,
        hovertemplate="Pipe %{x}<br>Repair Ratio: <b>%{y:.2%}</b><extra></extra>",
    )
    worst_fig.update_layout(template="plotly_white", margin=dict(l=40, r=20, t=50, b=40), hoverlabel=HOVER_STYLE)
    worst_fig.update_xaxes(type="category")
    worst_fig.update_yaxes(tickformat=".1%")

    detail_columns = [c for c in sheet_df.columns if c not in ("date", "surface_state", "repair_category")]
    detail_table = dash_table.DataTable(
        data=_table_records(sheet_df, detail_columns),
        columns=_table_columns(detail_columns),
        page_size=20,
        sort_action="native",
        filter_action="native",
        style_table={"overflowX": "auto"},
        style_cell=TABLE_CELL_STYLE,
        style_header=TABLE_HEADER_STYLE,
        style_data_conditional=TABLE_CONDITIONAL_STYLE,
    )

    return html.Div(
        [
            summary_section,
            html.H3(f"{selected_sheet} — Pipe Details"),
            dcc.Graph(figure=worst_fig),
            detail_table,
        ]
    )


# ---------------------------------------------------------------------------
# Callback 11: Project Grouping — populate the project sheet dropdown
# ---------------------------------------------------------------------------

@callback(
    Output("group-sheet-dropdown", "options"),
    Output("group-sheet-dropdown", "value"),
    Input("group-sheet-dropdown", "id"),  # fires once, on page load
)
def load_group_sheets(_id):
    df = load_pipe_repair_details()
    if df.empty:
        return [], None
    sheets = sorted(df["project_sheet"].unique())
    options = [{"label": s, "value": s} for s in sheets]
    return options, options[0]["value"] if options else None


# ---------------------------------------------------------------------------
# Callback 12: Project Grouping — load the saved config for a sheet
# ---------------------------------------------------------------------------

@callback(
    Output("pipe-groups-input", "value"),
    Output("machine-groups-input", "value"),
    Output("save-groups-result", "children", allow_duplicate=True),
    Input("group-sheet-dropdown", "value"),
    prevent_initial_call=True,
)
def load_group_config(selected_sheet):
    if not selected_sheet:
        return "", "", None
    config = load_project_group_config(selected_sheet, selected_sheet, "")
    if not config:
        return "", "", None
    return config.get("pipe_groups", ""), config.get("machine_groups", ""), None


# ---------------------------------------------------------------------------
# Callback 13: Project Grouping — save the group spec for a sheet
# ---------------------------------------------------------------------------

@callback(
    Output("save-groups-result", "children"),
    Input("save-groups-btn", "n_clicks"),
    State("group-sheet-dropdown", "value"),
    State("pipe-groups-input", "value"),
    State("machine-groups-input", "value"),
    prevent_initial_call=True,
)
def save_groups(n_clicks, selected_sheet, pipe_groups, machine_groups):
    if not n_clicks or not selected_sheet:
        return dash.no_update
    upsert_project_group_config(selected_sheet, selected_sheet, "", pipe_groups or "", machine_groups or "")
    return html.Div(f"✅ Groups saved for {selected_sheet}.", className="success-text")


# ---------------------------------------------------------------------------
# Callback 14: Dimension Detail — populate the dimension dropdown
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
# Callback 15: Dimension Detail — compare projects sharing the selected dimension
# ---------------------------------------------------------------------------

@callback(
    Output("dimension-detail-content", "children"),
    Input("dimension-detail-dropdown", "value"),
)
def render_dimension_detail(selected_dimension):
    if not selected_dimension:
        return _empty_state("No data available. Import an Excel file first.")

    master_df = load_master_data()
    latest_date = master_df["date"].max()
    latest_df = master_df[master_df["date"] == latest_date].copy()
    latest_df = apply_meter_based_repair_ratios(latest_df)

    dim_df = latest_df[latest_df["dimensions"] == selected_dimension].sort_values(
        "repair_ratio", ascending=False
    )
    if dim_df.empty:
        return _empty_state("No projects found for this dimension on the latest day.")

    fig = px.bar(
        dim_df,
        x="project_no",
        y="repair_ratio",
        labels={"project_no": "Project", "repair_ratio": "Repair Ratio"},
        title=f"Projects with Dimension {selected_dimension} (Latest Day)",
    )
    fig.update_traces(
        marker_color=COLOR_COIL,
        hovertemplate="%{x}<br>Repair Ratio: <b>%{y:.2%}</b><extra></extra>",
    )
    fig.update_layout(template="plotly_white", margin=dict(l=40, r=20, t=50, b=80), hoverlabel=HOVER_STYLE)
    fig.update_xaxes(tickangle=-45)
    fig.update_yaxes(tickformat=".1%")

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

    return html.Div([dcc.Graph(figure=fig), table])
