"""
Module 2: Data Conversion — bulk-upload coil cert PDFs (SDI) and merge the
extracted fields into a single "Report"-formatted Excel matching the
Excel-supplier's own layout (see pdf_cert_parser.py for the field mapping
and column order rationale).

Stateless by design: nothing here is written to the database. Parsed rows
accumulate in a dcc.Store for the current browser session only, and get
downloaded as an .xlsx on demand.
"""

from __future__ import annotations

import base64
import io

import dash
from dash import Input, Output, State, callback, dash_table, dcc, html

from pdf_cert_parser import REPORT_COLUMNS, build_report_rows, build_report_workbook, parse_sdi_cert_pdf

dash.register_page(__name__, path="/data-conversion", name="Data Conversion")

TABLE_HEADER_STYLE = {
    "backgroundColor": "#f4f6f8",
    "fontWeight": "600",
    "border": "none",
    "borderBottom": "2px solid #e2e8f0",
}
TABLE_CELL_STYLE = {"fontFamily": "inherit", "fontSize": "13px", "padding": "8px 10px"}
TABLE_CONDITIONAL_STYLE = [{"if": {"row_index": "odd"}, "backgroundColor": "#f8fafc"}]


def _decode_upload(contents: str) -> io.BytesIO:
    _, content_string = contents.split(",", 1)
    return io.BytesIO(base64.b64decode(content_string))


def layout():
    return html.Div(
        [
            dcc.Store(id="pdf-cert-store", data=[]),
            html.Section(
                [
                    html.H2("Coil Certs → Common Excel"),
                    html.P(
                        "Bulk-upload SDI coil cert PDFs. Extracted fields are merged into one "
                        "Excel matching the Excel-supplier's report format.",
                        className="help-text",
                    ),
                    html.Div(
                        [
                            html.Div(
                                [
                                    html.Label("Source"),
                                    dcc.RadioItems(
                                        id="cert-source-selector",
                                        options=[
                                            {"label": "PDF (SDI)", "value": "pdf"},
                                            {"label": "Excel (Supplier) — coming soon", "value": "excel", "disabled": True},
                                        ],
                                        value="pdf",
                                        inline=True,
                                    ),
                                ],
                                className="filter-field",
                            ),
                        ],
                        className="filter-row",
                    ),
                    dcc.Upload(
                        id="pdf-cert-upload",
                        children=html.Div(["Drag and drop PDF certs here or ", html.A("select files")]),
                        className="upload-box",
                        multiple=True,
                    ),
                    dcc.Loading(html.Div(id="pdf-cert-upload-status")),
                    html.Div(
                        [
                            html.Button("Clear", id="pdf-cert-clear-btn", className="secondary-btn"),
                            html.Button("Download Excel", id="pdf-cert-download-btn", className="primary-btn"),
                        ],
                        className="button-group",
                    ),
                    dcc.Download(id="pdf-cert-download"),
                    dcc.Loading(html.Div(id="pdf-cert-preview")),
                ],
                className="card",
            ),
        ]
    )


# ---------------------------------------------------------------------------
# Callback 1: parse newly-uploaded PDFs and add them to the session store
# ---------------------------------------------------------------------------


@callback(
    Output("pdf-cert-store", "data"),
    Output("pdf-cert-upload-status", "children"),
    Input("pdf-cert-upload", "contents"),
    Input("pdf-cert-clear-btn", "n_clicks"),
    State("pdf-cert-upload", "filename"),
    State("pdf-cert-store", "data"),
    prevent_initial_call=True,
)
def handle_pdf_upload(contents_list, _clear_clicks, filenames, stored_rows):
    triggered = dash.ctx.triggered_id
    if triggered == "pdf-cert-clear-btn":
        return [], None

    if not contents_list:
        return stored_rows or [], dash.no_update

    stored_rows = list(stored_rows or [])
    status_lines = []
    for contents, filename in zip(contents_list, filenames):
        try:
            file_obj = _decode_upload(contents)
            result = parse_sdi_cert_pdf(file_obj, source_file=filename)
        except Exception as exc:  # a single bad PDF shouldn't block the rest of the batch
            status_lines.append(html.P(f"{filename}: failed to read ({exc})", className="help-text"))
            continue

        stored_rows.append({"source_file": result.source_file, "fields": result.fields, "errors": result.errors})
        if result.errors:
            status_lines.append(
                html.P(f"{filename}: {', '.join(result.errors)}", className="help-text")
            )
        else:
            status_lines.append(
                html.P(
                    f"{filename}: OK — Heat {result.fields.get('heat_no')}, "
                    f"Coil {result.fields.get('coil_number')}",
                    className="help-text",
                )
            )

    return stored_rows, html.Div(status_lines)


# ---------------------------------------------------------------------------
# Callback 2: render the heat-grouped preview table from the session store
# ---------------------------------------------------------------------------


@callback(
    Output("pdf-cert-preview", "children"),
    Input("pdf-cert-store", "data"),
)
def render_pdf_cert_preview(stored_rows):
    if not stored_rows:
        return html.Div("No PDFs uploaded yet.", className="empty-state")

    from pdf_cert_parser import PdfCertResult

    results = [
        PdfCertResult(source_file=row["source_file"], fields=row["fields"], errors=row["errors"])
        for row in stored_rows
    ]
    df = build_report_rows(results)
    error_count = sum(1 for r in results if r.errors)

    summary = html.P(
        f"{len(results)} PDF(s) uploaded, {len(df)} heat(s) after grouping."
        + (f" {error_count} file(s) had missing fields — see messages above." if error_count else ""),
        className="help-text",
    )

    if df.empty:
        return html.Div([summary, html.Div("No valid rows to preview yet.", className="empty-state")])

    table = dash_table.DataTable(
        data=df.round(4).to_dict("records"),
        columns=[{"name": c, "id": c} for c in REPORT_COLUMNS],
        page_size=15,
        sort_action="native",
        style_table={"overflowX": "auto"},
        style_cell=TABLE_CELL_STYLE,
        style_header=TABLE_HEADER_STYLE,
        style_data_conditional=TABLE_CONDITIONAL_STYLE,
    )
    return html.Div([summary, table])


# ---------------------------------------------------------------------------
# Callback 3: build and download the merged Excel
# ---------------------------------------------------------------------------


@callback(
    Output("pdf-cert-download", "data"),
    Input("pdf-cert-download-btn", "n_clicks"),
    State("pdf-cert-store", "data"),
    prevent_initial_call=True,
)
def download_pdf_cert_report(n_clicks, stored_rows):
    if not n_clicks or not stored_rows:
        return dash.no_update

    from pdf_cert_parser import PdfCertResult

    results = [
        PdfCertResult(source_file=row["source_file"], fields=row["fields"], errors=row["errors"])
        for row in stored_rows
    ]
    df = build_report_rows(results)
    if df.empty:
        return dash.no_update

    xlsx_bytes = build_report_workbook(df)
    return dcc.send_bytes(xlsx_bytes, "coil_cert_report.xlsx")
