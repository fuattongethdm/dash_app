"""
Modül 1: Günlük Repair Rate Dashboard.

Akış (Streamlit'teki ile birebir aynı mantık):
  1) Kullanıcı Excel dosyasını yükler (dcc.Upload)
  2) parser.parse_daily_repair_rate() ile okunur, validators ile doğrulanır
  3) Doğrulama sonucu ve önizleme kullanıcıya gösterilir
  4) "Import'u Onayla" butonuna basılınca veritabanına yazılır (upsert)
  5) Aşağıdaki dashboard, veritabanındaki TÜM veriye göre güncellenir

Not: Excel'in okunma mantığı (parser.py, validators.py) hiç değiştirilmedi -
sadece üstündeki arayüz Streamlit yerine Dash ile yazıldı.
"""

from __future__ import annotations

import base64
import io

import dash
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import Input, Output, State, callback, dash_table, dcc, html

from calculations import (
    apply_meter_based_repair_ratios,
    daily_weighted_repair_ratios,
    repair_amount_trend_data,
)
from database import load_historical_baselines, load_master_data, upsert_repair_rates
from parser import parse_daily_repair_rate

dash.register_page(__name__, path="/", name="Dashboard")


# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------

def layout():
    return html.Div(
        [
            # Parse edilen (henüz DB'ye yazılmamış) veriyi tarayıcıda tutuyoruz.
            dcc.Store(id="parsed-data-store"),
            html.Section(
                [
                    html.H2("1) Günlük Excel Yükle"),
                    dcc.Upload(
                        id="excel-upload",
                        children=html.Div(
                            ["Excel dosyasını buraya sürükle ya da ", html.A("dosya seç")]
                        ),
                        className="upload-box",
                        multiple=False,
                    ),
                    html.Div(id="upload-validation-result"),
                    html.Div(id="upload-preview-table"),
                    html.Button(
                        "Import'u Onayla",
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
                    html.Div(
                        [
                            html.H2("2) Dashboard"),
                            html.Button("Verileri Yenile", id="refresh-dashboard-btn", className="secondary-btn"),
                        ],
                        className="section-header-row",
                    ),
                    dcc.Loading(html.Div(id="dashboard-content")),
                ],
                className="card",
            ),
        ]
    )


# ---------------------------------------------------------------------------
# Yardımcı fonksiyonlar
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
        children.append(html.H4(f"Hatalar ({len(report.errors)})", className="error-heading"))
        children.append(html.Ul(error_rows, className="error-list"))
    else:
        children.append(html.P(f"{report.import_rows} satır okundu, hata yok.", className="success-text"))

    return html.Div(children, className="validation-box ok" if report.ok else "validation-box fail")


# ---------------------------------------------------------------------------
# Callback 1: Excel yükle -> parse et -> doğrula -> önizleme göster
# ---------------------------------------------------------------------------

@callback(
    Output("upload-validation-result", "children"),
    Output("upload-preview-table", "children"),
    Output("confirm-import-btn", "style"),
    Output("parsed-data-store", "data"),
    Input("excel-upload", "contents"),
    State("excel-upload", "filename"),
    prevent_initial_call=True,
)
def handle_upload(contents, filename):
    if contents is None:
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update

    try:
        file_obj = _decode_upload(contents)
        df, report = parse_daily_repair_rate(file_obj)
    except Exception as exc:  # beklenmeyen bir hata olursa uygulamayı çökertme
        error_box = html.Div(
            f"Dosya okunurken beklenmeyen bir hata oluştu: {exc}",
            className="validation-box fail",
        )
        return error_box, None, {"display": "none"}, None

    validation_box = _validation_summary(report)

    if df.empty or not report.ok:
        return validation_box, None, {"display": "none"}, None

    preview = dash_table.DataTable(
        data=df.drop(columns=["excel_row"], errors="ignore").to_dict("records"),
        columns=[{"name": c, "id": c} for c in df.columns if c != "excel_row"],
        page_size=15,
        style_table={"overflowX": "auto"},
        style_cell={"fontFamily": "inherit", "fontSize": "13px", "padding": "6px"},
    )

    return (
        validation_box,
        html.Div([html.H4(f"Önizleme — {filename}"), preview], className="preview-box"),
        {"display": "inline-block"},
        df.to_json(date_format="iso", orient="split"),
    )


# ---------------------------------------------------------------------------
# Callback 2: "Import'u Onayla" -> veritabanına yaz
# ---------------------------------------------------------------------------

@callback(
    Output("import-confirm-result", "children"),
    Output("parsed-data-store", "data", allow_duplicate=True),
    Input("confirm-import-btn", "n_clicks"),
    State("parsed-data-store", "data"),
    prevent_initial_call=True,
)
def confirm_import(n_clicks, stored_json):
    if not n_clicks or not stored_json:
        return dash.no_update, dash.no_update

    df = pd.read_json(io.StringIO(stored_json), orient="split")
    written = upsert_repair_rates(df)
    return (
        html.Div(f"✅ {written} satır veritabanına kaydedildi.", className="success-text"),
        None,
    )


# ---------------------------------------------------------------------------
# Callback 3: Dashboard'u yükle / yenile
# ---------------------------------------------------------------------------

@callback(
    Output("dashboard-content", "children"),
    Input("refresh-dashboard-btn", "n_clicks"),
    Input("import-confirm-result", "children"),  # import başarılı olunca otomatik yenile
)
def render_dashboard(_n_clicks, _import_result):
    master_df = load_master_data()

    if master_df.empty:
        return html.P("Henüz veritabanında veri yok. Önce bir Excel yükleyip import edin.")

    baseline_df = load_historical_baselines()
    baseline_df = baseline_df[baseline_df.get("include_in_dashboard", True)] if not baseline_df.empty else baseline_df

    # --- Üst özet kartları ---
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
            _summary_card("Son Rapor Tarihi", latest_date.strftime("%d.%m.%Y")),
            _summary_card("Aktif Proje Sayısı", str(latest_df["project_no"].nunique())),
            _summary_card("Genel Repair Rate", f"%{current_overall_ratio * 100:.2f}"),
        ],
        className="summary-cards",
    )

    # --- Trend grafiği: genel repair rate zaman içinde ---
    trend_fig = go.Figure()
    trend_fig.add_trace(
        go.Scatter(
            x=overall_ratio["date"],
            y=overall_ratio["weighted_repair_ratio"] * 100,
            mode="lines+markers",
            name="Repair Rate (%)",
            line=dict(color="#2563eb", width=3),
        )
    )
    trend_fig.update_layout(
        title="Genel Repair Rate Trendi",
        yaxis_title="Repair Rate (%)",
        xaxis_title="Tarih",
        template="plotly_white",
        margin=dict(l=40, r=20, t=50, b=40),
    )

    # --- Günlük tamir miktarı (bar) ---
    daily_amount = repair_amount_trend_data(master_df, display_unit="m")
    bar_fig = px.bar(
        daily_amount,
        x="date",
        y="daily_repair_amount_display",
        labels={"date": "Tarih", "daily_repair_amount_display": "Günlük Tamir Miktarı (m)"},
        title="Günlük Tamir Miktarı",
    )
    bar_fig.update_layout(template="plotly_white", margin=dict(l=40, r=20, t=50, b=40))

    # --- En kötü performanslı projeler (son gün) ---
    worst = latest_df.sort_values("repair_ratio", ascending=False).head(10)
    worst_fig = px.bar(
        worst,
        x="repair_ratio",
        y="project_no",
        orientation="h",
        labels={"repair_ratio": "Repair Ratio", "project_no": "Proje"},
        title="En Yüksek Repair Ratio'ya Sahip 10 Proje (Son Gün)",
    )
    worst_fig.update_layout(
        template="plotly_white",
        yaxis=dict(autorange="reversed"),
        margin=dict(l=120, r=20, t=50, b=40),
    )

    # --- Son gün detay tablosu ---
    table_columns = [
        "project_no",
        "dimensions",
        "production_type",
        "qty",
        "project_status",
        "repair_ratio",
    ]
    detail_table = dash_table.DataTable(
        data=latest_df[table_columns].round(4).to_dict("records"),
        columns=[{"name": c, "id": c} for c in table_columns],
        page_size=15,
        sort_action="native",
        filter_action="native",
        style_table={"overflowX": "auto"},
        style_cell={"fontFamily": "inherit", "fontSize": "13px", "padding": "6px"},
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
            html.H3("Son Gün — Proje Detayları"),
            detail_table,
        ]
    )


def _summary_card(label: str, value: str) -> html.Div:
    return html.Div(
        [html.Div(value, className="summary-value"), html.Div(label, className="summary-label")],
        className="summary-card",
    )
