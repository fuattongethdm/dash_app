"""
Ana Dash uygulaması.

Bu dosya sadece "iskelet": sayfa yönlendirme (routing) ve genel görünüm (menü,
başlık) burada. Her modülün asıl mantığı `pages/` klasöründeki kendi
dosyasında - böylece modüller birbirine karışmıyor, her biri bağımsız
geliştirilip test edilebiliyor.

Render gibi bir platforma deploy ederken bu dosyadaki `server` değişkeni
kullanılıyor (gunicorn app:server).
"""

from __future__ import annotations

import dash
from dash import Dash, html, dcc

app = Dash(
    __name__,
    use_pages=True,
    suppress_callback_exceptions=True,
    title="Fabrika Takip Sistemi",
)

# Render / gunicorn bu değişkeni arıyor: `gunicorn app:server`
server = app.server


def _nav_link(label: str, path: str):
    return dcc.Link(label, href=path, className="nav-link")


app.layout = html.Div(
    [
        html.Header(
            [
                html.Div("Fabrika Takip Sistemi", className="app-title"),
                html.Nav(
                    [_nav_link(page["name"], page["path"]) for page in dash.page_registry.values()],
                    className="app-nav",
                ),
            ],
            className="app-header",
        ),
        html.Main(dash.page_container, className="app-main"),
    ]
)


if __name__ == "__main__":
    # Lokal geliştirme için. Production'da gunicorn kullanılacak.
    app.run(debug=True, host="0.0.0.0", port=8050)
