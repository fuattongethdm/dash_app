"""
Modül 3: Katalog / sunum modülü (resim yükleme, görsel içerik).

Şimdilik yer tutucu (placeholder).
"""

from __future__ import annotations

import dash
from dash import html

dash.register_page(__name__, path="/katalog", name="Katalog")


def layout():
    return html.Div(
        [
            html.H2("Katalog Modülü"),
            html.P("Bu modül henüz geliştirilmedi — Modül 1 ve 2'den sonra sırada bu var."),
        ],
        className="card",
    )
