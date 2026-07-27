"""
Modül 2: Excel/PDF'den veri çekme -> yeni Excel oluşturma/düzenleme.

Şimdilik yer tutucu (placeholder) - Modül 1 tamamlanıp test edildikten sonra
buraya geçeceğiz.
"""

from __future__ import annotations

import dash
from dash import html

dash.register_page(__name__, path="/veri-donusum", name="Veri Dönüşüm")


def layout():
    return html.Div(
        [
            html.H2("Veri Dönüşüm Modülü"),
            html.P("Bu modül henüz geliştirilmedi — sırada Modül 1'den sonra bu var."),
        ],
        className="card",
    )
