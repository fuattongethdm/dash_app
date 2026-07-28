"""
Module 3: Catalog / presentation module (image upload, visual content).

Placeholder for now.
"""

from __future__ import annotations

import dash
from dash import html

dash.register_page(__name__, path="/catalog", name="Catalog")


def layout():
    return html.Div(
        [
            html.H2("Catalog Module"),
            html.P("This module has not been built yet — it's next after Modules 1 and 2."),
        ],
        className="card",
    )
