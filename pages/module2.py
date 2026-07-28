"""
Module 2: Extract data from Excel/PDF -> generate/edit a new Excel.

Placeholder for now - we'll move on to this once Module 1 is complete and tested.
"""

from __future__ import annotations

import dash
from dash import html

dash.register_page(__name__, path="/data-conversion", name="Data Conversion")


def layout():
    return html.Div(
        [
            html.H2("Data Conversion Module"),
            html.P("This module has not been built yet — it's next after Module 1."),
        ],
        className="card",
    )
