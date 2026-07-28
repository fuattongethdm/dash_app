"""
Main Dash application.

This file is just the "skeleton": page routing and the overall shell (nav,
title) live here. Each module's actual logic lives in its own file under
`pages/` - that way modules stay independent and can be developed and
tested separately.

When deploying to a platform like Render, the `server` variable in this
file is used (gunicorn app:server).
"""

from __future__ import annotations

import dash
from dash import Dash, html, dcc

app = Dash(
    __name__,
    use_pages=True,
    suppress_callback_exceptions=True,
    title="Daily Tracking Report",
)

# Render / gunicorn looks for this variable: `gunicorn app:server`
server = app.server


def _nav_link(label: str, path: str):
    return dcc.Link(label, href=path, className="nav-link")


app.layout = html.Div(
    [
        html.Header(
            [
                html.Div("Daily Tracking Report", className="app-title"),
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
    # For local development. Production uses gunicorn.
    app.run(debug=True, host="0.0.0.0", port=8050)
