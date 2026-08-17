"""Gunicorn config — auto-loaded when the start command is plain
`gunicorn app:server` (no `-c` flag), which is how this app runs on Render.

The default 30s worker timeout is too short for the PDF report endpoint
(pages/home.py: download_pdf_report), which renders ~10 matplotlib charts
synchronously in a single request — on Render's shared CPU that can run
past 30s, so gunicorn aborts the worker mid-render (SIGABRT, then SIGKILL
if it doesn't exit fast enough) and the request comes back as a 500.
"""

timeout = 120
